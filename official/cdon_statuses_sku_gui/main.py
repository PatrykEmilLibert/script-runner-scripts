import base64
import csv
import json
import os
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import requests


APP_TITLE = "CDON Statusy Ofert po SKU"
DEFAULT_ACCOUNTS_PATH = rstr((Path(__file__).parent / "accounts.csv").resolve())
SANDBOX_BASE_URL = "https://merchants-api.sandbox.cdon.com/api"
PROD_BASE_URL = "https://merchants-api.cdon.com/api"
STATUS_BY_SKU_ENDPOINT = "/v1/statuses/sku"

ACCENT = "#ff69b4"
ACCENT_HOVER = "#ec4fa3"
BG_MAIN = "#fff7fb"
BG_PANEL = "#ffffff"
BG_INPUT = "#fff0f8"
BORDER = "#f3bdd8"
TEXT = "#3c2030"
TEXT_MUTED = "#6f4f60"
SUCCESS = "#1f8a4c"
ERROR = "#b42318"

DEFAULT_REQUESTS_PER_MINUTE = 300
MAX_REQUESTS_PER_MINUTE = 350
DEFAULT_MAX_RETRIES = 6
DEFAULT_BACKOFF_SECONDS = 1.5


def chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i : i + size]


class RequestRateLimiter:
    def __init__(self, requests_per_minute):
        if requests_per_minute < 1:
            raise ValueError("requests_per_minute musi byc > 0")
        self.min_interval = 60.0 / float(requests_per_minute)
        self._lock = threading.Lock()
        self._next_allowed_ts = 0.0

    def wait_turn(self):
        while True:
            with self._lock:
                now = time.monotonic()
                if now >= self._next_allowed_ts:
                    self._next_allowed_ts = now + self.min_interval
                    return
                wait_for = self._next_allowed_ts - now
            time.sleep(min(wait_for, 0.5))


class CDONStatusesClient:
    def __init__(
        self,
        merchant_id,
        api_token,
        use_sandbox=True,
        timeout=45,
        requests_per_minute=DEFAULT_REQUESTS_PER_MINUTE,
        max_retries=DEFAULT_MAX_RETRIES,
        backoff_seconds=DEFAULT_BACKOFF_SECONDS,
        log_callback=None,
    ):
        self.merchant_id = merchant_id
        self.api_token = api_token
        self.base_url = SANDBOX_BASE_URL if use_sandbox else PROD_BASE_URL
        self.timeout = timeout
        self.max_retries = max_retries
        self.backoff_seconds = float(backoff_seconds)
        self.rate_limiter = RequestRateLimiter(requests_per_minute)
        self.log = log_callback or (lambda _msg: None)

    def _headers(self):
        auth_raw = f"{self.merchant_id}:{self.api_token}".encode("ascii")
        auth_basic = base64.b64encode(auth_raw).decode("ascii")
        return {
            "Authorization": f"Basic {auth_basic}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "x-merchant-id": self.merchant_id,
            "User-Agent": "CDON-Statuses-SKU-GUI/1.0",
        }

    def fetch_statuses_by_skus(self, skus):
        url = f"{self.base_url}{STATUS_BY_SKU_ENDPOINT}"
        payload = {"skus": skus}
        transient_http_codes = {429, 500, 502, 503, 504}

        for attempt in range(self.max_retries + 1):
            self.rate_limiter.wait_turn()
            try:
                response = requests.post(url, headers=self._headers(), json=payload, timeout=self.timeout)
            except (requests.Timeout, requests.ConnectionError) as exc:
                if attempt >= self.max_retries:
                    raise RuntimeError(f"Blad polaczenia po {attempt + 1} probach: {exc}") from exc
                wait_time = min(90.0, self.backoff_seconds * (2**attempt))
                self.log(f"Polaczenie nieudane. Retry {attempt + 1}/{self.max_retries} za {wait_time:.1f}s")
                time.sleep(wait_time)
                continue

            if response.status_code == 200:
                try:
                    return response.json()
                except json.JSONDecodeError as exc:
                    raise RuntimeError("API zwrocilo niepoprawny JSON") from exc

            if response.status_code in transient_http_codes and attempt < self.max_retries:
                retry_after_header = response.headers.get("Retry-After", "").strip()
                wait_time = None
                if retry_after_header:
                    try:
                        wait_time = float(retry_after_header)
                    except ValueError:
                        wait_time = None

                if wait_time is None:
                    wait_time = min(90.0, self.backoff_seconds * (2**attempt))

                if response.status_code == 429:
                    self.log(
                        f"HTTP 429 (limit). Retry {attempt + 1}/{self.max_retries} za {wait_time:.1f}s"
                    )
                else:
                    self.log(
                        f"HTTP {response.status_code}. Retry {attempt + 1}/{self.max_retries} za {wait_time:.1f}s"
                    )

                time.sleep(wait_time)
                continue

            text = response.text[:500]
            raise RuntimeError(f"HTTP {response.status_code}: {text}")

        raise RuntimeError("Nie udalo sie pobrac statusow po wielu probach")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")

        self.title(APP_TITLE)
        self.geometry("1100x760")
        self.minsize(980, 660)
        self.configure(fg_color=BG_MAIN)

        self.accounts_path_var = ctk.StringVar(value=DEFAULT_ACCOUNTS_PATH)
        self.account_var = ctk.StringVar(value="")
        self.skus_csv_var = ctk.StringVar(value="")
        self.output_csv_var = ctk.StringVar(value="")
        self.use_sandbox_var = ctk.BooleanVar(value=False)
        self.batch_size_var = ctk.StringVar(value="200")
        self.rate_limit_var = ctk.StringVar(value=str(DEFAULT_REQUESTS_PER_MINUTE))

        self.accounts = {}
        self.is_running = False

        self._build_ui()
        self._load_accounts()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color=BG_PANEL, border_color=BORDER, border_width=1)
        header.grid(row=0, column=0, padx=18, pady=(18, 10), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text=APP_TITLE,
            text_color=TEXT,
            font=ctk.CTkFont(size=24, weight="bold"),
        ).grid(row=0, column=0, padx=16, pady=(14, 2), sticky="w")

        ctk.CTkLabel(
            header,
            text="Pobieranie statusow ofert CDON na podstawie SKU z pliku CSV.",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=13),
        ).grid(row=1, column=0, padx=16, pady=(0, 14), sticky="w")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(2, weight=1)

        config = ctk.CTkFrame(content, fg_color=BG_PANEL, border_color=BORDER, border_width=1)
        config.grid(row=0, column=0, sticky="ew")
        config.grid_columnconfigure(1, weight=1)

        row = 0
        ctk.CTkLabel(config, text="Plik kont", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="w")
        ctk.CTkEntry(config, textvariable=self.accounts_path_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=row, column=1, padx=8, pady=10, sticky="ew"
        )
        ctk.CTkButton(
            config,
            text="Wybierz",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._pick_accounts,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        ctk.CTkLabel(config, text="Konto", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="w")
        self.account_menu = ctk.CTkOptionMenu(
            config,
            variable=self.account_var,
            values=["Brak kont"],
            fg_color=ACCENT,
            button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            dropdown_fg_color=BG_PANEL,
            dropdown_text_color=TEXT,
            text_color="#ffffff",
        )
        self.account_menu.grid(row=row, column=1, padx=8, pady=10, sticky="ew")
        ctk.CTkButton(
            config,
            text="Odswiez",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._load_accounts,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        ctk.CTkLabel(config, text="Plik SKU CSV", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="w")
        ctk.CTkEntry(config, textvariable=self.skus_csv_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=row, column=1, padx=8, pady=10, sticky="ew"
        )
        ctk.CTkButton(
            config,
            text="Wybierz",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._pick_skus_csv,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        ctk.CTkLabel(config, text="Plik wynikowy CSV", text_color=TEXT).grid(row=row, column=0, padx=14, pady=10, sticky="w")
        ctk.CTkEntry(config, textvariable=self.output_csv_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=row, column=1, padx=8, pady=10, sticky="ew"
        )
        ctk.CTkButton(
            config,
            text="Wybierz",
            width=90,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            command=self._pick_output_csv,
        ).grid(row=row, column=2, padx=(8, 14), pady=10)

        row += 1
        options = ctk.CTkFrame(config, fg_color="transparent")
        options.grid(row=row, column=0, columnspan=3, padx=14, pady=(2, 12), sticky="ew")
        options.grid_columnconfigure(3, weight=1)

        ctk.CTkCheckBox(
            options,
            text="Sandbox",
            variable=self.use_sandbox_var,
            text_color=TEXT,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            border_color=BORDER,
        ).grid(row=0, column=0, padx=(0, 14), pady=4, sticky="w")

        ctk.CTkLabel(options, text="Batch size", text_color=TEXT).grid(row=0, column=1, padx=(0, 8), sticky="w")
        ctk.CTkEntry(options, width=80, textvariable=self.batch_size_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=0, column=2, padx=(0, 14), sticky="w"
        )

        ctk.CTkLabel(options, text="Limit req/min", text_color=TEXT).grid(row=0, column=3, padx=(0, 8), sticky="w")
        ctk.CTkEntry(options, width=90, textvariable=self.rate_limit_var, fg_color=BG_INPUT, text_color=TEXT).grid(
            row=0, column=4, padx=(0, 0), sticky="w"
        )

        controls = ctk.CTkFrame(content, fg_color="transparent")
        controls.grid(row=1, column=0, pady=(10, 10), sticky="ew")
        controls.grid_columnconfigure(2, weight=1)

        self.start_btn = ctk.CTkButton(
            controls,
            text="Pobierz statusy",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="#ffffff",
            width=180,
            command=self._start,
        )
        self.start_btn.grid(row=0, column=0, padx=(0, 8), sticky="w")

        self.stop_btn = ctk.CTkButton(
            controls,
            text="Wyczysc log",
            fg_color="#f2c7de",
            hover_color="#e9b5d1",
            text_color=TEXT,
            width=130,
            command=self._clear_log,
        )
        self.stop_btn.grid(row=0, column=1, padx=(0, 8), sticky="w")

        self.progress = ctk.CTkProgressBar(controls, progress_color=ACCENT, fg_color="#f4d5e5")
        self.progress.grid(row=0, column=2, sticky="ew")
        self.progress.set(0)

        log_panel = ctk.CTkFrame(content, fg_color=BG_PANEL, border_color=BORDER, border_width=1)
        log_panel.grid(row=2, column=0, sticky="nsew")
        log_panel.grid_columnconfigure(0, weight=1)
        log_panel.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_panel,
            text="Log",
            text_color=TEXT,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, padx=12, pady=(10, 6), sticky="w")

        self.log_box = ctk.CTkTextbox(
            log_panel,
            fg_color=BG_INPUT,
            text_color=TEXT,
            border_width=1,
            border_color=BORDER,
            wrap="word",
        )
        self.log_box.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")

    def _log(self, msg, color=None):
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {msg}\n"
        self.log_box.insert("end", line)
        self.log_box.see("end")

        if color:
            self.log_box.tag_add(color, "end-2l", "end-1l")
            self.log_box.tag_config(color, foreground=color)

    def _clear_log(self):
        self.log_box.delete("1.0", "end")

    def _pick_accounts(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik kont",
            filetypes=[("CSV", "*.csv"), ("Wszystkie", "*.*")],
        )
        if path:
            self.accounts_path_var.set(path)
            self._load_accounts()

    def _pick_skus_csv(self):
        path = filedialog.askopenfilename(
            title="Wybierz CSV z SKU",
            filetypes=[("CSV", "*.csv"), ("Wszystkie", "*.*")],
        )
        if path:
            self.skus_csv_var.set(path)
            if not self.output_csv_var.get().strip():
                stem = Path(path).stem
                out_path = str(Path(path).with_name(f"{stem}_statuses.csv"))
                self.output_csv_var.set(out_path)

    def _pick_output_csv(self):
        path = filedialog.asksaveasfilename(
            title="Zapisz wynik jako",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if path:
            self.output_csv_var.set(path)

    def _load_accounts(self):
        path = self.accounts_path_var.get().strip()
        self.accounts = {}

        if not path or not os.path.exists(path):
            self._set_account_menu(["Brak kont"])
            self._log("Nie znaleziono pliku kont.", ERROR)
            return

        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as handle:
                sample = handle.read(4096)
                handle.seek(0)
                delimiter = ";" if sample.count(";") >= sample.count(",") else ","
                reader = csv.DictReader(handle, delimiter=delimiter)

                for row in reader:
                    account_name = (row.get("Nazwa Konta") or row.get("account_name") or "").strip()
                    merchant_id = (row.get("MerchantID") or row.get("merchant_id") or "").strip()
                    api_token = (row.get("APIToken") or row.get("api_token") or "").strip()

                    if account_name and merchant_id and api_token:
                        self.accounts[account_name] = {
                            "merchant_id": merchant_id,
                            "api_token": api_token,
                        }

            if not self.accounts:
                self._set_account_menu(["Brak kont"])
                self._log("Plik kont wczytany, ale brak poprawnych rekordow.", ERROR)
                return

            names = sorted(self.accounts.keys())
            self._set_account_menu(names)
            self.account_var.set(names[0])
            self._log(f"Wczytano konta: {len(names)}", SUCCESS)
        except Exception as exc:
            self._set_account_menu(["Brak kont"])
            self._log(f"Blad czytania pliku kont: {exc}", ERROR)

    def _set_account_menu(self, values):
        self.account_menu.configure(values=values)
        if values:
            self.account_var.set(values[0])

    def _read_skus_csv(self, path):
        if not os.path.exists(path):
            raise FileNotFoundError("Nie znaleziono pliku SKU CSV")

        skus = []
        with open(path, "r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            delimiter = ";" if sample.count(";") >= sample.count(",") else ","

            reader = csv.DictReader(handle, delimiter=delimiter)
            headers = [h.strip() for h in (reader.fieldnames or [])]
            header_lc = {h.lower(): h for h in headers}

            sku_col = None
            for candidate in ["sku", "seller_sku", "item_sku", "id"]:
                if candidate in header_lc:
                    sku_col = header_lc[candidate]
                    break

            if sku_col is None:
                if headers:
                    sku_col = headers[0]
                else:
                    raise ValueError("CSV nie ma naglowkow")

            for row in reader:
                value = (row.get(sku_col) or "").strip()
                if value and value.lower() != "nan":
                    skus.append(value)

        # Usuniecie duplikatow z zachowaniem kolejnosci
        return list(dict.fromkeys(skus))

    def _flatten_statuses(self, account_name, response_json):
        rows = []
        statuses = response_json.get("statuses") or []

        for status_obj in statuses:
            base = {
                "account": account_name,
                "correlation_id": status_obj.get("correlation_id", ""),
                "article_id": status_obj.get("article_id", ""),
                "sku": status_obj.get("sku", ""),
                "action": status_obj.get("action", ""),
            }
            markets = status_obj.get("markets") or []

            if not markets:
                rows.append(
                    {
                        **base,
                        "market": "",
                        "status": "",
                        "error_code": "",
                    }
                )
                continue

            for market_info in markets:
                rows.append(
                    {
                        **base,
                        "market": market_info.get("market", ""),
                        "status": market_info.get("status", ""),
                        "error_code": market_info.get("error_code", ""),
                    }
                )

        return rows

    def _write_output_csv(self, path, rows):
        out_dir = os.path.dirname(path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        fieldnames = [
            "account",
            "correlation_id",
            "article_id",
            "sku",
            "action",
            "market",
            "status",
            "error_code",
        ]

        with open(path, "w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames, delimiter=";")
            writer.writeheader()
            writer.writerows(rows)

    def _validate(self):
        if self.is_running:
            return False

        account_name = self.account_var.get().strip()
        if not account_name or account_name not in self.accounts:
            messagebox.showerror("Blad", "Wybierz poprawne konto z listy.")
            return False

        skus_path = self.skus_csv_var.get().strip()
        if not skus_path:
            messagebox.showerror("Blad", "Wybierz plik CSV z SKU.")
            return False

        output_path = self.output_csv_var.get().strip()
        if not output_path:
            messagebox.showerror("Blad", "Wybierz plik wynikowy CSV.")
            return False

        try:
            batch_size = int(self.batch_size_var.get().strip())
            if batch_size < 1 or batch_size > 1000:
                raise ValueError
        except ValueError:
            messagebox.showerror("Blad", "Batch size musi byc liczba od 1 do 1000.")
            return False

        try:
            req_per_min = int(self.rate_limit_var.get().strip())
            if req_per_min < 1 or req_per_min > MAX_REQUESTS_PER_MINUTE:
                raise ValueError
        except ValueError:
            messagebox.showerror("Blad", f"Limit req/min musi byc liczba od 1 do {MAX_REQUESTS_PER_MINUTE}.")
            return False

        return True

    def _set_running(self, running):
        self.is_running = running
        state = "disabled" if running else "normal"
        self.start_btn.configure(state=state)

    def _start(self):
        if not self._validate():
            return

        self._set_running(True)
        self.progress.set(0)

        worker = threading.Thread(target=self._run_job, daemon=True)
        worker.start()

    def _run_job(self):
        try:
            account_name = self.account_var.get().strip()
            creds = self.accounts[account_name]
            skus_path = self.skus_csv_var.get().strip()
            output_path = self.output_csv_var.get().strip()
            batch_size = int(self.batch_size_var.get().strip())
            req_per_min = int(self.rate_limit_var.get().strip())
            use_sandbox = bool(self.use_sandbox_var.get())

            self._log(f"Konto: {account_name}")
            self._log(f"Srodowisko: {'SANDBOX' if use_sandbox else 'PRODUKCJA'}")
            self._log(f"Limit API: {req_per_min} req/min")
            self._log("Wczytywanie SKU...")
            skus = self._read_skus_csv(skus_path)
            if not skus:
                raise RuntimeError("Brak SKU do przetworzenia")

            self._log(f"Liczba SKU: {len(skus)}")

            client = CDONStatusesClient(
                merchant_id=creds["merchant_id"],
                api_token=creds["api_token"],
                use_sandbox=use_sandbox,
                requests_per_minute=req_per_min,
                max_retries=DEFAULT_MAX_RETRIES,
                backoff_seconds=DEFAULT_BACKOFF_SECONDS,
                log_callback=self._log,
            )

            all_rows = []
            all_found_skus = set()

            sku_batches = list(chunks(skus, batch_size))
            total = len(sku_batches)

            for idx, batch in enumerate(sku_batches, start=1):
                self._log(f"Batch {idx}/{total}: {len(batch)} SKU")
                response = client.fetch_statuses_by_skus(batch)
                rows = self._flatten_statuses(account_name, response)
                all_rows.extend(rows)

                for st in response.get("statuses") or []:
                    sku_val = st.get("sku")
                    if sku_val:
                        all_found_skus.add(str(sku_val))

                self.progress.set(idx / total)

            missing = [sku for sku in skus if sku not in all_found_skus]
            for sku in missing:
                all_rows.append(
                    {
                        "account": account_name,
                        "correlation_id": "",
                        "article_id": "",
                        "sku": sku,
                        "action": "",
                        "market": "",
                        "status": "not_returned",
                        "error_code": "",
                    }
                )

            self._write_output_csv(output_path, all_rows)
            self._log(f"Zapisano: {output_path}", SUCCESS)
            self._log(f"Rekordy: {len(all_rows)}", SUCCESS)

            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Gotowe",
                    f"Pobrano statusy dla {len(skus)} SKU.\nZapisano {len(all_rows)} rekordow.",
                ),
            )
        except Exception as exc:
            self._log(f"Blad: {exc}", ERROR)
            self.after(0, lambda: messagebox.showerror("Blad", str(exc)))
        finally:
            self.after(0, lambda: self._set_running(False))


if __name__ == "__main__":
    app = App()
    app.mainloop()
