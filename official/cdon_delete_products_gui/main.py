import csv
import base64
import json
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib import error, parse, request
import subprocess
import tempfile
import shutil
import time

DEFAULT_ACCOUNTS_PATH = Path(str((Path(__file__).parent / "accounts.csv").resolve()))


class CdonDeleteGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CDON - Usuwanie produktów po SKU")
        self.geometry("900x650")

        self.accounts = []
        self.account_map = {}

        self.accounts_path_var = tk.StringVar(value=str(DEFAULT_ACCOUNTS_PATH))
        self.selected_account_var = tk.StringVar()
        self.sku_csv_path_var = tk.StringVar()
        self.delimiter_var = tk.StringVar(value=",")
        self.sku_column_var = tk.StringVar(value="SKU")
        self.timeout_var = tk.StringVar(value="30")
        self.base_url_var = tk.StringVar(value="https://merchants-api.cdon.com/api")

        self._build_ui()
        self._load_accounts_initial()

    def _build_ui(self):
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        accounts_frame = ttk.LabelFrame(main, text="1) Konta CDON", padding=10)
        accounts_frame.pack(fill="x", padx=2, pady=6)

        ttk.Label(accounts_frame, text="Plik accounts.csv:").grid(row=0, column=0, sticky="w")
        ttk.Entry(accounts_frame, textvariable=self.accounts_path_var, width=80).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(accounts_frame, text="Wybierz...", command=self._choose_accounts_file).grid(row=0, column=2, padx=4)
        ttk.Button(accounts_frame, text="Wczytaj", command=self._load_accounts).grid(row=0, column=3, padx=4)

        ttk.Label(accounts_frame, text="Konto:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.account_combo = ttk.Combobox(
            accounts_frame,
            textvariable=self.selected_account_var,
            state="readonly",
            width=50
        )
        self.account_combo.grid(row=1, column=1, sticky="w", pady=(8, 0), padx=6)

        accounts_frame.columnconfigure(1, weight=1)

        sku_frame = ttk.LabelFrame(main, text="2) Plik z SKU do usunięcia", padding=10)
        sku_frame.pack(fill="x", padx=2, pady=6)

        ttk.Label(sku_frame, text="CSV z SKU:").grid(row=0, column=0, sticky="w")
        ttk.Entry(sku_frame, textvariable=self.sku_csv_path_var, width=80).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(sku_frame, text="Wybierz...", command=self._choose_sku_file).grid(row=0, column=2, padx=4)

        ttk.Label(sku_frame, text="Delimiter:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(sku_frame, textvariable=self.delimiter_var, width=8).grid(row=1, column=1, sticky="w", pady=(8, 0), padx=6)

        ttk.Label(sku_frame, text="Nazwa kolumny SKU:").grid(row=1, column=1, sticky="w", pady=(8, 0), padx=(110, 6))
        ttk.Entry(sku_frame, textvariable=self.sku_column_var, width=20).grid(row=1, column=1, sticky="w", pady=(8, 0), padx=(250, 6))

        ttk.Label(sku_frame, text="Timeout (s):").grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(sku_frame, textvariable=self.timeout_var, width=10).grid(row=1, column=2, sticky="e", pady=(8, 0))

        ttk.Label(sku_frame, text="API Base URL:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(sku_frame, textvariable=self.base_url_var, width=40).grid(row=2, column=1, sticky="w", pady=(8, 0), padx=6)

        sku_frame.columnconfigure(1, weight=1)

        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", padx=2, pady=8)

        self.start_btn = ttk.Button(action_frame, text="Usuń produkty", command=self._on_start)
        self.start_btn.pack(side="left")

        log_frame = ttk.LabelFrame(main, text="Log", padding=10)
        log_frame.pack(fill="both", expand=True, padx=2, pady=6)

        self.log_text = tk.Text(log_frame, wrap="word", height=20)
        self.log_text.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

    def _log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.update_idletasks()

    def _choose_accounts_file(self):
        path = filedialog.askopenfilename(
            title="Wybierz accounts.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.accounts_path_var.set(path)

    def _choose_sku_file(self):
        path = filedialog.askopenfilename(
            title="Wybierz CSV z SKU",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.sku_csv_path_var.set(path)

    def _load_accounts_initial(self):
        if DEFAULT_ACCOUNTS_PATH.exists():
            self._load_accounts()

    def _load_accounts(self):
        path = Path(self.accounts_path_var.get().strip())
        if not path.exists():
            messagebox.showerror("Błąd", f"Nie znaleziono pliku: {path}")
            return

        accounts = []
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f, delimiter=";")
                required = {"Nazwa Konta", "MerchantID", "APIToken"}
                if not reader.fieldnames or not required.issubset(set(reader.fieldnames)):
                    messagebox.showerror(
                        "Błąd",
                        "Plik accounts.csv musi zawierać kolumny: Nazwa Konta;MerchantID;APIToken"
                    )
                    return

                for row in reader:
                    name = (row.get("Nazwa Konta") or "").strip()
                    merchant_id = (row.get("MerchantID") or "").strip()
                    api_token = (row.get("APIToken") or "").strip()
                    if name and merchant_id and api_token:
                        accounts.append(
                            {
                                "name": name,
                                "merchant_id": merchant_id,
                                "api_token": api_token,
                            }
                        )
        except Exception as exc:
            messagebox.showerror("Błąd", f"Nie udało się wczytać kont: {exc}")
            return

        if not accounts:
            messagebox.showwarning("Uwaga", "Brak poprawnych kont w pliku.")
            return

        self.accounts = accounts
        self.account_map = {a["name"]: a for a in accounts}
        self.account_combo["values"] = [a["name"] for a in accounts]
        self.account_combo.current(0)
        self._log(f"Wczytano konta: {len(accounts)}")

    def _read_skus(self, file_path, delimiter, sku_column):
        skus = []
        with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f, delimiter=delimiter)
            if not reader.fieldnames:
                raise ValueError("Plik SKU jest pusty lub ma nieprawidłowy nagłówek.")

            if sku_column not in reader.fieldnames:
                raise ValueError(
                    f"Nie znaleziono kolumny '{sku_column}'. Dostępne: {', '.join(reader.fieldnames)}"
                )

            for row in reader:
                sku = (row.get(sku_column) or "").strip()
                if sku:
                    skus.append(sku)

        return list(dict.fromkeys(skus))

    def _delete_by_sku(self, base_url, merchant_id, api_token, sku, timeout):
        endpoint = f"{base_url.rstrip('/')}/v2/articles/bulk"
        payload = {
            "actions": [
                {"sku": sku, "action": "delete_article"}
            ]
        }

        credentials = f"{merchant_id}:{api_token}".encode("utf-8")
        basic_token = base64.b64encode(credentials).decode("ascii")

        # Try using system curl (avoids Python client fingerprint issues with Cloudflare/WAF)
        curl_path = shutil.which("curl")
        if curl_path:
            try:
                with tempfile.NamedTemporaryFile("w", encoding="utf-8", delete=False, suffix=".json") as tf:
                    json.dump(payload, tf, ensure_ascii=False)
                    tf.flush()
                    tmpname = tf.name

                cmd = [
                    curl_path,
                    "--silent",
                    "--show-error",
                    "--request",
                    "PUT",
                    endpoint,
                    "--header",
                    f"Authorization: Basic {basic_token}",
                    "--header",
                    "Content-Type: application/json",
                    "--header",
                    "Accept: application/json",
                    "--data-binary",
                    f"@{tmpname}",
                    "--write-out",
                    "\n%{http_code}"
                ]

                proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
                out = proc.stdout or ""
                # split off trailing HTTP status code
                if "\n" in out:
                    body, status_str = out.rsplit("\n", 1)
                else:
                    body, status_str = "", out

                try:
                    status = int(status_str.strip())
                except Exception:
                    status = None

                try:
                    Path(tmpname).unlink()
                except Exception:
                    pass

                if proc.returncode == 0 and status and 200 <= status < 300:
                    return True, status, body
                else:
                    return False, status, body or proc.stderr.decode("utf-8", errors="ignore") if isinstance(proc.stderr, bytes) else proc.stderr
            except subprocess.TimeoutExpired:
                return False, None, "curl timeout"
            except Exception as exc:
                # fall through to urllib fallback
                pass

        # Fallback to urllib if curl missing or fails
        req = request.Request(
            endpoint,
            data=json.dumps(payload).encode("utf-8"),
            method="PUT",
        )
        req.add_header("Authorization", f"Basic {basic_token}")
        req.add_header("Content-Type", "application/json")
        req.add_header("Accept", "application/json")
        # Spoof User-Agent to be more similar to common clients
        req.add_header("User-Agent", "curl/7.79.1")

        try:
            with request.urlopen(req, timeout=timeout) as resp:
                status = resp.getcode()
                body = resp.read().decode("utf-8", errors="ignore")
                return True, status, body
        except error.HTTPError as http_err:
            body = http_err.read().decode("utf-8", errors="ignore")
            return False, http_err.code, body
        except Exception as exc:
            return False, None, str(exc)

    def _on_start(self):
        if not self.accounts:
            messagebox.showwarning("Uwaga", "Najpierw wczytaj konta.")
            return

        account_name = self.selected_account_var.get().strip()
        if not account_name or account_name not in self.account_map:
            messagebox.showwarning("Uwaga", "Wybierz konto.")
            return

        sku_csv = self.sku_csv_path_var.get().strip()
        if not sku_csv or not Path(sku_csv).exists():
            messagebox.showwarning("Uwaga", "Wybierz poprawny plik CSV z SKU.")
            return

        delimiter = self.delimiter_var.get() or ","
        sku_column = self.sku_column_var.get().strip() or "SKU"

        try:
            timeout = int(self.timeout_var.get().strip())
            if timeout <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Uwaga", "Timeout musi być dodatnią liczbą całkowitą.")
            return

        self.start_btn.config(state="disabled")
        thread = threading.Thread(
            target=self._run_delete_process,
            args=(account_name, sku_csv, delimiter, sku_column, timeout),
            daemon=True,
        )
        thread.start()

    def _run_delete_process(self, account_name, sku_csv, delimiter, sku_column, timeout):
        account = self.account_map[account_name]
        merchant_id = account["merchant_id"]
        api_token = account["api_token"]
        base_url = self.base_url_var.get().strip() or "https://merchants-api.cdon.com/api"

        self._log("=" * 80)
        self._log(f"Konto: {account_name}")
        self._log(f"MerchantID: {merchant_id}")
        self._log(f"Plik SKU: {sku_csv}")

        try:
            skus = self._read_skus(sku_csv, delimiter, sku_column)
        except Exception as exc:
            self._log(f"[BŁĄD] Nie udało się wczytać SKU: {exc}")
            self.start_btn.config(state="normal")
            return

        if not skus:
            self._log("[INFO] Brak SKU do usunięcia.")
            self.start_btn.config(state="normal")
            return

        self._log(f"SKU do usunięcia: {len(skus)}")

        success = 0
        failed = 0

        for idx, sku in enumerate(skus, start=1):
            ok, status, body = self._delete_by_sku(
                base_url=base_url,
                merchant_id=merchant_id,
                api_token=api_token,
                sku=sku,
                timeout=timeout,
            )

            if ok:
                success += 1
                self._log(f"[{idx}/{len(skus)}] OK    SKU={sku} status={status}")
            else:
                failed += 1
                short_body = body[:300].replace("\n", " ") if body else ""
                self._log(f"[{idx}/{len(skus)}] FAIL  SKU={sku} status={status} msg={short_body}")

        self._log("-" * 80)
        self._log(f"Zakończono. Sukces: {success}, Błędy: {failed}, Razem: {len(skus)}")
        self._log("=" * 80)
        self.start_btn.config(state="normal")


if __name__ == "__main__":
    app = CdonDeleteGui()
    app.mainloop()
