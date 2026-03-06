from __future__ import annotations

import json
import math
import threading
import time
import base64
from dataclasses import dataclass
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any

import customtkinter as ctk
import pandas as pd
import requests
import yaml
from tkinter import filedialog, messagebox


APP_TITLE = "BOL Offers Creator (MVP)"
DEFAULT_API_URL = "https://api.bol.com/retailer/offers"
DEFAULT_AUTH_URL = "https://login.bol.com/token"
DEFAULT_SPEC_PATH = Path(r"C:\Users\super\Downloads\offers-v11 (1).yaml")
DEFAULT_TEMPLATE_XLSX_PATH = Path(
    r"C:\Users\super\Desktop\skrypty_py\product-enrichment_v0.8\product-enrichment_v0.12\data\furnitex_maxdywanik.xlsx"
)

ACCENT = "#ff69b4"
ACCENT_HOVER = "#e754a7"
APP_BG = "#fff7fb"
PANEL_BG = "#ffffff"
INPUT_BG = "#fff2f8"
BORDER = "#f7b3d2"
TEXT = "#3d2130"
MUTED = "#7d5a6b"
ERROR = "#b42318"
SUCCESS = "#2f7d4f"

SEND_RATE_LIMIT_PER_SEC = 40
SEND_MAX_WORKERS = 12
SEND_MAX_RETRIES = 4
SEND_REQUEST_TIMEOUT = 20


@dataclass(slots=True)
class ColumnMapping:
    ean: str = "EAN"
    price: str = "price"
    stock: str = "stock"
    reference: str = "id"
    on_hold: str = "on_hold"


@dataclass(slots=True)
class OfferDefaults:
    condition_category: str = "NEW"
    fulfilment_method: str = "FBR"
    fulfilment_schedule: str = "BOL_DELIVERY_PROMISE"
    min_days_to_customer: int = 3
    max_days_to_customer: int = 5
    country_codes: tuple[str, str] = ("NL", "BE")
    managed_by_retailer: bool = True
    on_hold_by_retailer: bool = False


class OpenApiSpecInspector:
    def __init__(self, spec_path: Path):
        self.spec_path = spec_path
        self.required_fields: list[str] = []
        self.condition_options: list[str] = ["NEW", "SECONDHAND", "REFURBISHED"]
        self.fulfilment_options: list[str] = ["FBR", "FBB"]

    def load(self) -> None:
        with self.spec_path.open("r", encoding="utf-8") as handle:
            content = yaml.safe_load(handle)

        schemas = content.get("components", {}).get("schemas", {})
        create_offer = schemas.get("CreateOfferRequest", {})
        self.required_fields = create_offer.get("required", [])

        condition_enum = schemas.get("ConditionCategory", {}).get("enum", [])
        fulfilment_enum = schemas.get("Fulfilment", {}).get("properties", {}).get("method", {}).get("enum", [])

        if condition_enum:
            self.condition_options = condition_enum
        if fulfilment_enum:
            self.fulfilment_options = fulfilment_enum


class RateLimiter:
    def __init__(self, rate_per_sec: int):
        self.rate_per_sec = rate_per_sec
        self.lock = threading.Lock()
        self.calls = deque()

    def acquire(self) -> None:
        while True:
            with self.lock:
                now = time.monotonic()
                while self.calls and now - self.calls[0] >= 1:
                    self.calls.popleft()

                if len(self.calls) < self.rate_per_sec:
                    self.calls.append(now)
                    return

                sleep_for = 1 - (now - self.calls[0])
            time.sleep(max(sleep_for, 0.001))


class OAuthTokenManager:
    def __init__(self, auth_url: str, client_id: str, client_secret: str):
        self.auth_url = auth_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.lock = threading.Lock()
        self.token: str | None = None
        self.expires_in = 0
        self.token_time = 0.0

    def get_token(self) -> str:
        with self.lock:
            now = time.time()
            if self.token and now < self.token_time + self.expires_in - 30:
                return self.token

            auth_raw = f"{self.client_id}:{self.client_secret}".encode("utf-8")
            auth_encoded = base64.b64encode(auth_raw).decode("utf-8")

            headers = {
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
                "Authorization": f"Basic {auth_encoded}",
            }

            response = requests.post(
                self.auth_url,
                params={"grant_type": "client_credentials"},
                headers=headers,
                timeout=15,
            )

            if response.status_code != 200:
                raise ValueError(f"Błąd pobierania tokena HTTP {response.status_code}: {response.text[:250]}")

            payload = response.json()
            access_token = payload.get("access_token")
            expires_in = int(payload.get("expires_in", 0))
            if not access_token:
                raise ValueError("Brak access_token w odpowiedzi OAuth")

            self.token = access_token
            self.expires_in = expires_in if expires_in > 0 else 300
            self.token_time = now
            return self.token


class OfferPayloadBuilder:
    @staticmethod
    def _normalize_ean(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float):
            if pd.isna(value):
                return ""
            if value.is_integer():
                return str(int(value))
        value_str = str(value).strip()
        if not value_str or value_str.lower() == "nan":
            return ""
        if value_str.endswith(".0") and value_str.replace(".", "", 1).isdigit():
            return value_str[:-2]
        return value_str

    @staticmethod
    def _to_bool(value: Any, fallback: bool) -> bool:
        if value is None:
            return fallback
        if isinstance(value, bool):
            return value
        lowered = str(value).strip().lower()
        if lowered in {"1", "true", "tak", "yes", "y"}:
            return True
        if lowered in {"0", "false", "nie", "no", "n"}:
            return False
        return fallback

    @staticmethod
    def _to_int(value: Any, fallback: int = 0) -> int:
        if value is None or str(value).strip() == "":
            return fallback
        return int(float(value))

    @staticmethod
    def _to_float(value: Any, fallback: float = 0.0) -> float:
        if value is None or str(value).strip() == "":
            return fallback
        if isinstance(value, str):
            value = value.replace(",", ".")
        return float(value)

    @staticmethod
    def _round_up_to_49_or_99(value: float) -> float:
        if value <= 0:
            return 0.0
        cents = int(math.ceil(value * 100))
        whole = cents // 100
        fractional = cents % 100

        if fractional <= 49:
            target_cents = whole * 100 + 49
        else:
            target_cents = whole * 100 + 99

        return target_cents / 100

    @staticmethod
    def build_payload(row: dict[str, Any], mapping: ColumnMapping, defaults: OfferDefaults) -> dict[str, Any]:
        ean_value = OfferPayloadBuilder._normalize_ean(row.get(mapping.ean, ""))
        if not ean_value:
            raise ValueError("Brak EAN w wierszu")

        condition_category = defaults.condition_category
        fulfilment_method = defaults.fulfilment_method
        price = OfferPayloadBuilder._to_float(row.get(mapping.price), fallback=0.0)
        if price <= 0:
            raise ValueError("Cena musi być > 0")

        main_price = OfferPayloadBuilder._round_up_to_49_or_99(price)
        bundle_prices = [
            {
                "quantity": 1,
                "unitPrice": main_price,
            }
        ]
        if main_price < 30:
            bundle_prices.extend(
                [
                    {
                        "quantity": 2,
                        "unitPrice": main_price * 0.98,
                    },
                    {
                        "quantity": 3,
                        "unitPrice": main_price * 0.97,
                    },
                    {
                        "quantity": 4,
                        "unitPrice": main_price * 0.96,
                    },
                ]
            )

        stock_amount = OfferPayloadBuilder._to_int(row.get(mapping.stock), fallback=0)
        on_hold = OfferPayloadBuilder._to_bool(row.get(mapping.on_hold), defaults.on_hold_by_retailer)

        fulfilment_payload: dict[str, Any] = {
            "method": fulfilment_method,
            "schedule": defaults.fulfilment_schedule,
        }
        if fulfilment_method == "FBR":
            fulfilment_payload["deliveryPromise"] = {
                "minimumDaysToCustomer": defaults.min_days_to_customer,
                "maximumDaysToCustomer": defaults.max_days_to_customer,
            }

        payload: dict[str, Any] = {
            "ean": ean_value,
            "condition": {"category": condition_category},
            "pricing": {
                "bundlePrices": bundle_prices
            },
            "fulfilment": fulfilment_payload,
            "countryAvailabilities": [{"countryCode": code} for code in defaults.country_codes],
            "onHoldByRetailer": on_hold,
        }

        reference = str(row.get(mapping.reference, "")).strip()
        if reference:
            payload["reference"] = reference

        payload["stock"] = {
            "amount": stock_amount,
            "managedByRetailer": defaults.managed_by_retailer,
        }

        return payload


class OffersCreatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")

        self.title(APP_TITLE)
        self.geometry("980x760")
        self.minsize(920, 680)
        self.configure(fg_color=APP_BG)

        self.xlsx_path: Path | None = None
        self.spec_path: Path | None = DEFAULT_SPEC_PATH if DEFAULT_SPEC_PATH.exists() else None
        self.payloads: list[dict[str, Any]] = []
        self.columns: list[str] = []

        self._build_ui()
        if self.spec_path:
            self.spec_entry.insert(0, str(self.spec_path))
            self.load_spec()

    def _build_ui(self) -> None:
        root = ctk.CTkFrame(self, fg_color=PANEL_BG, corner_radius=14)
        root.pack(fill="both", expand=True, padx=16, pady=16)

        title = ctk.CTkLabel(
            root,
            text="Tworzenie ofert BOL z XLSX (zarys MVP)",
            text_color=TEXT,
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.pack(anchor="w", padx=18, pady=(16, 4))

        subtitle = ctk.CTkLabel(
            root,
            text="Etap 1: ładowanie XLSX, budowa payloadów v11 i opcjonalne wysyłanie POST /retailer/offers. Warunki stałe: condition=NEW, fulfilment=FBR.",
            text_color=MUTED,
            font=ctk.CTkFont(size=13),
        )
        subtitle.pack(anchor="w", padx=18, pady=(0, 10))

        files_frame = ctk.CTkFrame(root, fg_color="transparent")
        files_frame.pack(fill="x", padx=18, pady=8)

        self.xlsx_entry = self._build_file_row(
            files_frame,
            label="Plik XLSX",
            button_text="Wybierz XLSX",
            row=0,
            command=self.select_xlsx,
        )
        self.spec_entry = self._build_file_row(
            files_frame,
            label="Spec YAML (offers-v11)",
            button_text="Wybierz YAML",
            row=1,
            command=self.select_spec,
        )

        api_frame = ctk.CTkFrame(root, fg_color=INPUT_BG, corner_radius=10)
        api_frame.pack(fill="x", padx=18, pady=(6, 10))
        api_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(api_frame, text="API URL", text_color=TEXT).grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.api_entry = ctk.CTkEntry(api_frame, fg_color="white", border_color=BORDER, text_color=TEXT)
        self.api_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=10)
        self.api_entry.insert(0, DEFAULT_API_URL)

        ctk.CTkLabel(api_frame, text="Auth URL", text_color=TEXT).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        self.auth_entry = ctk.CTkEntry(api_frame, fg_color="white", border_color=BORDER, text_color=TEXT)
        self.auth_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=(0, 10))
        self.auth_entry.insert(0, DEFAULT_AUTH_URL)

        ctk.CTkLabel(api_frame, text="Client ID", text_color=TEXT).grid(row=2, column=0, sticky="w", padx=10, pady=(0, 10))
        self.client_id_entry = ctk.CTkEntry(api_frame, fg_color="white", border_color=BORDER, text_color=TEXT)
        self.client_id_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=(0, 10))

        ctk.CTkLabel(api_frame, text="Client Secret", text_color=TEXT).grid(row=3, column=0, sticky="w", padx=10, pady=(0, 10))
        self.client_secret_entry = ctk.CTkEntry(api_frame, fg_color="white", border_color=BORDER, text_color=TEXT, show="*")
        self.client_secret_entry.grid(row=3, column=1, sticky="ew", padx=8, pady=(0, 10))

        self.required_label = ctk.CTkLabel(root, text="Wymagane pola z YAML: —", text_color=MUTED, wraplength=900, justify="left")
        self.required_label.pack(anchor="w", padx=18, pady=(0, 8))

        self.fixed_rules_label = ctk.CTkLabel(
            root,
            text="Stałe reguły: condition = NEW | fulfilment.method = FBR (zawsze przez nas) | schedule = BOL_DELIVERY_PROMISE | kraje: NL + BE",
            text_color=MUTED,
            wraplength=900,
            justify="left",
        )
        self.fixed_rules_label.pack(anchor="w", padx=18, pady=(0, 8))

        map_frame = ctk.CTkFrame(root, fg_color=INPUT_BG, corner_radius=10)
        map_frame.pack(fill="x", padx=18, pady=8)
        map_frame.grid_columnconfigure((1, 3), weight=1)

        self.map_entries: dict[str, ctk.CTkEntry] = {}
        mapping_fields = [
            ("ean", "Kolumna EAN"),
            ("price", "Kolumna cena"),
            ("stock", "Kolumna stan"),
            ("reference", "Kolumna reference"),
            ("on_hold", "Kolumna on_hold"),
        ]

        defaults = ColumnMapping()
        for index, (key, caption) in enumerate(mapping_fields):
            row = index // 2
            base_col = (index % 2) * 2
            ctk.CTkLabel(map_frame, text=caption, text_color=TEXT).grid(row=row, column=base_col, sticky="w", padx=10, pady=8)
            entry = ctk.CTkEntry(map_frame, fg_color="white", border_color=BORDER, text_color=TEXT)
            entry.grid(row=row, column=base_col + 1, sticky="ew", padx=8, pady=8)
            entry.insert(0, getattr(defaults, key))
            self.map_entries[key] = entry

        action_row = ctk.CTkFrame(root, fg_color="transparent")
        action_row.pack(fill="x", padx=18, pady=(10, 8))

        self.btn_load_xlsx = self._action_button(action_row, "Wczytaj XLSX", self.load_xlsx_preview)
        self.btn_save = self._action_button(action_row, "Zapisz JSON", self.save_payloads)
        self.btn_send = self._action_button(action_row, "One click processing (XLSX -> API)", self.send_payloads)

        self.progress_bar = ctk.CTkProgressBar(root, progress_color=ACCENT)
        self.progress_bar.pack(fill="x", padx=18, pady=(6, 6))
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(root, text="Gotowe.", text_color=MUTED)
        self.status_label.pack(anchor="w", padx=18, pady=(0, 8))

        self.log_box = ctk.CTkTextbox(root, fg_color="white", border_color=BORDER, text_color=TEXT)
        self.log_box.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        if DEFAULT_TEMPLATE_XLSX_PATH.exists():
            self.xlsx_path = DEFAULT_TEMPLATE_XLSX_PATH
            self.xlsx_entry.insert(0, str(DEFAULT_TEMPLATE_XLSX_PATH))

    def _build_file_row(self, parent: ctk.CTkFrame, label: str, button_text: str, row: int, command) -> ctk.CTkEntry:
        parent.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(parent, text=label, text_color=TEXT).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=6)
        entry = ctk.CTkEntry(parent, fg_color="white", border_color=BORDER, text_color=TEXT)
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        ctk.CTkButton(
            parent,
            text=button_text,
            command=command,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="white",
            width=140,
        ).grid(row=row, column=2, sticky="e", padx=(10, 0), pady=6)
        return entry

    def _action_button(self, parent: ctk.CTkFrame, text: str, command) -> ctk.CTkButton:
        button = ctk.CTkButton(
            parent,
            text=text,
            command=command,
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            text_color="white",
            height=36,
        )
        button.pack(side="left", padx=(0, 10))
        return button

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")

        def _update() -> None:
            self.log_box.insert("end", f"[{timestamp}] {message}\n")
            self.log_box.see("end")

        self.after(0, _update)

    def _set_status(self, text: str, is_error: bool = False) -> None:
        self.after(0, lambda: self.status_label.configure(text=text, text_color=ERROR if is_error else SUCCESS))

    def _set_progress(self, value: float) -> None:
        safe_value = min(max(value, 0.0), 1.0)
        self.after(0, lambda: self.progress_bar.set(safe_value))

    def select_xlsx(self) -> None:
        path = filedialog.askopenfilename(
            title="Wybierz plik XLSX",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self.xlsx_path = Path(path)
        self.xlsx_entry.delete(0, "end")
        self.xlsx_entry.insert(0, str(self.xlsx_path))
        self._log(f"Wybrano XLSX: {self.xlsx_path}")

    def select_spec(self) -> None:
        path = filedialog.askopenfilename(
            title="Wybierz specyfikację YAML",
            filetypes=[("YAML", "*.yaml *.yml"), ("All files", "*.*")],
        )
        if not path:
            return
        self.spec_path = Path(path)
        self.spec_entry.delete(0, "end")
        self.spec_entry.insert(0, str(self.spec_path))
        self.load_spec()

    def load_spec(self) -> None:
        if not self.spec_path or not self.spec_path.exists():
            self.required_label.configure(text="Wymagane pola z YAML: brak pliku spec")
            return

        try:
            inspector = OpenApiSpecInspector(self.spec_path)
            inspector.load()
            required = ", ".join(inspector.required_fields) if inspector.required_fields else "brak"
            self.required_label.configure(text=f"Wymagane pola z YAML: {required}")
            self._log(f"Wczytano spec: {self.spec_path.name}")
        except Exception as exc:
            self.required_label.configure(text="Wymagane pola z YAML: błąd odczytu")
            self._log(f"Błąd odczytu spec YAML: {exc}")
            self._set_status(f"Błąd YAML: {exc}", is_error=True)

    def load_xlsx_preview(self) -> None:
        if not self.xlsx_path or not self.xlsx_path.exists():
            messagebox.showwarning("Brak XLSX", "Wybierz plik XLSX.")
            return

        try:
            df = pd.read_excel(self.xlsx_path)
            self.columns = [str(column) for column in df.columns]
            self._log(f"Załadowano XLSX: {self.xlsx_path.name} | wiersze: {len(df)} | kolumny: {len(self.columns)}")
            self._log(f"Kolumny: {', '.join(self.columns[:20])}{' ...' if len(self.columns) > 20 else ''}")
            self._set_status("XLSX załadowany poprawnie.")
        except Exception as exc:
            self._set_status(f"Błąd XLSX: {exc}", is_error=True)
            self._log(f"Błąd odczytu XLSX: {exc}")

    def _build_mapping(self) -> ColumnMapping:
        return ColumnMapping(
            ean=self.map_entries["ean"].get().strip(),
            price=self.map_entries["price"].get().strip(),
            stock=self.map_entries["stock"].get().strip(),
            reference=self.map_entries["reference"].get().strip(),
            on_hold=self.map_entries["on_hold"].get().strip(),
        )

    def _generate_payloads_internal(self, preview_prices: bool = True, preview_first_payload: bool = True) -> bool:
        if not self.xlsx_path or not self.xlsx_path.exists():
            messagebox.showwarning("Brak XLSX", "Wybierz plik XLSX.")
            return False

        try:
            df = pd.read_excel(self.xlsx_path)
        except Exception as exc:
            self._set_status(f"Nie można odczytać XLSX: {exc}", is_error=True)
            return False

        mapping = self._build_mapping()
        defaults = OfferDefaults()

        payloads: list[dict[str, Any]] = []
        errors = 0
        price_preview_count = 0

        total = len(df)
        self._set_progress(0)

        for index, row in enumerate(df.to_dict(orient="records"), start=1):
            try:
                payload = OfferPayloadBuilder.build_payload(row, mapping, defaults)
                payloads.append(payload)

                if preview_prices and price_preview_count < 5:
                    raw_price = OfferPayloadBuilder._to_float(row.get(mapping.price), fallback=0.0)
                    bundle_prices = payload.get("pricing", {}).get("bundlePrices", [])
                    if bundle_prices:
                        tiers_text = " | ".join(
                            f"q{tier.get('quantity')}={tier.get('unitPrice')}" for tier in bundle_prices
                        )
                        self._log(f"Podgląd ceny: wejście={raw_price} -> {tiers_text}")
                        price_preview_count += 1
            except Exception as exc:
                errors += 1
                if errors <= 10:
                    self._log(f"Wiersz {index}: pominięty ({exc})")

            if total > 0:
                self._set_progress(index / total)

        self.payloads = payloads
        self._log(f"Wygenerowano payloadów: {len(payloads)} | pominięte: {errors}")

        if preview_first_payload and payloads:
            preview = json.dumps(payloads[0], ensure_ascii=False, indent=2)
            self._log("Podgląd pierwszego payloadu:")
            self._log(preview)

        if not payloads:
            self._set_status("Brak poprawnych payloadów do przetworzenia.", is_error=True)
            return False

        if errors > 0:
            self._set_status(f"Gotowe z ostrzeżeniami. Payloady: {len(payloads)} | Błędy: {errors}", is_error=True)
        else:
            self._set_status(f"Gotowe. Wygenerowano {len(payloads)} payloadów.")

        return True

    def generate_payloads(self) -> None:
        self._generate_payloads_internal(preview_prices=True, preview_first_payload=True)

    def save_payloads(self) -> None:
        if not self.payloads:
            self._log("Brak payloadów w pamięci. Automatyczne generowanie z XLSX przed zapisem JSON.")
            if not self._generate_payloads_internal(preview_prices=False, preview_first_payload=False):
                return

        output_path = filedialog.asksaveasfilename(
            title="Zapisz payloady JSON",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="offers_payloads.json",
        )
        if not output_path:
            return

        try:
            with open(output_path, "w", encoding="utf-8") as handle:
                json.dump(self.payloads, handle, ensure_ascii=False, indent=2)
            self._log(f"Zapisano plik JSON: {output_path}")
            self._set_status("Payloady zapisane.")
        except Exception as exc:
            self._set_status(f"Błąd zapisu JSON: {exc}", is_error=True)

    def send_payloads(self) -> None:
        self._log("One click processing: start automatycznego budowania payloadów z XLSX.")
        if not self._generate_payloads_internal(preview_prices=False, preview_first_payload=False):
            return

        client_id = self.client_id_entry.get().strip()
        client_secret = self.client_secret_entry.get().strip()
        if not client_id or not client_secret:
            messagebox.showwarning("Brak danych API", "Wpisz Client ID i Client Secret.")
            return

        api_url = self.api_entry.get().strip() or DEFAULT_API_URL
        auth_url = self.auth_entry.get().strip() or DEFAULT_AUTH_URL

        worker = threading.Thread(target=self._send_worker, args=(api_url, auth_url, client_id, client_secret), daemon=True)
        worker.start()

    def _send_worker(self, api_url: str, auth_url: str, client_id: str, client_secret: str) -> None:
        ok_count = 0
        fail_count = 0
        total = len(self.payloads)

        limiter = RateLimiter(SEND_RATE_LIMIT_PER_SEC)
        token_manager = OAuthTokenManager(auth_url=auth_url, client_id=client_id, client_secret=client_secret)
        self._set_progress(0)
        self._log(f"Start wysyłki: {total} ofert -> {api_url} | limit {SEND_RATE_LIMIT_PER_SEC}/s")

        def _send_single(payload: dict[str, Any]) -> tuple[bool, str]:
            offer_ref = str(payload.get("reference") or payload.get("ean") or "brak_ref")
            for attempt in range(1, SEND_MAX_RETRIES + 1):
                limiter.acquire()
                try:
                    token = token_manager.get_token()
                    headers = {
                        "Accept": "application/vnd.retailer.v11+json",
                        "Content-Type": "application/vnd.retailer.v11+json",
                        "Authorization": f"Bearer {token}",
                    }
                    response = requests.post(api_url, headers=headers, json=payload, timeout=SEND_REQUEST_TIMEOUT)
                    response_body = response.text if response.text else "<empty>"
                    self._log(
                        f"API odpowiedź | ref={offer_ref} | próba {attempt}/{SEND_MAX_RETRIES} | "
                        f"HTTP {response.status_code} | {response_body}"
                    )

                    if response.status_code in (200, 201):
                        return True, f"OK | ref={offer_ref} | HTTP {response.status_code}"

                    if response.status_code == 429:
                        retry_after_raw = response.headers.get("Retry-After")
                        reset_raw = response.headers.get("x-ratelimit-reset")
                        try:
                            retry_after = int(float(retry_after_raw)) if retry_after_raw is not None else 1
                        except (TypeError, ValueError):
                            retry_after = 1
                        try:
                            reset_after = int(float(reset_raw)) if reset_raw is not None else 0
                        except (TypeError, ValueError):
                            reset_after = 0

                        sleep_seconds = max(retry_after, reset_after + 1, 1)
                        if attempt < SEND_MAX_RETRIES:
                            time.sleep(sleep_seconds)
                            continue
                        return False, f"HTTP 429 po retry | ref={offer_ref}"

                    if 500 <= response.status_code <= 599 and attempt < SEND_MAX_RETRIES:
                        time.sleep(min(2 ** attempt, 8))
                        continue

                    return False, f"HTTP {response.status_code} | ref={offer_ref}"
                except Exception as exc:
                    exc_text = str(exc)
                    if "token" in exc_text.lower() and attempt < SEND_MAX_RETRIES:
                        time.sleep(1)
                        continue
                    if attempt < SEND_MAX_RETRIES:
                        time.sleep(min(2 ** attempt, 8))
                        continue
                    return False, f"Błąd request | ref={offer_ref}: {exc}"

            return False, "Nieznany błąd wysyłki"

        completed = 0
        with ThreadPoolExecutor(max_workers=SEND_MAX_WORKERS) as executor:
            futures = [executor.submit(_send_single, payload) for payload in self.payloads]

            for future in as_completed(futures):
                completed += 1
                success, message = future.result()
                self._log(message)
                if success:
                    ok_count += 1
                else:
                    fail_count += 1

                self._set_progress(completed / total if total else 1.0)

        self._log(f"Koniec wysyłki. Sukces: {ok_count} | Błędy: {fail_count}")
        if fail_count:
            self._set_status(f"Wysyłka zakończona z błędami. Sukces: {ok_count}, błędy: {fail_count}", is_error=True)
        else:
            self._set_status(f"Wysyłka zakończona. Sukces: {ok_count}")


if __name__ == "__main__":
    app = OffersCreatorApp()
    app.mainloop()
