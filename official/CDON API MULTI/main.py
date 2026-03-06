import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
import json
import base64
import csv
import os
import threading
import time
import io
import traceback
import concurrent.futures
import re
from datetime import datetime

# --- LOGIKA BIZNESOWA (API - IMPORTER POST) ---

class CDONClient:
    def __init__(self, merchant_id, api_token, use_sandbox=True, log_callback=None):
        self.merchant_id = merchant_id
        self.api_token = api_token
        self.base_url = "https://merchants-api.sandbox.cdon.com/api" if use_sandbox else "https://merchants-api.cdon.com/api"
        self.log = log_callback if log_callback else print
        self.request_count = 0

    def _get_headers(self):
        auth_string = f"{self.merchant_id}:{self.api_token}"
        auth_bytes = auth_string.encode('ascii')
        base64_bytes = base64.b64encode(auth_bytes)
        base64_auth = base64_bytes.decode('ascii')
        return {
            "Authorization": f"Basic {base64_auth}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "CDON-Python-Importer/3.0-Threaded"
        }

    def _parse_specifications(self, data, sku):
        market_languages = {
            "SE": "sv-SE",
            "DK": "da-DK",
            "FI": "fi-FI"
        }

        csv_specs = []
        for market_code, language in market_languages.items():
            group_key = f"specification_{market_code}_group"
            group_name = str(data.get(group_key, "")).strip()

            name_prefix = f"specification_{market_code}_name_"
            value_prefix = f"specification_{market_code}_value_"
            typo_value_prefix = f"specification_{market_code}_vaule_"

            indexes = set()
            for key in data.keys():
                if key.startswith(name_prefix):
                    suffix = key[len(name_prefix):]
                    if re.fullmatch(r"\d+", suffix):
                        indexes.add(int(suffix))

            attributes = []
            for idx in sorted(indexes):
                name_key = f"{name_prefix}{idx}"
                value_key = f"{value_prefix}{idx}"
                typo_value_key = f"{typo_value_prefix}{idx}"

                spec_name = str(data.get(name_key, "")).strip()
                spec_value = str(data.get(value_key, "")).strip()
                if not spec_value:
                    spec_value = str(data.get(typo_value_key, "")).strip()

                if not spec_name and not spec_value:
                    continue

                if not spec_name or not spec_value:
                    self.log(f"[WARN] {sku}: Niepełna specyfikacja {market_code} dla indeksu {idx} (name/value).")
                    continue

                attributes.append({
                    "name": spec_name,
                    "value": spec_value
                })

            if attributes:
                if not group_name:
                    self.log(f"[WARN] {sku}: Brak {group_key}. Pominięto specifications dla rynku {market_code}.")
                    continue

                csv_specs.append({
                    "language": language,
                    "value": [
                        {
                            "name": group_name,
                            "value": attributes
                        }
                    ]
                })

        if csv_specs:
            return csv_specs

        raw_specs = data.get('specifications_json') or data.get('specifications')
        if not raw_specs or not str(raw_specs).strip():
            return []

        try:
            parsed = json.loads(str(raw_specs).strip())
        except json.JSONDecodeError:
            preview = str(raw_specs).strip().replace("\n", " ")[:160]
            self.log(f"[WARN] {sku}: Niepoprawny JSON w specifications: {preview}")
            return []

        if isinstance(parsed, dict):
            parsed = [parsed]

        if not isinstance(parsed, list):
            self.log(f"[WARN] {sku}: specifications musi być listą lub obiektem JSON.")
            return []

        valid_specs = []
        for spec in parsed:
            if not isinstance(spec, dict):
                continue

            language = spec.get("language")
            sections = spec.get("value")
            if not language or not isinstance(sections, list):
                continue

            clean_sections = []
            for section in sections:
                if not isinstance(section, dict):
                    continue

                section_name = section.get("name")
                attributes = section.get("value")
                if not section_name or not isinstance(attributes, list):
                    continue

                clean_attributes = []
                for attr in attributes:
                    if not isinstance(attr, dict):
                        continue

                    attr_name = attr.get("name")
                    attr_value = attr.get("value")
                    if attr_name is None or attr_value is None:
                        continue

                    attr_entry = {
                        "name": str(attr_name),
                        "value": str(attr_value)
                    }

                    description = attr.get("description")
                    if description is not None and str(description).strip():
                        attr_entry["description"] = str(description)

                    clean_attributes.append(attr_entry)

                if clean_attributes:
                    clean_sections.append({
                        "name": str(section_name),
                        "value": clean_attributes
                    })

            if clean_sections:
                valid_specs.append({
                    "language": str(language),
                    "value": clean_sections
                })

        if not valid_specs:
            self.log(f"[WARN] {sku}: specifications ma niepoprawną strukturę i zostało pominięte.")

        return valid_specs

    def create_product_from_flat_data(self, data):
        """Wysyła jeden produkt metodą POST (tworzenie/nadpisywanie)"""
        sku = data.get('sku')
        if not sku:
            if any(data.values()):
                self.log("[SKIP] Pominięto wiersz: Brak SKU")
            return False

        market_config = {
            'Se': {'code': 'SE', 'lang': 'sv-SE', 'currency': 'SEK'},
            'Dk': {'code': 'DK', 'lang': 'da-DK', 'currency': 'DKK'},
            'Fi': {'code': 'FI', 'lang': 'fi-FI', 'currency': 'EUR'}
        }

        api_markets = []
        api_titles = []
        api_descriptions = []
        api_prices = []
        api_shipping = []
        api_delivery = []

        # --- Przetwarzanie rynków ---
        for suffix, config in market_config.items():
            price_key = f'originalPrice{suffix}'
            
            # Jeśli jest cena, dodajemy rynek
            if data.get(price_key) and str(data[price_key]).strip():
                code = config['code']
                lang = config['lang']
                api_markets.append(code)

                # Tytuł
                if data.get(f'title{suffix}'):
                    api_titles.append({"language": lang, "value": str(data[f'title{suffix}'])})
                
                # Opis
                if data.get(f'description{suffix}'):
                    api_descriptions.append({"language": lang, "value": str(data[f'description{suffix}'])})

                # Ceny i VAT
                try:
                    price_str = str(data[f'originalPrice{suffix}']).replace(',', '.').strip()
                    price_val = float(price_str)
                    
                    vat_str = str(data.get(f'vat{suffix}', '25')).replace(',', '.').strip()
                    if not vat_str: vat_str = '25'
                    
                    vat_rate = float(vat_str)
                    # Poprawka VAT: 25 -> 0.25
                    if vat_rate > 1: vat_rate = vat_rate / 100.0
                        
                    api_prices.append({
                        "market": code,
                        "value": {
                            "amount_including_vat": price_val,
                            "currency": config['currency'],
                            "vat_rate": vat_rate
                        }
                    })
                except ValueError:
                    self.log(f"[WARN] {sku}: Błąd ceny dla {suffix}.")

                # Czas dostawy
                min_time = data.get(f'deliveryTimeMin{suffix}')
                max_time = data.get(f'deliveryTimeMax{suffix}')
                if min_time and max_time:
                    try:
                        api_shipping.append({"market": code, "min": int(float(min_time)), "max": int(float(max_time))})
                    except: pass

                # Typ dostawy (mapowanie)
                del_type = data.get(f'delivery{suffix}')
                if del_type:
                    raw_val = del_type.strip()
                    lookup_key = raw_val.lower().replace(" ", "").replace("_", "")
                    d_map = {'homedelivery': 'home_delivery', 'servicepoint': 'service_point', 'mailbox': 'mailbox', 'digital': 'digital'}
                    val = d_map.get(lookup_key, raw_val.lower().replace(" ", "_"))
                    api_delivery.append({"market": code, "value": val})

        # Właściwości
        properties = []
        if data.get('weight') and str(data['weight']).strip():
            properties.append({"name": "weight_kg", "value": str(data['weight']).replace(',', '.')})

        specifications = self._parse_specifications(data, sku)

        # Zdjęcia (rozdzielane średnikiem)
        extra_images = []
        if data.get('extraImages') and str(data['extraImages']).strip():
            extra_images = [img.strip() for img in data['extraImages'].split(';') if img.strip()]

        # Stan magazynowy
        stock_int = 0
        if data.get('stock') and str(data['stock']).strip():
            try:
                stock_int = int(float(data['stock']))
            except: pass

        # --- BUDOWA BODY (POST) ---
        article_payload = {
            "sku": sku,
            "status": "for sale",
            "quantity": stock_int,
            "main_image": data.get('mainImage'),
            "markets": api_markets,
            "price": api_prices,
            "shipping_time": api_shipping,
            "delivery_type": api_delivery,
            "title": api_titles,
            "description": api_descriptions,
            "category": data.get('category'),
            "brand": data.get('brand'),
            "gtin": data.get('gtin')
        }
        
        # Dodaj zdjęcia tylko jeśli istnieją
        if extra_images:
            article_payload["images"] = extra_images
        if properties:
            article_payload["properties"] = properties
        if specifications:
            article_payload["specifications"] = specifications

        try:
            request_time = time.time()
            
            # POST /v2/articles/bulk (Tworzenie/Nadpisywanie)
            response = requests.post(
                f"{self.base_url}/v2/articles/bulk", 
                headers=self._get_headers(), 
                data=json.dumps({"articles": [article_payload]})
            )
            
            response_time = time.time() - request_time
            self.request_count += 1
            
            # Parsuj response body
            response_body = None
            try:
                response_body = response.json()
            except:
                response_body = response.text
            
            # ZAWSZE loguj szczegóły do TXT logu
            self.log(f"\n[REQUEST #{self.request_count}] SKU: {sku}")
            self.log(f"  Response Time: {response_time:.3f}s")
            self.log(f"  Status Code: {response.status_code}")
            self.log(f"  Response Body: {json.dumps(response_body) if isinstance(response_body, dict) else response_body[:500]}")
            
            if response.status_code in [200, 201, 202]:
                try:
                    response_data = response.json()
                    
                    if isinstance(response_data, dict):
                        # Sprawdzenie błędów
                        if response_data.get('errors') or response_data.get('message') or response_data.get('description'):
                            msg = response_data.get('message') or response_data.get('description')
                            if response_data.get('errors'):
                                msg += " " + json.dumps(response_data.get('errors'))
                            self.log(f"[BŁĄD API] {sku}: {msg}")
                            return False
                        
                        # Sukces (Batch ID lub Receipt)
                        if response_data.get('receipt') or response_data.get('batch_id'):
                            self.log(f"[OK] {sku}: Wysłano")
                            return True
                        
                        # OSTRZEŻENIE: OK status ale brak batch_id/receipt
                        self.log(f"[WARN] {sku}: Status OK ({response.status_code}) ale brak potwierdzenia (receipt/batch_id)")
                        return True

                    self.log(f"[OK] {sku}: Wysłano pomyślnie.")
                    return True

                except json.JSONDecodeError:
                    # OSTRZEŻENIE: Status OK ale non-JSON response
                    self.log(f"[WARN] {sku}: Status {response.status_code}, brak JSON")
                    return True
            else:
                self.log(f"[API ERROR] {sku}: {response.status_code}")
                return False
        except Exception as e:
            self.log(f"[EXCEPTION] {sku}: {str(e)}")
            return False

# --- GUI ---

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

ACCENT_COLOR = "#ff69b4"
ACCENT_HOVER = "#e754a8"
STOP_COLOR = "#c2188b"
STOP_HOVER = "#a31574"
APP_BG = "#fff8fc"
PANEL_BG = "#ffeaf5"
INPUT_BG = "#fff1f8"

class CDONApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CDON Product Importer v3.0 - Multithreaded POST")
        self.geometry("800x750")
        self.configure(fg_color=APP_BG)
        
        # Zmienne
        self.file_path = tk.StringVar()
        self.account_var = tk.StringVar()
        self.use_sandbox = tk.BooleanVar(value=False)
        self.is_running = False
        
        # Zmienne licznika
        self.processed_count = 0
        self.success_count = 0
        
        self.accounts_data = {}
        self.config_file = "accounts.csv"
        self.log_file = None

        self._load_accounts_config()
        self._create_widgets()

    def _load_accounts_config(self):
        if not os.path.exists(self.config_file):
            try:
                with open(self.config_file, "w", encoding="utf-8") as f:
                    f.write("Nazwa Konta;MerchantID;APIToken\n")
                    f.write("SandboxTest;12345;abc-token-przyklad\n")
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie można utworzyć pliku konfiguracyjnego: {e}")
        
        self.accounts_data = {}
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line: continue
                    parts = line.split(';')
                    if len(parts) >= 3:
                        name = parts[0].strip()
                        m_id = parts[1].strip()
                        token = parts[2].strip()
                        if name.lower() == "nazwa konta": continue
                        if name: self.accounts_data[name] = {"id": m_id, "token": token}
        except Exception as e:
            messagebox.showerror("Błąd konfiguracji", f"Błąd odczytu {self.config_file}:\n{e}")

    def _create_widgets(self):
        config_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        config_frame.pack(padx=20, pady=20, fill="x")

        ctk.CTkLabel(config_frame, text="Wybór Konta API", font=("Arial", 16, "bold")).pack(pady=10)

        grid_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        grid_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(grid_frame, text="Wybierz konto:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        account_names = list(self.accounts_data.keys())
        self.account_combo = ctk.CTkComboBox(
            grid_frame,
            values=account_names,
            variable=self.account_var,
            width=300,
            state="readonly",
            fg_color=INPUT_BG,
            border_color=ACCENT_COLOR,
            button_color=ACCENT_COLOR,
            button_hover_color=ACCENT_HOVER,
            text_color="#4a2a3a",
            dropdown_fg_color=INPUT_BG,
            dropdown_hover_color="#ffd6ea",
            dropdown_text_color="#4a2a3a"
        )
        self.account_combo.grid(row=0, column=1, padx=5, pady=5)
        if account_names: self.account_combo.set(account_names[0])
        else: self.account_combo.set("Brak kont w accounts.csv")

        ctk.CTkButton(
            grid_frame,
            text="Odśwież",
            command=self._refresh_accounts,
            width=80,
            fg_color=ACCENT_COLOR,
            hover_color=ACCENT_HOVER,
            text_color="white"
        ).grid(row=0, column=2, padx=5)
        ctk.CTkCheckBox(config_frame, text="Tryb Sandbox (Testowy)", variable=self.use_sandbox).pack(pady=10)

        file_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        file_frame.pack(padx=20, pady=10, fill="x")

        ctk.CTkLabel(file_frame, text="Dane Produktowe (CSV)", font=("Arial", 16, "bold")).pack(pady=10)
        file_sub_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_sub_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            file_sub_frame,
            text="Wybierz plik CSV",
            command=self.select_file,
            fg_color=ACCENT_COLOR,
            hover_color=ACCENT_HOVER,
            text_color="white"
        ).pack(side="left", padx=10)
        ctk.CTkLabel(file_sub_frame, textvariable=self.file_path, text_color="silver").pack(side="left", padx=10)

        action_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        action_frame.pack(padx=20, pady=10, fill="x")

        # Info o wątkach na przycisku
        self.start_btn = ctk.CTkButton(
            action_frame, 
            text="ROZPOCZNIJ IMPORT (5 wątków)", 
            command=self.start_process, 
            fg_color=ACCENT_COLOR,
            hover_color=ACCENT_HOVER,
            height=40, 
            font=("Arial", 14, "bold")
        )
        self.start_btn.pack(pady=(10, 5), fill="x", padx=50)

        self.stop_btn = ctk.CTkButton(
            action_frame, 
            text="ZATRZYMAJ", 
            command=self.stop_process, 
            fg_color=STOP_COLOR,
            hover_color=STOP_HOVER,
            text_color="white",
            height=40, 
            font=("Arial", 14, "bold"), 
            state="disabled"
        )
        self.stop_btn.pack(pady=(5, 10), fill="x", padx=50)

        self.progress_bar = ctk.CTkProgressBar(action_frame)
        self.progress_bar.pack(pady=10, fill="x", padx=20)
        self.progress_bar.configure(progress_color=ACCENT_COLOR)
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(action_frame, text="Gotowy")
        self.status_label.pack(pady=5)

        log_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        log_frame.pack(padx=20, pady=10, fill="both", expand=True)

        ctk.CTkLabel(log_frame, text="Logi operacji:", anchor="w").pack(fill="x", padx=5, pady=5)
        self.log_box = ctk.CTkTextbox(log_frame)
        self.log_box.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_box.configure(state="disabled")

    def _refresh_accounts(self):
        self._load_accounts_config()
        names = list(self.accounts_data.keys())
        self.account_combo.configure(values=names)
        if names: self.account_combo.set(names[0])

    def select_file(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if filename: self.file_path.set(filename)

    def log(self, message):
        def _update():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        
        # Zapisz do pliku jeśli jest otwarty
        if self.log_file:
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(message + "\n")
            except:
                pass
        
        self.after(0, _update)

    def update_progress(self, current, total):
        progress = current / total
        self.progress_bar.set(progress)
        self.status_label.configure(text=f"Przetwarzanie: {current}/{total} (Wątki aktywne)")

    def start_process(self):
        name = self.account_var.get()
        if not name or name not in self.accounts_data:
            messagebox.showwarning("Błąd", "Wybierz poprawne konto!")
            return
        
        creds = self.accounts_data[name]
        if not self.file_path.get():
            messagebox.showwarning("Brak pliku", "Wybierz plik CSV!")
            return
        
        self.is_running = True
        self.start_btn.configure(state="disabled", text="PRZETWARZANIE W TLE...")
        self.stop_btn.configure(state="normal")
        self.progress_bar.set(0)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        
        # Przygotuj plik logu
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = f"cdon_import_{timestamp}.log"
        
        # Uruchamiamy główny wątek zarządzający
        threading.Thread(target=self.run_import, args=(creds["id"], creds["token"]), daemon=True).start()

    def stop_process(self):
        if self.is_running:
            self.is_running = False
            self.stop_btn.configure(state="disabled", text="ZATRZYMYWANIE...")
            self.log("[INFO] Wysłano sygnał zatrzymania. Czekanie na zakończenie aktywnych wątków...")

    def run_import(self, merchant_id, api_token):
        file_path = self.file_path.get()
        client = CDONClient(merchant_id, api_token, self.use_sandbox.get(), self.log)
        
        self.log("=" * 80)
        self.log("--- START IMPORTU (WIELOWĄTKOWY) ---")
        self.log(f"Log file: {self.log_file}")
        self.log("Całe response'y z API są zapisywane do TXT logu")
        self.log("=" * 80)
        content = None
        encoding_used = None
        for enc in ['utf-8-sig', 'utf-8', 'cp1250', 'latin-1']:
            try:
                with open(file_path, mode='r', encoding=enc) as f:
                    content = f.read()
                encoding_used = enc
                break
            except UnicodeDecodeError: continue
            except Exception as e:
                self.log(f"[BŁĄD PLIKU] {str(e)}")
                self._reset_ui()
                return

        if not content:
            self.log("[BŁĄD] Nie rozpoznano kodowania.")
            self._reset_ui()
            return

        try:
            f_io = io.StringIO(content)
            sample = f_io.read(4096)
            f_io.seek(0)
            delimiter = ';'
            try:
                if len(sample) > 5:
                    delimiter = csv.Sniffer().sniff(sample, delimiters=[';', '\t', '|']).delimiter
            except: pass
            
            self.log(f"Separator: '{delimiter}' | Kodowanie: {encoding_used}")

            reader = csv.DictReader(f_io, delimiter=delimiter)
            if reader.fieldnames:
                reader.fieldnames = [name.strip().replace('"', '').replace("'", "") for name in reader.fieldnames]
            
            headers_lower = [h.lower() for h in reader.fieldnames]
            if 'sku' not in headers_lower:
                self.log(f"[BŁĄD] Brak kolumny 'sku'.")
                self._reset_ui()
                return

            rows = [row for row in reader if any(row.values())] 
            
            # Przygotowanie danych do wątków
            prepared_data = []
            for row in rows:
                clean_row = {}
                for k, v in row.items():
                    if k: clean_row[k.strip()] = v.strip() if v else ""
                prepared_data.append(clean_row)
            
            total_rows = len(prepared_data)
            self.log(f"Produktów do przetworzenia: {total_rows}")
            self.log("Uruchamianie puli wątków (Max: 5)...")

            self.processed_count = 0
            self.success_count = 0

            # Funkcja dla workera
            def process_item(item_data):
                if not self.is_running:
                    return False
                time.sleep(0.01) # Mały delay dla stabilności
                return client.create_product_from_flat_data(item_data)

            # --- WIELOWĄTKOWOŚĆ TUTAJ ---
            # Możesz zmienić 'max_workers=5' na inną liczbę, jeśli potrzebujesz
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(process_item, item) for item in prepared_data]
                
                for future in concurrent.futures.as_completed(futures):
                    if not self.is_running:
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    
                    try:
                        result = future.result()
                        if result:
                            self.success_count += 1
                    except Exception as exc:
                        self.log(f"[WĄTEK ERROR] {str(exc)}")
                    
                    self.processed_count += 1
                    self.after(0, lambda c=self.processed_count: self.update_progress(c, total_rows))

            self.log("\n" + "=" * 80)
            self.log("--- KONIEC IMPORTU ---")
            self.log(f"Sukcesy: {self.success_count} | Błędy: {total_rows - self.success_count}")
            self.log(f"Log zapisany do: {self.log_file}")
            self.log("=" * 80)
            
            msg = f"Sukcesy: {self.success_count}\nBłędy: {total_rows - self.success_count}"
            msg += f"\n\n📋 Pełne response'y z API w logu:"
            msg += f"\n{self.log_file}"
            if not self.is_running: msg += "\n\n(Zatrzymano)"
            self.after(0, lambda m=msg: messagebox.showinfo("Info", m))

        except Exception as e:
            self.log(f"[CRITICAL] {str(e)}")
            traceback.print_exc()
        finally:
            self.is_running = False
            self.after(0, self._reset_ui)

    def _reset_ui(self):
        self.start_btn.configure(state="normal", text="ROZPOCZNIJ IMPORT (5 wątków)")
        self.stop_btn.configure(state="disabled", text="ZATRZYMAJ")
        self.status_label.configure(text="Gotowy")

if __name__ == "__main__":
    app = CDONApp()
    app.mainloop()