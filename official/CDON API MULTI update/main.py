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
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- LOGIKA BIZNESOWA (API - UPDATER) ---

class CDONClient:
    def __init__(self, merchant_id, api_token, use_sandbox=True, log_callback=None):
        self.merchant_id = merchant_id
        self.api_token = api_token
        self.base_url = "https://merchants-api.sandbox.cdon.com/api" if use_sandbox else "https://merchants-api.cdon.com/api"
        self.log = log_callback if log_callback else print

    def _get_headers(self):
        auth_string = f"{self.merchant_id}:{self.api_token}"
        auth_bytes = auth_string.encode('ascii')
        base64_bytes = base64.b64encode(auth_bytes)
        base64_auth = base64_bytes.decode('ascii')
        return {
            "Authorization": f"Basic {base64_auth}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "CDON-Python-Updater/1.2-MultiThread"
        }

    def _parse_specifications(self, data):
        market_languages = {
            "SE": "sv-SE",
            "DK": "da-DK",
            "FI": "fi-FI"
        }

        parsed_specs = []

        for market_code, language in market_languages.items():
            group_key = f"specification_{market_code}_group"
            group_name = str(data.get(group_key, "")).strip()
            if not group_name:
                group_name = "Specifications"

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
                spec_name = str(data.get(f"{name_prefix}{idx}", "")).strip()
                spec_value = str(data.get(f"{value_prefix}{idx}", "")).strip()
                if not spec_value:
                    spec_value = str(data.get(f"{typo_value_prefix}{idx}", "")).strip()

                if not spec_name or not spec_value:
                    continue

                attributes.append({
                    "name": spec_name,
                    "value": spec_value
                })

            if attributes:
                parsed_specs.append({
                    "language": language,
                    "value": [
                        {
                            "name": group_name,
                            "value": attributes
                        }
                    ]
                })

        return parsed_specs

    def process_product_data(self, data):
        """
        Przetwarza jeden wiersz danych i wysyła żądanie do CDON.
        Zwraca True jeśli sukces, False jeśli błąd.
        """
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

        # --- PRZETWARZANIE DANYCH (CENY, OPISY, DOSTAWA) ---
        for suffix, config in market_config.items():
            code = config['code']
            lang = config['lang']
            
            market_active = False

            # 1. Ceny (tylko jeśli podane)
            price_key = f'originalPrice{suffix}'
            if data.get(price_key) and str(data[price_key]).strip():
                market_active = True
                try:
                    price_str = str(data[price_key]).replace(',', '.').strip()
                    price_val = float(price_str)
                    
                    vat_str = str(data.get(f'vat{suffix}', '25')).replace(',', '.').strip()
                    if not vat_str: vat_str = '25'
                    vat_rate = float(vat_str)
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
                    self.log(f"[WARN] {sku}: Błąd ceny dla {suffix}, pomijam cenę.")

            # 2. Tytuły
            title_key = f'title{suffix}'
            if data.get(title_key) and str(data[title_key]).strip():
                market_active = True
                api_titles.append({"language": lang, "value": str(data[title_key])})

            # 3. Opisy
            desc_key = f'description{suffix}'
            if data.get(desc_key) and str(data[desc_key]).strip():
                market_active = True
                api_descriptions.append({"language": lang, "value": str(data[desc_key])})
            
            # 4. Dostawa
            del_min = data.get(f'deliveryTimeMin{suffix}')
            del_max = data.get(f'deliveryTimeMax{suffix}')
            if del_min and del_max:
                market_active = True
                try:
                    api_shipping.append({"market": code, "min": int(float(del_min)), "max": int(float(del_max))})
                except: pass

            del_type = data.get(f'delivery{suffix}')
            if del_type and str(del_type).strip():
                market_active = True
                raw_val = del_type.strip()
                lookup = raw_val.lower().replace(" ", "").replace("_", "")
                d_map = {'homedelivery': 'home_delivery', 'servicepoint': 'service_point', 'mailbox': 'mailbox', 'digital': 'digital'}
                val = d_map.get(lookup, raw_val.lower().replace(" ", "_"))
                api_delivery.append({"market": code, "value": val})

            if market_active:
                api_markets.append(code)

        # Właściwości
        properties = []
        if data.get('weight') and str(data['weight']).strip():
            properties.append({"name": "weight_kg", "value": str(data['weight']).replace(',', '.')})

        specifications = self._parse_specifications(data)

        # Zdjęcia
        extra_images = []
        if data.get('extraImages') and str(data['extraImages']).strip():
            extra_images = [img.strip() for img in data['extraImages'].split(';') if img.strip()]

        # Stan magazynowy
        stock_int = None
        if data.get('stock') and str(data['stock']).strip():
            try:
                stock_int = int(float(data['stock']))
            except ValueError:
                pass 

        # --- BUDOWA BODY (ACTIONS STRUCTURE) ---
        article_body = {
            "status": "for sale"
        }

        if api_markets: article_body["markets"] = api_markets
        if api_titles: article_body["title"] = api_titles
        if api_descriptions: article_body["description"] = api_descriptions
        if api_shipping: article_body["shipping_time"] = api_shipping
        if api_delivery: article_body["delivery_type"] = api_delivery
        if extra_images: article_body["images"] = extra_images
        if properties: article_body["properties"] = properties
        if specifications: article_body["specifications"] = specifications

        if data.get('mainImage') and str(data['mainImage']).strip():
            article_body["main_image"] = str(data['mainImage']).strip()
        
        if data.get('category') and str(data['category']).strip():
            article_body["category"] = str(data['category']).strip()
            
        if data.get('brand') and str(data['brand']).strip():
            article_body["brand"] = str(data['brand']).strip()
            
        if data.get('gtin') and str(data['gtin']).strip():
            article_body["gtin"] = str(data['gtin']).strip()

        # --- BUDOWA LISTY AKCJI ---
        actions = []

        if len(article_body) > 1 and api_markets:
            actions.append({"sku": sku, "action": "update_article", "body": article_body})
        
        if api_prices:
            actions.append({"sku": sku, "action": "update_article_price", "body": {"price": api_prices}})
            
        if stock_int is not None:
            actions.append({"sku": sku, "action": "update_article_quantity", "body": {"quantity": stock_int}})

        if not actions:
            self.log(f"[SKIP] {sku}: Brak danych do aktualizacji w tym wierszu.")
            return True

        try:
            # Używamy timeout, aby wątki się nie zawieszały
            response = requests.put(
                f"{self.base_url}/v2/articles/bulk", 
                headers=self._get_headers(), 
                data=json.dumps({"actions": actions}),
                timeout=30 
            )
            
            if response.status_code in [200, 201, 202]:
                try:
                    response_data = response.json()
                    has_errors = False
                    items = []

                    if isinstance(response_data, dict):
                        if 'success' in response_data and isinstance(response_data['success'], list):
                             for s in response_data['success']:
                                 s['_is_success_flag'] = True
                                 items.append(s)
                        if 'failed' in response_data and isinstance(response_data['failed'], list):
                             for f in response_data['failed']:
                                 f['_is_success_flag'] = False
                                 items.append(f)
                        if not items and 'actions' in response_data:
                            items = response_data['actions']
                    elif isinstance(response_data, list):
                        items = response_data

                    if not items:
                        if isinstance(response_data, dict):
                            if response_data.get('receipt') or response_data.get('batch_id'):
                                 batch_id = response_data.get('receipt') or response_data.get('batch_id')
                                 self.log(f"[OK] {sku}: Wysłano (Batch: {batch_id}).")
                                 return True
                            if response_data.get('errors') or response_data.get('message'):
                                msg = response_data.get('message') or json.dumps(response_data.get('errors'))
                                self.log(f"[BŁĄD API] {sku}: {msg}")
                                return False
                        self.log(f"[WARN] {sku}: Pusta odpowiedź.")
                        return True

                    for item in items:
                        action_name = item.get('action', 'akcja')
                        is_success = item.get('_is_success_flag')
                        if is_success is None:
                             is_success = item.get('success') or item.get('status') in ['ok', 'success']
                        
                        if is_success is False: 
                            has_errors = True
                            msg = item.get('message') or item.get('errors')
                            if isinstance(msg, list):
                                msg = "; ".join([f"{e.get('location', '?')}: {e.get('message', '?')}" for e in msg if isinstance(e, dict)])
                            elif not msg:
                                msg = json.dumps(item)
                            self.log(f"[BŁĄD {action_name}]: {msg}")

                    if has_errors:
                        self.log(f"[FAIL] {sku}: Część zmian odrzucona.")
                        return False
                    else:
                        batch_info = ""
                        if isinstance(response_data, dict) and response_data.get('batch_id'):
                            batch_info = f" (Batch: {response_data.get('batch_id')})"
                        self.log(f"[OK] {sku}: Zaktualizowano{batch_info}.")
                        return True

                except json.JSONDecodeError:
                    self.log(f"[WARN] {sku}: Błąd JSON odpowiedzi.")
                    return True
            else:
                self.log(f"[API ERROR] {sku}: {response.status_code} - {response.text}")
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

class CDONUpdaterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CDON Product Updater (Multi-Thread) - v2.0")
        self.geometry("900x800")
        self.configure(fg_color=APP_BG)
        
        self.file_path = tk.StringVar()
        self.account_var = tk.StringVar()
        self.use_sandbox = tk.BooleanVar(value=False)
        self.thread_count = tk.IntVar(value=5) # Domyślnie 5 wątków
        self.is_running = False
        
        self.accounts_data = {}
        self.config_file = "accounts.csv"

        self._load_accounts_config()
        self._create_widgets()

    def _load_accounts_config(self):
        """Wczytuje konfigurację kont z pliku CSV."""
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
        """Tworzy elementy interfejsu graficznego."""
        header_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        header_frame.pack(padx=20, pady=10, fill="x")
        ctk.CTkLabel(header_frame, text="NARZĘDZIE AKTUALIZACJI (WIELOWĄTKOWE)", font=("Arial", 20, "bold"), text_color=ACCENT_COLOR).pack(pady=10)

        config_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        config_frame.pack(padx=20, pady=10, fill="x")

        # Konto
        ctk.CTkLabel(config_frame, text="Konfiguracja Połączenia", font=("Arial", 16, "bold")).pack(pady=5)
        grid_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        grid_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(grid_frame, text="Konto:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        account_names = list(self.accounts_data.keys())
        self.account_combo = ctk.CTkComboBox(
            grid_frame,
            values=account_names,
            variable=self.account_var,
            width=250,
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
        else: self.account_combo.set("Brak kont")
        ctk.CTkButton(
            grid_frame,
            text="Odśwież",
            command=self._refresh_accounts,
            width=80,
            fg_color=ACCENT_COLOR,
            hover_color=ACCENT_HOVER,
            text_color="white"
        ).grid(row=0, column=2, padx=5)
        
        # Opcje wątków i sandbox
        opts_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        opts_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkCheckBox(opts_frame, text="Tryb Sandbox", variable=self.use_sandbox).pack(side="left", padx=20)
        
        # Suwak wątków
        thread_frame = ctk.CTkFrame(opts_frame, fg_color="transparent")
        thread_frame.pack(side="right", padx=20)
        ctk.CTkLabel(thread_frame, text="Liczba wątków:").pack(side="left", padx=5)
        self.thread_slider = ctk.CTkSlider(thread_frame, from_=1, to=20, number_of_steps=19, variable=self.thread_count, width=150)
        self.thread_slider.configure(progress_color=ACCENT_COLOR, button_color=ACCENT_COLOR, button_hover_color=ACCENT_HOVER)
        self.thread_slider.pack(side="left", padx=5)
        self.thread_label = ctk.CTkLabel(thread_frame, textvariable=self.thread_count, width=30)
        self.thread_label.pack(side="left")

        # Plik
        file_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        file_frame.pack(padx=20, pady=10, fill="x")
        ctk.CTkLabel(file_frame, text="Dane (CSV)", font=("Arial", 16, "bold")).pack(pady=5)
        file_sub_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_sub_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkButton(
            file_sub_frame,
            text="Wybierz plik",
            command=self.select_file,
            fg_color=ACCENT_COLOR,
            hover_color=ACCENT_HOVER,
            text_color="white"
        ).pack(side="left", padx=10)
        ctk.CTkLabel(file_sub_frame, textvariable=self.file_path, text_color="silver").pack(side="left", padx=10)

        # Akcje
        action_frame = ctk.CTkFrame(self, fg_color=PANEL_BG)
        action_frame.pack(padx=20, pady=10, fill="x")
        self.start_btn = ctk.CTkButton(action_frame, text="START (ASYNC)", command=self.start_process, fg_color=ACCENT_COLOR, hover_color=ACCENT_HOVER, text_color="white", height=40, font=("Arial", 14, "bold"))
        self.start_btn.pack(pady=(10, 5), fill="x", padx=50)
        self.stop_btn = ctk.CTkButton(action_frame, text="ZATRZYMAJ", command=self.stop_process, fg_color=STOP_COLOR, hover_color=STOP_HOVER, text_color="white", height=40, font=("Arial", 14, "bold"), state="disabled")
        self.stop_btn.pack(pady=(5, 10), fill="x", padx=50)

        self.progress_bar = ctk.CTkProgressBar(action_frame)
        self.progress_bar.pack(pady=10, fill="x", padx=20)
        self.progress_bar.configure(progress_color=ACCENT_COLOR)
        self.progress_bar.set(0)
        self.status_label = ctk.CTkLabel(action_frame, text="Gotowy")
        self.status_label.pack(pady=5)

        # Logi
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
        """Dodaje wiadomość do logów w sposób bezpieczny dla wątków."""
        def _update():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        # Bezpieczne wywołanie aktualizacji GUI z innego wątku
        self.after(0, _update)

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
        self.thread_slider.configure(state="disabled")
        self.progress_bar.set(0)
        
        # Czyszczenie logów
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        
        # Uruchomienie wątku zarządczego (który uruchomi pulę workerów)
        threading.Thread(target=self.run_import, args=(creds["id"], creds["token"]), daemon=True).start()

    def stop_process(self):
        if self.is_running:
            self.is_running = False
            self.stop_btn.configure(state="disabled", text="ZATRZYMYWANIE...")
            self.log("[INFO] Wysłano sygnał zatrzymania. Oczekiwanie na zakończenie aktywnych wątków...")

    def run_import(self, merchant_id, api_token):
        """Główna logika przetwarzania działająca w osobnym wątku."""
        file_path = self.file_path.get()
        client = CDONClient(merchant_id, api_token, self.use_sandbox.get(), self.log)
        num_workers = self.thread_count.get()

        self.log(f"--- START (Wątków: {num_workers}) ---")
        
        # Wczytywanie pliku
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
                    delimiter = csv.Sniffer().sniff(sample, delimiters=[',', ';', '\t', '|']).delimiter
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
            total = len(rows)
            self.log(f"Produktów do przetworzenia: {total}")

            # Zmienne do śledzenia postępu (z Lockiem)
            lock = threading.Lock()
            progress_counter = {"processed": 0, "success": 0, "failed": 0}

            # Funkcja workera
            def process_row(row_data):
                if not self.is_running:
                    return None # Przerwij jeśli stop

                # Czyste dane
                clean_row = {}
                for k, v in row_data.items():
                    if k: clean_row[k.strip()] = v.strip() if v else ""
                
                # Wykonanie żądania API
                result = client.process_product_data(clean_row)
                
                # Aktualizacja liczników pod Lockiem
                with lock:
                    progress_counter["processed"] += 1
                    if result:
                        progress_counter["success"] += 1
                    else:
                        progress_counter["failed"] += 1
                    
                    curr = progress_counter["processed"]
                    prog = curr / total
                    
                    # UI Update (Thread Safe via .after)
                    self.after(0, lambda p=prog: self.progress_bar.set(p))
                    self.after(0, lambda t=f"Postęp: {curr}/{total} (OK: {progress_counter['success']})": self.status_label.configure(text=t))
                
                return result

            # ThreadPoolExecutor - Główna pętla równoległa
            with ThreadPoolExecutor(max_workers=num_workers) as executor:
                futures = {executor.submit(process_row, row): row for row in rows}
                
                for future in as_completed(futures):
                    # Sprawdzenie czy użytkownik wcisnął STOP
                    if not self.is_running:
                        self.log("[INFO] Anulowanie pozostałych zadań...")
                        # Wyjście z pętli powoduje, że Context Manager executora zacznie zamykanie (shutdown)
                        # Aktywne wątki dokończą pracę, oczekujące zostaną anulowane
                        break 
                    
                    try:
                        future.result() # Pobierz wynik (dla obsługi wyjątków)
                    except Exception as exc:
                        self.log(f"[CRITICAL WORKER ERROR] {exc}")

            self.log("--- KONIEC PRZETWARZANIA ---")
            msg = f"Przetworzono: {progress_counter['processed']}/{total}\nSukcesy: {progress_counter['success']}\nBłędy: {progress_counter['failed']}"
            if not self.is_running:
                msg += "\n(Zatrzymano przez użytkownika)"
            
            self.after(0, lambda: messagebox.showinfo("Info", msg))

        except Exception as e:
            self.log(f"[CRITICAL APP ERROR] {str(e)}")
            traceback.print_exc()
        finally:
            self.is_running = False
            self.after(0, self._reset_ui)

    def _reset_ui(self):
        self.start_btn.configure(state="normal", text="START (ASYNC)")
        self.stop_btn.configure(state="disabled", text="ZATRZYMAJ")
        self.thread_slider.configure(state="normal")
        self.status_label.configure(text="Gotowy")

if __name__ == "__main__":
    app = CDONUpdaterApp()
    app.mainloop()