import xml.etree.ElementTree as ET
import csv
import os
import urllib.request
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
from urllib.parse import urlparse
import tempfile
import time

# Ustawienia CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# --- Standardowa ścieżka zapisu (można zmienić) ---
DOMYSLNA_SCIEZKA_ZAPISU = "C:/Users/super/Downloads"
# ----------------------------------------------------

def clean_text(text):
    """Zastępuje znaki nowej linii i inne białe znaki pojedynczą spacją."""
    if not text:
        return ""
    return " ".join(text.split())

def pobierz_xml(url, sciezka_docelowa):
    """Pobiera plik XML z podanego URL."""
    try:
        opener = urllib.request.build_opener()
        opener.addheaders = [('User-agent', 'Mozilla/5.0')]
        urllib.request.install_opener(opener)
        urllib.request.urlretrieve(url, sciezka_docelowa)
        return True
    except Exception as e:
        messagebox.showerror(f"Błąd pobierania ({url})", f"Wystąpił błąd: {e}")
        return False

def parsuj_xml(sciezka_pliku):
    """Parsuje plik XML i ekstrahuje 'id', 'id_bl' oraz dodaje pustą kolumnę 'war'."""
    try:
        drzewo = ET.parse(sciezka_pliku)
        root = drzewo.getroot()
        dane = []

        for element in root.findall("o"):
            id_val = element.get("id")
            id_bl_val = ""  # Domyślnie puste

            # Szukanie atrybutu o nazwie 'id_bl' w tagu <attrs>
            atrybuty_elem = element.find("attrs")
            if atrybuty_elem is not None:
                for atrybut in atrybuty_elem.findall("a"):
                    if atrybut.get("name") == "id_bl":
                        id_bl_val = clean_text(atrybut.text) if atrybut.text else ""
                        break

            wiersz = {
                "id": id_val,
                "id_bl": id_bl_val,
                "war": ""
            }
            dane.append(wiersz)
            
        return dane
    except FileNotFoundError:
        messagebox.showerror("Błąd pliku", f"Nie znaleziono pliku: {sciezka_pliku}")
        return []
    except ET.ParseError as e:
        messagebox.showerror("Błąd parsowania XML", f"Błąd w pliku {os.path.basename(sciezka_pliku)}: {e}")
        return []
    except Exception as e:
        messagebox.showerror("Nieoczekiwany błąd parsowania", f"Wystąpił błąd: {e}")
        return []

def zapisz_do_csv(dane, sciezka_pliku):
    """Zapisuje dane do pliku CSV z określonymi kolumnami: id, id_bl, war."""
    pola = ['id', 'id_bl', 'war']

    try:
        with open(sciezka_pliku, 'w', encoding='utf-8-sig', newline='') as plik_csv:
            writer = csv.DictWriter(plik_csv, fieldnames=pola, delimiter=';', extrasaction='ignore')
            writer.writeheader()
            writer.writerows(dane)
        messagebox.showinfo("Sukces", f"Wszystkie dane zostały pomyślnie zapisane do:\n{os.path.abspath(sciezka_pliku)}")
        return True
    except Exception as e:
        messagebox.showerror(f"Błąd zapisu CSV ({os.path.basename(sciezka_pliku)})", f"Wystąpił błąd: {e}")
        return False

def przetworz_wiele_url_jeden_plik(urls, sciezka_zapisu_csv, app_instance):
    """Przetwarza wiele URL-i i zapisuje do jednego pliku CSV."""
    all_dane = []
    all_nazwy_bazowe = []
    katalog_tymczasowy = tempfile.gettempdir()
    liczba_url = len(urls)
    sukcesy_przetwarzania = 0
    bledy = 0

    if liczba_url == 0:
        app_instance.update_status("Nie podano żadnych nazw.", 0)
        return

    if not os.path.exists(sciezka_zapisu_csv):
        try: os.makedirs(sciezka_zapisu_csv)
        except Exception as e:
            messagebox.showerror("Błąd ścieżki zapisu", f"Nie można utworzyć katalogu: {sciezka_zapisu_csv}\nBłąd: {e}\nPliki będą zapisywane w katalogu roboczym.")
            sciezka_zapisu_csv = os.getcwd()
            app_instance.pole_sciezki_zapisu.delete(0, ctk.END)
            app_instance.pole_sciezki_zapisu.insert(0, sciezka_zapisu_csv)

    for i, url in enumerate(urls):
        url = url.strip()
        if not url: continue

        postep = (i + 1) / liczba_url
        nazwa_pliku_url = os.path.basename(urlparse(url).path) or f"feed_{i+1}.xml"
        app_instance.update_status(f"Przetwarzanie {i+1}/{liczba_url}: {nazwa_pliku_url}...", postep)

        nazwa_bazowa_xml = os.path.splitext(nazwa_pliku_url)[0]
        teraz_timestamp_temp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        sciezka_lokalna_xml = os.path.join(katalog_tymczasowy, f"temp_{nazwa_bazowa_xml}_{teraz_timestamp_temp}.xml")

        if pobierz_xml(url, sciezka_lokalna_xml):
            dane = parsuj_xml(sciezka_lokalna_xml)
            if dane:
                all_dane.extend(dane)
                all_nazwy_bazowe.append(nazwa_bazowa_xml)
                sukcesy_przetwarzania += 1
            else:
                bledy += 1
                app_instance.update_status(f"Błąd parsowania lub brak danych w {nazwa_pliku_url}. Pomijanie.", postep)
                time.sleep(0.5)
            
            try: os.remove(sciezka_lokalna_xml)
            except Exception as e: print(f"Ostrzeżenie: Nie udało się usunąć pliku tymczasowego {sciezka_lokalna_xml}: {e}")
        else:
            bledy += 1
            app_instance.update_status(f"Błąd pobierania {nazwa_pliku_url}. Pomijanie.", postep)
            time.sleep(0.5)

    if not all_dane:
        app_instance.update_status("Nie udało się przetworzyć żadnych danych.", 0)
        messagebox.showwarning("Brak danych", f"Nie udało się pobrać ani sparsować danych z żadnego podanego URL. Błędy: {bledy}")
        app_instance.reset_gui_after_delay()
        return

    # Tworzenie nazwy pliku
    nazwa_laczona = "_".join(all_nazwy_bazowe)
    if len(nazwa_laczona) > 100:
        nazwa_laczona = f"{all_nazwy_bazowe[0]}_and_{len(all_nazwy_bazowe)-1}_more"
        
    teraz_format_czasu_csv = datetime.now().strftime("%d%m%y-%H%M%S")
    nazwa_pliku_csv = os.path.join(sciezka_zapisu_csv, f"{nazwa_laczona}_{teraz_format_czasu_csv}.csv")

    app_instance.update_status("Zapisywanie połączonych danych...", 0.95)
    
    if zapisz_do_csv(all_dane, nazwa_pliku_csv):
        app_instance.update_status(f"Zakończono. Przetworzono: {sukcesy_przetwarzania}, Błędy: {bledy}", 1)
    else:
        app_instance.update_status(f"Błąd zapisu pliku! Przetworzono: {sukcesy_przetwarzania}, Błędy: {bledy}", 1)

    app_instance.reset_gui_after_delay()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Konwerter XML do CSV")
        
        window_width = 600
        window_height = 550
        self.minsize(500, 450)

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  
        self.grid_rowconfigure(1, weight=0)  
        self.grid_rowconfigure(2, weight=0)  
        self.grid_rowconfigure(3, weight=0)  
        self.grid_rowconfigure(4, weight=0)  

        input_frame_urls = ctk.CTkFrame(self)
        input_frame_urls.grid(row=0, column=0, padx=20, pady=(20,10), sticky="nsew")
        input_frame_urls.grid_columnconfigure(0, weight=1)
        input_frame_urls.grid_rowconfigure(1, weight=1) 

        # --- ZMIANA --- Etykieta prosi o podanie nazw, a nie URL-i
        etykieta_url = ctk.CTkLabel(input_frame_urls, text="Wpisz nazwy plików (każda w nowej linii):")
        etykieta_url.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.pole_url = ctk.CTkTextbox(input_frame_urls, height=150) 
        self.pole_url.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

        input_frame_path = ctk.CTkFrame(self)
        input_frame_path.grid(row=1, column=0, padx=20, pady=(5,10), sticky="ew")
        input_frame_path.grid_columnconfigure(1, weight=1)

        etykieta_sciezki = ctk.CTkLabel(input_frame_path, text="Katalog zapisu pliku CSV:")
        etykieta_sciezki.grid(row=0, column=0, padx=(10,5), pady=5, sticky="w")
        
        self.pole_sciezki_zapisu = ctk.CTkEntry(input_frame_path)
        self.pole_sciezki_zapisu.grid(row=0, column=1, padx=(0,5), pady=5, sticky="ew")
        self.pole_sciezki_zapisu.insert(0, DOMYSLNA_SCIEZKA_ZAPISU)

        przycisk_wybierz_sciezke = ctk.CTkButton(input_frame_path, text="Wybierz folder", width=120, command=self.wybierz_katalog_zapisu)
        przycisk_wybierz_sciezke.grid(row=0, column=2, padx=(0,10), pady=5, sticky="e")

        self.przycisk_przetworz = ctk.CTkButton(self, text="Przetwórz na JEDEN plik CSV", command=self.rozpocznij_przetwarzanie_action, height=40)
        self.przycisk_przetworz.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self, height=10)
        self.progress_bar.grid(row=3, column=0, padx=20, pady=(0, 5), sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(self, text="Gotowy.", text_color="gray")
        self.status_label.grid(row=4, column=0, padx=20, pady=(0,10), sticky="ew")

    def wybierz_katalog_zapisu(self):
        """Otwiera okno dialogowe do wyboru katalogu zapisu."""
        sciezka_katalogu = filedialog.askdirectory(initialdir=self.pole_sciezki_zapisu.get() or DOMYSLNA_SCIEZKA_ZAPISU)
        if sciezka_katalogu:
            self.pole_sciezki_zapisu.delete(0, ctk.END)
            self.pole_sciezki_zapisu.insert(0, sciezka_katalogu)

    def update_status(self, message, progress_value=None):
        """Aktualizuje etykietę statusu i pasek postępu."""
        self.status_label.configure(text=message)
        if progress_value is not None:
            self.progress_bar.set(progress_value)
        self.update_idletasks() 

    def reset_gui_after_delay(self, delay_ms=4000):
        """Resetuje pasek postępu i status po opóźnieniu."""
        self.after(delay_ms, lambda: self.update_status("Gotowy.", 0))
        self.przycisk_przetworz.configure(state="normal", text="Przetwórz na JEDEN plik CSV")

    def rozpocznij_przetwarzanie_action(self):
        """Pobiera nazwy plików, buduje URL-e i rozpoczyna proces."""
        # --- ZMIANA --- Odczytujemy nazwy, a nie pełne URL-e
        nazwy_text = self.pole_url.get("1.0", ctk.END) 
        nazwy = [nazwa.strip() for nazwa in nazwy_text.splitlines() if nazwa.strip()] 

        sciezka_zapisu_csv_gui = self.pole_sciezki_zapisu.get().strip()
        if not sciezka_zapisu_csv_gui:
            sciezka_zapisu_csv_gui = DOMYSLNA_SCIEZKA_ZAPISU
            self.pole_sciezki_zapisu.insert(0, sciezka_zapisu_csv_gui)
        
        # --- ZMIANA --- Budowanie listy URL-i na podstawie podanych nazw
        if not nazwy:
            messagebox.showerror("Błąd", "Musisz podać co najmniej jedną nazwę pliku.")
            return
            
        urls = [f"https://sm-prods.com/feeds/{nazwa}_zero.xml" for nazwa in nazwy]

        self.przycisk_przetworz.configure(state="disabled", text="Przetwarzanie...")
        self.update_idletasks()
        
        przetworz_wiele_url_jeden_plik(urls, sciezka_zapisu_csv_gui, self)

if __name__ == "__main__":
    app = App()
    app.mainloop()