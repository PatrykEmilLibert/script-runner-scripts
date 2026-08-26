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

# --- Standardowa ścieżka zapisu ---
DOMYSLNA_SCIEZKA_ZAPISU = os.path.join(os.path.expanduser("~"), "Downloads")
# ----------------------------------------------------

def clean_text(text):
    """Zastępuje znaki nowej linii i inne białe znaki pojedynczą spacją."""
    if not text:
        return ""
    # Rozdziela tekst po białych znakach (spacje, tabulatory, entery),
    # łączy je z powrotem pojedynczą spacją i usuwa wiodące/końcowe białe znaki.
    return " ".join(text.split())

def pobierz_xml(url, sciezka_docelowa):
    """Pobiera plik XML z podanego URL."""
    try:
        # Użycie nagłówka User-Agent, aby uniknąć blokowania przez niektóre serwery
        opener = urllib.request.build_opener()
        opener.addheaders = [('User-agent', 'Mozilla/5.0')]
        urllib.request.install_opener(opener)
        urllib.request.urlretrieve(url, sciezka_docelowa)
        return True
    except Exception as e:
        messagebox.showerror(f"Błąd pobierania ({url})", f"Wystąpił błąd: {e}")
        return False

def parsuj_xml(sciezka_pliku):
    """Parsuje plik XML i ekstrahuje dane, usuwając znaki nowej linii z pól tekstowych."""
    try:
        drzewo = ET.parse(sciezka_pliku)
        root = drzewo.getroot()
        dane = []

        for element in root.findall("o"):
            cat_elem = element.find("cat")
            name_elem = element.find("name")
            desc_elem = element.find("desc")

            # Utworzenie wiersza danych z zastosowaniem funkcji clean_text
            wiersz = {
                "id": element.get("id"),
                "url": element.get("url"),
                "price": element.get("price"),
                "avail": element.get("avail"),
                "weight": element.get("weight"),
                "stock": element.get("stock"),
                "cat": clean_text(cat_elem.text) if cat_elem is not None else "",
                "name": clean_text(name_elem.text) if name_elem is not None else "",
                "desc": clean_text(desc_elem.text) if desc_elem is not None else ""
            }
            
            # Przetwarzanie atrybutów z czyszczeniem ich wartości
            atrybuty_elem = element.find("attrs")
            if atrybuty_elem is not None:
                for atrybut in atrybuty_elem.findall("a"):
                    nazwa_atrybutu = atrybut.get("name")
                    if nazwa_atrybutu:
                        wiersz[nazwa_atrybutu] = clean_text(atrybut.text)

            # Przetwarzanie obrazów
            obrazy_elem = element.find("imgs")
            if obrazy_elem is not None:
                main_image = obrazy_elem.find("main")
                if main_image is not None and main_image.get("url"):
                    wiersz["image0"] = main_image.get("url")
                
                start_index = 1 if "image0" in wiersz else 0
                
                # Ograniczenie do pobierania tylko image1 (jeśli image0 istnieje)
                # lub image0 i image1 (jeśli main image nie istnieje)
                other_images = obrazy_elem.findall("i")
                for i, img in enumerate(other_images, start=start_index):
                    if i >= 2: # Zbieramy tylko image0 i image1
                        break
                    if img.get("url"):
                        wiersz[f"image{i}"] = img.get("url")
            
            dane.append(wiersz)
            
        # Zwracamy tylko dane, reszta nie jest potrzebna w nowej logice zapisu
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
    """Zapisuje dane do pliku CSV z predefiniowanymi kolumnami."""
    # Definiujemy na stałe, które kolumny mają się znaleźć w pliku wynikowym
    pola = ['id', 'weight', 'cat', 'name', 'desc', 'id_bl', 'sku_bl', 'EAN', 'Kod_producenta', 'image0', 'image1']

    try:
        with open(sciezka_pliku, 'w', encoding='utf-8-sig', newline='') as plik_csv:
            # Użycie separatora, extrasaction='ignore' pomija nadmiarowe dane w słowniku
            writer = csv.DictWriter(plik_csv, fieldnames=pola, delimiter='|', extrasaction='ignore')
            writer.writeheader()
            writer.writerows(dane)
        return True
    except Exception as e:
        messagebox.showerror(f"Błąd zapisu CSV ({os.path.basename(sciezka_pliku)})", f"Wystąpił błąd: {e}")
        return False

def przetworz_wiele_url_osobne_pliki(urls, sciezka_zapisu_csv, app_instance):
    """Przetwarza wiele URL-i i zapisuje każdy do osobnego pliku CSV."""
    katalog_tymczasowy = tempfile.gettempdir()
    liczba_url = len(urls)
    sukcesy = 0
    bledy_pobierania = 0
    bledy_parsowania = 0
    bledy_zapisu = 0

    if liczba_url == 0:
        app_instance.update_status("Nie podano żadnych URL-i.", 0)
        return

    if not os.path.exists(sciezka_zapisu_csv):
        try:
            os.makedirs(sciezka_zapisu_csv)
        except Exception as e:
            messagebox.showerror("Błąd ścieżki zapisu", f"Nie można utworzyć katalogu: {sciezka_zapisu_csv}\nBłąd: {e}\nPliki będą zapisywane w katalogu roboczym.")
            sciezka_zapisu_csv = os.getcwd()
            app_instance.pole_sciezki_zapisu.delete(0, ctk.END)
            app_instance.pole_sciezki_zapisu.insert(0, sciezka_zapisu_csv)

    for i, url in enumerate(urls):
        url = url.strip()
        if not url:
            continue

        postep = (i + 1) / liczba_url
        nazwa_pliku_url = os.path.basename(urlparse(url).path) or f"feed_{i+1}.xml"
        app_instance.update_status(f"Przetwarzanie {i+1}/{liczba_url}: {nazwa_pliku_url}...", postep)

        nazwa_bazowa_xml = os.path.splitext(nazwa_pliku_url)[0]
        teraz_timestamp_temp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        sciezka_lokalna_xml = os.path.join(katalog_tymczasowy, f"temp_{nazwa_bazowa_xml}_{teraz_timestamp_temp}.xml")

        if pobierz_xml(url, sciezka_lokalna_xml):
            dane = parsuj_xml(sciezka_lokalna_xml)
            if dane:
                teraz_format_czasu_csv = datetime.now().strftime("%d%m%y-%H%M%S")
                nazwa_pliku_csv = os.path.join(sciezka_zapisu_csv, f"{nazwa_bazowa_xml}_{teraz_format_czasu_csv}.csv")
                
                app_instance.update_status(f"Zapisywanie: {os.path.basename(nazwa_pliku_csv)}...", postep)
                if zapisz_do_csv(dane, nazwa_pliku_csv):
                    sukcesy += 1
                else:
                    bledy_zapisu += 1
            else:
                bledy_parsowania += 1
                app_instance.update_status(f"Błąd parsowania {nazwa_pliku_url}. Pomijanie.", postep)
                time.sleep(0.5)
            
            try:
                os.remove(sciezka_lokalna_xml)
            except Exception as e:
                print(f"Ostrzeżenie: Nie udało się usunąć pliku tymczasowego {sciezka_lokalna_xml}: {e}")
        else:
            bledy_pobierania += 1
            app_instance.update_status(f"Błąd pobierania {nazwa_pliku_url}. Pomijanie.", postep)
            time.sleep(0.5)

    if sukcesy > 0:
        messagebox.showinfo("Zakończono przetwarzanie", f"Pomyślnie przetworzono i zapisano {sukcesy} z {liczba_url} plików w:\n{os.path.abspath(sciezka_zapisu_csv)}\n\n"
                                                      f"Błędy pobierania: {bledy_pobierania}\n"
                                                      f"Błędy parsowania: {bledy_parsowania}\n"
                                                      f"Błędy zapisu: {bledy_zapisu}")
    else:
        messagebox.showwarning("Zakończono przetwarzanie", f"Nie udało się pomyślnie przetworzyć żadnego pliku.\n"
                                                          f"Błędy pobierania: {bledy_pobierania}\n"
                                                          f"Błędy parsowania: {bledy_parsowania}\n"
                                                          f"Błędy zapisu: {bledy_zapisu}")

    app_instance.update_status(f"Zakończono. Zapisano: {sukcesy}. Błędy: {bledy_pobierania+bledy_parsowania+bledy_zapisu}", 1)
    app_instance.reset_gui_after_delay()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Konwerter XML do CSV (Wybrane Kolumny)")
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

        etykieta_url = ctk.CTkLabel(input_frame_urls, text="Wklej URL-e plików XML (każdy w nowej linii):")
        etykieta_url.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        self.pole_url = ctk.CTkTextbox(input_frame_urls, height=150) 
        self.pole_url.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

        input_frame_path = ctk.CTkFrame(self)
        input_frame_path.grid(row=1, column=0, padx=20, pady=(5,10), sticky="ew")
        input_frame_path.grid_columnconfigure(1, weight=1)

        etykieta_sciezki = ctk.CTkLabel(input_frame_path, text="Katalog zapisu plików CSV:")
        etykieta_sciezki.grid(row=0, column=0, padx=(10,5), pady=5, sticky="w")
        
        self.pole_sciezki_zapisu = ctk.CTkEntry(input_frame_path)
        self.pole_sciezki_zapisu.grid(row=0, column=1, padx=(0,5), pady=5, sticky="ew")
        self.pole_sciezki_zapisu.insert(0, DOMYSLNA_SCIEZKA_ZAPISU)

        przycisk_wybierz_sciezke = ctk.CTkButton(input_frame_path, text="Wybierz folder", width=120, command=self.wybierz_katalog_zapisu)
        przycisk_wybierz_sciezke.grid(row=0, column=2, padx=(0,10), pady=5, sticky="e")

        self.przycisk_przetworz = ctk.CTkButton(self, text="Przetwórz na pliki CSV", command=self.rozpocznij_przetwarzanie_action, height=40)
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
        self.przycisk_przetworz.configure(state="normal", text="Przetwórz na pliki CSV")

    def rozpocznij_przetwarzanie_action(self):
        """Rozpoczyna proces przetwarzania wielu URL-i na osobne pliki."""
        urls_text = self.pole_url.get("1.0", ctk.END) 
        urls = [url.strip() for url in urls_text.splitlines() if url.strip()] 

        sciezka_zapisu_csv_gui = self.pole_sciezki_zapisu.get().strip()
        if not sciezka_zapisu_csv_gui:
            sciezka_zapisu_csv_gui = DOMYSLNA_SCIEZKA_ZAPISU
            self.pole_sciezki_zapisu.insert(0, sciezka_zapisu_csv_gui)

        if not urls:
            messagebox.showerror("Błąd", "Musisz podać co najmniej jeden URL pliku XML.")
            return

        self.przycisk_przetworz.configure(state="disabled", text="Przetwarzanie...")
        self.update_idletasks()
        
        przetworz_wiele_url_osobne_pliki(urls, sciezka_zapisu_csv_gui, self)

if __name__ == "__main__":
    app = App()
    app.mainloop()
