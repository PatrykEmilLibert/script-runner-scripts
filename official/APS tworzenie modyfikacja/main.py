import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import configparser

# Ustawienie motywu wyglądu (System, Dark, Light)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

def wczytaj_konfiguracje():
    """Wczytuje dane z pliku config.ini z zachowaniem wielkości liter w kluczach."""
    konfiguracja = configparser.ConfigParser()
    konfiguracja.optionxform = str
    sciezka_konfigu = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.ini')
    
    if not os.path.exists(sciezka_konfigu):
        konfiguracja['PROFIL_A'] = {'Kolumna_Przykladowa': 'Wartosc_A'}
        konfiguracja['PROFIL_B'] = {'Kolumna_Inna': 'Wartosc_B'}
        with open(sciezka_konfigu, 'w', encoding='utf-8') as plik_konfiguracyjny:
            plik_konfiguracyjny.write("; Plik konfiguracyjny został utworzony automatycznie.\n\n")
            konfiguracja.write(plik_konfiguracyjny)
        messagebox.showwarning("Informacja", f"Nie znaleziono pliku config.ini.\n\nUtworzono nowy plik z domyślnymi wartościami.\n\nUzupełnij go i uruchom program ponownie.")
        return None
        
    konfiguracja.read(sciezka_konfigu, encoding='utf-8')
    return konfiguracja

def procesuj_plik(sciezka_wejsciowa, sciezka_wyjsciowa, nazwa_sekcji, konfiguracja):
    """
    Dodaje lub modyfikuje kolumny w pliku CSV (z separatorem w postaci średnika).
    """
    try:
        if nazwa_sekcji not in konfiguracja:
            return False, f"Błąd: Brak sekcji [{nazwa_sekcji}] w pliku config.ini."
            
        wartosci_do_zastosowania = dict(konfiguracja[nazwa_sekcji])
        
        df = pd.read_csv(sciezka_wejsciowa, sep=';')
        
        istniejace_kolumny = df.columns.tolist()
        kolumny_dodane = []
        kolumny_zmodyfikowane = []

        for nazwa_kolumny, wartosc in wartosci_do_zastosowania.items():
            if nazwa_kolumny in istniejace_kolumny:
                kolumny_zmodyfikowane.append(nazwa_kolumny)
            else:
                kolumny_dodane.append(nazwa_kolumny)
            
            df[nazwa_kolumny] = wartosc

        df.to_csv(sciezka_wyjsciowa, index=False, sep=';')
        
        komunikaty = []
        if kolumny_dodane:
            komunikaty.append(f"Dodano kolumny: {', '.join(kolumny_dodane)}")
        if kolumny_zmodyfikowane:
            komunikaty.append(f"Zmodyfikowano kolumny: {', '.join(kolumny_zmodyfikowane)}")
        
        if not komunikaty:
            pelny_komunikat = "Nie zdefiniowano żadnych kolumn w wybranym profilu."
        else:
            pelny_komunikat = "\n".join(komunikaty)
        
        pelny_komunikat += f"\n\nPlik zapisano jako:\n{sciezka_wyjsciowa}"
        return True, pelny_komunikat

    except FileNotFoundError:
        return False, f"Błąd: Nie znaleziono pliku:\n{sciezka_wejsciowa}"
    except Exception as e:
        return False, f"Wystąpił nieoczekiwany błąd:\n{e}"

class App(ctk.CTk):
    def __init__(self, konfiguracja):
        super().__init__()

        self.konfiguracja = konfiguracja
        self.title("")
        # ZMIANA: Zwiększono wysokość okna na 750px dla 12 pozycji, kolorystyka różowa
        self.geometry("550x750")
        self.configure(fg_color="#FFB6D9")  # Light pink background
        
        # --- Custom Header ---
        self.header_frame = ctk.CTkFrame(self, fg_color="#FF1493", height=50)
        self.header_frame.grid(row=0, column=0, padx=0, pady=0, sticky="ew")
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_label = ctk.CTkLabel(self.header_frame, text="Nowoczesny Edytor CSV (separator: średnik)", text_color="white", font=("Arial", 14, "bold"))
        self.header_label.pack(padx=10, pady=10)
        
        self.grid_columnconfigure(0, weight=1)
        # ZMIANA: Ustawienie wagi dla ramki z opcjami - row 3 jest rozciągliwy
        self.grid_rowconfigure(3, weight=1)
        
        # --- Ramka pliku wejściowego ---
        self.frame_plik = ctk.CTkFrame(self, fg_color="#FFE4F0", border_width=2, border_color="#FF1493")
        self.frame_plik.grid(row=1, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.frame_plik.grid_columnconfigure(1, weight=1)
        self.label_plik = ctk.CTkLabel(self.frame_plik, text="Plik wejściowy:", text_color="black")
        self.label_plik.grid(row=0, column=0, padx=(10, 5), pady=10)
        self.entry_plik = ctk.CTkEntry(self.frame_plik, placeholder_text="Wybierz plik CSV do edycji...", border_width=2, border_color="#FF69B4", fg_color="#FFF0F5", text_color="black")
        self.entry_plik.grid(row=0, column=1, padx=5, sticky="ew")
        self.button_plik = ctk.CTkButton(self.frame_plik, text="Wybierz...", command=self.wybierz_plik, width=100, fg_color="#FF69B4", text_color="black")
        self.button_plik.grid(row=0, column=2, padx=(5, 10))

        # --- NOWA SEKCJA: Ramka pliku wyjściowego ---
        self.frame_plik_wyjsciowy = ctk.CTkFrame(self, fg_color="#FFE4F0", border_width=2, border_color="#FF1493")
        self.frame_plik_wyjsciowy.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.frame_plik_wyjsciowy.grid_columnconfigure(1, weight=1)
        self.label_plik_wyjsciowy = ctk.CTkLabel(self.frame_plik_wyjsciowy, text="Plik wyjściowy:", text_color="black")
        self.label_plik_wyjsciowy.grid(row=0, column=0, padx=(10, 5), pady=10)
        self.entry_plik_wyjsciowy = ctk.CTkEntry(self.frame_plik_wyjsciowy, placeholder_text="Wybierz miejsce zapisu...", border_width=2, border_color="#FF69B4", fg_color="#FFF0F5", text_color="black")
        self.entry_plik_wyjsciowy.grid(row=0, column=1, padx=5, sticky="ew")
        self.button_plik_wyjsciowy = ctk.CTkButton(self.frame_plik_wyjsciowy, text="Wybierz...", command=self.wybierz_plik_wyjsciowy, width=100, fg_color="#FF69B4", text_color="black")
        self.button_plik_wyjsciowy.grid(row=0, column=2, padx=(5, 10))

        # --- Ramka z opcjami profili (przewijalna) ---
        self.frame_opcje_outer = ctk.CTkFrame(self, fg_color="#FFE4F0", border_width=2, border_color="#FF1493")
        self.frame_opcje_outer.grid(row=3, column=0, padx=20, pady=10, sticky="nsew") 
        self.frame_opcje_outer.grid_columnconfigure(0, weight=1)
        self.frame_opcje_outer.grid_rowconfigure(1, weight=1)
        
        self.label_opcje = ctk.CTkLabel(self.frame_opcje_outer, text="Wybierz profil z pliku config.ini:", text_color="black")
        self.label_opcje.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        
        # Canvas ze scrollbarem dla radio buttonów
        self.canvas_opcje = ctk.CTkCanvas(self.frame_opcje_outer, bg="#FFE4F0", highlightthickness=0, height=100)
        self.canvas_opcje.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=(0, 10))
        
        self.scrollbar_opcje = ctk.CTkScrollbar(self.frame_opcje_outer, command=self.canvas_opcje.yview)
        self.scrollbar_opcje.grid(row=1, column=1, sticky="ns", padx=(0, 10), pady=(0, 10))
        
        self.canvas_opcje.configure(yscrollcommand=self.scrollbar_opcje.set)
        
        # Wewnętrzna ramka dla radio buttonów
        self.frame_opcje = ctk.CTkFrame(self.canvas_opcje, fg_color="#FFE4F0")
        self.canvas_window = self.canvas_opcje.create_window((0, 0), window=self.frame_opcje, anchor="nw")
        self.canvas_opcje.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas_opcje.bind("<Button-4>", self._on_mousewheel)
        self.canvas_opcje.bind("<Button-5>", self._on_mousewheel)
        
        self.nazwy_sekcji = self.konfiguracja.sections() if self.konfiguracja else []
        self.wybrana_opcja = ctk.StringVar()
        if not self.nazwy_sekcji:
            ctk.CTkLabel(self.frame_opcje, text="Brak profili w pliku config.ini!", text_color="black").pack(padx=30, anchor="w")
        else:
            for nazwa in self.nazwy_sekcji:
                ctk.CTkRadioButton(self.frame_opcje, text=nazwa, variable=self.wybrana_opcja, value=nazwa, text_color="black", fg_color="#FF69B4").pack(padx=30, pady=5, anchor="w")
            self.wybrana_opcja.set(self.nazwy_sekcji[0])
        
        # Aktualizuj wysokość canvas
        self.frame_opcje.update_idletasks()
        self.canvas_opcje.configure(scrollregion=self.canvas_opcje.bbox("all"))
        
        # Bind resize event
        self.frame_opcje.bind("<Configure>", self._on_frame_opcje_configure)
        
        # --- Przycisk uruchomienia ---
        self.button_uruchom = ctk.CTkButton(self, text="Przetwórz Plik", command=self.uruchom_przetwarzanie, height=40, fg_color="#FF1493", text_color="black")
        self.button_uruchom.grid(row=4, column=0, padx=20, pady=(10, 20), sticky="ew") 
        if self.konfiguracja is None or not self.nazwy_sekcji:
            self.button_uruchom.configure(state="disabled")

    def _on_mousewheel(self, event):
        """Obsługuje przewijanie canvas scrollbarem za pomocą kółka myszy."""
        self.canvas_opcje.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_frame_opcje_configure(self, event):
        """Aktualizuje scrollregion canvas gdy zmieni się zawartość."""
        self.canvas_opcje.configure(scrollregion=self.canvas_opcje.bbox("all"))
        
        # Aktualizuj szerokość wewnętrznej ramki
        self.canvas_opcje.itemconfig(self.canvas_window, width=self.frame_opcje_outer.winfo_width() - 30)

    def wybierz_plik(self):
        sciezka = filedialog.askopenfilename(title="Wybierz plik CSV", filetypes=(("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*")))
        if sciezka:
            self.entry_plik.delete(0, ctk.END)
            self.entry_plik.insert(0, sciezka)

            # ZMIANA: Automatycznie sugeruj nazwę i ścieżkę pliku wyjściowego
            sciezka_bazowa, _ = os.path.splitext(sciezka)
            nazwa_bazowa = os.path.basename(sciezka_bazowa)
            # Bierz pierwszy człon przed "_" i dodaj "_add_per_sku"
            pierwszy_czlon = nazwa_bazowa.split('_')[0]
            katalog = os.path.dirname(sciezka)
            sciezka_wyjsciowa_sugerowana = os.path.join(katalog, f"{pierwszy_czlon}_add_per_sku.csv")
            self.entry_plik_wyjsciowy.delete(0, ctk.END)
            self.entry_plik_wyjsciowy.insert(0, sciezka_wyjsciowa_sugerowana)

    # --- NOWA METODA ---
    def wybierz_plik_wyjsciowy(self):
        """Otwiera okno dialogowe 'Zapisz jako'."""
        sciezka_wejsciowa = self.entry_plik.get()
        if sciezka_wejsciowa:
            sciezka_bazowa, _ = os.path.splitext(sciezka_wejsciowa)
            nazwa_bazowa = os.path.basename(sciezka_bazowa)
            # Bierz pierwszy człon przed "_" i dodaj "_add_per_sku"
            pierwszy_czlon = nazwa_bazowa.split('_')[0]
            sugerowana_nazwa = f"{pierwszy_czlon}_add_per_sku.csv"
            katalog_poczatkowy = os.path.dirname(sciezka_wejsciowa)
        else:
            sugerowana_nazwa = "wynik.csv"
            katalog_poczatkowy = os.path.expanduser("~") # Katalog domowy użytkownika

        sciezka = filedialog.asksaveasfilename(
            title="Zapisz plik jako...",
            initialdir=katalog_poczatkowy,
            initialfile=sugerowana_nazwa,
            defaultextension=".csv",
            filetypes=(("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*"))
        )
        if sciezka:
            self.entry_plik_wyjsciowy.delete(0, ctk.END)
            self.entry_plik_wyjsciowy.insert(0, sciezka)

    def uruchom_przetwarzanie(self):
        sciezka_wejsciowa = self.entry_plik.get()
        # ZMIANA: Pobierz ścieżkę wyjściową z nowego pola
        sciezka_wyjsciowa = self.entry_plik_wyjsciowy.get() 
        
        if not sciezka_wejsciowa:
            messagebox.showerror("Błąd", "Proszę wybrać plik wejściowy CSV!")
            return
        
        # ZMIANA: Dodano walidację dla ścieżki wyjściowej
        if not sciezka_wyjsciowa:
            messagebox.showerror("Błąd", "Proszę wybrać ścieżkę zapisu dla pliku wyjściowego!")
            return

        wybrana_nazwa_sekcji = self.wybrana_opcja.get()
        if not wybrana_nazwa_sekcji:
            messagebox.showerror("Błąd", "Proszę wybrać profil do zastosowania!")
            return
        
        # ZMIANA: Usunięto automatyczne generowanie ścieżki wyjściowej
        sukces, komunikat = procesuj_plik(sciezka_wejsciowa, sciezka_wyjsciowa, wybrana_nazwa_sekcji, self.konfiguracja)
        
        if sukces:
            messagebox.showinfo("Sukces", komunikat)
        else:
            messagebox.showerror("Błąd", komunikat)

if __name__ == "__main__":
    konfiguracja = wczytaj_konfiguracje()
    app = App(konfiguracja)
    app.mainloop()