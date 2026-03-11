import os
import customtkinter as ctk
import xlwings as xw
from tkinter import filedialog, messagebox

# --- GŁÓWNA LOGIKA ŁĄCZENIA PLIKÓW (z użyciem xlwings) ---

def polacz_pliki_z_xlwings(folder_z_plikami, liczba_wierszy_naglowka, status_label, root):
    """
    Łączy pliki Excel za pomocą automatyzacji aplikacji Excel (xlwings).
    Wymaga zainstalowanego MS Excel.
    Aktualizuje etykietę statusu w GUI.
    """
    try:
        # Uruchom aplikację Excel w tle (będzie niewidoczna)
        with xw.App(visible=False, add_book=False) as app:
            app.display_alerts = False  # Wyłącz komunikaty typu "czy zapisać?"

            pliki_excel = [os.path.join(folder_z_plikami, f) for f in sorted(os.listdir(folder_z_plikami)) if f.endswith(('.xlsx', '.xls'))]
            if not pliki_excel:
                messagebox.showinfo("Informacja", "Nie znaleziono plików Excel w wybranym folderze.")
                status_label.configure(text="Zakończono. Nie znaleziono plików.")
                return

            # Utwórz nowy skoroszyt docelowy
            wb_docelowy = app.books.add()
            ws_docelowy = wb_docelowy.sheets[0]
            ws_docelowy.name = "Połączone Dane"
            
            nastepny_wolny_wiersz = 1
            naglowek_skopiowany = False

            # Przetwarzaj każdy plik po kolei
            for i, plik_path in enumerate(pliki_excel):
                # Aktualizacja statusu w interfejsie graficznym
                status_label.configure(text=f"Przetwarzanie ({i+1}/{len(pliki_excel)}): {os.path.basename(plik_path)}")
                root.update_idletasks() # Wymuś odświeżenie GUI, aby uniknąć "zamrożenia"

                try:
                    wb_zrodlowy = app.books.open(plik_path)
                    ws_zrodlowy = wb_zrodlowy.sheets[0]

                    # Skopiuj blok nagłówka (tylko z pierwszego pliku)
                    if not naglowek_skopiowany and liczba_wierszy_naglowka > 0:
                        zakres_naglowka = ws_zrodlowy.range((1, 1), (liczba_wierszy_naglowka, ws_zrodlowy.used_range.last_cell.column))
                        zakres_naglowka.copy(ws_docelowy.range((1, 1)))
                        nastepny_wolny_wiersz = liczba_wierszy_naglowka + 1
                        naglowek_skopiowany = True
                    
                    # 🎯 ZNAJDŹ OSTATNI WIERSZ I KOLUMNĘ W NIEZAWODNY SPOSÓB 🎯
                    # =======================================================
                    # <-- ZMIANA TUTAJ
                    ostatni_wiersz = ws_zrodlowy.range(f'A{ws_zrodlowy.cells.last_cell.row}').end('up').row
                    # =======================================================
                    
                    ostatnia_kolumna = ws_zrodlowy.used_range.last_cell.column
                    
                    start_wiersza_danych = liczba_wierszy_naglowka + 1
                    
                    if ostatni_wiersz >= start_wiersza_danych:
                        # Zdefiniuj zakres danych do skopiowania
                        zakres_danych = ws_zrodlowy.range((start_wiersza_danych, 1), (ostatni_wiersz, ostatnia_kolumna))
                        
                        # Skopiuj dane do pierwszego wolnego wiersza w pliku docelowym
                        zakres_danych.copy(ws_docelowy.range(f'A{nastepny_wolny_wiersz}'))
                        
                        nastepny_wolny_wiersz += zakres_danych.rows.count

                    wb_zrodlowy.close()

                except Exception as e:
                    messagebox.showwarning("Błąd pliku", f"Nie udało się przetworzyć pliku {os.path.basename(plik_path)}: {e}")

            # Zapisz i zamknij plik docelowy
            sciezka_wyjsciowa = os.path.join(folder_z_plikami, "POLACZONE_PLIKI.xlsx")
            wb_docelowy.save(sciezka_wyjsciowa)
            wb_docelowy.close()
            
            status_label.configure(text="Gotowe!")
            messagebox.showinfo("Sukces", f"Pliki połączone pomyślnie do:\n{sciezka_wyjsciowa}")

    except Exception as e:
        status_label.configure(text="Wystąpił błąd krytyczny.")
        messagebox.showerror("Błąd krytyczny", f"Wystąpił błąd z aplikacją Excel: {e}\n\nUpewnij się, że Excel jest poprawnie zainstalowany i zamknięty.")


# --- INTERFEJS GRAFICZNY (z użyciem CustomTkinter) ---

# Ustawienia wyglądu
ctk.set_appearance_mode("System")  # Tryby: "System", "Dark", "Light"
ctk.set_default_color_theme("blue") # Motywy: "blue", "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Konfiguracja okna
        self.title("Łączenie Plików Excel (xlwings)")
        self.geometry("600x400")
        self.grid_columnconfigure(0, weight=1)

        # Ramka wyboru folderu
        self.frame_folder = ctk.CTkFrame(self)
        self.frame_folder.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.frame_folder.grid_columnconfigure(1, weight=1)

        self.label_folder = ctk.CTkLabel(self.frame_folder, text="Folder z plikami:")
        self.label_folder.grid(row=0, column=0, padx=10, pady=10)
        self.entry_folder = ctk.CTkEntry(self.frame_folder, placeholder_text="Wybierz folder zawierający pliki Excel...")
        self.entry_folder.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.button_browse = ctk.CTkButton(self.frame_folder, text="Przeglądaj", width=100, command=self.wybierz_folder)
        self.button_browse.grid(row=0, column=2, padx=10, pady=10)

        # Ramka opcji
        self.frame_options = ctk.CTkFrame(self)
        self.frame_options.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.label_header = ctk.CTkLabel(self.frame_options, text="Ile wierszy od góry tworzy nagłówek?", anchor="w")
        self.label_header.pack(padx=10, pady=(10, 5), fill="x")
        self.entry_header_rows = ctk.CTkEntry(self.frame_options, width=120)
        self.entry_header_rows.insert(0, "1")
        self.entry_header_rows.pack(padx=10, pady=(0, 10), anchor="w")

        # Główny przycisk akcji
        self.button_merge = ctk.CTkButton(self, text="Połącz Pliki", height=40, font=("", 14, "bold"), command=self.uruchom_polaczenie)
        self.button_merge.grid(row=2, column=0, padx=20, pady=20, sticky="ew")

        # Etykieta statusu
        self.status_label = ctk.CTkLabel(self, text="Wybierz folder i kliknij 'Połącz Pliki'", text_color="gray")
        self.status_label.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
    
    def wybierz_folder(self):
        folder_wybrany = filedialog.askdirectory(title="Wybierz folder z plikami Excel")
        if folder_wybrany:
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, folder_wybrany)
            self.status_label.configure(text="Folder wybrany. Gotowy do łączenia.")

    def uruchom_polaczenie(self):
        folder = self.entry_folder.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Błąd", "Nie wybrano prawidłowego folderu.")
            return

        try:
            naglowek = int(self.entry_header_rows.get())
            if naglowek < 0:
                messagebox.showerror("Błąd", "Liczba wierszy nagłówka musi być nieujemna.")
                return
        except (ValueError, TypeError):
            messagebox.showerror("Błąd", "Nieprawidłowa liczba wierszy nagłówka. Wpisz liczbę całkowitą.")
            return
        
        # Wyłącz przycisk na czas operacji, aby uniknąć wielokrotnego klikania
        self.button_merge.configure(state="disabled", text="Pracuję...")
        try:
            # Uruchom główną funkcję
            polacz_pliki_z_xlwings(folder, naglowek, self.status_label, self)
        finally:
            # Włącz przycisk z powrotem, nawet jeśli wystąpił błąd
            self.button_merge.configure(state="normal", text="Połącz Pliki")

if __name__ == "__main__":
    app = App()
    app.mainloop()