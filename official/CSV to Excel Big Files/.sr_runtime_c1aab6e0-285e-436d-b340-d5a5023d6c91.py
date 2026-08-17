import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# Ustawienie wyglądu CustomTkinter
ctk.set_appearance_mode("System")  # Tryby: "System" (domyślny), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Motywy: "blue" (domyślny), "dark-blue", "green"

class CSVToExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Import CSV do Excela")
        self.geometry("800x400") # Zwiększona szerokość okna GUI
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Ramka główna
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(1, weight=1) # Allow second column to expand

        # Tytuł
        self.title_label = ctk.CTkLabel(self.main_frame, text="Import danych CSV do pliku Excel", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, columnspan=3, pady=(20, 30), padx=20)

        # Ścieżka do pliku CSV
        self.csv_label = ctk.CTkLabel(self.main_frame, text="Wybierz pliki CSV:")
        self.csv_label.grid(row=1, column=0, padx=20, pady=5, sticky="w")
        self.csv_path_entry = ctk.CTkEntry(self.main_frame)
        self.csv_path_entry.grid(row=1, column=1, padx=20, pady=5, sticky="ew")
        self.csv_button = ctk.CTkButton(self.main_frame, text="Przeglądaj", command=self.select_csv_files, width=100)
        self.csv_button.grid(row=1, column=2, padx=(0, 20), pady=5)

        # Zmienna do przechowywania listy wybranych plików CSV
        self.selected_csv_files = []

        # Ścieżka do folderu zapisu
        self.output_label = ctk.CTkLabel(self.main_frame, text="Wybierz folder docelowy:")
        self.output_label.grid(row=2, column=0, padx=20, pady=5, sticky="w")
        self.output_path_entry = ctk.CTkEntry(self.main_frame)
        self.output_path_entry.grid(row=2, column=1, padx=20, pady=5, sticky="ew")
        self.output_button = ctk.CTkButton(self.main_frame, text="Przeglądaj", command=self.select_output_folder, width=100)
        self.output_button.grid(row=2, column=2, padx=(0, 20), pady=5)

        # Separator CSV
        self.separator_label = ctk.CTkLabel(self.main_frame, text="Separator CSV (np. ',' lub ';'):")
        self.separator_label.grid(row=3, column=0, padx=20, pady=5, sticky="w")
        self.separator_entry = ctk.CTkEntry(self.main_frame, width=50)
        self.separator_entry.insert(0, ";") # Domyślny separator to średnik
        self.separator_entry.grid(row=3, column=1, padx=20, pady=5, sticky="w")

        # Czy plik CSV ma nagłówki
        self.header_var = ctk.BooleanVar(value=True)
        self.header_checkbox = ctk.CTkCheckBox(self.main_frame, text="Plik CSV ma nagłówki", variable=self.header_var)
        self.header_checkbox.grid(row=4, column=0, columnspan=2, padx=20, pady=10, sticky="w")

        # Przycisk importu
        self.import_button = ctk.CTkButton(self.main_frame, text="Importuj i Zapisz do Excela", command=self.import_and_save, font=ctk.CTkFont(size=16, weight="bold"))
        self.import_button.grid(row=5, column=0, columnspan=3, pady=30, padx=20)

        # Komunikat o statusie
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="green", font=ctk.CTkFont(size=14))
        self.status_label.grid(row=6, column=0, columnspan=3, pady=(0, 20), padx=20)

    def select_csv_files(self):
        """Otwiera okno dialogowe do wyboru jednego lub wielu plików CSV."""
        file_paths = filedialog.askopenfilenames(
            title="Wybierz pliki CSV",
            filetypes=(("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*"))
        )
        if file_paths:
            self.selected_csv_files = list(file_paths)
            self.csv_path_entry.delete(0, ctk.END)
            if len(self.selected_csv_files) == 1:
                self.csv_path_entry.insert(0, self.selected_csv_files[0])
            else:
                self.csv_path_entry.insert(0, f"Wybrano {len(self.selected_csv_files)} plików CSV")
            self.status_label.configure(text="")
        else:
            self.selected_csv_files = []
            self.csv_path_entry.delete(0, ctk.END)
            self.status_label.configure(text="Nie wybrano żadnych plików CSV.", text_color="orange")

    def select_output_folder(self):
        """Otwiera okno dialogowe do wyboru folderu docelowego."""
        folder_path = filedialog.askdirectory(title="Wybierz folder docelowy")
        if folder_path:
            self.output_path_entry.delete(0, ctk.END)
            self.output_path_entry.insert(0, folder_path)
            self.status_label.configure(text="")

    def import_and_save(self):
        """Importuje dane z CSV i zapisuje je do pliku Excela, dzieląc duże pliki na arkusze."""
        csv_files = self.selected_csv_files
        output_folder = self.output_path_entry.get()
        separator = self.separator_entry.get()
        has_header = self.header_var.get()

        # Walidacja wejścia
        if not csv_files:
            messagebox.showerror("Błąd", "Proszę wybrać przynajmniej jeden plik CSV.")
            return
        if not output_folder:
            messagebox.showerror("Błąd", "Proszę wybrać folder docelowy.")
            return
        if not separator:
            messagebox.showerror("Błąd", "Proszę podać separator CSV.")
            return

        # Utwórz folder docelowy, jeśli nie istnieje
        try:
            os.makedirs(output_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się utworzyć folderu docelowego: {e}")
            return

        successful_imports = 0
        failed_imports = []

        self.status_label.configure(text="Przetwarzanie plików...", text_color="blue")
        self.update_idletasks() # Aktualizuj GUI natychmiast

        # *** POCZĄTEK ZMIAN ***
        # Definicja stałych dla podziału plików
        ROWS_PER_SHEET = 900000
        ROW_LIMIT_FOR_SPLIT = 1000000
        # *** KONIEC ZMIAN ***

        for csv_file in csv_files:
            try:
                if not os.path.exists(csv_file):
                    failed_imports.append(f"Nie znaleziono pliku: {os.path.basename(csv_file)}")
                    continue

                # Generowanie nazwy pliku Excela na podstawie nazwy pliku CSV
                csv_base_name = os.path.splitext(os.path.basename(csv_file))[0]
                prefix_parts = csv_base_name.split('_')
                excel_prefix = prefix_parts[0] if prefix_parts else csv_base_name
                excel_file_name = f"{excel_prefix}.xlsx"
                excel_full_path = os.path.join(output_folder, excel_file_name)

                self.status_label.configure(text=f"Wczytywanie: {os.path.basename(csv_file)}...", text_color="blue")
                self.update_idletasks()

                # Odczyt pliku CSV za pomocą pandas
                df = pd.read_csv(csv_file, sep=separator, header=0 if has_header else None, encoding='utf-8', on_bad_lines='skip')

                self.status_label.configure(text=f"Zapisywanie: {excel_file_name}...", text_color="blue")
                self.update_idletasks()

                # *** POCZĄTEK ZMIAN ***
                # Sprawdź, czy liczba wierszy przekracza próg podziału
                if len(df) > ROW_LIMIT_FOR_SPLIT:
                    # Użyj ExcelWritera do zapisu w wielu arkuszach
                    # Do zapisu w wielu arkuszach wymagana jest biblioteka 'openpyxl'
                    # pip install openpyxl
                    with pd.ExcelWriter(excel_full_path, engine='openpyxl') as writer:
                        # Podziel DataFrame na części i zapisz każdą w osobnym arkuszu
                        for i in range(0, len(df), ROWS_PER_SHEET):
                            chunk_df = df.iloc[i:i + ROWS_PER_SHEET]
                            sheet_name = f'Arkusz_{i // ROWS_PER_SHEET + 1}'
                            chunk_df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # Jeśli plik nie jest za duży, zapisz go w jednym arkuszu
                    df.to_excel(excel_full_path, index=False, sheet_name='Arkusz_1')
                # *** KONIEC ZMIAN ***
                
                successful_imports += 1
                self.status_label.configure(text=f"Przetworzono: {os.path.basename(csv_file)}", text_color="blue")
                self.update_idletasks()

            except pd.errors.EmptyDataError:
                failed_imports.append(f"Plik pusty: {os.path.basename(csv_file)}")
            except pd.errors.ParserError as e:
                failed_imports.append(f"Błąd parsowania w {os.path.basename(csv_file)}: {e}")
            except Exception as e:
                failed_imports.append(f"Ogólny błąd w {os.path.basename(csv_file)}: {e}")

        # Wyświetl podsumowanie po przetworzeniu wszystkich plików
        final_message = f"Zakończono przetwarzanie. Pomyślnie zaimportowano {successful_imports} plików."
        if failed_imports:
            final_message += f"\n\nBłędy w {len(failed_imports)} plikach:\n" + "\n".join(failed_imports)
            self.status_label.configure(text="Zakończono z błędami.", text_color="red")
            messagebox.showwarning("Zakończono z błędami", final_message)
        else:
            self.status_label.configure(text=f"Pomyślnie zaimportowano wszystkie {successful_imports} plików.", text_color="green")
            messagebox.showinfo("Sukces", final_message)

if __name__ == "__main__":
    app = CSVToExcelApp()
    app.mainloop()

