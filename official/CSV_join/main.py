import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox

class CSVMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Łączenie kolumn CSV po nazwie")
        self.geometry("550x420")

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        self.grid_columnconfigure(0, weight=1)

        # --- Elementy GUI ---
        self.files_label = ctk.CTkLabel(self, text="Nie wybrano żadnych plików", wraplength=480)
        self.files_label.grid(row=0, column=0, columnspan=2, padx=20, pady=(10, 5), sticky="ew")

        self.select_files_button = ctk.CTkButton(self, text="Wybierz pliki CSV", command=self.select_files)
        self.select_files_button.grid(row=1, column=0, columnspan=2, padx=20, pady=5)

        self.column_label = ctk.CTkLabel(self, text="Podaj dokładną nazwę kolumny (nagłówek):")
        self.column_label.grid(row=2, column=0, columnspan=2, padx=20, pady=(20, 5))

        self.column_entry = ctk.CTkEntry(self, placeholder_text="np. 'Produkt' lub 'ID_Transakcji'")
        self.column_entry.grid(row=3, column=0, columnspan=2, padx=20, pady=5, sticky="ew")
        
        # NOWE POLE: WYBÓR SEPARATORA
        self.sep_label = ctk.CTkLabel(self, text="Separator w plikach CSV:")
        self.sep_label.grid(row=4, column=0, padx=(20, 5), pady=(20, 5), sticky="w")
        
        self.sep_entry = ctk.CTkEntry(self, width=100)
        self.sep_entry.insert(0, ",") # Domyślna wartość to przecinek
        self.sep_entry.grid(row=4, column=1, padx=(5, 20), pady=(20, 5), sticky="w")

        self.merge_button = ctk.CTkButton(self, text="Połącz i zapisz", command=self.merge_and_save)
        self.merge_button.grid(row=5, column=0, columnspan=2, padx=20, pady=20)

        self.selected_files = []

    def _read_csv_with_fallback(self, file_path, delimiter):
        encodings_to_try = ["utf-8-sig", "utf-8", "cp1250", "latin1"]
        last_error = None

        for enc in encodings_to_try:
            try:
                return pd.read_csv(
                    file_path,
                    header=0,
                    sep=delimiter,
                    engine="python",
                    encoding=enc,
                )
            except UnicodeDecodeError as exc:
                last_error = exc

        raise last_error

    def select_files(self):
        self.selected_files = filedialog.askopenfilenames(
            title="Wybierz pliki CSV",
            filetypes=[("Pliki CSV", "*.csv")]
        )
        if self.selected_files:
            self.files_label.configure(text=f"Wybrano {len(self.selected_files)} plików.")
        else:
            self.files_label.configure(text="Nie wybrano żadnych plików")

    def merge_and_save(self):
        if not self.selected_files:
            messagebox.showwarning("Brak plików", "Najpierw wybierz pliki CSV!")
            return

        column_name = self.column_entry.get()
        if not column_name:
            messagebox.showwarning("Brak nazwy kolumny", "Musisz podać nazwę kolumny do scalenia.")
            return
            
        # Pobranie separatora z pola
        delimiter = self.sep_entry.get()
        if not delimiter:
            messagebox.showwarning("Brak separatora", "Musisz podać znak separatora (np. ',' lub ';').")
            return

        all_columns_data = []

        try:
            for file_path in self.selected_files:
                # Użycie podanego separatora podczas wczytywania pliku
                df = self._read_csv_with_fallback(file_path, delimiter)
                
                if column_name in df.columns:
                    all_columns_data.append(df[column_name])
                else:
                    messagebox.showerror(
                        "Błąd: Brak kolumny",
                        f"W pliku:\n'{file_path.split('/')[-1]}'\n\n"
                        f"nie znaleziono kolumny o nazwie: '{column_name}'."
                    )
                    return
            
            merged_series = pd.concat(all_columns_data, ignore_index=True)
            final_df = pd.DataFrame(merged_series)
            final_df.columns = [column_name]

            save_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("Pliki CSV", "*.csv")],
                title="Zapisz połączony plik jako..."
            )

            if save_path:
                final_df.to_csv(save_path, index=False)
                messagebox.showinfo("Sukces", f"Plik został pomyślnie zapisany w lokalizacji:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Wystąpił błąd", f"Nie udało się przetworzyć plików.\nSzczegóły: {e}")

if __name__ == "__main__":
    app = CSVMergerApp()
    app.mainloop()