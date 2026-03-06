import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- Konfiguracja ---
# Nazwa kolumny, według której będą usuwane duplikaty
id_column_name = 'id'
# Lista kolumn, które mają pozostać w pliku wynikowym
columns_to_keep = ['id', 'weight', 'EAN']
# --- Koniec konfiguracji ---

def process_excel_logic(input_path, status_callback):
    """
    Zawiera logikę przetwarzania pliku.
    :param input_path: Ścieżka do pliku wejściowego.
    :param status_callback: Funkcja do raportowania postępu w GUI.
    """
    try:
        # Krok 1: Wczytanie pliku Excel
        status_callback(f"Wczytywanie pliku: {os.path.basename(input_path)}...")
        df = pd.read_excel(input_path)
        status_callback("Plik wczytany pomyślnie.")

        initial_rows = len(df)
        status_callback(f"Liczba wszystkich wierszy: {initial_rows}")

        # Krok 2: Usunięcie duplikatów
        status_callback(f"Usuwanie duplikatów na podstawie kolumny '{id_column_name}'...")
        df.drop_duplicates(subset=[id_column_name], keep='first', inplace=True)

        rows_after_deduplication = len(df)
        status_callback(f"Usunięto {initial_rows - rows_after_deduplication} zduplikowanych wierszy.")

        # Krok 3: Wybranie określonych kolumn
        status_callback(f"Wybieranie kolumn: {', '.join(columns_to_keep)}...")
        missing_columns = [col for col in columns_to_keep if col not in df.columns]
        if missing_columns:
            raise ValueError(f"W pliku brakuje następujących kolumn: {', '.join(missing_columns)}")
            
        df_final = df[columns_to_keep]

        # Krok 4: Zapisanie wyniku
        dir_name = os.path.dirname(input_path)
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_filename = f"{base_name}_oczyszczone.xlsx"
        output_path = os.path.join(dir_name, output_filename)
        
        status_callback(f"Zapisywanie wyniku do pliku: {output_filename}...")
        df_final.to_excel(output_path, index=False)
        
        return True, f"Operacja zakończona sukcesem!\n\nPlik zapisano jako:\n{output_path}"

    except Exception as e:
        return False, f"Wystąpił błąd:\n\n{e}"

class ExcelCleanerApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel - Czyszczenie duplikatów")
        master.geometry("500x300")
        
        self.selected_file_path = ""

        # Styl
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", font=('Helvetica', 10))
        style.configure("TLabel", padding=5, font=('Helvetica', 10))
        style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))

        # Ramka główna
        main_frame = ttk.Frame(master, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Elementy GUI
        ttk.Label(main_frame, text="Narzędzie do czyszczenia plików Excel", style="Header.TLabel").pack(pady=(0, 10))
        
        self.file_label = ttk.Label(main_frame, text="Nie wybrano pliku", wraplength=480)
        self.file_label.pack(pady=5)
        
        self.select_button = ttk.Button(main_frame, text="Wybierz plik Excel", command=self.select_file)
        self.select_button.pack(pady=10)

        self.process_button = ttk.Button(main_frame, text="Przetwórz plik", state=tk.DISABLED, command=self.process_file)
        self.process_button.pack(pady=5)
        
        self.status_label = ttk.Label(main_frame, text="Wybierz plik, aby rozpocząć.", justify=tk.LEFT)
        self.status_label.pack(pady=10, fill=tk.X, expand=True)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Wybierz plik Excel",
            filetypes=(("Pliki Excel", "*.xlsx"), ("Wszystkie pliki", "*.*"))
        )
        if file_path:
            self.selected_file_path = file_path
            self.file_label.config(text=f"Wybrano: {os.path.basename(file_path)}")
            self.process_button.config(state=tk.NORMAL)
            self.status_label.config(text="Plik gotowy do przetworzenia.")

    def process_file(self):
        if not self.selected_file_path:
            messagebox.showwarning("Brak pliku", "Najpierw wybierz plik do przetworzenia.")
            return

        self.process_button.config(state=tk.DISABLED)
        self.select_button.config(state=tk.DISABLED)
        
        # Użycie master.after do odświeżenia GUI przed rozpoczęciem przetwarzania
        self.master.after(100, self.run_processing)
        
    def run_processing(self):
        success, message = process_excel_logic(self.selected_file_path, self.update_status)
        
        if success:
            messagebox.showinfo("Sukces", message)
        else:
            messagebox.showerror("Błąd", message)
        
        # Resetowanie stanu GUI
        self.selected_file_path = ""
        self.file_label.config(text="Nie wybrano pliku")
        self.status_label.config(text="Wybierz kolejny plik.")
        self.select_button.config(state=tk.NORMAL)
        self.process_button.config(state=tk.DISABLED)

    def update_status(self, text):
        self.status_label.config(text=text)
        self.master.update_idletasks() # Odświeżenie interfejsu

# Uruchomienie aplikacji GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()

