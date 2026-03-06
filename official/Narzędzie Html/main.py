import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import os
import re
import json

# --- Ustawienia CustomTkinter ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# --- Definicje Markerów i Nazw Kolumn ---
# Używamy \u00A0 (non-breaking space) dla unikalności markera, jeśli jest to wymagane.
# Można uprościć do prostszego stringa, jeśli nie ma ryzyka kolizji.
NEW_COMPLEX_MARKER = "\n\u00A0\u00A0\u00A0___HtM1P14CEhO1D3R___\u00A0\u00A0\u00A0\n"
EXTRACTOR_REPLACEMENT_MARKER = NEW_COMPLEX_MARKER
COMBINER_MARKER_FULL_COMPLEX = NEW_COMPLEX_MARKER
COMBINER_MARKER_CORE_ONLY = "___HtM1P14CEhO1D3R___" # Rdzeń markera do wyszukiwania węższego zakresu
CSV_HTML_JSON_COLUMN_NAME = 'HTML_JSON_List'

# --- Funkcja Ekstrahująca HTML ---
def perform_extraction(input_xlsx_path, description_column_number,
                       replacement_marker=EXTRACTOR_REPLACEMENT_MARKER, progress_callback=None):
    """
    Ekstrahuje sekwencje HTML z określonej kolumny pliku XLSX.
    Zapisuje wyekstrahowany HTML jako listę stringów JSON do pliku CSV.
    Zastępuje HTML w oryginalnym pliku XLSX markerem.
    """
    try:
        df = pd.read_excel(input_xlsx_path, header=None)

        if not (0 < description_column_number <= len(df.columns)):
            messagebox.showerror("Błąd ekstrakcji", f"Nieprawidłowy numer kolumny: {description_column_number}.")
            if progress_callback: progress_callback(100, "Błąd")
            return

        description_column_index = description_column_number - 1
        html_json_list_for_csv = []
        modified_descriptions_for_excel = []
        
        # Wzorzec regex do znajdowania sekwencji HTML:
        # ((?:<[^>]+>\s*)+)
        # - Outer (...) creates a capturing group for the whole sequence.
        # - (?:<[^>]+>\s*) matches a single HTML tag (e.g., <p>, <div>, <span>, <br/>)
        #   followed by zero or more whitespace characters. <[^>]+> matches anything inside <...>.
        # - + after the non-capturing group means one or more such tags.
        html_sequence_pattern = r"((?:<[^>]+>\s*)+)"
        total_rows = len(df)

        for index, row in df.iterrows():
            if progress_callback:
                progress_callback(int((index + 1) / total_rows * 100),"Przetwarzanie...")

            current_cell_html_segments = []
            if description_column_index < len(row) and pd.notna(row[description_column_index]):
                description = str(row[description_column_index])
                
                # Znajdowanie wszystkich sekwencji HTML w komórce
                # re.findall zwraca listę stringów, jeśli wzorzec ma jedną grupę przechwytującą
                html_sequences_in_cell = re.findall(html_sequence_pattern, description)
                
                if html_sequences_in_cell:
                    for seq_match in html_sequences_in_cell:
                        cleaned_seq = seq_match.strip() # seq_match jest stringiem
                        current_cell_html_segments.append(cleaned_seq)
                    
                    json_string_for_cell = json.dumps(current_cell_html_segments, ensure_ascii=False)
                    html_json_list_for_csv.append(json_string_for_cell)
                else:
                    html_json_list_for_csv.append("[]") 
                
                # Zastępowanie sekwencji HTML markerem
                modified_description_for_excel_cell, num_substitutions = re.subn(html_sequence_pattern, replacement_marker, description)
                
                # Sprawdzenie (opcjonalne, można usunąć jeśli niepotrzebne bez logowania)
                if len(html_sequences_in_cell) != num_substitutions:
                    # Ta sytuacja może wymagać dokładniejszej analizy wzorca regex lub danych wejściowych
                    # W tej wersji bez logowania, po prostu kontynuujemy.
                    pass

                modified_descriptions_for_excel.append(modified_description_for_excel_cell)
            else: 
                html_json_list_for_csv.append("[]")
                modified_descriptions_for_excel.append(str(row.get(description_column_index, ""))) 
            
        while len(modified_descriptions_for_excel) < len(df):
            modified_descriptions_for_excel.append("")
        while len(html_json_list_for_csv) < len(df):
            html_json_list_for_csv.append("[]")

        output_xlsx_path = os.path.splitext(input_xlsx_path)[0] + "_modified.xlsx"
        modified_df = df.copy()
        if description_column_index < len(modified_df.columns):
            modified_df[modified_df.columns[description_column_index]] = modified_descriptions_for_excel
        modified_df.to_excel(output_xlsx_path, index=False, header=False)

        output_csv_path = os.path.splitext(input_xlsx_path)[0] + "_html_json.csv"
        html_df_for_csv = pd.DataFrame({CSV_HTML_JSON_COLUMN_NAME: html_json_list_for_csv})
        html_df_for_csv.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

        if progress_callback: progress_callback(100, "Zakończono")
        messagebox.showinfo("Sukces ekstrakcji", f"Ekstrakcja zakończona.\nZmodyfikowany XLSX: '{output_xlsx_path}'.\nHTML JSON CSV: '{output_csv_path}'.")

    except Exception as e:
        if progress_callback: progress_callback(100, "Błąd")
        messagebox.showerror("Błąd ekstrakcji", f"Wystąpił błąd: {e}\nTyp: {type(e).__name__}")
    finally:
        if progress_callback: progress_callback(0, "Gotowy")

# --- Funkcja Łącząca HTML ---
def perform_combination(xlsx_path, csv_path, column_number, output_xlsx_path=None,
                        marker_full_complex=COMBINER_MARKER_FULL_COMPLEX,
                        marker_core_only=COMBINER_MARKER_CORE_ONLY,
                        progress_callback=None):
    try:
        df_opisy = pd.read_excel(xlsx_path, header=None)
        df_html_csv = pd.read_csv(csv_path, encoding='utf-8-sig', keep_default_na=False, na_filter=False)

        if df_opisy.empty:
            raise ValueError(f"Plik XLSX '{xlsx_path}' jest pusty.")
        
        if CSV_HTML_JSON_COLUMN_NAME not in df_html_csv.columns and not df_html_csv.empty :
            available_cols = list(df_html_csv.columns)
            raise ValueError(f"Plik CSV '{csv_path}' nie zawiera kolumny '{CSV_HTML_JSON_COLUMN_NAME}'. Dostępne: {available_cols}")
        elif df_html_csv.empty and CSV_HTML_JSON_COLUMN_NAME not in df_html_csv.columns:
            df_html_csv = pd.DataFrame(columns=[CSV_HTML_JSON_COLUMN_NAME])


        if not 1 <= column_number <= len(df_opisy.columns):
            raise ValueError(f"Nieprawidłowy numer kolumny dla XLSX: {column_number}.")

        # Regex do znajdowania pełnego markera lub tylko jego rdzenia
        # Użycie re.escape jest ważne, jeśli markery mogą zawierać specjalne znaki regex
        marker_pattern_str = f"({re.escape(marker_full_complex)}|{re.escape(marker_core_only)})"
        marker_regex = re.compile(marker_pattern_str)

        def replace_markers_in_cell(excel_cell_text, html_segments_list_from_json, cell_row_idx_for_log):
            if not isinstance(excel_cell_text, str):
                excel_cell_text = str(excel_cell_text)
            
            parts = marker_regex.split(excel_cell_text)
            num_markers_in_excel_cell = (len(parts) -1) // 2 
            num_html_segments = len(html_segments_list_from_json)

            if len(parts) == 1: # Nie znaleziono markerów
                # Jeśli są segmenty HTML, ale nie ma markerów, można to zasygnalizować lub zignorować
                if num_html_segments > 0:
                    # messagebox.showwarning("Niezgodność", f"Wiersz {cell_row_idx_for_log + 1}: Znaleziono {num_html_segments} segmentów HTML, ale brak markerów w komórce Excela.")
                    pass # Cicha kontynuacja
                return excel_cell_text 

            if num_html_segments != num_markers_in_excel_cell:
                # messagebox.showwarning("Niezgodność ilości", f"Wiersz {cell_row_idx_for_log + 1}: Liczba segmentów HTML ({num_html_segments}) różni się od liczby markerów ({num_markers_in_excel_cell}).")
                pass # Cicha kontynuacja, proces spróbuje wstawić co może

            new_cell_elements = []
            part_0_text = parts[0]
            new_cell_elements.append(part_0_text)
            
            html_segment_idx_for_this_cell = 0
            
            for i in range(1, len(parts), 2): 
                marker_found_in_xlsx = parts[i] 
                text_after_marker_in_xlsx = parts[i+1] if (i+1) < len(parts) else ""
                
                if html_segment_idx_for_this_cell < num_html_segments:
                    replacement_html = str(html_segments_list_from_json[html_segment_idx_for_this_cell])
                    new_cell_elements.append(replacement_html)
                    html_segment_idx_for_this_cell += 1
                else:
                    # Zabrakło segmentów HTML, wstawiamy oryginalny marker z powrotem
                    new_cell_elements.append(marker_found_in_xlsx)

                new_cell_elements.append(text_after_marker_in_xlsx) 
            
            return "".join(new_cell_elements)

        target_column_index_in_xlsx = column_number - 1
        processed_xlsx_column_data = []
        total_rows = len(df_opisy)

        for xlsx_row_idx in range(total_rows):
            if progress_callback:
                progress_callback(int((xlsx_row_idx + 1) / total_rows * 100), "Przetwarzanie...")

            excel_cell_content = str(df_opisy.iloc[xlsx_row_idx, target_column_index_in_xlsx])
            html_segments_for_current_cell = []

            if xlsx_row_idx < len(df_html_csv):
                if CSV_HTML_JSON_COLUMN_NAME in df_html_csv.columns:
                    json_string_from_csv = df_html_csv.loc[xlsx_row_idx, CSV_HTML_JSON_COLUMN_NAME]
                    try:
                        if json_string_from_csv and isinstance(json_string_from_csv, str) and json_string_from_csv.strip():
                            loaded_json = json.loads(json_string_from_csv)
                            if isinstance(loaded_json, list):
                                html_segments_for_current_cell = loaded_json
                            else:
                                html_segments_for_current_cell = []
                        else:
                            html_segments_for_current_cell = [] 
                    except json.JSONDecodeError:
                        html_segments_for_current_cell = []
                else:
                     html_segments_for_current_cell = []
            else:
                html_segments_for_current_cell = []
            
            processed_cell = replace_markers_in_cell(excel_cell_content, html_segments_for_current_cell, xlsx_row_idx)
            processed_xlsx_column_data.append(processed_cell)
        
        df_opisy.iloc[:, target_column_index_in_xlsx] = processed_xlsx_column_data

        def strip_if_string(value):
            if isinstance(value, str): return value.strip()
            return value
        df_opisy = df_opisy.applymap(strip_if_string)

        output_file_name = os.path.splitext(xlsx_path)[0] + "_combined_data.xlsx"
        if output_xlsx_path:
            output_file_name = output_xlsx_path

        df_opisy.to_excel(output_file_name, index=False, header=False)
        
        if progress_callback: progress_callback(100, "Zakończono")
        messagebox.showinfo("Sukces łączenia", f"Dane połączone.\nPlik wynikowy: '{output_file_name}'.")

    except ValueError as ve:
        if progress_callback: progress_callback(100, "Błąd")
        messagebox.showerror("Błąd łączenia (dane)", f"Wystąpił błąd danych: {ve}")
    except Exception as e:
        if progress_callback: progress_callback(100, "Błąd")
        messagebox.showerror("Błąd łączenia (ogólny)", f"Wystąpił błąd: {e}\nTyp: {type(e).__name__}")
    finally:
        if progress_callback: progress_callback(0, "Gotowy")

# --- Główna Aplikacja Launchera ---
class AppLauncher(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Narzędzia Excel - HTML (v3.0 - Bez Logowania)")
        window_width = 750
        window_height = 520
        self.geometry(f"{window_width}x{window_height}")
        self.center_window(window_width, window_height)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.operation_frame = ctk.CTkFrame(self)
        self.operation_frame.grid(row=0, column=0, padx=20, pady=(20,10), sticky="ew")
        
        self.operation_label = ctk.CTkLabel(self.operation_frame, text="Wybierz operację:", font=ctk.CTkFont(size=14, weight="bold"))
        self.operation_label.pack(side="left", padx=(10,10))
        
        self.operation_var = tk.StringVar(value="Ekstrakcja")
        
        self.radio_extract = ctk.CTkRadioButton(self.operation_frame, text="Ekstrahuj HTML", variable=self.operation_var, value="Ekstrakcja", command=self.update_ui)
        self.radio_extract.pack(side="left", padx=10)
        
        self.radio_combine = ctk.CTkRadioButton(self.operation_frame, text="Połącz/Wstaw HTML", variable=self.operation_var, value="Łączenie", command=self.update_ui)
        self.radio_combine.pack(side="left", padx=10)

        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.extract_ui_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.combine_ui_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")

        self.progress_label = ctk.CTkLabel(self, text="Gotowy", font=ctk.CTkFont(size=10))
        self.progress_label.grid(row=2, column=0, padx=20, pady=(5,0), sticky="ew")

        self.progressbar = ctk.CTkProgressBar(self, mode="determinate")
        self.progressbar.set(0)
        self.progressbar.grid(row=3, column=0, padx=20, pady=(0,10), sticky="ew")

        self.run_button = ctk.CTkButton(self, text="Uruchom wybraną operację", height=40, command=self.run_selected_operation)
        self.run_button.grid(row=4, column=0, padx=20, pady=(10,20), sticky="ew")
        
        self.update_ui()

    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def update_progress(self, value, text):
        self.progressbar.set(value / 100)
        self.progress_label.configure(text=f"{text} ({value}%)")
        self.update_idletasks()

    def setup_extract_ui(self):
        if hasattr(self, 'extract_ui_frame') and self.extract_ui_frame.winfo_exists():
            self.extract_ui_frame.destroy()
            
        self.extract_ui_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.extract_ui_frame.grid(row=0, column=0, sticky="nsew")
        self.extract_ui_frame.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(self.extract_ui_frame, text="Ekstrakcja HTML z pliku XLSX", font=ctk.CTkFont(size=16, weight="bold"))
        title.grid(row=0, column=0, columnspan=3, pady=(10,15), padx=20, sticky="ew")

        ctk.CTkLabel(self.extract_ui_frame, text="Plik XLSX do przetworzenia:", anchor="w").grid(row=1, column=0, padx=(20,5), pady=5, sticky="w")
        self.extract_xlsx_entry = ctk.CTkEntry(self.extract_ui_frame, placeholder_text="Ścieżka do pliku .xlsx")
        self.extract_xlsx_entry.grid(row=1, column=1, padx=(0,5), pady=5, sticky="ew")
        ctk.CTkButton(self.extract_ui_frame, text="Przeglądaj...", width=100, command=lambda: self.browse_file_for_entry(self.extract_xlsx_entry, "xlsx")).grid(row=1, column=2, padx=(0,20), pady=5, sticky="e")

        ctk.CTkLabel(self.extract_ui_frame, text="Numer kolumny z opisami (od 1):", anchor="w").grid(row=2, column=0, padx=(20,5), pady=(5,20), sticky="w")
        self.extract_column_entry = ctk.CTkEntry(self.extract_ui_frame, width=120)
        self.extract_column_entry.grid(row=2, column=1, padx=(0,5), pady=(5,20), sticky="w")

    def setup_combine_ui(self):
        if hasattr(self, 'combine_ui_frame') and self.combine_ui_frame.winfo_exists():
            self.combine_ui_frame.destroy()
            
        self.combine_ui_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.combine_ui_frame.grid(row=0, column=0, sticky="nsew")
        self.combine_ui_frame.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(self.combine_ui_frame, text="Łączenie danych i wstawianie HTML", font=ctk.CTkFont(size=16, weight="bold"))
        title.grid(row=0, column=0, columnspan=3, pady=(10,15), padx=20, sticky="ew")

        ctk.CTkLabel(self.combine_ui_frame, text="Plik XLSX z markerami:", anchor="w").grid(row=1, column=0, padx=(20,5), pady=5, sticky="w")
        self.combine_xlsx_entry = ctk.CTkEntry(self.combine_ui_frame, placeholder_text="Ścieżka do pliku .xlsx (zmodyfikowany)")
        self.combine_xlsx_entry.grid(row=1, column=1, padx=(0,5), pady=5, sticky="ew")
        ctk.CTkButton(self.combine_ui_frame, text="Przeglądaj...", width=100, command=lambda: self.browse_file_for_entry(self.combine_xlsx_entry, "xlsx")).grid(row=1, column=2, padx=(0,20), pady=5, sticky="e")

        ctk.CTkLabel(self.combine_ui_frame, text="Plik CSV z HTML (JSON):", anchor="w").grid(row=2, column=0, padx=(20,5), pady=5, sticky="w")
        self.combine_csv_entry = ctk.CTkEntry(self.combine_ui_frame, placeholder_text="Ścieżka do pliku .csv (z JSON)")
        self.combine_csv_entry.grid(row=2, column=1, padx=(0,5), pady=5, sticky="ew")
        ctk.CTkButton(self.combine_ui_frame, text="Przeglądaj...", width=100, command=lambda: self.browse_file_for_entry(self.combine_csv_entry, "csv")).grid(row=2, column=2, padx=(0,20), pady=5, sticky="e")

        ctk.CTkLabel(self.combine_ui_frame, text="Numer kolumny do podmiany (od 1):", anchor="w").grid(row=3, column=0, padx=(20,5), pady=(5,20), sticky="w")
        self.combine_column_entry = ctk.CTkEntry(self.combine_ui_frame, width=120)
        self.combine_column_entry.grid(row=3, column=1, padx=(0,5), pady=(5,20), sticky="w")

    def update_ui(self):
        self.update_progress(0, "Gotowy")
        operation = self.operation_var.get()
        
        self.extract_ui_frame.grid_remove()
        self.combine_ui_frame.grid_remove()

        if operation == "Ekstrakcja":
            self.setup_extract_ui()
            self.extract_ui_frame.grid()
        elif operation == "Łączenie":
            self.setup_combine_ui()
            self.combine_ui_frame.grid()

    def browse_file_for_entry(self, entry_widget, file_type):
        if file_type == "xlsx": filetypes = [("Pliki Excel", "*.xlsx"), ("Wszystkie pliki", "*.*")]
        elif file_type == "csv": filetypes = [("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*")]
        else: 
            tk.messagebox.showerror("Błąd", f"Nieobsługiwany typ pliku: {file_type}")
            return
            
        filename = filedialog.askopenfilename(title=f"Wybierz plik {file_type.upper()}", filetypes=filetypes, defaultextension=f".{file_type}")
        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)

    def run_selected_operation(self):
        operation = self.operation_var.get()
        self.update_progress(0, "Rozpoczynanie...")
        
        if operation == "Ekstrakcja":
            xlsx_path = self.extract_xlsx_entry.get()
            column_str = self.extract_column_entry.get()
            if not xlsx_path or not column_str: 
                messagebox.showerror("Błąd danych", "Wszystkie pola dla ekstrakcji muszą być wypełnione.")
                self.update_progress(0, "Błąd danych")
                return
            try: 
                column_num = int(column_str)
                if column_num <= 0: raise ValueError("Numer kolumny musi być dodatni.")
            except ValueError: 
                messagebox.showerror("Błąd danych", "Nieprawidłowy numer kolumny. Musi to być liczba dodatnia.")
                self.update_progress(0, "Błąd danych")
                return
            perform_extraction(xlsx_path, column_num, progress_callback=self.update_progress)
            
        elif operation == "Łączenie":
            xlsx_path = self.combine_xlsx_entry.get()
            csv_path = self.combine_csv_entry.get()
            column_str = self.combine_column_entry.get()
            if not xlsx_path or not csv_path or not column_str: 
                messagebox.showerror("Błąd danych", "Wszystkie pola dla łączenia muszą być wypełnione.")
                self.update_progress(0, "Błąd danych")
                return
            try: 
                column_num = int(column_str)
                if column_num <= 0: raise ValueError("Numer kolumny musi być dodatni.")
            except ValueError: 
                messagebox.showerror("Błąd danych", "Nieprawidłowy numer kolumny. Musi to być liczba dodatnia.")
                self.update_progress(0, "Błąd danych")
                return
            perform_combination(xlsx_path, csv_path, column_num, progress_callback=self.update_progress)

if __name__ == "__main__":
    app = AppLauncher()
    app.mainloop()
