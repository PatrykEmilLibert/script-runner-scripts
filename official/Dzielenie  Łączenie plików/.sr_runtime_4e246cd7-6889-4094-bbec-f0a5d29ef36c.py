import pandas as pd
import customtkinter
from tkinter import filedialog, messagebox
import os
import shutil
import re # Do rozpoznawania plików
from collections import Counter

# Spróbuj zaimportować win32com, jeśli nie ma, poinformuj użytkownika
try:
    import win32com.client
    import pythoncom
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

import openpyxl

# --- Konfiguracja GUI ---
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")
customtkinter.set_widget_scaling(0.8)

#================================================================================
# FUNKCJE POMOCNICZE
#================================================================================

def get_sheet_names(filepath):
    """Zwraca listę nazw arkuszy z danego pliku Excel."""
    if not filepath or not os.path.exists(filepath):
        return None
    try:
        workbook = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        return workbook.sheetnames
    except Exception as e:
        messagebox.showerror("Błąd odczytu pliku", f"Nie można odczytać arkuszy z pliku: {filepath}\nBłąd: {e}")
        return None

def find_last_row(ws):
    """Znajduje ostatni używany wiersz w arkuszu."""
    # xlUp = -4162
    return ws.Cells(ws.Rows.Count, 1).End(-4162).Row

#================================================================================
# LOGIKA DZIELENIA
#================================================================================

def podziel_excel(plik_wejsciowy, wierszy_na_plik, folder_wyjsciowy, liczba_wierszy_naglowka=0, sheet_identifier=0, app_instance=None):
    """Dzieli plik Excel, tworząc kopie oryginału i usuwając zbędne wiersze."""
    if not PYWIN32_AVAILABLE:
        messagebox.showerror("Błąd krytyczny", "Biblioteka pywin32 nie jest zainstalowana.")
        return

    excel = None
    pythoncom.CoInitialize()

    try:
        df = pd.read_excel(plik_wejsciowy, sheet_name=sheet_identifier, header=None, engine='openpyxl')
        liczba_wierszy_df = len(df)
        if liczba_wierszy_df == 0:
            messagebox.showerror("Błąd", f"Wybrany arkusz '{sheet_identifier}' jest pusty.")
            return

        if liczba_wierszy_naglowka >= liczba_wierszy_df:
            messagebox.showerror("Błąd", "Liczba wierszy nagłówka jest większa lub równa liczbie wszystkich wierszy.")
            return

        liczba_wierszy_danych = liczba_wierszy_df - liczba_wierszy_naglowka
        liczba_plikow_do_utworzenia = (liczba_wierszy_danych + wierszy_na_plik - 1) // wierszy_na_plik

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        nazwa_pliku_base, rozszerzenie_pliku = os.path.splitext(os.path.basename(plik_wejsciowy))
        arkusz_sufix = f"_arkusz_{sheet_identifier}".replace(" ", "_")

        pliki_utworzone = 0
        for i in range(liczba_plikow_do_utworzenia):
            if app_instance:
                 app_instance.update_status_split(f"Dzielenie: część {i+1}/{liczba_plikow_do_utworzenia}...")

            nazwa_wyjsciowa = f"{nazwa_pliku_base}{arkusz_sufix}_czesc_{i + 1}{rozszerzenie_pliku}"
            sciezka_wyjsciowa = os.path.abspath(os.path.join(folder_wyjsciowy, nazwa_wyjsciowa))
            shutil.copy2(plik_wejsciowy, sciezka_wyjsciowa)

            workbook = None
            try:
                workbook = excel.Workbooks.Open(sciezka_wyjsciowa)
                ws = workbook.Sheets(sheet_identifier)
                ws.Activate()

                start_danych_idx = i * wierszy_na_plik
                end_danych_idx = start_danych_idx + wierszy_na_plik
                pierwszy_wiersz_danych_do_zachowania = liczba_wierszy_naglowka + start_danych_idx + 1
                ostatni_wiersz_danych_do_zachowania = liczba_wierszy_naglowka + end_danych_idx

                if ostatni_wiersz_danych_do_zachowania < liczba_wierszy_df:
                    ws.Range(f"{ostatni_wiersz_danych_do_zachowania + 1}:{liczba_wierszy_df}").Delete()
                if pierwszy_wiersz_danych_do_zachowania > liczba_wierszy_naglowka + 1:
                    ws.Range(f"{liczba_wierszy_naglowka + 1}:{pierwszy_wiersz_danych_do_zachowania - 1}").Delete()

                workbook.Save()
                pliki_utworzone += 1
            except Exception as e_edit:
                messagebox.showerror("Błąd edycji pliku", f"Błąd podczas edycji {nazwa_wyjsciowa}.\n{e_edit}")
            finally:
                if workbook:
                    workbook.Close(SaveChanges=False)

        messagebox.showinfo("Sukces", f"Plik podzielono na {pliki_utworzone} części.\nZapisano w: {folder_wyjsciowy}")

    except Exception as e_main:
        messagebox.showerror("Błąd krytyczny", f"Wystąpił błąd: {e_main}")
    finally:
        if excel:
            excel.ScreenUpdating = True
            excel.Quit()
        pythoncom.CoUninitialize()
        if app_instance:
            app_instance.update_status_split("Gotowy.")

#================================================================================
# LOGIKA SKŁADANIA STANDARDOWEGO
#================================================================================

def find_parts_for_base_file(base_file_path):
    """Znajduje wszystkie pasujące części dla danego pliku bazowego."""
    folder_path = os.path.dirname(base_file_path)
    base_filename = os.path.basename(base_file_path)
    
    pattern = re.compile(r"(.+?)_arkusz_(.+?)_czesc_(\d+)(\..+)")
    match = pattern.match(base_filename)
    if not match:
        return None, None, None
    
    base_name, _, _, extension = match.groups()
    
    parts = []
    part_pattern_str = fr"^{re.escape(base_name)}_arkusz_.+?_czesc_(\d+){re.escape(extension)}$"
    part_pattern = re.compile(part_pattern_str)

    for filename in os.listdir(folder_path):
        part_match = part_pattern.match(filename)
        if part_match:
            part_num = int(part_match.group(1))
            parts.append((part_num, os.path.join(folder_path, filename)))
    
    parts.sort()
    return parts, base_name, extension

def scal_pliki(base_file_path, sheet_identifier, header_rows_count, app_instance=None):
    """Scala grupę plików w jeden, bazując na pliku wzorcowym i wybranym arkuszu."""
    if not PYWIN32_AVAILABLE: return

    excel = None
    pythoncom.CoInitialize()
    
    app_instance.update_status_merge(f"Wyszukiwanie części plików...")
    file_group, base_name, extension = find_parts_for_base_file(base_file_path)

    if not file_group:
        messagebox.showerror("Błąd", "Nie można znaleźć pasujących części. Sprawdź, czy nazwa pliku jest prawidłowa (np. ..._czesc_1.xlsx).")
        app_instance.update_status_merge("Błąd wyszukiwania.")
        return
        
    all_data_frames = []
    app_instance.update_status_merge("Wczytywanie danych...")
    try:
        # Wczytaj wszystkie części do pamięci
        for i, (part_num, file_path) in enumerate(file_group):
            app_instance.update_status_merge(f"Wczytywanie części {part_num}/{len(file_group)}...")
            df = pd.read_excel(file_path, sheet_name=sheet_identifier, header=None, skiprows=header_rows_count if i > 0 else 0)
            all_data_frames.append(df)
        
        final_df = pd.concat(all_data_frames, ignore_index=True)

    except Exception as e_pandas:
        messagebox.showerror("Błąd wczytywania danych", f"Nie udało się wczytać danych z jednej z części. Upewnij się, że wszystkie części zawierają arkusz '{sheet_identifier}'.\n\nBłąd: {e_pandas}")
        return

    folder_path = os.path.dirname(base_file_path)
    output_filename = f"{base_name}_SCALONY{extension}"
    output_path = os.path.join(folder_path, output_filename)
    
    if os.path.exists(output_path):
        if not messagebox.askyesno("Potwierdzenie", f"Plik '{output_filename}' już istnieje.\n\nCzy chcesz go nadpisać?"):
            app_instance.update_status_merge("Anulowano.")
            return

    template_file_path = file_group[0][1]
    try:
        shutil.copy2(template_file_path, output_path)
    except Exception as e_copy:
        messagebox.showerror("Błąd kopiowania", f"Nie można utworzyć pliku wyjściowego.\n{e_copy}")
        return

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        main_workbook = excel.Workbooks.Open(output_path)
        main_ws = main_workbook.Sheets(sheet_identifier)
        main_ws.Cells.ClearContents()
        
        app_instance.update_status_merge("Zapisywanie scalonych danych...")
        data_to_write = final_df.where(pd.notna(final_df), None).values.tolist()

        start_row, start_col = 1, 1
        end_row = start_row + len(data_to_write) - 1
        end_col = start_col + len(data_to_write[0]) - 1 if data_to_write else start_col
        range_to_write = main_ws.Range(main_ws.Cells(start_row, start_col), main_ws.Cells(end_row, end_col))
        range_to_write.Value = data_to_write
        
        main_workbook.Save()
        messagebox.showinfo("Sukces", f"Pliki zostały pomyślnie scalone.\n\nZapisano w:\n{output_path}")

    except Exception as e_main:
        messagebox.showerror("Błąd krytyczny", f"Wystąpił błąd podczas zapisywania scalonego pliku: {e_main}")
    finally:
        if 'main_workbook' in locals() and main_workbook:
            main_workbook.Close(SaveChanges=False)
        if excel:
            excel.ScreenUpdating = True
            excel.Quit()
        pythoncom.CoUninitialize()
        if app_instance:
            app_instance.update_status_merge("Gotowy.")

#================================================================================
# LOGIKA SKŁADANIA NIESTANDARDOWEGO
#================================================================================

def scal_pliki_niestandardowo(file_paths, sheet_index, app_instance=None):
    """Scala dane z wielu dowolnych plików, tworząc nowy plik z nowym nagłówkiem."""
    all_data_frames = []
    app_instance.update_status_custom_merge("Wczytywanie danych...")
    try:
        for i, file_path in enumerate(file_paths):
            app_instance.update_status_custom_merge(f"Wczytywanie pliku {i+1}/{len(file_paths)}...")
            df = pd.read_excel(file_path, sheet_name=sheet_index, header=None, skiprows=1)
            all_data_frames.append(df)

        if not all_data_frames:
            messagebox.showwarning("Brak danych", "Nie udało się wczytać danych z żadnego z wybranych plików.")
            return

        final_df = pd.concat(all_data_frames, ignore_index=True)
        
        # Tworzenie nowego nagłówka
        num_cols = len(final_df.columns)
        new_header = ['cat', 'id']
        if num_cols > 2:
            new_header.extend([f'Kolumna_{i+1}' for i in range(2, num_cols)])
        final_df.columns = new_header

    except Exception as e:
        messagebox.showerror("Błąd wczytywania", f"Nie można przetworzyć plików. Upewnij się, że każdy plik zawiera arkusz o podanym numerze.\n\nBłąd: {e}")
        return

    output_path = filedialog.asksaveasfilename(
        title="Zapisz scalony plik jako...",
        defaultextension=".xlsx",
        filetypes=[("Plik Excel", "*.xlsx")])
        
    if not output_path:
        app_instance.update_status_custom_merge("Anulowano zapis.")
        return

    try:
        app_instance.update_status_custom_merge("Zapisywanie pliku...")
        final_df.to_excel(output_path, index=False)
        messagebox.showinfo("Sukces", f"Pliki zostały pomyślnie scalone i zapisane w:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Błąd zapisu", f"Nie udało się zapisać pliku.\n\nBłąd: {e}")
    finally:
        app_instance.update_status_custom_merge("Gotowy.")


#================================================================================
# GŁÓWNA APLIKACJA GUI
#================================================================================

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Dzielenie i Składanie Plików Excel v4.0")
        self.geometry("540x580")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.custom_merge_files = []
        
        self.tab_view = customtkinter.CTkTabview(self)
        self.tab_view.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.tab_view.add("Dzielenie")
        self.tab_view.add("Składanie")
        self.tab_view.add("Składanie Niestandardowe")
        
        self.create_split_tab(self.tab_view.tab("Dzielenie"))
        self.create_merge_tab(self.tab_view.tab("Składanie"))
        self.create_custom_merge_tab(self.tab_view.tab("Składanie Niestandardowe"))

    # --- ZAKŁADKI ---
    def create_split_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        
        etykieta_plik = customtkinter.CTkLabel(tab, text="1. Wybierz plik Excel do podziału:")
        etykieta_plik.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.entry_plik_split = customtkinter.CTkEntry(tab)
        self.entry_plik_split.grid(row=1, column=0, padx=(10, 5), pady=5, sticky="ew")
        przycisk_plik = customtkinter.CTkButton(tab, text="Przeglądaj...", command=self.wybierz_plik_split)
        przycisk_plik.grid(row=1, column=1, padx=(5, 10), pady=5, sticky="ew")

        etykieta_arkusz = customtkinter.CTkLabel(tab, text="2. Wybierz arkusz:")
        etykieta_arkusz.grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.optionmenu_arkusz_split = customtkinter.CTkOptionMenu(tab, values=["Wybierz plik..."], state="disabled")
        self.optionmenu_arkusz_split.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        etykieta_parametry = customtkinter.CTkLabel(tab, text="3. Ustaw parametry:")
        etykieta_parametry.grid(row=4, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.entry_wiersze_split = customtkinter.CTkEntry(tab, placeholder_text="Liczba wierszy danych na plik")
        self.entry_wiersze_split.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self.entry_naglowek_split = customtkinter.CTkEntry(tab, placeholder_text="Liczba wierszy nagłówka (np. 0)")
        self.entry_naglowek_split.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        etykieta_folder = customtkinter.CTkLabel(tab, text="4. Wybierz folder zapisu:")
        etykieta_folder.grid(row=7, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.entry_folder_split = customtkinter.CTkEntry(tab)
        self.entry_folder_split.grid(row=8, column=0, padx=(10, 5), pady=5, sticky="ew")
        przycisk_folder = customtkinter.CTkButton(tab, text="Przeglądaj...", command=self.wybierz_folder_split)
        przycisk_folder.grid(row=8, column=1, padx=(5, 10), pady=5, sticky="ew")
        
        self.przycisk_podziel = customtkinter.CTkButton(tab, text="Podziel Plik", command=self.uruchom_podzial, height=40)
        self.przycisk_podziel.grid(row=9, column=0, columnspan=2, padx=10, pady=(20, 5), sticky="ew")
        self.status_label_split = customtkinter.CTkLabel(tab, text="Gotowy.")
        self.status_label_split.grid(row=10, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

    def create_merge_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)

        etykieta_plik_merge = customtkinter.CTkLabel(tab, text="1. Wybierz dowolną część pliku do scalenia:")
        etykieta_plik_merge.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.entry_plik_merge = customtkinter.CTkEntry(tab)
        self.entry_plik_merge.grid(row=1, column=0, padx=(10, 5), pady=5, sticky="ew")
        przycisk_plik_merge = customtkinter.CTkButton(tab, text="Przeglądaj...", command=self.wybierz_plik_merge)
        przycisk_plik_merge.grid(row=1, column=1, padx=(5, 10), pady=5, sticky="ew")
        
        etykieta_arkusz_merge = customtkinter.CTkLabel(tab, text="2. Wybierz arkusz do scalenia:")
        etykieta_arkusz_merge.grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.optionmenu_arkusz_merge = customtkinter.CTkOptionMenu(tab, values=["Wybierz plik..."], state="disabled")
        self.optionmenu_arkusz_merge.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        etykieta_naglowek_merge = customtkinter.CTkLabel(tab, text="3. Podaj liczbę wierszy nagłówka w plikach:")
        etykieta_naglowek_merge.grid(row=4, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.entry_naglowek_merge = customtkinter.CTkEntry(tab, placeholder_text="np. 1")
        self.entry_naglowek_merge.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        self.przycisk_scal = customtkinter.CTkButton(tab, text="Scal Pliki", command=self.uruchom_scalanie, height=40)
        self.przycisk_scal.grid(row=6, column=0, columnspan=2, padx=10, pady=(20, 5), sticky="ew")
        self.status_label_merge = customtkinter.CTkLabel(tab, text="Wybierz plik-część, aby rozpocząć.")
        self.status_label_merge.grid(row=7, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

    def create_custom_merge_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)

        etykieta_pliki_custom = customtkinter.CTkLabel(tab, text="1. Wybierz pliki do scalenia:")
        etykieta_pliki_custom.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.label_pliki_custom_count = customtkinter.CTkLabel(tab, text="Nie wybrano plików.")
        self.label_pliki_custom_count.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        przycisk_pliki_custom = customtkinter.CTkButton(tab, text="Wybierz wiele plików...", command=self.wybierz_wiele_plikow)
        przycisk_pliki_custom.grid(row=1, column=1, padx=(5, 10), pady=5, sticky="e")

        etykieta_arkusz_custom = customtkinter.CTkLabel(tab, text="2. Podaj numer porządkowy arkusza do scalenia:")
        etykieta_arkusz_custom.grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        self.entry_arkusz_custom = customtkinter.CTkEntry(tab, placeholder_text="np. 1 dla pierwszego arkusza")
        self.entry_arkusz_custom.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        info_label = customtkinter.CTkLabel(tab, text="Info: Skrypt pominie pierwszy wiersz (nagłówek) z każdego pliku i utworzy nowy nagłówek 'cat', 'id' itd.", wraplength=400)
        info_label.grid(row=4, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="w")

        self.przycisk_scal_custom = customtkinter.CTkButton(tab, text="Scal i utwórz nowy plik", command=self.uruchom_scalanie_niestandardowe, height=40)
        self.przycisk_scal_custom.grid(row=5, column=0, columnspan=2, padx=10, pady=(20, 5), sticky="ew")
        self.status_label_custom_merge = customtkinter.CTkLabel(tab, text="Wybierz pliki i podaj numer arkusza.")
        self.status_label_custom_merge.grid(row=6, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")


    # --- METODY OBSŁUGI ZDARZEŃ ---
    def update_status_split(self, message): self.status_label_split.configure(text=message); self.update_idletasks()
    def update_status_merge(self, message): self.status_label_merge.configure(text=message); self.update_idletasks()
    def update_status_custom_merge(self, message): self.status_label_custom_merge.configure(text=message); self.update_idletasks()

    def wybierz_plik_split(self):
        filename = filedialog.askopenfilename(title="Wybierz plik Excel", filetypes=(("Pliki Excel", "*.xlsx *.xls *.xlsm"), ("Wszystkie pliki", "*.*")))
        if filename: self.entry_plik_split.delete(0, "end"); self.entry_plik_split.insert(0, filename); self.wczytaj_arkusze_split()

    def wybierz_folder_split(self):
        foldername = filedialog.askdirectory(title="Wybierz folder do zapisu")
        if foldername: self.entry_folder_split.delete(0, "end"); self.entry_folder_split.insert(0, foldername)
            
    def wybierz_plik_merge(self):
        filename = filedialog.askopenfilename(title="Wybierz dowolną część pliku", filetypes=(("Pliki Excel", "*.xlsx *.xls *.xlsm"), ("Wszystkie pliki", "*.*")))
        if filename: self.entry_plik_merge.delete(0, "end"); self.entry_plik_merge.insert(0, filename); self.wczytaj_arkusze_merge()

    def wybierz_wiele_plikow(self):
        filenames = filedialog.askopenfilenames(title="Wybierz pliki do scalenia", filetypes=(("Pliki Excel", "*.xlsx *.xls *.xlsm"), ("Wszystkie pliki", "*.*")))
        if filenames:
            self.custom_merge_files = filenames
            self.label_pliki_custom_count.configure(text=f"Wybrano plików: {len(self.custom_merge_files)}")
            self.update_status_custom_merge("Gotowy do scalenia.")

    def wczytaj_arkusze_split(self):
        filepath = self.entry_plik_split.get()
        if not filepath: return
        self.update_status_split("Wczytywanie arkuszy...")
        sheet_names = get_sheet_names(filepath)
        if sheet_names:
            self.optionmenu_arkusz_split.configure(values=sheet_names, state="normal"); self.optionmenu_arkusz_split.set(sheet_names[0])
            self.update_status_split("Wybierz arkusz i ustaw parametry.")
        else:
            self.optionmenu_arkusz_split.configure(values=["Błąd odczytu"], state="disabled"); self.update_status_split("Błąd: Nie można wczytać arkuszy.")

    def wczytaj_arkusze_merge(self):
        filepath = self.entry_plik_merge.get()
        if not filepath: return
        self.update_status_merge("Wczytywanie arkuszy...")
        sheet_names = get_sheet_names(filepath)
        if sheet_names:
            self.optionmenu_arkusz_merge.configure(values=sheet_names, state="normal"); self.optionmenu_arkusz_merge.set(sheet_names[0])
            self.update_status_merge("Wybierz arkusz, podaj liczbę nagłówków i scal.")
        else:
            self.optionmenu_arkusz_merge.configure(values=["Błąd odczytu"], state="disabled"); self.update_status_merge("Błąd: Nie można wczytać arkuszy.")

    # --- URUCHAMIANIE OPERACJI ---
    def uruchom_podzial(self):
        if not all([self.entry_plik_split.get(), self.entry_folder_split.get(), self.entry_wiersze_split.get(), self.entry_naglowek_split.get()]):
            messagebox.showerror("Błąd", "Wypełnij wszystkie pola w zakładce Dzielenie.")
            return
        try:
            wierszy = int(self.entry_wiersze_split.get()); naglowek = int(self.entry_naglowek_split.get())
            if wierszy <= 0 or naglowek < 0: raise ValueError
        except ValueError: messagebox.showerror("Błąd", "Liczba wierszy i nagłówka musi być prawidłową liczbą dodatnią."); return
            
        self.przycisk_podziel.configure(state="disabled")
        podziel_excel(self.entry_plik_split.get(), wierszy, self.entry_folder_split.get(), naglowek, self.optionmenu_arkusz_split.get(), self)
        self.przycisk_podziel.configure(state="normal")
        
    def uruchom_scalanie(self):
        base_file_path = self.entry_plik_merge.get(); sheet_identifier = self.optionmenu_arkusz_merge.get(); header_rows_str = self.entry_naglowek_merge.get()
        if not base_file_path or not header_rows_str or "Wybierz plik" in sheet_identifier or "Błąd" in sheet_identifier:
            messagebox.showerror("Błąd", "Wypełnij wszystkie pola w zakładce Składanie.")
            return
        try:
            header_rows_count = int(header_rows_str)
            if header_rows_count < 0: raise ValueError
        except ValueError: messagebox.showerror("Błąd", "Liczba wierszy nagłówka musi być prawidłową liczbą (0 lub więcej)."); return
            
        self.przycisk_scal.configure(state="disabled")
        scal_pliki(base_file_path, sheet_identifier, header_rows_count, self)
        self.przycisk_scal.configure(state="normal")

    def uruchom_scalanie_niestandardowe(self):
        if not self.custom_merge_files: messagebox.showerror("Błąd", "Najpierw wybierz pliki do scalenia."); return
        sheet_num_str = self.entry_arkusz_custom.get()
        if not sheet_num_str: messagebox.showerror("Błąd", "Podaj numer arkusza."); return
        
        try:
            sheet_index = int(sheet_num_str) - 1
            if sheet_index < 0: raise ValueError
        except ValueError: messagebox.showerror("Błąd", "Numer arkusza musi być liczbą całkowitą większą od 0."); return

        self.przycisk_scal_custom.configure(state="disabled")
        scal_pliki_niestandardowo(self.custom_merge_files, sheet_index, self)
        self.przycisk_scal_custom.configure(state="normal")


if __name__ == '__main__':
    if os.name != 'nt' or not PYWIN32_AVAILABLE:
        messagebox.showerror("Błąd krytyczny", "Ta aplikacja wymaga systemu Windows, MS Excel oraz biblioteki pywin32.")
    else:
        app = App()
        app.mainloop()
