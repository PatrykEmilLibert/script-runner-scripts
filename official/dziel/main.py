import pandas as pd
import customtkinter
from tkinter import filedialog, messagebox
import os
# shutil nie będzie już potrzebny do kopiowania w głównej logice
# import shutil
import time # Dodano do obsługi opóźnień przy COM

# Spróbuj zaimportować win32com, jeśli nie ma, poinformuj użytkownika
try:
    import win32com.client
    import pythoncom # Potrzebne do CoInitialize/CoUninitialize
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

# openpyxl i dataframe_to_rows są nadal potrzebne do odczytu początkowego pliku przez pandas
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


# Ustawienie wyglądu i motywu dla customtkinter
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

def excel_value_converter(val):
    """Konwertuje wartości pandas na typy odpowiednie dla Excel COM."""
    if pd.isna(val):
        return None  # Puste komórki
    # Można dodać więcej konwersji typów w razie potrzeby, np. daty
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime() # Konwersja Timestamp z pandas na datetime Pythona
    return val

def podziel_excel(plik_wejsciowy, wierszy_na_plik, folder_wyjsciowy, liczba_wierszy_nagłówka=0):
    """
    Dzieli plik Excela na mniejsze pliki. Używa pywin32 do zapisu, tworząc nowe skoroszyty.
    Makra z oryginalnych plików .xlsm NIE są przenoszone tą metodą.
    """
    if not PYWIN32_AVAILABLE:
        messagebox.showerror("Błąd krytyczny",
                             "Biblioteka pywin32 nie jest zainstalowana lub wystąpił błąd podczas jej importu. "
                             "Ta biblioteka jest wymagana do poprawnego działania skryptu w systemie Windows. "
                             "Zainstaluj ją używając: pip install pywin32")
        return

    excel = None
    # Inicjalizacja COM dla bieżącego wątku. Ważne, jeśli GUI działa w innym wątku niż logika.
    # Dla prostych aplikacji tkinter/customtkinter może nie być ściśle konieczne, ale to dobra praktyka.
    if PYWIN32_AVAILABLE:
        pythoncom.CoInitialize()

    try:
        original_file_ext = os.path.splitext(plik_wejsciowy)[1].lower()

        if not original_file_ext in ('.xlsx', '.xls', '.xlsm'):
            messagebox.showerror("Błąd", "Nieobsługiwany format pliku. Wybierz plik .xlsx, .xls lub .xlsm.")
            return

        # Sprawdzenie i ewentualne utworzenie folderu wyjściowego
        if not os.path.exists(folder_wyjsciowy):
            try:
                os.makedirs(folder_wyjsciowy)
                print(f"Utworzono folder wyjściowy: {folder_wyjsciowy}")
            except Exception as e_mkdir:
                messagebox.showerror("Błąd folderu", f"Nie można utworzyć folderu wyjściowego: {folder_wyjsciowy}\nBłąd: {e_mkdir}")
                return
        
        df = pd.read_excel(plik_wejsciowy, header=None, engine='openpyxl')
        liczba_wierszy_df = len(df)

        if liczba_wierszy_df == 0:
            messagebox.showerror("Błąd", "Plik wejściowy jest pusty.")
            return

        if liczba_wierszy_nagłówka > liczba_wierszy_df:
            messagebox.showerror("Błąd", "Liczba wierszy nagłówka nie może być większa niż całkowita liczba wierszy w pliku.")
            return

        liczba_wierszy_danych = liczba_wierszy_df - liczba_wierszy_nagłówka
        
        if liczba_wierszy_danych <= 0 and liczba_wierszy_nagłówka > 0:
            messagebox.showinfo("Informacja", "Plik zawiera tylko wiersze nagłówka lub mniej wierszy niż zadeklarowany nagłówek. Tworzenie pojedynczego pliku z nagłówkiem (jeśli istnieje).")
            liczba_plikow_do_utworzenia = 1
        elif liczba_wierszy_danych <= 0 and liczba_wierszy_nagłówka == 0:
            return 
        else:
            liczba_plikow_do_utworzenia = (liczba_wierszy_danych + wierszy_na_plik - 1) // wierszy_na_plik

        nazwa_pliku_base, _ = os.path.splitext(os.path.basename(plik_wejsciowy))

        naglowki_df = pd.DataFrame()
        if liczba_wierszy_nagłówka > 0:
            naglowki_df = df.iloc[:liczba_wierszy_nagłówka]

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception as e_excel_dispatch:
            messagebox.showerror("Błąd Excel", f"Nie można uruchomić aplikacji Excel. Upewnij się, że jest zainstalowana.\nBłąd: {e_excel_dispatch}")
            return
            
        xls_to_xlsx_warning_shown = False
        xlsm_macros_not_copied_warning_shown = False


        for i in range(liczba_plikow_do_utworzenia):
            start_index_danych_w_df = i * wierszy_na_plik
            end_index_danych_w_df = min((i + 1) * wierszy_na_plik, liczba_wierszy_danych)
            
            df_part_dane = df.iloc[liczba_wierszy_nagłówka + start_index_danych_w_df : liczba_wierszy_nagłówka + end_index_danych_w_df]

            if not naglowki_df.empty:
                df_do_zapisu_w_excelu = pd.concat([naglowki_df, df_part_dane], ignore_index=True)
            else:
                df_do_zapisu_w_excelu = df_part_dane

            if liczba_wierszy_danych <= 0 and i == 0 and not naglowki_df.empty:
                df_do_zapisu_w_excelu = naglowki_df
            elif df_do_zapisu_w_excelu.empty and liczba_plikow_do_utworzenia == 1 and not naglowki_df.empty :
                df_do_zapisu_w_excelu = naglowki_df
            elif df_do_zapisu_w_excelu.empty and liczba_plikow_do_utworzenia > 0 :
                print(f"Pominięto tworzenie pustego pliku dla części: {i + 1}")
                continue
            
            docelowe_rozszerzenie = original_file_ext
            # Przy tworzeniu nowych plików, jeśli oryginał był .xlsm, nowy będzie .xlsx, chyba że jawnie zapiszemy jako .xlsm
            # Skupmy się na poprawnym zapisie danych. Makra nie będą kopiowane.
            if original_file_ext == '.xls':
                docelowe_rozszerzenie = '.xlsx'
                if not xls_to_xlsx_warning_shown:
                    messagebox.showwarning("Konwersja .xls",
                                           "Oryginalny plik był w formacie .xls. Podzielone pliki zostaną zapisane w nowszym formacie .xlsx.")
                    xls_to_xlsx_warning_shown = True
            elif original_file_ext == '.xlsm':
                 # Nowy plik nie będzie miał makr, więc można zapisać jako .xlsx lub .xlsm (ale pusty z makr)
                 # Dla bezpieczeństwa i mniejszego zamieszania, zapiszmy jako .xlsx
                 docelowe_rozszerzenie = '.xlsx'
                 if not xlsm_macros_not_copied_warning_shown:
                    messagebox.showwarning("Ostrzeżenie .xlsm",
                                           "Oryginalny plik był w formacie .xlsm. Podzielone pliki są tworzone jako nowe skoroszyty .xlsx i NIE będą zawierać oryginalnych makr.")
                    xlsm_macros_not_copied_warning_shown = True

            
            nazwa_wyjsciowa = f"{nazwa_pliku_base}_{i + 1}{docelowe_rozszerzenie}"
            # Używamy os.path.abspath, aby uzyskać pełną, znormalizowaną ścieżkę
            sciezka_wyjsciowa = os.path.abspath(os.path.join(folder_wyjsciowy, nazwa_wyjsciowa))


            workbook = None
            sheet = None
            try:
                # Tworzenie nowego skoroszytu
                workbook = excel.Workbooks.Add()
                sheet = workbook.Worksheets(1) # Domyślnie pierwszy arkusz
                # sheet.Name = "Podzielone Dane" # Można ustawić nazwę arkusza

                if not df_do_zapisu_w_excelu.empty:
                    # Konwersja NaN na None i Timestamp na datetime przed zapisem
                    data_to_write = df_do_zapisu_w_excelu.applymap(excel_value_converter).values.tolist()
                    
                    start_row, start_col = 1, 1
                    # Sprawdzenie czy data_to_write nie jest pusta i czy wewnętrzne listy nie są puste
                    if data_to_write and data_to_write[0]:
                        end_row = start_row + len(data_to_write) - 1
                        end_col = start_col + len(data_to_write[0]) - 1
                        
                        excel_range = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
                        excel_range.Value = data_to_write
                    elif data_to_write and not data_to_write[0]: # Pusta lista wewnętrzna (np. DataFrame z jedną kolumną i bez wierszy)
                        print(f"Dane dla pliku {nazwa_wyjsciowa} są puste (brak kolumn).")
                    else: # data_to_write jest pusta
                        print(f"Brak danych do zapisu dla pliku {nazwa_wyjsciowa}.")

                file_format_code = None
                if docelowe_rozszerzenie == '.xlsx':
                    file_format_code = 51  # xlOpenXMLWorkbook
                # Jeśli chcielibyśmy zapisać jako .xlsm (nawet bez makr), kod byłby 52
                # elif docelowe_rozszerzenie == '.xlsm':
                #     file_format_code = 52  # xlOpenXMLWorkbookMacroEnabled
                
                # Zapis skoroszytu pod nową, pełną ścieżką
                if file_format_code:
                    workbook.SaveAs(sciezka_wyjsciowa, FileFormat=file_format_code)
                else: # Dla .xls (które konwertujemy na .xlsx) lub innych nieprzewidzianych
                    workbook.SaveAs(sciezka_wyjsciowa) 

                workbook.Close(SaveChanges=False) # False, bo SaveAs już zapisało

            except Exception as e_save:
                # Wypisz pełną ścieżkę w błędzie
                messagebox.showerror("Błąd zapisu (Excel COM)", f"Nie można zapisać pliku: {sciezka_wyjsciowa}\nBłąd: {e_save}")
                if workbook:
                    workbook.Close(SaveChanges=False) # Spróbuj zamknąć, jeśli otwarty
                continue
            finally:
                if sheet:
                    del sheet
                if workbook:
                    del workbook
        
        messagebox.showinfo("Sukces",
                            f"Plik został podzielony na {liczba_plikow_do_utworzenia} części i zapisany w folderze:\n{folder_wyjsciowy}")

    except FileNotFoundError:
        messagebox.showerror("Błąd", "Nie znaleziono wskazanego pliku.")
    except pd.errors.EmptyDataError:
        messagebox.showerror("Błąd", "Plik Excel jest pusty lub nie zawiera danych w oczekiwanym formacie.")
    except ImportError: 
        messagebox.showerror("Błąd importu", "Nie znaleziono wymaganej biblioteki. Upewnij się, że pandas, openpyxl i customtkinter są zainstalowane.")
    except Exception as e_main:
        messagebox.showerror("Błąd krytyczny", f"Wystąpił nieoczekiwany błąd główny: {e_main}\nTyp błędu: {type(e_main)}")
    finally:
        if excel:
            excel.Quit()
            del excel
        if PYWIN32_AVAILABLE:
            pythoncom.CoUninitialize()


def wybierz_plik():
    filename = filedialog.askopenfilename(initialdir=".",
                                          title="Wybierz plik Excel",
                                          filetypes=(("Pliki Excel", "*.xlsx *.xls *.xlsm"),
                                                     ("Wszystkie pliki", "*.*")))
    if filename: 
        entry_plik.delete(0, customtkinter.END)
        entry_plik.insert(customtkinter.END, filename)

def wybierz_folder():
    foldername = filedialog.askdirectory(initialdir=".", title="Wybierz folder do zapisu")
    if foldername: 
        entry_folder.delete(0, customtkinter.END)
        entry_folder.insert(customtkinter.END, foldername)

def uruchom_podzial():
    if not PYWIN32_AVAILABLE:
        messagebox.showerror("Błąd krytyczny",
                             "Biblioteka pywin32 nie jest dostępna. Skrypt nie może kontynuować. "
                             "Zainstaluj ją używając: pip install pywin32 i uruchom skrypt ponownie.")
        return

    plik_wejsciowy = entry_plik.get()
    folder_wyjsciowy = entry_folder.get()
    liczba_wierszy_str = entry_wiersze.get()
    liczba_wierszy_naglowka_str = entry_naglowek.get()

    if not plik_wejsciowy:
        messagebox.showerror("Błąd", "Nie wybrano pliku.")
        return 
    if not folder_wyjsciowy: 
        messagebox.showerror("Błąd", "Nie wybrano folderu do zapisu.")
        return 
    if not liczba_wierszy_str: 
        messagebox.showerror("Błąd", "Nie podano liczby wierszy na plik.")
        return 

    try:
        wierszy_na_plik = int(liczba_wierszy_str)
        if wierszy_na_plik <= 0:
            messagebox.showerror("Błąd", "Liczba wierszy na plik musi być większa od zera.")
            return 

        liczba_wierszy_nagłówka = int(liczba_wierszy_naglowka_str) if liczba_wierszy_naglowka_str else 0
        if liczba_wierszy_nagłówka < 0:
            messagebox.showerror("Błąd", "Liczba wierszy nagłówka nie może być ujemna.")
            return 

        podziel_excel(plik_wejsciowy, wierszy_na_plik, folder_wyjsciowy, liczba_wierszy_nagłówka)

    except ValueError:
        messagebox.showerror("Błąd", "Podana liczba wierszy lub liczba wierszy nagłówka jest nieprawidłowa. Wprowadź liczbę całkowitą.")
    except Exception as e: 
        messagebox.showerror("Błąd krytyczny", f"Wystąpił nieprzewidziany błąd podczas uruchamiania: {e}")


# --- Konfiguracja GUI z CustomTkinter ---
root = customtkinter.CTk()
root.title("Dzielenie Plików Excel v1.6 (Excel COM - Nowe Skoroszyty)") 
root.geometry("550x500") 

padding_options = {'padx': 10, 'pady': (10, 5)}
entry_width = 250 

main_frame = customtkinter.CTkFrame(master=root) 
main_frame.pack(padx=20, pady=20, fill="both", expand=True)

# --- Elementy GUI z CustomTkinter ---
etykieta_plik = customtkinter.CTkLabel(master=main_frame, text="Wybierz plik Excel:")
etykieta_plik.grid(row=0, column=0, columnspan=2, sticky="w", padx=padding_options['padx'], pady=padding_options['pady'])

entry_plik = customtkinter.CTkEntry(master=main_frame, width=entry_width + 100)
entry_plik.grid(row=1, column=0, sticky="ew", padx=padding_options['padx'], pady=padding_options['pady'])

przycisk_plik = customtkinter.CTkButton(master=main_frame, text="Przeglądaj...", command=wybierz_plik)
przycisk_plik.grid(row=1, column=1, sticky="ew", padx=(5, padding_options['padx']), pady=padding_options['pady'])

etykieta_wiersze = customtkinter.CTkLabel(master=main_frame, text="Liczba wierszy danych na plik (bez nagłówka):")
etykieta_wiersze.grid(row=2, column=0, columnspan=2, sticky="w", padx=padding_options['padx'], pady=padding_options['pady'])

entry_wiersze = customtkinter.CTkEntry(master=main_frame, width=150)
entry_wiersze.grid(row=3, column=0, sticky="w", padx=padding_options['padx'], pady=padding_options['pady'])

etykieta_naglowek = customtkinter.CTkLabel(master=main_frame, text="Liczba wierszy nagłówka (opcjonalnie):")
etykieta_naglowek.grid(row=4, column=0, columnspan=2, sticky="w", padx=padding_options['padx'], pady=padding_options['pady'])

entry_naglowek = customtkinter.CTkEntry(master=main_frame, width=150)
entry_naglowek.insert(0, "0") 
entry_naglowek.grid(row=5, column=0, sticky="w", padx=padding_options['padx'], pady=padding_options['pady'])

etykieta_folder = customtkinter.CTkLabel(master=main_frame, text="Wybierz folder do zapisu:")
etykieta_folder.grid(row=6, column=0, columnspan=2, sticky="w", padx=padding_options['padx'], pady=padding_options['pady'])

entry_folder = customtkinter.CTkEntry(master=main_frame, width=entry_width + 100)
entry_folder.grid(row=7, column=0, sticky="ew", padx=padding_options['padx'], pady=padding_options['pady'])

przycisk_folder = customtkinter.CTkButton(master=main_frame, text="Przeglądaj...", command=wybierz_folder)
przycisk_folder.grid(row=7, column=1, sticky="ew", padx=(5, padding_options['padx']), pady=padding_options['pady'])

info_label_text = "Uwaga: Ta wersja skryptu używa automatyzacji MS Excel (wymaga Excela i pywin32 w Windows). Makra NIE są kopiowane."
if not PYWIN32_AVAILABLE:
    info_label_text = "BŁĄD: Biblioteka pywin32 nie jest dostępna! Funkcjonalność zapisu przez Excel jest wyłączona."

info_label = customtkinter.CTkLabel(master=main_frame, text=info_label_text, wraplength=480, justify="left",
                                    text_color="orange" if not PYWIN32_AVAILABLE else ("#FFCC00" if customtkinter.get_appearance_mode() == "Dark" else "#8B4513") ) # Lepszy kontrast dla ostrzeżenia
info_label.grid(row=8, column=0, columnspan=2, padx=padding_options['padx'], pady=(10,0), sticky="ew")


przycisk_podziel = customtkinter.CTkButton(master=main_frame, text="Podziel Plik", command=uruchom_podzial, height=40)
przycisk_podziel.grid(row=9, column=0, columnspan=2, padx=padding_options['padx'], pady=(10,padding_options['pady'][1]), sticky="ew")

main_frame.grid_columnconfigure(0, weight=3)
main_frame.grid_columnconfigure(1, weight=1)

if __name__ == '__main__':
    if not PYWIN32_AVAILABLE and os.name == 'nt': 
        root.withdraw()
        messagebox.showerror("Krytyczny błąd zależności",
                             "Nie udało się zaimportować biblioteki pywin32. Jest ona niezbędna do działania tej wersji skryptu.\n\n"
                             "Proszę zainstalować ją za pomocą polecenia:\n"
                             "pip install pywin32\n\n"
                             "Aplikacja zostanie teraz zamknięta.")
    elif os.name != 'nt' and PYWIN32_AVAILABLE: # Jeśli pywin32 jest, ale nie jesteśmy na Windows (teoretycznie niemożliwe, ale dla pewności)
        root.withdraw()
        messagebox.showerror("Błąd platformy",
                                 "Biblioteka pywin32 jest specyficzna dla systemu Windows. Ta aplikacja nie będzie działać poprawnie na bieżącym systemie.")
    elif os.name != 'nt' and not PYWIN32_AVAILABLE: # Jeśli nie Windows i nie ma pywin32
         root.withdraw()
         messagebox.showerror("Błąd platformy i zależności",
                                 "Ta wersja skryptu jest przeznaczona dla systemu Windows i wymaga Microsoft Excel oraz biblioteki pywin32.\n\n"
                                 "Na innych systemach operacyjnych ta funkcjonalność nie będzie działać poprawnie.")
    else: # Jesteśmy na Windows i pywin32 jest dostępne
        root.mainloop()
