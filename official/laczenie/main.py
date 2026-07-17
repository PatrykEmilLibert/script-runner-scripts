import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
from collections import Counter

def znajdz_wspolna_nazwe(folder_z_plikami):
    """
    Próbuje znaleźć wspólną nazwę pliku na podstawie plików w folderze.
    """
    nazwy_bazowe = []
    if not folder_z_plikami or not os.path.isdir(folder_z_plikami):
        return None
    for filename in os.listdir(folder_z_plikami):
        if (filename.endswith(".xlsx") or filename.endswith(".xls")) and re.match(r'.*_(\d+)\.(xlsx|xls)$', filename):
            nazwa_bez_numeru = re.sub(r'_(\d+)\.(xlsx|xls)$', '', filename)
            nazwy_bazowe.append(nazwa_bez_numeru)
    if nazwy_bazowe:
        return Counter(nazwy_bazowe).most_common(1)[0][0]
    return None

def usun_nadmiarowe_spacje(df):
    """
    Usuwa białe znaki z początku i końca każdej komórki w DataFrame 
    (tylko dla stringów w kolumnach typu object) i konwertuje 'nan' na ''.
    """
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip().replace('nan', '')
    return df

def polacz_pliki(folder_z_plikami, oryginalna_nazwa_pliku, liczba_wierszy_bloku_naglowka):
    """
    Łączy pliki Excela. Cały blok nagłówka z pierwszego pliku jest umieszczany na górze.
    Ostatni wiersz tego bloku definiuje nazwy kolumn dla danych.
    """
    try:
        pliki_do_polaczenia = []
        for filename in os.listdir(folder_z_plikami):
            if filename.startswith(f"{oryginalna_nazwa_pliku}_") and \
               (filename.endswith(".xlsx") or filename.endswith(".xls")):
                match = re.search(r'_(\d+)\.(xlsx|xls)$', filename)
                if match:
                    numer = int(match.group(1))
                    pliki_do_polaczenia.append((numer, os.path.join(folder_z_plikami, filename)))

        if not pliki_do_polaczenia:
            messagebox.showinfo("Informacja", f"Nie znaleziono plików pasujących do wzorca '{oryginalna_nazwa_pliku}_numer.xlsx/xls' w folderze: {folder_z_plikami}")
            return

        pliki_do_polaczenia.sort(key=lambda x: x[0])

        df_header_block = pd.DataFrame()
        list_of_data_dfs = []
        actual_data_column_names = None

        # --- Przetwarzanie pierwszego pliku (dla bloku nagłówka i danych) ---
        path_pierwszy_plik = pliki_do_polaczenia[0][1]
        nazwa_pierwszego_pliku_base = os.path.basename(path_pierwszy_plik)

        if liczba_wierszy_bloku_naglowka > 0:
            try:
                # Wczytaj cały blok nagłówka z pierwszego pliku
                df_header_block = pd.read_excel(path_pierwszy_plik, header=None, nrows=liczba_wierszy_bloku_naglowka)
                df_header_block = usun_nadmiarowe_spacje(df_header_block.copy()) # Usuń spacje z bloku nagłówka

                if len(df_header_block) < liczba_wierszy_bloku_naglowka:
                    messagebox.showwarning("Ostrzeżenie (Nagłówek)", 
                                           f"Pierwszy plik ('{nazwa_pierwszego_pliku_base}') nie zawiera {liczba_wierszy_bloku_naglowka} wierszy dla bloku nagłówka. "
                                           "Blok nagłówka nie zostanie użyty, a nazwy kolumn nie zostaną ustalone.")
                    df_header_block = pd.DataFrame() # Resetuj blok nagłówka
                    actual_data_column_names = None
                    # Spróbuj wczytać dane z pierwszego pliku bez pomijania wierszy, jakby nie było nagłówka
                    df_data_first_file = pd.read_excel(path_pierwszy_plik, header=None)
                else:
                    # Użyj ostatniego wiersza bloku nagłówka jako nazwy kolumn dla danych
                    actual_data_column_names = df_header_block.iloc[liczba_wierszy_bloku_naglowka - 1].astype(str).tolist()
                    # Wczytaj dane z pierwszego pliku, pomijając blok nagłówka
                    df_data_first_file = pd.read_excel(path_pierwszy_plik, header=None, skiprows=liczba_wierszy_bloku_naglowka)

            except Exception as e:
                messagebox.showerror("Błąd (Pierwszy Plik)", f"Błąd podczas przetwarzania nagłówka/danych z pierwszego pliku ('{nazwa_pierwszego_pliku_base}'): {e}")
                return
        else: # liczba_wierszy_bloku_naglowka == 0
            df_data_first_file = pd.read_excel(path_pierwszy_plik, header=None)
            actual_data_column_names = None # Brak zdefiniowanych nazw kolumn

        # Przetwarzanie danych z pierwszego pliku
        if not df_data_first_file.empty:
            if actual_data_column_names:
                if len(actual_data_column_names) == df_data_first_file.shape[1]:
                    df_data_first_file.columns = actual_data_column_names
                else:
                    messagebox.showwarning("Niezgodność Kolumn (Pierwszy Plik)",
                                           f"Liczba kolumn danych ({df_data_first_file.shape[1]}) w pierwszym pliku ('{nazwa_pierwszego_pliku_base}') "
                                           f"nie zgadza się z liczbą nazw kolumn z bloku nagłówka ({len(actual_data_column_names)}). "
                                           "Dane z tego pliku zostaną dołączone z domyślnymi nazwami kolumn.")
                    actual_data_column_names = None # Resetuj, aby nie używać dla kolejnych plików
                    df_data_first_file.columns = [f"Kol_{j}" for j in range(df_data_first_file.shape[1])] # Generyczne nazwy
            df_data_first_file = usun_nadmiarowe_spacje(df_data_first_file)
            list_of_data_dfs.append(df_data_first_file)

        # --- Przetwarzanie kolejnych plików (tylko dane) ---
        for i in range(1, len(pliki_do_polaczenia)):
            _num, path_kolejny_plik = pliki_do_polaczenia[i]
            nazwa_kolejnego_pliku_base = os.path.basename(path_kolejny_plik)
            try:
                df_data_subsequent_file = pd.read_excel(path_kolejny_plik, header=None, skiprows=liczba_wierszy_bloku_naglowka)
                if not df_data_subsequent_file.empty:
                    if actual_data_column_names: # Jeśli nazwy kolumn zostały ustalone z pierwszego pliku
                        if len(actual_data_column_names) == df_data_subsequent_file.shape[1]:
                            df_data_subsequent_file.columns = actual_data_column_names
                        else:
                            messagebox.showwarning(f"Niezgodność Kolumn ({nazwa_kolejnego_pliku_base})",
                                                   f"Liczba kolumn danych ({df_data_subsequent_file.shape[1]}) w pliku '{nazwa_kolejnego_pliku_base}' "
                                                   f"nie zgadza się z oczekiwaną liczbą ({len(actual_data_column_names)}). "
                                                   "Dane z tego pliku zostaną pominięte, aby uniknąć błędów łączenia.")
                            continue # Pomiń ten DataFrame
                    # Jeśli actual_data_column_names to None, dane zostaną dołączone z domyślnymi nagłówkami numerycznymi
                    df_data_subsequent_file = usun_nadmiarowe_spacje(df_data_subsequent_file)
                    list_of_data_dfs.append(df_data_subsequent_file)
            except Exception as e:
                messagebox.showwarning(f"Błąd Odczytu ({nazwa_kolejnego_pliku_base})", 
                                       f"Nie udało się odczytać danych z pliku '{nazwa_kolejnego_pliku_base}': {e}. Plik zostanie pominięty.")

        # --- Łączenie wszystkich części danych ---
        merged_data_df = pd.DataFrame()
        if list_of_data_dfs:
            try:
                merged_data_df = pd.concat(list_of_data_dfs, ignore_index=True)
            except ValueError as ve: # Może się zdarzyć, jeśli kolumny nadal nie pasują (np. różne typy danych w tych samych kolumnach)
                messagebox.showerror("Błąd Łączenia Danych", 
                                     f"Wystąpił błąd podczas łączenia danych: {ve}. "
                                     "Może to być spowodowane niezgodnością typów danych lub strukturą kolumn pomimo prób ich ujednolicenia. "
                                     "Sprawdź pliki źródłowe.")
                # Zapisz tylko blok nagłówka, jeśli istnieje
                if not df_header_block.empty:
                    merged_data_df = pd.DataFrame() # Upewnij się, że dane nie są zapisywane
                else:
                    return
        
        if df_header_block.empty and merged_data_df.empty:
            messagebox.showinfo("Informacja", "Nie znaleziono ani bloku nagłówka, ani danych do zapisania. Plik wyjściowy nie zostanie utworzony.")
            return

        # --- Zapis do pliku Excel ---
        nazwa_wyjsciowa = f"POLACZONY_{oryginalna_nazwa_pliku}{os.path.splitext(path_pierwszy_plik)[1]}"
        sciezka_wyjsciowa = os.path.join(folder_z_plikami, nazwa_wyjsciowa)

        try:
            with pd.ExcelWriter(sciezka_wyjsciowa, engine='openpyxl') as writer:
                if not df_header_block.empty:
                    df_header_block.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                
                if not merged_data_df.empty:
                    start_row_for_data = len(df_header_block) if not df_header_block.empty else 0
                    # Zapisujemy dane BEZ ich własnych nagłówków, ponieważ ostatni wiersz df_header_block pełni tę rolę
                    merged_data_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=start_row_for_data)
            
            messagebox.showinfo("Sukces", f"Pliki połączone pomyślnie do:\n{sciezka_wyjsciowa}")

        except Exception as e:
            messagebox.showerror("Błąd Zapisu", f"Nie udało się zapisać połączonego pliku: {e}")

    except Exception as e:
        messagebox.showerror("Błąd Krytyczny", f"Wystąpił nieoczekiwany błąd: {e}")


# --- Funkcje GUI ---
def wybierz_folder():
    folder_wybrany = filedialog.askdirectory(initialdir=".", title="Wybierz folder z podzielonymi plikami")
    if folder_wybrany: 
        root.folder = folder_wybrany
        entry_folder.delete(0, tk.END)
        entry_folder.insert(tk.END, root.folder)
        automatyczna_nazwa = znajdz_wspolna_nazwe(root.folder)
        if automatyczna_nazwa:
            entry_nazwa.delete(0, tk.END)
            entry_nazwa.insert(tk.END, automatyczna_nazwa)
        elif root.folder: 
            messagebox.showinfo("Informacja", "Nie udało się automatycznie wykryć wspólnej nazwy pliku. Proszę wprowadzić ją ręcznie.")

def uruchom_polaczenie():
    folder_z_plikami = entry_folder.get()
    oryginalna_nazwa = entry_nazwa.get()
    liczba_wierszy_bloku_naglowka_val = 0 

    try:
        wiersze_str = entry_liczba_wierszy_bloku_naglowka.get()
        if wiersze_str: 
            liczba_wierszy_bloku_naglowka_val = int(wiersze_str)
            if liczba_wierszy_bloku_naglowka_val < 0:
                messagebox.showerror("Błąd", "Liczba wierszy w bloku nagłówka musi być wartością nieujemną (>= 0).")
                return
    except ValueError:
        messagebox.showerror("Błąd", "Liczba wierszy w bloku nagłówka jest nieprawidłowa. Proszę podać liczbę całkowitą.")
        return

    if not folder_z_plikami or not os.path.isdir(folder_z_plikami):
        messagebox.showerror("Błąd", "Nie wybrano prawidłowego folderu z plikami.")
    elif not oryginalna_nazwa:
        messagebox.showerror("Błąd", "Nie podano oryginalnej (bazowej) nazwy pliku.")
    else:
        polacz_pliki(folder_z_plikami, oryginalna_nazwa, liczba_wierszy_bloku_naglowka_val)

# --- Tworzenie okna głównego ---
root = tk.Tk()
root.title("Łączenie Podzielonych Plików Excel")
root.folder = "" 

ramka_folderu = tk.Frame(root, padx=10, pady=5)
ramka_folderu.pack(fill=tk.X)
etykieta_folder = tk.Label(ramka_folderu, text="Folder z plikami:")
etykieta_folder.pack(side=tk.LEFT, padx=(0,5))
entry_folder = tk.Entry(ramka_folderu, width=50)
entry_folder.pack(side=tk.LEFT, expand=True, fill=tk.X)
przycisk_folder = tk.Button(ramka_folderu, text="Przeglądaj", command=wybierz_folder)
przycisk_folder.pack(side=tk.LEFT, padx=(5,0))

ramka_nazwy = tk.Frame(root, padx=10, pady=5)
ramka_nazwy.pack(fill=tk.X)
etykieta_nazwa = tk.Label(ramka_nazwy, text="Bazowa nazwa pliku (bez _numeru i rozszerzenia):")
etykieta_nazwa.pack(side=tk.LEFT, anchor='w')
entry_nazwa = tk.Entry(ramka_nazwy, width=50)
entry_nazwa.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5,0))

ramka_opcji = tk.Frame(root, padx=10, pady=5)
ramka_opcji.pack(fill=tk.X)
etykieta_liczba_wierszy_bloku_naglowka = tk.Label(ramka_opcji,
    text="Ile wierszy od góry PIERWSZEGO pliku tworzy blok nagłówka (np. 3)?\n"
         "Ten blok zostanie w całości umieszczony na górze połączonego pliku.\n"
         "OSTATNI wiersz tego bloku (np. 3-ci) posłuży jako nazwy kolumn dla DANYCH.\n"
         "W kolejnych plikach tyle samo wierszy od góry zostanie pominiętych.\n"
         "Wpisz 0, jeśli pliki nie mają bloku nagłówka (wszystkie wiersze to dane).",
    justify=tk.LEFT)
etykieta_liczba_wierszy_bloku_naglowka.pack(anchor='w')
entry_liczba_wierszy_bloku_naglowka = tk.Entry(ramka_opcji, width=10)
entry_liczba_wierszy_bloku_naglowka.insert(0, "1") 
entry_liczba_wierszy_bloku_naglowka.pack(anchor='w', pady=(0,10))

przycisk_polacz = tk.Button(root, text="Połącz Pliki", command=uruchom_polaczenie, width=20, height=2)
przycisk_polacz.pack(pady=20)

root.minsize(560, 350) 
root.mainloop()
