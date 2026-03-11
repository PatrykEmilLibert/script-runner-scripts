import tkinter as tk
from tkinter import filedialog, messagebox
import os
import csv

def export_filenames_to_csv():
    """
    Otwiera okno dialogowe do wyboru folderu, a następnie eksportuje
    nazwy plików z tego folderu do pliku CSV.
    """
    # 1. Otwórz okno dialogowe i poproś użytkownika o wybranie folderu
    folder_path = filedialog.askdirectory(title="Wybierz folder, z którego chcesz wyciągnąć nazwy plików")

    # 2. Sprawdź, czy użytkownik wybrał folder (jeśli nie, zakończ funkcję)
    if not folder_path:
        messagebox.showinfo("Informacja", "Nie wybrano żadnego folderu.")
        return

    try:
        # 3. Pobierz listę wszystkich pozycji w folderze i odfiltruj tylko pliki
        # os.path.isfile() zapewnia, że nie dodamy do listy nazw podfolderów
        all_items = os.listdir(folder_path)
        filenames = [item for item in all_items if os.path.isfile(os.path.join(folder_path, item))]

        # Sprawdź, czy znaleziono jakiekolwiek pliki
        if not filenames:
            messagebox.showinfo("Informacja", "W wybranym folderze nie znaleziono żadnych plików.")
            return

        # 4. Zdefiniuj ścieżkę do docelowego pliku CSV
        output_csv_path = os.path.join(folder_path, 'lista_plikow.csv')

        # 5. Zapisz nazwy plików do pliku CSV
        with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Zapisz nagłówek kolumny
            writer.writerow(['Nazwa Pliku'])
            
            # Zapisz każdą nazwę pliku w nowym wierszu
            for filename in filenames:
                writer.writerow([filename])

        # 6. Poinformuj użytkownika o sukcesie
        messagebox.showinfo(
            "Sukces!",
            f"Pomyślnie zapisano {len(filenames)} nazw plików.\n\n"
            f"Plik został zapisany w lokalizacji:\n{output_csv_path}"
        )

    except Exception as e:
        # W razie błędu, poinformuj użytkownika
        messagebox.showerror("Błąd!", f"Wystąpił nieoczekiwany błąd:\n{e}")

# --- Konfiguracja Głównego Okna Aplikacji (GUI) ---

# Stworzenie głównego okna
root = tk.Tk()
root.title("Ekstraktor Nazw Plików do CSV")
root.geometry("400x200") # Ustawienie rozmiaru okna

# Ustawienie paddingu (marginesów wewnętrznych) dla estetyki
main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(expand=True, fill=tk.BOTH)

# Etykieta z instrukcją dla użytkownika
instruction_label = tk.Label(
    main_frame,
    text="Kliknij przycisk poniżej, aby wybrać folder i wygenerować plik CSV z listą jego plików.",
    wraplength=360, # Automatyczne zawijanie tekstu
    justify=tk.CENTER
)
instruction_label.pack(pady=(0, 20)) # Dodatkowy margines na dole

# Przycisk uruchamiający funkcję
run_button = tk.Button(
    main_frame,
    text="Wybierz Folder i Zapisz CSV",
    command=export_filenames_to_csv,
    font=("Helvetica", 10, "bold"),
    bg="#4CAF50", # Kolor tła
    fg="white",   # Kolor tekstu
    padx=10,
    pady=5
)
run_button.pack()

# Uruchomienie pętli zdarzeń okna
root.mainloop()