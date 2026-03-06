import tkinter as tk
from tkinter import filedialog, messagebox
import os
import csv

def export_foldernames_to_csv():
    """
    Opens a dialog box to select a folder, then exports
    the names of subfolders from that folder to a CSV file.
    """
    # 1. Open dialog and ask the user to select a folder
    folder_path = filedialog.askdirectory(title="Wybierz folder, z którego chcesz wyciągnąć nazwy podfolderów")

    # 2. Check if the user selected a folder (if not, exit the function)
    if not folder_path:
        messagebox.showinfo("Informacja", "Nie wybrano żadnego folderu.")
        return

    try:
        # 3. Get a list of all items in the folder and filter only directories
        # os.path.isdir() ensures that we only add subfolder names to the list
        all_items = os.listdir(folder_path)
        foldernames = [item for item in all_items if os.path.isdir(os.path.join(folder_path, item))]

        # Check if any folders were found
        if not foldernames:
            messagebox.showinfo("Informacja", "W wybranym folderze nie znaleziono żadnych podfolderów.")
            return

        # 4. Define the path to the target CSV file
        output_csv_path = os.path.join(folder_path, 'lista_podfolderow.csv')

        # 5. Save the folder names to the CSV file
        with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write the column header
            writer.writerow(['Nazwa Podfolderu'])
            
            # Write each folder name in a new row
            for foldername in foldernames:
                writer.writerow([foldername])

        # 6. Inform the user about the success
        messagebox.showinfo(
            "Sukces!",
            f"Pomyślnie zapisano {len(foldernames)} nazw podfolderów.\n\n"
            f"Plik został zapisany w lokalizacji:\n{output_csv_path}"
        )

    except Exception as e:
        # In case of an error, inform the user
        messagebox.showerror("Błąd!", f"Wystąpił nieoczekiwany błąd:\n{e}")

# --- Application Main Window Configuration (GUI) ---

# Create the main window
root = tk.Tk()
root.title("Ekstraktor Nazw Podfolderów do CSV")
root.geometry("400x200") # Set window size

# Set padding (internal margins) for aesthetics
main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(expand=True, fill=tk.BOTH)

# Instruction label for the user
instruction_label = tk.Label(
    main_frame,
    text="Kliknij przycisk poniżej, aby wybrać folder i wygenerować plik CSV z listą jego podfolderów.",
    wraplength=360, # Automatic text wrapping
    justify=tk.CENTER
)
instruction_label.pack(pady=(0, 20)) # Additional bottom margin

# Button to run the function
run_button = tk.Button(
    main_frame,
    text="Wybierz Folder i Zapisz CSV",
    command=export_foldernames_to_csv, # Changed to the new function
    font=("Helvetica", 10, "bold"),
    bg="#4CAF50", # Background color
    fg="white",   # Text color
    padx=10,
    pady=5
)
run_button.pack()

# Start the window's event loop
root.mainloop()