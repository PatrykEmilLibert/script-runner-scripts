import csv
import math
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox


ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


HOT_PINK = "#ff69b4"


class CSVSplitterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Dzielenie pliku CSV")
        self.geometry("820x430")
        self.minsize(760, 390)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=14)
        self.main_frame.grid(row=0, column=0, padx=18, pady=18, sticky="nsew")

        self.main_frame.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(
            self.main_frame,
            text="Dzielenie CSV na mniejsze pliki",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="#111111",
        )
        title.grid(row=0, column=0, columnspan=3, padx=18, pady=(16, 24), sticky="w")

        self._add_path_row(
            row=1,
            label="Plik CSV:",
            button_text="Wybierz",
            browse_command=self.select_input_file,
            entry_attr="input_entry",
        )

        self._add_path_row(
            row=2,
            label="Folder wyjściowy:",
            button_text="Wybierz",
            browse_command=self.select_output_folder,
            entry_attr="output_entry",
        )

        rows_label = ctk.CTkLabel(self.main_frame, text="Liczba wierszy danych na plik:")
        rows_label.grid(row=3, column=0, padx=18, pady=10, sticky="w")
        self.rows_entry = ctk.CTkEntry(self.main_frame, width=180)
        self.rows_entry.insert(0, "1000")
        self.rows_entry.grid(row=3, column=1, padx=18, pady=10, sticky="w")

        sep_label = ctk.CTkLabel(self.main_frame, text="Separator:")
        sep_label.grid(row=4, column=0, padx=18, pady=10, sticky="w")
        self.sep_entry = ctk.CTkEntry(self.main_frame, width=180)
        self.sep_entry.insert(0, ";")
        self.sep_entry.grid(row=4, column=1, padx=18, pady=10, sticky="w")

        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            text_color="#333333",
            anchor="w",
            justify="left",
            wraplength=760,
        )
        self.status_label.grid(row=6, column=0, columnspan=3, padx=18, pady=(8, 12), sticky="ew")

        self.split_button = ctk.CTkButton(
            self.main_frame,
            text="Podziel CSV",
            command=self.split_csv,
            height=42,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=HOT_PINK,
            hover_color="#ff4fa6",
            text_color="#ffffff",
        )
        self.split_button.grid(row=5, column=0, columnspan=3, padx=18, pady=(18, 10), sticky="ew")

    def _add_path_row(self, row: int, label: str, button_text: str, browse_command, entry_attr: str) -> None:
        row_label = ctk.CTkLabel(self.main_frame, text=label)
        row_label.grid(row=row, column=0, padx=18, pady=10, sticky="w")

        entry = ctk.CTkEntry(self.main_frame, placeholder_text="Wskaż ścieżkę...")
        entry.grid(row=row, column=1, padx=18, pady=10, sticky="ew")
        setattr(self, entry_attr, entry)

        browse_btn = ctk.CTkButton(
            self.main_frame,
            text=button_text,
            width=120,
            command=browse_command,
            fg_color=HOT_PINK,
            hover_color="#ff4fa6",
            text_color="#ffffff",
        )
        browse_btn.grid(row=row, column=2, padx=(0, 18), pady=10, sticky="e")

    def set_status(self, text: str, color: str = "#333333") -> None:
        self.status_label.configure(text=text, text_color=color)
        self.update_idletasks()

    def select_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Wybierz plik CSV",
            filetypes=(("CSV", "*.csv"), ("Wszystkie pliki", "*.*")),
        )
        if path:
            self.input_entry.delete(0, ctk.END)
            self.input_entry.insert(0, path)
            if not self.output_entry.get().strip():
                self.output_entry.insert(0, str(Path(path).parent))
            self.set_status("")

    def select_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Wybierz folder wyjściowy")
        if path:
            self.output_entry.delete(0, ctk.END)
            self.output_entry.insert(0, path)
            self.set_status("")

    def split_csv(self) -> None:
        input_path = Path(self.input_entry.get().strip())
        output_dir = Path(self.output_entry.get().strip())
        separator = self.sep_entry.get() or ";"

        if not input_path.exists() or not input_path.is_file():
            messagebox.showerror("Błąd", "Wybierz poprawny plik CSV.")
            return
        if not output_dir:
            messagebox.showerror("Błąd", "Wybierz folder wyjściowy.")
            return

        try:
            rows_per_file = int(self.rows_entry.get().strip())
            if rows_per_file <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Błąd", "Liczba wierszy musi być dodatnią liczbą całkowitą.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            created_files, total_rows = split_csv_with_header(
                input_file=input_path,
                output_dir=output_dir,
                rows_per_file=rows_per_file,
                delimiter=separator,
            )
        except ValueError as error:
            messagebox.showerror("Błąd danych", str(error))
            return
        except Exception as error:
            messagebox.showerror("Błąd", f"Nie udało się podzielić pliku:\n{error}")
            return

        summary = (
            f"Gotowe. Utworzono {created_files} plików w folderze:\n{output_dir}\n"
            f"Liczba wierszy danych: {total_rows}"
        )
        self.set_status(summary, color="#0f7a2f")
        messagebox.showinfo("Sukces", summary)


def split_csv_with_header(input_file: Path, output_dir: Path, rows_per_file: int, delimiter: str = ";") -> tuple[int, int]:
    with input_file.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.reader(source, delimiter=delimiter)
        header = next(reader, None)

        if header is None:
            raise ValueError("Plik CSV jest pusty.")

        rows = list(reader)

    total_rows = len(rows)
    if total_rows == 0:
        raise ValueError("Plik zawiera tylko nagłówki i brak wierszy danych.")

    parts_count = math.ceil(total_rows / rows_per_file)
    base_name = input_file.stem

    for part_index in range(parts_count):
        start = part_index * rows_per_file
        end = start + rows_per_file
        chunk = rows[start:end]
        output_name = f"{base_name}_part_{part_index + 1:03d}.csv"
        output_path = output_dir / output_name

        with output_path.open("w", encoding="utf-8", newline="") as target:
            writer = csv.writer(target, delimiter=delimiter)
            writer.writerow(header)
            writer.writerows(chunk)

    return parts_count, total_rows


if __name__ == "__main__":
    app = CSVSplitterApp()
    app.mainloop()