import argparse
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd


EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}


def load_table(input_path: Path, sheet_name: str | None, input_sep: str) -> pd.DataFrame:
    ext = input_path.suffix.lower()

    if ext in EXCEL_EXTENSIONS:
        return pd.read_excel(input_path, sheet_name=sheet_name or 0, dtype=object)

    return pd.read_csv(input_path, sep=input_sep, dtype=object)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_columns = []
    for i, col in enumerate(df.columns, start=1):
        if pd.isna(col):
            new_columns.append(f"kolumna_{i}")
            continue

        col_name = str(col).strip()
        new_columns.append(col_name if col_name else f"kolumna_{i}")

    df = df.copy()
    df.columns = new_columns
    return df


def convert_to_pairwise_layout(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    out_rows: list[list[object]] = []

    for _, row in df.iterrows():
        out_row: list[object] = []
        for col_name, value in zip(df.columns, row.tolist()):
            if pd.isna(value):
                continue

            if isinstance(value, str) and value.strip() == "":
                continue

            if len(str(value)) > 50:
                continue

            out_row.extend([col_name, value])
        out_rows.append(out_row)

    return pd.DataFrame(out_rows)


def save_table(df: pd.DataFrame, output_path: Path, output_sep: str) -> None:
    ext = output_path.suffix.lower()

    if ext in EXCEL_EXTENSIONS:
        df.to_excel(output_path, index=False, header=False)
        return

    df.to_csv(output_path, sep=output_sep, index=False, header=False, encoding="utf-8-sig")


def build_default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_pairwise{input_path.suffix}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Konwertuje tabelę do układu: nazwa1;wartość1;nazwa2;wartość2... "
            "dla każdego wiersza danych."
        )
    )
    parser.add_argument("-i", "--input", help="Ścieżka do pliku wejściowego (xlsx/xls/csv).")
    parser.add_argument("-o", "--output", help="Ścieżka do pliku wyjściowego. Domyślnie: *_pairwise.")
    parser.add_argument("--sheet", help="Nazwa arkusza (tylko dla Excel). Domyślnie pierwszy arkusz.")
    parser.add_argument("--input-sep", default=";", help="Separator wejściowy dla CSV (domyślnie ;).")
    parser.add_argument("--output-sep", default=";", help="Separator wyjściowy dla CSV (domyślnie ;).")
    parser.add_argument("--gui", action="store_true", help="Uruchom interfejs graficzny (CTk).")
    return parser.parse_args()


def run_conversion(input_path: Path, output_path: Path, sheet: str | None, input_sep: str, output_sep: str) -> tuple[int, int]:
    if not input_path.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku wejściowego: {input_path}")

    source_df = load_table(input_path=input_path, sheet_name=sheet, input_sep=input_sep)
    converted_df = convert_to_pairwise_layout(source_df)
    save_table(converted_df, output_path=output_path, output_sep=output_sep)
    return len(source_df), len(converted_df)


class ConverterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")

        self.title("Konwerter tabeli -> para nazwa/wartość")
        self.geometry("860x500")
        self.minsize(820, 480)
        self.configure(fg_color="#fff8fc")

        self.hotpink = "#ff69b4"
        self.hotpink_hover = "#e754a5"
        self.white = "#ffffff"

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.input_sep_var = tk.StringVar(value=";")
        self.output_sep_var = tk.StringVar(value=";")
        self.status_var = tk.StringVar(value="Wybierz plik wejściowy i kliknij: Konwertuj")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ctk.CTkFrame(self, fg_color=self.white, corner_radius=14, border_width=2, border_color=self.hotpink)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        title = ctk.CTkLabel(
            container,
            text="Konwerter skomplikowanej transpozycji",
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color="#2a2a2a",
        )
        title.pack(anchor="w", padx=20, pady=(18, 6))

        subtitle = ctk.CTkLabel(
            container,
            text="Układ: nazwa1;wartość1;nazwa2;wartość2... dla każdego wiersza",
            font=ctk.CTkFont(size=14),
            text_color="#4b4b4b",
        )
        subtitle.pack(anchor="w", padx=20, pady=(0, 16))

        self._path_row(container, "Plik wejściowy", self.input_var, self.pick_input_file)
        self._path_row(container, "Plik wyjściowy", self.output_var, self.pick_output_file)

        options_frame = ctk.CTkFrame(container, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=(4, 12))
        options_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self._labeled_entry(options_frame, 0, "Arkusz (Excel, opcjonalnie)", self.sheet_var, "np. Sheet1")
        self._labeled_entry(options_frame, 1, "Separator wejściowy CSV", self.input_sep_var, ";")
        self._labeled_entry(options_frame, 2, "Separator wyjściowy CSV", self.output_sep_var, ";")

        buttons_frame = ctk.CTkFrame(container, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=20, pady=(8, 8))
        buttons_frame.grid_columnconfigure((0, 1, 2), weight=1)

        convert_button = ctk.CTkButton(
            buttons_frame,
            text="Konwertuj",
            command=self.convert,
            height=42,
            fg_color=self.hotpink,
            hover_color=self.hotpink_hover,
            text_color="#ffffff",
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        convert_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        auto_button = ctk.CTkButton(
            buttons_frame,
            text="Ustaw domyślny plik wyjściowy",
            command=self.set_default_output,
            height=42,
            fg_color="#ffc0de",
            hover_color="#ffadd4",
            text_color="#2f2f2f",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        auto_button.grid(row=0, column=1, padx=8, sticky="ew")

        clear_button = ctk.CTkButton(
            buttons_frame,
            text="Wyczyść",
            command=self.clear_fields,
            height=42,
            fg_color="#ffe6f3",
            hover_color="#ffd4ea",
            text_color="#2f2f2f",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        clear_button.grid(row=0, column=2, padx=(8, 0), sticky="ew")

        status_label = ctk.CTkLabel(
            container,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            wraplength=770,
            text_color="#2f2f2f",
            font=ctk.CTkFont(size=13),
            fg_color="#fff0f8",
            corner_radius=10,
            padx=12,
            pady=10,
        )
        status_label.pack(fill="x", padx=20, pady=(8, 18))

    def _path_row(self, parent: ctk.CTkFrame, label_text: str, variable: tk.StringVar, browse_command) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=6)
        row.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(row, text=label_text, text_color="#333333", font=ctk.CTkFont(size=14, weight="bold"))
        label.grid(row=0, column=0, sticky="w", pady=(0, 6))

        entry = ctk.CTkEntry(
            row,
            textvariable=variable,
            height=38,
            border_color=self.hotpink,
            fg_color="#fffafb",
            text_color="#222222",
        )
        entry.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        button = ctk.CTkButton(
            row,
            text="Wybierz",
            command=browse_command,
            width=120,
            height=38,
            fg_color=self.hotpink,
            hover_color=self.hotpink_hover,
            text_color="#ffffff",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        button.grid(row=1, column=1, sticky="e")

    def _labeled_entry(
        self,
        parent: ctk.CTkFrame,
        column: int,
        label_text: str,
        variable: tk.StringVar,
        placeholder: str,
    ) -> None:
        label = ctk.CTkLabel(parent, text=label_text, text_color="#333333", font=ctk.CTkFont(size=13, weight="bold"))
        label.grid(row=0, column=column, sticky="w", padx=(0, 10), pady=(0, 6))

        entry = ctk.CTkEntry(
            parent,
            textvariable=variable,
            height=36,
            placeholder_text=placeholder,
            border_color=self.hotpink,
            fg_color="#fffafb",
            text_color="#222222",
        )
        entry.grid(row=1, column=column, sticky="ew", padx=(0, 10))

    def pick_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Wybierz plik wejściowy",
            filetypes=[
                ("Obsługiwane", "*.xlsx *.xls *.xlsm *.xlsb *.csv"),
                ("Excel", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("CSV", "*.csv"),
            ],
        )
        if not file_path:
            return

        self.input_var.set(file_path)
        if not self.output_var.get().strip():
            self.set_default_output()

    def pick_output_file(self) -> None:
        suggested_name = "wynik_pairwise.xlsx"
        current_input = self.input_var.get().strip()
        if current_input:
            suggested_name = build_default_output_path(Path(current_input)).name

        file_path = filedialog.asksaveasfilename(
            title="Wybierz plik wyjściowy",
            initialfile=suggested_name,
            defaultextension=".xlsx",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
                ("Wszystkie", "*.*"),
            ],
        )
        if file_path:
            self.output_var.set(file_path)

    def set_default_output(self) -> None:
        input_text = self.input_var.get().strip()
        if not input_text:
            self.status_var.set("Najpierw wybierz plik wejściowy, aby ustawić domyślną ścieżkę wyjściową.")
            return

        input_path = Path(input_text)
        self.output_var.set(str(build_default_output_path(input_path)))
        self.status_var.set("Ustawiono domyślny plik wyjściowy.")

    def clear_fields(self) -> None:
        self.input_var.set("")
        self.output_var.set("")
        self.sheet_var.set("")
        self.input_sep_var.set(";")
        self.output_sep_var.set(";")
        self.status_var.set("Pola wyczyszczone. Wybierz plik wejściowy i kliknij: Konwertuj")

    def convert(self) -> None:
        input_text = self.input_var.get().strip()
        output_text = self.output_var.get().strip()
        sheet_text = self.sheet_var.get().strip() or None
        input_sep = self.input_sep_var.get() or ";"
        output_sep = self.output_sep_var.get() or ";"

        if not input_text:
            messagebox.showerror("Brak danych", "Wybierz plik wejściowy.")
            return

        input_path = Path(input_text)
        if not output_text:
            output_text = str(build_default_output_path(input_path))
            self.output_var.set(output_text)

        output_path = Path(output_text)

        try:
            in_rows, out_rows = run_conversion(
                input_path=input_path,
                output_path=output_path,
                sheet=sheet_text,
                input_sep=input_sep,
                output_sep=output_sep,
            )
            self.status_var.set(f"Gotowe. Zapisano: {output_path} | Wiersze: {in_rows} -> {out_rows}")
            messagebox.showinfo("Sukces", f"Konwersja zakończona.\n\nZapisano:\n{output_path}")
        except Exception as exc:
            self.status_var.set(f"Błąd: {exc}")
            messagebox.showerror("Błąd konwersji", str(exc))


def run_gui() -> None:
    app = ConverterApp()
    app.mainloop()


def main() -> None:
    args = parse_args()
    if args.gui or not args.input:
        run_gui()
        return

    input_path = Path(args.input)

    output_path = Path(args.output) if args.output else build_default_output_path(input_path)
    in_rows, out_rows = run_conversion(
        input_path=input_path,
        output_path=output_path,
        sheet=args.sheet,
        input_sep=args.input_sep,
        output_sep=args.output_sep,
    )

    print(f"Zapisano: {output_path}")
    print(f"Wierszy wejściowych: {in_rows} | Wierszy wyjściowych: {out_rows}")


if __name__ == "__main__":
    main()
