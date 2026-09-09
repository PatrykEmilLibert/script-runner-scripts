import customtkinter as ctk
from tkinter import filedialog, messagebox
import csv
import io
import os
import math
import threading

# ── motyw ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

HOTPINK = "#FF69B4"
HOTPINK_HOVER = "#E0559A"
LIGHT_BG = "#F5F5F5"
FRAME_BG = "#FFFFFF"
BYTES_IN_MB = 1024 * 1024

# ── logika podziału ────────────────────────────────────────────────────────────

def _row_size_bytes(row: list[str]) -> int:
    buf = io.StringIO(newline="")
    writer = csv.writer(buf)
    writer.writerow(row)
    return len(buf.getvalue().encode("utf-8"))


def split_csv(filepath: str, out_dir: str, split_mode: str,
              log_callback, done_callback,
              chunk_size: int | None = None,
              max_size_mb: float | None = None):
    try:
        with open(filepath, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            headers = next(reader)
            rows = list(reader)

        total = len(rows)
        if total == 0:
            done_callback(0, 0, "Plik nie zawiera danych.")
            return

        chunks: list[list[list[str]]] = []
        if split_mode == "rows":
            if chunk_size is None or chunk_size < 1:
                done_callback(0, 0, "Nieprawidłowy rozmiar paczki po wierszach.")
                return

            num_chunks = math.ceil(total / chunk_size)
            for i in range(num_chunks):
                chunks.append(rows[i * chunk_size:(i + 1) * chunk_size])

        elif split_mode == "size":
            if max_size_mb is None or max_size_mb <= 0:
                done_callback(0, 0, "Podaj poprawny limit rozmiaru pliku w MB.")
                return

            max_size_bytes = int(max_size_mb * BYTES_IN_MB)
            if max_size_bytes < 1:
                done_callback(0, 0, "Podany limit MB jest zbyt mały.")
                return

            # BOM (UTF-8-SIG) jest dodawany na początku każdego pliku wynikowego.
            header_bytes = 3 + _row_size_bytes(headers)
            current_chunk: list[list[str]] = []
            current_size = header_bytes

            for row_no, row in enumerate(rows, start=1):
                row_bytes = _row_size_bytes(row)

                if header_bytes + row_bytes > max_size_bytes:
                    done_callback(
                        0,
                        0,
                        (
                            f"Wiersz {row_no} jest zbyt duży dla limitu {max_size_mb:g} MB "
                            "(nawet sam z nagłówkiem przekracza limit)."
                        )
                    )
                    return

                if current_chunk and (current_size + row_bytes > max_size_bytes):
                    chunks.append(current_chunk)
                    current_chunk = [row]
                    current_size = header_bytes + row_bytes
                else:
                    current_chunk.append(row)
                    current_size += row_bytes

            if current_chunk:
                chunks.append(current_chunk)

            num_chunks = len(chunks)
        else:
            done_callback(0, 0, "Nieznany tryb dzielenia.")
            return

        base = os.path.splitext(os.path.basename(filepath))[0]

        for i, chunk in enumerate(chunks, start=1):
            out_name = f"{base}_czesc_{i:03d}.csv"
            out_path = os.path.join(out_dir, out_name)
            with open(out_path, "w", newline="", encoding="utf-8-sig") as out:
                writer = csv.writer(out)
                writer.writerow(headers)
                writer.writerows(chunk)

            if split_mode == "size":
                file_mb = os.path.getsize(out_path) / BYTES_IN_MB
                log_callback(f"  ✓ {out_name}  ({len(chunk)} wierszy, {file_mb:.2f} MB)")
            else:
                log_callback(f"  ✓ {out_name}  ({len(chunk)} wierszy)")

        done_callback(total, num_chunks, None)

    except Exception as exc:
        done_callback(0, 0, str(exc))


# ── GUI ────────────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Dzielenie CSV")
        self.geometry("620x660")
        self.resizable(False, False)
        self.configure(fg_color=LIGHT_BG)

        self._selected_file: str = ""
        self._out_dir: str = ""
        self._split_mode = ctk.StringVar(value="rows")

        self._build_ui()

    # ── widżety ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 20, "pady": 6}

        # tytuł
        title = ctk.CTkLabel(
            self, text="Dzielenie pliku CSV",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=HOTPINK
        )
        title.pack(pady=(20, 4))

        subtitle = ctk.CTkLabel(
            self, text="Dzieli CSV po liczbie wierszy lub po docelowym rozmiarze pliku (MB)",
            font=ctk.CTkFont(size=12), text_color="#666666"
        )
        subtitle.pack(pady=(0, 16))

        # ── karta: plik wejściowy ──
        card1 = self._card(self, "Plik wejściowy")
        card1.pack(fill="x", **pad)

        self.lbl_file = ctk.CTkLabel(card1, text="Nie wybrano pliku",
                                     text_color="#888888",
                                     font=ctk.CTkFont(size=11),
                                     anchor="w", wraplength=440)
        self.lbl_file.pack(side="left", padx=12, pady=10, expand=True, fill="x")

        ctk.CTkButton(card1, text="Wybierz CSV",
                      fg_color=HOTPINK, hover_color=HOTPINK_HOVER,
                      width=120, command=self._pick_file).pack(side="right", padx=12, pady=10)

        # ── karta: folder wyjściowy ──
        card2 = self._card(self, "Folder wyjściowy")
        card2.pack(fill="x", **pad)

        self.lbl_out = ctk.CTkLabel(card2, text="(domyślnie: folder pliku wejściowego)",
                                    text_color="#888888",
                                    font=ctk.CTkFont(size=11),
                                    anchor="w", wraplength=440)
        self.lbl_out.pack(side="left", padx=12, pady=10, expand=True, fill="x")

        ctk.CTkButton(card2, text="Wybierz folder",
                      fg_color=HOTPINK, hover_color=HOTPINK_HOVER,
                      width=120, command=self._pick_out).pack(side="right", padx=12, pady=10)

        # ── karta: tryb dzielenia ──
        card3 = self._card(self, "Tryb dzielenia")
        card3.pack(fill="x", **pad)

        mode_wrap = ctk.CTkFrame(card3, fg_color="transparent")
        mode_wrap.pack(fill="x", padx=12, pady=10)

        ctk.CTkRadioButton(
            mode_wrap,
            text="Po liczbie wierszy",
            variable=self._split_mode,
            value="rows",
            command=self._on_mode_change,
            fg_color=HOTPINK,
            hover_color=HOTPINK_HOVER,
        ).pack(side="left", padx=(0, 20))

        ctk.CTkRadioButton(
            mode_wrap,
            text="Po rozmiarze pliku (MB)",
            variable=self._split_mode,
            value="size",
            command=self._on_mode_change,
            fg_color=HOTPINK,
            hover_color=HOTPINK_HOVER,
        ).pack(side="left")

        # ── karta: wiersze na plik ──
        card4 = self._card(self, "Wiersze na plik (bez nagłówka)")
        card4.pack(fill="x", **pad)

        self.entry_rows = ctk.CTkEntry(card4, width=110,
                                       justify="center",
                                       font=ctk.CTkFont(size=14, weight="bold"),
                                       border_color=HOTPINK)
        self.entry_rows.insert(0, "1000")
        self.entry_rows.pack(padx=12, pady=10)

        # ── karta: limit MB ──
        card5 = self._card(self, "Maksymalny rozmiar pliku (MB)")
        card5.pack(fill="x", **pad)

        size_wrap = ctk.CTkFrame(card5, fg_color="transparent")
        size_wrap.pack(padx=12, pady=10)

        self.entry_mb = ctk.CTkEntry(size_wrap, width=110,
                                     justify="center",
                                     font=ctk.CTkFont(size=14, weight="bold"),
                                     border_color=HOTPINK)
        self.entry_mb.insert(0, "5")
        self.entry_mb.pack(side="left")

        ctk.CTkLabel(size_wrap, text=" MB", text_color="#666666").pack(side="left", padx=(8, 0))

        # ── przycisk START ──
        self.btn_start = ctk.CTkButton(
            self, text="▶  Podziel plik",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=HOTPINK, hover_color=HOTPINK_HOVER,
            height=44, corner_radius=10,
            command=self._start
        )
        self.btn_start.pack(padx=20, pady=(8, 4), fill="x")

        # ── pasek postępu ──
        self.progress = ctk.CTkProgressBar(self, progress_color=HOTPINK,
                                           height=10, corner_radius=5)
        self.progress.set(0)
        self.progress.pack(padx=20, pady=(4, 2), fill="x")

        self.lbl_status = ctk.CTkLabel(self, text="", text_color="#555555",
                                       font=ctk.CTkFont(size=11))
        self.lbl_status.pack()

        # ── log ──
        log_frame = ctk.CTkFrame(self, fg_color=FRAME_BG, corner_radius=8,
                                 border_width=1, border_color="#E0E0E0")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(6, 16))

        self.log_box = ctk.CTkTextbox(log_frame, activate_scrollbars=True,
                                      font=ctk.CTkFont(family="Consolas", size=11),
                                      fg_color=FRAME_BG, text_color="#333333",
                                      state="disabled", wrap="none")
        self.log_box.pack(fill="both", expand=True, padx=4, pady=4)

        self._on_mode_change()

    # ── pomocnicze ─────────────────────────────────────────────────────────────

    @staticmethod
    def _card(parent, label: str) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color=FRAME_BG, corner_radius=8,
                     border_width=1, border_color="#E0E0E0")
        ctk.CTkLabel(frame, text=f" {label}",
                 font=ctk.CTkFont(size=11, weight="bold"),
                 text_color=HOTPINK, anchor="w").pack(anchor="w", padx=10, pady=(6, 0))
        return frame

    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.yview("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _on_mode_change(self):
        mode = self._split_mode.get()
        if mode == "rows":
            self.entry_rows.configure(state="normal")
            self.entry_mb.configure(state="disabled")
        else:
            self.entry_rows.configure(state="disabled")
            self.entry_mb.configure(state="normal")

    # ── zdarzenia ──────────────────────────────────────────────────────────────

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik CSV",
            filetypes=[("Pliki CSV", "*.csv"), ("Wszystkie pliki", "*.*")]
        )
        if path:
            self._selected_file = path
            self.lbl_file.configure(text=path, text_color="#222222")
            if not self._out_dir:
                self.lbl_out.configure(
                    text=f"(domyślnie: {os.path.dirname(path)})",
                    text_color="#888888"
                )

    def _pick_out(self):
        folder = filedialog.askdirectory(title="Wybierz folder wyjściowy")
        if folder:
            self._out_dir = folder
            self.lbl_out.configure(text=folder, text_color="#222222")

    def _start(self):
        if not self._selected_file:
            messagebox.showwarning("Brak pliku", "Najpierw wybierz plik CSV.")
            return

        split_mode = self._split_mode.get()
        chunk_size: int | None = None
        max_size_mb: float | None = None

        if split_mode == "rows":
            try:
                chunk_size = int(self.entry_rows.get())
                if chunk_size < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Błąd", "Podaj poprawną liczbę wierszy (≥ 1).")
                return
        else:
            try:
                max_size_mb = float(self.entry_mb.get().strip().replace(",", "."))
                if max_size_mb <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Błąd", "Podaj poprawny rozmiar pliku w MB (> 0).")
                return

        out_dir = self._out_dir or os.path.dirname(self._selected_file)

        self._log_clear()
        self.progress.set(0)
        self.lbl_status.configure(text="Przetwarzanie…", text_color="#555555")
        self.btn_start.configure(state="disabled")

        if split_mode == "rows":
            self._log(f"Tryb: po liczbie wierszy ({chunk_size} na plik)")
        else:
            self._log(f"Tryb: po rozmiarze pliku (maks. {max_size_mb:g} MB)")

        def run():
            split_csv(
                filepath=self._selected_file,
                out_dir=out_dir,
                split_mode=split_mode,
                chunk_size=chunk_size,
                max_size_mb=max_size_mb,
                log_callback=lambda msg: self.after(0, self._log, msg),
                done_callback=lambda total, n, err: self.after(
                    0, self._finish, total, n, err)
            )

        threading.Thread(target=run, daemon=True).start()

    def _finish(self, total: int, num_chunks: int, error: str | None):
        self.btn_start.configure(state="normal")
        self.progress.set(1)

        if error:
            self.lbl_status.configure(
                text=f"Błąd: {error}", text_color="#CC0000")
            messagebox.showerror("Błąd", error)
        else:
            self.lbl_status.configure(
                text=f"Gotowe! {total} wierszy → {num_chunks} plik{'ów' if num_chunks != 1 else ''}.",
                text_color="#228B22"
            )
            self._log(f"\n✅  Podzielono {total} wierszy na {num_chunks} pliki.")


# ── uruchomienie ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
