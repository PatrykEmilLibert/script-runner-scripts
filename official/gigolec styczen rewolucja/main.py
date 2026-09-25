import csv
import requests
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import tempfile
import queue
from collections import deque
from concurrent.futures import ThreadPoolExecutor, wait
import customtkinter as ctk

API_URL = "https://sm-prods.com/api/files/feeds"
FEED_PREFIX = "https://sm-prods.com/feeds/"
HTTP_TIMEOUT = 60

MODE_ALL = "all_products"
MODE_DEDUP = "deduplicated"

HOT_PINK = "#ff1493"
HOT_PINK_HOVER = "#ff4db8"
HOT_PINK_DARK = "#c71585"
PANEL_BG = "#ffe3f3"
INNER_BG = "#fff4fb"
CARD_BG = "#ffd9ef"
SOFT_TEXT = "#8a2b63"

# ---------------- FEED DISCOVERY ----------------

def _iter_strings(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, list):
        for x in obj:
            yield from _iter_strings(x)
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from _iter_strings(v)

def get_feed_urls():
    r = requests.get(API_URL, timeout=HTTP_TIMEOUT)
    r.raise_for_status()
    data = r.json()

    urls = []
    for s in _iter_strings(data):
        if isinstance(s, str) and s.endswith("_zero.xml"):
            name = s.split("/")[-1]
            urls.append(f"{FEED_PREFIX}{name}")

    return sorted(set(urls))

def extract_merchant(url):
    return url.rsplit("/", 1)[-1].replace("_zero.xml", "").strip()

# ---------------- XML HELPERS ----------------

def attr(offer, key):
    return (offer.attrib.get(key) or "").strip()

def text(parent, tag):
    el = parent.find(tag)
    return el.text.strip() if el is not None and el.text else ""

def collect_named_attrs(offer):
    # słownik <a name=...> budowany raz na ofertę zamiast XPath per pole
    attrs = {}
    for a in offer.iter("a"):
        n = a.attrib.get("name")
        if n and n not in attrs:
            attrs[n] = a.text.strip() if a.text else ""
    return attrs

def extract_images(offer):
    imgs = offer.find("imgs")
    if imgs is None:
        return []

    urls = []

    main = imgs.find("main")
    if main is not None and main.attrib.get("url"):
        urls.append(main.attrib["url"].strip())

    for i in imgs.findall("i"):
        if i.attrib.get("url"):
            urls.append(i.attrib["url"].strip())

    return urls

# każda funkcja dostaje (offer, słownik <a name=...>)
FIELD_MAP = {
    "id": lambda o, a: attr(o, "id"),
    # id_bl nie jest atrybutem <o>, siedzi jako <attrs><a name="id_bl">…</a>
    "id_bl": lambda o, a: a.get("id_bl", ""),
    "price": lambda o, a: attr(o, "price"),
    "stock": lambda o, a: attr(o, "stock"),
    "avail": lambda o, a: attr(o, "avail"),
    "weight": lambda o, a: attr(o, "weight"),
    "product_url": lambda o, a: attr(o, "url"),
    "name": lambda o, a: text(o, "name"),
    "category": lambda o, a: text(o, "cat"),
    "description": lambda o, a: text(o, "desc"),
    "ean": lambda o, a: a.get("EAN", ""),
    "sku": lambda o, a: a.get("sku_bl", ""),
    "producer": lambda o, a: a.get("Producent", ""),
    "images": lambda o, a: extract_images(o),
}

IMG_SEP = "\x1f"            # separator zdjęć w pliku tymczasowym
DOWNLOAD_CHUNK = 1 << 20    # 1 MB
WRITE_BUFFER = 1 << 20
PROGRESS_EVERY = 100_000
PARALLEL_DOWNLOADS = 6      # domyślna liczba feedów pobieranych naraz
PARALLEL_CHOICES = ["1", "2", "4", "6", "8", "12", "16"]

# długie opisy w pliku tymczasowym przekraczają domyślne 128 KB na pole
csv.field_size_limit(2**31 - 1)


def download_to_file(url, path, cancel_event=None):
    """Pobiera feed strumieniowo na dysk — XML nie ląduje w RAM."""
    with requests.get(url, timeout=HTTP_TIMEOUT, stream=True) as r:
        r.raise_for_status()
        with open(path, "wb") as f:
            for chunk in r.iter_content(DOWNLOAD_CHUNK):
                if cancel_event and cancel_event.is_set():
                    return False
                f.write(chunk)
    return True


def _silent_remove(path):
    try:
        os.remove(path)
    except OSError:
        pass


def fetch_feed(url, tmp_dir, cancel_event=None):
    """Wątek pobierający: każdy feed do własnego pliku tymczasowego. Zwraca ścieżkę albo None po Stop."""
    fd, path = tempfile.mkstemp(suffix=".xml", prefix="shumee_feed_", dir=tmp_dir)
    os.close(fd)
    try:
        ok = download_to_file(url, path, cancel_event)
    except BaseException:
        _silent_remove(path)
        raise
    if not ok:
        _silent_remove(path)
        return None
    return path


def iter_offers(path):
    """iterparse po <o>; po obsłudze oferta jest usuwana z rodzica, więc drzewo nie rośnie."""
    stack = []
    for event, elem in ET.iterparse(path, events=("start", "end")):
        if event == "start":
            stack.append(elem)
            continue
        stack.pop()
        if elem.tag == "o":
            yield elem
            elem.clear()
            if stack:
                stack[-1].remove(elem)

# ---------------- CORE PROCESSING ----------------

def _parse_price(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return float("inf")


def run_processing(merchant_set, fields, mode, output_file, log, cancel_event=None,
                   parallel=PARALLEL_DOWNLOADS):
    """
    Dwa przebiegi, oba strumieniowe:
      1) feedy pobierane równolegle (parallel wątków) do osobnych plików XML na dysku,
         parsowane po kolei w kolejności listy -> iterparse -> wiersz od razu do tymczasowego CSV
         (w RAM tylko jedna oferta naraz + przy dedup mały słownik EAN -> najtańszy wiersz)
      2) tymczasowy CSV -> docelowy CSV, gdy już wiadomo ile kolumn img_* potrzeba
    """
    def cancelled():
        if cancel_event and cancel_event.is_set():
            log("Stopped by user")
            return True
        return False

    if cancelled():
        return

    log("Fetching feed list…")
    urls = get_feed_urls()

    if merchant_set:
        urls = [u for u in urls if extract_merchant(u) in merchant_set]

    log(f"Feeds selected: {len(urls)}")

    base_fields = [f for f in fields if f != "images"]
    want_images = "images" in fields
    tmp_header = base_fields + ["source_feed", "_images"]

    out_dir = os.path.dirname(os.path.abspath(output_file))
    fd, tmp_csv = tempfile.mkstemp(suffix=".csv", prefix="shumee_rows_", dir=out_dir)
    os.close(fd)

    parallel = max(1, int(parallel))
    # ile feedów może czekać na dysku przed parsowaniem — ogranicza zajęte miejsce
    window = parallel * 2
    pool = ThreadPoolExecutor(max_workers=parallel, thread_name_prefix="feed-dl")
    pending = deque()   # (url, future) w kolejności listy feedów
    url_iter = iter(urls)

    def refill():
        while len(pending) < window:
            url = next(url_iter, None)
            if url is None:
                return
            pending.append((url, pool.submit(fetch_feed, url, out_dir, cancel_event)))

    ean_idx = base_fields.index("ean") if "ean" in base_fields else None
    price_idx = base_fields.index("price") if "price" in base_fields else None
    best_by_ean = {}   # ean -> (cena, numer wiersza)
    max_imgs = 0
    total = 0

    try:
        # ---------- PRZEBIEG 1 ----------
        with open(tmp_csv, "w", newline="", encoding="utf-8", buffering=WRITE_BUFFER) as tf:
            tw = csv.writer(tf)
            tw.writerow(tmp_header)

            log(f"Downloading up to {parallel} feeds in parallel")
            refill()
            done_feeds = 0

            while pending:
                if cancelled():
                    return

                url, fut = pending[0]
                merchant = extract_merchant(url)

                # czekaj na pobranie, ale reaguj na Stop
                while not fut.done():
                    if cancelled():
                        return
                    wait([fut], timeout=0.5)
                pending.popleft()

                try:
                    tmp_xml = fut.result()
                except Exception as e:
                    log(f"ERROR {merchant}: {e}")
                    refill()
                    continue
                finally:
                    done_feeds += 1

                refill()
                if tmp_xml is None:
                    cancelled()
                    return

                log(f"Processing {merchant} ({done_feeds}/{len(urls)})")
                merchant_rows = 0
                try:
                    for offer in iter_offers(tmp_xml):
                        if cancel_event and cancel_event.is_set():
                            break

                        named = collect_named_attrs(offer)
                        row = [FIELD_MAP[f](offer, named) for f in base_fields]
                        images = extract_images(offer) if want_images else []
                        if len(images) > max_imgs:
                            max_imgs = len(images)

                        row.append(merchant)
                        row.append(IMG_SEP.join(images))
                        tw.writerow(row)

                        if mode == MODE_DEDUP:
                            ean = row[ean_idx] if ean_idx is not None else ""
                            price = _parse_price(row[price_idx]) if price_idx is not None else 0.0
                            best = best_by_ean.get(ean)
                            if best is None or price < best[0]:
                                best_by_ean[ean] = (price, total)

                        total += 1
                        merchant_rows += 1
                        if total % PROGRESS_EVERY == 0:
                            log(f"  …{total:,} rows collected")
                except ET.ParseError as e:
                    log(f"ERROR {merchant}: XML parse error after {merchant_rows} offers: {e}")
                finally:
                    _silent_remove(tmp_xml)

                if cancelled():
                    return
                log(f"  {merchant}: {merchant_rows:,} offers")

        keep = None
        if mode == MODE_DEDUP:
            log("Deduplicating by EAN (lowest price)…")
            keep = {idx for _, idx in best_by_ean.values()}
            best_by_ean.clear()

        # ---------- PRZEBIEG 2 ----------
        image_columns = []
        if want_images:
            image_columns = ["img_main"] + [f"img_{i}" for i in range(1, max_imgs)]
        final_fields = base_fields + image_columns + ["source_feed"]
        img_slots = len(image_columns)

        log("Writing output CSV…")
        written = 0
        with open(tmp_csv, newline="", encoding="utf-8", buffering=WRITE_BUFFER) as tf, \
             open(output_file, "w", newline="", encoding="utf-8", buffering=WRITE_BUFFER) as f:
            reader = csv.reader(tf)
            next(reader)  # nagłówek tymczasowy
            writer = csv.writer(f)
            writer.writerow(final_fields)

            for idx, r in enumerate(reader):
                if keep is not None and idx not in keep:
                    continue
                if written % PROGRESS_EVERY == 0 and cancelled():
                    return

                base = r[:-2]
                source_feed = r[-2]
                imgs = r[-1].split(IMG_SEP) if r[-1] else []
                if img_slots:
                    base.extend(imgs + [""] * (img_slots - len(imgs)))
                base.append(source_feed)
                writer.writerow(base)

                written += 1
                if written % PROGRESS_EVERY == 0:
                    log(f"  …{written:,} rows written")

        log(f"DONE — rows written: {written:,}")
    finally:
        # Stop/błąd: nie startuj kolejnych pobrań, a już ściągnięte pliki usuń
        for _, fut in pending:
            fut.cancel()
        pool.shutdown(wait=True)
        for _, fut in pending:
            if fut.done() and not fut.cancelled() and fut.exception() is None and fut.result():
                _silent_remove(fut.result())
        _silent_remove(tmp_csv)


# ---------------- GUI ----------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("Shumee XML → CSV Tool")
        self.geometry("980x760")

        self.merchant_csv = ctk.StringVar()
        self.use_all_merchants = ctk.BooleanVar(value=True)
        self.mode = ctk.StringVar(value=MODE_ALL)
        self.parallel = ctk.StringVar(value=str(PARALLEL_DOWNLOADS))
        self.output_file = ctk.StringVar(value="output.csv")
        self.merchant_status = ctk.StringVar(value="Using all merchants")
        self.cancel_event = threading.Event()
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self.field_vars = {f: ctk.BooleanVar(value=True) for f in FIELD_MAP}

        self.build_ui()
        self._drain_log()

    def build_ui(self):
        container = ctk.CTkFrame(self, corner_radius=12, fg_color=PANEL_BG)
        container.pack(fill="both", expand=True, padx=12, pady=12)

        inner = ctk.CTkScrollableFrame(container, corner_radius=10, fg_color=INNER_BG)
        inner.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            inner,
            text="Merchants",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=HOT_PINK,
        ).pack(anchor="w", pady=(4, 4))

        ctk.CTkCheckBox(
            inner,
            text="Use all merchants",
            variable=self.use_all_merchants,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            border_color=HOT_PINK_DARK,
            checkmark_color="white",
        ).pack(fill="x")

        ctk.CTkButton(
            inner,
            text="Select merchant CSV",
            command=self.select_merchant_csv,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            text_color="white",
        ).pack(fill="x", pady=4)
        ctk.CTkLabel(inner, textvariable=self.merchant_status, text_color=SOFT_TEXT).pack(anchor="w", pady=(0, 8))

        ctk.CTkFrame(inner, height=2, fg_color=HOT_PINK_DARK, corner_radius=2).pack(fill="x", pady=8)

        ctk.CTkLabel(
            inner,
            text="Mode",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=HOT_PINK,
        ).pack(anchor="w", pady=(4, 4))

        ctk.CTkRadioButton(
            inner,
            text="All products",
            variable=self.mode,
            value=MODE_ALL,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            border_color=HOT_PINK_DARK,
        ).pack(anchor="w")
        ctk.CTkRadioButton(
            inner,
            text="Deduplicated (EAN, lowest price)",
            variable=self.mode,
            value=MODE_DEDUP,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            border_color=HOT_PINK_DARK,
        ).pack(anchor="w", pady=(0, 8))

        parallel_row = ctk.CTkFrame(inner, fg_color="transparent")
        parallel_row.pack(anchor="w", pady=(0, 8))
        ctk.CTkLabel(
            parallel_row,
            text="Feeds downloaded in parallel:",
            text_color=SOFT_TEXT,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkOptionMenu(
            parallel_row,
            values=PARALLEL_CHOICES,
            variable=self.parallel,
            width=80,
            fg_color=HOT_PINK,
            button_color=HOT_PINK_DARK,
            button_hover_color=HOT_PINK_HOVER,
        ).pack(side="left")

        ctk.CTkFrame(inner, height=2, fg_color=HOT_PINK_DARK, corner_radius=2).pack(fill="x", pady=8)

        ctk.CTkLabel(
            inner,
            text="Fields",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=HOT_PINK,
        ).pack(anchor="w", pady=(4, 4))
        ctk.CTkButton(
            inner,
            text="Select / Unselect all fields",
            command=self.toggle_all_fields,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            text_color="white",
        ).pack(fill="x", pady=(0, 6))

        fields_frame = ctk.CTkFrame(inner, corner_radius=10, fg_color=CARD_BG)
        fields_frame.pack(fill="x", pady=(0, 8))

        columns = 3
        for col in range(columns):
            fields_frame.grid_columnconfigure(col, weight=1)

        for idx, (f, v) in enumerate(self.field_vars.items()):
            row, col = divmod(idx, columns)
            ctk.CTkCheckBox(
                fields_frame,
                text=f,
                variable=v,
                fg_color=HOT_PINK,
                hover_color=HOT_PINK_HOVER,
                border_color=HOT_PINK_DARK,
                checkmark_color="white",
            ).grid(row=row, column=col, padx=8, pady=6, sticky="w")

        ctk.CTkFrame(inner, height=2, fg_color=HOT_PINK_DARK, corner_radius=2).pack(fill="x", pady=8)

        ctk.CTkLabel(
            inner,
            text="Output file",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=HOT_PINK,
        ).pack(anchor="w", pady=(4, 4))
        ctk.CTkEntry(inner, textvariable=self.output_file).pack(fill="x")
        ctk.CTkButton(
            inner,
            text="Change output CSV",
            command=self.select_output_csv,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            text_color="white",
        ).pack(fill="x", pady=4)
        actions_frame = ctk.CTkFrame(inner, fg_color="transparent")
        actions_frame.pack(fill="x", pady=8)
        actions_frame.grid_columnconfigure(0, weight=1)
        actions_frame.grid_columnconfigure(1, weight=1)

        self.run_button = ctk.CTkButton(
            actions_frame,
            text="RUN",
            command=self.run,
            fg_color=HOT_PINK,
            hover_color=HOT_PINK_HOVER,
            text_color="white",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=38,
        )
        self.run_button.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        self.stop_button = ctk.CTkButton(
            actions_frame,
            text="STOP",
            command=self.stop,
            fg_color=HOT_PINK_DARK,
            hover_color=HOT_PINK_HOVER,
            text_color="white",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=38,
            state="disabled",
        )
        self.stop_button.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        self.log_box = ctk.CTkTextbox(inner, height=220, corner_radius=10)
        self.log_box.pack(fill="both", expand=True)

    def toggle_all_fields(self):
        any_unchecked = any(not v.get() for v in self.field_vars.values())
        for v in self.field_vars.values():
            v.set(any_unchecked)

    def log(self, msg):
        # wołane z wątku roboczego — tylko kolejka, widżety dotyka _drain_log w wątku GUI
        self.log_queue.put(msg)

    def _drain_log(self):
        lines = []
        try:
            while True:
                lines.append(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        if lines:
            self.log_box.insert("end", "\n".join(lines) + "\n")
            self.log_box.see("end")
        self.after(100, self._drain_log)

    def select_merchant_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not path:
            return

        self.merchant_csv.set(path)
        self.use_all_merchants.set(False)

        with open(path, newline="") as f:
            merchants = [row[0].strip() for row in csv.reader(f) if row and row[0].strip()]

        self.merchant_status.set(f"{os.path.basename(path)} ({len(merchants)} merchants)")

    def select_output_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if path:
            self.output_file.set(path)

    def run(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.log("Process is already running")
            return

        fields = [f for f, v in self.field_vars.items() if v.get()]
        if not fields:
            messagebox.showerror("Error", "Select at least one field")
            return

        merchants = None
        if not self.use_all_merchants.get() and self.merchant_csv.get():
            with open(self.merchant_csv.get(), newline="") as f:
                merchants = {row[0].strip() for row in csv.reader(f) if row and row[0].strip()}

        self.cancel_event.clear()
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")

        self.worker_thread = threading.Thread(
            target=self._run_worker,
            args=(merchants, fields, self.mode.get(), self.output_file.get(), int(self.parallel.get())),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_worker(self, merchants, fields, mode, output_file, parallel):
        # zmienne Tk odczytane w run() — wątek roboczy nie dotyka widżetów
        try:
            run_processing(
                merchants,
                fields,
                mode,
                output_file,
                self.log,
                self.cancel_event,
                parallel=parallel,
            )
        except Exception as e:
            self.log(f"ERROR: {e}")
        finally:
            self.after(0, self._set_idle_state)

    def _set_idle_state(self):
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")

    def stop(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.cancel_event.set()
            self.log("Stopping process…")

# ---------------- START ----------------

if __name__ == "__main__":
    App().mainloop()
