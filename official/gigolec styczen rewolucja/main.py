import csv
import requests
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict
import threading
import os
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

def attr_by_name(offer, name):
    el = offer.find(f".//a[@name='{name}']")
    return el.text.strip() if el is not None and el.text else ""

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

FIELD_MAP = {
    "id": lambda o: attr(o, "id"),
    # id_bl nie jest atrybutem <o>, siedzi jako <attrs><a name="id_bl">…</a>
    "id_bl": lambda o: attr_by_name(o, "id_bl"),
    "price": lambda o: attr(o, "price"),
    "stock": lambda o: attr(o, "stock"),
    "avail": lambda o: attr(o, "avail"),
    "weight": lambda o: attr(o, "weight"),
    "product_url": lambda o: attr(o, "url"),
    "name": lambda o: text(o, "name"),
    "category": lambda o: text(o, "cat"),
    "description": lambda o: text(o, "desc"),
    "ean": lambda o: attr_by_name(o, "EAN"),
    "sku": lambda o: attr_by_name(o, "sku_bl"),
    "producer": lambda o: attr_by_name(o, "Producent"),
    "images": extract_images,
}

# ---------------- CORE PROCESSING ----------------

def run_processing(merchant_set, fields, mode, output_file, log, cancel_event=None):
    if cancel_event and cancel_event.is_set():
        log("Stopped by user")
        return

    log("Fetching feed list…")
    urls = get_feed_urls()

    if merchant_set:
        urls = [u for u in urls if extract_merchant(u) in merchant_set]

    log(f"Feeds selected: {len(urls)}")

    rows = []

    for url in urls:
        if cancel_event and cancel_event.is_set():
            log("Stopped by user")
            return

        merchant = extract_merchant(url)
        log(f"Processing {merchant}")

        try:
            r = requests.get(url, timeout=HTTP_TIMEOUT)
            r.raise_for_status()
            root = ET.fromstring(r.content)
        except Exception as e:
            log(f"ERROR {merchant}: {e}")
            continue

        for offer in root.findall(".//o"):
            if cancel_event and cancel_event.is_set():
                log("Stopped by user")
                return

            row = {}
            images = []

            for f in fields:
                if f == "images":
                    images = FIELD_MAP[f](offer)
                else:
                    row[f] = FIELD_MAP[f](offer)

            row["_images"] = images
            row["source_feed"] = merchant
            rows.append(row)

    if mode == MODE_DEDUP:
        log("Deduplicating by EAN (lowest price)…")
        by_ean = defaultdict(list)

        for r in rows:
            by_ean[r.get("ean", "")].append(r)

        deduped = []
        for items in by_ean.values():
            try:
                items.sort(key=lambda x: float(x.get("price", "0")))
            except ValueError:
                pass
            deduped.append(items[0])

        rows = deduped

    log("Preparing CSV columns…")

    max_imgs = max((len(r["_images"]) for r in rows), default=0)

    image_columns = []
    if "images" in fields:
        image_columns = ["img_main"] + [f"img_{i}" for i in range(1, max_imgs)]

    final_fields = [f for f in fields if f != "images"] + image_columns + ["source_feed"]

    log("Writing output CSV…")
    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=final_fields)
        writer.writeheader()

        for r in rows:
            if cancel_event and cancel_event.is_set():
                log("Stopped by user")
                return

            out = {k: r.get(k, "") for k in final_fields}

            imgs = r["_images"]
            if imgs:
                out["img_main"] = imgs[0]
                for i, url in enumerate(imgs[1:], start=1):
                    out[f"img_{i}"] = url

            writer.writerow(out)

    log(f"DONE — rows written: {len(rows)}")

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
        self.output_file = ctk.StringVar(value="output.csv")
        self.merchant_status = ctk.StringVar(value="Using all merchants")
        self.cancel_event = threading.Event()
        self.worker_thread = None

        self.field_vars = {f: ctk.BooleanVar(value=True) for f in FIELD_MAP}

        self.build_ui()

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
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.update_idletasks()

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
            args=(merchants, fields),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_worker(self, merchants, fields):
        try:
            run_processing(
                merchants,
                fields,
                self.mode.get(),
                self.output_file.get(),
                self.log,
                self.cancel_event,
            )
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
