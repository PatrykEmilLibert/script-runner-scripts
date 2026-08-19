# -*- coding: utf-8 -*-
"""
GUI (PySide6/Qt) + logika dla feedów Joom.

Na wzór gui_qt_pobieranie.py, ale:
  - pobiera JEDEN feed na nazwę: {BASE}/{nazwa}_joom_SE.xml
  - pakuje dane bezpośrednio do szablonu Joom "products_template" (arkusz
    "Records", 3 wiersze nagłówka + dane od wiersza 4) — bez potrzeby pliku
    szablonu; nagłówki są wpisane na sztywno w tym skrypcie.
  - CAŁA logika ORAZ warstwa GUI (styl Qt, most sygnałów, pasek postępu)
    znajdują się w TYM jednym pliku — brak zależności od innych modułów
    projektu (nie importujemy gui_qt_amno_common).

Uruchom:  python gui_qt_joom.py
"""

from __future__ import annotations

import concurrent.futures
import os
import re
import sys
import tempfile
import threading
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
from urllib.parse import urlparse

import openpyxl
import requests

from PySide6.QtCore import QObject, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QGridLayout, QGroupBox, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
    QVBoxLayout, QWidget,
)

# --- Ustawienia ogólne ---
DOMYSLNA_SCIEZKA_ZAPISU = os.path.join(os.path.expanduser("~"), "Downloads")
MAX_WORKERS = 10
EXCEL_MAX_LEN = 32767
EXCEL_MAX_ROWS = 1048576
_ILLEGAL_CTRL_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

# --- Ustawienia specyficzne dla Joom ---
FEED_URL_TEMPLATE = "https://sm-prods.com/feeds/{nazwa}_joomSE.xml"
DEFAULT_STORE_ID = "63d25e2be5b9375110813782"
DEFAULT_CURRENCY = "EUR"
VAT_DIVISOR = 1.25          # cena w feedzie jest brutto -> netto = brutto / 1.25
DEFAULT_SHIPPING_PRICE = 0  # Shipping Price (without VAT) (default warehouse)
MAX_EXTRA_IMAGES = 20       # "Extra Image URLs (max 20)"
EXTRA_IMAGES_SEP = "\n"     # każdy dodatkowy URL w nowej linii

# Rozpoznawanie atrybutów (case-insensitive) w <attrs>
BRAND_KEYS = {"producent", "producer", "brand", "marka"}
EAN_KEYS = {"ean", "gtin", "barcode"}
COLOR_KEYS = {"color", "kolor"}
SIZE_KEYS = {"size", "rozmiar", "rozmiar produktu"}

# --- API sm-prods: lista zakazanych EAN-ów (forbidden_eans_paginated) ---
SMPRODS_TOKEN = "y7SeKeGSfVZtH9dCxwVULWTbcfWBrVq2WKcJssq8Pz8o5t3DFDpQ12BGRGc1S3fOJ2UC3tRMi29ChrseLAsl4GhHKR3Y9ALr9Zfq8pyeYtlExRas7rOfvRTrqKdEOJ8y"
FORBIDDEN_EANS_PAGINATED_URL = "https://api-sm-prods.sm-prods.com/forbidden_eans_paginated"

# Podmieniane przez GUI na MessageboxShim (dialogi kierowane do Qt).
messagebox = None

# --- Nagłówek szablonu Joom "Records" (3 wiersze), wpisany na sztywno ---
JOOM_COLUMNS = [
    "Product SKU", "Name", "Description", "Search Tags", "Labels", "Brand",
    "Landing Page URL", "Product Main Image URL", "Extra Image URLs (max 20)",
    "Suggested Category ID", "Dangerous Kind", "Video ID", "Store ID",
    "Rich Content EN", "Rich Content RU", "Rich Content DE", "Rich Content ES",
    "Variant SKU", "Variant Main Image URL", "Shipping Weight (kg)",
    "Shipping Length (cm)", "Shipping Width (cm)", "Shipping Height (cm)",
    "Manufacture GTIN", "HS Code", "Color", "Size", "Price (without VAT)",
    "Currency", "Inventory (default warehouse)",
    "Shipping Price (without VAT) (default warehouse)", "Declared Value",
    "MSRP", "EPREL ID",
]

JOOM_ROW_REQUIRED = [
    "(required, irreversible)", "(required)", "(required)", "(recommended)",
    "(optional)", "(recommended)", "(recommended)", "(required)",
    "(recommended)", "(recommended)", "(optional)", "(recommended)",
    "(required, irreversible)", "(optional)", "(optional)", "(optional)",
    "(optional)", "(required, irreversible)", "(recommended)", "(optional)",
    "(optional)", "(optional)", "(optional)", "(recommended)", "(optional)",
    "(recommended)", "(recommended)", "(required)", "(required)", "(required)",
    "(required)", "(optional)", "(recommended)", "(optional)",
]

# Kolumny 1-17 dotyczą produktu, 18-34 wariantu.
JOOM_ROW_SCOPE = (
    ["(product related field)"] * 17 + ["(variant related field)"] * 17
)


# ============================ Warstwa Qt (styl/most/postęp) ============================
# (Wcześniej w gui_qt_amno_common — teraz wbudowane, aby skrypt był samodzielny.)

# Paleta (jasne tło / różowe akcenty).
ACCENT = "#ff69b4"
ACCENT_HOVER = "#e754a7"
APP_BG = "#fff7fb"
PANEL_BG = "#ffffff"
INPUT_BG = "#fff2f8"
BORDER = "#f7b3d2"
TEXT = "#3d2130"
MUTED = "#7d5a6b"
TROUGH = "#f5d7e6"

QSS = f"""
QWidget {{ background-color: {APP_BG}; color: {TEXT}; font-size: 13px; }}
QGroupBox {{
    background-color: {PANEL_BG};
    border: 1px solid {BORDER};
    border-radius: 12px;
    margin-top: 16px;
    padding: 10px 10px 8px 10px;
    font-weight: bold;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 2px 6px;
    color: {ACCENT_HOVER};
}}
QLabel {{ background: transparent; color: {TEXT}; font-weight: normal; }}
QLineEdit, QPlainTextEdit, QComboBox {{
    background-color: {INPUT_BG};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 5px 8px;
    color: {TEXT};
    selection-background-color: {ACCENT};
    selection-color: #ffffff;
}}
QLineEdit:focus, QPlainTextEdit:focus {{ border: 1px solid {ACCENT}; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background-color: {PANEL_BG};
    border: 1px solid {BORDER};
    selection-background-color: {ACCENT};
    selection-color: #ffffff;
}}
QPushButton {{
    background-color: {ACCENT};
    color: #ffffff;
    border: none;
    border-radius: 10px;
    padding: 9px 16px;
    font-weight: bold;
}}
QPushButton:hover {{ background-color: {ACCENT_HOVER}; }}
QPushButton:pressed {{ background-color: {ACCENT_HOVER}; }}
QPushButton:disabled {{ background-color: {TROUGH}; color: #ffffff; }}
QCheckBox {{ background: transparent; color: {TEXT}; spacing: 8px; font-weight: normal; }}
QProgressBar {{
    background-color: {TROUGH};
    border: 1px solid {BORDER};
    border-radius: 8px;
    text-align: center;
    color: {TEXT};
    min-height: 18px;
}}
QProgressBar::chunk {{ background-color: {ACCENT}; border-radius: 7px; }}
QScrollBar:vertical {{ background: {INPUT_BG}; width: 12px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {BORDER}; border-radius: 6px; min-height: 24px; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
"""


def apply_scaling(factor: str = "0.9") -> None:
    """Skala całego UI. Musi być ustawione PRZED utworzeniem QApplication."""
    os.environ.setdefault("QT_SCALE_FACTOR", factor)


class Bridge(QObject):
    """Sygnały do bezpiecznego przekazywania zdarzeń z wątków workerów do GUI."""
    status_sig = Signal(str, float)      # tekst, ratio (-1 = nie zmieniaj paska)
    reset_sig = Signal(int)              # opóźnienie w ms do resetu statusu
    dialog_sig = Signal(str, str, str)   # rodzaj(info/warn/error), tytuł, treść
    savedir_sig = Signal(str)            # nowa ścieżka zapisu (fallback z workera)


class MessageboxShim:
    """Kieruje dialogi wołane z wątku roboczego do Qt (wątkowo-bezpiecznie)."""
    def __init__(self, bridge: "Bridge"):
        self._bridge = bridge

    def showinfo(self, title, message):
        self._bridge.dialog_sig.emit("info", str(title), str(message))

    def showwarning(self, title, message):
        self._bridge.dialog_sig.emit("warn", str(title), str(message))

    def showerror(self, title, message):
        self._bridge.dialog_sig.emit("error", str(title), str(message))


class _SavedirShim:
    """Zastępuje pole ścieżki, gdy worker wpisuje fallbackową ścieżkę zapisu."""
    def __init__(self, bridge: "Bridge"):
        self._bridge = bridge

    def delete(self, *args):
        pass

    def insert(self, index, text):
        self._bridge.savedir_sig.emit(str(text))

    def get(self):
        return ""


class BaseController:
    """Interfejs app_instance oczekiwany przez funkcję przetwarzania."""
    def __init__(self, bridge: "Bridge"):
        self.bridge = bridge
        self.pole_sciezki_zapisu = _SavedirShim(bridge)

    def after(self, delay_ms, callback=None, *args):
        """Zgodność z Tk: wywołania self.after(0, ...) dla dialogów.

        Callbacki wołają tylko messagebox (emisję sygnału), więc można je
        wykonać od razu — sygnał i tak trafi do głównego wątku.
        """
        if callable(callback):
            callback(*args)
        return None

    def update_status(self, message, progress_value=None):
        self.bridge.status_sig.emit(
            str(message), -1.0 if progress_value is None else float(progress_value))

    def reset_gui_after_delay(self, delay_ms=4000):
        self.bridge.reset_sig.emit(int(delay_ms))


class ProgressMixin:
    """Wspólne sloty sygnałów i sekcja 'Postęp'."""

    def _connect_bridge(self):
        self.bridge.status_sig.connect(self._on_status)
        self.bridge.reset_sig.connect(self._on_reset)
        self.bridge.dialog_sig.connect(self._on_dialog)
        self.bridge.savedir_sig.connect(self._on_savedir)

    def _on_status(self, text: str, ratio: float):
        self.status_label.setText(text)
        if ratio >= 0.0:
            self.progress_bar.setValue(int(max(0.0, min(1.0, ratio)) * 1000))

    def _on_reset(self, delay_ms: int):
        self.process_btn.setEnabled(True)
        self.process_btn.setText(self._process_btn_text)
        QTimer.singleShot(int(delay_ms), self._reset_status)

    def _reset_status(self):
        self.status_label.setText("Gotowy.")
        self.progress_bar.setValue(0)

    def _on_dialog(self, kind: str, title: str, text: str):
        if kind == "warn":
            QMessageBox.warning(self, title, text)
        elif kind == "error":
            QMessageBox.critical(self, title, text)
        else:
            QMessageBox.information(self, title, text)

    def _on_savedir(self, path: str):
        if hasattr(self, "savedir_edit"):
            self.savedir_edit.setText(path)

    def _begin_running(self, busy_text: str):
        self.process_btn.setEnabled(False)
        self.process_btn.setText(busy_text)


# ============================ Funkcje pomocnicze ============================
def sanitize_for_excel(value):
    """Usuwa nielegalne znaki i przycina tekst do limitu Excela (32 767)."""
    if value is None:
        return ""
    text = str(value)
    if _ILLEGAL_CTRL_CHARS.search(text):
        text = _ILLEGAL_CTRL_CHARS.sub(" ", text)
    if len(text) > EXCEL_MAX_LEN:
        text = text[:EXCEL_MAX_LEN]
    return text


def to_number(value):
    """Próbuje zrzutować wartość na float; zwraca None, gdy się nie da."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        txt = value.strip().replace(",", ".")
        if not txt:
            return None
        try:
            return float(txt)
        except ValueError:
            return None
    return None


def clean_text(text):
    """Zastępuje znaki nowej linii i inne białe znaki pojedynczą spacją."""
    if not text:
        return ""
    return " ".join(text.split())


def cena_netto(raw_price):
    """Cena brutto z feedu -> netto (brutto / VAT_DIVISOR), zaokrąglona do 2 miejsc."""
    num = to_number(raw_price)
    if num is None:
        return sanitize_for_excel(raw_price)
    return round(num / VAT_DIVISOR, 2)


def pobierz_forbidden_eans():
    """
    Pobiera zbiór zakazanych EAN-ów z API sm-prods (paginowany endpoint
    /forbidden_eans_paginated). Zwraca (zbior_ean, blad).
    """
    headers = {"Authorization": f"Bearer {SMPRODS_TOKEN}"}
    try:
        eans = set()
        page = 1
        while True:
            r = requests.get(
                FORBIDDEN_EANS_PAGINATED_URL,
                params={"page": page, "page_size": 1000},
                timeout=60, headers=headers)
            r.raise_for_status()
            payload = r.json()
            for x in payload.get("data", []):
                if x.get("ean"):
                    eans.add(str(x["ean"]).strip())
            if page >= payload.get("total_pages", page):
                break
            page += 1
        return eans, None
    except Exception as e:
        return set(), f"forbidden_eans_paginated: {e}"


def pobierz_xml(url, sciezka_docelowa):
    """Pobiera plik XML z URL. Zwraca (True, None) lub (False, str(e))."""
    try:
        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=30) as response:
            with open(sciezka_docelowa, "wb") as f:
                f.write(response.read())
        return True, None
    except Exception as e:
        return False, str(e)


def _dopasuj_atrybut(attrs_lower, klucze):
    """Zwraca wartość pierwszego atrybutu, którego nazwa (lower) pasuje do zbioru kluczy."""
    for nazwa, wartosc in attrs_lower.items():
        if nazwa in klucze and wartosc:
            return wartosc
    return ""


def parsuj_xml(sciezka_pliku, forbidden_eans=None):
    """
    Parsuje feed Joom (<offers><o>...</o></offers>) i zwraca listę produktów
    już zmapowanych do pól potrzebnych szablonowi Joom.
    Odrzuca produkty z zakazanym EAN (forbidden_eans).
    Zwraca (produkty, None) lub ([], str(e)).
    """
    try:
        produkty = []
        for _, element in ET.iterparse(sciezka_pliku, events=("end",)):
            if element.tag != "o":
                continue

            name_elem = element.find("name")
            cat_elem = element.find("cat")
            desc_elem = element.find("desc")

            # Atrybuty z <attrs>
            attrs_lower = {}
            attrs_elem = element.find("attrs")
            if attrs_elem is not None:
                for a in attrs_elem.findall("a"):
                    nazwa = a.get("name")
                    if nazwa:
                        attrs_lower[nazwa.strip().lower()] = clean_text(a.text)

            ean = _dopasuj_atrybut(attrs_lower, EAN_KEYS)

            # FILTR ZAKAZANYCH EAN-ów
            if forbidden_eans and ean and ean.strip() in forbidden_eans:
                element.clear()
                continue

            # Obrazy: główny + dodatkowe
            main_image = ""
            extra_images = []
            imgs_elem = element.find("imgs")
            if imgs_elem is not None:
                m = imgs_elem.find("main")
                if m is not None and m.get("url"):
                    main_image = m.get("url")
                for img in imgs_elem.findall("i"):
                    if img.get("url"):
                        extra_images.append(img.get("url"))
            # Gdy brak <main>, pierwszy z <i> traktujemy jako główny
            if not main_image and extra_images:
                main_image = extra_images.pop(0)

            produkty.append({
                "id": element.get("id") or "",
                "url": element.get("url") or "",
                "price": element.get("price"),
                "stock": element.get("stock"),
                "weight": element.get("weight"),
                "name": clean_text(name_elem.text) if name_elem is not None else "",
                "cat": clean_text(cat_elem.text) if cat_elem is not None else "",
                "desc": clean_text(desc_elem.text) if desc_elem is not None else "",
                "brand": _dopasuj_atrybut(attrs_lower, BRAND_KEYS),
                "ean": ean,
                "color": _dopasuj_atrybut(attrs_lower, COLOR_KEYS),
                "size": _dopasuj_atrybut(attrs_lower, SIZE_KEYS),
                "main_image": main_image,
                "extra_images": extra_images,
            })
            element.clear()

        return produkty, None

    except FileNotFoundError as e:
        return [], f"Nie znaleziono pliku: {sciezka_pliku} ({e})"
    except ET.ParseError as e:
        return [], f"Błąd parsowania XML w {os.path.basename(sciezka_pliku)}: {e}"
    except Exception as e:
        return [], f"Nieoczekiwany błąd parsowania: {e}"


def zbuduj_wiersz_joom(p, store_id):
    """Buduje listę 34 wartości w kolejności JOOM_COLUMNS dla jednego produktu."""
    sku = sanitize_for_excel(p["id"])
    extra = EXTRA_IMAGES_SEP.join(p["extra_images"][:MAX_EXTRA_IMAGES])
    stock_num = to_number(p["stock"])
    weight_num = to_number(p["weight"])

    return [
        sku,                                    # Product SKU
        sanitize_for_excel(p["name"]),          # Name
        sanitize_for_excel(p["desc"]),          # Description
        "",                                     # Search Tags
        "",                                     # Labels
        "",         # Brand
        "",           # Landing Page URL
        sanitize_for_excel(p["main_image"]),    # Product Main Image URL
        sanitize_for_excel(extra),              # Extra Image URLs (max 20)
        "",                                     # Suggested Category ID
        "",                                     # Dangerous Kind
        "",                                     # Video ID
        sanitize_for_excel(store_id),           # Store ID
        "",                                     # Rich Content EN
        "",                                     # Rich Content RU
        "",                                     # Rich Content DE
        "",                                     # Rich Content ES
        sku,                                    # Variant SKU
        sanitize_for_excel(p["main_image"]),    # Variant Main Image URL
        "",  # Shipping Weight (kg)
        "",                                     # Shipping Length (cm)
        "",                                     # Shipping Width (cm)
        "",                                     # Shipping Height (cm)
        sanitize_for_excel(p["ean"]),           # Manufacture GTIN
        "",                                     # HS Code
        sanitize_for_excel(p["color"]),         # Color
        sanitize_for_excel(p["size"]),          # Size
        cena_netto(p["price"]),                 # Price (without VAT)
        DEFAULT_CURRENCY,                       # Currency
        stock_num if stock_num is not None else 0,  # Inventory (default warehouse)
        DEFAULT_SHIPPING_PRICE,                 # Shipping Price (without VAT)
        "",                                     # Declared Value
        "",                                     # MSRP
        "",                                     # EPREL ID
    ]


def zapisz_do_szablonu_joom(produkty, store_id, sciezka_pliku):
    """Zapisuje produkty do pliku XLSX w formacie szablonu Joom (arkusz 'Records')."""
    try:
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(title="Records")
        ws.append(JOOM_COLUMNS)
        ws.append(JOOM_ROW_REQUIRED)
        ws.append(JOOM_ROW_SCOPE)

        limit = EXCEL_MAX_ROWS - 3  # 3 wiersze nagłówka
        zapisane = 0
        for p in produkty:
            if zapisane >= limit:
                break
            ws.append(zbuduj_wiersz_joom(p, store_id))
            zapisane += 1

        wb.save(sciezka_pliku)
        obciete = max(0, len(produkty) - zapisane)
        return True, None, zapisane, obciete
    except Exception as e:
        return False, str(e), 0, 0


def pobierz_i_parsuj_url(url, forbidden_eans=None):
    """Pobiera i parsuje jeden URL (do uruchamiania w wątku). Zwraca (status, dane)."""
    katalog_tymczasowy = tempfile.gettempdir()
    nazwa_pliku_url = os.path.basename(urlparse(url).path) or f"feed_{hash(url)}.xml"
    nazwa_bazowa = os.path.splitext(nazwa_pliku_url)[0]
    znacznik = datetime.now().strftime("%Y%m%d%H%M%S%f")
    sciezka_lokalna = os.path.join(katalog_tymczasowy, f"temp_{nazwa_bazowa}_{znacznik}.xml")

    sukces, powod = pobierz_xml(url, sciezka_lokalna)
    if not sukces:
        return ("blad_pobierania", (url, powod, nazwa_pliku_url))

    produkty, powod_pars = parsuj_xml(sciezka_lokalna, forbidden_eans)

    try:
        os.remove(sciezka_lokalna)
    except Exception as e:
        print(f"Ostrzeżenie: nie usunięto pliku tymczasowego {sciezka_lokalna}: {e}")

    if powod_pars:
        return ("blad_parsowania", (url, powod_pars, nazwa_pliku_url))
    if not produkty:
        return ("blad_parsowania", (url, "Brak elementów <o> po parsowaniu", nazwa_pliku_url))

    return ("sukces", (produkty, nazwa_pliku_url))


# ============================ Główne przetwarzanie ============================
def przetworz_na_joom(nazwy, sciezka_zapisu, store_id, app_instance):
    """Pobiera feedy Joom dla nazw i zapisuje jeden plik XLSX w formacie szablonu Joom."""
    store_id = (store_id or DEFAULT_STORE_ID).strip() or DEFAULT_STORE_ID

    # Lista zakazanych EAN-ów (zawsze włączona).
    app_instance.update_status("Pobieram listę zakazanych EAN-ów z sm-prods...", 0)
    forbidden_eans, blad_forbidden = pobierz_forbidden_eans()
    if blad_forbidden:
        app_instance.after(0, lambda m=blad_forbidden: messagebox.showwarning(
            "Brak listy forbidden EAN",
            f"Nie udało się pobrać listy zakazanych EAN-ów — pobieranie przebiegnie BEZ tego filtra.\n\n{m}"))

    urls = [FEED_URL_TEMPLATE.format(nazwa=nazwa) for nazwa in nazwy]
    liczba_url = len(urls)
    if liczba_url == 0:
        app_instance.update_status("Nie podano żadnych nazw.", 0)
        app_instance.reset_gui_after_delay()
        return

    if not os.path.exists(sciezka_zapisu):
        try:
            os.makedirs(sciezka_zapisu)
        except Exception as e:
            app_instance.after(0, lambda s=sciezka_zapisu, err=e: messagebox.showerror(
                "Błąd ścieżki zapisu",
                f"Nie można utworzyć katalogu: {s}\nBłąd: {err}\nZapis w katalogu roboczym."))
            sciezka_zapisu = os.getcwd()
            app_instance.pole_sciezki_zapisu.insert(0, sciezka_zapisu)

    dane_po_id = {}   # id -> produkt (deduplikacja)
    sukcesy = 0
    bledy_pobierania = 0
    bledy_parsowania = 0
    bledne_linki = []

    info_eans = f" Odsiew zakazanych EAN: {len(forbidden_eans)}." if forbidden_eans else ""
    app_instance.update_status(
        f"Rozpoczynam pobieranie {liczba_url} feedów (max {MAX_WORKERS} wątków).{info_eans}", 0)

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(pobierz_i_parsuj_url, u, forbidden_eans): u for u in urls}
        zakonczone = 0
        for future in concurrent.futures.as_completed(futures):
            zakonczone += 1
            postep = zakonczone / liczba_url
            try:
                status, data = future.result()
                if status == "sukces":
                    produkty, nazwa_pliku_url = data
                    for p in produkty:
                        pid = p["id"]
                        if pid and pid not in dane_po_id:
                            dane_po_id[pid] = p
                    sukcesy += 1
                    app_instance.update_status(f"Pobrano {zakonczone}/{liczba_url}: {nazwa_pliku_url}", postep)
                elif status == "blad_pobierania":
                    url, powod, nazwa_pliku_url = data
                    bledy_pobierania += 1
                    bledne_linki.append((url, powod))
                    app_instance.update_status(f"Błąd pobierania {zakonczone}/{liczba_url}: {nazwa_pliku_url}", postep)
                elif status == "blad_parsowania":
                    url, powod, nazwa_pliku_url = data
                    bledy_parsowania += 1
                    bledne_linki.append((url, powod))
                    app_instance.update_status(f"Błąd parsowania {zakonczone}/{liczba_url}: {nazwa_pliku_url}", postep)
            except Exception as e:
                bledy_parsowania += 1
                app_instance.update_status(f"Błąd krytyczny wątku {zakonczone}/{liczba_url}: {e}", postep)
                print(f"Błąd krytyczny w wątku: {e}")

    liczba_bledow = bledy_pobierania + bledy_parsowania
    podsumowanie_bledow = (
        f"Błędy pobierania: {bledy_pobierania}\n"
        f"Błędy parsowania: {bledy_parsowania}"
    )

    if not dane_po_id:
        app_instance.update_status(f"Brak danych. Błędy: {liczba_bledow}", 0)
        tresc = "Nie udało się pobrać ani sparsować żadnych produktów.\n\n" + podsumowanie_bledow
        if bledne_linki:
            tresc += "\n\nPrzykład: " + bledne_linki[0][0] + "\n" + bledne_linki[0][1]
        app_instance.after(0, lambda m=tresc: messagebox.showwarning("Brak danych", m))
        app_instance.reset_gui_after_delay()
        return

    # Nazwa pliku
    teraz = datetime.now().strftime("%d%m%y-%H%M%S")
    if len(nazwy) == 1:
        nazwa_pliku = f"{nazwy[0]}_joom_SE_{teraz}.xlsx"
    else:
        nazwa_pliku = f"{nazwy[0]}_and_{len(nazwy)-1}_more_joom_SE_{teraz}.xlsx"
    sciezka_wyjscia = os.path.join(sciezka_zapisu, nazwa_pliku)

    app_instance.update_status("Zapisywanie do szablonu Joom...", 0.98)
    sukces_zapisu, powod_zapisu, zapisane, obciete = zapisz_do_szablonu_joom(
        list(dane_po_id.values()), store_id, sciezka_wyjscia)

    if not sukces_zapisu:
        app_instance.update_status(f"Błąd zapisu: {powod_zapisu}", 1)
        app_instance.after(0, lambda m=powod_zapisu: messagebox.showerror(
            "Błąd zapisu", f"Nie udało się zapisać pliku Joom:\n{m}"))
        app_instance.reset_gui_after_delay()
        return

    app_instance.update_status(f"Zakończono. Produkty: {zapisane}, Błędy: {liczba_bledow}", 1)

    wiadomosc = (
        f"Pomyślnie pobrano {sukcesy} z {liczba_url} feedów.\n"
        f"Zapisano {zapisane} produktów do szablonu Joom:\n{os.path.abspath(sciezka_wyjscia)}\n\n"
        f"Store ID: {store_id}\n"
        f"{podsumowanie_bledow}"
    )
    if obciete:
        wiadomosc += f"\n\nUWAGA: pominięto {obciete} produktów (limit wierszy Excela)."
    app_instance.after(0, lambda m=wiadomosc: messagebox.showinfo("Sukces", m))
    app_instance.reset_gui_after_delay()


# ================================ GUI (Qt) ================================
class MainWindow(QMainWindow, ProgressMixin):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Konwerter feedów Joom do Excel (Qt)")
        self.resize(760, 620)

        self.bridge = Bridge()
        self.controller = BaseController(self.bridge)
        self._process_btn_text = "Pobierz i zbuduj plik Joom"

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.addWidget(self._build_names_group(), stretch=1)
        root.addWidget(self._build_config_group())
        root.addWidget(self._build_action_button())
        root.addWidget(self._build_progress_group())

        self._connect_bridge()

        # Kierujemy dialogi z logiki do Qt (wątkowo-bezpiecznie).
        global messagebox
        messagebox = MessageboxShim(self.bridge)

    # ---------- budowa UI ----------
    def _build_names_group(self) -> QGroupBox:
        g = QGroupBox("Nazwy feedów")
        lay = QVBoxLayout(g)
        lay.addWidget(QLabel(
            "Wprowadź nazwy dla generowania linków XML (każda w nowej linii).\n"
            "Link: " + FEED_URL_TEMPLATE.format(nazwa="{nazwa}")))
        self.names_edit = QPlainTextEdit()
        self.names_edit.setPlaceholderText("np. goodidea")
        lay.addWidget(self.names_edit, stretch=1)
        return g

    def _build_config_group(self) -> QGroupBox:
        g = QGroupBox("Konfiguracja")
        grid = QGridLayout(g)
        grid.setColumnStretch(1, 1)

        grid.addWidget(QLabel("Katalog zapisu pliku Excel:"), 0, 0)
        self.savedir_edit = QLineEdit(DOMYSLNA_SCIEZKA_ZAPISU)
        grid.addWidget(self.savedir_edit, 0, 1)
        btn_dir = QPushButton("Wybierz folder")
        btn_dir.clicked.connect(self._pick_savedir)
        grid.addWidget(btn_dir, 0, 2)

        grid.addWidget(QLabel("Store ID (Joom):"), 1, 0)
        self.store_edit = QLineEdit(DEFAULT_STORE_ID)
        self.store_edit.setPlaceholderText(DEFAULT_STORE_ID)
        grid.addWidget(self.store_edit, 1, 1, 1, 2)
        return g

    def _build_action_button(self) -> QPushButton:
        self.process_btn = QPushButton(self._process_btn_text)
        self.process_btn.clicked.connect(self._on_process_clicked)
        return self.process_btn

    def _build_progress_group(self) -> QGroupBox:
        g = QGroupBox("Postęp")
        lay = QVBoxLayout(g)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1000)
        self.status_label = QLabel("Gotowy.")
        lay.addWidget(self.progress_bar)
        lay.addWidget(self.status_label)
        return g

    # ---------- akcje ----------
    def _pick_savedir(self):
        d = QFileDialog.getExistingDirectory(
            self, "Wybierz folder zapisu",
            self.savedir_edit.text() or DOMYSLNA_SCIEZKA_ZAPISU)
        if d:
            self.savedir_edit.setText(d)

    def _on_process_clicked(self):
        if not self.process_btn.isEnabled():
            return
        nazwy = [n.strip() for n in self.names_edit.toPlainText().splitlines() if n.strip()]
        if not nazwy:
            QMessageBox.critical(self, "Błąd", "Musisz wprowadzić co najmniej jedną nazwę.")
            return

        sciezka_zapisu = self.savedir_edit.text().strip() or DOMYSLNA_SCIEZKA_ZAPISU
        store_id = self.store_edit.text().strip() or DEFAULT_STORE_ID

        self._begin_running("Przetwarzanie...")
        threading.Thread(
            target=przetworz_na_joom,
            args=(nazwy, sciezka_zapisu, store_id, self.controller),
            daemon=True,
        ).start()


def main():
    apply_scaling()
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
