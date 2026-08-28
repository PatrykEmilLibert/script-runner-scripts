from pathlib import Path
"""
Generator wsadu CDON z feedów XML sm-prods - PySide6

Produkuje plik .xlsx z KOMPLETEM kolumn, które potrafią wczytać:
  - CDON_API_MULTI.py        (import, POST /v2/articles/bulk)
  - CDON_API_update_MULTI.py (aktualizacja, PUT /v2/articles/bulk)

Plik .xlsx wrzuca się do tamtych narzędzi bezpośrednio - bez konwersji do CSV.

UWAGA: ten skrypt oraz oba powyższe tworzą jedną rodzinę - każdą zmianę
formatu wsadu trzeba nanieść na wszystkie trzy.

Wdrożenie do Centrum Zarządzania (wymaga uprawnień administratora):
  copy /Y "cdon_for_Magda.py" str((Path(__file__).parent / "scripts").resolve())
"""

import os
import re
import sys
import threading
import traceback
import xml.etree.ElementTree as ET

import openpyxl
import requests

from PySide6.QtCore import Qt, QObject, Signal
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QLabel, QPushButton, QLineEdit, QCheckBox, QComboBox, QSpinBox,
    QProgressBar, QPlainTextEdit, QFileDialog, QMessageBox
)

URL_GROUP_SUFFIXES = {
    "Normalna": {"se": "se", "dk": "dk", "fi": "fi"},
    "4%": {"se": "se_4", "dk": "dk_4", "fi": "fi_4"},
    "5,6%": {"se": "se_56", "dk": "dk_56", "fi": "fi_56"},
    "6,4%": {"se": "se_64", "dk": "dk_64", "fi": "fi_64"},
    "8%": {"se": "se_8", "dk": "dk_8", "fi": "fi_8"},
    "10%": {"se": "se_10", "dk": "dk_10", "fi": "fi_10"},
    "12%": {"se": "se_12", "dk": "dk_12", "fi": "fi_12"},
}

GROUP_ORDER = ["Normalna", "4%", "5,6%", "6,4%", "8%", "10%", "12%"]

# --- API sm-prods: lista zakazanych EAN-ów (forbidden_eans_paginated) ---
SMPRODS_TOKEN = "y7SeKeGSfVZtH9dCxwVULWTbcfWBrVq2WKcJssq8Pz8o5t3DFDpQ12BGRGc1S3fOJ2UC3tRMi29ChrseLAsl4GhHKR3Y9ALr9Zfq8pyeYtlExRas7rOfvRTrqKdEOJ8y"
FORBIDDEN_EANS_PAGINATED_URL = "https://api-sm-prods.sm-prods.com/forbidden_eans_paginated"

# Kraje UE (GPSR): jeśli producent jest spoza tej listy, CDON wymaga responsible_person.
EU_COUNTRY_CODES = {
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR",
    "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK",
    "SI", "ES", "SE"
}

SPEC_MARKETS = ["SE", "DK", "FI"]

STYLESHEET = """
QWidget { background-color: #ffffff; color: #1a1a1a; font-family: 'Segoe UI', Arial; font-size: 13px; }
QGroupBox { background-color: #fff5fa; border: 1px solid #ff69b4; border-radius: 8px;
            margin-top: 14px; padding: 14px 10px 10px 10px; font-weight: bold; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #c2188b; }
QPushButton { background-color: #ff69b4; color: white; border: none; border-radius: 6px;
              padding: 8px 14px; font-weight: bold; }
QPushButton:hover { background-color: #e754a6; }
QPushButton:disabled { background-color: #f2cde1; color: #ffffff; }
QPushButton#stopButton { background-color: #c2188b; }
QPushButton#stopButton:hover { background-color: #a31574; }
QLineEdit, QComboBox, QSpinBox, QPlainTextEdit { background-color: #ffffff; border: 1px solid #ff69b4;
              border-radius: 6px; padding: 5px; color: #1a1a1a; }
QComboBox QAbstractItemView { background-color: #ffffff; selection-background-color: #ffd6ea;
              color: #1a1a1a; border: 1px solid #ff69b4; }
QProgressBar { border: 1px solid #ff69b4; border-radius: 6px; text-align: center;
               background-color: #fff5fa; height: 18px; }
QProgressBar::chunk { background-color: #ff69b4; border-radius: 5px; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #ff69b4;
                       border-radius: 4px; background: #ffffff; }
QCheckBox::indicator:checked { background: #ff69b4; }
"""


# --- Pobieranie i parsowanie XML ---

def download_xml(url):
    """Downloads and parses an XML file from a URL."""
    try:
        response = requests.get(url, timeout=120)
        response.raise_for_status()
        root = ET.fromstring(response.content)
        return root
    except requests.exceptions.RequestException as e:
        raise ConnectionError(f"Błąd podczas pobierania pliku z URL {url}: {e}")
    except ET.ParseError as e:
        raise ValueError(f"Błąd podczas parsowania XML z URL {url}: {e}")
    except Exception as e:
        raise Exception(f"Nieoczekiwany błąd podczas przetwarzania URL {url}: {e}")


def pobierz_forbidden_eans():
    """
    Pobiera zbiór zakazanych EAN-ów z API sm-prods przez paginowany endpoint
    /forbidden_eans_paginated (?page=&page_size=, odpowiedź {page, page_size,
    total, total_pages, data:[{id, ean}]}).
    Zwraca (zbior_ean, blad).
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


def build_products_dict(root):
    """Builds a dictionary of products from the XML root for quick lookup."""
    products = {}
    if root is None:
        return products
    for offer in root.findall('o'):
        product_id = offer.get('id')
        if product_id:
            products[product_id] = offer
    return products


def get_attr(offer, attr_name, default=""):
    """Gets a specific attribute value from an offer's 'attrs' section."""
    attrs = offer.find('attrs')
    if attrs is not None:
        for a in attrs.findall('a'):
            if a.get('name') == attr_name:
                return (a.text or "").strip()
    return default


def get_category(offer):
    cat = offer.find('cat')
    return (cat.text or "").strip() if cat is not None else ""


def get_name(offer):
    name = offer.find('name')
    return (name.text or "").strip() if name is not None else ""


def get_desc(offer):
    desc = offer.find('desc')
    return (desc.text or "").strip() if desc is not None else ""


def get_main_image(offer):
    imgs = offer.find('imgs')
    if imgs is not None:
        main = imgs.find('main')
        if main is not None:
            return main.get('url', '')
    return ""


def get_extra_images(offer):
    """Gets up to 9 extra image URLs from an offer."""
    imgs = offer.find('imgs')
    urls = []
    if imgs is not None:
        all_i = imgs.findall('i')
        for i in all_i[:9]:
            url = i.get('url', '')
            if url:
                urls.append(url)
    return ";".join(urls)


def short(text, length):
    return text if len(text) <= length else text[:length]


def get_price(offer):
    return offer.get('price', "")


def get_stock(offer):
    return offer.get('stock', "")


def get_weight(offer):
    return offer.get('weight', "")


def strip_html_tags(text):
    if not isinstance(text, str):
        return text
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)


def get_brand(offer):
    """Determines the brand from attributes or product ID."""
    producent = get_attr(offer, "Producent")
    if producent:
        return producent.strip()
    prod_id = offer.get('id', '')
    if '_' in prod_id:
        return prod_id.split('_')[0]
    return prod_id


# --- Definicja kolumn wsadu ---

def build_headers(spec_slots):
    """
    Pełna lista kolumn czytanych przez CDON_API_MULTI.py i CDON_API_update_MULTI.py.
    spec_slots = liczba par name/value na rynek dla sekcji 'specifications'.
    """
    headers = [
        "sku", "weight", "brand", "gtin", "stock", "mainImage", "extraImages",
        "titleSe", "descriptionSe", "titleDk", "descriptionDk", "titleFi", "descriptionFi",
        "category",
        "originalPriceSe", "originalPriceDk", "originalPriceFi",
        "shippingCostSe", "shippingCostDk", "shippingCostFi",
        "deliveryTimeMinSe", "deliveryTimeMinDk", "deliveryTimeMinFi",
        "deliveryTimeMaxSe", "deliveryTimeMaxDk", "deliveryTimeMaxFi",
        "vatSe", "vatDk", "vatFi",
        "deliverySe", "deliveryDk", "deliveryFi",
        "shipped_from",
        # --- GPSR: producent i osoba odpowiedzialna w UE ---
        "manufacturer_name", "manufacturer_street_address", "manufacturer_city",
        "manufacturer_postal_code", "manufacturer_country", "manufacturer_website",
        "manufacturer_email",
        "responsible_person_name", "responsible_person_phone", "responsible_person_email",
    ]

    # --- specyfikacje techniczne per rynek ---
    for market in SPEC_MARKETS:
        headers.append(f"specification_{market}_group")
        for idx in range(1, spec_slots + 1):
            headers.append(f"specification_{market}_name_{idx}")
            headers.append(f"specification_{market}_value_{idx}")

    # awaryjna kolumna: gotowy JSON specyfikacji (używana, gdy kolumny wyżej są puste)
    headers.append("specifications_json")
    return headers


class GeneratorSettings:
    """Wartości wpisywane do każdego wiersza - z GUI."""

    def __init__(self):
        self.spec_slots = 3
        self.shipped_from = "EU"
        self.vat = {"Se": 25, "Dk": 25, "Fi": 25.5}
        self.delivery_time_min = 4
        self.delivery_time_max = 6
        self.delivery_type = "HomeDelivery"
        self.manufacturer = {
            "manufacturer_name": "",
            "manufacturer_street_address": "",
            "manufacturer_city": "",
            "manufacturer_postal_code": "",
            "manufacturer_country": "",
            "manufacturer_website": "",
            "manufacturer_email": "",
            "responsible_person_name": "",
            "responsible_person_phone": "",
            "responsible_person_email": "",
        }

    def manufacturer_warnings(self):
        """Zwraca listę ostrzeżeń o niekompletnych danych producenta (te same reguły co importer)."""
        m = self.manufacturer
        warnings = []
        filled = [v for v in m.values() if v.strip()]
        if not filled:
            return warnings

        if not m["manufacturer_name"].strip():
            warnings.append("Brak 'manufacturer_name' - importer pominie cały blok producenta.")
            return warnings

        missing = [key for key in (
            "manufacturer_street_address", "manufacturer_city",
            "manufacturer_postal_code", "manufacturer_country"
        ) if not m[key].strip()]
        if missing:
            warnings.append(
                "Niekompletny adres producenta (" + ", ".join(missing) +
                ") - importer pominie cały blok producenta.")
            return warnings

        country = m["manufacturer_country"].strip().upper()
        if len(country) != 2:
            warnings.append(f"'manufacturer_country' = '{country}' musi być 2-literowym kodem ISO.")
            return warnings

        if not m["responsible_person_name"].strip():
            if m["responsible_person_phone"].strip() or m["responsible_person_email"].strip():
                warnings.append("Podano telefon/e-mail osoby odpowiedzialnej, ale brak 'responsible_person_name'.")
            if country not in EU_COUNTRY_CODES:
                warnings.append(
                    f"Producent z kraju '{country}' (spoza UE) - CDON wymaga danych osoby odpowiedzialnej w UE.")
        return warnings


def build_row(headers, se_offer, dk_offer, fi_offer, settings):
    """Buduje jeden wiersz wsadu w kolejności zgodnej z 'headers'."""
    name = get_name(se_offer)
    desc = get_desc(se_offer)

    short_name = short(name, 135)
    processed_desc = strip_html_tags(desc) if len(desc) > 9500 else desc
    short_desc = short(processed_desc, 9500)

    values = {
        "sku": se_offer.get('id', ''),
        "weight": get_weight(se_offer),
        "brand": get_brand(se_offer),
        "gtin": get_attr(se_offer, "EAN"),
        "stock": get_stock(se_offer),
        "mainImage": get_main_image(se_offer),
        "extraImages": get_extra_images(se_offer),
        "titleSe": short_name, "descriptionSe": short_desc,
        "titleDk": short_name, "descriptionDk": short_desc,
        "titleFi": short_name, "descriptionFi": short_desc,
        "category": get_category(se_offer),
        "originalPriceSe": get_price(se_offer),
        "originalPriceDk": get_price(dk_offer) if dk_offer is not None else "",
        "originalPriceFi": get_price(fi_offer) if fi_offer is not None else "",
        "shippingCostSe": "0", "shippingCostDk": "0", "shippingCostFi": "0",
        "deliveryTimeMinSe": settings.delivery_time_min,
        "deliveryTimeMinDk": settings.delivery_time_min,
        "deliveryTimeMinFi": settings.delivery_time_min,
        "deliveryTimeMaxSe": settings.delivery_time_max,
        "deliveryTimeMaxDk": settings.delivery_time_max,
        "deliveryTimeMaxFi": settings.delivery_time_max,
        "vatSe": settings.vat["Se"], "vatDk": settings.vat["Dk"], "vatFi": settings.vat["Fi"],
        "deliverySe": settings.delivery_type,
        "deliveryDk": settings.delivery_type,
        "deliveryFi": settings.delivery_type,
        "shipped_from": settings.shipped_from,
    }
    values.update(settings.manufacturer)

    # kolumny specyfikacji i specifications_json zostają puste - do uzupełnienia w Excelu
    return [values.get(header, "") for header in headers]


# --- GUI (PySide6) ---

class WorkerSignals(QObject):
    log = Signal(str)
    progress = Signal(int, int)
    status = Signal(str)
    warning = Signal(str, str)
    finished = Signal(bool, str)


class FeedGeneratorWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Generator wsadu CDON - Qt")
        self.resize(880, 900)
        self.setStyleSheet(STYLESHEET)

        self.is_running = False
        self.stop_event = threading.Event()

        self.signals = WorkerSignals()
        self.signals.log.connect(self._append_log)
        self.signals.progress.connect(self._set_progress)
        self.signals.status.connect(self.status_label_set)
        self.signals.warning.connect(self._show_warning)
        self.signals.finished.connect(self._on_finished)

        self._create_widgets()

    # --- Budowa interfejsu ---

    def _create_widgets(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        header = QLabel("GENERATOR WSADU CDON")
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("font-size: 18px; font-weight: bold; color: #ff69b4; padding: 4px;")
        layout.addWidget(header)

        # --- Źródło feedów ---
        source_box = QGroupBox("Źródło feedów")
        source_grid = QGridLayout(source_box)
        source_grid.setColumnStretch(1, 1)

        source_grid.addWidget(QLabel("Przedrostek pliku XML:"), 0, 0)
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("np. armaservis")
        source_grid.addWidget(self.prefix_edit, 0, 1, 1, 3)

        source_grid.addWidget(QLabel("Grupy końcówek URL:"), 1, 0)
        groups_row = QGridLayout()
        self.group_checks = {}
        for index, group_name in enumerate(GROUP_ORDER):
            check = QCheckBox(group_name)
            check.setChecked(group_name == "Normalna")
            groups_row.addWidget(check, index // 4, index % 4)
            self.group_checks[group_name] = check
        source_grid.addLayout(groups_row, 1, 1, 1, 3)

        layout.addWidget(source_box)

        # --- Wartości domyślne wiersza ---
        defaults_box = QGroupBox("Wartości wpisywane do każdego wiersza")
        defaults_grid = QGridLayout(defaults_box)

        defaults_grid.addWidget(QLabel("shipped_from:"), 0, 0)
        self.shipped_from_combo = QComboBox()
        self.shipped_from_combo.addItems(["EU", "NON_EU"])
        defaults_grid.addWidget(self.shipped_from_combo, 0, 1)

        defaults_grid.addWidget(QLabel("Typ dostawy:"), 0, 2)
        self.delivery_combo = QComboBox()
        self.delivery_combo.addItems(["HomeDelivery", "ServicePoint", "Mailbox"])
        defaults_grid.addWidget(self.delivery_combo, 0, 3)

        defaults_grid.addWidget(QLabel("Czas dostawy min:"), 1, 0)
        self.delivery_min_spin = QSpinBox()
        self.delivery_min_spin.setRange(1, 9)
        self.delivery_min_spin.setValue(4)
        defaults_grid.addWidget(self.delivery_min_spin, 1, 1)

        defaults_grid.addWidget(QLabel("Czas dostawy max:"), 1, 2)
        self.delivery_max_spin = QSpinBox()
        self.delivery_max_spin.setRange(1, 9)
        self.delivery_max_spin.setValue(6)
        defaults_grid.addWidget(self.delivery_max_spin, 1, 3)

        defaults_grid.addWidget(QLabel("VAT SE / DK / FI (%):"), 2, 0)
        vat_row = QHBoxLayout()
        self.vat_se_edit = QLineEdit("25")
        self.vat_dk_edit = QLineEdit("25")
        self.vat_fi_edit = QLineEdit("25.5")
        for widget in (self.vat_se_edit, self.vat_dk_edit, self.vat_fi_edit):
            widget.setMaximumWidth(80)
            vat_row.addWidget(widget)
        vat_row.addStretch(1)
        defaults_grid.addLayout(vat_row, 2, 1, 1, 3)

        defaults_grid.addWidget(QLabel("Pary specyfikacji na rynek:"), 3, 0)
        self.spec_slots_spin = QSpinBox()
        self.spec_slots_spin.setRange(0, 20)
        self.spec_slots_spin.setValue(3)
        self.spec_slots_spin.setToolTip(
            "Ile pustych par kolumn specification_XX_name_N / specification_XX_value_N\n"
            "wygenerować dla każdego rynku (SE, DK, FI) do ręcznego uzupełnienia.")
        defaults_grid.addWidget(self.spec_slots_spin, 3, 1)

        layout.addWidget(defaults_box)

        # --- Producent (GPSR) ---
        manufacturer_box = QGroupBox("Producent i osoba odpowiedzialna w UE (GPSR) - opcjonalne")
        manufacturer_grid = QGridLayout(manufacturer_box)
        manufacturer_grid.setColumnStretch(1, 1)
        manufacturer_grid.setColumnStretch(3, 1)

        self.manufacturer_edits = {}
        manufacturer_fields = [
            ("manufacturer_name", "Nazwa producenta:", 0, 0),
            ("manufacturer_email", "E-mail producenta:", 0, 2),
            ("manufacturer_street_address", "Ulica i numer:", 1, 0),
            ("manufacturer_city", "Miasto:", 1, 2),
            ("manufacturer_postal_code", "Kod pocztowy:", 2, 0),
            ("manufacturer_country", "Kraj (ISO-2, np. CN):", 2, 2),
            ("manufacturer_website", "Strona www:", 3, 0),
            ("responsible_person_name", "Osoba odpowiedzialna - nazwa:", 4, 0),
            ("responsible_person_email", "Osoba odpowiedzialna - e-mail:", 4, 2),
            ("responsible_person_phone", "Osoba odpowiedzialna - telefon:", 5, 0),
        ]
        for key, label, row, col in manufacturer_fields:
            manufacturer_grid.addWidget(QLabel(label), row, col)
            edit = QLineEdit()
            manufacturer_grid.addWidget(edit, row, col + 1)
            self.manufacturer_edits[key] = edit

        hint = QLabel(
            "Zostaw puste, jeśli dane producenta uzupełnisz później w Excelu. "
            "Producent spoza UE wymaga wypełnienia osoby odpowiedzialnej.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #8a6b7c; font-weight: normal;")
        manufacturer_grid.addWidget(hint, 6, 0, 1, 4)

        layout.addWidget(manufacturer_box)

        # --- Plik wyjściowy ---
        output_box = QGroupBox("Plik wyjściowy")
        output_grid = QGridLayout(output_box)
        output_grid.setColumnStretch(1, 1)

        output_grid.addWidget(QLabel("Folder docelowy:"), 0, 0)
        self.output_edit = QLineEdit(os.path.join(os.path.expanduser('~'), 'Desktop'))
        output_grid.addWidget(self.output_edit, 0, 1)
        self.browse_btn = QPushButton("Przeglądaj...")
        self.browse_btn.clicked.connect(self.browse_output_directory)
        output_grid.addWidget(self.browse_btn, 0, 2)

        layout.addWidget(output_box)

        # --- Akcje ---
        self.generate_btn = QPushButton("Generuj plik Excel")
        self.generate_btn.setMinimumHeight(42)
        self.generate_btn.clicked.connect(self.run_processing)
        layout.addWidget(self.generate_btn)

        self.stop_btn = QPushButton("ZATRZYMAJ")
        self.stop_btn.setObjectName("stopButton")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_process)
        layout.addWidget(self.stop_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Gotowy")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        log_box = QGroupBox("Postęp")
        log_layout = QVBoxLayout(log_box)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(5000)
        log_layout.addWidget(self.log_view)
        layout.addWidget(log_box, 1)

    # --- Sloty GUI ---

    def browse_output_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Wybierz folder docelowy", self.output_edit.text())
        if directory:
            self.output_edit.setText(directory)

    def _append_log(self, message):
        self.log_view.appendPlainText(message)

    def _set_progress(self, current, total):
        self.progress_bar.setRange(0, max(total, 1))
        self.progress_bar.setValue(current)

    def status_label_set(self, text):
        self.status_label.setText(text)

    def _show_warning(self, title, message):
        QMessageBox.warning(self, title, message)

    def log(self, message):
        self.signals.log.emit(message)

    def collect_settings(self):
        """Zbiera ustawienia z GUI; zwraca (settings, blad)."""
        settings = GeneratorSettings()
        settings.spec_slots = self.spec_slots_spin.value()
        settings.shipped_from = self.shipped_from_combo.currentText()
        settings.delivery_type = self.delivery_combo.currentText()
        settings.delivery_time_min = self.delivery_min_spin.value()
        settings.delivery_time_max = self.delivery_max_spin.value()

        if settings.delivery_time_min > settings.delivery_time_max:
            return None, "Czas dostawy 'min' nie może być większy niż 'max'."

        for market, widget in (("Se", self.vat_se_edit), ("Dk", self.vat_dk_edit), ("Fi", self.vat_fi_edit)):
            raw = widget.text().strip().replace(',', '.')
            if not raw:
                return None, f"Podaj stawkę VAT dla rynku {market.upper()}."
            try:
                settings.vat[market] = float(raw)
            except ValueError:
                return None, f"Nieprawidłowa stawka VAT dla rynku {market.upper()}: '{widget.text()}'."

        for key, widget in self.manufacturer_edits.items():
            value = widget.text().strip()
            if key == "manufacturer_country":
                value = value.upper()
            settings.manufacturer[key] = value

        return settings, None

    def run_processing(self):
        prefix = self.prefix_edit.text().strip()
        output_dir = self.output_edit.text().strip()
        selected_groups = [name for name, check in self.group_checks.items() if check.isChecked()]

        if not prefix:
            QMessageBox.warning(self, "Brakujące dane", "Proszę wprowadzić przedrostek pliku.")
            return
        if not output_dir:
            QMessageBox.warning(self, "Brakujące dane", "Proszę podać folder docelowy.")
            return
        if not os.path.isdir(output_dir):
            QMessageBox.critical(self, "Błąd", f"Podany folder docelowy nie istnieje:\n{output_dir}")
            return
        if not selected_groups:
            QMessageBox.warning(self, "Brakujące dane", "Wybierz co najmniej jedną grupę końcówek URL.")
            return

        settings, error = self.collect_settings()
        if error:
            QMessageBox.warning(self, "Nieprawidłowe ustawienia", error)
            return

        warnings = settings.manufacturer_warnings()
        if warnings:
            answer = QMessageBox.question(
                self, "Dane producenta niekompletne",
                "\n\n".join(warnings) + "\n\nGenerować mimo to?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if answer != QMessageBox.Yes:
                return

        filename = f"{prefix}_output_feeds.xlsx"
        output_file_path = os.path.join(output_dir, filename)

        if os.path.exists(output_file_path):
            answer = QMessageBox.question(
                self, "Potwierdzenie",
                f"Plik '{filename}' już istnieje w wybranej lokalizacji.\n\nCzy chcesz go nadpisać?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if answer != QMessageBox.Yes:
                return

        # Kolejność grup zgodna z GROUP_ORDER, żeby wynik był powtarzalny
        ordered_groups = [name for name in GROUP_ORDER if name in selected_groups]

        self.is_running = True
        self.stop_event.clear()
        self.generate_btn.setEnabled(False)
        self.generate_btn.setText("Przetwarzanie...")
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.log_view.clear()

        threading.Thread(
            target=self.process_feeds,
            args=(ordered_groups, prefix, output_file_path, settings),
            daemon=True
        ).start()

    def stop_process(self):
        if self.is_running:
            self.stop_event.set()
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("ZATRZYMYWANIE...")
            self.log("[INFO] Wysłano STOP - przerwę po bieżącym feedzie.")

    # --- Wątek roboczy ---

    def process_feeds(self, selected_groups, prefix, output_file_path, settings):
        """Pobiera feedy i buduje plik Excel. Działa w osobnym wątku."""
        try:
            self.signals.status.emit("Pobieram listę zakazanych EAN-ów z sm-prods...")
            self.log("Pobieram listę zakazanych EAN-ów z sm-prods...")
            forbidden_eans, blad_forbidden = pobierz_forbidden_eans()
            if blad_forbidden:
                self.log(f"[WARN] {blad_forbidden}")
                self.signals.warning.emit(
                    "Brak listy forbidden EAN",
                    "Nie udało się pobrać listy zakazanych EAN-ów - "
                    f"generowanie przebiegnie BEZ tego filtra.\n\n{blad_forbidden}")
            else:
                self.log(f"Zakazanych EAN-ów na liście: {len(forbidden_eans)}")

            headers = build_headers(settings.spec_slots)
            self.log(f"Kolumn w wsadzie: {len(headers)} (w tym {settings.spec_slots} par specyfikacji na rynek)")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Dane"
            ws.append(headers)

            base_url = "https://sm-prods.com/feeds/"
            total_groups = len(selected_groups)
            skipped_forbidden = 0
            written_rows = 0

            for group_index, group_name in enumerate(selected_groups, start=1):
                if self.stop_event.is_set():
                    self.log(f"[INFO] Przerwano przed grupą '{group_name}'.")
                    break

                group_suffixes = URL_GROUP_SUFFIXES[group_name]
                se_url = f"{base_url}{prefix}_cdon_{group_suffixes['se']}.xml"
                dk_url = f"{base_url}{prefix}_cdon_{group_suffixes['dk']}.xml"
                fi_url = f"{base_url}{prefix}_cdon_{group_suffixes['fi']}.xml"

                self.signals.status.emit(f"[{group_index}/{total_groups}] Pobieranie feeda SE ({group_name})...")
                root_se = download_xml(se_url)
                if self.stop_event.is_set():
                    break
                self.signals.status.emit(f"[{group_index}/{total_groups}] Pobieranie feeda DK ({group_name})...")
                root_dk = download_xml(dk_url)
                if self.stop_event.is_set():
                    break
                self.signals.status.emit(f"[{group_index}/{total_groups}] Pobieranie feeda FI ({group_name})...")
                root_fi = download_xml(fi_url)

                self.signals.status.emit(f"[{group_index}/{total_groups}] Budowanie słowników produktów ({group_name})...")
                se_offers = build_products_dict(root_se)
                dk_offers = build_products_dict(root_dk)
                fi_offers = build_products_dict(root_fi)

                total = len(se_offers)
                count = 0
                self.log(f"[{group_index}/{total_groups}] Grupa '{group_name}': {total} produktów w feedzie SE")

                for pid, se_offer in se_offers.items():
                    if self.stop_event.is_set():
                        self.log(f"[INFO] Przerwano w trakcie grupy '{group_name}'.")
                        break

                    count += 1
                    if count % 200 == 0:
                        self.signals.status.emit(f"[{group_index}/{total_groups}] {group_name}: produkt {count}/{total}...")
                        self.signals.progress.emit(count, total)

                    gtin = get_attr(se_offer, "EAN")

                    # FILTR ZAKAZANYCH EAN-ów - pomiń produkt, jeśli EAN na liście.
                    if forbidden_eans and gtin and gtin.strip() in forbidden_eans:
                        skipped_forbidden += 1
                        continue

                    ws.append(build_row(
                        headers, se_offer, dk_offers.get(pid), fi_offers.get(pid), settings
                    ))
                    written_rows += 1

                self.signals.progress.emit(total, total)

            self.signals.status.emit("Zapisywanie pliku Excel...")
            self.log(f"Zapisywanie pliku: {output_file_path}")
            wb.save(output_file_path)

            summary = f"Plik '{os.path.basename(output_file_path)}' został zapisany."
            summary += f"\n\nWierszy: {written_rows}\nKolumn: {len(headers)}"
            if forbidden_eans:
                summary += f"\nOdsiano produktów z zakazanym EAN: {skipped_forbidden}"
            if self.stop_event.is_set():
                summary += "\n\n(Przerwano przez użytkownika - plik zawiera dane częściowe)"

            self.log("--- GOTOWE ---")
            self.signals.finished.emit(True, summary)

        except Exception as e:
            self.log(f"[BŁĄD] {e}")
            traceback.print_exc()
            self.signals.finished.emit(False, str(e))

    def _on_finished(self, ok, message):
        self.is_running = False
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("Generuj plik Excel")
        self.stop_btn.setEnabled(False)
        self.stop_btn.setText("ZATRZYMAJ")
        self.status_label.setText("Gotowy")
        if ok:
            QMessageBox.information(self, "Sukces", message)
        else:
            QMessageBox.critical(self, "Błąd", message)

    def closeEvent(self, event):
        if self.is_running:
            answer = QMessageBox.question(
                self, "Generowanie w toku",
                "Generowanie wciąż trwa. Zatrzymać i zamknąć?",
                QMessageBox.Yes | QMessageBox.No
            )
            if answer != QMessageBox.Yes:
                event.ignore()
                return
            self.stop_event.set()
        event.accept()


def main():
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    window = FeedGeneratorWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()