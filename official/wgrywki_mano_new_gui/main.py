import re
import pandas as pd
import json
import yaml
import os
import random
import sys
import threading
from pathlib import Path

import gui_qt_amno_common as common

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QGridLayout, QGroupBox,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
    QPushButton, QVBoxLayout, QWidget,
)

MAX_DPD_WEIGHT_KG = 10.0
ALLOWED_DPD_CARRIERS = {"DPD", "DPD2", "DPD3", "DPD4"}

# Wektoryzowany odpowiednik _has_allowed_dpd_carrier: dopasowuje przewoźnika jako
# cały token (granice = znaki spoza [A-Z0-9], tak jak split w wersji per-wiersz).
_ALLOWED_DPD_PATTERN = (
    r"(?<![A-Z0-9])(?:"
    + "|".join(re.escape(c) for c in sorted(ALLOWED_DPD_CARRIERS, key=len, reverse=True))
    + r")(?![A-Z0-9])"
)


def _clean_text(value):
    if value is None:
        return ""
    return " ".join(str(value).split())


def _to_float(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _normalize_aps_key(value):
    return _clean_text(value).lower()


def _has_allowed_dpd_carrier(value):
    if value is None:
        return False
    text = str(value).upper()
    if not text.strip():
        return False
    tokens = [t for t in re.split(r"[^A-Z0-9]+", text) if t]
    return any(t in ALLOWED_DPD_CARRIERS for t in tokens)


def _build_prefix_ean_key(product_sku, ean):
    sku_text = _clean_text(product_sku)
    ean_text = _clean_text(ean)
    if not sku_text or not ean_text:
        return ""
    if "_" not in sku_text:
        return ""
    prefix = sku_text.split("_", 1)[0]
    if not prefix:
        return ""
    return f"{prefix}_{ean_text}"


def _resolve_aps_row(aps_rows, product_sku, ean):
    prefix_ean_key = _build_prefix_ean_key(product_sku, ean)
    if prefix_ean_key:
        normalized = _normalize_aps_key(prefix_ean_key)
        if normalized in aps_rows:
            return aps_rows[normalized]
    sku_key = _normalize_aps_key(product_sku)
    if sku_key in aps_rows:
        return aps_rows[sku_key]
    return None


def _load_aps_rows(csv_path, needed_keys=None):
    """Load full APS CSV with columns: sku, aps_weight, FR, DE, IT, ES.

    Wektoryzowane wczytywanie pandas (chunki), semantycznie równoważne wersji
    per-wiersz opartej o csv.DictReader — liczy się z plikami rzędu milionów wierszy.

    needed_keys: opcjonalny zbiór znormalizowanych kluczy SKU, o które feed może
    kiedykolwiek zapytać (klucze prefix_ean oraz sku z kolumn id/EAN). Gdy podany,
    budujemy tylko pasujące rekordy — wynik odpytań _resolve_aps_row jest identyczny,
    bo każdy klucz, który mógłby zostać wyszukany, i tak zostaje zachowany.
    """
    rows_map = {}
    try:
        # Wykryj nazwy kolumn (bez rozróżniania wielkości liter) z samego nagłówka.
        header_df = pd.read_csv(csv_path, nrows=0, encoding="utf-8-sig")
        headers = {}
        for name in header_df.columns:
            normalized = _clean_text(name).lower()
            if normalized and normalized not in headers:
                headers[normalized] = name

        sku_col = headers.get("sku")
        aps_col = headers.get("aps_weight")

        if not sku_col or not aps_col:
            return {}, "Brak kolumn sku i/lub aps_weight w pliku APS CSV."

        market_cols = {country: headers.get(country.lower()) for country in ["FR", "DE", "IT", "ES"]}
        missing = [c for c, col in market_cols.items() if not col]
        if missing:
            return {}, "Brak kolumn rynkow w APS CSV: " + ", ".join(missing)

        usecols = [sku_col, aps_col] + [market_cols[c] for c in ["FR", "DE", "IT", "ES"]]
        reader = pd.read_csv(
            csv_path,
            usecols=usecols,
            dtype=str,
            encoding="utf-8-sig",
            na_filter=False,      # puste pola -> "" (jak row.get(...) w DictReader)
            chunksize=500_000,    # strumieniowo, żeby nie trzymać całego pliku w RAM
        )

        fr_col = market_cols["FR"]
        de_col = market_cols["DE"]
        it_col = market_cols["IT"]
        es_col = market_cols["ES"]

        # 16 możliwych kombinacji rynków jako współdzielone frozenset (FR=1,DE=2,IT=4,ES=8),
        # żeby nie alokować oddzielnego set() dla każdego z milionów wierszy.
        _bits = (("FR", 1), ("DE", 2), ("IT", 4), ("ES", 8))
        combos = [frozenset(c for c, b in _bits if idx & b) for idx in range(16)]

        for chunk in reader:
            # sku -> _normalize_aps_key: collapse whitespace + lower()
            sku_norm = chunk[sku_col].str.split().str.join(" ").str.lower()

            # Zawężenie do kluczy odpytywanych przez feed — kosztowne operacje
            # (waga, regex DPD) liczymy dopiero na dopasowanych wierszach.
            if needed_keys is not None:
                keep = sku_norm.isin(needed_keys)
                if not keep.any():
                    continue
                chunk = chunk[keep]
                sku_norm = sku_norm[keep]

            # aps_weight -> _to_float: strip, przecinek->kropka, float (NaN = odrzuć)
            aps_weight = pd.to_numeric(
                chunk[aps_col].str.strip().str.replace(",", ".", regex=False),
                errors="coerce",
            )
            # DPD w danym rynku jako cały token (odpowiednik _has_allowed_dpd_carrier),
            # zsumowane w maskę bitową rynków.
            bits = (
                chunk[fr_col].str.upper().str.contains(_ALLOWED_DPD_PATTERN, regex=True, na=False).to_numpy(dtype="int8")
                + chunk[de_col].str.upper().str.contains(_ALLOWED_DPD_PATTERN, regex=True, na=False).to_numpy(dtype="int8") * 2
                + chunk[it_col].str.upper().str.contains(_ALLOWED_DPD_PATTERN, regex=True, na=False).to_numpy(dtype="int8") * 4
                + chunk[es_col].str.upper().str.contains(_ALLOWED_DPD_PATTERN, regex=True, na=False).to_numpy(dtype="int8") * 8
            )

            for k, w, b in zip(sku_norm.tolist(), aps_weight.tolist(), bits.tolist()):
                if not k or w != w:  # k == "" lub w == NaN -> pomiń (jak w oryginale)
                    continue
                rows_map[k] = {"aps_weight": w, "dpd_markets": combos[b]}

        return rows_map, None
    except Exception as exc:
        return {}, str(exc)


# Hardcoded template mapping from CSV
TEMPLATE_MAPPING = {
    'sku': 'id',
    'ean': 'EAN',
    'sku_manufacturer': 'id',
    'brand': 'AUTRES',
    'manufacturer': '',
    'mm_category_id': 'cat_id',
    'merchant_category': '',
    'title': 'title',
    'description': 'desc',
    'product_url': '',
    'image_1': 'image0',
    'image_2': 'image1',
    'image_3': 'image2',
    'image_4': 'image3',
    'image_5': 'image4',
    'cross_sell_sku': '',
    'Sample_SKU': '',
    'manufacturer_pdf': '',
    'product_information_pdf': '',
    'repairability_index_pdf': '',
    'product_instructions_pdf': '',
    'safety_information_pdf': '',
    'refrigeration_devices_information_pdf': '',
    'eu_energy_efficiency_class_url': '',
    'unit_count': '',
    'unit_count_type': '',
    'ParentSKU': '',
    'parent_title': '',
    'length': 'losowa liczba 1-100',
    'length_unit': 'cm',
    'width': 'losowa liczba 1-100',
    'width_unit': 'cm',
    'height': 'losowa liczba 1-100',
    'height_unit': 'cm',
    'weight': 'weight',
    'weight_unit': 'kg',
    'volume': '100',
    'volume_unit': 'L',
    'light_output': '100',
    'light_output_unit': 'lm',
    'power': '100',
    'power_unit': 'W',
    'voltage': '230',
    'voltage_unit': 'V',
    'battery_life': '2',
    'battery_life_unit': 'year(s)',
    'amperage': '10',
    'amperage_unit': 'A',
    'rotational_speed': '250',
    'rotational_speed_unit': 'rpm',
    'max._flow_rate': '10',
    'max._flow_rate_unit': 'liter per minute',
    'noise_level': '65',
    'noise_level_unit': 'dB',
    'pressure': '100',
    'pressure_unit': 'PSI',
    'colour': 'Multicolour',
    'product_finish': 'Matt',
    'shape': 'Square',
    'ip_rating': 'IPXX',
    'cap_fitting': 'Integrated LED',
    'energy_efficiency_rating': 'C',
    'clothing_size': 'Unique size',
    'connection_configuration': 'No pipe thread',
    'pipe_thread_size': '3/4" (20x27)',
    'max_working_diameter': '125',
    'max_working_diameter_unit': 'mm',
    'disc_/_blade_diameter': '125',
    'disc_/_blade_diameter_unit': 'mm',
    'coverage': '100',
    'coverage_unit': 'm²',
    'light_colour': 'White',
    'center-to-center_distance': '250',
    'center-to-center_distance_unit': 'mm',
    'engine_capacity': '250',
    'engine_capacity_unit': 'cm³',
    'number_of_seats': '1',
    'number_of_seats_unit': 'seats',
    'power_source': 'Corded',
    'fixing_method': 'For wall',
    'maximum_load': '100',
    'maximum_load_unit': 'kg',
    'power_supply': 'Electricity',
    'style': 'Modern',
    'battery_technology': 'Li-po',
    'number_of_batteries': '1',
    'number_of_batteries_unit': 'batteries',
    'range': '100',
    'range_unit': 'm',
    'tilt_angle': '45',
    'tilt_angle_unit': 'degree',
    'pcs_per_pack': '1',
    'pcs_per_pack_unit': 'products',
    'warranty': '1',
    'warranty_unit': 'year(s)',
    'availability_of_spare_parts': '1',
    'availability_of_spare_parts_unit': 'year(s)',
    'volume_in_litres': '100',
    'volume_in_litres_unit': 'l',
    'collection_bin_capacity': '100',
    'collection_bin_capacity_unit': 'l',
    'intended_use': 'Indoor',
    'type_of_electrical_installation': 'Visible',
    'environmental_certification': 'Energy Star',
    'connection_port_type': 'Connection',
    'reparability_index': '1',
    'reparability_index_unit': 'part(s)',
    'finishing_appearance': 'Matt',
    'colour_name': 'Black',
    'colour_code': '#000000',
    'battery_capacity': '2',
    'battery_capacity_unit': 'Ah',
    'max._energy_efficiency_rating': 'A',
    'min._energy_efficiency_rating': 'D',
    'main_material': 'Two-material',
    'folding_thickness': '10',
    'folding_thickness_unit': 'mm',
    'folding_angle': '45',
    'folding_angle_unit': 'degree',
    'box_length': '10',
    'box_length_unit': 'cm',
    'box_width': '10',
    'box_width_unit': 'cm',
    'box_height': '10',
    'box_height_unit': 'cm',
    'screen_size_(in_inches)': '32',
    'screen_size_(in_inches)_unit': 'in',
    'product_price_vat_inc': '',
    'retail_price_vat_inc': '',
    'eco_participation': '',
    'min_quantity': '',
    'increment': '',
    'quantity_lower_bound_1': '',
    'quantity_price_1': '',
    'quantity_lower_bound_2': '',
    'quantity_price_2': '',
    'quantity_lower_bound_3': '',
    'quantity_price_3': '',
    'quantity': 'stock',
    'carrier_grid_1': '',
    'carrier_grid_2': '',
    'carrier_grid_3': '',
    'carrier_grid_4': '',
    'carrier_grid_5': '',
    'carrier_grid_6': '',
    'shipping_time_carrier_grid_1': '',
    'shipping_time_carrier_grid_2': '',
    'shipping_time_carrier_grid_3': '',
    'shipping_time_carrier_grid_4': '',
    'shipping_time_carrier_grid_5': '',
    'shipping_time_carrier_grid_6': '',
    'DisplayWeight': 'weight',
}

class ManoManoFeedGenerator:
    def __init__(self, taxonomy_file, log_callback=None, aps_csv_path=None):
        """
        Initialize the feed generator with taxonomy file (YAML format).
        
        Parameters:
        - taxonomy_file: path to YAML file with ManoMano taxonomy (simplified format)
        - log_callback: function to call for logging messages
        - aps_csv_path: optional path to full APS CSV (sku, aps_weight, FR, DE, IT, ES)
        """
        self.taxonomy_file = taxonomy_file
        self.taxonomy = {}
        self.template_mapping = TEMPLATE_MAPPING.copy()
        self.template_headers = list(TEMPLATE_MAPPING.keys())
        self.log_callback = log_callback or print
        self.aps_rows = {}
        # APS CSV wczytujemy leniwie w process_excel_to_csv — dopiero znając klucze
        # (id/EAN) z feedu, filtrujemy do rekordów, które mogą kiedykolwiek zostać
        # odpytane. Dzięki temu z pliku 6 mln wierszy budujemy słownik rzędu tysięcy.
        self.aps_csv_path = aps_csv_path

        self.load_taxonomy()
    
    def log(self, message):
        """Log message using callback or print."""
        self.log_callback(message)
    
    def load_taxonomy(self):
        """Load simplified taxonomy from YAML file."""
        self.log(f"Loading taxonomy from: {self.taxonomy_file}")
        try:
            with open(self.taxonomy_file, 'r', encoding='utf-8') as f:
                self.taxonomy = yaml.safe_load(f) or {}
            self.log(f"Loaded {len(self.taxonomy)} categories")
        except Exception as e:
            self.log(f"Error loading taxonomy: {e}")
            raise
    
    def get_required_attributes_for_category(self, category_id):
        """Get list of required attribute names for a category."""
        if category_id not in self.taxonomy:
            return []
        
        required_attrs = self.taxonomy[category_id]['required_attributes']
        return [attr['name'] for attr in required_attrs if attr.get('mandatory', False)]
    
    def generate_value(self, column_name, source_row=None, country_code=None, country_mapping=None):
        """
        Generate value based on template mapping and instructions.
        
        Parameters:
        - column_name: name of the column
        - source_row: row from source Excel file
        - country_code: country code (FR, DE, IT, ES) for dynamic mappings
        - country_mapping: dict with country-specific column mappings for title/description
        """
        # Handle country-specific mappings for title and description
        if country_mapping and column_name in country_mapping:
            instruction = country_mapping[column_name]
        else:
            # Get the mapping/instruction for this column
            instruction = self.template_mapping.get(column_name, '')
        
        # If instruction is empty, return empty
        if not instruction or instruction == '':
            return ''
        
        # Check if it's a source column reference (like 'id', 'ean', 'aps_weight', etc.)
        # These are column names from the Excel file
        if source_row is not None and instruction in source_row.index:
            value = source_row[instruction]
            if pd.notna(value):
                return value
        
        # Process instruction if no source value
        instruction_str = str(instruction).lower()
        
        # Handle random number generation
        if 'losowa liczba' in instruction_str or 'random' in instruction_str:
            # Extract range from instruction like "losowa liczba 1-100"
            match = re.search(r'(\d+)-(\d+)', instruction_str)
            if match:
                min_val = int(match.group(1))
                max_val = int(match.group(2))
                return random.randint(min_val, max_val)
        
        # Return the instruction as default value
        return instruction
    
    def process_excel_to_csv(self, excel_file_path, output_dir=None, progress_callback=None, allow_missing_cat_id=False, skip_weight_verification=False):
        """
        Process Excel file and create CSV files with category-aware required fields.
        
        Parameters:
        - excel_file_path: path to the input Excel file
        - output_dir: directory where CSV files will be saved (default: same as input file)
        - progress_callback: function to call with progress updates (0-100)
        - allow_missing_cat_id: include products without cat_id when True
        - skip_weight_verification: skip validation for weight < 40 when True
        """
        # Columns that must always be filled
        ALWAYS_REQUIRED_COLUMNS = {
            'sku', 'ean', 'sku_manufacturer', 'brand', 'mm_category_id',
            'title', 'description', 'image_1', 'image_2', 'image_3', 'image_4', 'image_5',
            'product_price_vat_inc', 'quantity', 'carrier_grid_1', 'shipping_time_carrier_grid_1',
            'DisplayWeight'
        }
        
        # Set base output directory
        if output_dir is None:
            output_dir = os.path.dirname(excel_file_path)

        # Always create subfolder based on input filename stem
        input_stem = Path(excel_file_path).stem
        output_dir = os.path.join(output_dir, input_stem)

        # Create output directory if it doesn't exist
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        self.log(f"Output directory: {output_dir}")
        
        # Read Excel file (calamine engine is much faster; fall back to default)
        self.log(f"\nReading Excel file: {excel_file_path}")
        try:
            try:
                df = pd.read_excel(excel_file_path, engine="calamine")
            except Exception:
                df = pd.read_excel(excel_file_path)
        except Exception as e:
            self.log(f"Error reading Excel file: {e}")
            raise
        
        # Check if required columns exist
        required_cols = ['avail']
        if not allow_missing_cat_id:
            required_cols.append('cat_id')
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            self.log(f"Warning: Missing columns in Excel: {missing_cols}")
        elif allow_missing_cat_id and 'cat_id' not in df.columns:
            self.log("Info: Missing column 'cat_id' - continuing because option to include products without cat_id is enabled")

        # Wczytaj APS CSV filtrując do kluczy, o które ten feed może zapytać
        # (klucze prefix_ean i sku z kolumn id/EAN). Identyczny wynik APS_BOX,
        # a z pliku milionów wierszy budujemy słownik rzędu tysięcy rekordów.
        if self.aps_csv_path:
            needed_keys = set()
            if 'id' in df.columns:
                ids = df['id'].tolist()
                eans = df['EAN'].tolist() if 'EAN' in df.columns else [None] * len(ids)
                for sku_v, ean_v in zip(ids, eans):
                    source_sku = str(sku_v) if pd.notna(sku_v) else ''
                    source_ean = str(ean_v) if pd.notna(ean_v) else ''
                    prefix_ean = _build_prefix_ean_key(source_sku, source_ean)
                    if prefix_ean:
                        needed_keys.add(_normalize_aps_key(prefix_ean))
                    sku_key = _normalize_aps_key(source_sku)
                    if sku_key:
                        needed_keys.add(sku_key)
            if not needed_keys:
                self.log("\nAPS CSV pominiety: brak kluczy id/EAN w feedzie (APS_BOX wylaczone)")
            else:
                self.log(f"\nWczytywanie APS CSV: {self.aps_csv_path}")
                rows, err = _load_aps_rows(self.aps_csv_path, needed_keys=needed_keys)
                if err:
                    self.log(f"Uwaga: nie mozna wczytac APS CSV: {err}")
                else:
                    self.aps_rows = rows
                    self.log(f"APS CSV wczytany: {len(self.aps_rows)} rekordow SKU dopasowanych do feedu")

        # Define the configuration for each output file
        configs = [
            {
                'name': 'DE',
                'country_code': 'DE',
                'price_col': 'price_manoDE',
                'carrier_grid_1': 'manomanode',
                'shipping_time': '2#4',
                'output_file': os.path.join(output_dir, 'output_DE.csv'),
                'title_col': 'DE',
                'description_col': 'DE2'
            },
            {
                'name': 'IT',
                'country_code': 'IT',
                'price_col': 'price_manoIT',
                'carrier_grid_1': 'manomanoit',
                'shipping_time': '4#6',
                'output_file': os.path.join(output_dir, 'output_IT.csv'),
                'title_col': 'IT',
                'description_col': 'IT2'
            },
            {
                'name': 'ES',
                'country_code': 'ES',
                'price_col': 'price_manoES',
                'carrier_grid_1': 'manomanoes',
                'shipping_time': '4#6',
                'output_file': os.path.join(output_dir, 'output_ES.csv'),
                'title_col': 'ES',
                'description_col': 'ES2'
            },
            {
                'name': 'FR',
                'country_code': 'FR',
                'price_col': 'price_manoFR',
                'carrier_grid_1': 'manomanofr',
                'shipping_time': '4#6',
                'output_file': os.path.join(output_dir, 'output_FR.csv'),
                'title_col': 'FR',
                'description_col': 'FR2'
            },
            {
                'name': 'DE_extra',
                'country_code': 'DE',
                'price_col': 'price_manoDE_extra',
                'carrier_grid_1': 'manomanode_APS',
                'shipping_time': '2#4',
                'output_file': os.path.join(output_dir, 'output_DE_extra.csv'),
                'title_col': 'DE',
                'description_col': 'DE2'
            },
            {
                'name': 'IT_extra',
                'country_code': 'IT',
                'price_col': 'price_manoIT_extra',
                'carrier_grid_1': 'manomanoit_APS',
                'shipping_time': '4#6',
                'output_file': os.path.join(output_dir, 'output_IT_extra.csv'),
                'title_col': 'IT',
                'description_col': 'IT2'
            },
            {
                'name': 'ES_extra',
                'country_code': 'ES',
                'price_col': 'price_manoES_extra',
                'carrier_grid_1': 'manomanoes_APS',
                'shipping_time': '4#6',
                'output_file': os.path.join(output_dir, 'output_ES_extra.csv'),
                'title_col': 'ES',
                'description_col': 'ES2'
            },
            {
                'name': 'FR_extra',
                'country_code': 'FR',
                'price_col': 'price_manoFR_extra',
                'carrier_grid_1': 'manomanofr_APS',
                'shipping_time': '4#6',
                'output_file': os.path.join(output_dir, 'output_FR_extra.csv'),
                'title_col': 'FR',
                'description_col': 'FR2'
            }
        ]
        
        total_configs = len(configs)
        
        # --- Precompute once (independent of row/config) to speed up the hot loop ---
        df_columns = set(df.columns)
        base_empty_row = {col: '' for col in self.template_headers}
        always_only_fill = [c for c in self.template_headers if c in ALWAYS_REQUIRED_COLUMNS]

        # Classify each template column so per-cell work avoids regex/.lower() in the loop
        col_class = {}
        for col in self.template_headers:
            instr = self.template_mapping.get(col, '')
            if not instr or instr == '':
                col_class[col] = ('empty',)
                continue
            instr_low = str(instr).lower()
            m = re.search(r'(\d+)-(\d+)', instr_low)
            is_rand = ('losowa liczba' in instr_low or 'random' in instr_low) and m is not None
            if instr in df_columns:
                if is_rand:
                    col_class[col] = ('src_rand', instr, int(m.group(1)), int(m.group(2)))
                else:
                    col_class[col] = ('src_const', instr, instr)
            elif is_rand:
                col_class[col] = ('rand', int(m.group(1)), int(m.group(2)))
            else:
                col_class[col] = ('const', instr)

        def _cell(kind, row):
            t = kind[0]
            if t == 'const':
                return kind[1]
            if t == 'src_const':
                v = row[kind[1]]
                return v if pd.notna(v) else kind[2]
            if t == 'rand':
                return random.randint(kind[1], kind[2])
            if t == 'src_rand':
                v = row[kind[1]]
                return v if pd.notna(v) else random.randint(kind[2], kind[3])
            return ''  # empty

        image_src = {c: self.template_mapping.get(c, '')
                     for c in self.template_headers if c.startswith('image_')}
        has_weight = 'weight' in df_columns
        has_id = 'id' in df_columns
        has_ean = 'EAN' in df_columns
        req_cache = {}      # category_id (int) -> set of required attr names
        fill_cache = {}     # category_id (int) -> list of columns to fill

        # Process each configuration
        for config_idx, config in enumerate(configs):
            self.log(f"\nProcessing {config['name']} file...")

            # Filter data: avail = 1 and price column is not empty
            filter_condition = df['avail'] == 1

            if config['price_col'] in df.columns:
                filter_condition = filter_condition & df[config['price_col']].notna() & (df[config['price_col']] != '')
            else:
                self.log(f"Warning: Price column {config['price_col']} not found in Excel")
                continue

            filtered_df = df[filter_condition]

            self.log(f"Found {len(filtered_df)} products with price in {config['price_col']}")

            # Per-config constants
            price_col = config['price_col']
            carrier_1 = config['carrier_grid_1']
            shipping_1 = config['shipping_time']
            title_col = config.get('title_col')
            desc_col = config.get('description_col')
            has_title_col = title_col in df_columns
            has_desc_col = desc_col in df_columns
            is_extra = config['name'].endswith('_extra')
            country_code = config['country_code']

            config_base = base_empty_row.copy()
            config_base['carrier_grid_1'] = carrier_1
            config_base['shipping_time_carrier_grid_1'] = shipping_1

            # Create output dataframe
            output_data = []
            skipped_count = 0

            for row in filtered_df.to_dict('records'):
                # Get category ID and required attributes
                category_id = row.get('cat_id')
                required_attr_names = None

                if category_id and pd.notna(category_id):
                    try:
                        category_id = int(category_id)
                        required_attr_names = req_cache.get(category_id)
                        if required_attr_names is None:
                            required_attr_names = set(self.get_required_attributes_for_category(category_id))
                            req_cache[category_id] = required_attr_names
                    except (ValueError, TypeError):
                        pass

                # Get weight value for validation
                weight = None
                if has_weight:
                    try:
                        weight = float(row['weight'])
                    except (ValueError, TypeError):
                        pass

                # Get title and description values for this language
                title_value = row.get(title_col) if has_title_col else None
                description_value = row.get(desc_col) if has_desc_col else None

                # Validation: Check if product meets all requirements
                # 1. Must have category (optional)
                if (not allow_missing_cat_id) and (not category_id or pd.isna(category_id)):
                    skipped_count += 1
                    continue

                # 2. Must have weight lower than 40 (optional)
                if not skip_weight_verification:
                    if weight is None or pd.isna(weight) or weight >= 40:
                        skipped_count += 1
                        continue

                # 3. Must have title in the language
                if pd.isna(title_value) or title_value == '':
                    skipped_count += 1
                    continue

                # 4. Must have description in the language
                if pd.isna(description_value) or description_value == '':
                    skipped_count += 1
                    continue

                # Determine which columns to fill (cached per category)
                if required_attr_names:
                    fill_cols = fill_cache.get(category_id)
                    if fill_cols is None:
                        fill_cols = [c for c in self.template_headers
                                     if c in ALWAYS_REQUIRED_COLUMNS or c in required_attr_names]
                        fill_cache[category_id] = fill_cols
                else:
                    fill_cols = always_only_fill

                # Fill columns (start from an all-empty template copy)
                output_row = config_base.copy()
                for col in fill_cols:
                    if col == 'title':
                        output_row[col] = title_value
                    elif col == 'description':
                        output_row[col] = description_value
                    elif col == 'product_price_vat_inc':
                        output_row[col] = row[price_col]
                    elif col == 'mm_category_id':
                        output_row[col] = category_id if (category_id and not pd.isna(category_id)) else ''
                    elif col == 'carrier_grid_1' or col == 'shipping_time_carrier_grid_1':
                        continue  # already set in config_base
                    elif col in image_src:
                        src = image_src[col]
                        if src and src in df_columns:
                            v = row[src]
                            output_row[col] = v if pd.notna(v) else ''
                        else:
                            output_row[col] = ''
                    else:
                        output_row[col] = _cell(col_class[col], row)

                # APS_BOX: for _extra configs set carrier_grid_2 when DPD conditions are met
                if is_extra and self.aps_rows:
                    source_sku = str(row['id']) if has_id and pd.notna(row['id']) else ''
                    source_ean = str(row['EAN']) if has_ean and pd.notna(row['EAN']) else ''
                    aps_row_data = _resolve_aps_row(self.aps_rows, source_sku, source_ean)
                    if aps_row_data:
                        aps_w = _to_float(aps_row_data.get('aps_weight'))
                        dpd_markets = aps_row_data.get('dpd_markets', set())
                        if aps_w is not None and aps_w <= MAX_DPD_WEIGHT_KG and country_code in dpd_markets:
                            box_carrier = f"manomano{country_code.lower()}_APS_BOX"
                            output_row['carrier_grid_2'] = box_carrier
                            output_row['shipping_time_carrier_grid_2'] = shipping_1

                output_data.append(output_row)
            
            # Create DataFrame from output data
            output_df = pd.DataFrame(output_data, columns=self.template_headers)
            
            # Save to CSV with semicolon separator
            output_df.to_csv(config['output_file'], index=False, sep=';', encoding='utf-8')
            self.log(f"Saved: {config['output_file']}")
            self.log(f"Total rows: {len(output_df)}")
            if skipped_count > 0:
                if allow_missing_cat_id:
                    if skip_weight_verification:
                        self.log(f"Skipped: {skipped_count} products (title or description)")
                    else:
                        self.log(f"Skipped: {skipped_count} products (weight ≥ 40, title, or description)")
                else:
                    if skip_weight_verification:
                        self.log(f"Skipped: {skipped_count} products (missing category, title, or description)")
                    else:
                        self.log(f"Skipped: {skipped_count} products (missing category, weight ≥ 40, title, or description)")

            # Update progress
            if progress_callback:
                progress = int((config_idx + 1) / total_configs * 100)
                progress_callback(progress)


class Worker(threading.Thread):
    def __init__(self, taxonomy_file, excel_file, output_dir, allow_missing_cat_id, skip_weight_verification, aps_csv_path, log_callback, progress_callback, finished_callback, failed_callback):
        super().__init__(daemon=True)
        self.taxonomy_file = taxonomy_file
        self.excel_file = excel_file
        self.output_dir = output_dir
        self.allow_missing_cat_id = allow_missing_cat_id
        self.skip_weight_verification = skip_weight_verification
        self.aps_csv_path = aps_csv_path
        self.log_callback = log_callback
        self.progress_callback = progress_callback
        self.finished_callback = finished_callback
        self.failed_callback = failed_callback

    def run(self):
        try:
            self.log_callback("Initializing generator...")
            generator = ManoManoFeedGenerator(
                self.taxonomy_file,
                log_callback=self.log_callback,
                aps_csv_path=self.aps_csv_path if self.aps_csv_path else None,
            )

            generator.process_excel_to_csv(
                self.excel_file,
                output_dir=self.output_dir,
                progress_callback=self.progress_callback,
                allow_missing_cat_id=self.allow_missing_cat_id,
                skip_weight_verification=self.skip_weight_verification
            )

            self.log_callback("\n" + "=" * 50)
            self.log_callback("Done! CSV files have been created.")
            self.log_callback("=" * 50)
            self.finished_callback()
        except Exception as e:
            self.failed_callback(str(e))


class Bridge(QObject):
    """Sygnały do bezpiecznego przekazywania zdarzeń z wątku workera do GUI."""
    log_sig = Signal(str)        # linia logu
    progress_sig = Signal(int)   # postęp 0-100
    finished_sig = Signal()      # zakończono pomyślnie
    failed_sig = Signal(str)     # błąd (treść)


class ManoManoGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ManoMano Feed Generator")
        self.resize(900, 760)

        self.bridge = Bridge()
        self.worker = None
        self._process_btn_text = "Generate CSV Files"

        # Domyślne ścieżki (Pulpit bieżącego użytkownika)
        user_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.default_taxonomy = os.path.join(user_desktop, "manomano_taxonomy.yml")

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        title = QLabel("ManoMano Feed Generator")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(
            f"font-size: 22px; font-weight: bold; color: {common.ACCENT_HOVER}; padding: 4px;")
        root.addWidget(title)

        root.addWidget(self._build_files_group())
        root.addWidget(self._build_options_group())
        root.addWidget(self.process_btn_widget())
        root.addWidget(self._build_progress_group())
        root.addWidget(self._build_log_group(), stretch=1)

        self._connect_bridge()

    # ---------- budowa UI ----------
    def _file_row(self, grid, row, label, browse_slot, default_text=""):
        grid.addWidget(QLabel(label), row, 0)
        edit = QLineEdit(default_text)
        grid.addWidget(edit, row, 1)
        btn = QPushButton("Browse")
        btn.clicked.connect(browse_slot)
        grid.addWidget(btn, row, 2)
        return edit

    def _build_files_group(self) -> QGroupBox:
        g = QGroupBox("Pliki")
        grid = QGridLayout(g)
        grid.setColumnStretch(1, 1)

        self.excel_edit = self._file_row(grid, 0, "Excel File:", self.browse_excel)
        taxonomy_default = self.default_taxonomy if os.path.exists(self.default_taxonomy) else ""
        self.taxonomy_edit = self._file_row(grid, 1, "Taxonomy YAML:", self.browse_taxonomy, taxonomy_default)
        self.output_edit = self._file_row(grid, 2, "Output Directory:", self.browse_output)
        self.aps_csv_edit = self._file_row(
            grid, 3, "APS CSV (pelny: sku, aps_weight, FR/DE/IT/ES):", self.browse_aps_csv)

        info = QLabel(
            "(Leave Output Directory empty to use Excel file location; "
            "output goes to a subfolder named after input file)")
        info.setStyleSheet(f"color: {common.MUTED}; font-size: 11px;")
        info.setWordWrap(True)
        grid.addWidget(info, 4, 0, 1, 3)
        return g

    def _build_options_group(self) -> QGroupBox:
        g = QGroupBox("Opcje")
        lay = QVBoxLayout(g)
        self.allow_missing_cat_id_checkbox = QCheckBox("Include products without cat_id")
        self.skip_weight_verification_checkbox = QCheckBox("Skip weryfikacji wagi (< 40)")
        lay.addWidget(self.allow_missing_cat_id_checkbox)
        lay.addWidget(self.skip_weight_verification_checkbox)
        return g

    def process_btn_widget(self) -> QPushButton:
        self.process_btn = QPushButton(self._process_btn_text)
        self.process_btn.clicked.connect(self.process_files)
        return self.process_btn

    def _build_progress_group(self) -> QGroupBox:
        g = QGroupBox("Postęp")
        lay = QVBoxLayout(g)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        lay.addWidget(self.progress_bar)
        return g

    def _build_log_group(self) -> QGroupBox:
        g = QGroupBox("Log")
        lay = QVBoxLayout(g)
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        lay.addWidget(self.log_text)
        return g

    def _connect_bridge(self):
        self.bridge.log_sig.connect(self._append_log)
        self.bridge.progress_sig.connect(self.progress_bar.setValue)
        self.bridge.finished_sig.connect(self.on_finished)
        self.bridge.failed_sig.connect(self.on_failed)

    # ---------- akcje wyboru plików ----------
    def browse_excel(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "",
            "Excel files (*.xlsx *.xls);;All files (*.*)")
        if filename:
            self.excel_edit.setText(filename)
            if not self.output_edit.text().strip():
                self.output_edit.setText(os.path.dirname(filename))

    def browse_taxonomy(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Taxonomy YAML File", "",
            "YAML files (*.yml *.yaml);;All files (*.*)")
        if filename:
            self.taxonomy_edit.setText(filename)

    def browse_output(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if directory:
            self.output_edit.setText(directory)

    def browse_aps_csv(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Wybierz APS CSV (pelny z kolumnami FR/DE/IT/ES)", "",
            "CSV files (*.csv);;All files (*.*)")
        if filename:
            self.aps_csv_edit.setText(filename)

    # ---------- sloty GUI (główny wątek) ----------
    def _append_log(self, message: str):
        self.log_text.appendPlainText(message)

    # ---------- callbacki wołane z wątku workera ----------
    def log(self, message):
        self.bridge.log_sig.emit(str(message))

    def update_progress(self, value):
        self.bridge.progress_sig.emit(int(value))

    def process_files(self):
        excel_file = self.excel_edit.text().strip()
        taxonomy_file = self.taxonomy_edit.text().strip()
        output_dir = self.output_edit.text().strip() or None
        allow_missing_cat_id = self.allow_missing_cat_id_checkbox.isChecked()
        skip_weight_verification = self.skip_weight_verification_checkbox.isChecked()
        aps_csv_path = self.aps_csv_edit.text().strip() or None

        if not excel_file:
            QMessageBox.critical(self, "Error", "Please select an Excel file")
            return
        if not taxonomy_file:
            QMessageBox.critical(self, "Error", "Please select a taxonomy YAML file")
            return
        if not os.path.exists(excel_file):
            QMessageBox.critical(self, "Error", f"Excel file not found:\n{excel_file}")
            return
        if not os.path.exists(taxonomy_file):
            QMessageBox.critical(self, "Error", f"Taxonomy file not found:\n{taxonomy_file}")
            return
        if aps_csv_path and not os.path.exists(aps_csv_path):
            QMessageBox.critical(self, "Error", f"APS CSV file not found:\n{aps_csv_path}")
            return

        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.process_btn.setEnabled(False)
        self.process_btn.setText("Przetwarzanie...")
        self.log(f"Include products without cat_id: {'ON' if allow_missing_cat_id else 'OFF'}")
        self.log(f"Skip weryfikacji wagi: {'ON' if skip_weight_verification else 'OFF'}")
        self.log(f"APS CSV: {aps_csv_path if aps_csv_path else 'nie wybrany (APS_BOX wylaczone)'}")

        self.worker = Worker(
            taxonomy_file, excel_file, output_dir, allow_missing_cat_id, skip_weight_verification, aps_csv_path,
            log_callback=self.log,
            progress_callback=self.update_progress,
            finished_callback=self.bridge.finished_sig.emit,
            failed_callback=self.bridge.failed_sig.emit,
        )
        self.worker.start()

    def on_finished(self):
        self.process_btn.setEnabled(True)
        self.process_btn.setText(self._process_btn_text)
        QMessageBox.information(self, "Success", "CSV files have been generated successfully!")

    def on_failed(self, error_msg):
        self.process_btn.setEnabled(True)
        self.process_btn.setText(self._process_btn_text)
        self.log(f"Error: {error_msg}")
        QMessageBox.critical(self, "Error", error_msg)


def main():
    """Main function to run the GUI."""
    common.apply_scaling()
    app = QApplication(sys.argv)
    app.setStyleSheet(common.QSS)
    win = ManoManoGUI()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
