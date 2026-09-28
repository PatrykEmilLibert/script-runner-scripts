"""Analizator SKU: Partnerzy i Grupy (CDON) — samodzielne okno PySide6.

Liczy aktywne SKU (real_status == "aktywne") per partner (prefiks SKU przed "_")
w grupach SHUMEE / GREATSTORE / EXTRASTORE. Kolumna SUMA liczy unikalne SKU
partnera (ten sam SKU w kilku grupach liczony raz).

Układ i styl jak w dawnej wersji Streamlit, kolorystyka przestawiona na różową.
"""
import datetime
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, Qt, QThread, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableView,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

# --- STAŁE ---
GROUPS = ["SHUMEE", "GREATSTORE", "EXTRASTORE"]
ACTIVE_REAL_STATUS_VALUES = {"aktywne"}
SUM_COL = "SUMA (Unikalne)"
TOTAL_LABEL = "SUMA CAŁKOWITA"

# --- PALETA (układ Streamlit, niebieskie akcenty zamienione na róż) ---
APP_BG = "#fdf2f8"          # było #f3f4f6
TEXT = "#1f2937"
HEADING = "#9d174d"         # było #1e3a8a
METRIC = "#111827"
CARD_BG = "#ffffff"
UPLOAD_BORDER = "#f9a8d4"   # było #cbd5e1 (przerywana ramka uploadu)
BTN_FROM = "#ec4899"        # gradient przycisku, było #2563eb -> #1e40af
BTN_TO = "#be185d"
BTN_FROM_HOVER = "#f472b6"
BTN_TO_HOVER = "#db2777"
INFO_BG = "#fce7f3"         # st.info
INFO_TEXT = "#9d174d"
SUCCESS_BG = "#fbcfe8"      # st.success
SUCCESS_TEXT = "#831843"
ERROR_BG = "#ffe4e6"        # st.error / st.warning
ERROR_TEXT = "#9f1239"
BORDER = "#f5d0e6"
MUTED = "#6b7280"
HEADER_BG = "#fce7f3"
ALT_ROW = "#fff7fb"
TOTAL_BG = "#fbcfe8"

QSS = f"""
QWidget {{ background-color: {APP_BG}; color: {TEXT}; font-family: "Segoe UI"; font-size: 14px; }}
QScrollArea {{ border: none; }}
QLabel {{ background: transparent; }}
QLabel#h1 {{ font-size: 30px; font-weight: 700; color: {HEADING}; }}
QLabel#h2 {{ font-size: 24px; font-weight: 700; color: {HEADING}; }}
QLabel#h3 {{ font-size: 19px; font-weight: 700; color: {HEADING}; }}
QLabel#muted {{ color: {MUTED}; }}
QLabel#info {{
    background-color: {INFO_BG}; color: {INFO_TEXT};
    border-radius: 8px; padding: 14px 16px;
}}
QLabel#success {{
    background-color: {SUCCESS_BG}; color: {SUCCESS_TEXT};
    border-radius: 8px; padding: 14px 16px;
}}
QLabel#error {{
    background-color: {ERROR_BG}; color: {ERROR_TEXT};
    border-radius: 8px; padding: 14px 16px;
}}
QFrame#uploader {{
    background-color: {CARD_BG};
    border: 1px dashed {UPLOAD_BORDER};
    border-radius: 12px;
}}
QFrame#uploader[drag="true"] {{ border: 2px dashed {BTN_FROM}; }}
QFrame#card {{
    background-color: {CARD_BG};
    border: 1px solid {BORDER};
    border-radius: 8px;
}}
QFrame#divider {{ background-color: {BORDER}; max-height: 1px; min-height: 1px; border: none; }}
QLabel#metricLabel {{ font-size: 14px; color: {TEXT}; }}
QLabel#metricValue {{ font-size: 28px; color: {METRIC}; }}
QListWidget {{
    background-color: {CARD_BG}; border: none; color: {TEXT};
    selection-background-color: {INFO_BG}; selection-color: {HEADING};
}}
QPushButton {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {BTN_FROM}, stop:1 {BTN_TO});
    border: none; color: white; font-weight: bold;
    padding: 8px 16px; border-radius: 8px;
}}
QPushButton:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {BTN_FROM_HOVER}, stop:1 {BTN_TO_HOVER});
}}
QPushButton:disabled {{ background: {BORDER}; color: white; }}
QPushButton#browse {{
    background: {CARD_BG}; color: {TEXT}; font-weight: normal;
    border: 1px solid {BORDER}; padding: 6px 12px;
}}
QPushButton#browse:hover {{ border: 1px solid {BTN_FROM}; color: {BTN_FROM}; }}
QPushButton#link {{
    background: transparent; color: {MUTED}; font-weight: normal; padding: 2px 6px;
}}
QPushButton#link:hover {{ color: {BTN_TO}; }}
QToolButton#expander {{
    background-color: {CARD_BG}; border: 1px solid {BORDER}; border-radius: 8px;
    padding: 10px 12px; text-align: left; color: {TEXT};
}}
QToolButton#expander:hover {{ color: {BTN_TO}; }}
QTableView {{
    background-color: {CARD_BG};
    alternate-background-color: {ALT_ROW};
    border: 1px solid {BORDER};
    border-radius: 8px;
    gridline-color: {BORDER};
    selection-background-color: {INFO_BG};
    selection-color: {HEADING};
}}
QHeaderView::section {{
    background-color: {HEADER_BG}; color: {MUTED};
    border: none; border-right: 1px solid {BORDER}; border-bottom: 1px solid {BORDER};
    padding: 6px;
}}
QTableCornerButton::section {{ background-color: {HEADER_BG}; border: none; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {UPLOAD_BORDER}; border-radius: 5px; min-height: 24px; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 0; }}
QScrollBar::handle:horizontal {{ background: {UPLOAD_BORDER}; border-radius: 5px; min-width: 24px; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
QStatusBar {{ color: {MUTED}; }}
"""


# --- GŁÓWNA LOGIKA ---
def analyze(uploads):
    """uploads: {GRUPA: [ścieżki CSV]} -> (df_final | None, total_processed, liczba_partnerów, błędy)."""
    # Struktura danych: partner_group_data[GRUPA][PARTNER] = {zbiór_sku}
    partner_group_data = {g: defaultdict(set) for g in GROUPS}
    errors = []
    all_partners = set()
    total_skus_processed_count = 0

    for group in GROUPS:
        files = uploads.get(group) or []
        if not files:
            continue

        # Deduplikacja per grupa
        unique_group_skus = set()

        for path in files:
            name = Path(path).name
            try:
                # Nowy format pliku: separator ';' i kolumna aktywności 'real_status'
                df = pd.read_csv(path, sep=";", on_bad_lines="skip", dtype=str)
                df.columns = [str(col).strip().lower() for col in df.columns]

                if "sku" in df.columns and "real_status" in df.columns:
                    active_mask = (
                        df["real_status"]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .isin(ACTIVE_REAL_STATUS_VALUES)
                    )
                    active_skus = df.loc[active_mask, "sku"].dropna().astype(str).str.strip()
                    unique_group_skus.update(sku for sku in active_skus if sku)
                else:
                    errors.append(f"❌ {group}/{name}: Brak kolumn 'sku' lub 'real_status'")
            except Exception as e:
                errors.append(f"❌ {group}/{name}: Błąd - {e}")

        total_skus_processed_count += len(unique_group_skus)

        # Rozdzielanie SKU na partnerów dla danej grupy
        for sku in unique_group_skus:
            if "_" in sku:
                partner = sku.split("_", 1)[0]
                partner_group_data[group][partner].add(sku)
                all_partners.add(partner)

    if not all_partners:
        return None, total_skus_processed_count, 0, errors

    rows_list = []
    for partner in sorted(all_partners):
        row = {"Partner": partner}
        # Zbiór wszystkich SKU tego partnera ze wszystkich grup (do deduplikacji)
        partner_all_skus = set()
        for group in GROUPS:
            skus_in_group = partner_group_data[group][partner]
            row[group] = len(skus_in_group)
            partner_all_skus.update(skus_in_group)
        row[SUM_COL] = len(partner_all_skus)
        rows_list.append(row)

    df_matrix = pd.DataFrame(rows_list).sort_values(SUM_COL, ascending=False)

    # Wiersz podsumowania (TOTAL) na dole
    sum_row_dict = df_matrix.drop(columns=["Partner"]).sum().to_dict()
    sum_row_dict["Partner"] = TOTAL_LABEL
    df_final = pd.concat([df_matrix, pd.DataFrame([sum_row_dict])], ignore_index=True)

    return df_final, total_skus_processed_count, len(all_partners), errors


# --- MODEL TABELI ---
class DataFrameModel(QAbstractTableModel):
    def __init__(self, df=None):
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()

    def set_df(self, df):
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        value = self._df.iat[index.row(), index.column()]
        is_total = self._df.iat[index.row(), 0] == TOTAL_LABEL
        if role == Qt.DisplayRole:
            if isinstance(value, (int, float)) or hasattr(value, "item"):
                return f"{int(value):,}".replace(",", " ")
            return str(value)
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignLeft | Qt.AlignVCenter) if index.column() == 0 else int(Qt.AlignRight | Qt.AlignVCenter)
        if role == Qt.BackgroundRole and is_total:
            return QColor(TOTAL_BG)
        if role == Qt.FontRole and (is_total or self._df.columns[index.column()] == SUM_COL):
            font = QFont()
            font.setBold(True)
            return font
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return None


# --- WORKER (analiza poza wątkiem GUI) ---
class AnalyzeWorker(QObject):
    finished = Signal(object, int, int, list)

    def __init__(self, uploads):
        super().__init__()
        self.uploads = uploads

    def run(self):
        try:
            result = analyze(self.uploads)
        except Exception as e:
            result = (None, 0, 0, [f"❌ Nieoczekiwany błąd: {e}"])
        self.finished.emit(*result)




# --- WIDGETY W STYLU STREAMLIT ---
def label(text, name=None, wrap=True):
    lbl = QLabel(text)
    lbl.setTextFormat(Qt.RichText)
    lbl.setWordWrap(wrap)
    if name:
        lbl.setObjectName(name)
    return lbl


def divider():
    line = QFrame()
    line.setObjectName("divider")
    return line


def metric_card(caption):
    card = QFrame()
    card.setObjectName("card")
    value = QLabel("—")
    value.setObjectName("metricValue")
    cap = QLabel(caption)
    cap.setObjectName("metricLabel")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(16, 12, 16, 12)
    lay.addWidget(cap)
    lay.addWidget(value)
    return card, value


class FileUploader(QFrame):
    """Odpowiednik st.file_uploader: przerywana ramka, drag&drop, 'Browse files' i lista plików."""
    changed = Signal()

    def __init__(self, group):
        super().__init__()
        self.group = group
        self.setObjectName("uploader")
        self.setAcceptDrops(True)

        icon = QLabel("☁️")
        icon.setStyleSheet("font-size: 26px;")
        texts = QVBoxLayout()
        texts.setSpacing(0)
        texts.addWidget(label("Drag and drop files here"))
        texts.addWidget(label("Pliki CSV (wiele naraz)", "muted"))
        browse = QPushButton("Browse files")
        browse.setObjectName("browse")
        browse.setCursor(Qt.PointingHandCursor)
        browse.clicked.connect(self.add_dialog)

        top = QHBoxLayout()
        top.addWidget(icon)
        top.addLayout(texts, stretch=1)
        top.addWidget(browse)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list.setFixedHeight(96)
        self.list.hide()

        self.btn_remove = QPushButton("✕ Usuń zaznaczone")
        self.btn_remove.setObjectName("link")
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear = QPushButton("✕ Wyczyść")
        self.btn_clear.setObjectName("link")
        self.btn_clear.clicked.connect(self.clear)
        self.actions = QWidget()
        self.actions.setStyleSheet("background: transparent;")
        act = QHBoxLayout(self.actions)
        act.setContentsMargins(0, 0, 0, 0)
        act.addStretch(1)
        act.addWidget(self.btn_remove)
        act.addWidget(self.btn_clear)
        self.actions.hide()

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 12)
        lay.addLayout(top)
        lay.addWidget(self.list)
        lay.addWidget(self.actions)

    def files(self):
        return [self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())]

    def add_files(self, paths):
        existing = set(self.files())
        for p in paths:
            if p.lower().endswith(".csv") and p not in existing:
                self.list.addItem(f"📄 {Path(p).name}")
                item = self.list.item(self.list.count() - 1)
                item.setData(Qt.UserRole, p)
                item.setToolTip(p)
                existing.add(p)
        self._refresh()

    def add_dialog(self):
        paths, _ = QFileDialog.getOpenFileNames(self, f"Pliki dla {self.group}", "", "CSV (*.csv)")
        if paths:
            self.add_files(paths)

    def remove_selected(self):
        for item in self.list.selectedItems():
            self.list.takeItem(self.list.row(item))
        self._refresh()

    def clear(self):
        self.list.clear()
        self._refresh()

    def _refresh(self):
        has = self.list.count() > 0
        self.list.setVisible(has)
        self.actions.setVisible(has)
        self.changed.emit()

    def _set_drag(self, on):
        self.setProperty("drag", "true" if on else "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            self._set_drag(True)
            event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self._set_drag(False)

    def dropEvent(self, event):
        self._set_drag(False)
        self.add_files([u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()])
        event.acceptProposedAction()


class GroupColumn(QWidget):
    """Kolumna grupy: st.info z nazwą, uploader, st.success 'Wgrano: N'."""
    changed = Signal()

    def __init__(self, group):
        super().__init__()
        self.uploader = FileUploader(group)
        self.uploader.changed.connect(self._refresh)
        self.success = label("", "success")
        self.success.hide()

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(label(f"📂 <b>{group}</b>", "info"))
        lay.addWidget(self.uploader)
        lay.addWidget(self.success)
        lay.addStretch(1)

    def files(self):
        return self.uploader.files()

    def _refresh(self):
        n = len(self.files())
        self.success.setText(f"Wgrano: {n}")
        self.success.setVisible(n > 0)
        self.changed.emit()


class Expander(QWidget):
    """Odpowiednik st.expander."""

    def __init__(self, title, expanded=False):
        super().__init__()
        self.title = title
        self.button = QToolButton()
        self.button.setObjectName("expander")
        self.button.setCheckable(True)
        self.button.setChecked(expanded)
        self.button.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.button.toggled.connect(self._toggle)
        self.body = label("", None)
        self.body.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.body.setContentsMargins(12, 8, 12, 8)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self.button)
        lay.addWidget(self.body)
        self._toggle(expanded)

    def set_lines(self, lines):
        self.body.setText("<br>".join(lines))

    def _toggle(self, on):
        self.button.setText(("▾ " if on else "▸ ") + self.title)
        self.body.setVisible(on)


# --- OKNO GŁÓWNE ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📦 Analizator SKU Partner-Grupy")
        self.resize(1200, 900)
        self.df_final = None
        self._thread = None
        self._worker = None

        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(48, 32, 48, 32)
        root.setSpacing(14)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setWidget(page)
        self.setCentralWidget(scroll)

        root.addWidget(label("📦 Analizator SKU: Partnerzy i Grupy (Unikalne)", "h1"))
        root.addWidget(label(
            "Analiza plików CSV pod kątem aktywności produktu w kolumnie <b><code>real_status</code></b>.<br>"
            "Za aktywne uznawane są tylko rekordy z wartością <b><code>aktywne</code></b>.<br>"
            "Kolumna <b>SUMA</b> pokazuje liczbę unikalnych SKU per partner (eliminuje duplikaty między grupami)."
        ))

        # 1. Interfejs uploadu (3 kolumny)
        root.addWidget(label("1. Wgraj pliki CSV dla poszczególnych grup", "h3"))
        cols = QHBoxLayout()
        cols.setSpacing(16)
        self.columns = {}
        for g in GROUPS:
            col = GroupColumn(g)
            col.changed.connect(self._update_run_state)
            self.columns[g] = col
            col.setMinimumWidth(0)
            col.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
            cols.addWidget(col, stretch=1)
        root.addLayout(cols)

        root.addWidget(divider())

        self.start_info = label("Wgraj pliki CSV do sekcji powyżej, aby rozpocząć.", "info")
        root.addWidget(self.start_info)
        self.btn_run = QPushButton("🚀 Uruchom Analizę")
        self.btn_run.setMinimumHeight(40)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        self.btn_run.clicked.connect(self.run_analysis)
        root.addWidget(self.btn_run)

        self.message = label("", "error")
        self.message.hide()
        root.addWidget(self.message)

        # Wyniki
        self.results = QWidget()
        res = QVBoxLayout(self.results)
        res.setContentsMargins(0, 0, 0, 0)
        res.setSpacing(14)
        res.addWidget(label("📊 Wyniki: Partnerzy vs Grupy", "h2"))
        res.addWidget(label(
            "Kolumna <b>SUMA (Unikalne)</b> weryfikuje duplikaty. Jeśli ten sam SKU jest w grupie SHUMEE "
            "i GREATSTORE, zostanie policzony tylko raz w sumie.", "info"
        ))
        metrics = QHBoxLayout()
        metrics.setSpacing(16)
        m1, self.metric_rows = metric_card("Przetworzone wiersze (Suma grup)")
        m2, self.metric_partners = metric_card("Liczba Partnerów")
        metrics.addWidget(m1)
        metrics.addWidget(m2)
        res.addLayout(metrics)

        self.model = DataFrameModel()
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setMinimumHeight(420)
        res.addWidget(self.table)

        self.errors = Expander("⚠️ Wykryto błędy w niektórych plikach")
        res.addWidget(self.errors)

        res.addWidget(label("📥 Pobierz Raport", "h2"))
        self.btn_save = QPushButton("💾 Pobierz Podsumowanie")
        self.btn_save.setMinimumHeight(40)
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.clicked.connect(self.save_csv)
        res.addWidget(self.btn_save)

        self.results.hide()
        root.addWidget(self.results)
        root.addStretch(1)

        self._update_run_state()

    def _uploads(self):
        return {g: c.files() for g, c in self.columns.items()}

    def _update_run_state(self):
        has_files = sum(len(v) for v in self._uploads().values()) > 0
        self.start_info.setVisible(not has_files)
        self.btn_run.setVisible(has_files)
        self.btn_run.setEnabled(self._thread is None)

    def run_analysis(self):
        self.btn_run.setEnabled(False)
        self.btn_run.setText("⏳ Przetwarzanie danych...")
        self.message.hide()

        self._thread = QThread()
        self._worker = AnalyzeWorker(self._uploads())
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_finished)
        self._worker.finished.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)
        self._thread.start()

    def _cleanup_thread(self):
        self._thread.deleteLater()
        self._worker.deleteLater()
        self._thread = None
        self._worker = None
        self.btn_run.setText("🚀 Uruchom Analizę")
        self._update_run_state()

    def _on_finished(self, df_final, total_processed, partners, errors):
        self.errors.set_lines(errors)
        self.errors.setVisible(bool(errors))

        if df_final is None:
            self.df_final = None
            self.results.hide()
            if errors:
                self.message.setText("Nie znaleziono danych. Sprawdź błędy poniżej.<br><br>" + "<br>".join(errors))
            else:
                self.message.setText(
                    "Nie znaleziono żadnych aktywnych SKU (wg kolumny 'real_status') z poprawnymi prefiksami."
                )
            self.message.show()
            return

        self.df_final = df_final
        self.metric_rows.setText(str(total_processed))
        self.metric_partners.setText(str(partners))
        self.model.set_df(df_final)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.btn_save.setText(f"💾 Pobierz Podsumowanie ({self._default_filename()})")
        self.results.show()

    @staticmethod
    def _default_filename():
        return f"podsumowanie_CDON_{datetime.datetime.now().strftime('%Y-%m-%d')}.csv"

    def save_csv(self):
        if self.df_final is None:
            return
        default = str(Path.home() / "Downloads" / self._default_filename())
        path, _ = QFileDialog.getSaveFileName(self, "Zapisz podsumowanie", default, "CSV (*.csv)")
        if not path:
            return
        try:
            self.df_final.to_csv(path, sep=";", index=False, encoding="utf-8-sig")
        except Exception as e:
            QMessageBox.critical(self, "Błąd zapisu", str(e))
            return
        self.statusBar().showMessage(f"Zapisano: {path}", 8000)


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
