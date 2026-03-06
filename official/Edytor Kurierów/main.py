"""
Config.ini GUI Editor - Elegancki edytor konfiguracji transportu
Autor: Patryk Libert
Styl: PyQt6 z różowymi gradientami i efektami przezroczystości
"""

import sys
import json
import os
from pathlib import Path
from configparser import ConfigParser
from typing import Optional, Dict, List, Tuple
from dataclasses import dataclass
from enum import Enum

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QLineEdit, QPushButton, QLabel,
    QTabWidget, QDialog, QDialogButtonBox, QMessageBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QSplitter, QSpinBox, QDoubleSpinBox,
    QComboBox, QCheckBox, QTextEdit, QFileDialog, QStyleFactory, QScrollArea,
    QGroupBox, QAbstractItemView, QMenu
)
from PyQt6.QtCore import (
    Qt, QSize, QTimer, QPropertyAnimation, QEasingCurve,
    pyqtSignal, QObject, QRect
)
from PyQt6.QtGui import (
    QColor, QFont, QIcon, QPixmap, QLinearGradient, QPainter,
    QBrush, QPen, QFontMetrics, QTextCursor, QAction
)
from PyQt6.QtWidgets import QWidget


class Theme(Enum):
    """Paleta kolorów dla eleganccy różowych tematów"""
    PRIMARY_PINK = "#FF1493"
    LIGHT_PINK = "#FFB6D9"
    PALE_PINK = "#FFE4E1"
    DARK_PINK = "#C71585"
    ACCENT_PURPLE = "#9932CC"
    LIGHT_PURPLE = "#E6D5FF"
    WHITE = "#FFFFFF"
    LIGHT_GRAY = "#F5F5F5"
    DARK_GRAY = "#2C2C2C"
    TEXT_DARK = "#1A1A1A"


class GradientButton(QPushButton):
    """Przycisk z gradientem różowo-fioletowym i efektami hover"""
    
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.is_hovered = False
        self.setMinimumHeight(40)
        self.setMinimumWidth(100)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
    def enterEvent(self, event):
        self.is_hovered = True
        self.update()
        
    def leaveEvent(self, event):
        self.is_hovered = False
        self.update()
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Gradient
        gradient = QLinearGradient(0, 0, 0, self.height())
        if self.is_hovered:
            gradient.setColorAt(0, QColor("#FF1493"))
            gradient.setColorAt(1, QColor("#C71585"))
        else:
            gradient.setColorAt(0, QColor("#FFB6D9"))
            gradient.setColorAt(1, QColor("#FF1493"))
        
        # Zaokrąglone krawędzie
        painter.fillRect(self.rect(), QBrush(gradient))
        
        # Cień pod przyciskiem
        if self.is_hovered:
            painter.setPen(QPen(QColor("#FF1493"), 2))
        else:
            painter.setPen(QPen(QColor("#FFE4E1"), 2))
        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 8, 8)
        
        # Tekst
        painter.setPen(QColor("white"))
        font = QFont()
        font.setPointSize(10)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.text())


class StyledLineEdit(QLineEdit):
    """Eleganckie pole tekstowe z różowym obramowaniem"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(35)
        self.setStyleSheet("""
            QLineEdit {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
                border-radius: 6px;
                padding: 5px 10px;
                font-size: 11pt;
                color: #1A1A1A;
            }
            QLineEdit:focus {
                border: 2px solid #FF1493;
                background-color: #fff9fc;
            }
        """)


class ConfigEditorDialog(QDialog):
    """Dialog do dodawania/edycji sekcji config"""
    
    def __init__(self, title: str, section_name: str = "", values: Dict = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 500, 400)
        self.setStyleSheet(self._get_stylesheet())
        
        layout = QVBoxLayout()
        
        # Nazwa sekcji
        name_layout = QHBoxLayout()
        name_label = QLabel("Nazwa grupy:")
        name_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.name_input = StyledLineEdit()
        self.name_input.setText(section_name)
        self.name_input.setReadOnly(section_name != "")  # Nie można zmieniać istniejącej
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.name_input)
        layout.addLayout(name_layout)
        
        # Tabela parametrów
        layout.addWidget(QLabel("Parametry:"))
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Klucz (np. PL)", "Wartość (np. DPD)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
                border-radius: 6px;
            }
            QHeaderView::section {
                background-color: #FFB6D9;
                color: white;
                padding: 5px;
                border: none;
                font-weight: bold;
            }
        """)
        
        if values:
            self.table.setRowCount(len(values))
            for i, (key, val) in enumerate(values.items()):
                self.table.setItem(i, 0, QTableWidgetItem(key))
                self.table.setItem(i, 1, QTableWidgetItem(val))
        else:
            self.table.setRowCount(0)
        
        layout.addWidget(self.table)
        
        # Przycisk dodaj wiersz
        add_row_btn = GradientButton("+ Dodaj wiersz")
        add_row_btn.clicked.connect(self._add_row)
        layout.addWidget(add_row_btn)
        
        # Przycisk usuń wiersz
        del_row_btn = GradientButton("- Usuń ostatni wiersz")
        del_row_btn.clicked.connect(self._delete_row)
        layout.addWidget(del_row_btn)
        
        # Przyciski OK/Anuluj
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
    
    def _add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(""))
        self.table.setItem(row, 1, QTableWidgetItem(""))
    
    def _delete_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.rowCount() - 1)
    
    def get_data(self) -> Tuple[str, Dict]:
        """Zwraca (nazwa_sekcji, {klucz: wartość})"""
        name = self.name_input.text().strip()
        values = {}
        for i in range(self.table.rowCount()):
            key_item = self.table.item(i, 0)
            val_item = self.table.item(i, 1)
            if key_item and val_item:
                key = key_item.text().strip()
                val = val_item.text().strip()
                if key and val:
                    values[key] = val
        return name, values
    
    def _get_stylesheet(self) -> str:
        return """
            QDialog {
                background-color: #FFF9FC;
            }
            QLabel {
                color: #1A1A1A;
                font-size: 11pt;
            }
            QTableWidget {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
            }
            QPushButton {
                border: none;
            }
        """


class ConfigGUIEditor(QMainWindow):
    """Główne okno edytora konfiguracji"""
    
    def __init__(self, config_file: str = None):
        super().__init__()
        self.config_file = config_file or str(
            Path(__file__).parent / "z pulpitu" / "config.ini"
        )
        self.config = ConfigParser()
        # Case-sensitive - nie konwertuje kluczy na małe litery
        self.config.optionxform = str
        
        # Sprawdź czy plik istnieje
        if not Path(self.config_file).exists():
            print(f"⚠️  Plik nie znaleziony: {self.config_file}")
            self.config_file = None
        else:
            self.config.read(self.config_file, encoding='utf-8')
            print(f"✅ Plik załadowany: {self.config_file}")
            print(f"📋 Sekcje: {self.config.sections()}")
        
        self._setup_ui()
        self._load_config()
        
        # Jeśli brakuje pliku, pokaż dialog
        if self.config_file is None:
            self.config_file = ""
            self._show_no_file_warning()
        
    def _setup_ui(self):
        """Konfiguracja interfejsu użytkownika"""
        self.setWindowTitle("🎀 Config.ini Editor - Transportowe Zarządzanie")
        self.setGeometry(100, 100, 1200, 700)
        self.setStyleSheet(self._get_main_stylesheet())
        
        # Menu bar
        self._create_menu_bar()
        
        # Widget centralny
        central_widget = QWidget()
        main_layout = QHBoxLayout()
        
        # Panel lewy - lista sekcji
        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("📋 Grupy konfiguracji:"))
        left_panel.setContentsMargins(10, 10, 10, 10)
        
        self.section_list = QListWidget()
        self.section_list.itemClicked.connect(self._on_section_selected)
        self.section_list.setStyleSheet("""
            QListWidget {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
                border-radius: 6px;
            }
            QListWidget::item:selected {
                background-color: #FFB6D9;
                color: white;
            }
        """)
        left_panel.addWidget(self.section_list)
        
        # Przyciski panelu lewego
        buttons_layout = QVBoxLayout()
        new_section_btn = GradientButton("➕ Nowa grupa")
        new_section_btn.clicked.connect(self._new_section)
        buttons_layout.addWidget(new_section_btn)
        
        edit_section_btn = GradientButton("✏️ Edytuj")
        edit_section_btn.clicked.connect(self._edit_section)
        buttons_layout.addWidget(edit_section_btn)
        
        delete_section_btn = GradientButton("🗑️ Usuń")
        delete_section_btn.clicked.connect(self._delete_section)
        buttons_layout.addWidget(delete_section_btn)
        
        left_panel.addLayout(buttons_layout)
        
        # Panel prawy - edycja zawartości
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("🔧 Edycja zawartości:"))
        
        self.content_table = QTableWidget()
        self.content_table.setColumnCount(2)
        self.content_table.setHorizontalHeaderLabels(["Klucz", "Wartość"])
        self.content_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.content_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        # Edycja - wszystkie sposoby
        self.content_table.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
        self.content_table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
                border-radius: 6px;
            }
            QHeaderView::section {
                background-color: #FF1493;
                color: white;
                padding: 5px;
                border: none;
                font-weight: bold;
            }
        """)
        right_panel.addWidget(self.content_table)
        
        # Przyciski edycji
        content_buttons = QHBoxLayout()
        add_param_btn = GradientButton("➕ Dodaj parametr")
        add_param_btn.clicked.connect(self._add_parameter)
        content_buttons.addWidget(add_param_btn)
        
        delete_param_btn = GradientButton("🗑️ Usuń parametr")
        delete_param_btn.clicked.connect(self._delete_parameter)
        content_buttons.addWidget(delete_param_btn)
        
        right_panel.addLayout(content_buttons)
        
        # Przycisk zapisu
        save_btn = GradientButton("💾 Zapisz zmiany")
        save_btn.clicked.connect(self._save_config)
        right_panel.addWidget(save_btn)
        
        # Układ poziomy
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        left_widget.setMaximumWidth(300)
        
        right_widget = QWidget()
        right_widget.setLayout(right_panel)
        
        main_layout.addWidget(left_widget)
        main_layout.addWidget(right_widget, 1)
        
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
    
    def _load_config(self):
        """Wczytaj sekcje z config.ini"""
        self.section_list.clear()
        for section in self.config.sections():
            item = QListWidgetItem(section)
            item.setIcon(QIcon())
            self.section_list.addItem(item)
    
    def _show_no_file_warning(self):
        """Pokaż ostrzeżenie jeśli brakuje pliku"""
        reply = QMessageBox.question(
            self,
            "Plik nie znaleziony",
            "Domyślny plik config.ini nie został znaleziony.\n\nChcesz go teraz wybrać?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._open_config_file()
    
    def _on_section_selected(self, item):
        """Załaduj zawartość wybranej sekcji"""
        section = item.text()
        self.content_table.setRowCount(0)
        
        if self.config.has_section(section):
            items = self.config.items(section)
            self.content_table.setRowCount(len(items))
            for i, (key, val) in enumerate(items):
                self.content_table.setItem(i, 0, QTableWidgetItem(key))
                self.content_table.setItem(i, 1, QTableWidgetItem(val))
    
    def _new_section(self):
        """Dodaj nową sekcję"""
        dialog = ConfigEditorDialog("Nowa grupa konfiguracji", parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name, values = dialog.get_data()
            if name and values:
                if not self.config.has_section(name):
                    self.config.add_section(name)
                for key, val in values.items():
                    self.config.set(name, key, val)
                self._load_config()
                QMessageBox.information(self, "Sukces", f"Dodano grupę '{name}'")
    
    def _edit_section(self):
        """Edytuj wybraną sekcję"""
        item = self.section_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Info", "Wybierz grupę do edycji!")
            return
        
        section = item.text()
        values = dict(self.config.items(section))
        
        dialog = ConfigEditorDialog(f"Edycja grupy: {section}", section, values, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name, new_values = dialog.get_data()
            
            # Usuń stare wpisy
            for key in self.config.options(section):
                self.config.remove_option(section, key)
            
            # Dodaj nowe wpisy
            for key, val in new_values.items():
                self.config.set(section, key, val)
            
            self._load_config()
            QMessageBox.information(self, "Sukces", f"Zaktualizowano grupę '{section}'")
    
    def _delete_section(self):
        """Usuń wybraną sekcję"""
        item = self.section_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Info", "Wybierz grupę do usunięcia!")
            return
        
        section = item.text()
        reply = QMessageBox.question(
            self, "Potwierdzenie",
            f"Na pewno usunąć grupę '{section}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.config.remove_section(section)
            self._load_config()
            QMessageBox.information(self, "Sukces", f"Usunięto grupę '{section}'")
    
    def _add_parameter(self):
        """Dodaj parametr do wybranej sekcji"""
        item = self.section_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Info", "Wybierz grupę!")
            return
        
        row = self.content_table.rowCount()
        self.content_table.insertRow(row)
        self.content_table.setItem(row, 0, QTableWidgetItem(""))
        self.content_table.setItem(row, 1, QTableWidgetItem(""))
    
    def _delete_parameter(self):
        """Usuń parametr z wybranej sekcji"""
        if self.content_table.rowCount() > 0:
            self.content_table.removeRow(self.content_table.rowCount() - 1)
    
    def _save_config(self):
        """Zapisz zmiany do pliku"""
        item = self.section_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Info", "Wybierz grupę by ją zapisać!")
            return
        
        section = item.text()
        
        # Wyczyść sekcję
        for key in self.config.options(section):
            self.config.remove_option(section, key)
        
        # Dodaj nowe dane z tabeli
        for i in range(self.content_table.rowCount()):
            key_item = self.content_table.item(i, 0)
            val_item = self.content_table.item(i, 1)
            if key_item and val_item:
                key = key_item.text().strip()
                val = val_item.text().strip()
                if key and val:
                    self.config.set(section, key, val)
        
        # Zapisz do pliku
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            QMessageBox.information(self, "Sukces", f"Zapisano konfigurację!")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Błąd zapisu: {str(e)}")
    
    def _create_menu_bar(self):
        """Tworzenie menu bar'a"""
        menubar = self.menuBar()
        menubar.setStyleSheet("""
            QMenuBar {
                background-color: #FFFFFF;
                color: #000000;
                border-bottom: 2px solid #FFB6D9;
            }
            QMenuBar::item:selected {
                background-color: #FFB6D9;
            }
            QMenu {
                background-color: #FFFFFF;
                color: #000000;
                border: 1px solid #FFB6D9;
            }
            QMenu::item:selected {
                background-color: #FFB6D9;
            }
        """)
        
        # Menu Plik
        file_menu = menubar.addMenu("📁 Plik")
        
        # Akcja: Otwórz config
        open_action = QAction("📂 Otwórz config.ini...", self)
        open_action.triggered.connect(self._open_config_file)
        file_menu.addAction(open_action)
        
        file_menu.addSeparator()
        
        # Akcja: Wyjście
        exit_action = QAction("❌ Wyjście", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
    
    def _open_config_file(self):
        """Dialog do wyboru pliku config"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz plik config.ini",
            str(Path.home() / "Desktop"),
            "Config Files (*.ini);;All Files (*.*)"
        )
        
        if file_path:
            self.config_file = file_path
            # Przeładuj config z nowego pliku!
            self.config = ConfigParser()
            self.config.optionxform = str  # Case-sensitive
            self.config.read(self.config_file, encoding='utf-8')
            # Odśwież UI
            self._load_config()
            self.content_table.setRowCount(0)
            self.setWindowTitle(f"🎀 Config.ini Editor - {Path(file_path).name}")
            QMessageBox.information(self, "Sukces", f"Załadowano:\n{file_path}")
    
    def _get_main_stylesheet(self) -> str:
        """Global stylesheet z różowymi gradientami"""
        return """
            QMainWindow {
                background-color: #FFF9FC;
            }
            QLabel {
                color: #000000;
                font-weight: bold;
                font-size: 11pt;
            }
            QListWidget {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
                border-radius: 6px;
                color: #000000;
            }
            QListWidget::item {
                color: #000000;
            }
            QTableWidget {
                background-color: #FFFFFF;
                border: 2px solid #FFB6D9;
                border-radius: 6px;
                gridline-color: #FFE4E1;
                color: #000000;
            }
            QHeaderView::section {
                background-color: #FF1493;
                color: white;
                padding: 5px;
                border: none;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 5px;
                border: 1px solid #FFE4E1;
                color: #000000;
            }
            QTableWidget::item:selected {
                background-color: #FFB6D9;
                color: #000000;
            }
            QMessageBox {
                background-color: #FFF9FC;
                color: #000000;
            }
        """


def main():
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('Fusion'))
    
    editor = ConfigGUIEditor()
    editor.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
