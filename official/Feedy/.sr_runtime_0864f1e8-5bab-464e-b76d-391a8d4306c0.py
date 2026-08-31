import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox, Dialog
from tkinter.filedialog import askopenfilename
import csv
import os
import time
import json
import keyboard
import pyautogui
from functools import partial

# --- NOWOŚĆ: Poprawka skalowania dla Windows ---
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except ImportError:
    pass # Ten kod zadziała tylko na Windows

# --- Zmienne globalne i konfiguracja ---
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
# NOWOŚĆ: Zmiana na plik JSON jako główny plik konfiguracyjny
JSON_FILE = os.path.join(desktop_path, 'macros.json')
# Ścieżka do starego pliku CSV na potrzeby konwersji
OLD_CSV_FILE = os.path.join(desktop_path, 'macros.csv')


# Znaczniki akcji
TAB_MARKER = "{TAB}"
ENTER_MARKER = "{ENTER}"
SHIFT_TAB_MARKER = "{SHIFT_TAB}"
CTRL_TAB_MARKER = "{CTRL_TAB}"
UP_MARKER = "{UP}"
DOWN_MARKER = "{DOWN}"
LEFT_MARKER = "{LEFT}"
RIGHT_MARKER = "{RIGHT}"
CLICK_IMAGE_MARKER = "{CLICK_IMAGE}"
SCROLL_CLICK_MARKER = "{SCROLL_CLICK}"

HOTKEY_NONE = "Brak"
HOTKEY_BASE_KEYS = [f'F{i}' for i in range(1, 13)]

# --- Sekwencja klawiszy między makrami ---
INTERMEDIATE_ACTIONS = [
    {'type': 'key', 'value': 'tab'}, {'type': 'key', 'value': 'tab'},
    {'type': 'key', 'value': 'tab'}, {'type': 'key', 'value': 'tab'},
    {'type': 'key', 'value': 'tab'}, {'type': 'key', 'value': 'tab'},
    {'type': 'key', 'value': 'enter'},
    {'type': 'key', 'value': 'shift+tab'}, {'type': 'key', 'value': 'shift+tab'},
    {'type': 'key', 'value': 'shift+tab'}, {'type': 'key', 'value': 'shift+tab'},
    {'type': 'key', 'value': 'enter'},
    {'type': 'key', 'value': 'tab'}, {'type': 'key', 'value': 'tab'},
    {'type': 'key', 'value': 'tab'}, {'type': 'key', 'value': 'tab'},
    {'type': 'key', 'value': 'right'}
]

class MacroApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title("Menedżer Makr z Rozpoznawaniem Obrazu")
        self.geometry("1200x650") 

        self.macros = {}
        self.is_running = False
        self._setup_ui()
        self.load_and_display_macros()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _setup_ui(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=BOTH, expand=True)

        ttk.Label(main_frame, text="Dostępne Makra (Ctrl+Click, aby zaznaczyć wiele):", font=("-size 12 -weight bold")).pack(fill=X, pady=(0, 5))
        
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=BOTH, expand=True)
        
        self.macro_listbox = ttk.Treeview(list_frame, columns=("macro_name", "hotkey"), show="headings", selectmode="extended")
        self.macro_listbox.heading("macro_name", text="Nazwa Makra")
        self.macro_listbox.heading("hotkey", text="Skrót Indywidualny")
        self.macro_listbox.column("hotkey", width=150, anchor='center')
        self.macro_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.macro_listbox.bind("<Double-1>", lambda event: self.edit_focused_macro())
        
        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.macro_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.macro_listbox.configure(yscrollcommand=scrollbar.set)

        run_settings_frame = ttk.Frame(main_frame, padding=(0, 10))
        run_settings_frame.pack(fill=X)

        ttk.Button(run_settings_frame, text="Uruchom Zaznaczone", command=self.run_macro_with_repeat, bootstyle="primary").grid(row=0, column=0, pady=5)
        
        ttk.Label(run_settings_frame, text="Ilość powtórzeń sekwencji:").grid(row=0, column=1, padx=(10, 5))
        self.repeat_entry = ttk.Entry(run_settings_frame, width=5)
        self.repeat_entry.insert(0, "1")
        self.repeat_entry.grid(row=0, column=2)
        
        self.ctrl_tab_var = ttk.BooleanVar(value=False)
        ttk.Checkbutton(run_settings_frame, text="Przełącz kartę po sekwencji (Ctrl+Tab)", variable=self.ctrl_tab_var, bootstyle="round-toggle").grid(row=0, column=3, padx=10)

        ttk.Label(run_settings_frame, text="Skrót do uruchomienia sekwencji:").grid(row=1, column=0, pady=(10,0), sticky=W)
        hotkey_options = [HOTKEY_NONE] + HOTKEY_BASE_KEYS
        self.sequence_hotkey_combo = ttk.Combobox(run_settings_frame, values=hotkey_options, state="readonly", width=8)
        self.sequence_hotkey_combo.set("F11")
        self.sequence_hotkey_combo.grid(row=1, column=1, columnspan=2, pady=(10,0), sticky=W, padx=5)
        self.sequence_hotkey_combo.bind("<<ComboboxSelected>>", self.on_sequence_hotkey_change)
        
        self.skip_intermediate_var = ttk.BooleanVar(value=False)
        ttk.Checkbutton(run_settings_frame, text="Pomiń akcje po 1. makrze", variable=self.skip_intermediate_var, bootstyle="round-toggle").grid(row=1, column=3, padx=10, pady=(10,0))

        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=X, pady=15)

        ttk.Button(controls_frame, text="Dodaj Makro Strukturalne", command=self.add_structured_macro, bootstyle="success").pack(side=LEFT, fill=X, expand=True, padx=5, pady=5)
        ttk.Button(controls_frame, text="Dodaj Nowe (Standardowe)", command=self.add_new_macro, bootstyle="success-outline").pack(side=LEFT, fill=X, expand=True, padx=5, pady=5)
        ttk.Button(controls_frame, text="Edytuj", command=self.edit_focused_macro, bootstyle="info").pack(side=LEFT, fill=X, expand=True, padx=5, pady=5)
        ttk.Button(controls_frame, text="Kopiuj", command=self.copy_focused_macro, bootstyle="warning").pack(side=LEFT, fill=X, expand=True, padx=5, pady=5)
        ttk.Button(controls_frame, text="Usuń", command=self.delete_focused_macro, bootstyle="danger").pack(side=LEFT, fill=X, expand=True, padx=5, pady=5)
        ttk.Button(controls_frame, text="Odśwież", command=self.load_and_display_macros, bootstyle="secondary").pack(side=LEFT, fill=X, expand=True, padx=5, pady=5)

        # AKTUALIZACJA: Pasek stanu wskazuje na plik JSON
        self.status_bar = ttk.Label(self, text=f"Gotowy. Plik: {JSON_FILE} | Naciśnij F12, aby zatrzymać makro.", relief=SUNKEN, anchor=W, padding=5)
        self.status_bar.pack(side=BOTTOM, fill=X)

    def on_sequence_hotkey_change(self, event=None):
        self.start_hotkey_listener()

    def start_hotkey_listener(self):
        keyboard.unhook_all()
        keyboard.add_hotkey('f12', self.stop_all_macros)
        
        for name, data in self.macros.items():
            hotkey_str_user = data.get('hotkey')
            if not hotkey_str_user or hotkey_str_user == HOTKEY_NONE:
                continue
            try:
                lib_hotkey_str = hotkey_str_user.lower()
                keyboard.add_hotkey(lib_hotkey_str, partial(self._schedule_macro_execution, [name]), suppress=True)
            except Exception as e:
                Messagebox.show_warning(f"Nie udało się zarejestrować skrótu '{hotkey_str_user}' dla makra '{name}'.\nBłąd: {e}", "Błąd skrótu")

        sequence_hotkey = self.sequence_hotkey_combo.get()
        if sequence_hotkey and sequence_hotkey != HOTKEY_NONE:
            try:
                keyboard.add_hotkey(sequence_hotkey.lower(), self.run_macro_with_repeat, suppress=True)
                self.status_bar.config(text=f"Gotowy. Skrót sekwencji: {sequence_hotkey}. Naciśnij F12, aby zatrzymać.")
            except Exception as e:
                Messagebox.show_warning(f"Nie udało się zarejestrować skrótu sekwencji '{sequence_hotkey}'.\nBłąd: {e}", "Błąd skrótu")


    def stop_all_macros(self):
        self.is_running = False
        self.status_bar.config(text="Wykonywanie makra zatrzymane przez użytkownika (F12).")

    def _schedule_macro_execution(self, macro_names):
        if self.is_running: return
        
        try:
            repeats = int(self.repeat_entry.get())
            if repeats < 1: repeats = 1
        except ValueError:
            repeats = 1
        
        if not macro_names: return

        self.is_running = True
        self.after(1, self._execute_macro_loop, macro_names, repeats)

    def on_closing(self):
        keyboard.unhook_all()
        self.destroy()

    def get_focused_macro_name(self):
        focused_item = self.macro_listbox.focus()
        if not focused_item:
            return None
        return self.macro_listbox.item(focused_item, "values")[0]

    def get_selected_macro_names(self):
        if not self.macro_listbox.selection():
            return []
        
        all_items = self.macro_listbox.get_children('')
        selected_ids = self.macro_listbox.selection()
        
        ordered_selected_names = [self.macro_listbox.item(item_id, "values")[0] for item_id in all_items if item_id in selected_ids]
        return ordered_selected_names

    def load_and_display_macros(self):
        self.macros = self._load_macros_from_file()
        for item in self.macro_listbox.get_children():
            self.macro_listbox.delete(item)
        for name, data in sorted(self.macros.items()):
            hotkey = data.get('hotkey', HOTKEY_NONE)
            self.macro_listbox.insert("", END, values=(name, hotkey))
        
        self.start_hotkey_listener()

    def run_macro_with_repeat(self):
        if self.is_running:
            Messagebox.show_warning("Inne makro jest już w trakcie wykonywania.", "Makro aktywne")
            return

        macro_names = self.get_selected_macro_names()
        if not macro_names:
            if self.focus_get() != self.macro_listbox:
                 Messagebox.show_warning("Proszę wybrać co najmniej jedno makro z listy.", "Brak zaznaczenia")
            return
        
        try:
            repeats = int(self.repeat_entry.get())
            if repeats < 1: repeats = 1
        except ValueError:
            repeats = 1
        
        self.is_running = True
        self.countdown(1, macro_names, repeats)

    def countdown(self, seconds_left, macro_names, repeats):
        if not self.is_running: return
        if seconds_left > 0:
            self.status_bar.config(text=f"Start za {seconds_left} sekund... Naciśnij F12, aby anulować.")
            self.after(1000, self.countdown, seconds_left - 1, macro_names, repeats)
        else:
            self.after(1, self._execute_macro_loop, macro_names, repeats)

    def _execute_macro_loop(self, macro_names, repeats):
        for i in range(repeats):
            if not self.is_running: break

            for j, macro_name in enumerate(macro_names):
                if not self.is_running: break

                self.status_bar.config(text=f"Powt. {i + 1}/{repeats} | Krok {j + 1}/{len(macro_names)}: '{macro_name}'...")
                self.update_idletasks()

                actions = self.macros.get(macro_name, {}).get('actions', [])
                self.execute_actions(actions)

                if not self.is_running: break

                is_last_macro_in_sequence = (j == len(macro_names) - 1)
                if not is_last_macro_in_sequence and (not self.skip_intermediate_var.get() or j > 0):
                    self.status_bar.config(text="Wykonywanie akcji pośrednich...")
                    self.update_idletasks()
                    time.sleep(0.1)
                    self.execute_actions(INTERMEDIATE_ACTIONS)
                    if not self.is_running: break
                    time.sleep(0.1)
            
            if not self.is_running:
                self.status_bar.config(text="Pętla powtórzeń zatrzymana.")
                break

            if i < repeats - 1:
                if self.ctrl_tab_var.get():
                    self.status_bar.config(text=f"Sekwencja {i + 1} zakończona. Przełączanie karty...")
                    self.update_idletasks()
                    time.sleep(0.1)
                    pyautogui.hotkey('ctrl', 'tab')
                    time.sleep(0.2)
                else:
                    time.sleep(0.2)
        
        if self.is_running:
            self.status_bar.config(text="Zakończono wykonywanie wszystkich makr.")
        self.is_running = False

    def execute_actions(self, actions):
        try:
            time.sleep(0.1)
            for action in actions:
                if not self.is_running: break
                
                action_type = action['type']
                action_value = action['value']
                confidence = action.get('confidence', 0.8)

                if action_type == "key":
                    keyboard.press_and_release(action_value)
                elif action_type == "text":
                    keyboard.write(str(action_value), delay=0.01)
                elif action_type == "click_image":
                    self._find_and_click(action_value, confidence=confidence)
                elif action_type == "scroll_click":
                    self._scroll_and_click(action_value, confidence=confidence)
                
                time.sleep(0.08) # Lekko zwiększona pauza dla większej stabilności
        except Exception as e:
            self.status_bar.config(text=f"Błąd podczas wykonywania makra: {e}")
            Messagebox.show_error(f"Wystąpił błąd: {e}", "Błąd wykonania makra")
            self.is_running = False

    def _find_and_click(self, image_path, confidence=0.8):
        if not self.is_running: return False
        try:
            location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
            if location:
                pyautogui.moveTo(location)
                pyautogui.click()
                return True
            else:
                raise Exception(f"Nie znaleziono obrazka: {os.path.basename(image_path)}")
        except Exception as e:
            raise Exception(f"Błąd podczas szukania obrazka: {e}")

    def _scroll_and_click(self, image_path, confidence=0.8, max_scrolls=50):
        for _ in range(max_scrolls):
            if not self.is_running: return False
            try:
                if self._find_and_click(image_path, confidence):
                    return True
            except Exception:
                pyautogui.scroll(-500)
                time.sleep(0.1)
        raise Exception(f"Nie znaleziono obrazka po {max_scrolls} próbach przewijania: {os.path.basename(image_path)}")

    def add_structured_macro(self):
        dialog = AddStructuredMacroDialog(self)
        dialog.show()
        if dialog.result:
            name, data = dialog.result
            if name in self.macros:
                if Messagebox.show_question(f"Makro o nazwie '{name}' już istnieje. Nadpisać?", "Pytanie") != "Yes":
                    return
            self.macros[name] = data
            self._save_macros_to_file()
            self.load_and_display_macros()

    def add_new_macro(self):
        dialog = AddEditMacroDialog(self)
        dialog.show()
        if dialog.result:
            name, data = dialog.result
            if name in self.macros:
                if Messagebox.show_question(f"Makro o nazwie '{name}' już istnieje. Nadpisać?", "Pytanie") != "Yes":
                    return
            self.macros[name] = data
            self._save_macros_to_file()
            self.load_and_display_macros()

    def edit_focused_macro(self):
        original_name = self.get_focused_macro_name()
        if not original_name:
            Messagebox.show_warning("Proszę wybrać makro do edycji.", "Brak zaznaczenia")
            return
        
        original_data = self.macros.get(original_name, {})
        
        parsed_data = parse_structured_macro(original_data.get('actions', []))

        if parsed_data:
            dialog = AddStructuredMacroDialog(self, name=original_name, initial_data=parsed_data)
        else:
            dialog = AddEditMacroDialog(self, name=original_name, data=original_data)
        
        dialog.show()

        if dialog.result:
            new_name, new_data = dialog.result
            if new_name != original_name:
                if new_name in self.macros:
                    if Messagebox.show_question(f"Makro o nazwie '{new_name}' już istnieje. Nadpisać?", "Pytanie") != "Yes":
                        return
                del self.macros[original_name]
            self.macros[new_name] = new_data
            self._save_macros_to_file()
            self.load_and_display_macros()

    def copy_focused_macro(self):
        original_name = self.get_focused_macro_name()
        if not original_name:
            Messagebox.show_warning("Proszę wybrać makro do skopiowania.", "Brak zaznaczenia")
            return

        new_name = f"Kopia {original_name}"
        counter = 2
        while new_name in self.macros:
            new_name = f"Kopia {original_name} ({counter})"
            counter += 1

        actions_to_copy = self.macros.get(original_name, {}).get('actions', []).copy()
        self.macros[new_name] = {'actions': actions_to_copy, 'hotkey': HOTKEY_NONE}
        self._save_macros_to_file()
        self.load_and_display_macros()

    def delete_focused_macro(self):
        macro_name = self.get_focused_macro_name()
        if not macro_name:
            Messagebox.show_warning("Proszę wybrać makro do usunięcia.", "Brak zaznaczenia")
            return
        
        if Messagebox.show_question(f"Czy na pewno chcesz usunąć makro '{macro_name}'?", "Potwierdzenie") == "Yes":
            del self.macros[macro_name]
            self._save_macros_to_file()
            self.load_and_display_macros()

    # --- NOWA LOGIKA WCZYTYWANIA PLIKÓW ---
    def _load_macros_from_file(self):
        """
        Wczytuje makra z pliku JSON.
        Jeśli plik JSON nie istnieje, próbuje konwertować stary plik CSV.
        Jeśli żaden plik nie istnieje, tworzy nowy, pusty plik JSON.
        """
        macros = {}
        if os.path.exists(JSON_FILE):
            # --- 1. Wczytaj plik JSON (preferowany) ---
            try:
                with open(JSON_FILE, 'r', encoding='utf-8') as f:
                    macros = json.load(f)
            except json.JSONDecodeError:
                Messagebox.show_error(f"Plik {JSON_FILE} jest uszkodzony. Tworzenie kopii zapasowej i nowego pliku.", "Błąd JSON")
                try:
                    os.rename(JSON_FILE, JSON_FILE + f".bak_{int(time.time())}")
                except Exception as e:
                    pass # Ignoruj błąd zmiany nazwy, jeśli plik jest zablokowany
                with open(JSON_FILE, 'w', encoding='utf-8') as f:
                    json.dump({}, f) # Zapisz pusty plik
                macros = {}
            except Exception as e:
                Messagebox.show_error(f"Nieoczekiwany błąd odczytu pliku {JSON_FILE}.\n{e}", "Błąd pliku")
                return {} # Zwróć pusty słownik
        
        elif os.path.exists(OLD_CSV_FILE):
            # --- 2. Konwertuj stary plik CSV (jeśli JSON nie istnieje) ---
            Messagebox.show_info(
                f"Wykryto stary plik {os.path.basename(OLD_CSV_FILE)}. Trwa konwersja do nowego formatu {os.path.basename(JSON_FILE)}...",
                "Konwersja pliku konfiguracyjnego"
            )
            # Użyj starej logiki do wczytania CSV
            macros = self._load_macros_from_csv_legacy() 
            
            if macros:
                # Zapisz wczytane dane do nowego pliku JSON
                self.macros = macros # Tymczasowo ustaw, aby funkcja _save mogła zadziałać
                if self._save_macros_to_file(): # Ta funkcja zapisuje już do JSON
                    Messagebox.show_info(
                        "Konwersja zakończona pomyślnie. Nowy plik .json został utworzony. Stary plik .csv nie będzie już używany.",
                        "Konwersja udana"
                    )
                else:
                    Messagebox.show_error(f"Błąd podczas zapisywania nowego pliku {JSON_FILE} po konwersji.", "Błąd konwersji")
            else:
                 Messagebox.show_warning("Stary plik CSV został znaleziony, ale wydaje się być pusty lub uszkodzony. Tworzenie nowego pliku JSON.", "Konwersja")
                 # Stwórz pusty plik JSON, aby uniknąć ponownej próby konwersji
                 try:
                     with open(JSON_FILE, 'w', encoding='utf-8') as f:
                        json.dump({}, f)
                 except Exception: pass
                 macros = {}
        else:
            # --- 3. Stwórz nowy pusty plik JSON (jeśli żaden nie istnieje) ---
            try:
                with open(JSON_FILE, 'w', encoding='utf-8') as f:
                    json.dump({}, f)
                self.status_bar.config(text=f"Utworzono nowy plik konfiguracyjny: {JSON_FILE}")
            except Exception as e:
                Messagebox.show_error(f"Nie udało się utworzyć pliku konfiguracyjnego {JSON_FILE}.\n{e}", "Błąd pliku")
        
        return macros

    def _load_macros_from_csv_legacy(self):
        """
        Logika wczytywania makr ze starego pliku CSV (na potrzeby konwersji).
        To jest ciało starej funkcji _load_macros_from_file.
        """
        macros = {}
        if not os.path.exists(OLD_CSV_FILE):
            # Nie twórz już pliku CSV, jeśli go nie ma. 
            # Nowy plik JSON zostanie utworzony przez _load_macros_from_file
            return macros

        try:
            with open(OLD_CSV_FILE, 'r', newline='', encoding='utf-8') as f:
                header_line = f.readline()
                if not header_line: return macros
                header = [h.strip() for h in header_line.split(',')]
                f.seek(0)

                if 'akcje' in header and 'hotkey' in header:
                    reader = csv.DictReader(f)
                    for row in reader:
                        name = row.get('nazwa_makra')
                        hotkey = row.get('hotkey', HOTKEY_NONE)
                        actions_json = row.get('akcje', '[]')
                        if name:
                            try:
                                actions = json.loads(actions_json)
                                if actions and isinstance(actions[0], str):
                                    raise TypeError("Old string format inside JSON detected")
                            except (json.JSONDecodeError, TypeError):
                                actions_list = json.loads(actions_json) if isinstance(actions_json, str) else []
                                actions = self._convert_old_actions(actions_list)
                            macros[name] = {'actions': actions, 'hotkey': hotkey}
                else:
                    # Logika dla bardzo starego formatu CSV
                    reader = csv.reader(f)
                    try:
                        next(reader) # Pomiń nagłówek
                    except StopIteration:
                        return macros # Pusty plik
                    
                    for row in reader:
                        if not row: continue
                        name = row[0].strip()
                        raw_actions = [value.strip() for value in row[1:]]
                        converted_actions = self._convert_old_actions(raw_actions)
                        macros[name] = {'actions': converted_actions, 'hotkey': HOTKEY_NONE}
                    
                    Messagebox.show_info(
                        "Wykryto bardzo stary format pliku `macros.csv`. Został on wczytany i zostanie przekonwertowany do formatu .json.",
                        "Aktualizacja formatu pliku"
                    )

        except Exception as e:
             Messagebox.show_error(f"Błąd odczytu lub przetwarzania pliku {OLD_CSV_FILE}.\nPlik może być uszkodzony.\nBłąd: {e}", "Błąd pliku")
        return macros

    def _convert_old_actions(self, raw_actions):
        converted_actions = []
        marker_map = {
            TAB_MARKER: ('key', 'tab'), ENTER_MARKER: ('key', 'enter'),
            SHIFT_TAB_MARKER: ('key', 'shift+tab'), CTRL_TAB_MARKER: ('key', 'ctrl+tab'),
            UP_MARKER: ('key', 'up'), DOWN_MARKER: ('key', 'down'), 
            LEFT_MARKER: ('key', 'left'), RIGHT_MARKER: ('key', 'right')
        }
        for act in raw_actions:
            if act in marker_map:
                action_type, action_value = marker_map[act]
                converted_actions.append({'type': action_type, 'value': action_value})
            else:
                converted_actions.append({'type': 'text', 'value': act})
        return converted_actions

    # --- NOWA LOGIKA ZAPISU DO PLIKU ---
    def _save_macros_to_file(self):
        """
        Zapisuje bieżący stan słownika self.macros do pliku JSON
        z formatowaniem "pretty-print" (wcięcia).
        """
        try:
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                # Użyj indent=4 dla "pretty-printing"
                # ensure_ascii=False pozwala poprawnie zapisywać polskie znaki
                json.dump(self.macros, f, indent=4, ensure_ascii=False) 
            return True
        except Exception as e:
            Messagebox.show_error(f"Błąd zapisu do pliku {JSON_FILE}.\n{e}", "Błąd zapisu")
            return False

#
# --- Reszta klas (AddStructuredMacroDialog, AddEditMacroDialog, ImageDialog) ---
# --- oraz funkcja (parse_structured_macro) pozostają BEZ ZMIAN ---
#

class AddStructuredMacroDialog(ttk.Toplevel):
    def __init__(self, parent, name="", initial_data=None):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Edytuj Makro Strukturalne" if name else "Dodaj Nowe Makro Strukturalne")
        self.geometry("800x650") # Zmniejszono wysokość po usunięciu pól
        
        self.result = None
        self.price_modifier_rows = []

        self._setup_widgets()

        if name:
            self.entries['macro_name'].insert(0, name)
        if initial_data:
            self._populate_form(initial_data)

    def _setup_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=BOTH, expand=True)
        
        # --- Pola główne ---
        fields_frame = ttk.Frame(main_frame)
        fields_frame.grid(row=0, column=0, columnspan=2, sticky=EW, padx=5)
        self.entries = {}
        # Usunięto VAT i Minimalny stan z formularza
        fields = {
            "Nazwa makra": "macro_name", "Plik wynikowy": "result_file", "Para walutowa": "currency_pair",
            "Maksymalny stan": "max_stock", "Grupa cenowa dla CSV": "price_group_csv", 
            "Nazwa kolumny per SKU": "per_sku_name", "Kolumna do mapowania": "merge_on_column"
        }
        for i, (label_text, key) in enumerate(fields.items()):
            ttk.Label(fields_frame, text=label_text + ":").grid(row=i, column=0, sticky=W, padx=5, pady=5)
            entry = ttk.Entry(fields_frame, width=60)
            entry.grid(row=i, column=1, sticky=EW, padx=5, pady=5)
            self.entries[key] = entry
        fields_frame.columnconfigure(1, weight=1)

        # --- Modyfikatory cen ---
        pm_frame = ttk.Labelframe(main_frame, text="Modyfikatory Cen", padding=10)
        pm_frame.grid(row=1, column=0, columnspan=2, sticky=EW, pady=15, padx=5)
        pm_frame.columnconfigure(0, weight=1); pm_frame.columnconfigure(1, weight=1)
        pm_frame.columnconfigure(2, weight=1); pm_frame.columnconfigure(3, weight=1)

        headers = ["Przedział (do)", "Mnożnik", "Kwota dodana", "Wysyłka"]
        for i, header in enumerate(headers):
            ttk.Label(pm_frame, text=header, font=("-weight bold")).grid(row=0, column=i, padx=5, pady=5)

        self.pm_container = ttk.Frame(pm_frame)
        self.pm_container.grid(row=1, column=0, columnspan=4, sticky=EW)

        for _ in range(4): self._add_price_modifier_row()
        self._add_price_modifier_row(is_last=True)

        # --- Przyciski dolne ---
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(bottom_frame, text="Zapisz", command=self.on_save, bootstyle="success").pack(side=RIGHT, padx=5)
        ttk.Button(bottom_frame, text="Anuluj", command=self.destroy, bootstyle="secondary").pack(side=RIGHT)

    def _add_price_modifier_row(self, is_last=False):
        row_frame = ttk.Frame(self.pm_container)
        row_frame.pack(fill=X, pady=2)
        row_entries = {}
        
        if is_last:
            ttk.Label(row_frame, text="Powyżej poprzedniego").pack(side=LEFT, padx=5, expand=True, fill=X)
            row_entries['interval'] = None 
        else:
            entry_interval = ttk.Entry(row_frame); entry_interval.pack(side=LEFT, padx=5, expand=True, fill=X)
            row_entries['interval'] = entry_interval

        entry_multiplier = ttk.Entry(row_frame); entry_multiplier.pack(side=LEFT, padx=5, expand=True, fill=X)
        row_entries['multiplier'] = entry_multiplier
        entry_added = ttk.Entry(row_frame); entry_added.pack(side=LEFT, padx=5, expand=True, fill=X)
        row_entries['added'] = entry_added
        entry_ship = ttk.Entry(row_frame); entry_ship.pack(side=LEFT, padx=5, expand=True, fill=X)
        row_entries['ship'] = entry_ship
        
        self.price_modifier_rows.append(row_entries)

    def _populate_form(self, data):
        """Uzupełnia formularz danymi z parsowanego makra."""
        for key, entry in self.entries.items():
            if key != 'macro_name' and key in data:
                entry.insert(0, data[key])
        
        pm_data = data.get('price_modifiers', [])
        for i, row_widgets in enumerate(self.price_modifier_rows):
            if i < len(pm_data):
                item_data = pm_data[i]
                if row_widgets['interval']:
                    row_widgets['interval'].insert(0, item_data.get('interval', ''))
                row_widgets['multiplier'].insert(0, item_data.get('multiplier', ''))
                row_widgets['added'].insert(0, item_data.get('added', ''))
                if 'ship' in item_data:
                    row_widgets['ship'].insert(0, item_data.get('ship', ''))

    def on_save(self):
        macro_name = self.entries['macro_name'].get().strip()
        if not macro_name:
            Messagebox.show_error("Nazwa makra nie może być pusta.", "Brak nazwy", parent=self)
            return

        actions = []

        def add_text(key):
            value = self.entries[key].get().strip()
            if value or value == "0":
                actions.append({'type': 'text', 'value': value})
        
        def add_key(key_name): actions.append({'type': 'key', 'value': key_name})

        # Sekwencja generowania akcji
        add_text('result_file'); add_key('tab'); add_text('currency_pair'); add_key('tab'); add_key('tab'); add_key('tab')
        add_text('max_stock'); add_key('tab'); add_text('price_group_csv'); add_key('tab'); add_key('tab')
        
        for i, row in enumerate(self.price_modifier_rows):
            is_last_row = (i == len(self.price_modifier_rows) - 1)
            
            # Interval
            if not is_last_row:
                if row['interval']:
                    val = row['interval'].get().strip()
                    if val or val == "0": actions.append({'type': 'text', 'value': val})
                add_key('tab')
            
            # Multiplier
            val = row['multiplier'].get().strip()
            if val or val == "0": actions.append({'type': 'text', 'value': val})
            add_key('tab')
            
            # Amount added
            val = row['added'].get().strip()
            if val or val == "0": actions.append({'type': 'text', 'value': val})
            add_key('tab')
            
            # Ship
            val = row['ship'].get().strip()
            if val or val == "0": actions.append({'type': 'text', 'value': val})
            
            add_key('tab')

        add_key('enter'); add_key('shift+tab'); add_key('shift+tab')
        add_text('per_sku_name'); add_key('tab'); add_text('merge_on_column')

        self.result = (macro_name, {'actions': actions, 'hotkey': HOTKEY_NONE})
        self.destroy()

    def show(self):
        self.deiconify()
        self.wait_window()


class AddEditMacroDialog(ttk.Toplevel):
    def __init__(self, parent, name="", data=None):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Edytuj Makro" if name else "Dodaj Nowe Makro (Standardowe)")
        self.geometry("750x750")
        
        self.result = None
        self.tree_item_data = {}

        self._setup_widgets(name, data or {})

    def _setup_widgets(self, name, data):
        actions = data.get('actions', [])
        hotkey = data.get('hotkey', HOTKEY_NONE)

        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=BOTH, expand=True)

        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=X, pady=(0, 15))
        
        ttk.Label(top_frame, text="Nazwa Makra:").grid(row=0, column=0, sticky=W, padx=(0,10))
        self.name_entry = ttk.Entry(top_frame)
        self.name_entry.grid(row=0, column=1, sticky=EW, columnspan=2)
        self.name_entry.insert(0, name)

        ttk.Label(top_frame, text="Skrót klawiszowy (indywidualny):").grid(row=1, column=0, sticky=W, pady=(10,0))
        
        hotkey_frame = ttk.Frame(top_frame)
        hotkey_frame.grid(row=1, column=1, sticky=EW, pady=(10,0), columnspan=2)

        self.ctrl_var = ttk.BooleanVar()
        self.alt_var = ttk.BooleanVar()
        self.shift_var = ttk.BooleanVar()
        
        ttk.Checkbutton(hotkey_frame, text="Ctrl", variable=self.ctrl_var, bootstyle="round-toggle").pack(side=LEFT)
        ttk.Checkbutton(hotkey_frame, text="Alt", variable=self.alt_var, bootstyle="round-toggle").pack(side=LEFT, padx=10)
        ttk.Checkbutton(hotkey_frame, text="Shift", variable=self.shift_var, bootstyle="round-toggle").pack(side=LEFT)

        hotkey_options = [HOTKEY_NONE] + HOTKEY_BASE_KEYS
        self.hotkey_combo = ttk.Combobox(hotkey_frame, values=hotkey_options, state="readonly", width=8)
        self.hotkey_combo.pack(side=LEFT, padx=10)
        
        hotkey_parts = hotkey.split('+')
        base_key = next((p for p in hotkey_parts if p.upper() in HOTKEY_BASE_KEYS), HOTKEY_NONE)
        self.hotkey_combo.set(base_key)
        self.ctrl_var.set('Ctrl' in hotkey_parts)
        self.alt_var.set('Alt' in hotkey_parts)
        self.shift_var.set('Shift' in hotkey_parts)
        top_frame.columnconfigure(1, weight=1)

        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill=X, pady=5)

        add_btn_frame = ttk.Frame(controls_frame)
        add_btn_frame.pack(side=LEFT, fill=X, expand=True)
        
        insert_options_frame = ttk.Frame(controls_frame)
        insert_options_frame.pack(side=LEFT, padx=20)
        
        ttk.Button(controls_frame, text="Usuń zaznaczone", command=self._delete_selected_action, bootstyle="danger-outline").pack(side=RIGHT)

        ttk.Button(add_btn_frame, text="Dodaj Wartość", command=self._add_value_row, bootstyle="success-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_btn_frame, text="Dodaj Tab", command=self._add_tab_row, bootstyle="info-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_btn_frame, text="Dodaj Enter", command=self._add_enter_row, bootstyle="primary-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_btn_frame, text="Dodaj Shift+Tab", command=self._add_shift_tab_row, bootstyle="primary-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_btn_frame, text="Dodaj Ctrl+Tab", command=self._add_ctrl_tab_row, bootstyle="primary-outline").pack(side=LEFT, padx=5, pady=2)
        
        add_arrows_frame = ttk.Frame(add_btn_frame)
        add_arrows_frame.pack(side=LEFT, pady=(0, 5))
        ttk.Button(add_arrows_frame, text="↑", command=self._add_up_row, bootstyle="secondary-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_arrows_frame, text="↓", command=self._add_down_row, bootstyle="secondary-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_arrows_frame, text="←", command=self._add_left_row, bootstyle="secondary-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_arrows_frame, text="→", command=self._add_right_row, bootstyle="secondary-outline").pack(side=LEFT, padx=5, pady=2)

        add_image_frame = ttk.Frame(add_btn_frame)
        add_image_frame.pack(side=LEFT, pady=(0, 5))
        ttk.Button(add_image_frame, text="Kliknij Obrazek", command=self._add_click_image_row, bootstyle="danger-outline").pack(side=LEFT, padx=5, pady=2)
        ttk.Button(add_image_frame, text="Przewiń i Kliknij", command=self._add_scroll_click_row, bootstyle="danger-outline").pack(side=LEFT, padx=5, pady=2)
        
        self.insert_mode = ttk.StringVar(value="end")
        ttk.Radiobutton(insert_options_frame, text="Na końcu", variable=self.insert_mode, value="end").pack(anchor=W)
        ttk.Radiobutton(insert_options_frame, text="Przed zaznaczonym", variable=self.insert_mode, value="before").pack(anchor=W)
        ttk.Radiobutton(insert_options_frame, text="Po zaznaczonym", variable=self.insert_mode, value="after").pack(anchor=W)

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=10)
        
        self.action_tree = ttk.Treeview(tree_frame, columns=("type", "value"), show="headings", selectmode="browse")
        self.action_tree.heading("type", text="Typ Akcji")
        self.action_tree.heading("value", text="Wartość / Ścieżka")
        self.action_tree.column("type", width=150)
        
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.action_tree.yview)
        self.action_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.action_tree.pack(side=LEFT, fill=BOTH, expand=True)
        tree_scrollbar.pack(side=RIGHT, fill=Y)

        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=X, pady=(10, 0))
        ttk.Button(bottom_frame, text="Zapisz", command=self.on_save, bootstyle="success").pack(side=RIGHT, padx=5)
        ttk.Button(bottom_frame, text="Anuluj", command=self.destroy, bootstyle="secondary").pack(side=RIGHT)
        
        self.name_entry.focus_set()
        
        for action in actions:
            self._insert_action(action, 'end', None)

    def _insert_action(self, action_data, mode, selection):
        parent = ''
        index = 'end'

        if mode != 'end' and selection:
            if mode == 'before':
                index = self.action_tree.index(selection)
            elif mode == 'after':
                index = self.action_tree.index(selection) + 1
        
        display_type, display_value = self._get_display_values(action_data)
        
        item_id = self.action_tree.insert(parent, index, values=(display_type, display_value))
        self.tree_item_data[item_id] = action_data
        
        self.action_tree.selection_set(item_id)
        self.action_tree.focus(item_id)
        self.action_tree.see(item_id)

    def _get_display_values(self, action_data):
        action_type = action_data['type']
        value = action_data.get('value', '')
        confidence = action_data.get('confidence')

        if action_type == "text":
            return "Wpisz tekst", value
        elif action_type in ["click_image", "scroll_click"]:
            type_text = "Kliknij Obrazek" if action_type == "click_image" else "Przewiń i Kliknij"
            value_text = f"{os.path.basename(value)} (Trafność: {confidence})" if value else "Brak pliku"
            return type_text, value_text
        elif action_type == "key":
            key_map = {
                "tab": "Tabulator", "enter": "Enter", "shift+tab": "Shift + Tab", "ctrl+tab": "Ctrl + Tab",
                "up": "↑ Strzałka w górę", "down": "↓ Strzałka w dół",
                "left": "← Strzałka w lewo", "right": "→ Strzałka w prawo"
            }
            return "Naciśnij klawisz", key_map.get(value, value)
        return "", ""

    def _add_action_based_on_selection(self, action_data):
        mode = self.insert_mode.get()
        selection = self.action_tree.focus() if self.action_tree.selection() else None
        
        if mode != 'end' and not selection:
            Messagebox.show_warning("Proszę najpierw zaznaczyć element na liście, aby wstawić akcję przed lub po nim.", "Brak zaznaczenia", parent=self)
            return
            
        self._insert_action(action_data, mode, selection)

    def _add_value_row(self): self._add_action_based_on_selection({'type': 'text', 'value': ''})
    def _add_tab_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'tab'})
    def _add_enter_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'enter'})
    def _add_shift_tab_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'shift+tab'})
    def _add_ctrl_tab_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'ctrl+tab'})
    def _add_up_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'up'})
    def _add_down_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'down'})
    def _add_left_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'left'})
    def _add_right_row(self): self._add_action_based_on_selection({'type': 'key', 'value': 'right'})

    def _add_image_action(self, action_type):
        dialog = ImageDialog(self, "Wybierz obrazek i ustaw trafność")
        dialog.show()
        if dialog.result:
            filepath, confidence = dialog.result
            action_data = {'type': action_type, 'value': filepath, 'confidence': confidence}
            self._add_action_based_on_selection(action_data)

    def _add_click_image_row(self): self._add_image_action("click_image")
    def _add_scroll_click_row(self): self._add_image_action("scroll_click")

    def _delete_selected_action(self):
        selection = self.action_tree.selection()
        if not selection:
            Messagebox.show_warning("Proszę zaznaczyć akcję do usunięcia.", "Brak zaznaczenia", parent=self)
            return
        
        for item_id in selection:
            if item_id in self.tree_item_data:
                del self.tree_item_data[item_id]
            self.action_tree.delete(item_id)

    def show(self):
        self.deiconify()
        self.wait_window()

    def on_save(self):
        name = self.name_entry.get().strip()
        if not name:
            Messagebox.show_error("Nazwa makra nie może być pusta.", "Brak nazwy", parent=self)
            return
        
        actions = [self.tree_item_data[item_id] for item_id in self.action_tree.get_children()]

        hotkey_parts = []
        if self.ctrl_var.get(): hotkey_parts.append('Ctrl')
        if self.alt_var.get(): hotkey_parts.append('Alt')
        if self.shift_var.get(): hotkey_parts.append('Shift')
        
        base_key = self.hotkey_combo.get()
        if base_key != HOTKEY_NONE:
            hotkey_parts.append(base_key)
        
        hotkey_str = '+'.join(hotkey_parts) if hotkey_parts and base_key != HOTKEY_NONE else HOTKEY_NONE

        self.result = (name, {'actions': actions, 'hotkey': hotkey_str})
        self.destroy()

class ImageDialog(ttk.Toplevel):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()

        self.filepath = None
        self.result = None

        self._create_widgets()
        
    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(expand=True, fill=BOTH)

        self.path_var = ttk.StringVar(value="Nie wybrano pliku...")
        
        ttk.Label(main_frame, text="Plik obrazka:").pack(padx=10, pady=5, anchor=W)
        path_frame = ttk.Frame(main_frame)
        path_frame.pack(padx=10, fill=X)
        ttk.Entry(path_frame, textvariable=self.path_var, state="readonly").pack(side=LEFT, fill=X, expand=True)
        ttk.Button(path_frame, text="Wybierz...", command=self._select_file).pack(side=LEFT, padx=5)

        ttk.Label(main_frame, text="Poziom trafności (0.0 - 1.0):").pack(padx=10, pady=5, anchor=W)
        self.confidence_entry = ttk.Entry(main_frame, width=10)
        self.confidence_entry.insert(0, "0.8")
        self.confidence_entry.pack(padx=10, pady=(0, 10))

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Zatwierdź", command=self.apply, bootstyle="success").pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="Anuluj", command=self.destroy, bootstyle="secondary").pack(side=LEFT, padx=5)

    def _select_file(self):
        path = askopenfilename(title="Wybierz obrazek (.png)", filetypes=[("Pliki PNG", "*.png")], parent=self)
        if path:
            self.filepath = path
            self.path_var.set(os.path.basename(path))

    def show(self):
        self.deiconify()
        self.wait_window()

    def apply(self):
        if not self.filepath:
            Messagebox.show_warning("Nie wybrano pliku obrazka.", "Brak pliku", parent=self)
            return

        try:
            confidence = float(self.confidence_entry.get())
            if not (0.0 <= confidence <= 1.0):
                raise ValueError
        except ValueError:
            Messagebox.show_error("Poziom trafności musi być liczbą od 0.0 do 1.0.", "Nieprawidłowa wartość", parent=self)
            return
        
        self.result = (self.filepath, confidence)
        self.destroy()

def parse_structured_macro(actions):
    """Próbuje zinterpretować listę akcji jako makro strukturalne.
    Zwraca słownik z danymi, jeśli się powiedzie, w przeciwnym razie None."""
    
    actions = actions.copy()
    parsed_data = {'price_modifiers': []}
    
    def expect_key(key_name):
        if not actions or actions[0] != {'type': 'key', 'value': key_name}:
            return False
        actions.pop(0)
        return True

    def expect_text(data_key):
        if not actions or actions[0]['type'] != 'text':
            return True 
        parsed_data[data_key] = actions.pop(0)['value']
        return True

    # Parsowanie głównej sekwencji
    if not (expect_text('result_file') and expect_key('tab') and
            expect_text('currency_pair') and expect_key('tab') and expect_key('tab') and expect_key('tab') and
            expect_text('max_stock') and expect_key('tab') and
            expect_text('price_group_csv') and expect_key('tab') and expect_key('tab')):
        return None

    # Parsowanie modyfikatorów cen
    while True:
        if not actions: return None
        
        if actions[0] == {'type': 'key', 'value': 'enter'}:
            break
            
        pm = {}
        # Dla pierwszych 4 wierszy, spodziewamy się potencjalnego 'interval'
        if len(parsed_data['price_modifiers']) < 4:
            if actions[0]['type'] == 'text':
                pm['interval'] = actions.pop(0)['value']
            if not expect_key('tab'): return None
        
        # Dla wszystkich wierszy spodziewamy się reszty
        if actions[0]['type'] == 'text':
            pm['multiplier'] = actions.pop(0)['value']
        if not expect_key('tab'): return None

        if actions[0]['type'] == 'text':
            pm['added'] = actions.pop(0)['value']
        if not expect_key('tab'): return None

        if actions[0]['type'] == 'text':
            pm['ship'] = actions.pop(0)['value']
        
        if not expect_key('tab'): return None
        
        parsed_data['price_modifiers'].append(pm)
    
    # Parsowanie końcowej sekwencji
    if not (expect_key('enter') and expect_key('shift+tab') and expect_key('shift+tab') and
            expect_text('per_sku_name') and expect_key('tab') and
            expect_text('merge_on_column')):
        return None

    if actions:
        return None

    return parsed_data


if __name__ == "__main__":
    app = MacroApp()
    app.mainloop()