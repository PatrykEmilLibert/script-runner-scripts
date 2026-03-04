import flet as ft
import pandas as pd
import datetime
import os
from tkinter import filedialog
import tkinter as tk

# --- STAŁE ---
GROUPS = ["SHUMEE", "GREATSTORE", "EXTRASTORE"]
COUNTRIES = ["FR", "DE", "IT", "ES"]

def main(page: ft.Page):
    # --- KONFIGURACJA STRONY ---
    page.title = "Analizator SKU Multi-Market"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 20
    page.window_width = 1200
    page.window_height = 900
    # Kolory inspirowane poprzednim CSS - używamy stringa "indigo" zamiast ft.colors.INDIGO
    page.theme = ft.Theme(color_scheme_seed="indigo")
    
    # Przechowywanie ścieżek do plików: uploads[group][country] = lista ścieżek jako stringi
    uploads = {g: {c: [] for c in COUNTRIES} for g in GROUPS}
    
    # Elementy UI, które będziemy aktualizować
    results_container = ft.Column(scroll=ft.ScrollMode.AUTO)
    progress_ring = ft.ProgressBar(visible=False)
    # Używamy HEX code dla GREY_700 (#616161)
    status_text = ft.Text("", color="#616161")
    export_row = ft.Row(visible=False)

    # --- LOGIKA BIZNESOWA (PANDAS) ---
    def process_data(e):
        # Reset UI
        results_container.controls.clear()
        progress_ring.visible = True
        status_text.value = "Przetwarzanie plików..."
        export_row.visible = False
        page.update()

        all_raw_data = []
        errors = []
        total_files_count = sum(len(uploads[g][c]) for g in GROUPS for c in COUNTRIES)

        if total_files_count == 0:
            progress_ring.visible = False
            status_text.value = "Nie wybrano żadnych plików!"
            status_text.color = "red"
            page.update()
            return

        try:
            for group in GROUPS:
                for country in COUNTRIES:
                    files = uploads[group][country]
                    if not files:
                        continue

                    for file_path in files:
                        try:
                            # Czytamy CSV
                            df = pd.read_csv(file_path, sep=";", on_bad_lines="skip", dtype=str)

                            if "SKU" in df.columns and "STATUS" in df.columns:
                                # Filtrujemy ONLINE
                                online_skus = df[df["STATUS"] == "ONLINE"]["SKU"].dropna().astype(str)

                                for sku in online_skus:
                                    # Wyznaczanie partnera
                                    partner = sku.split("_", 1)[0] if "_" in sku else "INNE"
                                    
                                    all_raw_data.append({
                                        "Grupa": group,
                                        "Kraj": country,
                                        "SKU": sku,
                                        "Partner": partner
                                    })
                            else:
                                file_name = os.path.basename(file_path)
                                errors.append(f"❌ {group}/{country}/{file_name}: Brak kolumn SKU/STATUS")
                        except Exception as ex:
                            file_name = os.path.basename(file_path) if file_path else "unknown"
                            errors.append(f"❌ {group}/{country}/{file_name}: Błąd - {str(ex)}")

            if not all_raw_data:
                raise Exception("Nie znaleziono żadnych SKU Online w wybranych plikach.")

            # --- PRZETWARZANIE DANYCH ---
            df_all = pd.DataFrame(all_raw_data)

            # 1. Unikalność per kraj
            df_unique_countries = df_all.drop_duplicates(subset=['Kraj', 'SKU'])
            df_counts = df_unique_countries.groupby(['Kraj', 'Partner']).size().reset_index(name='Ilość')
            
            # Pivot Matrix
            df_matrix = df_counts.pivot(index='Partner', columns='Kraj', values='Ilość').fillna(0).astype(int)
            
            # Wymuszenie kolejności kolumn
            df_matrix = df_matrix.reindex(columns=COUNTRIES, fill_value=0)

            # 2. Unikalność globalna (SUMA)
            df_partner_unique = df_all.drop_duplicates(subset=['Partner', 'SKU'])
            partner_global_counts = df_partner_unique.groupby('Partner').size()

            df_matrix['SUMA'] = df_matrix.index.map(partner_global_counts).fillna(0).astype(int)
            df_matrix = df_matrix.sort_values('SUMA', ascending=False)

            # Wiersz podsumowania
            total_row = df_matrix.sum().to_frame().T
            total_row.index = ['SUMA CAŁKOWITA']
            df_final_matrix = pd.concat([df_matrix, total_row])

            # Reset indeksu, żeby Partner był kolumną, a nie indeksem (łatwiej wyświetlić)
            df_display = df_final_matrix.reset_index().rename(columns={'index': 'Partner'})

            # --- GENEROWANIE TABELI W FLET ---
            # Kolumny
            dt_columns = [ft.DataColumn(ft.Text(col, weight=ft.FontWeight.BOLD)) for col in df_display.columns]
            
            # Wiersze
            dt_rows = []
            for _, row in df_display.iterrows():
                cells = []
                for col in df_display.columns:
                    val = row[col]
                    # Pogrubienie wiersza sumy
                    is_total_row = row['Partner'] == 'SUMA CAŁKOWITA'
                    weight = ft.FontWeight.BOLD if is_total_row or col == 'Partner' or col == 'SUMA' else ft.FontWeight.NORMAL
                    # Kolory: BLUE_900 -> #0d47a1, BLACK -> black
                    color = "#0d47a1" if is_total_row else "black"
                    
                    cells.append(ft.DataCell(ft.Text(str(val), weight=weight, color=color)))
                dt_rows.append(ft.DataRow(cells=cells))

            data_table = ft.DataTable(
                columns=dt_columns,
                rows=dt_rows,
                border=ft.border.all(1, "#e0e0e0"), # GREY_300
                vertical_lines=ft.border.BorderSide(1, "#eeeeee"), # GREY_200
                horizontal_lines=ft.border.BorderSide(1, "#eeeeee"), # GREY_200
                heading_row_color="#e3f2fd", # BLUE_50
            )

            results_container.controls.append(data_table)
            
            # Raport błędów
            if errors:
                err_col = ft.Column()
                err_col.controls.append(ft.Text("⚠️ Raport błędów:", color="red", weight=ft.FontWeight.BOLD))
                for err in errors:
                    err_col.controls.append(ft.Text(err, size=12, color="#ef5350")) # RED_400
                results_container.controls.append(err_col)

            status_text.value = f"Przetworzono pomyślnie! Unikalnych SKU w sumie: {len(df_partner_unique)}"
            status_text.color = "#388e3c" # GREEN_700
            
            # Przygotowanie eksportu
            def save_csv(e):
                try:
                    output_path = f"podsumowanie_ManoMano_{datetime.datetime.now().strftime('%Y-%m-%d')}.csv"
                    df_final_matrix.to_csv(output_path, sep=";", encoding="utf-8-sig")
                    page.show_snack_bar(ft.SnackBar(ft.Text(f"Zapisano plik: {output_path}"), bgcolor="green"))
                except Exception as ex:
                    page.show_snack_bar(ft.SnackBar(ft.Text(f"Błąd zapisu: {str(ex)}"), bgcolor="red"))
            
            # Odświeżenie przycisku eksportu
            export_row.controls.clear()
            export_row.controls.append(
                ft.ElevatedButton(
                    "💾 Pobierz Raport CSV",
                    on_click=save_csv,
                    style=ft.ButtonStyle(color="white", bgcolor="#43a047") # GREEN_600
                )
            )
            export_row.visible = True

        except Exception as e:
            status_text.value = f"Wystąpił błąd krytyczny: {str(e)}"
            status_text.color = "#d32f2f" # RED_700
            print(e) # Log do konsoli

        finally:
            progress_ring.visible = False
            page.update()

    # --- BUDOWANIE UI UPLOADU ---
    
    # Funkcja pomocnicza do tworzenia "kafelka" uploadu
    def create_upload_tile(group_name, country_code):
        selected_files_text = ft.Text("Brak plików", size=12, color="grey")
        
        def select_files(e):
            # Ukrycie głównego okna tkinter
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            # Otwarcie dialogu wyboru plików
            file_paths = filedialog.askopenfilenames(
                title=f"Wybierz pliki CSV dla {group_name} - {country_code}",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            root.destroy()
            
            if file_paths:
                uploads[group_name][country_code] = list(file_paths)
                count = len(file_paths)
                selected_files_text.value = f"Wybrano: {count} plik(ów)"
                selected_files_text.color = "blue"
                selected_files_text.weight = ft.FontWeight.BOLD
            else:
                uploads[group_name][country_code] = []
                selected_files_text.value = "Brak plików"
                selected_files_text.color = "grey"
                selected_files_text.weight = ft.FontWeight.NORMAL
            
            page.update()
        
        return ft.Container(
            content=ft.Column([
                ft.Text(country_code, weight=ft.FontWeight.BOLD, size=16),
                ft.ElevatedButton(
                    "Wybierz pliki",
                    on_click=select_files,
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=5))
                ),
                selected_files_text
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            padding=10,
            border=ft.border.all(1, "#e0e0e0"), # GREY_300
            border_radius=8,
            bgcolor="white",
            width=180,
            height=130
        )

    # Tworzenie zakładek dla Grup - używamy prostego rozwiązania z przyciskami
    current_group_index = [0]  # Używamy listy jako mutable reference
    upload_container = ft.Container(padding=20, bgcolor="#fafafa")
    
    def update_upload_view():
        group = GROUPS[current_group_index[0]]
        row_controls = [create_upload_tile(group, country) for country in COUNTRIES]
        upload_container.content = ft.Row(row_controls, wrap=True, alignment=ft.MainAxisAlignment.START, spacing=20)
        page.update()
    
    def change_group(index):
        def handler(e):
            current_group_index[0] = index
            # Aktualizacja kolorów przycisków
            for i, btn in enumerate(group_buttons):
                if i == index:
                    btn.style = ft.ButtonStyle(bgcolor="#3949ab", color="white")
                else:
                    btn.style = ft.ButtonStyle(bgcolor="#e0e0e0", color="black")
            update_upload_view()
        return handler
    
    # Przyciski wyboru grupy
    group_buttons = []
    for i, group in enumerate(GROUPS):
        btn = ft.ElevatedButton(
            f"🏢 {group}",
            on_click=change_group(i),
            style=ft.ButtonStyle(
                bgcolor="#3949ab" if i == 0 else "#e0e0e0",
                color="white" if i == 0 else "black"
            )
        )
        group_buttons.append(btn)
    
    group_selector = ft.Row(group_buttons, spacing=10)
    
    # Inicjalizacja widoku dla pierwszej grupy
    update_upload_view()
    
    tabs_section = ft.Container(
        content=ft.Column([group_selector, upload_container]),
        height=250
    )

    # --- GŁÓWNY UKŁAD STRONY ---
    
    header = ft.Container(
        content=ft.Column([
            ft.Text("🌍 Analizator SKU Multi-Market", size=30, weight=ft.FontWeight.BOLD, color="#1a237e"), # INDIGO_900
            ft.Text("Wgraj pliki CSV, aby wygenerować raport Matrix. Algorytm usuwa duplikaty i sumuje unikalne SKU.", size=14, color="#616161") # GREY_700
        ]),
        padding=ft.padding.only(bottom=20)
    )

    action_bar = ft.Row([
        ft.ElevatedButton(
            "🚀 Generuj Raport Matrix", 
            on_click=process_data,
            style=ft.ButtonStyle(
                color="white", 
                bgcolor="#3949ab", # INDIGO_600
                padding=15,
                shape=ft.RoundedRectangleBorder(radius=8)
            ),
            height=45
        ),
        progress_ring,
        status_text
    ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER)

    # Dodawanie wszystkiego do strony
    page.add(
        header,
        tabs_section,
        ft.Divider(),
        action_bar,
        ft.Divider(),
        export_row,
        results_container
    )

if __name__ == "__main__":
    ft.app(target=main)