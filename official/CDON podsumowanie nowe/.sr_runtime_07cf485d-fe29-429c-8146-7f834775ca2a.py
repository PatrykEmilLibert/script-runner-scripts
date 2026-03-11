import streamlit as st
import pandas as pd
import sys
import io
import datetime
from collections import Counter, defaultdict

# --- LOGIKA STARTOWA (LAUNCHER) ---
if __name__ == "__main__":
    try:
        from streamlit.web import cli as stcli
        from streamlit import runtime
    except ImportError:
        import streamlit.cli as stcli
        runtime = None

    if not runtime or not runtime.exists():
        # Limit uploadu 2GB
        sys.argv = ["streamlit", "run", sys.argv[0], "--server.maxUploadSize=2000"]
        sys.exit(stcli.main())

# --- KONFIGURACJA STRONY ---
def setup_page():
    st.set_page_config(
        page_title="Analizator SKU Partner-Grupy",
        page_icon="📦",
        layout="wide"
    )
    
    st.markdown("""
        <style>
        .stApp { background-color: #f3f4f6; color: #1f2937; }
        h1, h2, h3 { color: #1e3a8a !important; }
        div[data-testid="stFileUploader"] {
            background-color: white;
            border: 1px dashed #cbd5e1;
            padding: 20px;
            border-radius: 12px;
            text-align: center;
        }
        div[data-testid="stMetricValue"] { font-size: 1.8rem; color: #111827; }
        .stButton button {
            background: linear-gradient(90deg, #2563eb 0%, #1e40af 100%);
            border: none;
            color: white;
            font-weight: bold;
            padding: 0.5rem 1rem;
            border-radius: 8px;
        }
        </style>
    """, unsafe_allow_html=True)

# --- STAŁE ---
GROUPS = ["SHUMEE", "GREATSTORE", "EXTRASTORE"]
ACTIVE_REAL_STATUS_VALUES = {"aktywne"}

# --- GŁÓWNA LOGIKA ---
def main():
    setup_page()

    st.title("📦 Analizator SKU: Partnerzy i Grupy (Unikalne)")
    st.markdown("""
    Analiza plików CSV pod kątem aktywności produktu w kolumnie **`real_status`**. 
    Za aktywne uznawane są tylko rekordy z wartością **`aktywne`**.
    Kolumna **SUMA** pokazuje liczbę unikalnych SKU per partner (eliminuje duplikaty między grupami).
    """)

    # Słownik do przechowywania plików: uploads['SHUMEE'] = [file1, file2...]
    uploads = {g: [] for g in GROUPS}

    # 1. INTERFEJS UPLOADU (3 KOLUMNY)
    st.markdown("### 1. Wgraj pliki CSV dla poszczególnych grup")
    cols = st.columns(3)

    for i, group in enumerate(GROUPS):
        with cols[i]:
            st.info(f"📂 **{group}**")
            files = st.file_uploader(
                f"Pliki dla {group}",
                type=['csv'],
                accept_multiple_files=True,
                key=f"upl_{group}",
                label_visibility="collapsed"
            )
            if files:
                uploads[group] = files
                st.success(f"Wgrano: {len(files)}")

    st.divider()

    # Sprawdzenie czy cokolwiek wgrano
    total_files = sum(len(uploads[g]) for g in GROUPS)

    if total_files > 0:
        if st.button("🚀 Uruchom Analizę", type="primary", use_container_width=True):
            with st.spinner("Przetwarzanie danych..."):
                
                # Struktura danych: partner_group_data[GRUPA][PARTNER] = {zbiór_sku}
                partner_group_data = {g: defaultdict(set) for g in GROUPS}
                
                errors = []
                all_partners = set() 
                
                total_skus_processed_count = 0

                # Pętla po grupach
                for group in GROUPS:
                    files = uploads[group]
                    if not files:
                        continue
                    
                    # Deduplikacja per plik/grupa
                    unique_group_skus = set()

                    for file in files:
                        try:
                            file.seek(0)
                            # Nowy format pliku: separator ';' i kolumna aktywności 'real_status'
                            df = pd.read_csv(file, sep=";", on_bad_lines="skip", dtype=str)
                            df.columns = [str(col).strip().lower() for col in df.columns]
                            
                            # Walidacja kolumn
                            if "sku" in df.columns and "real_status" in df.columns:
                                # Filtrowanie aktywnych produktów wg real_status
                                active_mask = (
                                    df["real_status"]
                                    .fillna("")
                                    .astype(str)
                                    .str.strip()
                                    .str.lower()
                                    .isin(ACTIVE_REAL_STATUS_VALUES)
                                )
                                active_skus = (
                                    df.loc[active_mask, "sku"]
                                    .dropna()
                                    .astype(str)
                                    .str.strip()
                                )
                                unique_group_skus.update(sku for sku in active_skus if sku)
                            else:
                                errors.append(f"❌ {group}/{file.name}: Brak kolumn 'sku' lub 'real_status'")
                        except Exception as e:
                            errors.append(f"❌ {group}/{file.name}: Błąd - {str(e)}")
                    
                    total_skus_processed_count += len(unique_group_skus)

                    # Rozdzielanie SKU na partnerów dla danej grupy
                    for sku in unique_group_skus:
                        if "_" in sku:
                            partner = sku.split("_", 1)[0]
                            # Dodajemy SKU do zbioru (set) zamiast zwiększać licznik
                            partner_group_data[group][partner].add(sku)
                            all_partners.add(partner)
                        else:
                            pass

                # --- GENEROWANIE RAPORTU MATRIX ---
                if not all_partners:
                    if errors:
                        st.error("Nie znaleziono danych. Sprawdź błędy poniżej.")
                        with st.expander("Raport błędów"):
                            for e in errors: st.write(e)
                    else:
                        st.warning("Nie znaleziono żadnych aktywnych SKU (wg kolumny 'real_status') z poprawnymi prefiksami.")
                    return

                # Tworzymy listę wierszy
                rows_list = []
                sorted_partners = sorted(list(all_partners))

                for partner in sorted_partners:
                    row = {'Partner': partner}
                    
                    # Zbiór wszystkich SKU tego partnera ze wszystkich grup (do deduplikacji)
                    partner_all_skus = set()
                    
                    for group in GROUPS:
                        skus_in_group = partner_group_data[group][partner]
                        count = len(skus_in_group)
                        row[group] = count
                        
                        # Dodajemy do globalnego worka partnera
                        partner_all_skus.update(skus_in_group)
                    
                    # SUMA to wielkość zbioru wszystkich unikalnych SKU partnera
                    row['SUMA (Unikalne)'] = len(partner_all_skus)
                    rows_list.append(row)

                df_matrix = pd.DataFrame(rows_list)
                
                # Sortowanie malejąco po SUMIE
                df_matrix = df_matrix.sort_values('SUMA (Unikalne)', ascending=False)
                
                # Dodanie wiersza podsumowania (TOTAL)
                # Sumujemy kolumny liczbowe
                sum_row = df_matrix.drop(columns=['Partner']).sum()
                sum_row_dict = sum_row.to_dict()
                sum_row_dict['Partner'] = 'SUMA CAŁKOWITA'
                
                # Dołączenie wiersza sumy na dole
                df_final = pd.concat([df_matrix, pd.DataFrame([sum_row_dict])], ignore_index=True)

                # --- WYŚWIETLANIE WYNIKÓW ---
                st.markdown("## 📊 Wyniki: Partnerzy vs Grupy")
                st.info("Kolumna **SUMA (Unikalne)** weryfikuje duplikaty. Jeśli ten sam SKU jest w grupie SHUMEE i GREATSTORE, zostanie policzony tylko raz w sumie.")
                
                m1, m2 = st.columns(2)
                with m1: st.container(border=True).metric("Przetworzone wiersze (Suma grup)", total_skus_processed_count)
                with m2: st.container(border=True).metric("Liczba Partnerów", len(all_partners))

                # Tabela
                st.dataframe(df_final, use_container_width=True, hide_index=True)

                if errors:
                    with st.expander("⚠️ Wykryto błędy w niektórych plikach", expanded=False):
                        for e in errors: st.write(e)

                # --- POBIERANIE ---
                st.markdown("## 📥 Pobierz Raport")
                
                current_date = datetime.datetime.now().strftime("%Y-%m-%d")
                filename = f"podsumowanie_CDON_{current_date}.csv"
                
                csv_data = df_final.to_csv(sep=";", index=False, encoding="utf-8-sig")
                
                st.download_button(
                    label=f"💾 Pobierz Podsumowanie ({filename})",
                    data=csv_data,
                    file_name=filename,
                    mime="text/csv",
                    type="primary",
                    use_container_width=True
                )

    else:
        st.info("Wgraj pliki CSV do sekcji powyżej, aby rozpocząć.")

if __name__ == "__main__":
    main()