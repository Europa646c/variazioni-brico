import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Gestione Variazioni Prezzo - BricoIO", layout="wide")

st.title("🏷️ Gestione Variazioni Prezzo & Giacenze - BricoIO")
st.markdown("Carica i due file Excel per elaborare le variazioni prezzo escludendo automaticamente i passaggi sicuri.")

# Sezione caricamento file
col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    file_var = st.file_uploader("1. File Variazioni Future (es. VAR_FUTURE...xlsx)", type=["xlsx", "xls"])
with col_upload2:
    file_stock = st.file_uploader("2. File Stock / Facing (es. StockGrottammare.xlsx)", type=["xlsx", "xls"])

if file_var and file_stock:
    try:
        with st.spinner("Elaborazione dati in corso..."):
            df_var = pd.read_excel(file_var)
            
            # Lettura fogli stock
            excel_stock = pd.ExcelFile(file_stock)
            sheet_names = excel_stock.sheet_names
            
            # Giacenze
            target_stock = "Stock" if "Stock" in sheet_names else sheet_names[0]
            df_stock = pd.read_excel(excel_stock, sheet_name=target_stock)
            
            # Facing (Zkit)
            df_zkit = None
            if "Zkit" in sheet_names:
                df_zkit = pd.read_excel(excel_stock, sheet_name="Zkit")
            
            # Normalizzazione nomi colonne
            df_var.columns = [str(c).strip() for c in df_var.columns]
            df_stock.columns = [str(c).strip() for c in df_stock.columns]
            if df_zkit is not None:
                df_zkit.columns = [str(c).strip() for c in df_zkit.columns]

            # Riconoscimento colonne chiave
            col_art_var = next((c for c in df_var.columns if "art" in c.lower() or "cod" in c.lower()), df_var.columns[0])
            col_art_stock = next((c for c in df_stock.columns if "art" in c.lower() or "cod" in c.lower()), df_stock.columns[0])
            col_qta = next((c for c in df_stock.columns if "qta" in c.lower() or "giac" in c.lower() or "stock" in c.lower()), None)
            
            # Conversione codice articolo in stringa pulita
            df_var[col_art_var] = df_var[col_art_var].astype(str).str.strip()
            df_stock[col_art_stock] = df_stock[col_art_stock].astype(str).str.strip()
            
            # Calcolo giacenza per articolo
            if col_qta:
                df_stock[col_qta] = pd.to_numeric(df_stock[col_qta], errors='coerce').fillna(0)
                stock_map = df_stock.groupby(col_art_stock)[col_qta].sum().to_dict()
            else:
                stock_map = {}

            # Calcolo presenza a scaffale (Facing / Zkit)
            facing_set = set()
            if df_zkit is not None:
                col_art_zkit = next((c for c in df_zkit.columns if "art" in c.lower() or "cod" in c.lower()), df_zkit.columns[0])
                col_mod = next((c for c in df_zkit.columns if "mod" in c.lower() or "facing" in c.lower()), None)
                df_zkit[col_art_zkit] = df_zkit[col_art_zkit].astype(str).str.strip()
                if col_mod:
                    valid_zkit = df_zkit[df_zkit[col_mod].notna() & (df_zkit[col_mod].astype(str).str.strip() != "")]
                    facing_set = set(valid_zkit[col_art_zkit].unique())
                else:
                    facing_set = set(df_zkit[col_art_zkit].unique())

            # Assegnazione stock e facing a df_var
            df_var["Giacenza"] = df_var[col_art_var].map(stock_map).fillna(0)
            df_var["In_Facing"] = df_var[col_art_var].apply(lambda x: "SI" if x in facing_set else "NO")

            # --- NUOVA SCREMATURA (PASSAGGIO SICURO) ---
            # 1. Filtro Tipo (se presente, tenere solo Prz. Vendita - base)
            col_tipo = next((c for c in df_var.columns if "tipo" in c.lower()), None)
            mask_tipo_ok = True
            if col_tipo:
                mask_tipo_ok = df_var[col_tipo].astype(str).str.contains("Prz. Vendita - base", case=False, na=False)

            # 2. Filtro Giacenza / Facing (esclude stock 0 e facing NO)
            mask_presenza = (df_var["Giacenza"] > 0) | (df_var["In_Facing"] == "SI")

            # 3. Filtro Variazione Prezzo Vendita (Esclude variazioni solo di costo base senza modifica prezzi)
            col_p_base = next((c for c in df_var.columns if c.upper() == "PREZZO BASE"), None)
            col_p_nuovo = next((c for c in df_var.columns if c.upper() == "PREZZO NUOVO"), None)
            
            if col_p_base and col_p_nuovo:
                val_p_base = pd.to_numeric(df_var[col_p_base].astype(str).str.replace(',', '.'), errors='coerce')
                val_p_nuovo = pd.to_numeric(df_var[col_p_nuovo].astype(str).str.replace(',', '.'), errors='coerce')
                
                # Ha una variazione di prezzo di vendita valida se:
                # - Il nuovo prezzo non è vuoto/nullo
                # - Il nuovo prezzo è diverso dal prezzo base
                mask_prezzo_modificato = val_p_nuovo.notna() & (val_p_nuovo != val_p_base) & (val_p_nuovo > 0)
            else:
                mask_prezzo_modificato = True

            # Righe da lavorare (escluse quelle di "Passaggio Sicuro")
            df_lavoro = df_var[mask_tipo_ok & mask_presenza & mask_prezzo_modificato].copy()

            # Calcolo Urgenze (Data Inizio <= 3 giorni)
            col_data = next((c for c in df_lavoro.columns if "data" in c.lower() and "inizio" in c.lower()), None)
            if not col_data:
                col_data = next((c for c in df_lavoro.columns if "data" in c.lower()), None)

            if col_data:
                df_lavoro[col_data] = pd.to_datetime(df_lavoro[col_data], errors='coerce')
                oggi = pd.to_datetime(datetime.now().date())
                df_lavoro["Giorni_Rimanenti"] = (df_lavoro[col_data] - oggi).dt.days
                df_lavoro["Urgente"] = df_lavoro["Giorni_Rimanenti"].apply(
                    lambda g: "🚨 URGENTE (<=3gg)" if pd.notna(g) and g <= 3 else "Regolare"
                )
            else:
                df_lavoro["Giorni_Rimanenti"] = None
                df_lavoro["Urgente"] = "Regolare"

            # Individuazione colonna Fornitore
            col_forn = next((c for c in df_lavoro.columns if "forn" in c.lower() or "ragione" in c.lower()), None)

        # Dashboard Metriche
        st.success("File elaborati con successo!")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Totale Righe File", len(df_var))
        m2.metric("Passaggi Sicuri Esclusi", len(df_var) - len(df_lavoro))
        m3.metric("Da Gestire (Scaffale/Giacenza)", len(df_lavoro))
        urgenze_tot = (df_lavoro["Urgente"] == "🚨 URGENTE (<=3gg)").sum()
        m4.metric("🚨 Alert Imminenti (<= 3gg)", urgenze_tot)

        # Tabella di Riepilogo per Fornitore
        if col_forn:
            st.subheader("📊 Riepilogo per Fornitore")
            recap_forn = df_lavoro.groupby(col_forn).agg(
                Totale_Articoli=(col_art_var, 'count'),
                Articoli_Urgenti=('Urgente', lambda x: (x == "🚨 URGENTE (<=3gg)").sum()),
                Giacenza_Totale=('Giacenza', 'sum')
            ).reset_index().sort_values(by=["Articoli_Urgenti", "Totale_Articoli"], ascending=False)
            
            st.dataframe(recap_forn, use_container_width=True)

        # Dettaglio Articoli
        st.subheader("📋 Dettaglio Articoli da Gestire")
        st.dataframe(df_lavoro, use_container_width=True)

        # Esportazione Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_lavoro.to_excel(writer, sheet_name="Da_Gestire", index=False)
            if col_forn:
                recap_forn.to_excel(writer, sheet_name="Riepilogo_Fornitori", index=False)
        output.seek(0)

        st.download_button(
            label="📥 Scarica Report Excel Filtrato",
            data=output,
            file_name=f"Variazioni_Prezzo_Gestione_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione dei file: {e}")
