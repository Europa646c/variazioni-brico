import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Gestione Variazioni Brico", layout="wide", page_icon="🏷️")

st.title("🏷️ Dashboard Variazioni Prezzo - Brico IO Grottammare")
st.markdown("Incrocio automatico tra listino futuro e giacenze/facing con filtro **Passaggio Sicuro**.")

# 1. Caricamento File
col1, col2 = st.columns(2)
with col1:
    var_file = st.file_uploader("1. File Variazioni Future (es. VAR_FUTURE...xlsx)", type=["xlsx", "xls"])
with col2:
    stock_file = st.file_uploader("2. File Giacenze / Facing (Zkit)", type=["xlsx", "xls"])

def clean_code(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

if var_file and stock_file:
    with st.spinner("Elaborazione dati in corso..."):
        try:
            # Caricamento variazioni future
            df_var = pd.read_excel(var_file)
            
            # Caricamento file stock/facing (Zkit)
            excel_stock = pd.ExcelFile(stock_file)
            sheet_names = excel_stock.sheet_names
            
            # Trova il foglio giusto (Zkit o primo foglio)
            target_sheet = next((s for s in sheet_names if "zkit" in s.lower()), sheet_names[0])
            df_stock = pd.read_excel(stock_file, sheet_name=target_sheet)

            # Normalizzazione colonne
            df_var.columns = [str(c).strip().upper() for c in df_var.columns]
            df_stock.columns = [str(c).strip().upper() for c in df_stock.columns]

            # Riconoscimento colonne codice articolo
            col_art_var = next((c for c in df_var.columns if "ART" in c or "COD" in c), df_var.columns[0])
            col_art_stock = next((c for c in df_stock.columns if "ART" in c or "COD" in c), df_stock.columns[0])

            # Pulizia codici per il merge
            df_var['COD_KEY'] = df_var[col_art_var].apply(clean_code)
            df_stock['COD_KEY'] = df_stock[col_art_stock].apply(clean_code)

            # Riconoscimento colonne Stock e Facing
            col_giacenza = next((c for c in df_stock.columns if "GIAC" in c or "QTA" in c or "QUANT" in c or "STOCK" in c), None)
            col_facing = next((c for c in df_stock.columns if "FACING" in c or "MODULO" in c), None)

            # Calcolo Giacenza e In_Facing
            stock_summary = pd.DataFrame({'COD_KEY': df_stock['COD_KEY'].unique()})
            
            if col_giacenza:
                giac_agg = df_stock.groupby('COD_KEY')[col_giacenza].sum().reset_index()
                giac_agg.rename(columns={col_giacenza: 'GIACENZA'}, inplace=True)
                stock_summary = stock_summary.merge(giac_agg, on='COD_KEY', how='left')
            else:
                stock_summary['GIACENZA'] = 0

            if col_facing:
                facing_items = df_stock[df_stock[col_facing].notna() & (df_stock[col_facing].astype(str).str.strip() != "") & (df_stock[col_facing] != 0)]['COD_KEY'].unique()
                stock_summary['IN_FACING'] = stock_summary['COD_KEY'].apply(lambda x: 'SI' if x in facing_items else 'NO')
            else:
                stock_summary['IN_FACING'] = 'NO'

            stock_summary['GIACENZA'] = stock_summary['GIACENZA'].fillna(0)

            # Unione dei due file
            merged = df_var.merge(stock_summary, on='COD_KEY', how='left')
            merged['GIACENZA'] = merged['GIACENZA'].fillna(0)
            merged['IN_FACING'] = merged['IN_FACING'].fillna('NO')

            # --- IDENTIFICAZIONE PASSAGGIO SICURO ---
            # 1. Nessuna presenza fisica: Giacenza <= 0 e non in Facing
            cond_zero_presenza = (merged['GIACENZA'] <= 0) & (merged['IN_FACING'] == 'NO')

            # 2. Variazioni solo costo: PREZZO BASE e PREZZO NUOVO entrambi vuoti o uguali
            col_prz_base = next((c for c in merged.columns if "PREZZO" in c and "BASE" in c), None)
            col_prz_nuovo = next((c for c in merged.columns if "PREZZO" in c and "NUOVO" in c), None)

            if col_prz_base and col_prz_nuovo:
                cond_no_variazione_prezzo = (
                    (merged[col_prz_base].isna() | (merged[col_prz_base].astype(str).str.strip() == "")) &
                    (merged[col_prz_nuovo].isna() | (merged[col_prz_nuovo].astype(str).str.strip() == ""))
                ) | (merged[col_prz_base] == merged[col_prz_nuovo])
            else:
                cond_no_variazione_prezzo = pd.Series(False, index=merged.index)

            # 3. Tipo variazione non di vendita (se colonna Tipo presente)
            col_tipo = next((c for c in merged.columns if "TIPO" in c), None)
            if col_tipo:
                cond_tipo_escluso = ~merged[col_tipo].astype(str).str.contains("Prz. Vendita - base", case=False, na=False)
            else:
                cond_tipo_escluso = pd.Series(False, index=merged.index)

            merged['PASSAGGIO_SICURO'] = cond_zero_presenza | cond_no_variazione_prezzo | cond_tipo_escluso

            # Calcolo giorni a scadenza / urgenza
            col_data = next((c for c in merged.columns if "DATA" in c and ("INIZIO" in c or "DECOR" in c or "PASS" in c)), None)
            today = pd.to_datetime(datetime.today().date())
            
            if col_data:
                merged['DATA_DT'] = pd.to_datetime(merged[col_data], errors='coerce')
                merged['GIORNI_RIMANENTI'] = (merged['DATA_DT'] - today).dt.days
                merged['URGENTE'] = merged['GIORNI_RIMANENTI'].apply(lambda d: '🔴 URGENTE' if pd.notna(d) and d <= 3 else '🟢 NORMALE')
            else:
                merged['GIORNI_RIMANENTI'] = None
                merged['URGENTE'] = '🟢 NORMALE'

            # Identificazione colonna Fornitore e Descrizione
            col_forn = next((c for c in merged.columns if "FORN" in c or "RAGIONE" in c), "FORNITORE")
            if col_forn not in merged.columns:
                merged[col_forn] = "Non specificato"

            col_desc = next((c for c in merged.columns if "DESC" in c), None)

            # Filtro degli articoli "Da Gestire" (esclude Passaggio Sicuro)
            df_da_gestire = merged[~merged['PASSAGGIO_SICURO']].copy()

            # --- VISUALIZZAZIONE DASHBOARD ---
            st.divider()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Righe Totali File", len(merged))
            m2.metric("Passaggio Sicuro (Esclusi)", int(merged['PASSAGGIO_SICURO'].sum()))
            m3.metric("Da Gestire", len(df_da_gestire))
            urgenti_count = int((df_da_gestire['URGENTE'] == '🔴 URGENTE').sum())
            m4.metric("Alert Imminenti (<= 3 gg)", urgenti_count, delta="Urgente" if urgenti_count > 0 else None, delta_color="inverse")

            # Grafico e riepilogo per fornitore
            st.subheader("📊 Riepilogo Articoli da Gestire per Fornitore")
            if not df_da_gestire.empty:
                recap_forn = df_da_gestire.groupby(col_forn).agg(
                    Totale_Articoli=('COD_KEY', 'count'),
                    Urgenti=('URGENTE', lambda s: (s == '🔴 URGENTE').sum())
                ).reset_index().sort_values(by=['Urgenti', 'Totale_Articoli'], ascending=[False, False])
                
                col_chart, col_tbl = st.columns([1, 1])
                with col_chart:
                    st.bar_chart(recap_forn.set_index(col_forn)['Totale_Articoli'])
                with col_tbl:
                    st.dataframe(recap_forn, use_container_width=True, hide_index=True)
            else:
                st.info("Nessun articolo rimasto da gestire: tutti i record rientrano nel Passaggio Sicuro.")

            # Tabella di dettaglio
            st.subheader("📋 Elenco Dettagliato Articoli da Gestire")
            cols_to_show = ['URGENTE', col_art_var]
            if col_desc and col_desc in df_da_gestire.columns:
                cols_to_show.append(col_desc)
            cols_to_show.extend([col_forn, 'GIACENZA', 'IN_FACING'])
            if col_prz_base and col_prz_base in df_da_gestire.columns:
                cols_to_show.append(col_prz_base)
            if col_prz_nuovo and col_prz_nuovo in df_da_gestire.columns:
                cols_to_show.append(col_prz_nuovo)
            if col_data and col_data in df_da_gestire.columns:
                cols_to_show.extend([col_data, 'GIORNI_RIMANENTI'])

            valid_cols = [c for c in cols_to_show if c in df_da_gestire.columns]
            st.dataframe(df_da_gestire[valid_cols], use_container_width=True, hide_index=True)

            # Export Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_da_gestire[valid_cols].to_excel(writer, index=False, sheet_name='Da_Gestire')
                merged[merged['PASSAGGIO_SICURO']].to_excel(writer, index=False, sheet_name='Passaggio_Sicuro')
            
            st.download_button(
                label="📥 Scarica Report Excel (Da Gestire & Passaggio Sicuro)",
                data=buffer.getvalue(),
                file_name=f"Variazioni_Prezzo_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Errore durante l'elaborazione dei file: {e}")
