import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Variazioni Prezzo Brico", layout="wide")

st.title("Gestione Variazioni Prezzo - Brico")
st.markdown("Carica i due file Excel per elaborare le variazioni e visualizzare la dashboard.")

col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    file_var = st.file_uploader("1. File Variazioni Future (es. VAR_FUTURE...xlsx)", type=["xlsx", "xls"])

with col_upload2:
    file_stock = st.file_uploader("2. File Stock / Facing (es. StockGrottammare.xlsx)", type=["xlsx", "xls"])

def clean_col_name(c):
    return str(c).strip().upper()

def trova_colonna(colonne, nomi_possibili):
    for nome in nomi_possibili:
        for c in colonne:
            if nome.upper() in str(c).strip().upper():
                return c
    return None

if file_var and file_stock:
    try:
        with st.spinner("Elaborazione dei file in corso..."):
            # 1. Caricamento Variazioni
            df_var = pd.read_excel(file_var)
            df_var.columns = [str(c).strip() for c in df_var.columns]

            # 2. Caricamento Stock e Zkit
            xl_stock = pd.ExcelFile(file_stock)
            
            # Foglio Stock
            sheet_stock = next((s for s in xl_stock.sheet_names if "STOCK" in s.upper()), xl_stock.sheet_names[0])
            df_stock_raw = pd.read_excel(xl_stock, sheet_name=sheet_stock)
            df_stock_raw.columns = [str(c).strip() for c in df_stock_raw.columns]

            # Foglio Zkit per il Facing
            sheet_zkit = next((s for s in xl_stock.sheet_names if "ZKIT" in s.upper() or "FACING" in s.upper()), None)
            df_zkit_raw = pd.read_excel(xl_stock, sheet_name=sheet_zkit) if sheet_zkit else pd.DataFrame()
            if not df_zkit_raw.empty:
                df_zkit_raw.columns = [str(c).strip() for c in df_zkit_raw.columns]

            # Riconoscimento colonne Variazioni
            col_art_var = trova_colonna(df_var.columns, ["ARTICOLO", "CODICE", "COD_ART"])
            col_forn = trova_colonna(df_var.columns, ["FORNITORE", "DESCR_FORN", "RAGIONE SOCIALE"])
            col_tipo = trova_colonna(df_var.columns, ["TIPO"])
            col_data_ini = trova_colonna(df_var.columns, ["DATA INIZIO", "DATA_INIZIO", "VALIDO DAL"])
            col_p_base = trova_colonna(df_var.columns, ["PREZZO BASE", "PRZ BASE", "PRZ. BASE"])
            col_p_nuovo = trova_colonna(df_var.columns, ["PREZZO NUOVO", "PRZ NUOVO", "PRZ. NUOVO"])
            col_c_base = trova_colonna(df_var.columns, ["COSTO BASE", "CST BASE"])
            col_c_nuovo = trova_colonna(df_var.columns, ["COSTO NUOVO", "CST NUOVO"])
            col_p_ins = trova_colonna(df_var.columns, ["P_NUOVO_INS", "PREZZO_INS"])

            # Riconoscimento colonne Stock
            col_art_stock = trova_colonna(df_stock_raw.columns, ["ARTICOLO", "CODICE", "COD_ART"])
            col_qta_stock = trova_colonna(df_stock_raw.columns, ["GIACENZA", "QTA", "DISPONIBILE", "ESISTENZA"])

            # Riconoscimento colonne Zkit (Facing)
            col_art_zkit = trova_colonna(df_zkit_raw.columns, ["ARTICOLO", "CODICE", "COD_ART"]) if not df_zkit_raw.empty else None
            col_mod_zkit = trova_colonna(df_zkit_raw.columns, ["MODULO", "FACING"]) if not df_zkit_raw.empty else None

            # Normalizzazione codici
            df_var['KEY_ART'] = df_var[col_art_var].astype(str).str.strip().str.lstrip("0")
            df_stock_raw['KEY_ART'] = df_stock_raw[col_art_stock].astype(str).str.strip().str.lstrip("0")

            # Mappa Giacenze
            df_stock_raw[col_qta_stock] = pd.to_numeric(df_stock_raw[col_qta_stock], errors='coerce').fillna(0)
            stock_map = df_stock_raw.groupby('KEY_ART')[col_qta_stock].sum().to_dict()
            df_var['Giacenza'] = df_var['KEY_ART'].map(stock_map).fillna(0)

            # Mappa Facing
            if col_art_zkit and col_mod_zkit and not df_zkit_raw.empty:
                df_zkit_raw['KEY_ART'] = df_zkit_raw[col_art_zkit].astype(str).str.strip().str.lstrip("0")
                facing_set = set(df_zkit_raw[df_zkit_raw[col_mod_zkit].notna() & (df_zkit_raw[col_mod_zkit].astype(str).str.strip() != "")]['KEY_ART'])
                df_var['In_Facing'] = df_var['KEY_ART'].isin(facing_set)
            else:
                df_var['In_Facing'] = False

            # Converti valori numerici per prezzo e costo
            if col_p_base:
                df_var['P_BASE_NUM'] = pd.to_numeric(df_var[col_p_base].astype(str).str.replace(',', '.'), errors='coerce')
            else:
                df_var['P_BASE_NUM'] = float('nan')

            if col_p_nuovo:
                df_var['P_NUOVO_NUM'] = pd.to_numeric(df_var[col_p_nuovo].astype(str).str.replace(',', '.'), errors='coerce')
            else:
                df_var['P_NUOVO_NUM'] = float('nan')

            # REGOLE PASSAGGIO SICURO
            # 1. Zero stock e zero facing
            cond_zero_stock_facing = (df_var['Giacenza'] <= 0) & (~df_var['In_Facing'])

            # 2. Tipo non vendita base
            if col_tipo:
                cond_non_base = ~df_var[col_tipo].astype(str).str.contains("Prz. Vendita - base", case=False, na=False)
            else:
                cond_non_base = False

            # 3. Solo variazione costo senza modifica del prezzo di vendita
            cond_prezzo_invariato = (
                (df_var['P_BASE_NUM'].isna() & df_var['P_NUOVO_NUM'].isna()) |
                (df_var['P_BASE_NUM'] == df_var['P_NUOVO_NUM'])
            )

            df_var['Passaggio_Sicuro'] = cond_zero_stock_facing | cond_non_base | cond_prezzo_invariato

            # Calcolo giorni rimanenti e urgenza
            today = pd.to_datetime(datetime.now().date())
            if col_data_ini:
                df_var['Data_Inizio_DT'] = pd.to_datetime(df_var[col_data_ini], errors='coerce')
                df_var['Giorni_Rimanenti'] = (df_var['Data_Inizio_DT'] - today).dt.days
            else:
                df_var['Giorni_Rimanenti'] = 999

            if col_p_ins:
                df_var['Da_Gestire'] = df_var[col_p_ins].isna() | (df_var[col_p_ins].astype(str).str.strip() == "")
            else:
                df_var['Da_Gestire'] = True

            df_var['Urgente'] = df_var['Da_Gestire'] & (df_var['Giorni_Rimanenti'] <= 3) & (~df_var['Passaggio_Sicuro'])

            # Dataset per la gestione (escludendo Passaggio Sicuro)
            df_operativo = df_var[~df_var['Passaggio_Sicuro']].copy()

        # METRICHE
        st.write("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Totale Variazioni", len(df_var))
        m2.metric("Passaggio Sicuro (Esclusi)", int(df_var['Passaggio_Sicuro'].sum()))
        m3.metric("Da Gestire (Effettivi)", len(df_operativo))
        m4.metric("🚨 Alert Urgenti (<= 3 gg)", int(df_operativo['Urgente'].sum()))

        # RIEPILOGO FORNITORI E
