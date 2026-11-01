import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.set_page_config(
    page_title="Gestione Variazioni Prezzo - Brico",
    page_icon="🏷️",
    layout="wide"
)

st.title("🏷️ Dashboard Variazioni Prezzo Negozio")
st.markdown("Carica i file di esportazione per verificare le variazioni future con giacenza o facing e identificare le urgenze.")

# Sezione caricamento file
col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    file_variazioni = st.file_uploader(
        "1. File Variazioni Listino Future (Excel)",
        type=["xlsx", "xls"]
    )

with col_upload2:
    file_stock = st.file_uploader(
        "2. File Giacenze e Facing (Zkit / Gestionale)",
        type=["xlsx", "xls"]
    )

def pulisci_codice(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

if file_variazioni and file_stock:
    try:
        with st.spinner("Elaborazione dati in corso..."):
            df_var = pd.read_excel(file_variazioni)
            df_stk = pd.read_excel(file_stock)

            # Normalizzazione nomi colonne
            df_var.columns = [str(c).strip().upper() for c in df_var.columns]
            df_stk.columns = [str(c).strip().upper() for c in df_stk.columns]

            # Riconoscimento colonna codice articolo
            col_art_var = next((c for c in df_var.columns if any(k in c for k in ["ARTICOLO", "CODICE", "COD_ART"])), df_var.columns[0])
            col_art_stk = next((c for c in df_stk.columns if any(k in c for k in ["ARTICOLO", "CODICE", "COD_ART"])), df_stk.columns[0])

            df_var["KEY_ART"] = df_var[col_art_var].apply(pulisci_codice)
            df_stk["KEY_ART"] = df_stk[col_art_stk].apply(pulisci_codice)

            # Riconoscimento giacenza e facing
            col_giac = next((c for c in df_stk.columns if any(k in c for k in ["GIAC", "ESISTENZA", "QTA", "QUANT"])), None)
            col_fac = next((c for c in df_stk.columns if any(k in c for k in ["FACING", "FAC", "FRONT"])), None)

            if col_giac:
                df_stk[col_giac] = pd.to_numeric(df_stk[col_giac], errors="coerce").fillna(0)
            else:
                df_stk["GIACENZA"] = 0
                col_giac = "GIACENZA"

            if col_fac:
                df_stk[col_fac] = pd.to_numeric(df_stk[col_fac], errors="coerce").fillna(0)
            else:
                df_stk["FACING"] = 0
                col_fac = "FACING"

            # Merge dei due file
            df_merged = pd.merge(
                df_var,
                df_stk[["KEY_ART", col_giac, col_fac]].drop_duplicates(subset=["KEY_ART"]),
                on="KEY_ART",
                how="inner"
            )

            # Filtro: giacenza > 0 oppure facing > 0
            df_filtrato = df_merged[
                (df_merged[col_giac] > 0) | (df_merged[col_fac] > 0)
            ].copy()

            # Riconoscimento data decorrenza / validità
            col_data = next((c for c in df_filtrato.columns if any(k in c for k in ["DATA", "DECORRENZA", "VALIDITA", "SCADENZA"])), None)
            
            oggi = datetime.now()
            soglia_urgenza = oggi + timedelta(days=3)

            if col_data:
                df_filtrato["DATA_CONV"] = pd.to_datetime(df_filtrato[col_data], errors="coerce")
                df_filtrato["URGENTE"] = df_filtrato["DATA_CONV"].apply(
                    lambda d: "🚨 URGENTE" if pd.notna(d) and d <= soglia_urgenza else "Normale"
                )
            else:
                df_filtrato["URGENTE"] = "Normale"

            # Riconoscimento colonna fornitore
            col_forn = next((c for c in df_filtrato.columns if any(k in c for k in ["FORNITORE", "FORN", "RAGIONE_SOC", "DESC_FORN"])), None)

        st.success("Elaborazione completata con successo!")

        # Metric cards
        col_m1, col_m2, col_m3 = st.columns(3)
        tot_var = len(df_var)
        tot_presenti = len(df_filtrato)
        tot_urgenti = (df_filtrato["URGENTE"] == "🚨 URGENTE").sum()

        col_m1.metric("Totale Variazioni Listino", tot_var)
        col_m2.metric("In Negozio (Giacenza > 0 o Facing)", tot_presenti)
        col_m3.metric("Scadenze Urgenti (≤ 3 gg)", tot_urgenti)

        # Grafico e riepilogo fornitori se presente la colonna
        if col_forn:
            st.subheader("📊 Riepilogo Variazioni per Fornitore")
            recap_forn = df_filtrato[col_forn].value_counts().reset_index()
            recap_forn.columns = ["Fornitore", "Numero Variazioni"]
            st.bar_chart(recap_forn.set_index("Fornitore"))

        # Visualizzazione Tabella
        st.subheader("📋 Dettaglio Articoli da Verificare")
        colonne_da_mostrare = [c for c in df_filtrato.columns if c not in ["KEY_ART", "DATA_CONV"]]
        st.dataframe(df_filtrato[colonne_da_mostrare], use_container_width=True)

        # Esportazione Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_filtrato[colonne_da_mostrare].to_excel(writer, index=False, sheet_name="Variazioni_Negozio")
        buffer.seek(0)

        st.download_button(
            label="📥 Scarica File Excel Elaborato",
            data=buffer,
            file_name=f"Variazioni_Negozio_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Si è verificato un errore durante l'elaborazione dei file: {e}")
