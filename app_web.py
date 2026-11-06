import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Gestione Variazioni Prezzi", layout="wide")

# Creiamo due Tab: una per il filtro e una per la dashboard
tab_filtro, tab_dashboard = st.tabs(["Filtro Variazioni", "📊 Dashboard Statistiche"])

# ==========================================
# TAB 1: FILTRO VARIAZIONI PREZZI
# ==========================================
with tab_filtro:
    st.title("Filtro Variazioni Prezzi in Giacenza")
    st.write("Carica il file delle variazioni e l'export delle giacenze per generare il report incrociato.")

    col1, col2 = st.columns(2)
    with col1:
        file_var = st.file_uploader("Carica File Variazioni (es. VAR_FUTURE)", type=['xlsx', 'xls'])
    with col2:
        file_exp = st.file_uploader("Carica File Export Giacenze (es. export)", type=['xlsx', 'xls'])

    if file_var and file_exp:
        if st.button("Elabora Dati"):
            try:
                # Lettura dei file
                df_var = pd.read_excel(file_var)
                df_exp = pd.read_excel(file_exp)

                # --- 1. ESTRAZIONE NOME NEGOZIO PER LA DASHBOARD ---
                # Prende la prima riga utile della colonna DESCRIZIONE
                try:
                    nome_negozio = str(df_var['DESCRIZIONE'].dropna().iloc[0])
                except:
                    nome_negozio = "Negozio Sconosciuto"

                # Pulizia Variazioni
                df_var_clean = df_var.dropna(subset=['ARTICOLO']).copy()
                df_var_clean['ARTICOLO'] = pd.to_numeric(df_var_clean['ARTICOLO'], errors='coerce')
                
                # Filtro 1: Rimuove i prezzi vuoti
                df_var_clean = df_var_clean.dropna(subset=['PREZZO BASE', 'PREZZO NUOVO'])
                
                # Filtro 2: Esclude gli articoli in promozione IS-R dalla colonna TIPO
                if 'TIPO' in df_var_clean.columns:
                    df_var_clean = df_var_clean[df_var_clean['TIPO'] != 'Prz. vend. prom. IS-R']

                # Pulizia Giacenze
                df_exp['Materiale'] = pd.to_numeric(df_exp['Materiale'], errors='coerce')
                df_exp_stock = df_exp[df_exp['Utilizz. liber.'] > 0]

                # Incrocio Dati
                df_result = pd.merge(
                    df_var_clean,
                    df_exp_stock[['Materiale', 'Utilizz. liber.']],
                    left_on='ARTICOLO',
                    right_on='Materiale',
                    how='inner'
                )
                df_result = df_result.rename(columns={'Utilizz. liber.': 'GIACENZA'})
                df_result = df_result.drop(columns=['Materiale'])

                st.success(f"Elaborazione completata! Trovati {len(df_result)} articoli in giacenza da aggiornare.")
                st.dataframe(df_result.head(10))

                # --- 2. INVIO DATI NASCOSTO AL DATABASE ---
                try:
                    totale_var = len(df_result)
                    # Crea un dizionario con il conteggio dei fornitori
                    if 'DESC.FORNITORE' in df_result.columns:
                        dettaglio_forn = json.dumps(df_result['DESC.FORNITORE'].value_counts().to_dict())
                    else:
                        dettaglio_forn = "{}"

                    url_modulo = "https://docs.google.com/forms/d/e/1FAIpQLSd3zGADI6BtQrnj4QV2Vh4rBRxgGtp0vbdfOcgbrH9PTQHDzA/formResponse"
                    dati_invio = {
                        "entry.1553744770": nome_negozio,
                        "entry.591249622": totale_var,
                        "entry.1751674726": dettaglio_forn
                    }
                    requests.post(url_modulo, data=dati_invio)
                except Exception as e:
                    pass # Evita di bloccare il download in caso di errori di rete

                # --- 3. CREAZIONE EXCEL DA SCARICARE ---
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Variazioni in Giacenza"

                for r in dataframe_to_rows(df_result, index=False, header=True):
                    ws.append(r)

                # Stili
                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                border = Border(left=Side(style='thin', color='D9D9D9'),
                                right=Side(style='thin', color='D9D9D9'),
                                top=Side(style='thin', color='D9D9D9'),
                                bottom=Side(style='thin', color='D9D9D9'))

                # Applica stili intestazione
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                # Applica bordi a tutte le celle
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.border = border

                # Auto-adatta larghezza colonne
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    ws.column_dimensions[column].width = min(max_length + 2, 35)

                wb.save(output)
                output.seek(0)

                # Tasto per il download
                st.download_button(
                    label="📥 Scarica il Report Excel",
                    data=output,
                    file_name="Report_Variazioni_In_Giacenza.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Si è verificato un errore durante l'elaborazione: {e}")

# ==========================================
# TAB 2: DASHBOARD STATISTICHE
# ==========================================
with tab_dashboard:
    st.header("📊 Dashboard Variazioni per Punto Vendita")

    URL_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQURIymDxwh_J3MhmQY72yqvsqn2hL0rciLgAdfgdzDI_vRLzqQBOAElGA1ndr8rVqR_P3PkOZPS6i1/pub?output=csv"

    @st.cache_data(ttl=60)
    def carica_dati_dashboard(url):
        return pd.read_csv(url)

    try:
        df_storico = carica_dati_dashboard(URL_CSV)
        df_storico.rename(columns={df_storico.columns[0]: 'Data'}, inplace=True)
            
        if not df_storico.empty:
            negozi = df_storico['Negozio'].dropna().unique()
            negozio_selezionato = st.selectbox("Scegli il punto vendita da visualizzare:", negozi)
            
            dati_negozio = df_storico[df_storico['Negozio'] == negozio_selezionato].iloc[-1]
            
            st.caption(f"Ultimo aggiornamento per questo negozio: {dati_negozio['Data']}")
            st.metric("Totale Variazioni in Giacenza Elaborate", dati_negozio['Totale'])
            
            dettaglio_str = dati_negozio['Dettaglio']
            if pd.notna(dettaglio_str) and dettaglio_str != "{}":
                dizionario_fornitori = json.loads(dettaglio_str)
                df_grafico = pd.DataFrame(list(dizionario_fornitori.items()), columns=['Fornitore', 'Variazioni'])
                df_grafico['Percentuale (%)'] = ((df_grafico['Variazioni'] / df_grafico['Variazioni'].sum()) * 100).round(2)
                df_grafico = df_grafico.sort_values(by='Percentuale (%)', ascending=False)
                
                st.subheader("Impatto per Fornitore (%)")
                st.bar_chart(df_grafico.set_index('Fornitore')['Percentuale (%)'])
                
                with st.expander("Vedi tabella dettagliata dei fornitori"):
                    st.dataframe(df_grafico, use_container_width=True)
        else:
            st.info("In attesa che venga elaborato il primo file da un negozio per mostrare i dati.")
            
    except Exception as e:
        st.warning("Il database è pronto. Elabora il tuo primo file nella scheda 'Filtro Variazioni' per vedere i grafici!")
