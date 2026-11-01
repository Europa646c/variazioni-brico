import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# Configurazione pagina web ALLARGATA per far spazio ai grafici
st.set_page_config(page_title="Dashboard Variazioni", page_icon="🛒", layout="wide")

st.title("🛒 Dashboard Variazioni Prezzo Brico")
st.markdown("Analizza le variazioni giornaliere, visualizza il riepilogo per fornitore e scarica il report.")

# Caricamento file
col_file1, col_file2 = st.columns(2)
with col_file1:
    file_var = st.file_uploader("1. Variazioni Future (Excel)", type=['xlsx'])
with col_file2:
    file_stock = st.file_uploader("2. Stock e Facing (Excel)", type=['xlsx'])

if file_var and file_stock:
    if st.button("Analizza e Genera Dashboard", type="primary", use_container_width=True):
        with st.spinner("Incrocio dei dati in corso..."):
            try:
                # Lettura dati
                df_var = pd.read_excel(file_var, sheet_name=0)
                df_var = df_var.dropna(subset=['ARTICOLO']).copy()
                df_var['ARTICOLO'] = df_var['ARTICOLO'].astype(float).astype(int).astype(str)
                df_var['DATA INIZIO'] = pd.to_datetime(df_var['DATA INIZIO'], format='%d/%m/%Y', errors='coerce')
                df_var['DATA INIZIO_STR'] = df_var['DATA INIZIO'].dt.strftime('%d/%m/%Y')

                df_stock = pd.read_excel(file_stock, sheet_name='Stock')
                df_zkit = pd.read_excel(file_stock, sheet_name='Zkit')

                # Elaborazione
                df_stock['Articolo'] = df_stock['Articolo'].astype(str)
                df_stock['Stock qtà'] = pd.to_numeric(df_stock['Stock qtà'], errors='coerce').fillna(0)
                stock_dict = df_stock[df_stock['Stock qtà'] > 0].set_index('Articolo')['Stock qtà'].to_dict()

                df_zkit['Articolo'] = df_zkit['Articolo'].astype(str)
                zkit_articles = set(df_zkit[df_zkit['Modulo'].notna() & (df_zkit['Modulo'] != '')]['Articolo'])

                # Filtri base
                df_var['Ha_Stock_o_Facing'] = df_var['ARTICOLO'].isin(stock_dict.keys()) | df_var['ARTICOLO'].isin(zkit_articles)
                df_var['Da_Variare'] = df_var['P_NUOVO_INS'].isna() & df_var['C_NUOVO_INS'].isna()

                today = pd.Timestamp.today()
                df_var['Giorni_a_Inizio'] = (df_var['DATA INIZIO'] - today).dt.days
                
                df_da_gestire = df_var[df_var['Ha_Stock_o_Facing'] & df_var['Da_Variare']].copy()
                df_da_gestire['Giacenza'] = df_da_gestire['ARTICOLO'].map(stock_dict).fillna(0)
                df_da_gestire['In_Facing'] = df_da_gestire['ARTICOLO'].apply(lambda x: 'SI' if x in zkit_articles else 'NO')
                df_da_gestire['Urgenza (<=3gg)'] = df_da_gestire['Giorni_a_Inizio'].apply(lambda x: 'URGENTE' if pd.notna(x) and x <= 3 else '')

                colonne = ['DESC.FORNITORE', 'ARTICOLO', 'EAN', 'DESCRIZIONE', 'Giacenza', 'In_Facing', 'DATA INIZIO_STR', 'Giorni_a_Inizio', 'Urgenza (<=3gg)']
                df_report = df_da_gestire[colonne].rename(columns={
                    'DESC.FORNITORE': 'Fornitore', 'ARTICOLO': 'Codice Articolo', 
                    'DATA INIZIO_STR': 'Data Passaggio Auto', 'Giorni_a_Inizio': 'Giorni Rimanenti'
                }).sort_values(by=['Fornitore', 'Giorni Rimanenti'])

                st.success("Dati elaborati con successo!")
                st.divider()

                # DASHBOARD E GRAFICI
                tot_articoli = len(df_report)
                tot_urgenti = df_report[df_report['Urgenza (<=3gg)'] == 'URGENTE'].shape[0]

                met1, met2, met3 = st.columns(3)
                met1.metric(label="Da Gestire (Totale)", value=tot_articoli)
                met2.metric(label="Urgenti (Passaggio entro 3 gg)", value=tot_urgenti)

                st.subheader("📊 Analisi per Fornitore")
                summary_df = df_report.groupby('Fornitore').agg(
                    Da_Gestire=('Codice Articolo', 'count'),
                    Urgenti=('Urgenza (<=3gg)', lambda x: (x == 'URGENTE').sum())
                ).reset_index().sort_values('Da_Gestire', ascending=False)

                col_tab, col_grafico = st.columns(2)
                with col_tab:
                    st.dataframe(summary_df, hide_index=True, use_container_width=True)
                with col_grafico:
                    st.bar_chart(summary_df.set_index('Fornitore'))
                
                st.subheader("📝 Dettaglio Articoli")
                st.dataframe(df_report, hide_index=True, height=250, use_container_width=True)
                st.divider()

                # EXCEL
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Articoli da Gestire"
                headers = list(df_report.columns)
                ws.append(headers)

                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                urgent_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                urgent_font = Font(color="9C0006", bold=True)
                thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), 
                                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border

                for r_idx, row in enumerate(dataframe_to_rows(df_report, index=False, header=False), 2):
                    for c_idx, value in enumerate(row, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=value)
                        cell.border = thin_border
                        if headers[c_idx-1] in ['Giacenza', 'In_Facing', 'Data Passaggio Auto', 'Giorni Rimanenti', 'Urgenza (<=3gg)']:
                            cell.alignment = Alignment(horizontal="center")
                        if row[-1] == 'URGENTE':
                            cell.fill = urgent_fill
                            if headers[c_idx-1] == 'Urgenza (<=3gg)': cell.font = urgent_font

                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 15
                
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                
                wb.save(output)
                output.seek(0)
                
                col_btn1, col_btn2, col_btn3 = st.columns([1,2,1])
                with col_btn2:
                    st.download_button(
                        label="📥 SCARICA IL REPORT EXCEL PRONTO",
                        data=output,
                        file_name=f"Variazioni_Grottammare_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Si è verificato un errore durante la lettura dei file: {str(e)}")