import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Configurazione della pagina
st.set_page_config(page_title="Filtro Variazioni Prezzi", layout="wide")
st.title("Filtro Variazioni Prezzi in Giacenza")
st.write("Carica il file delle variazioni e l'export delle giacenze per generare il report incrociato.")

# Box per il caricamento dei file
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

            # Pulizia Variazioni
            df_var_clean = df_var.dropna(subset=['ARTICOLO']).copy()
            df_var_clean['ARTICOLO'] = pd.to_numeric(df_var_clean['ARTICOLO'], errors='coerce')
            df_var_clean = df_var_clean.dropna(subset=['PREZZO BASE', 'PREZZO NUOVO'])

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

            st.success(f"Elaborazione completata! Trovati {len(df_result)} articoli in giacenza.")
            st.dataframe(df_result.head(10)) # Mostra un'anteprima a schermo

            # Creazione file Excel formattato in memoria
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
