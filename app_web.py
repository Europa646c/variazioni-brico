import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Variazioni Brico", page_icon="📦", layout="wide")

st.title("📦 Gestione Variazioni Listino Brico")
st.write("Carica i file qui sotto per iniziare l'analisi.")

# --- UPLOAD DEI FILE ---
col1, col2 = st.columns(2)
with col1:
    file_variazioni = st.file_uploader("1. File Variazioni Future", type=['xlsx', 'xls'])
with col2:
    file_stock = st.file_uploader("2. File Stock & Facing", type=['xlsx', 'xls'])

# --- LOGICA DI ELABORAZIONE ---
if file_variazioni is not None and file_stock is not None:
    if st.button("Elabora Dati", type="primary"):
        with st.spinner("Lettura e analisi dei file in corso..."):
            try:
                # 1. LETTURA FILE
                df_var = pd.read_excel(file_variazioni, sheet_name="Variazioni Listino")
                df_stock = pd.read_excel(file_stock, sheet_name="Stock")
                df_zkit = pd.read_excel(file_stock, sheet_name="Zkit")

                # 2. PULIZIA NOMI COLONNE (Rimuove spazi nascosti e a capo)
                df_var.columns = df_var.columns.astype(str).str.strip()
                df_stock.columns = df_stock.columns.astype(str).str.strip()
                df_zkit.columns = df_zkit.columns.astype(str).str.strip()

                # Controllo di sicurezza diagnostico
                if 'Cod. Materiale' not in df_stock.columns:
                    st.error("ERRORE: Non trovo la colonna 'Cod. Materiale' nel foglio Stock.")
                    st.warning(f"Colonne rilevate nel file: {df_stock.columns.tolist()}")
                    st.stop()
                if 'ARTICOLO' not in df_var.columns:
                    st.error("ERRORE: Non trovo la colonna 'ARTICOLO' nel file Variazioni.")
                    st.warning(f"Colonne rilevate nel file: {df_var.columns.tolist()}")
                    st.stop()

                # 3. PREPARAZIONE DATI
                df_var = df_var.dropna(subset=['ARTICOLO']).copy()
                df_var['ARTICOLO'] = df_var['ARTICOLO'].astype(float).astype(int).astype(str)

                df_stock['Cod. Materiale'] = df_stock['Cod. Materiale'].astype(str)
                stock_agg = df_stock.groupby('Cod. Materiale')['Stock qtà'].sum().reset_index()

                df_zkit['Materiale'] = df_zkit['Materiale'].astype(str)
                df_zkit['Has_Facing'] = df_zkit['Modulo'].notna() & (df_zkit['Modulo'] != '')
                facing_agg = df_zkit.groupby('Materiale')['Has_Facing'].any().reset_index()

                # 4. UNIONE DATI
                df_merged = df_var.merge(stock_agg, left_on='ARTICOLO', right_on='Cod. Materiale', how='left')
                df_merged['Stock qtà'] = df_merged['Stock qtà'].fillna(0)

                df_merged = df_merged.merge(facing_agg, left_on='ARTICOLO', right_on='Materiale', how='left')
                df_merged['Has_Facing'] = df_merged['Has_Facing'].fillna(False)

                # 5. FILTRI E URGENZE (Passaggio Sicuro)
                def is_passaggio_sicuro(row):
                    if "Prz. Vendita - base" not in str(row.get('TIPO', '')):
                        return True
                    if row['Stock qtà'] <= 0 and not row['Has_Facing']:
                        return True
                    return False

                df_merged['Passaggio_Sicuro'] = df_merged.apply(is_passaggio_sicuro, axis=1)
                
                df_output = df_merged[~df_merged['Passaggio_Sicuro']].copy()

                today = datetime.now()
                df_output['DATA INIZIO'] = pd.to_datetime(df_output['DATA INIZIO'], format='%d/%m/%Y', errors='coerce')
                df_output['Giorni_Rimanenti'] = (df_output['DATA INIZIO'] - today).dt.days

                df_output['Da_Gestire'] = df_output['P_NUOVO_INS'].isna() & df_output['C_NUOVO_INS'].isna()
                df_output['Alert_Imminente'] = df_output['Da_Gestire'] & (df_output['Giorni_Rimanenti'] <= 3)

                # 6. CREAZIONE INTERFACCIA OUTPUT
                if df_output.empty:
                    st.success("🎉 Nessun articolo da gestire! Tutti gli articoli rientrano nel passaggio automatico.")
                else:
                    recap = []
                    for forn, group in df_output.groupby('DESC.FORNITORE'):
                        da_gestire_count = group['Da_Gestire'].sum()
                        alert_count = group['Alert_Imminente'].sum()
                        if da_gestire_count > 0 or alert_count > 0:
                            recap.append({
                                'Fornitore': forn,
                                'Da Gestire': da_gestire_count,
                                'Alert Imminenti': alert_count,
                                'Scadenza (gg)': group['Giorni_Rimanenti'].min()
                            })
                    
                    df_recap = pd.DataFrame(recap)
                    if not df_recap.empty:
                        df_recap = df_recap.sort_values(by=['Alert Imminenti', 'Da Gestire'], ascending=[False, False])
                        st.subheader("📊 Riepilogo Fornitori")
                        st.dataframe(df_recap, use_container_width=True)

                    # PREPARAZIONE EXCEL DA SCARICARE
                    df_excel = pd.DataFrame({
                        'Fornitore': df_output['DESC.FORNITORE'],
                        'Codice Articolo': df_output['ARTICOLO'],
                        'EAN': df_output['EAN'],
                        'Descrizione': df_output['DESCRIZIONE'],
                        'Giacenza': df_output['Stock qtà'],
                        'In_Facing': df_output['Has_Facing'].map({True: 'SI', False: 'NO'}),
                        'Data Passaggio Auto': df_output['DATA INIZIO'].dt.strftime('%d/%m/%Y'),
                        'Giorni Rimanenti': df_output['Giorni_Rimanenti'],
                        'Urgente': df_output['Alert_Imminente'].map({True: 'SI', False: 'NO'})
                    })

                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                        df_excel.to_excel(writer, index=False, sheet_name='Da Gestire')
                    output_excel.seek(0)

                    st.download_button(
                        label="📥 Scarica File Excel Dettagliato",
                        data=output_excel,
                        file_name="Report_Variazioni.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Errore critico durante l'esecuzione: {str(e)}")
