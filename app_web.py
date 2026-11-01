def elabora_variazioni_brico(file_variazioni, file_stock):
    # 1. Caricamento Dati
    df_var = pd.read_excel(file_variazioni, sheet_name="Variazioni Listino")
    df_stock = pd.read_excel(file_stock, sheet_name="Stock")
    df_zkit = pd.read_excel(file_stock, sheet_name="Zkit")

    # --- NOVITÀ: Pulizia dei nomi delle colonne da spazi vuoti accidentali ---
    df_var.columns = df_var.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    df_zkit.columns = df_zkit.columns.str.strip()
    # -------------------------------------------------------------------------

    # 2. Pulizia e Preparazione
    df_var = df_var.dropna(subset=['ARTICOLO']).copy()
    df_var['ARTICOLO'] = df_var['ARTICOLO'].astype(float).astype(int).astype(str)

    df_stock['Cod. Materiale'] = df_stock['Cod. Materiale'].astype(str)
    stock_agg = df_stock.groupby('Cod. Materiale')['Stock qtà'].sum().reset_index()

    df_zkit['Materiale'] = df_zkit['Materiale'].astype(str)
    df_zkit['Has_Facing'] = df_zkit['Modulo'].notna() & (df_zkit['Modulo'] != '')
    facing_agg = df_zkit.groupby('Materiale')['Has_Facing'].any().reset_index()

    # 3. Merge dei Dati
    df_merged = df_var.merge(stock_agg, left_on='ARTICOLO', right_on='Cod. Materiale', how='left')
    df_merged['Stock qtà'] = df_merged['Stock qtà'].fillna(0)

    df_merged = df_merged.merge(facing_agg, left_on='ARTICOLO', right_on='Materiale', how='left')
    df_merged['Has_Facing'] = df_merged['Has_Facing'].fillna(False)

    # 4. Logica di Scrematura
    def check_passaggio_sicuro(row):
        if "Prz. Vendita - base" not in str(row['TIPO']):
            return True
        if row['Stock qtà'] <= 0 and not row['Has_Facing']:
            return True
        return False

    df_merged['Passaggio_Sicuro'] = df_merged.apply(check_passaggio_sicuro, axis=1)

    today = datetime.now()
    df_merged['DATA INIZIO'] = pd.to_datetime(df_merged['DATA INIZIO'], format='%d/%m/%Y', errors='coerce')
    df_merged['Giorni_Rimanenti'] = (df_merged['DATA INIZIO'] - today).dt.days

    df_merged['Da_Gestire'] = df_merged['P_NUOVO_INS'].isna() & df_merged['C_NUOVO_INS'].isna()
    df_merged['Alert_Imminente'] = df_merged['Da_Gestire'] & (df_merged['Giorni_Rimanenti'] <= 3)

    # 5. Generazione Output
    df_output = df_merged[~df_merged['Passaggio_Sicuro']].copy()

    recap = []
    for forn, group in df_output.groupby('DESC.FORNITORE'):
        da_gestire_count = group['Da_Gestire'].sum()
        alert_count = group['Alert_Imminente'].sum()
        
        if da_gestire_count > 0 or alert_count > 0:
            min_days = group['Giorni_Rimanenti'].min()
            recap.append({
                'Fornitore': forn,
                'Da Gestire': da_gestire_count,
                'Alert Imminenti': alert_count,
                'Scadenza (gg)': min_days
            })
            
    df_recap = pd.DataFrame(recap)
    if not df_recap.empty:
        df_recap = df_recap.sort_values(by=['Alert Imminenti', 'Da Gestire'], ascending=[False, False])

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

    return df_recap, output_excel
