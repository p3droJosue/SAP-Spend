import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from datetime import datetime
import threading
import os
import pandas as pd
import numpy as np
from utils import get_path_dates
from SAP_Spend import (
    clean_transf_ekko,
    clean_transf_ekpo,
    clean_transf_fbl1n,
    clean_transf_fbl3n,
    concat_yt
)
import re

class P2PApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("P2P Spend Extraction")
        self.geometry("600x500")
        self.build_ui()

    def build_ui(self):
        tk.Label(self, text="Carpeta destino:").pack(anchor="w", padx=10, pady=5)
        frame = tk.Frame(self); frame.pack(fill="x", padx=10)
        self.folder_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.folder_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame, text="Elegir", command=self.choose_folder).pack(side="right")

        tk.Label(self, text="Fecha Desde (dd.mm.yyyy):").pack(anchor="w", padx=10, pady=5)
        self.from_var = tk.StringVar(); tk.Entry(self, textvariable=self.from_var).pack(fill="x", padx=10)

        tk.Label(self, text="Fecha Hasta (dd.mm.yyyy):").pack(anchor="w", padx=10, pady=5)
        self.to_var = tk.StringVar(); tk.Entry(self, textvariable=self.to_var).pack(fill="x", padx=10)

        tk.Button(self, text="Ejecutar", command=self.run_pipeline).pack(pady=10)

        self.log = scrolledtext.ScrolledText(self, state='disabled', height=15)
        self.log.pack(fill="both", padx=10, pady=10, expand=True)

    def choose_folder(self):
        path = filedialog.askdirectory(title="Selecciona carpeta destino")
        if path: self.folder_var.set(path)

    def log_message(self, msg):
        self.log.configure(state='normal')
        self.log.insert("end", msg + "\n")
        self.log.configure(state='disabled')
        self.log.see("end")

    def run_pipeline(self):
        folder = self.folder_var.get(); d1 = self.from_var.get(); d2 = self.to_var.get()
        if not os.path.isdir(folder):
            messagebox.showerror("Error", "Carpeta inválida"); return
        if not (re.match(r'^\d{2}\.\d{2}\.\d{4}$', d1) and re.match(r'^\d{2}\.\d{2}\.\d{4}$', d2)):
            messagebox.showerror("Error", "Fechas con formato incorrecto"); return
        threading.Thread(target=self.pipeline_thread, args=(folder,d1,d2), daemon=True).start()

    def concat_final_db(folder_path, df):
        output_path = os.path.join(folder_path, f"Spend Data Base/Concat_Spend_CSV/PowerBI_DataBase.csv")
        old_db_path = os.path.join(folder_path, f"Spend Data Base/Concat_Spend_CSV/PowerBI_DataBase.csv")

        if os.path.exists(old_db_path):
            chunks = pd.read_csv(old_db_path, chunksize=1000000, low_memory=False)
            df_old_db = pd.concat(chunks, ignore_index=True)
        else:
            df_old_db = pd.DataFrame()

        df_new_db = pd.concat([df_old_db, df], ignore_index=True)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        df_new_db.to_csv(output_path, index=False)
        print(f"Years in final Data Frame: {df_new_db['Year/month'].unique()}")
        print(f"Final database saved to {output_path}")

    def pipeline_thread(self, folder, date_from, date_to):
        self.log_message("=== Iniciando Extracción P2P ===")
        now = datetime.now()
        months_spanish = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',
                          5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',
                          9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
        month = now.month - 1 if now.month>1 else 12; year = now.year

        # 1) Download via SAP_Spend
        df_ekko, _, _ = clean_transf_ekko(folder, date_from, date_to, month, year, months_spanish)
        self.log_message(f"EKKO registros: {len(df_ekko)}")
        df_ekpo = clean_transf_ekpo(df_ekko, folder, month, year, months_spanish)
        self.log_message(f"EKPO registros: {len(df_ekpo)}")
        df_fbl1n = clean_transf_fbl1n(folder, date_from, date_to, month, year, months_spanish)
        self.log_message(f"FBL1N registros: {len(df_fbl1n)}")
        df_fbl3n = clean_transf_fbl3n(folder, date_from, date_to, month, year, months_spanish)
        self.log_message(f"FBL3N registros: {len(df_fbl3n)}")

        # 2) Data Cleaning and trasnformation
        self.log_message("\n=== Limpiando y transformando datos ===\n")
        self.log_message("Limpiando EKKO...")

        df_ekko = df_ekko[df_ekko['EBELN'].notna()]
        for col in ['WKURS', 'RLWRT']:
            df_ekko[col] = df_ekko[col].replace(',','', regex=True).astype(float)

        df_ekko.rename(columns={'EBELN':'Purchasing Document'},inplace=True)
        self.log_message(f"Total POs únicos:{df_ekko['Purchasing Document'].nunique()}")

        self.log_message("Limpiando EKPO...")
        for col in ['MENGE', 'NETPR', 'NETWR', 'PEINH', 'EFFWR']:
            df_ekpo[col] = df_ekpo[col].replace(',','', regex=True).astype(float)
        
        df_ekpo['PO_Item'] = df_ekpo.apply(lambda x: f"{x['EBELN']}_{x['EBELP']}", axis=1)

        self.log_message("- Limpiando y combinando FBL3N + FBL1N...")
        # prepare FBL3N
        fbl3n_columns = [
        'G/L Account','Reference Key','Purchasing Document','Document Number','Document Date',
        'Posting Date','Assignment','Clearing Document','Amount in local currency','Local Currency',
        'Reference','Vendor','Material','Company Code','Account','User name',
        'Document Type','Amount in doc. curr.','Document currency','Plant',
        'Amount in loc.curr.2','Local currency 2','Year/month','Transaction',
        'Fiscal Year','Item'
        ]
        df_fbl3n = df_fbl3n[fbl3n_columns]
        df_fbl3n['Index'] = df_fbl3n['Vendor'].astype(str).str[0]
        numeric_cols = ['Amount in local currency','Amount in doc. curr.','Amount in loc.curr.2']
        for col in numeric_cols:
            df_fbl3n[col] = df_fbl3n[col].replace(',','', regex=True).astype(float)

        df_fbl3n['PO_Item'] = df_fbl3n.apply(lambda x: f"{x['Purchasing Document']}_{x['Item']}", axis=1)
        df_fbl3n['Purchasing Document']=df_fbl3n['Purchasing Document'].fillna('Missing')
        df_fbl3n['Item'] = df_fbl3n['Item'].fillna(0).astype(int)
        df_fbl3n['DocNum_Item'] = df_fbl3n.apply(lambda row: f"{row['Document Number']}_{row['Item']}",axis=1)
        df_fbl3n['PO_Item'] = df_fbl3n['PO_Item'].replace('Missing_0','')
        df_fbl3n = df_fbl3n[df_fbl3n['Document Type'].isin(['WE','WI','WA','RE'])]
        df_fbl3n['Report Name'] = 'FBL3N'

        # Prepare FBL1N
        self.log_message("- Limpiando y combinando FBL1N...")
        fbl1n_columns = [
        'Account','Assignment','Company Code','Document Number','Purchasing Document','Line item',
        'Document Type','Document Date','Posting Date','Amount in local currency','User name',
        'Amount in doc. curr.','Document currency','Amount in loc.curr.2','Text',
        'Year/month','Profit Center','Invoice reference','Terms of Payment',
        'G/L Account','Reference Key'
        ]

        df_fbl1n = df_fbl1n[fbl1n_columns]
        df_fbl1n.rename(columns={'Account':'Vendor','Line item':'Item'},inplace=True)
        df_fbl1n['Index'] = df_fbl1n['Vendor'].astype(str).str[0]
        df_fbl1n['Item'] = df_fbl1n['Item'].fillna(0).astype(int)
        for col in numeric_cols:
            df_fbl1n[col] = df_fbl1n[col].replace(',','', regex=True).astype(float)
        
        df_fbl1n = df_fbl1n[df_fbl1n['Document Type'].isin(['KG','KA','KR','RE'])]
        df_fbl1n['Report Name'] = 'FBL1N'

        df_nonpo = pd.concat([df_fbl3n, df_fbl1n], ignore_index=True)
        for dcol in ['Document Date','Posting Date']:
            df_nonpo[dcol] = pd.to_datetime(df_nonpo[dcol], dayfirst=True, errors='coerce')
        
        # Join EKKO
        self.log_message("- Combinando EKKO con FBL3N/FBL1N...")
        df_merge1 = df_nonpo.merge(
            df_ekko[['Purchasing Document','EKORG','EKGRP','LIFNR','ZBD1T']],
            on='Purchasing Document', how='left'
        ).rename(columns={'EKGRP':'Purchasing Group','EKORG':'Purchasing Organization','ZBD1T':'Payment Terms'})
        df_merge1['Purchasing Group'] = df_merge1['Purchasing Group'].fillna('Missing')

        # Join EKPO
        self.log_message("- Combinando EKPO con FBL3N/FBL1N...")
        keep_columns_ekpo = ['PO_Item','MENGE','NETPR','PEINH','NETWR','EFFWR']
        df_ekpo = df_ekpo[keep_columns_ekpo]
        df_merge2 = df_merge1.merge(
            df_ekpo[['PO_Item','MENGE','NETPR','NETWR','EFFWR','PEINH']],
            on='PO_Item', how='left'
        )
        
        # load catalogs
        self.log_message("\n- Aplicando catálogos y categorizaciones… -\n")
        catalogue_path = os.path.join(folder, 'Catalogues')
        columns_cat4 = [
        'Vendor','Name','Policy Compliance','GP/Non-GP','Category','Sub Category'
        ]
        cat4 = pd.read_excel(os.path.join(catalogue_path,'PMF_Category_NonPO_Final.xlsx'),
                             sheet_name='NON-PO PMF_Final (2)', usecols=columns_cat4)
        cat4.rename(columns={'Category':'Category 2','Sub Category':'Sub Category 2'},inplace=True)
        columns_cat11 = [
        'PGr','Description','Class','Category','Sub Category'
        ]
        cat11 = pd.read_excel(os.path.join(catalogue_path,'PMF_Category_NonPO_Final.xlsx'),
                              sheet_name='EKGRP_2', usecols=columns_cat11)
        columns_fag =[
        'WBS element','Document Number','Line item'
        ]
        df_fag = pd.read_excel(os.path.join(catalogue_path,'CAPEX 2018-2025.xlsx'), usecols=columns_fag)
        df_fag['DocNum_Item'] = df_fag.apply(lambda r: f"{r['Document Number']}_{r['Line item']}", axis=1)
        wbs_map = df_fag.drop_duplicates(subset=['DocNum_Item']).set_index('DocNum_Item')['WBS element'].to_dict()

        # map catalogs (example for Sub Category)
        sub2 = cat4.drop_duplicates(subset=['Vendor']).set_index('Vendor')['Sub Category 2'].to_dict()
        df_merge2['Sub Category 2'] = df_merge2['Vendor'].map(sub2)
        df_merge2['WBS Element'] = df_merge2['DocNum_Item'].map(wbs_map)
        sub11 = cat11.drop_duplicates(subset=['PGr']).set_index('PGr')['Sub Category'].to_dict()
        df_merge2['Sub Category'] = df_merge2['Purchasing Group'].map(sub11)

        
        vendor_categoria_mapping = cat4.drop_duplicates(subset=['Vendor']).set_index('Vendor')['Sub Category 2'].to_dict()
        df_merge2['Sub Category 2'] = df_merge2['Vendor'].map(vendor_categoria_mapping)
        df_merge2['WBS Element'] = df_merge2['DocNum_Item'].map(wbs_map)

        ekgrp_category_mapping = cat11.set_index('PGr')['Sub Category'].to_dict()
        df_merge2['Sub Category'] = df_merge2['Purchasing Group'].map(ekgrp_category_mapping)

        gp_nongp_cat4 = cat4.drop_duplicates(subset=['Vendor']).set_index('Vendor')['GP/Non-GP'].to_dict()
        gp_nongp_cat11 = cat11.set_index('PGr')['Class'].to_dict()
        df_merge2['Class'] = df_merge2['Purchasing Group'].map(gp_nongp_cat11)
        df_merge2['GP/Non-GP'] = df_merge2['Vendor'].map(gp_nongp_cat4)

        # Define Full Category Conditions
        Condition1 = (df_merge2['Purchasing Group']!='ARI')&(df_merge2['Purchasing Document'].notna())
        Condition2 = (df_merge2['Document Type'].isin(['KR','KA']))&(df_merge2['Purchasing Document'].isna())
        Condition3 = (df_merge2['Document Type']=='KG')&(df_merge2['Purchasing Document'].isna())
        Condition4 = ((df_merge2['Purchasing Group']=='ARI')|(df_merge2['Purchasing Group']=='Missing')) & (df_merge2['Purchasing Document'].notna())  
        Condition5 = (df_merge2['WBS Element'].notna()) #&(df_merge2['Purchasing Document'].notna())
        Condition6 = (df_merge2['Vendor'].astype(str).str.startswith('4',na=False))
        Condition7 = (df_merge2['Index'].between('6','8'))
        Condition8 = (df_merge2['G/L Account']=='1026002')&(df_merge2['Document Type']=='WA')&(df_merge2['Document Number'].astype(str).str.startswith('49',na=False))

        df_merge2['Full Category'] = np.select(
            [Condition5,Condition6,Condition7,Condition8,Condition1,Condition2,Condition3,Condition4],
            ['Capex','Employee','Interco','Popato Obsolete',df_merge2['Sub Category'],df_merge2['Sub Category 2'],'Credit Note',
            df_merge2['Sub Category 2']],
            default=''
        )

        # Define Full GP/Non-GP Conditions
        Cond1 = (df_merge2['Purchasing Group']!='Missing')&(df_merge2['Purchasing Document'].notna())
        Cond2 = (df_merge2['Document Type'].isin(['KR','KA','KG','RE']))&(df_merge2['Purchasing Document'].isna())
        Cond3 = (df_merge2['Purchasing Group']=='Missing')&(df_merge2['Purchasing Document'].notna())
        Cond4 = (df_merge2['G/L Account']=='1026002')&(df_merge2['Document Type']=='WA')&(df_merge2['Document Number'].astype(str).str.startswith('49',na=False))

        df_merge2['Full GP/Non-GP'] = np.select(
            [Cond4,Cond1,Cond2,Cond3],
            ['GP',df_merge2['Class'],df_merge2['GP/Non-GP'],df_merge2['GP/Non-GP']],default=''
        )
        df_merge2['Posting Date'] = df_merge2['Posting Date'].dt.strftime('%d.%m.%Y')
        df_merge2['Document Date'] = df_merge2['Document Date'].dt.strftime('%d.%m.%Y')


        # Final formatting and write
        merge_path = os.path.join(folder, f"Spend Data Base/{months_spanish[month]}_{year}")
        os.makedirs(merge_path, exist_ok=True)
        out_file = os.path.join(merge_path, "PowerBI_DataBase.csv")
        df_merge2.to_csv(out_file, index=False)
        self.log_message(f"PowerBI_DataBase.csv saved")

        # Append history and YT
        self.log_message("\n=== Actualizando Spend Data Base ===\n")
        old_db_path = os.path.join(folder, f"Spend Data Base/Concat_Spend_CSV/PowerBI_DataBase.csv")
        if os.path.exists(old_db_path):
            chunks = pd.read_csv(old_db_path, chunksize=1000000, low_memory=False)
            df_old_db = pd.concat(chunks, ignore_index=True)
        else:
            df_old_db = pd.DataFrame()

        df_new_db = pd.concat([df_old_db, df_merge2], ignore_index=True)
        merge_path = os.path.join(folder, f"Spend Data Base/Concat_Spend_CSV")
        os.makedirs(merge_path, exist_ok=True)
        df_new_db.to_csv(os.path.join(merge_path,f"PowerBI_DataBase.csv"),index=False)
        self.log_message(f"PowerBI_DataBase.csv concatenated and saved")

        self.log_message("\n=== Actualizando YT DB ===\n")
        concat_yt(folder, df_ekko, df_ekpo, df_fbl1n, df_fbl3n)
        self.log_message("=== Proceso completado ===")


if __name__ == "__main__":
    app = P2PApp()
    app.mainloop()
