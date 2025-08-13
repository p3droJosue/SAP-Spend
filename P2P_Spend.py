import os
import re
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
from datetime import datetime

def check_numbers(source_list,target_list):
    missing_numbers = [number for number in source_list if number not in target_list]
    df_missing_numbers = pd.DataFrame(missing_numbers,columns=['Missing POs'])
    len_missing_numbers = len(missing_numbers)
    if not df_missing_numbers.empty:
        print(f'Missing numbers stored in df_missing_numbers')
        print('Count of missing POs: ',str(f'{len_missing_numbers:,}'))
        return df_missing_numbers
    else:
        return None


def validate_date(date_input):
    pattern = r'^\d{2}\.\d{2}\.\d{4}$'
    return re.match(pattern, date_input)


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


def main():
    # --- 1) get folder & dates, set up month/year lookups ---
    folder_path, Date_From, Date_To = get_path_dates()
    now = datetime.now()
    months_spanish = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    month = now.month - 1 if now.month > 1 else 12; year = now.year

    print("\n========= STARTING P2P SPEND EXTRACTION =========\n")
    print('\nCode created by Pedro García, GPID: 81010435\n')

    ## --- 2) download & collect raw files via SAP_Spend functions ---
    df_ekko, _, _    = clean_transf_ekko(folder_path, Date_From, Date_To, month, year, months_spanish)
    df_ekpo          = clean_transf_ekpo(df_ekko, folder_path, month, year, months_spanish)
    df_fbl1n         = clean_transf_fbl1n(folder_path, Date_From, Date_To, month, year, months_spanish)
    df_fbl3n         = clean_transf_fbl3n(folder_path, Date_From, Date_To, month, year, months_spanish)


    # --- 3) YOUR DATA CLEANING & TRANSFORM STEPS (moved in here) ---
    # Example: numeric conversions & renames on EKKO
    print('\n- Concatenating the EKKO files -')

    df_ekko = df_ekko[df_ekko['EBELN'] != np.nan]
    numeric_cols = ['WKURS','RLWRT']
    df_ekko[numeric_cols] = df_ekko[numeric_cols].replace(',','',regex=True)
    df_ekko[numeric_cols] = df_ekko[numeric_cols].apply(pd.to_numeric)

    list_EKKO_PO = df_ekko['EBELN'].unique()
    Total_unique_PO_EKKO = len(list_EKKO_PO)
    df_ekko.rename(columns={'EBELN':'Purchasing Document'},inplace=True)
    print('Total created POs: \n',str(f'{Total_unique_PO_EKKO:,}'))


    # Creating EKPO Data Frame
    print('\n- CREATING EKPO DATA FRAME -')

    print('Shape of the EKPO report, # Columns: ', str(f'{df_ekpo.shape[1]:,}'),
        ', # Rows: ',str(f'{df_ekpo.shape[0]:,}'))

    print('\nEKPO Data Frame:')
    numeric_cols2 = ['MENGE','NETPR','NETWR','EFFWR','PEINH']
    df_ekpo[numeric_cols2] = df_ekpo[numeric_cols2].replace(',','',regex=True)
    df_ekpo[numeric_cols2] = df_ekpo[numeric_cols2].apply(pd.to_numeric)
    df_ekpo['PO_Item'] = df_ekpo.apply(lambda row: f"{row['EBELN']}_{row['EBELP']}",axis=1)
 

    # Creating FBL3N Data Frame
    print('\n- CREATING FBL3N DATA FRAME -')
    FBL3N_columns = [
        'G/L Account','Reference Key','Purchasing Document','Document Number','Document Date',
        'Posting Date','Assignment','Clearing Document','Amount in local currency','Local Currency',
        'Reference','Vendor','Material','Company Code','Account','User name',
        'Document Type','Amount in doc. curr.','Document currency','Plant',
        'Amount in loc.curr.2','Local currency 2','Year/month','Transaction',
        'Fiscal Year','Item'
    ]

    df_fbl3n = df_fbl3n[FBL3N_columns]
    fillna_cols = [
        'Amount in local currency','Amount in doc. curr.','Vendor',
        'Amount in loc.curr.2','Item','Fiscal Year'
    ]
    df_fbl3n[fillna_cols] = df_fbl3n[fillna_cols].fillna(0)
    df_fbl3n = df_fbl3n.astype(
    {
        'Item':int,'Fiscal Year':int
    })

    df_fbl3n['Index'] = df_fbl3n['Vendor'].astype(str).str[0]
    numeric_cols3 = ['Amount in local currency','Amount in doc. curr.','Amount in loc.curr.2']
    df_fbl3n[numeric_cols3] = df_fbl3n[numeric_cols3].replace(',','',regex=True)
    df_fbl3n[numeric_cols3] = df_fbl3n[numeric_cols3].apply(pd.to_numeric)

    df_fbl3n['Purchasing Document']=df_fbl3n['Purchasing Document'].fillna('Missing')
    df_fbl3n['PO_Item'] = df_fbl3n.apply(lambda row: f"{row['Purchasing Document']}_{row['Item']}",axis=1)
    df_fbl3n['DocNum_Item'] = df_fbl3n.apply(lambda row: f"{row['Document Number']}_{row['Item']}",axis=1)
    df_fbl3n['PO_Item'] = df_fbl3n['PO_Item'].replace('Missing_0','')
    df_fbl3n = df_fbl3n[df_fbl3n['Document Type'].isin(['WE','WI','WA','RE'])]
    df_fbl3n['Report Name'] = 'FBL3N'
    print(f"Document types in FBL3N: {df_fbl3n['Document Type'].unique()}")


    # Creating FBL1N Data Frame
    print('\n- CREATING FBL1N DATA FRAME -')
    FBL1N_columns = [
        'Account','Assignment','Company Code','Document Number','Purchasing Document','Line item',
        'Document Type','Document Date','Posting Date','Amount in local currency','User name',
        'Amount in doc. curr.','Document currency','Amount in loc.curr.2','Text',
        'Year/month','Profit Center','Invoice reference','Terms of Payment',
        'G/L Account','Reference Key'
    ]

    df_fbl1n = df_fbl1n[FBL1N_columns]
    df_fbl1n.rename(columns={'Account':'Vendor','Line item':'Item'},inplace=True)
    fillna_cols = [
        'Amount in local currency','Amount in doc. curr.','Amount in loc.curr.2',
        'Item','Vendor'
    ]
    df_fbl1n[fillna_cols] = df_fbl1n[fillna_cols].fillna(0)
    df_fbl1n = df_fbl1n.astype(
    {
        'Item':int
    })

    df_fbl1n['Index'] = df_fbl1n['Vendor'].astype(str).str[0]
    numeric_cols4 = ['Amount in local currency','Amount in doc. curr.','Amount in loc.curr.2']
    df_fbl1n[numeric_cols4] = df_fbl1n[numeric_cols4].replace(',','',regex=True)
    df_fbl1n[numeric_cols4] = df_fbl1n[numeric_cols4].apply(pd.to_numeric)
    df_fbl1n = df_fbl1n[df_fbl1n['Document Type'].isin(['KG','KA','KR','RE'])]
    df_fbl1n['Report Name'] = 'FBL1N'
    print(f"Document types in FBL1N: {df_fbl1n['Document Type'].unique()}")


    # Cataloging the Data Frames
    catalogue_path = os.path.join(folder_path,'Catalogues')
    columns_cat4 = [
        'Vendor','Name','Policy Compliance','GP/Non-GP','Category','Sub Category'
    ]
    Cataloge_4 = pd.read_excel(os.path.join(catalogue_path,'PMF_Category_NonPO_Final.xlsx'),
                            sheet_name='NON-PO PMF_Final (2)',usecols=columns_cat4)
    Cataloge_4.rename(columns={'Category':'Category 2','Sub Category':'Sub Category 2'},inplace=True)
    columns_cat11 = [
    'PGr','Description','Class','Category','Sub Category'
    ]
    Cataloge_11 = pd.read_excel(os.path.join(catalogue_path,'PMF_Category_NonPO_Final.xlsx'),
                                sheet_name='EKGRP_2',usecols=columns_cat11) 
    columns_fagll03 =[
        'WBS element','Document Number','Line item'
    ]
    df_fagll03 = pd.read_excel(
        os.path.join(catalogue_path,'CAPEX 2018-2025.xlsx'),
        usecols=columns_fagll03
        )
    df_fagll03['DocNum_Item'] = df_fagll03.apply(
        lambda row: f"{row['Document Number']}_{row['Line item']}",
        axis=1
        )

    wbs_element_docnum = df_fagll03.set_index('DocNum_Item')['WBS element'].to_dict()


    # Merge FBL3N and FBL1N Data Frames
    df_fb3_fb1 = pd.concat([df_fbl3n,df_fbl1n],ignore_index=True)
    columns_to_date = ['Document Date','Posting Date']

    df_fb3_fb1[columns_to_date] = df_fb3_fb1[columns_to_date].apply(
        lambda x: pd.to_datetime(x,format='mixed',dayfirst=True,errors='coerce')
        )

    print(f"FBL3N & FBL1N Concatenated Document Types:{df_fb3_fb1['Document Type'].unique()}")


    # Join EKKO to Merged Data Frame
    keep_columns_ekko = ['Purchasing Document','EKORG','EKGRP','LIFNR','ZBD1T']
    df_ekko = df_ekko[keep_columns_ekko]
    df_merge1 = pd.merge(
        df_fb3_fb1,df_ekko,
        on='Purchasing Document',
        how='left'
        )
    df_merge1['EKGRP'] = df_merge1['EKGRP'].fillna('Missing')
    df_merge1.rename(
        columns={'EKGRP':'Purchasing Group','EKORG':'Purchasing Organization','ZBD1T':'Payment Terms'},
        inplace=True
        )


    # Join EKPO to Merged Data Frame
    keep_columns_ekpo = ['PO_Item','MENGE','NETPR','PEINH','NETWR','EFFWR']
    df_ekpo = df_ekpo[keep_columns_ekpo]
    df_merge2 = pd.merge(
        df_merge1,df_ekpo,
        on='PO_Item',
        how='left'
        )


    ######
    # Join Cataloge 4 and 11 to Merged Data Frame
    print('\n- JOINING CATALOGE 4 AND 11 TO MERGED DATA FRAME -')

    #vendor_paymentterm_mapping = Cataloge_4.set_index('Vendor')['Payment Terms'].to_dict()
    vendor_categoria_mapping = Cataloge_4.drop_duplicates(subset=['Vendor']).set_index('Vendor')['Sub Category 2'].to_dict()
    df_merge2['Sub Category 2'] = df_merge2['Vendor'].map(vendor_categoria_mapping)
    df_merge2['WBS Element'] = df_merge2['DocNum_Item'].map(wbs_element_docnum)
    #df_merge2['Payment Terms'] = df_merge2['Vendor'].map(vendor_paymentterm_mapping)

    ekgrp_category_mapping = Cataloge_11.set_index('PGr')['Sub Category'].to_dict()
    df_merge2['Sub Category'] = df_merge2['Purchasing Group'].map(ekgrp_category_mapping)

    gp_nongp_cat4 = Cataloge_4.drop_duplicates(subset=['Vendor']).set_index('Vendor')['GP/Non-GP'].to_dict()
    gp_nongp_cat11 = Cataloge_11.set_index('PGr')['Class'].to_dict()
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

    # Final Data Frame
    merge_path = os.path.join(folder_path, f"Spend Data Base/{months_spanish[month]}_{year}")
    os.makedirs(merge_path, exist_ok=True)
    df_merge2.to_csv(os.path.join(merge_path,f"PowerBI_DataBase.csv"),index=False)

    concat_final_db(folder_path, df_merge2)
    concat_yt(folder_path, df_ekko, df_ekpo, df_fbl1n, df_fbl3n)

    print("\n========= P2P SPEND EXTRACTION COMPLETED =========\n")

if __name__ == '__main__':
    main()