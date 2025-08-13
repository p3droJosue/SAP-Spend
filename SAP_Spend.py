import pandas as pd
import os
import time
from datetime import datetime, timedelta
import win32com.client



def Download_EKKO(Date_From,Date_To,folder_path,file_name):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZSE16"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGV_TABNAME").text = "EKKO"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/btn%_I4_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "UB"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").text = "ZMTO"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").text = "ZINC"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").text = "ZFTD"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").text = "WK"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").text = "WM"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").caretPosition = 2
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/btn%_I12_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "MX00"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "MXA1"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtI15-LOW").text = Date_From
    session.findById("wnd[0]/usr/ctxtI15-HIGH").text = Date_To
    session.findById("wnd[0]/usr/ctxtI12-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtI12-HIGH").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 31
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()


def Download_EKPO(df_ekko,folder_path):
    PO_EKPO = list()
    i = 0
    print('rows leidas en PO_List = ', str(f'{df_ekko.shape[0]:,}'))

    while i < df_ekko.shape[0]:
        count_plus20 = i + 20000
        File_name = f"EKPO_{i}_{count_plus20}.xlsx"  
        partial = 0

        while partial < 20000 and i < df_ekko.shape[0]:
            PO_EKPO.append(df_ekko.iloc[i,0])
            partial += 1
            i += 1

        my_data = pd.DataFrame(data=PO_EKPO)
        my_data.to_clipboard(index=False, index_label=False)
        print('Data in clipboard')
        print(str(f'{i:,}'),'of',str(f'{df_ekko.shape[0]:,}'))

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZSE16"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtGV_TABNAME").text = "EKPO"
        session.findById("wnd[0]/usr/ctxtGV_TABNAME").caretPosition = 4
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = File_name
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]").sendVKey (3)
        session.findById("wnd[0]").sendVKey (3)
        session.findById("wnd[0]").sendVKey (3)

        PO_EKPO = []
        try:
            os.system('taskkill /f /im excel.exe')
        except Exception as e:
            print(f'An exception occurred while tryign to close Excel: {e}')
        
    print('\n----------------------------------------------------------------------')
    print('############ ¡REPORT EKPO EXTRACTED SUCCESSFULLY! ##############')
    print('----------------------------------------------------------------------\n')


def Download_FBL3N(Date_From, Date_To, folder_path, file_name):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL3N"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "2010001"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "2010002"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "2010000"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "2033910"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "2031195"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "1026002"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "MX**"
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = Date_From
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "LO_FBL3N"
    session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 8
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]").sendVKey (16)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey (3)
    session.findById("wnd[0]").sendVKey (3)


def Download_FBL1N(Date_From, Date_To, folder_path, file_name):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "FBL1N"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[16]").press()
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN015_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "KG"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "KR"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "KA"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "RE"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "X1"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "ZX"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 2
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN015_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = "MX**"
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = Date_From
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Date_To
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "FBL1N_HRQ"
    session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 8
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]").sendVKey (16)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey (3)
    session.findById("wnd[0]").sendVKey (3)


def Execute_EKKO(folder_path, Date_From, Date_To):
    print('\n- DOWNLOADING EKKO -\n')
    From_Time = datetime.strptime(Date_From, "%d.%m.%Y")
    To_Time = datetime.strptime(Date_To, "%d.%m.%Y")
    if To_Time > From_Time or To_Time == From_Time:
        days_r = 60
        counter_Ex = 0
        Sec_1 = From_Time
        Sec_2 = From_Time + timedelta(days= days_r)
        New_From = Date_From
        days_r = days_r + 1

        while (To_Time > Sec_2):
            New_From = Sec_1.strftime('%d.%m.%Y')
            New_To = Sec_2.strftime('%d.%m.%Y')
            counter_Ex += 1
            print(f"EXECUTION {counter_Ex} from {New_From} to {New_To} --In progress--")
            file_name = f"EKKO_{str(New_From)}_{str(New_To)}.xlsx"
            Download_EKKO(New_From, New_To, folder_path, file_name)
            Sec_1 =  Sec_2 + timedelta(days= 1)
            Sec_2 = Sec_2 + timedelta(days= days_r)
            
            try:
                os.system('taskkill /f /im excel.exe')
            except Exception as e:
                print(f'\nAn exception occurred while tryging to close Excel: {e}\n')
                
    New_From = Sec_1.strftime('%d.%m.%Y')
    counter_Ex += 1
    print(f"EXECUTION {counter_Ex} from {New_From} to {Date_To} --Last One--")
    file_name = f"EKKO_{str(New_From)}_{str(Date_To)}.xlsx"
    Download_EKKO(New_From,Date_To,folder_path,file_name)

    try:
        os.system('taskkill /f /im excel.exe')
    except Exception as e:
        print(f'\nAn exception occurred while tryging to close Excel: {e}\n')
    wait_and_continue()


# === FUNCIONES UTILITARIAS ===
def wait_and_continue():
    time.sleep(1.5)

def close_excel_process():
    try:
        os.system('taskkill /f /im excel.exe')
    except Exception as e:
        print(f'Error closing Excel: {e}')

# === FUNCIONES PRINCIPALES ===
def execute_download_report(folder_path, Date_From, Date_To, days_r, file_prefix, download_function):
    From_Time = datetime.strptime(Date_From, "%d.%m.%Y")
    To_Time = datetime.strptime(Date_To, "%d.%m.%Y")

    if To_Time < From_Time:
        print("Error: La fecha final es anterior a la inicial.")
        return

    counter_Ex = 0
    Sec_1 = From_Time
    Sec_2 = From_Time + timedelta(days=days_r)
    days_r += 1

    while To_Time > Sec_2:
        New_From = Sec_1.strftime('%d.%m.%Y')
        New_To = Sec_2.strftime('%d.%m.%Y')
        counter_Ex += 1
        print(f'EXECUTION {counter_Ex}: from {New_From} to {New_To} --In progress--')

        file_name = f"{file_prefix}_{New_From}_{New_To}.xlsx"
        download_function(New_From, New_To, folder_path, file_name)
        close_excel_process()

        Sec_1 = Sec_2 + timedelta(days=1)
        Sec_2 = Sec_2 + timedelta(days=days_r)

    New_From = Sec_1.strftime('%d.%m.%Y')
    counter_Ex += 1
    print(f'EXECUTION {counter_Ex}: from {New_From} to {Date_To} --Last One--')

    file_name = f"{file_prefix}_{New_From}_{Date_To}.xlsx"
    download_function(New_From, Date_To, folder_path, file_name)
    close_excel_process()

    print('-----------------------------------------------------')
    print(f'############### ¡{file_prefix} COMPLETED! ################')
    print('-----------------------------------------------------')

def collect_concat_files(folder_path, report_name, keep_columns):
    xlsx_files = [
        pd.read_excel(os.path.join(root, file), usecols=keep_columns)
        for root, _, files in os.walk(folder_path)
        for file in files
        if file.startswith(report_name) and file.endswith('.xlsx')
    ]
    if not xlsx_files:
        print(f"Warning: No files found for {report_name} in {folder_path}")
        return pd.DataFrame()
    df = pd.concat(xlsx_files, ignore_index=True)
    return df

def concat_final_db(folder_path, report_name, df):
    output_path = os.path.join(folder_path, f"{report_name}/{report_name}_YT/{report_name} DB.csv")
    old_db_path = os.path.join(folder_path, f"{report_name}/{report_name}_YT/{report_name} DB.csv")

    if os.path.exists(old_db_path):
        chunks = pd.read_csv(old_db_path, chunksize=1000000, low_memory=False)
        df_old_db = pd.concat(chunks, ignore_index=True)
    else:
        df_old_db = pd.DataFrame()

    df_new_db = pd.concat([df_old_db, df], ignore_index=True)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df_new_db.to_csv(output_path, index=False)
    print(f"Final database saved to {output_path}")

def clean_transf_ekko(folder_path, Date_From, Date_To, month, year, months_spanish):
    ekko_path = os.path.join(folder_path, f"EKKO/{months_spanish[month]}_{year}")
    os.makedirs(ekko_path, exist_ok=True)
    Execute_EKKO(ekko_path, Date_From, Date_To)

    print('\n- Concatenating the EKKO files -\n')
    keep_columns_ekko = ['EBELN','BUKRS','BSTYP','BSART','AEDAT','ERNAM','LPONR','LIFNR',
                         'ZTERM','ZBD1T','EKORG','EKGRP','WAERS','WKURS','BEDAT','FRGSX',
                         'FRGKE','RLWRT']
    df_ekko = collect_concat_files(ekko_path, 'EKKO', keep_columns_ekko)

    df_ekko['EBELN'] = df_ekko['EBELN'].astype(str)
    df_ekko['BEDAT'] = pd.to_datetime(df_ekko['BEDAT'], dayfirst=True, errors='coerce')
    df_ekko = df_ekko[df_ekko['EBELN'] != '']
    list_EKKO_PO = df_ekko['EBELN'].unique()
    Total_unique_PO_EKKO = len(list_EKKO_PO)
    print(f"\nTotal created POs between {Date_From} to {Date_To}: {Total_unique_PO_EKKO:,}\n")
    return df_ekko, list_EKKO_PO, Total_unique_PO_EKKO

def clean_transf_ekpo(df_ekko, folder_path, month, year, months_spanish):
    print('\n- DOWNLOADING EKPO -')
    ekpo_path = os.path.join(folder_path, f"EKPO/{months_spanish[month]}_{year}")
    os.makedirs(ekpo_path, exist_ok=True)
    Download_EKPO(df_ekko, ekpo_path)
    wait_and_continue()

    print('- CREATING EKPO DATA FRAME -')
    keep_columns_EKPO = ['EBELN','EBELP','LOEKZ','TXZ01','MATNR','BUKRS','WERKS','MATKL',
                         'INFNR','MENGE','MEINS','BPRME','NETPR','PEINH','NETWR','EFFWR',
                         'ELIKZ','KONNR','KTPNR','BANFN','BNFPO','MTART','LEWED','STATUS',
                         'KNTTP','EREKZ','PSTYP']
    df_ekpo = collect_concat_files(ekpo_path, 'EKPO', keep_columns_EKPO)
    print(f"Shape of the EKPO report: # Columns = {df_ekpo.shape[1]:,} & # Rows = {df_ekpo.shape[0]:,}")

    df_ekpo['EBELP'] = df_ekpo['EBELP'].astype(int)
    df_ekpo['EBELN'] = df_ekpo['EBELN'].astype(str)
    df_ekpo['LEWED'] = pd.to_datetime(df_ekpo['LEWED'], errors='coerce')

    return df_ekpo

def clean_transf_fbl1n(folder_path, Date_From, Date_To, month, year, months_spanish):
    print('\n- DOWNLOADING FBL1N -\n')
    fbl1n_path = os.path.join(folder_path, f"FBL1N/{months_spanish[month]}_{year}")
    os.makedirs(fbl1n_path, exist_ok=True)
    execute_download_report(fbl1n_path, Date_From, Date_To, 60, 'FBL1N', Download_FBL1N)
    df_fbl1n = collect_concat_files(fbl1n_path, 'FBL1N', None)
    if df_fbl1n.empty:
        print("Warning: FBL1N data is empty.")
    return df_fbl1n

def clean_transf_fbl3n(folder_path, Date_From, Date_To, month, year, months_spanish):
    print('\n- DOWNLOADING FBL3N -\n')
    fbl3n_path = os.path.join(folder_path, f"FBL3N/{months_spanish[month]}_{year}")
    os.makedirs(fbl3n_path, exist_ok=True)
    execute_download_report(fbl3n_path, Date_From, Date_To, 16, 'FBL3N', Download_FBL3N)
    df_fbl3n = collect_concat_files(fbl3n_path, 'FBL3N', None)
    if df_fbl3n.empty:
        print("Warning: FBL3N data is empty.")
    return df_fbl3n

def concat_yt(folder_path, df_ekko, df_ekpo, df_fbl1n, df_fbl3n):
    print("===== Concatenating EKKO DB =====")
    concat_final_db(folder_path, 'EKKO', df_ekko)
    print("===== Concatenating EKPO DB =====")
    concat_final_db(folder_path, 'EKPO', df_ekpo)
    print("===== Concatenating FBL1N DB =====")
    concat_final_db(folder_path, 'FBL1N', df_fbl1n)
    print("===== Concatenating FBL3N DB =====")
    concat_final_db(folder_path, 'FBL3N', df_fbl3n)
    print("===== Concatenation YT DB Completed =====")