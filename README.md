# P2P Spend — SAP Extraction + Cleaning + GUI
A Windows‑only toolkit to extract P2P (Procure‑to‑Pay) data from SAP via SAP GUI Scripting, transform/merge it into a single dataset, and export PowerBI_DataBase.csv for analytics. It provides both a CLI pipeline and a simple Tkinter GUI, plus packaging to a single .exe with PyInstaller.

✨ Features

- Automates SAP downloads (EKKO, EKPO, FBL1N, FBL3N) in date chunks.

- Cleans and merges data into a unified dataset for Power BI.

- Saves outputs with month/year folders and appends to historical (YT) CSVs.

- GUI to select folder, enter date range, and watch progress logs.

- Optional one‑file .exe build for non‑technical users.

SPEND CODE/

├─ P2P_Spend.py            # CLI entrypoint: orchestrates full pipeline

├─ SAP_Spend.py            # SAP scripting, downloads, helpers, concat utils

├─ utils.py                # Folder/date prompts (Tk) & path helpers

├─ run_p2p_spend_gui.py    # Tkinter GUI launcher (calls same pipeline pieces)

├─ Catalogues/             # External Excel catalog files used in mapping

│  ├─ PMF_Category_NonPO_Final.xlsx  (sheets: "NON-PO PMF_Final (2)", "EKGRP_2")

│  └─ CAPEX 2018-2025.xlsx (WBS mapping)

└─ README.md               # This file


🖥️ Requirements

Windows with SAP GUI for Windows installed and configured.

SAP GUI Scripting enabled

On client: Options → Accessibility & Scripting → Enable scripting

On server: sapgui/user_scripting = TRUE (Basis setting)

Python 3.10+ (tested on 3.13.5) and pip

Access to the SAP variants used in scripts (examples):

FBL1N variant: FBL1N_HRQ

FBL3N variant: LO_FBL3N

ZSE16 access for EKKO/EKPO extracts

