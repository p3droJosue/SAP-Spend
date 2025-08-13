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
