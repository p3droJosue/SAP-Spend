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

⚙️ Installation
# (Recommended) create & activate a virtual env
python -m venv .venv
.venv\Scripts\activate

# install dependencies
pip install -r requirements.txt

▶️ Usage (CLI)
Run the full pipeline in console mode:
python P2P_Spend.py

You’ll be prompted for:
Download folder (root path where monthly subfolders & outputs go)

Date range: dd.mm.yyyy → dd.mm.yyyy

Outputs:

Spend Data Base/<Mes>_<Año>/PowerBI_DataBase.csv

Historical “YT” CSVs under each report’s _YT subfolder


🪟 Usage (GUI)
Launch the Tkinter app:
python run_p2p_spend_gui.py
Pick a destination folder

Enter Desde / Hasta dates (dd.mm.yyyy)

Click Ejecutar and watch progress in the log panel

Build single‑file .exe (optional)
pyinstaller --clean --onefile --windowed run_p2p_spend_gui.py
# Result: dist/run_p2p_spend_gui.exe


pyinstaller --clean --onefile --windowed run_p2p_spend_gui.py
# Result: dist/run_p2p_spend_gui.exe

Rebuild after code changes by re‑running the same command.


🧠 How it works (High‑level)

Chunked SAP downloads: date intervals are split to avoid SAP limits.

EKKO/EKPO: Purchase orders & items; copied via ZSE16 and clipboard batching.

FBL1N/FBL3N: Vendor and G/L line items; saved to XLSX by SAP GUI.

Cleaning: Type casts, key creation (PO_Item), date parsing, currency to float.

Merging: Non‑PO + PO joins (EKKO/EKPO), vendor/category/WBS lookups via catalog files.

Exports: PowerBI_DataBase.csv for the selected period, and append to long‑term CSVs.


🔧 Configuration hotspots (adjust for your SAP)

Company code mask in FBL1N / FBL3N: "MX**" (change if needed)

Document types (e.g., KG, KR, KA, RE, X1, ZX)

Account ranges used in FBL3N filter

Layout variants (FBL1N_HRQ, LO_FBL3N) — must exist in your SAP user

Date chunk size in downloads (e.g., 60 days for FBL1N; 16 for FBL3N)

Catalog files & sheet names in Catalogues/


🧪 Troubleshooting

InvalidIndexError during .map() → ensure catalog keys are unique. The GUI pipeline de‑duplicates via drop_duplicates before mapping.

FileNotFoundError when creating subfolders → Paths with mixed \\ and /. The code normalizes paths; ensure your base folder exists and you have write permissions.

Excel left open / locked files → The scripts force‑kill Excel after exports to avoid locks.

PyInstaller can’t find the script → run from the correct folder or pass the full path; delete build/, dist/, *.spec and rebuild with --clean if necessary.

Hidden imports (PyInstaller warnings) → add --hidden-import=<module> if a missing module is reported at runtime.
