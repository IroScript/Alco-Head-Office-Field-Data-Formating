# Alco Head Office Field Data Formatting

A lightweight Python GUI application that automates the mapping and formatting of hierarchical field force targets from Excel spreadsheets.

## Features
- **Automatic Hierarchy Parsing:** Automatically resolves parent-child assignments between field forces (MPO, AFM, MR, etc.) and managers (FM, AM, RSM, SH, etc.) by scanning upward from summary rows.
- **Dynamic Formula Shifting:** Replicates complex sub-total formulas in Column I by applying relative cell reference shifting corresponding to Column E formulas.
- **Exceptions & Special Cases Handling:** Resolves custom formatting rules for special roles (e.g. `DIC`, standalone `FM(MPO)` units reporting directly to Sector Heads).
- **Tiny GUI Interface:** Provides a lightweight Tkinter window (occupies less than 5% of the screen) to select the input `.xlsx` file and destination directory easily.

## Requirements
Make sure you have Python 3 installed. You'll need the `openpyxl` library to read and write Excel files:

```bash
pip install openpyxl
```

## How to Run
Run the script using the terminal:

```bash
python automate_excel.py
```

Choose your target file and output folder using the GUI, then click **Automate Formatting** to generate your formatted file.
