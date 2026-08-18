# Econ Project

Python analysis for non-perishable purchasing costs, meal swipe demand, item-level regressions, and forecast outputs.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python3 Econ_proj.py
```

By default, the script reads `A.5_Non-Perishables.xlsx` from this folder and uses the `Data` sheet. You can pass a different workbook and sheet:

```bash
python3 Econ_proj.py path/to/file.xlsx "Data"
```

The script writes CSV, PNG, TXT, and HTML analysis outputs into the project folder.
