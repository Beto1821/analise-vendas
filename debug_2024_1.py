
import pandas as pd
import os
import glob

BASE_DIR = "/Users/beto1821uol.com.br/Library/CloudStorage/OneDrive-Personal/Atual/analise grafo"
PATTERN = "*1º SEMESTRE 2024*.xlsx"

print(f"Searching for: {PATTERN}")
files = glob.glob(os.path.join(BASE_DIR, PATTERN))

if not files:
    print("No file found!")
else:
    f = files[0]
    print(f"File found: {f}")
    
    # Inspect Sheets
    xl = pd.ExcelFile(f, engine="openpyxl")
    print(f"Sheets: {xl.sheet_names}")
    
    # Inspect January content raw
    if "JANEIRO" in xl.sheet_names:
        print("\n--- RAW CONTENT (First 10 rows) ---")
        df = pd.read_excel(f, sheet_name="JANEIRO", header=None, nrows=10, engine="openpyxl")
        print(df)
        
        print("\n--- TRYING HEADER 0 ---")
        df0 = pd.read_excel(f, sheet_name="JANEIRO", header=0, nrows=5, engine="openpyxl")
        print(df0.columns)
        
        print("\n--- TRYING HEADER 1 ---")
        df1 = pd.read_excel(f, sheet_name="JANEIRO", header=1, nrows=5, engine="openpyxl")
        print(df1.columns)
