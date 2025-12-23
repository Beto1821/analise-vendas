
import pandas as pd
import os

base_dir = "/Users/beto1821uol.com.br/Library/CloudStorage/OneDrive-Personal/Atual/analise grafo"
files = [
    "COBERTURA DE PREÇOS 1º SEMESTRE 2024.xlsx",
    "COBERTURA DE PREÇOS 1º SEMESTRE 2025.xlsb"
]

for f in files:
    path = os.path.join(base_dir, f)
    print(f"\n--- INSPECTING SHEET 'JANEIRO' IN {f} ---")
    try:
        if f.endswith('.xlsx'):
            df = pd.read_excel(path, sheet_name='JANEIRO', nrows=10)
        else:
            df = pd.read_excel(path, engine='pyxlsb', sheet_name='JANEIRO', nrows=10)
        
        print(df.head(10))
        print("Columns:", df.columns.tolist())
    except Exception as e:
        print(f"Error reading {f}: {e}")
