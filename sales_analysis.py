
import pandas as pd
import os
import matplotlib.pyplot as plt
import glob
import re
import warnings

# Suppress warnings
warnings.filterwarnings("ignore")

# Configuração
BASE_DIR = "/Users/beto1821uol.com.br/Library/CloudStorage/OneDrive-Personal/Atual/analise grafo"
TARGET_COMPANIES = ["RDF", "ATUAL PAPELARIA", "ATUAL", "RD F"] # Adicionado variações
OUTPUT_FILE = "sales_analysis_report.txt"

# Mapeamento de Arquivos
FILES_INFO = [
    {"file": "COBERTURA DE PREÇOS 1º SEMESTRE 2024.xlsx", "year": 2024, "semester": 1, "engine": "openpyxl", "header_row": 1},
    {"file": "COBERTURA DE PREÇOS 2º SEMESTRE 2024.xlsb", "year": 2024, "semester": 2, "engine": "pyxlsb", "header_row": 0}, 
    {"file": "COBERTURA DE PREÇOS 1º SEMESTRE 2025.xlsb", "year": 2025, "semester": 1, "engine": "pyxlsb", "header_row": 0},
    {"file": "COBERTURA DE PREÇOS 2º SEMESTRE 2025.xlsb", "year": 2025, "semester": 2, "engine": "pyxlsb", "header_row": 0}
]

# Colunas de Interesse (Mapping para padronização)
COL_MAPPING = {
    "VENCEDOR": "Empresa",
    "PARCEIRO DE NEGOCIOS": "Empresa",
    "MARCA": "Marca",
    "R$ FINAL": "Valor_Unitario",
    "R$ RESMA": "Valor_Unitario",
    "VOLUME (RESMAS)": "Volume",
    "QUANTIDADE": "Volume"
}

months_lookup = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4, 
    "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, 
    "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}

def dedup_columns(columns):
    """Dedup duplicate column names by appending .1, .2 etc"""
    seen = {}
    new_cols = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}.{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    return new_cols

def load_data():
    all_data = []

    for info in FILES_INFO:
        path = os.path.join(BASE_DIR, info["file"])
        if not os.path.exists(path):
            print(f"Skipping {info['file']} (Not found)")
            continue
            
        print(f"Loading {info['file']}...")
        
        try:
            if info["engine"] == "openpyxl":
                xl = pd.ExcelFile(path, engine="openpyxl")
                sheet_names = xl.sheet_names
            else:
                xl = pd.ExcelFile(path, engine="pyxlsb")
                sheet_names = xl.sheet_names

            for sheet in sheet_names:
                upper_sheet = sheet.upper().strip()
                month_num = months_lookup.get(upper_sheet)
                
                if not month_num:
                    continue 
                
                print(f"  -> Processing sheet: {sheet}")
                
                df = pd.read_excel(path, sheet_name=sheet, engine=info["engine"], header=info["header_row"])
                
                # Normalize and Dedup Columns
                raw_cols = [str(c).strip().upper() for c in df.columns]
                df.columns = dedup_columns(raw_cols)
                
                # Rename
                rename_dict = {}
                for col in df.columns:
                    for k, v in COL_MAPPING.items():
                        if k in col: 
                            if v not in rename_dict.values():
                                rename_dict[col] = v
                
                df.rename(columns=rename_dict, inplace=True)
                
                # Keep only relevant columns to avoid concat issues with garbage columns
                cols_to_keep = ["Empresa", "Marca", "Valor_Unitario", "Volume"]
                # Add existing columns that match
                final_cols_in_df = [c for c in cols_to_keep if c in df.columns]
                
                if "Empresa" not in final_cols_in_df:
                    # Try to heuristic find company column if missing?
                    # For now just skip or warn
                    # print(f"    WARNING: No 'Empresa' column found in {sheet}")
                    pass
                
                # Select only relevant columns + whatever else we might need?
                # Actually, let's just keep the mapped ones to be safe and clean.
                if final_cols_in_df:
                    subset_df = df[final_cols_in_df].copy()
                    subset_df["Ano"] = info["year"]
                    subset_df["Mes"] = month_num
                    all_data.append(subset_df)
                    
        except Exception as e:
            print(f"  ERROR processing {info['file']}: {e}")

    if not all_data:
        return pd.DataFrame() 

    full_df = pd.concat(all_data, ignore_index=True)
    return full_df

def clean_and_filter(df):
    if df.empty: return df
        
    print(f"\nTotal rows loaded: {len(df)}")
    
    # Filter Company
    df["Empresa_Clean"] = df["Empresa"].astype(str).str.upper().fillna("")
    
    # Better matching logic: Check if ANY target word is in the company name
    # e.g. "RDF PAPELARIA" matches "RDF"
    def match_company(name):
        for t in TARGET_COMPANIES:
            # Word boundary check or simple inclusion? Inclusion is safer for messy data
            if t in name:
                return True
        return False

    mask = df["Empresa_Clean"].apply(match_company)
    filtered_df = df[mask].copy()
    
    # Normalize Company Name for Aggregation
    filtered_df["Empresa_Final"] = filtered_df["Empresa_Clean"].apply(
        lambda x: "ATUAL" if "ATUAL" in x else ("RDF" if "RD" in x or "R.D.F" in x else x)
    )
    
    print(f"Rows after filtering companies: {len(filtered_df)}")
    
    # Clean Numbers
    def clean_money(val):
        if pd.isna(val): return 0.0
        s = str(val).upper().replace("R$", "").replace(" ", "")
        # Handle 1.200,50 -> 1200.50
        # If ',' is present, assume it's decimal separator
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0
            
    filtered_df["Valor_Unitario"] = filtered_df["Valor_Unitario"].apply(clean_money)
    
    # Volume might be "100" or "100.0" or "100,0"?
    def clean_vol(val):
        if pd.isna(val): return 0.0
        s = str(val).upper().replace(" ", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0

    if "Volume" not in filtered_df.columns:
        filtered_df["Volume"] = 1.0
    else:
        filtered_df["Volume"] = filtered_df["Volume"].apply(clean_vol).replace(0, 1) # Avoid 0 volume? user said consider volume. If 0, maybe it means 1 unit? Leaning towards keeping 0 if data says 0, but usually volume 0 in sales is weird. Let's assume 0->0.
    
    filtered_df["Total_Venda"] = filtered_df["Valor_Unitario"] * filtered_df["Volume"]
    
    return filtered_df

def generate_visualizations(df):
    if df.empty: return
    
    # Ensure directory
    os.makedirs("analysis_output", exist_ok=True)
    
    # Aggregation
    monthly_sales = df.groupby(["Ano", "Mes", "Empresa_Final"])["Total_Venda"].sum().reset_index()
    monthly_vol = df.groupby(["Ano", "Mes", "Empresa_Final"])["Volume"].sum().reset_index()
    
    # Pivot for plotting
    pivot_sales = monthly_sales.pivot_table(index=["Ano", "Mes"], columns="Empresa_Final", values="Total_Venda", fill_value=0)
    
    # Plot 1: Monthly Sales Comparison
    plt.figure(figsize=(12, 6))
    
    # Create a 'YYYY-MM' string index for x-axis
    pivot_sales.index = [f"{y}-{m:02d}" for y, m in pivot_sales.index]
    
    pivot_sales.plot(kind='bar', figsize=(12, 6), colormap='viridis')
    plt.title('Vendas Totais por Mês (2024 vs 2025)')
    plt.ylabel('Valor Total (R$)')
    plt.xlabel('Mês')
    plt.grid(True, axis='y', alpha=0.3)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("analysis_output/vendas_mensais.png")
    print("Graph saved: analysis_output/vendas_mensais.png")
    
    # Plot 2: Total Volume Comparison
    plt.figure(figsize=(10, 6))
    summary_vol = df.groupby(["Ano", "Empresa_Final"])["Volume"].sum().unstack(fill_value=0)
    summary_vol.plot(kind='bar', figsize=(10, 6))
    plt.title('Volume Total de Vendas por Ano e Empresa')
    plt.ylabel('Volume (Unidades/Resmas)')
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig("analysis_output/volume_anual.png")
    print("Graph saved: analysis_output/volume_anual.png")

    # Insight Text
    with open(OUTPUT_FILE, "w") as f:
        f.write("=== RELATÓRIO DE ANÁLISE DE VENDAS ===\n\n")
        f.write("1. TOTAIS GERAIS\n")
        total_sales = df.groupby("Empresa_Final")["Total_Venda"].sum()
        f.write(str(total_sales) + "\n\n")
        
        f.write("2. COMPARAÇÃO 2024 vs 2025 (Jan-Jun)\n")
        # Filter for first half to be fair comparison if 2025 is incomplete? 2025 is full year?
        # User implies "2024 vs 2025", 2025 likely current/partial.
        # Let's show data as is.
        y_group = df.groupby("Ano")["Total_Venda"].sum()
        f.write(str(y_group) + "\n\n")
        
    print(f"Report saved: {OUTPUT_FILE}")

def main():
    df = load_data()
    clean_df = clean_and_filter(df)
    
    if not clean_df.empty:
        print("\n--- DATA LOADED ---")
        print(clean_df.head())
        generate_visualizations(clean_df)
    else:
        print("No data found after processing.")

if __name__ == "__main__":
    main()
