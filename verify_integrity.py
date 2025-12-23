
import pandas as pd
import os
import glob
import warnings
import traceback

# Suppress warnings
warnings.filterwarnings("ignore")

BASE_DIR = "/Users/beto1821uol.com.br/Library/CloudStorage/OneDrive-Personal/Atual/analise grafo"

FILE_PATTERNS = [
    {"pattern": "*1º SEMESTRE 2024*.xlsx", "year": 2024, "semester": 1, "engine": "openpyxl", "header_row": 1},
    {"pattern": "*2º SEMESTRE 2024*.xlsb", "year": 2024, "semester": 2, "engine": "pyxlsb", "header_row": 0}, 
    {"pattern": "*1º SEMESTRE 2025*.xlsb", "year": 2025, "semester": 1, "engine": "pyxlsb", "header_row": 0},
    {"pattern": "*2º SEMESTRE 2025*.xlsb", "year": 2025, "semester": 2, "engine": "pyxlsb", "header_row": 0}
]

COLUMN_PRIORITIES = [
    ("Valor_Unitario", ["R$ FINAL", "R$ RESMA", "R$ TOTAL", "VALOR"]),
    ("Empresa", ["VENCEDOR", "RAZÃO SOCIAL", "PARCEIRO", "FORNECEDOR"]),
    ("Marca", ["MARCA"]),
    ("Volume", ["VOLUME (RESMAS)", "VOLUME", "QUANTIDADE", "QTD"])
]

BLACKLIST_TERMS = {
    "Empresa": ["ANTERIOR", "STATUS", "SITUAÇÃO", "RESULTADO", "COLOCAÇÃO", "ULTIMO"],
    "Valor_Unitario": ["ANTERIOR", "ESTIMADO", "DIFERENÇA"]
}

STATUS_KEYWORDS_SET = {"GANHAMOS", "PERDEMOS", "SUSPENSA", "SUSPENSO", "ADIADO", "ADIOU", "CANCELADO", "FRACASSADO", "DESCLASSIFICADO", "NÃO PARTICIPAMOS"}

months_lookup = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4, 
    "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, 
    "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}

def dedup_columns(columns):
    seen = {}
    new_cols = []
    for col in columns:
        col_str = str(col).strip().upper()
        if col_str in seen:
            seen[col_str] += 1
            new_cols.append(f"{col_str}.{seen[col_str]}")
        else:
            seen[col_str] = 0
            new_cols.append(col_str)
    return new_cols

def is_status_column(series):
     if isinstance(series, pd.DataFrame):
         # If duplicate columns, take the first one
         series = series.iloc[:, 0]
         
     sample = series.dropna().astype(str).str.upper().str.strip()
     if sample.empty: return False
     match_count = sample.apply(lambda x: any(k in x for k in STATUS_KEYWORDS_SET) or x in STATUS_KEYWORDS_SET).sum()
     return (match_count / len(sample)) > 0.3

def detect_header_row(df_preview):
    keywords = ["DATA DO EVENTO", "NRO DO PREGÃO", "VOLUME", "VENCEDOR", "VALOR", "EMPRESA", "PARCEIRO", "R$ FINAL"]
    for i, row in df_preview.iterrows():
        row_vals = [str(x).upper() for x in row.values if pd.notna(x)]
        matches = 0
        for k in keywords:
            if any(k in val for val in row_vals):
                matches += 1
        if matches >= 3:
            return i
    return 0

def load_data_and_audit():
    processed_count_rdf = 0
    processed_count_atual = 0
    
    audit_log = []

    for info in FILE_PATTERNS:
        search_path = os.path.join(BASE_DIR, info["pattern"])
        found_files = glob.glob(search_path)
        
        if not found_files:
            continue
            
        actual_path = found_files[0]
        filename = os.path.basename(actual_path)
        print(f"Scanning {filename}...")
            
        try:
            if info["engine"] == "openpyxl":
                xl = pd.ExcelFile(actual_path, engine="openpyxl")
            else:
                xl = pd.ExcelFile(actual_path, engine="pyxlsb")

            for sheet in xl.sheet_names:
                upper_sheet = sheet.upper().strip()
                if upper_sheet not in months_lookup:
                    continue
                
                # Load with logic
                try:
                    df_preview = pd.read_excel(actual_path, sheet_name=sheet, engine=info["engine"], header=None, nrows=10)
                    header_idx = detect_header_row(df_preview)
                    df = pd.read_excel(actual_path, sheet_name=sheet, engine=info["engine"], header=header_idx)
                except Exception as e:
                    print(f"Error reading {sheet}: {e}")
                    continue

                # Ensure columns are unique strings
                raw_cols = [str(c).strip().upper() for c in df.columns]
                df.columns = dedup_columns(raw_cols)

                # --- MAPPING LOGIC ---
                rename_dict = {}
                found_targets = set()
                
                mapped_empresa_col = None

                for target, candidates in COLUMN_PRIORITIES:
                    best_match = None
                    valid_candidates = []
                    for candidate in candidates:
                        matches = [c for c in df.columns if candidate in c]
                        if target in BLACKLIST_TERMS:
                            matches = [c for c in matches if not any(bad in c for bad in BLACKLIST_TERMS[target])]
                        
                        if target == "Empresa":
                             filtered = []
                             for m in matches:
                                 if not is_status_column(df[m]):
                                    filtered.append(m)
                             matches = filtered
                        
                        if matches:
                             valid_candidates.extend(matches)
                    
                    if valid_candidates:
                         for candidate in candidates:
                             matches = [c for c in df.columns if candidate in c]
                             if target in BLACKLIST_TERMS:
                                matches = [c for c in matches if not any(bad in c for bad in BLACKLIST_TERMS[target])]
                             if target == "Empresa":
                                 filtered = []
                                 for m in matches:
                                     if not is_status_column(df[m]):
                                        filtered.append(m)
                                 matches = filtered
                             
                             if matches:
                                 best_match = min(matches, key=len)
                                 break
                    
                    if best_match:
                        rename_dict[best_match] = target
                        found_targets.add(target)
                        if target == "Empresa":
                            mapped_empresa_col = best_match

                df.rename(columns=rename_dict, inplace=True)
                
                if "Empresa" in df.columns:
                     # Clean and Categorize
                     df_proc = df[df["Empresa"].notna()].copy()
                     df_proc["Empresa_Clean"] = df_proc["Empresa"].astype(str).str.strip().str.upper()
                     
                     # Filter invalid status companies
                     status_keywords_raw = ["GANHAMOS", "PERDEMOS", "DESCLASSIFICADO", "FRACASSADO"]
                     df_proc = df_proc[~df_proc["Empresa_Clean"].isin(status_keywords_raw)]

                     rdf_in_proc = df_proc[df_proc["Empresa_Clean"].str.contains("RDF|RD F|R.D.F", regex=True, na=False)]
                     atual_in_proc = df_proc[df_proc["Empresa_Clean"].str.contains("ATUAL", na=False)]
                     
                     n_rdf = len(rdf_in_proc)
                     n_atual = len(atual_in_proc)
                     processed_count_rdf += n_rdf
                     processed_count_atual += n_atual
                     
                     audit_log.append(f"{sheet:<10} | {str(mapped_empresa_col):<30} | {n_rdf:<5} | {n_atual:<5}")
                else:
                     audit_log.append(f"{sheet:<10} | NOT FOUND                      | 0     | 0")

        except Exception as e:
            tb = traceback.format_exc()
            print(f"Error processing {filename}: {tb}")

    print("\n" + "="*80)
    print(f"{'SHEET':<10} | {'MAPPED COLUMN':<30} | {'RDF':<5} | {'ATUAL':<5}")
    print("="*80)
    for l in audit_log:
        print(l)
    print("="*80)
    print(f"TOTAL PROCESSED: RDF={processed_count_rdf}, ATUAL={processed_count_atual}")

if __name__ == "__main__":
    load_data_and_audit()
