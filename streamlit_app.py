
import streamlit as st
import pandas as pd
import os
import plotly.express as px
import plotly.graph_objects as go
import warnings
import glob

# Suppress warnings
warnings.filterwarnings("ignore")

# Configuração da Página
st.set_page_config(
    page_title="Dashboard de Vendas - Comparativo 2024/2025",
    page_icon="xC4",
    layout="wide"
)

# Constantes
BASE_DIR = "/Users/beto1821uol.com.br/Library/CloudStorage/OneDrive-Personal/Atual/analise grafo"

FILE_PATTERNS = [
    {"pattern": "*1* SEMESTRE 2024*.xlsx", "year": 2024, "semester": 1, "engine": "openpyxl", "header_row": 1},
    {"pattern": "*2* SEMESTRE 2024*.xlsb", "year": 2024, "semester": 2, "engine": "pyxlsb", "header_row": 0}, 
    {"pattern": "*1* SEMESTRE 2025*.xlsb", "year": 2025, "semester": 1, "engine": "pyxlsb", "header_row": 0},
    {"pattern": "*2* SEMESTRE 2025*.xlsb", "year": 2025, "semester": 2, "engine": "pyxlsb", "header_row": 0}
]

# ... existing code ...



# PRIORIDADES DE MAPEAMENTO (Ordem importa!)
# Lista de tuplas (Campo Destino, [Lista de Candidatos em Ordem de Prioridade])
COLUMN_PRIORITIES = [
    ("Valor_Unitario", ["R$ FINAL", "R$ RESMA", "R$ TOTAL", "VALOR"]),
    ("Empresa", ["VENCEDOR", "RAZÃO SOCIAL", "PARCEIRO", "FORNECEDOR"]),
    ("Marca", ["MARCA"]),
    ("Volume", ["VOLUME (RESMAS)", "VOLUME", "QUANTIDADE", "QTD"])
]

# Termos proibidos em nomes de colunas para certos campos
BLACKLIST_TERMS = {
    "Empresa": ["ANTERIOR", "STATUS", "SITUAÇÃO", "RESULTADO", "COLOCAÇÃO", "ULTIMO"],
    "Valor_Unitario": ["ANTERIOR", "ESTIMADO", "DIFERENÇA"]
}

months_lookup = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4, 
    "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, 
    "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}

month_names = {v: k for k, v in months_lookup.items() if k != "MARCO"}

def dedup_columns(columns):
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

@st.cache_data
def load_data():
    all_data = []
    debug_logs = []

    # Helper to find header row dynamically
    def detect_header_row(df_preview):
        # Keywords to identify the header row
        keywords = ["DATA DO EVENTO", "NRO DO PREGÃO", "VOLUME", "VENCEDOR", "VALOR", "EMPRESA", "PARCEIRO", "R$ FINAL"]
        
        for i, row in df_preview.iterrows():
            row_vals = [str(x).upper() for x in row.values if pd.notna(x)]
            matches = 0
            for k in keywords:
                if any(k in val for val in row_vals):
                    matches += 1
            
            # If we find enough keywords, assume this is the header
            if matches >= 3:
                return i
        return 0 # Fallback

    for info in FILE_PATTERNS:
        search_path = os.path.join(BASE_DIR, info["pattern"])
        found_files = glob.glob(search_path)
        
        if not found_files:
            debug_logs.append(f"ARQUIVO NÃO ENCONTRADO: {info['pattern']}")
            continue
            
        actual_path = found_files[0]
        filename = os.path.basename(actual_path)
            
        try:
            if info["engine"] == "openpyxl":
                # For xlsx in 2024/2025, headers can be inconsistent 
                xl = pd.ExcelFile(actual_path, engine="openpyxl")
                sheet_names = xl.sheet_names
            else:
                 # xlsb
                xl = pd.ExcelFile(actual_path, engine="pyxlsb")
                sheet_names = xl.sheet_names

            for sheet in sheet_names:
                upper_sheet = sheet.upper().strip()
                month_num = months_lookup.get(upper_sheet)
                
                if not month_num:
                    continue 
                
                # Step 1: Read valid preview to find header
                try:
                    df_preview = pd.read_excel(actual_path, sheet_name=sheet, engine=info["engine"], header=None, nrows=10)
                    header_idx = detect_header_row(df_preview)
                except Exception as e:
                    debug_logs.append(f"Error previewing {filename} [{sheet}]: {e}")
                    header_idx = info["header_row"] # Fallback to config
                
                # Step 2: Read full sheet with detected header
                try:
                    df = pd.read_excel(actual_path, sheet_name=sheet, engine=info["engine"], header=header_idx)
                except Exception as e:
                     debug_logs.append(f"Error reading {filename} [{sheet}] with header={header_idx}: {e}")
                     continue

                raw_cols = [str(c).strip().upper() for c in df.columns]
                df.columns = dedup_columns(raw_cols)
                
                # --- LOGICA DE MAPEAMENTO POR PRIORIDADE ---
                rename_dict = {}
                found_targets = set()
                
                # Status keywords to detect if a column is actually a status column
                STATUS_KEYWORDS_SET = {"GANHAMOS", "PERDEMOS", "SUSPENSA", "SUSPENSO", "ADIADO", "ADIOU", "CANCELADO", "FRACASSADO", "DESCLASSIFICADO", "NÃO PARTICIPAMOS"}

                def is_status_column(series):
                     # Check if a significant portion of the non-null values are status keywords
                     sample = series.dropna().astype(str).str.upper().str.strip()
                     if sample.empty: return False
                     
                     # Check precise matches or partial matches
                     match_count = sample.apply(lambda x: any(k in x for k in STATUS_KEYWORDS_SET) or x in STATUS_KEYWORDS_SET).sum()
                     return (match_count / len(sample)) > 0.3 # If >30% looks like status, it's a status column

                for target, candidates in COLUMN_PRIORITIES:
                    # Tenta encontrar o melhor candidato
                    best_match = None
                    valid_candidates = []

                    for candidate in candidates:
                        # Busca colunas que contêm o termo candidato
                        matches = [c for c in df.columns if candidate in c]
                        
                        # Filtrar blacklist
                        if target in BLACKLIST_TERMS:
                            matches = [c for c in matches if not any(bad in c for bad in BLACKLIST_TERMS[target])]
                        
                        # Apply content validation for 'Empresa'
                        if target == "Empresa":
                             filtered_matches = []
                             for m in matches:
                                 if not is_status_column(df[m]):
                                     filtered_matches.append(m)
                                 else:
                                     # Optionally log this rejection?
                                     pass
                             matches = filtered_matches
                        
                        if matches:
                             valid_candidates.extend(matches)
                    
                    if valid_candidates:
                        # Find the shortest match among all valid candidates found across all candidate keywords
                        # But we should prioritize the order of candidates (e.g. VENCEDOR > RAZÃO SOCIAL)
                        # So we take the matches from the *first* candidate keyword that produced valid matches
                        # Reworking the loop slightly to stop at first HIT
                        pass 

                    # Re-run logic correctly with early break:
                    for candidate in candidates:
                         matches = [c for c in df.columns if candidate in c]
                         if target in BLACKLIST_TERMS:
                            matches = [c for c in matches if not any(bad in c for bad in BLACKLIST_TERMS[target])]
                         
                         if target == "Empresa":
                             matches = [m for m in matches if not is_status_column(df[m])]
                         
                         if matches:
                             best_match = min(matches, key=len)
                             break
                    
                    if best_match:
                        rename_dict[best_match] = target
                        found_targets.add(target)
                
                df.rename(columns=rename_dict, inplace=True)
                
                # Validation
                missing = [t[0] for t in COLUMN_PRIORITIES if t[0] not in found_targets]
                if not missing:
                    cols_to_keep = ["Empresa", "Marca", "Valor_Unitario", "Volume"]
                    subset_df = df[cols_to_keep].copy()
                    subset_df["Ano"] = info["year"]
                    subset_df["Mes"] = month_num
                    subset_df["Origem"] = filename
                    all_data.append(subset_df)
                else:
                    debug_logs.append(f"MISSING {missing} in {filename} [{sheet}] (Header Row: {header_idx}). Found: {df.columns.tolist()}")
                    
        except Exception as e:
            debug_logs.append(f"ERROR reading {filename}: {str(e)}")

    if not all_data:
        return pd.DataFrame(), debug_logs

    full_df = pd.concat(all_data, ignore_index=True)
    return full_df, debug_logs

@st.cache_data
def clean_and_process(df):
    if df.empty: return df
        
    def clean_money(val):
        if pd.isna(val): return 0.0
        s = str(val).upper().replace("R$", "").replace(" ", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0
            
    df["Valor_Unitario"] = df["Valor_Unitario"].apply(clean_money)
    df = df[df["Valor_Unitario"] > 0]
    
    def clean_vol(val):
        if pd.isna(val): return 0.0
        s = str(val).upper().replace(" ", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0
            
    if "Volume" not in df.columns:
        df["Volume"] = 1.0
    else:
        df["Volume"] = df["Volume"].apply(clean_vol).replace(0, 1)
        
    df["Total_Venda"] = df["Valor_Unitario"] * df["Volume"]
    
    # Categorização
    df = df[df["Empresa"].notna()]
    df = df[df["Empresa"].astype(str).str.strip() != ""]
    
    df["Empresa_Clean"] = df["Empresa"].astype(str).str.upper()
    
    # Filtro Extra: Remover nomes de empresa que parecem status ("GANHAMOS", "PERDEMOS") caso tenha passado
    status_keywords = ["GANHAMOS", "PERDEMOS", "DESCLASSIFICADO", "FRACASSADO"]
    df = df[~df["Empresa_Clean"].isin(status_keywords)]
    
    def categorize(name):
        if any(t in name for t in ["RDF", "RD F", "R.D.F"]):
            return "RDF"
        elif "ATUAL" in name:
            return "ATUAL"
        else:
            return "OUTROS"
            
    df["Categoria"] = df["Empresa_Clean"].apply(categorize)
    
    return df

def generate_insights(df, df_filtered_my_companies):
    insights = []
    
    total_sales = df["Total_Venda"].sum()
    my_sales = df_filtered_my_companies["Total_Venda"].sum()
    share = (my_sales / total_sales * 100) if total_sales > 0 else 0
    
    insights.append(f"**Market Share Global**: As empresas RDF e ATUAL representam **{share:.2f}%** do faturamento total analisado (R$ {total_sales:,.2f}).")
    
    sales_by_year = df_filtered_my_companies.groupby("Ano")["Total_Venda"].sum()
    if 2024 in sales_by_year and 2025 in sales_by_year:
        growth = ((sales_by_year[2025] - sales_by_year[2024]) / sales_by_year[2024]) * 100
        trend = "CRESCIMENTO" if growth > 0 else "QUEDA"
        emoji = "xC4" if growth > 0 else "xCE"
        insights.append(f"**Comparativo Anual (RDF+ATUAL)**: Houve um(a) {emoji} **{trend} de {growth:.1f}%** em 2025 comparado a 2024.")
    
    return insights

# --- MAIN APP ---

st.title("xC4 Análise de Vendas: RDF & ATUAL vs Mercado")

# Carga de Dados
with st.spinner("Carregando planilhas..."):
    raw_df, debug_logs = load_data()
    # clean_and_process agora aplica os filtros de exclusão
    df = clean_and_process(raw_df)

# Sidebar Debug
with st.sidebar.expander("Debug Logs", expanded=False):
    for log in debug_logs:
        st.write(log)
    if not df.empty:
        st.write("Amostra de Categorias:", df["Categoria"].value_counts())

if df.empty:
    st.error("Nenhum dado encontrado após filtros.")
    
    st.markdown("### Diagnóstico do Servidor")
    st.write(f"Caminho Base: `{BASE_DIR}`")
    
    try:
        files_on_server = os.listdir(BASE_DIR)
        st.write("Arquivos encontrados na pasta:")
        st.code("\n".join(files_on_server))
    except Exception as e:
        st.error(f"Erro ao listar diretório: {e}")

    with st.stop():
        pass

# Filtros
st.sidebar.header("Filtros")
selected_years = st.sidebar.multiselect("Anos", options=sorted(df["Ano"].unique()), default=sorted(df["Ano"].unique()))
selected_months_nums = st.sidebar.multiselect(
    "Meses", 
    options=sorted(df["Mes"].unique()), 
    format_func=lambda x: month_names.get(x, str(x)),
    default=sorted(df["Mes"].unique())
)

filtered_df = df[
    (df["Ano"].isin(selected_years)) & 
    (df["Mes"].isin(selected_months_nums))
]

my_companies_df = filtered_df[filtered_df["Categoria"].isin(["RDF", "ATUAL"])]
others_df = filtered_df[filtered_df["Categoria"] == "OUTROS"]

col1, col2, col3 = st.columns(3)
total_market = filtered_df["Total_Venda"].sum()
total_mine = my_companies_df["Total_Venda"].sum()
total_others = others_df["Total_Venda"].sum()

col1.metric("Vendas Totais", f"R$ {total_market:,.2f}")
col2.metric("Vendas RDF + ATUAL", f"R$ {total_mine:,.2f}", delta=f"{(total_mine/total_market)*100:.1f}% Share" if total_market else 0)
col3.metric("Vendas Outras Empresas", f"R$ {total_others:,.2f}")

st.divider()

st.subheader("xC9 Insights")
insights = generate_insights(df, df[(df["Categoria"].isin(["RDF", "ATUAL"]))]) 
for i in insights:
    st.markdown(f"- {i}")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs(["Comparativo Mensal", "Market Share", "Dados Brutos", "Data Inspector (Debug)"])

with tab1:
    st.markdown("### Evolução Mensal")
    date_df = filtered_df.assign(year=filtered_df["Ano"], month=filtered_df["Mes"], day=1)
    filtered_df["Data"] = pd.to_datetime(date_df[["year", "month", "day"]])
    
    monthly_cat = filtered_df.groupby(["Data", "Categoria"])["Total_Venda"].sum().reset_index()
    monthly_cat = monthly_cat.sort_values("Data")
    
    fig = px.bar(
        monthly_cat, x="Data", y="Total_Venda", color="Categoria",
        title="Vendas Mensais por Categoria",
        color_discrete_map={"RDF": "#1f77b4", "ATUAL": "#ff7f0e", "OUTROS": "#d62728"}
    )
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.markdown("### Participação")
    total_by_cat = filtered_df.groupby("Categoria")["Total_Venda"].sum().reset_index()
    fig_pie = px.pie(
        total_by_cat, values="Total_Venda", names="Categoria", 
        color="Categoria",
        color_discrete_map={"RDF": "#1f77b4", "ATUAL": "#ff7f0e", "OUTROS": "#d62728"},
        hole=0.4
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with tab3:
    st.dataframe(filtered_df)

with tab4:
    st.markdown("### Inspeção de Arquivos")
    st.write("Anos encontrados:", df["Ano"].value_counts())
    st.write("Origem dos dados:", df["Origem"].value_counts())
    st.write("Columns in df:", df.columns.tolist())
    st.dataframe(others_df.head(50))
