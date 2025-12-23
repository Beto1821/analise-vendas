
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

# CORREÇÃO: 2024 S1 tem header na linha 1 (index 1). Outros parecem ser linha 0.
FILE_PATTERNS = [
    {"pattern": "*1º SEMESTRE 2024*.xlsx", "year": 2024, "semester": 1, "engine": "openpyxl", "header_row": 1},
    {"pattern": "*2º SEMESTRE 2024*.xlsb", "year": 2024, "semester": 2, "engine": "pyxlsb", "header_row": 0}, 
    {"pattern": "*1º SEMESTRE 2025*.xlsb", "year": 2025, "semester": 1, "engine": "pyxlsb", "header_row": 0},
    {"pattern": "*2º SEMESTRE 2025*.xlsb", "year": 2025, "semester": 2, "engine": "pyxlsb", "header_row": 0}
]

COL_MAPPING = {
    "VENCEDOR": "Empresa",
    "RAZÃO SOCIAL": "Empresa", # CORREÇÃO: 2024 usa Razão Social
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
                # Check extension roughly
                xl = pd.ExcelFile(actual_path, engine="openpyxl")
                sheet_names = xl.sheet_names
            else:
                xl = pd.ExcelFile(actual_path, engine="pyxlsb")
                sheet_names = xl.sheet_names

            for sheet in sheet_names:
                upper_sheet = sheet.upper().strip()
                month_num = months_lookup.get(upper_sheet)
                
                if not month_num:
                    continue 
                
                # Ler dados
                df = pd.read_excel(actual_path, sheet_name=sheet, engine=info["engine"], header=info["header_row"])
                
                # Dedup colunas
                raw_cols = [str(c).strip().upper() for c in df.columns]
                df.columns = dedup_columns(raw_cols)
                
                # Rename
                rename_dict = {}
                for col in df.columns:
                    # Proteção: Ignorar 'Vencedor Ultimo Evento' ou similar
                    if "ULTIMO EVENTO" in col:
                        continue 
                        
                    for k, v in COL_MAPPING.items():
                        if k in col: 
                            # Prioridade: Se já mapeou 'Empresa', não sobrescreve fácil, a menos que seja match melhor?
                            # Simplificação: Primeiro match vence. RAZAO SOCIAL aparece, pega.
                            if v not in rename_dict.values():
                                rename_dict[col] = v
                                break
                
                df.rename(columns=rename_dict, inplace=True)
                
                cols_to_keep = ["Empresa", "Marca", "Valor_Unitario", "Volume"]
                final_cols_in_df = [c for c in cols_to_keep if c in df.columns]
                
                if final_cols_in_df:
                    subset_df = df[final_cols_in_df].copy()
                    subset_df["Ano"] = info["year"]
                    subset_df["Mes"] = month_num
                    subset_df["Origem"] = filename
                    all_data.append(subset_df)
                else:
                    debug_logs.append(f"MISSING COLS in {filename} [{sheet}]. Found: {df.columns.tolist()}")
                    
        except Exception as e:
            debug_logs.append(f"ERROR reading {filename}: {str(e)}")

    if not all_data:
        return pd.DataFrame(), debug_logs

    full_df = pd.concat(all_data, ignore_index=True)
    return full_df, debug_logs

@st.cache_data
def clean_and_process(df):
    if df.empty: return df
        
    # Limpeza de Valores
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
    df = clean_and_process(raw_df)

# Sidebar Debug
with st.sidebar.expander("Debug Logs", expanded=False):
    for log in debug_logs:
        st.write(log)
    if not df.empty:
        st.write("Anos Carregados:", df["Ano"].unique())
        st.write("Categorias:", df["Categoria"].value_counts())

if df.empty:
    st.error("Nenhum dado encontrado.")
    st.stop()

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
    
    st.dataframe(others_df.head(50))
