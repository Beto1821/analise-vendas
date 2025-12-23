
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
TARGET_COMPANIES = ["RDF", "ATUAL PAPELARIA", "ATUAL", "RD F"] 

# Usando padrões para encontrar arquivos (mais robusto contra Unicode/Espaços)
FILE_PATTERNS = [
    {"pattern": "*1º SEMESTRE 2024*.xlsx", "year": 2024, "semester": 1, "engine": "openpyxl", "header_row": 1},
    {"pattern": "*2º SEMESTRE 2024*.xlsb", "year": 2024, "semester": 2, "engine": "pyxlsb", "header_row": 0}, 
    {"pattern": "*1º SEMESTRE 2025*.xlsb", "year": 2025, "semester": 1, "engine": "pyxlsb", "header_row": 0},
    {"pattern": "*2º SEMESTRE 2025*.xlsb", "year": 2025, "semester": 2, "engine": "pyxlsb", "header_row": 0}
]

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
        # Buscar arquivo real usando glob
        search_path = os.path.join(BASE_DIR, info["pattern"])
        found_files = glob.glob(search_path)
        
        if not found_files:
            debug_logs.append(f"ARQUIVO NÃO ENCONTRADO para padrão: {info['pattern']}")
            continue
            
        # Pega o primeiro match (deve ser único)
        actual_path = found_files[0]
        filename = os.path.basename(actual_path)
        debug_logs.append(f"Carregando: {filename}")
            
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
                
                df = pd.read_excel(actual_path, sheet_name=sheet, engine=info["engine"], header=info["header_row"])
                
                raw_cols = [str(c).strip().upper() for c in df.columns]
                df.columns = dedup_columns(raw_cols)
                
                rename_dict = {}
                for col in df.columns:
                    for k, v in COL_MAPPING.items():
                        if k in col: 
                            if v not in rename_dict.values():
                                rename_dict[col] = v
                
                df.rename(columns=rename_dict, inplace=True)
                
                cols_to_keep = ["Empresa", "Marca", "Valor_Unitario", "Volume"]
                final_cols_in_df = [c for c in cols_to_keep if c in df.columns]
                
                if final_cols_in_df:
                    subset_df = df[final_cols_in_df].copy()
                    subset_df["Ano"] = info["year"]
                    subset_df["Mes"] = month_num
                    all_data.append(subset_df)
                    
        except Exception as e:
            debug_logs.append(f"Erro ao ler {filename}: {str(e)}")

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
    
    # Categorização de Empresas
    df["Empresa_Clean"] = df["Empresa"].astype(str).str.upper().fillna("DESCONHECIDO")
    
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
    
    # Total Geral
    total_sales = df["Total_Venda"].sum()
    my_sales = df_filtered_my_companies["Total_Venda"].sum()
    share = (my_sales / total_sales * 100) if total_sales > 0 else 0
    
    insights.append(f"**Market Share Global**: As empresas RDF e ATUAL representam **{share:.2f}%** do faturamento total analisado (R$ {total_sales:,.2f}).")
    
    # Crescimento 2024 vs 2025 (Minhas Empresas)
    sales_by_year = df_filtered_my_companies.groupby("Ano")["Total_Venda"].sum()
    if 2024 in sales_by_year and 2025 in sales_by_year:
        growth = ((sales_by_year[2025] - sales_by_year[2024]) / sales_by_year[2024]) * 100
        trend = "CRESCIMENTO" if growth > 0 else "QUEDA"
        emoji = "xC4" if growth > 0 else "xCE"
        insights.append(f"**Comparativo Anual (RDF+ATUAL)**: Houve um(a) {emoji} **{trend} de {growth:.1f}%** em 2025 comparado a 2024 (considerando os meses disponíveis).")
        insights.append(f"- 2024: R$ {sales_by_year[2024]:,.2f}")
        insights.append(f"- 2025: R$ {sales_by_year[2025]:,.2f}")
    
    # Melhor Mês 2025
    df_2025 = df_filtered_my_companies[df_filtered_my_companies["Ano"] == 2025]
    if not df_2025.empty:
        best_month_num = df_2025.groupby("Mes")["Total_Venda"].sum().idxmax()
        best_month_val = df_2025.groupby("Mes")["Total_Venda"].sum().max()
        best_month_name = month_names.get(best_month_num, str(best_month_num))
        insights.append(f"**Destaque 2025**: O melhor mês para RDF+ATUAL em 2025 foi **{best_month_name}** com vendas de R$ {best_month_val:,.2f}.")

    return insights

# --- MAIN APP ---

st.title("xC4 Análise de Vendas: RDF & ATUAL vs Mercado")
st.markdown("Comparativo detalhado de performance 2024 vs 2025.")

# Carga de Dados
with st.spinner("Carregando e processando planilhas..."):
    raw_df, debug_logs = load_data()
    df = clean_and_process(raw_df)

# Sidebar Debug Logs
with st.sidebar.expander("Logs de Carregamento (Debug)"):
    for log in debug_logs:
        st.write(log)

if df.empty:
    st.error("Nenhum dado encontrado. Verifique os 'Logs de Carregamento' na barra lateral para entender o motivo.")
    st.stop()

# Filtros Sidebar
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

# Separar Dados
my_companies_df = filtered_df[filtered_df["Categoria"].isin(["RDF", "ATUAL"])]
others_df = filtered_df[filtered_df["Categoria"] == "OUTROS"]

# --- KPIS ---
col1, col2, col3 = st.columns(3)
total_market = filtered_df["Total_Venda"].sum()
total_mine = my_companies_df["Total_Venda"].sum()
total_others = others_df["Total_Venda"].sum()

col1.metric("Vendas Totais do Mercado", f"R$ {total_market:,.2f}")
col2.metric("Vendas RDF + ATUAL", f"R$ {total_mine:,.2f}", delta=f"{(total_mine/total_market)*100:.1f}% Share" if total_market else 0)
col3.metric("Vendas Outras Empresas", f"R$ {total_others:,.2f}")

st.divider()

# --- INSIGHTS AUTOMÁTICOS ---
st.subheader("xC9 Insights Gerados")
insights = generate_insights(df, df[(df["Categoria"].isin(["RDF", "ATUAL"]))]) # Insights globais baseados no dataset completo para contexto histórico
for i in insights:
    st.markdown(f"- {i}")

st.divider()

# --- VISUALIZAÇÕES ---

tab1, tab2, tab3 = st.tabs(["Comparativo Mensal", "Market Share", "Dados Brutos"])

with tab1:
    st.markdown("### Evolução de Vendas por Categoria")
    
    # Group by Year-Month-Category
    # Create a Date column for better plotting
    filtered_df["Data"] = pd.to_datetime(filtered_df.assign(Day=1)[["Ano", "Mes", "Day"]])
    
    monthly_cat = filtered_df.groupby(["Data", "Categoria"])["Total_Venda"].sum().reset_index()
    monthly_cat = monthly_cat.sort_values("Data")
    
    fig = px.bar(
        monthly_cat, 
        x="Data", 
        y="Total_Venda", 
        color="Categoria", 
        title="Vendas Mensais por Categoria (Comparativo)",
        labels={"Total_Venda": "Valor (R$)", "Categoria": "Empresa"},
        color_discrete_map={"RDF": "#1f77b4", "ATUAL": "#ff7f0e", "OUTROS": "#d62728"},
        barmode='stack'
    )
    fig.update_layout(xaxis_tickformat="%b %Y")
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("### Comparativo Anual (Apenas RDF e ATUAL)")
    # Pivot to compare 2024 vs 2025 side by side for months
    my_monthly = my_companies_df.groupby(["Mes", "Ano"])["Total_Venda"].sum().reset_index()
    my_monthly["Mes_Nome"] = my_monthly["Mes"].map(month_names)
    
    fig2 = px.bar(
        my_monthly,
        x="Mes_Nome",
        y="Total_Venda",
        color="Ano",
        barmode="group",
        title="RDF+ATUAL: Comparativo Mês a Mês (2024 vs 2025)",
        category_orders={"Mes_Nome": list(month_names.values())} # Ensure Month Order
    )
    st.plotly_chart(fig2, use_container_width=True)

with tab2:
    st.markdown("### Distribuição de Faturamento")
    
    # Donut Chart
    total_by_cat = filtered_df.groupby("Categoria")["Total_Venda"].sum().reset_index()
    fig_pie = px.pie(
        total_by_cat,
        values="Total_Venda",
        names="Categoria",
        title="Participação no Período Selecionado",
        color="Categoria",
        color_discrete_map={"RDF": "#1f77b4", "ATUAL": "#ff7f0e", "OUTROS": "#d62728"},
        hole=0.4
    )
    st.plotly_chart(fig_pie, use_container_width=True)
    
    st.markdown("### Top 10 'Outras Empresas'")
    top_others = others_df.groupby("Empresa_Clean")["Total_Venda"].sum().nlargest(10).reset_index()
    fig_bar_others = px.bar(
        top_others,
        x="Total_Venda",
        y="Empresa_Clean",
        orientation='h',
        title="Maiores Concorrentes (Outros)",
        text_auto='.2s'
    )
    fig_bar_others.update_layout(yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig_bar_others, use_container_width=True)

with tab3:
    st.markdown("### Detalhamento dos Dados")
    st.dataframe(filtered_df[["Ano", "Mes", "Empresa", "Categoria", "Valor_Unitario", "Volume", "Total_Venda"]].sort_values(["Ano", "Mes"]))
