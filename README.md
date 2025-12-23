
# 📊 Dashboard de Análise de Vendas: 2024 vs 2025

Este projeto consiste em um dashboard interativo desenvolvido em **Python** e **Streamlit** para analisar e comparar a performance de vendas entre os anos de **2024** e **2025**.

O foco principal é o monitoramento das empresas **RDF** e **ATUAL**, permitindo uma visão clara de seu *Market Share* (participação de mercado) frente aos concorrentes ("Outras Empresas").

## 🎯 Objetivos
- **Comparativo Anual**: Analisar o crescimento ou retração das vendas em 2025 comparado a 2024.
- **Market Share**: Medir a representatividade das empresas do grupo no mercado total analisado.
- **Integridade dos Dados**: Garantir a leitura correta de múltiplas planilhas Excel com formatos variados (cabeçalhos dinâmicos).

## 🚀 Funcionalidades
- **Filtros Dinâmicos**: Seleção de Anos e Meses na barra lateral.
- **KPIs**: Indicadores de Vendas Totais, Vendas do Grupo e Vendas de Concorrentes.
- **Gráficos Interativos**:
    - Evolução Mensal de Vendas (Barras por Categoria).
    - Gráfico de Pizza de Participação de Mercado.
- **Insights Automáticos**: Geração de comentários textuais sobre tendências de crescimento.
- **Inspector de Dados**: Aba para auditoria e visualização dos dados brutos carregados.

## 🛠️ Tecnologias Utilizadas
- **Streamlit**: Interface web interativa.
- **Pandas**: Manipulação e limpeza de dados (ETL).
- **Plotly**: Visualização de dados.
- **OpenPyXL / PyXLSB**: Leitura de arquivos Excel (.xlsx e .xlsb).

## ⚙️ Como Executar
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Execute a aplicação:
   ```bash
   streamlit run streamlit_app.py
   ```
3. O dashboard abrirá automaticamente no seu navegador.

## 📂 Estrutura de Arquivos
- `streamlit_app.py`: Código principal da aplicação.
- `verify_integrity.py`: Script auxiliar para auditoria de dados (conta ocorrências de RDF/ATUAL).
- `requirements.txt`: Lista de bibliotecas necessárias.
