import streamlit as st
import pandas as pd

st.set_page_config(page_title="Analisador Dinâmico de Planilhas", layout="wide")
st.title("📊 Analisador Dinâmico de Planilhas")

# Upload
uploaded_file = st.file_uploader("📂 Suba sua planilha (.xlsx ou .csv)", type=["xlsx", "csv"])

if uploaded_file is not None:
    # Detectar abas
    if uploaded_file.name.endswith(".xlsx"):
        xls = pd.ExcelFile(uploaded_file)
        aba = st.selectbox("📑 Escolha a aba", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=aba)
    else:
        df = pd.read_csv(uploaded_file, sep=None, engine="python")

    st.subheader("🔎 Pré-visualização")
    st.dataframe(df.head())

    # Listar colunas
    cols = df.columns.tolist()
    st.write("📋 Colunas detectadas:", cols)

    # Seleção de colunas para análise categórica
    col_analise = st.selectbox("👉 Escolha uma coluna categórica para contar valores", cols)
    col_grupo = st.selectbox("👉 (Opcional) Agrupar por outra coluna", ["Nenhum"] + cols)

    # Frequência
    freq = None
    if col_analise:
        st.subheader(f"📊 Frequência de valores em: {col_analise}")
        if col_grupo != "Nenhum":
            freq = df.groupby(col_grupo)[col_analise].value_counts().unstack(fill_value=0)
            st.bar_chart(freq)
        else:
            freq = df[col_analise].value_counts().reset_index()
            freq.columns = [col_analise, "Ocorrências"]
            st.bar_chart(freq.set_index(col_analise))

    # Análise numérica automática
    num_cols = df.select_dtypes(include="number").columns.tolist()
    desc, corr = None, None
    if num_cols:
        st.subheader("📈 Estatísticas de colunas numéricas")
        desc = df[num_cols].describe().T[["mean", "50%", "std", "min", "max"]]
        desc.rename(columns={"mean": "Média", "50%": "Mediana", "std": "Desvio Padrão",
                             "min": "Mínimo", "max": "Máximo"}, inplace=True)
        st.dataframe(desc)

        st.subheader("🔗 Correlação entre variáveis numéricas")
        corr = df[num_cols].corr()
        st.dataframe(corr)
        st.line_chart(corr)

    # Exportar relatório Excel com gráficos
    if st.button("📥 Gerar Relatório Excel"):
        saida = "relatorio_dinamico.xlsx"
        with pd.ExcelWriter(saida, engine="xlsxwriter") as writer:
            wb = writer.book

            # Base
            df.to_excel(writer, sheet_name="Base", index=False)

            # Frequência
            if freq is not None:
                freq.to_excel(writer, sheet_name="Analise", index=True)
                ws = writer.sheets["Analise"]
                chart = wb.add_chart({"type": "column"})
                chart.add_series({
                    "categories": ["Analise", 1, 0, len(freq), 0],
                    "values": ["Analise", 1, 1, len(freq), 1],
                    "name": "Frequência"
                })
                chart.set_title({"name": f"Frequência de {col_analise}"})
                ws.insert_chart("E2", chart)

            # Estatísticas
            if desc is not None:
                desc.to_excel(writer, sheet_name="Estatísticas")
                ws = writer.sheets["Estatísticas"]
                chart = wb.add_chart({"type": "column"})
                chart.add_series({
                    "categories": ["Estatísticas", 1, 0, len(desc), 0],
                    "values": ["Estatísticas", 1, 1, len(desc), 1],
                    "name": "Média"
                })
                chart.set_title({"name": "Médias Numéricas"})
                ws.insert_chart("H2", chart)

            # Correlação
            if corr is not None:
                corr.to_excel(writer, sheet_name="Correlações")
                ws = writer.sheets["Correlações"]
                chart = wb.add_chart({"type": "line"})
                for i in range(len(corr)):
                    chart.add_series({
                        "name": ["Correlações", i+1, 0],
                        "categories": ["Correlações", 0, 1, 0, len(corr.columns)],
                        "values": ["Correlações", i+1, 1, i+1, len(corr.columns)]
                    })
                chart.set_title({"name": "Correlação entre Variáveis"})
                ws.insert_chart("B10", chart)

        with open(saida, "rb") as f:
            st.download_button("⬇️ Baixar Relatório", f, file_name=saida)
