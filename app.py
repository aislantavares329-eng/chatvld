import streamlit as st
import pandas as pd

# Configuração inicial do app
st.set_page_config(page_title="Detector de Padrões VLD", layout="wide")
st.title("📊 Detector de Padrões VLD")

# Dias da semana em português
DIAS_PT = {0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta", 5: "Sábado", 6: "Domingo"}

# -----------------------------
# Função preparar a base
# -----------------------------
def preparar_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip().upper() for c in df.columns]

    # Datas
    if "DATA" in df.columns:
        df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
        df["MES"] = df["DATA"].dt.to_period("M").astype(str)
        df["DIA_SEMANA"] = df["DATA"].dt.weekday.map(DIAS_PT)

    # Garantir que tempos sejam numéricos
    for col in ["TEMPO DE SOLUÇÃO", "TEMPO_DE_SOLUCAO", "TEMPO_DE_SOLUCAO_MIN", "PARADA_MIN"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df

# -----------------------------
# Função gerar relatório Excel com gráficos
# -----------------------------
def gerar_relatorio(df: pd.DataFrame, saida="relatorio.xlsx"):
    with pd.ExcelWriter(saida, engine="xlsxwriter") as writer:
        wb = writer.book

        # Aba base completa
        df.to_excel(writer, sheet_name="Base", index=False)

        # 1) Top Defeitos
        if "DEFEITO" in df.columns:
            top_defeitos = df["DEFEITO"].value_counts().rename_axis("DEFEITO").reset_index(name="QTD")
            top_defeitos.to_excel(writer, sheet_name="Top Defeitos", index=False)
            ws = writer.sheets["Top Defeitos"]
            chart = wb.add_chart({"type": "column"})
            chart.add_series({
                "categories": ["Top Defeitos", 1, 0, len(top_defeitos), 0],
                "values":     ["Top Defeitos", 1, 1, len(top_defeitos), 1],
                "name": "Ocorrências"
            })
            chart.set_title({"name": "Top Defeitos"})
            ws.insert_chart("D2", chart)

        # 2) Defeitos x Fábrica
        if set(["FÁBRICA", "DEFEITO"]).issubset(df.columns):
            defeito_fab = df.groupby(["FÁBRICA", "DEFEITO"]).size().reset_index(name="QTD")
            defeito_fab.to_excel(writer, sheet_name="Defeitos x Fábrica", index=False)
            ws = writer.sheets["Defeitos x Fábrica"]
            chart = wb.add_chart({"type": "column"})
            chart.add_series({
                "categories": ["Defeitos x Fábrica", 1, 1, len(defeito_fab), 1],
                "values":     ["Defeitos x Fábrica", 1, 2, len(defeito_fab), 2],
                "name": "Qtd por defeito"
            })
            chart.set_title({"name": "Defeitos por Fábrica"})
            ws.insert_chart("E2", chart)

        # 3) Tempo médio por defeito
        tempo_col = None
        for c in ["TEMPO DE SOLUÇÃO", "TEMPO_DE_SOLUCAO", "TEMPO_DE_SOLUCAO_MIN"]:
            if c in df.columns:
                tempo_col = c
                break
        if tempo_col and "DEFEITO" in df.columns:
            tempo_medio = (df.groupby("DEFEITO")[tempo_col].mean().round(1)
                           .reset_index(name="TEMPO_MEDIO_MIN"))
            tempo_medio.to_excel(writer, sheet_name="Tempo Médio", index=False)
            ws = writer.sheets["Tempo Médio"]
            chart = wb.add_chart({"type": "bar"})
            chart.add_series({
                "categories": ["Tempo Médio", 1, 0, len(tempo_medio), 0],
                "values":     ["Tempo Médio", 1, 1, len(tempo_medio), 1],
                "name": "Média (min)"
            })
            chart.set_title({"name": "Tempo Médio por Defeito"})
            ws.insert_chart("D2", chart)

        # 4) Defeito x Mês
        if set(["DEFEITO", "MES"]).issubset(df.columns):
            pivot = pd.pivot_table(df, index="DEFEITO", columns="MES",
                                   values=("FÁBRICA" if "FÁBRICA" in df.columns else "DEFEITO"),
                                   aggfunc="count", fill_value=0)
            pivot.to_excel(writer, sheet_name="Defeito x Mês")
            ws = writer.sheets["Defeito x Mês"]
            chart = wb.add_chart({"type": "line"})
            for i in range(len(pivot)):
                chart.add_series({
                    "name":       ["Defeito x Mês", i+1, 0],
                    "categories": ["Defeito x Mês", 0, 1, 0, len(pivot.columns)],
                    "values":     ["Defeito x Mês", i+1, 1, i+1, len(pivot.columns)],
                })
            chart.set_title({"name": "Ocorrências por Mês"})
            ws.insert_chart("B10", chart)

    return saida

# -----------------------------
# Interface Streamlit
# -----------------------------
uploaded_file = st.file_uploader("📂 Suba sua base (.csv ou .xlsx)", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Lê a base
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, encoding="utf-8", sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file)

    df = preparar_df(df)

    # Pré-visualização
    st.subheader("🔎 Pré-visualização da base")
    st.dataframe(df.head())

    # Top Defeitos (gráfico no app)
    if "DEFEITO" in df.columns:
        st.subheader("🔥 Top Defeitos")
        top_defeitos = df["DEFEITO"].value_counts().reset_index()
        top_defeitos.columns = ["Defeito", "Ocorrências"]
        st.bar_chart(top_defeitos.set_index("Defeito"))

    # Botão gerar Excel
    if st.button("📥 Gerar Relatório Excel"):
        saida = gerar_relatorio(df)
        with open(saida, "rb") as f:
            st.download_button("⬇️ Baixar Relatório", f, file_name="relatorio.xlsx")
