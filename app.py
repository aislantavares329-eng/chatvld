import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

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

    cols = df.columns.tolist()
    st.write("📋 Colunas detectadas:", cols)

    # -----------------------------
    # Correlação entre duas colunas numéricas
    # -----------------------------
    col_x = st.selectbox("👉 Escolha a primeira coluna numérica (X)", cols)
    col_y = st.selectbox("👉 Escolha a segunda coluna numérica (Y)", cols)

    corr_val, insight, df_corr = None, None, None

    if col_x and col_y:
        try:
            df_corr = df[[col_x, col_y]].dropna()
            df_corr[col_x] = pd.to_numeric(df_corr[col_x], errors="coerce")
            df_corr[col_y] = pd.to_numeric(df_corr[col_y], errors="coerce")
            df_corr = df_corr.dropna()

            if not df_corr.empty:
                st.subheader(f"📉 Correlação entre {col_x} e {col_y}")
                fig, ax = plt.subplots()
                ax.scatter(df_corr[col_x], df_corr[col_y], alpha=0.6)
                ax.set_xlabel(col_x)
                ax.set_ylabel(col_y)
                st.pyplot(fig)

                corr_val = df_corr[col_x].corr(df_corr[col_y])
                st.write(f"🔗 Correlação de Pearson: **{corr_val:.2f}**")

                if corr_val > 0.7:
                    insight = "📈 Forte correlação positiva → quando X aumenta, Y tende a aumentar."
                elif corr_val < -0.7:
                    insight = "📉 Forte correlação negativa → quando X aumenta, Y tende a diminuir."
                elif -0.3 < corr_val < 0.3:
                    insight = "⚪ Correlação fraca ou inexistente → não há padrão claro."
                else:
                    insight = "🟡 Correlação moderada → existe relação, mas não muito forte."

                st.write(insight)

        except Exception as e:
            st.error(f"Erro ao calcular correlação: {e}")

    # -----------------------------
    # Relação categórica tipo FÁBRICA x DEFEITO
    # -----------------------------
    st.subheader("📊 Relação entre duas colunas categóricas")
    col_a = st.selectbox("👉 Escolha a primeira coluna categórica", cols)
    col_b = st.selectbox("👉 Escolha a segunda coluna categórica", cols)

    relacao, diag = None, None
    if col_a and col_b:
        relacao = df.groupby([col_a, col_b]).size().reset_index(name="QTD")

        if not relacao.empty:
            pivot = relacao.pivot(index=col_a, columns=col_b, values="QTD").fillna(0)
            st.bar_chart(pivot)

            st.subheader(f"🥧 Distribuição de {col_b}")
            dist = df[col_b].value_counts()
            st.pyplot(dist.plot.pie(autopct="%1.1f%%", figsize=(5, 5)).get_figure())

            maior = relacao.loc[relacao["QTD"].idxmax()]
            diag = (
                f"⚠️ Diagnóstico Preventivo:\n\n"
                f"- A combinação **{maior[col_a]} x {maior[col_b]}** apresentou **{maior['QTD']} ocorrências**.\n"
                f"- Recomenda-se intensificar manutenção preventiva em **{maior[col_a]}**, "
                f"com foco em evitar novos casos de **{maior[col_b]}**."
            )
            st.info(diag)

    # -----------------------------
    # Exportar para Excel
    # -----------------------------
    if st.button("📥 Gerar Relatório Excel"):
        saida = "relatorio_dinamico.xlsx"
        with pd.ExcelWriter(saida, engine="xlsxwriter") as writer:
            wb = writer.book

            # Aba base
            df.to_excel(writer, sheet_name="Base", index=False)

            # Aba Correlação
            if df_corr is not None and corr_val is not None:
                df_corr.to_excel(writer, sheet_name="Correlação", index=False)
                ws = writer.sheets["Correlação"]
                ws.write(len(df_corr) + 2, 0, "Coeficiente de Correlação (Pearson):")
                ws.write(len(df_corr) + 2, 1, corr_val)
                if insight:
                    ws.write(len(df_corr) + 3, 0, "Insight:")
                    ws.write(len(df_corr) + 3, 1, insight)

                chart = wb.add_chart({"type": "scatter"})
                chart.add_series({
                    "categories": ["Correlação", 1, 0, len(df_corr), 0],
                    "values": ["Correlação", 1, 1, len(df_corr), 1],
                    "name": f"{col_x} vs {col_y}"
                })
                chart.set_title({"name": f"Dispersão: {col_x} x {col_y}"})
                ws.insert_chart("E2", chart)

            # Aba Relação categórica
            if relacao is not None:
                relacao.to_excel(writer, sheet_name="Relação", index=False)
                ws = writer.sheets["Relação"]

                chart = wb.add_chart({"type": "column"})
                chart.add_series({
                    "categories": ["Relação", 1, 0, len(relacao), 0],
                    "values": ["Relação", 1, 2, len(relacao), 2],
                    "name": f"{col_a} x {col_b}"
                })
                chart.set_title({"name": f"{col_a} x {col_b}"})
                ws.insert_chart("E2", chart)

                if diag:
                    ws.write(len(relacao) + 3, 0, "Diagnóstico Preventivo:")
                    ws.write(len(relacao) + 4, 0, diag)

        with open(saida, "rb") as f:
            st.download_button("⬇️ Baixar Relatório", f, file_name=saida)
