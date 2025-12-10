import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Gerador AGF", layout="wide")

st.title("Gerador de Relação Nominal – Instituto AGF")

st.write("""
Este aplicativo gera a planilha **no layout oficial do Instituto AGF**, usando apenas a planilha de funcionários.
Todos os funcionários com **Data de Desligamento preenchida serão excluídos** automaticamente.
""")

uploaded_file = st.file_uploader("Carregar planilha de funcionários (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Pré-visualização da base carregada")
    st.dataframe(df.head())

    # Normalizar nomes das colunas para evitar problemas
    df.columns = [col.strip().lower() for col in df.columns]

    # Verificar colunas obrigatórias
    colunas_necessarias = [
        "cpf", "nome", "rg", "matricula", "cargo", "sindicato", 
        "data de desligamento"
    ]

    faltando = [c for c in colunas_necessarias if c not in df.columns]

    if faltando:
        st.error(f"As seguintes colunas estão faltando na planilha: {faltando}")
    else:
        # Filtrar funcionários ativos (sem desligamento)
        df_filtrado = df[df["data de desligamento"].isna()].copy()

        # Montar o layout final do AGF
        df_final = pd.DataFrame({
            "CPF": df_filtrado["cpf"],
            "NOME": df_filtrado["nome"],
            "RG": df_filtrado["rg"],
            "RE": df_filtrado["matricula"],
            "FUNCAO": df_filtrado["cargo"],
            "MUNICIPIO_PRESTACAO_SERVICO": df_filtrado["sindicato"],
            "CNPJ_EMPREGADOR": "65.035.552/0001-80"
        })

        st.subheader("Registros que serão enviados ao Instituto AGF")
        st.dataframe(df_final)

        # Gerar arquivo Excel para download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="AGF")

        output.seek(0)

        st.download_button(
            label="Baixar Planilha Final (Excel)",
            data=output,
            file_name="relacao_nominal_AGF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success(f"Total de funcionários incluídos: {len(df_final)}")
