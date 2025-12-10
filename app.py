import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Gerador AGF", layout="wide")

# Funções auxiliares
def validar_cpf(cpf):
    """Valida CPF e retorna True se válido"""
    if pd.isna(cpf):
        return False
    
    # Remove caracteres não numéricos
    cpf_limpo = re.sub(r'\D', '', str(cpf))
    
    # Adicionar zeros à esquerda se necessário
    cpf_limpo = cpf_limpo.zfill(11)
    
    # Verifica se tem 11 dígitos
    if len(cpf_limpo) != 11:
        return False
    
    # Verifica se todos os dígitos são iguais
    if cpf_limpo == cpf_limpo[0] * 11:
        return False
    
    # Valida primeiro dígito verificador
    soma = sum(int(cpf_limpo[i]) * (10 - i) for i in range(9))
    digito1 = (soma * 10 % 11) % 10
    
    if int(cpf_limpo[9]) != digito1:
        return False
    
    # Valida segundo dígito verificador
    soma = sum(int(cpf_limpo[i]) * (11 - i) for i in range(10))
    digito2 = (soma * 10 % 11) % 10
    
    if int(cpf_limpo[10]) != digito2:
        return False
    
    return True

def formatar_cpf(cpf):
    """Formata CPF no padrão 000.000.000-00"""
    if pd.isna(cpf):
        return ""
    
    cpf_limpo = re.sub(r'\D', '', str(cpf))
    
    # Adicionar zeros à esquerda se necessário para completar 11 dígitos
    cpf_limpo = cpf_limpo.zfill(11)
    
    if len(cpf_limpo) == 11:
        return f"{cpf_limpo[:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:]}"
    
    return str(cpf)

def limpar_rg(rg):
    """Remove caracteres especiais do RG"""
    if pd.isna(rg):
        return ""
    
    return re.sub(r'\D', '', str(rg))

def validar_campo_obrigatorio(valor):
    """Verifica se campo obrigatório está preenchido"""
    if pd.isna(valor) or str(valor).strip() == "":
        return False
    return True

st.set_page_config(page_title="Gerador AGF", layout="wide")

st.title("Gerador de Relação Nominal – Instituto AGF")

st.write("""
Este aplicativo gera a planilha **no layout oficial do Instituto AGF**, usando apenas a planilha de funcionários.
Todos os funcionários com **Data de Desligamento preenchida serão excluídos** automaticamente.
""")

# Configurações
with st.expander("⚙️ Configurações Avançadas"):
    cnpj_empregador = st.text_input(
        "CNPJ do Empregador",
        value="65.035.552/0001-80",
        help="Informe o CNPJ que será usado na coluna CNPJ_EMPREGADOR"
    )
    
    incluir_desligados = st.checkbox(
        "Incluir funcionários desligados",
        value=False,
        help="Marque para incluir também os funcionários com data de desligamento"
    )
    
    validar_cpf_checkbox = st.checkbox(
        "Validar CPF (dígitos verificadores)",
        value=True,
        help="Valida se os CPFs são válidos conforme algoritmo oficial"
    )

uploaded_file = st.file_uploader("Carregar planilha de funcionários (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Forçar CPF e RG como texto para preservar zeros à esquerda
        # Salvar temporariamente para usar xlsxwriter
        temp_file = BytesIO(uploaded_file.read())
        uploaded_file.seek(0)
        
        try:
            df = pd.read_excel(uploaded_file, dtype={'cpf': str, 'rg': str, 'matricula': str})
        except ImportError:
            # Se openpyxl não estiver disponível, tentar sem especificar engine
            df = pd.read_excel(temp_file, dtype={'cpf': str, 'rg': str, 'matricula': str})

        st.subheader("Pré-visualização da base carregada")
        st.dataframe(df.head())
        
        total_registros = len(df)
        st.info(f"📊 Total de registros carregados: {total_registros}")

        # Normalizar nomes das colunas para evitar problemas
        df.columns = [col.strip().lower() for col in df.columns]
        
        # Converter CPF, RG e Matricula para string para preservar zeros à esquerda
        if 'cpf' in df.columns:
            df['cpf'] = df['cpf'].astype(str)
        if 'rg' in df.columns:
            df['rg'] = df['rg'].astype(str)
        if 'matricula' in df.columns:
            df['matricula'] = df['matricula'].astype(str)

        # Verificar colunas obrigatórias
        colunas_necessarias = [
            "cpf", "nome", "rg", "matricula", "cargo", "sindicato", 
            "data de desligamento"
        ]

        faltando = [c for c in colunas_necessarias if c not in df.columns]

        if faltando:
            st.error(f"❌ As seguintes colunas estão faltando na planilha: {', '.join(faltando)}")
        else:
            # Filtrar funcionários ativos (sem desligamento)
            if incluir_desligados:
                df_filtrado = df.copy()
                st.warning("⚠️ Incluindo todos os funcionários (ativos e desligados)")
            else:
                df_filtrado = df[df["data de desligamento"].isna()].copy()
                desligados = total_registros - len(df_filtrado)
                if desligados > 0:
                    st.warning(f"⚠️ {desligados} funcionário(s) desligado(s) foram excluídos")
            
            # Validações e limpeza de dados
            problemas = []
            indices_invalidos = []
            
            for idx, row in df_filtrado.iterrows():
                erros_linha = []
                
                # Validar campos obrigatórios
                if not validar_campo_obrigatorio(row["cpf"]):
                    erros_linha.append("CPF vazio")
                elif validar_cpf_checkbox and not validar_cpf(row["cpf"]):
                    erros_linha.append("CPF inválido")
                
                if not validar_campo_obrigatorio(row["nome"]):
                    erros_linha.append("Nome vazio")
                
                if not validar_campo_obrigatorio(row["rg"]):
                    erros_linha.append("RG vazio")
                
                if not validar_campo_obrigatorio(row["matricula"]):
                    erros_linha.append("Matrícula vazia")
                
                if not validar_campo_obrigatorio(row["cargo"]):
                    erros_linha.append("Cargo vazio")
                
                if not validar_campo_obrigatorio(row["sindicato"]):
                    erros_linha.append("Sindicato vazio")
                
                if erros_linha:
                    nome_funcionario = row["nome"] if validar_campo_obrigatorio(row["nome"]) else "Nome não informado"
                    problemas.append(f"Linha {idx + 2}: {nome_funcionario} - {', '.join(erros_linha)}")
                    indices_invalidos.append(idx)
            
            # Remover registros com problemas
            if indices_invalidos:
                st.error(f"❌ {len(indices_invalidos)} registro(s) com problemas foram identificados:")
                with st.expander("Ver detalhes dos problemas"):
                    for problema in problemas:
                        st.write(f"- {problema}")
                
                df_filtrado = df_filtrado.drop(indices_invalidos)
                st.info(f"✅ Continuando com {len(df_filtrado)} registro(s) válido(s)")
            
            if len(df_filtrado) == 0:
                st.error("❌ Nenhum registro válido para processar!")
            else:
                # Aplicar formatações
                df_filtrado["cpf_formatado"] = df_filtrado["cpf"].apply(formatar_cpf)
                df_filtrado["rg_limpo"] = df_filtrado["rg"].apply(limpar_rg)
                df_filtrado["nome_limpo"] = df_filtrado["nome"].str.strip().str.upper()

                # Montar o layout final do AGF
                df_final = pd.DataFrame({
                    "CPF": df_filtrado["cpf_formatado"],
                    "NOME": df_filtrado["nome_limpo"],
                    "RG": df_filtrado["rg_limpo"],
                    "RE": df_filtrado["matricula"],
                    "FUNCAO": df_filtrado["cargo"],
                    "MUNICIPIO_PRESTACAO_SERVICO": df_filtrado["sindicato"],
                    "CNPJ_EMPREGADOR": cnpj_empregador
                })

                st.subheader("✅ Registros que serão enviados ao Instituto AGF")
                st.dataframe(df_final)

                # Gerar arquivo Excel para download
                output = BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df_final.to_excel(writer, index=False, sheet_name="AGF")

                output.seek(0)

                st.download_button(
                    label="📥 Baixar Planilha Final (Excel 97-2003)",
                    data=output,
                    file_name="relacao_nominal_AGF.xls",
                    mime="application/vnd.ms-excel"
                )

                st.success(f"✅ Total de funcionários incluídos: {len(df_final)}")
                
                # Resumo do processamento
                with st.expander("📊 Resumo do Processamento"):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Registros Carregados", total_registros)
                    
                    with col2:
                        excluidos = total_registros - len(df_final)
                        st.metric("Registros Excluídos", excluidos)
                    
                    with col3:
                        st.metric("Registros Válidos", len(df_final))
    
    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {str(e)}")
        st.info("Verifique se o arquivo está no formato correto (.xlsx) e não está corrompido.")
