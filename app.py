import streamlit as st
import os
import pandas as pd

# Função para carregar todas as planilhas Excel de um diretório
def carregar_planilhas(diretorio):
    extensoes_excel = ['.xlsx', '.xls', '.xlsm', '.xlsb']
    arquivos_excel = [os.path.join(diretorio, f) for f in os.listdir(diretorio) if any(f.endswith(ext) for ext in extensoes_excel) and not f.startswith('~$')]
    planilhas = {}
    for arquivo in arquivos_excel:
        try:
            if arquivo.endswith('.xls'):
                planilhas[os.path.basename(arquivo)] = pd.read_excel(arquivo, sheet_name=0, engine='xlrd')
            else:
                planilhas[os.path.basename(arquivo)] = pd.read_excel(arquivo, sheet_name=0)
        except Exception as e:
            st.write(f"Erro ao carregar {arquivo}: {e}")
    return planilhas

# Outras funções...

def executar_verificacao(diretorio, arquivo_saida):
    planilhas = carregar_planilhas(diretorio)
    if len(planilhas) < 2:
        st.write("É necessário pelo menos duas planilhas para a verificação.")
        return
    try:
        duplicados = verificar_requisitos(planilhas)
        if duplicados.empty:
            st.write("Nenhum requisito duplicado encontrado.")
        else:
            duplicados.to_excel(arquivo_saida, index=False)
            st.write(f"Verificação concluída. Resultado salvo em: {arquivo_saida}")
            st.dataframe(duplicados)
    except ValueError as e:
        st.write(f"Erro na verificação: {e}")
    except Exception as e:
        st.write(f"Ocorreu um erro inesperado: {e}")

st.title('Verificação de Planilhas')
diretorio_planilhas = st.text_input('Caminho do Diretório das Planilhas', '/Users/arizona/Documents/agily/planilhas antigas')
arquivo_saida = st.text_input('Caminho do Arquivo de Saída', '/Users/arizona/Documents/agily/resultados_verificacao.xlsx')

if st.button('Executar Verificação'):
    executar_verificacao(diretorio_planilhas, arquivo_saida)
