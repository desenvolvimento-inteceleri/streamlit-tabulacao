#Juntar varias abas de sheets em uma só
#Padrao de colunas: Ano	Nome	Escola	Pontuação	Tempo	Deficiência/Transtorno	ETAPA


import streamlit as st
import pandas as pd
from io import BytesIO

def main():
    st.title("Combinar Dados de Múltiplas Abas em apenas uma")

    st.write("Exemplo da estrutura de dados esperada de upload:") 
    # Exemplo de tabela
    exemplo_data = {
        "Ano": ["1ª ANO", "1ª ANO", "1ª ANO"],
        "Nome": ["ALUNO 1", "ALUNO 2", "ALUNO 3"],
        "Escola": ["ESCOLA 1", "ESCOLA 2", "ESCOLA 3"],
        "Pontuação": [45, 45, 45],
        "Tempo": ["00:15:00", "00:11:00", "00:15:00"],
        "Se for aluno com deficiência/transtorno:": ["NÃO POSSUI DEFICIÊNCIA/TRANSTORNO"] * 3,
        "Etapa de Classificação": ["1º CLASSIFICATÓRIA"] * 3
    }
    exemplo_df = pd.DataFrame(exemplo_data)
    st.table(exemplo_df)

    # Carregar o arquivo Excel via upload
    uploaded_file = st.file_uploader("Carregue o arquivo do Google Sheets em formato Excel (.xlsx)", type="xlsx")

    if uploaded_file is not None:
        # Nome das colunas que você deseja extrair
        colunas_desejadas = [
            "Ano", "Nome", "Escola", "Pontuação", "Tempo", 
            "Se for aluno com deficiência/transtorno:", "Etapa de Classificação"
        ]

        # Inicializar um DataFrame vazio para armazenar todos os dados
        all_data = pd.DataFrame(columns=colunas_desejadas)

        # Carregar todas as sheets do arquivo
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)

        # Iterar sobre cada sheet
        for sheet_name, data in all_sheets.items():
            # Ignorar as duas primeiras linhas
            data = data.iloc[2:].reset_index(drop=True)

            # Renomear as colunas para garantir que estamos pegando as corretas
            data.columns = colunas_desejadas

            # Converter todos os dados para maiúsculas
            data = data.applymap(lambda x: x.upper() if isinstance(x, str) else x)

            # Concatenar a sheet ao DataFrame principal
            all_data = pd.concat([all_data, data], ignore_index=True)

        # Exibir os dados combinados no Streamlit
        if not all_data.empty:
            st.write("Dados combinados de todas as sheets:")
            st.dataframe(all_data)

            # Botão para download do DataFrame combinado em CSV
            csv = all_data.to_csv(index=False).encode('utf-8')
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                all_data.to_excel(writer, index=False, sheet_name='Dados Combinados')

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Baixar dados combinados em CSV",
                    data=csv,
                    file_name="dados_combinados.csv",
                    mime="text/csv",
                )
            with col2:
                st.download_button(
                    label="Baixar dados combinados em Excel",
                    data=excel_buffer.getvalue(),
                    file_name="dados_combinados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.warning("Nenhuma aba foi carregada corretamente.")
    else:
        st.info("Por favor, carregue um arquivo .xlsx para continuar.")
