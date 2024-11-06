# up dos 2 arquivos (1 e 2 etapas),
#gera um arquivo unido esses dois
#classifica os 10 melhores (ou qtd que o usuario desejar)

import streamlit as st
import pandas as pd
from io import BytesIO

# Função para carregar e ordenar os dados de uma sheet específica
def carregar_e_ordenar_dados_por_sheet(df, etapa):
    # Adicionar coluna para identificar a fase
    df['Etapa'] = etapa
    
    # Ordenar por Ano, Pontuação (decrescente), Tempo (crescente) e Etapa
    df = df.sort_values(by=["Ano", "Pontuação", "Tempo", "Etapa"], ascending=[True, False, True, True])
    
    return df

# Função principal do aplicativo
def main():
    st.title("Organizador de Classificatórias para Múltiplas Escolas")
    
    st.write("Carregue os arquivos Excel das 1ª e 2ª classificatórias.")
    
    # Upload dos arquivos
    arquivo1 = st.file_uploader("Upload da 1ª Classificatória", type=["xlsx"])
    arquivo2 = st.file_uploader("Upload da 2ª Classificatória", type=["xlsx"])
    
    # Verifica se ambos os arquivos foram carregados
    if arquivo1 and arquivo2:
        # Carrega todas as sheets em dicionários de DataFrames
        sheets_1 = pd.read_excel(arquivo1, sheet_name=None, header=1)
        sheets_2 = pd.read_excel(arquivo2, sheet_name=None, header=1)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name in sheets_1.keys():
                # Verifica se a sheet está presente em ambos os arquivos
                if sheet_name in sheets_2:
                    # Carrega e processa cada sheet
                    df1 = carregar_e_ordenar_dados_por_sheet(sheets_1[sheet_name], '1ª CLASSIFICATÓRIA')
                    df2 = carregar_e_ordenar_dados_por_sheet(sheets_2[sheet_name], '2ª CLASSIFICATÓRIA')
                    
                    # Concatenar as duas etapas
                    df_total = pd.concat([df1, df2], ignore_index=True)
                    
                    # Escreve o nome da escola como cabeçalho e a tabela ordenada para cada aba
                    df_total.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                    worksheet = writer.sheets[sheet_name]
                    worksheet.merge_range('A1:G1', sheet_name, writer.book.add_format({'align': 'center', 'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white'}))
                else:
                    st.warning(f"Sheet '{sheet_name}' não encontrada em ambos os arquivos.")
        
        output.seek(0)
        
        # Botão para download do arquivo organizado
        st.download_button(
            label="Baixar Classificação Organizada para Todas as Escolas",
            data=output,
            file_name="classificacao_organizada_todas_escolas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
