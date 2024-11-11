import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px

# Função para filtrar os melhores alunos de cada ano em uma aba específica
def filtrar_melhores_alunos(df, top_n):
    df.columns = df.columns.str.strip().str.title()  # Remove espaços e coloca em Title Case
    df = df.rename(columns={
        "Ano Escolar": "Ano",
        "Pontuação": "Pontuação",
        "Tempo": "Tempo"
    })
    
    # Ordena pelo ano, pontuação (decrescente), e tempo (crescente)
    df = df.sort_values(by=["Ano", "Pontuação", "Tempo"], ascending=[True, False, True])
    
    # Seleciona os melhores alunos de cada ano
    top_alunos_df = df.groupby("Ano").head(top_n)
    
    return top_alunos_df

# Função principal do aplicativo
def main():
    st.title("Classificação dos Melhores Alunos por escola")
    
    # Passo 1: Upload do arquivo
    uploaded_file = st.file_uploader("Carregue o arquivo Excel com os dados dos alunos", type=["xlsx"])
    
    if uploaded_file is not None:
        # Carrega todas as sheets em um dicionário de DataFrames
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=1)  # Pula a primeira linha de cabeçalho extra
        
        # Seleção do número de melhores alunos por ano para exibir
        top_n = st.selectbox("Escolha o número de melhores alunos por ano para exibir", [1, 2, 3, 4, 5])
        
        # Processamento dos dados e criação do arquivo Excel para download
        output = BytesIO()
        total_counts = pd.Series(dtype=int)  # Série para armazenar a contagem total de cada ano escolar
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Itera sobre cada aba e aplica a filtragem
            for sheet_name, df in all_sheets.items():
                try:
                    # Extrai o nome da escola da aba atual
                    nome_escola = sheet_name
                    
                    # Filtra os melhores alunos
                    top_alunos_df = filtrar_melhores_alunos(df, top_n)
                    
                    # Atualiza a contagem total de alunos por ano
                    total_counts = total_counts.add(top_alunos_df['Ano'].value_counts(), fill_value=0)
                    
                    # Insere o nome da escola como a primeira linha (cabeçalho)
                    top_alunos_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                    
                    # Adiciona o nome da escola na primeira linha com formatação
                    worksheet = writer.sheets[sheet_name]
                    format_center_bold = writer.book.add_format({'align': 'center', 'bold': True})
                    worksheet.merge_range('A1:G1', nome_escola, format_center_bold)
                    
                except KeyError as e:
                    st.error(f"Erro na aba {sheet_name}: Coluna {str(e)} não encontrada.")
        
        # Mover o ponteiro para o início do buffer
        output.seek(0)
        
        # Botão para download do arquivo Excel gerado
        st.download_button(
            label="Baixar Classificação dos Melhores Alunos",
            data=output,
            file_name="melhores_alunos_classificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Exibir gráfico de quantidade de alunos por unidade organizacional (Org Unit Path) interativo
        st.write("Quantidade total de alunos por Org Unit Path:")
        total_counts = total_counts.sort_index()  # Ordena os anos para o gráfico
        fig = px.bar(total_counts, x=total_counts.index, y=total_counts.values,
                     labels={'x': 'Ano', 'y': 'Quantidade'},
                     title="Quantidade por Ano")
        st.plotly_chart(fig)

        # Exibir a tabela de contagem de alunos por ano
        st.write("Tabela de Quantidade de Alunos por Ano:")
        table_data = pd.DataFrame({'Ano': total_counts.index, 'Quantidade': total_counts.values})
        st.dataframe(table_data)

if __name__ == '__main__':
    main()
