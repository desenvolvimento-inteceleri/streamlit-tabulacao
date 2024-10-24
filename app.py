import streamlit as st
import pandas as pd
import xlsxwriter

# Função para filtrar o formulário de resposta e gerar o arquivo classificatória
def gerar_classificatoria(formulario_df):
    # Colunas que você deseja manter do formulário de resposta
    colunas_selecionadas = [
        'Nome do aluno?',
        'Qual é o nome da sua escola?',
        'Escreva o nome da escola caso ela no esteja listada',
        'Ano escolar do aluno:',
        'Total de pontuação?',
        'Quanto tempo de realização?',
        'Se for aluno com deficiência/transtorno:'
    ]
    
    # Filtrando as colunas que você deseja do formulário de resposta
    filtrado_df = formulario_df[colunas_selecionadas]
    
    # Substituindo o valor "Escola não está na lista" pelo valor da coluna "Escreva o nome da escola caso ela no esteja listada"
    filtrado_df['Qual é o nome da sua escola?'] = filtrado_df.apply(
        lambda row: row['Escreva o nome da escola caso ela no esteja listada']
        if row['Qual é o nome da sua escola?'] == "Escola não está na lista"
        else row['Qual é o nome da sua escola?'],
        axis=1
    )
    
    # Remover a coluna "Escreva o nome da escola caso ela no esteja listada" após a substituição
    filtrado_df = filtrado_df.drop(columns=['Escreva o nome da escola caso ela no esteja listada'])
    
    # Renomeando as colunas para o formato esperado do arquivo classificatória
    filtrado_df.columns = [
        'Nome',
        'Escola',
        'Ano',
        'Pontuação',
        'Tempo',
        'Se for aluno com deficiência/transtorno:'
    ]
    
    # Garantir que as colunas "Nome" e "Escola" estejam em letras maiúsculas
    filtrado_df['Nome'] = filtrado_df['Nome'].str.upper()
    filtrado_df['Escola'] = filtrado_df['Escola'].str.upper()
    
    # Criar uma coluna extra chamada ETAPA (com valores fictícios de exemplo)
    filtrado_df['ETAPA'] = '2° CLASSIFICATÓRIA'
    
    # Reordenando as colunas conforme solicitado
    filtrado_df = filtrado_df[['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Se for aluno com deficiência/transtorno:', 'ETAPA']]
    
    return filtrado_df

# Função para salvar o arquivo com escolas separadas em diferentes sheets
def salvar_excel_separado_por_escola(classificatoria_df):
    # Nome do arquivo a ser gerado
    output_path = "classificatoria_por_escola.xlsx"
    
    # Usando o XlsxWriter para personalizar a saída com múltiplas planilhas
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Separar por escolas
        escolas = classificatoria_df['Escola'].unique()
        
        for escola in escolas:
            # Tratar valores nulos na coluna "Escola" e garantir que a string não ultrapasse 31 caracteres
            escola = str(escola)[:31] if pd.notna(escola) else "ESCOLA_DESCONHECIDA"
            
            # Filtrar os dados para cada escola
            df_escola = classificatoria_df[classificatoria_df['Escola'] == escola]
            
            # Adiciona a aba (sheet) para cada escola
            df_escola.to_excel(writer, sheet_name=escola, index=False, startrow=1)  # Escreve a partir da linha 2
            
            # Acessar o worksheet da escola atual
            worksheet = writer.sheets[escola]
            
            # Mesclar as células na primeira linha e centralizar o nome da escola
            worksheet.merge_range('A1:G1', escola, writer.book.add_format({'align': 'center', 'bold': True}))
    
    return output_path

# Função principal do aplicativo Streamlit
def main():
    st.title("Gerador de Classificatória por Escola")
    
    # Upload do arquivo formulário de resposta
    uploaded_file = st.file_uploader("Envie o arquivo do Formulário de Resposta", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Carregar o arquivo Excel enviado pelo usuário
        formulario_df = pd.read_excel(uploaded_file)
        
        # Gerar o arquivo classificatória com base no formulário de resposta filtrado
        classificatoria_df = gerar_classificatoria(formulario_df)
        
        # Mostrar a tabela filtrada na tela
        st.write("Dados filtrados para o arquivo classificatória:")
        st.dataframe(classificatoria_df)
        
        # Botão para baixar o arquivo classificatória gerado
        if st.button("Gerar o arquivo Classificatória"):
            # Salvando o arquivo com cada escola em sua própria aba (sheet)
            output_path = salvar_excel_separado_por_escola(classificatoria_df)
            
            # Gerando o botão de download
            with open(output_path, "rb") as file:
                st.download_button(
                    label="Baixar Classificatória",
                    data=file,
                    file_name="classificatoria_por_escola.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# Executa a aplicação no Streamlit
if __name__ == '__main__':
    main()
