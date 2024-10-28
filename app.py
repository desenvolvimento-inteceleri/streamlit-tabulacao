import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import re

# Função para padronizar o nome da escola, removendo prefixos e espaços extras
def padronizar_nome_escola(nome):
    if not isinstance(nome, str):
        return ""
    nome = re.sub(r'\b[Ee]\s*\.?\s*[Mm]\s*\.?\s*[Ee]\s*\.?\s*[Ff]\s*\.?\s*\b', '', nome)
    nome = re.sub(r'\s+', ' ', nome)
    return nome.strip().upper()

# Função para limpar e padronizar o valor de pontuação, convertendo para número
def padronizar_pontuacao(pontuacao):
    if isinstance(pontuacao, str):
        pontuacao = re.sub(r'[^\d]', '', pontuacao)
    try:
        return int(pontuacao)
    except ValueError:
        return 0

# Função para gerar um valor de ordenação baseado na coluna "Ano"
def obter_ordem_ano(ano):
    # Regex para identificar anos numéricos (ex.: "2ª ano", "3º ano", etc.)
    match_ano = re.match(r'(\d+)[ªº]?\s*ano', str(ano).lower())
    if match_ano:
        return int(match_ano.group(1))  # Retorna o ano como número (ex.: 2 para "2ª ano")
    
    # Regex para identificar etapas EJAI (ex.: "EJAI 1ª etapa", "EJAI 2ª etapa")
    match_ejai = re.match(r'ejai\s*(\d+)[ªº]?\s*etapa', str(ano).lower())
    if match_ejai:
        return 100 + int(match_ejai.group(1))  # Adiciona 100 para ordenar as etapas EJAI após os anos
    
    return float('inf')  # Coloca valores desconhecidos no final

# Função para filtrar e ordenar o formulário de resposta
def gerar_classificatoria(formulario_df):
    colunas_selecionadas = [
        'Nome do aluno?',
        'Qual é o nome da sua escola?',
        'Escreva o nome da escola caso ela no esteja listada',
        'Ano escolar do aluno:',
        'Total de pontuação?',
        'Quanto tempo de realização?',
        'Se for aluno com deficiência/transtorno:'
    ]
    
    filtrado_df = formulario_df[colunas_selecionadas]
    filtrado_df.loc[filtrado_df['Qual é o nome da sua escola?'] == "Escola não está na lista", 'Qual é o nome da sua escola?'] = filtrado_df['Escreva o nome da escola caso ela no esteja listada']
    filtrado_df = filtrado_df.drop(columns=['Escreva o nome da escola caso ela no esteja listada'])
    
    filtrado_df.columns = [
        'Nome',
        'Escola',
        'Ano',
        'Pontuação',
        'Tempo',
        'Se for aluno com deficiência/transtorno:'
    ]
    
    filtrado_df['Nome'] = filtrado_df['Nome'].str.upper()
    filtrado_df['Escola'] = filtrado_df['Escola'].apply(padronizar_nome_escola)
    filtrado_df['Pontuação'] = filtrado_df['Pontuação'].apply(padronizar_pontuacao)
    
    # Adiciona uma coluna de ordenação de anos baseada na função obter_ordem_ano
    filtrado_df['Ordem_Ano'] = filtrado_df['Ano'].apply(obter_ordem_ano)
    
    # Ordenar o DataFrame
    filtrado_df = filtrado_df.sort_values(by=['Ordem_Ano', 'Pontuação', 'Tempo'], ascending=[True, False, True])
    
    # Remover a coluna auxiliar de ordenação
    filtrado_df = filtrado_df.drop(columns=['Ordem_Ano'])
    
    # Criar uma coluna extra chamada ETAPA
    filtrado_df['ETAPA'] = '2° CLASSIFICATÓRIA'
    filtrado_df = filtrado_df[['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Se for aluno com deficiência/transtorno:', 'ETAPA']]
    
    return filtrado_df

# Função para salvar o arquivo com escolas separadas em diferentes sheets
def salvar_excel_separado_por_escola(classificatoria_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        escolas = classificatoria_df['Escola'].unique()
        
        for escola in escolas:
            escola_sheet_name = escola[:31] if pd.notna(escola) and escola != "" else "ESCOLA_DESCONHECIDA"
            df_escola = classificatoria_df[classificatoria_df['Escola'] == escola]
            df_escola.to_excel(writer, sheet_name=escola_sheet_name, index=False, startrow=1)
            
            worksheet = writer.sheets[escola_sheet_name]
            worksheet.merge_range('A1:G1', escola, writer.book.add_format({'align': 'center', 'bold': True}))
    
    output.seek(0)
    return output

# Função principal do aplicativo Streamlit
def main():
    st.title("Gerador de Classificatória por Escola")
    uploaded_file = st.file_uploader("Envie o arquivo do Formulário de Resposta", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        formulario_df = pd.read_excel(uploaded_file)
        classificatoria_df = gerar_classificatoria(formulario_df)
        
        st.write("Dados filtrados e ordenados para o arquivo classificatória:")
        st.dataframe(classificatoria_df)
        
        if st.button("Gerar o arquivo Classificatória"):
            output = salvar_excel_separado_por_escola(classificatoria_df)
            st.download_button(
                label="Baixar Classificatória",
                data=output,
                file_name="classificatoria_por_escola.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == '__main__':
    main()
