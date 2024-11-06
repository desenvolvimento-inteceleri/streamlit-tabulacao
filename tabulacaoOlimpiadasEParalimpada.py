#tabulacao do pedagogico
#pega o formulario de respota enviado pelos professores
#separa sheets por escola, e classifica dos maiores pontos no menor tempo
#retorna 2 arquivos: para a olimpiadas e para paralimpiada

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
    match_ano = re.match(r'(\d+)[ªº]?\s*ano', str(ano).lower())
    if match_ano:
        return int(match_ano.group(1))
    match_ejai = re.match(r'ejai\s*(\d+)[ªº]?\s*etapa', str(ano).lower())
    if match_ejai:
        return 100 + int(match_ejai.group(1))
    return float('inf')

# Função para filtrar e ordenar o formulário de resposta
def gerar_classificatoria(formulario_df, etapa):
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
        'Deficiência/Transtorno'
    ]
    
    filtrado_df['Nome'] = filtrado_df['Nome'].str.upper()
    filtrado_df['Escola'] = filtrado_df['Escola'].apply(padronizar_nome_escola)
    filtrado_df['Pontuação'] = filtrado_df['Pontuação'].apply(padronizar_pontuacao)
    filtrado_df['Ordem_Ano'] = filtrado_df['Ano'].apply(obter_ordem_ano)
    filtrado_df = filtrado_df.sort_values(by=['Ordem_Ano', 'Pontuação', 'Tempo'], ascending=[True, False, True])
    filtrado_df = filtrado_df.drop(columns=['Ordem_Ano'])
    
    filtrado_df['ETAPA'] = etapa
    filtrado_df = filtrado_df[['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Deficiência/Transtorno', 'ETAPA']]
    
    return filtrado_df

# Função para salvar arquivos Excel separados por categoria de deficiência
def salvar_excel_por_categoria(classificatoria_df):
    output_olimpiada = BytesIO()
    output_paralimpiada = BytesIO()
    
    # Dados para Alunos Olimpíada
    olimpiada_df = classificatoria_df[classificatoria_df['Deficiência/Transtorno'] == "Não possui deficiência/transtorno"]
    with pd.ExcelWriter(output_olimpiada, engine='xlsxwriter') as writer:
        escolas = olimpiada_df['Escola'].unique()
        for escola in escolas:
            escola_sheet_name = escola[:31] if pd.notna(escola) and escola != "" else "ESCOLA_DESCONHECIDA"
            df_escola = olimpiada_df[olimpiada_df['Escola'] == escola]
            df_escola.to_excel(writer, sheet_name=escola_sheet_name, index=False, startrow=1)
            worksheet = writer.sheets[escola_sheet_name]
            worksheet.merge_range('A1:G1', escola, writer.book.add_format({'align': 'center', 'bold': True}))
    
    # Dados para Alunos Paralimpíada
    paralimpiada_df = classificatoria_df[classificatoria_df['Deficiência/Transtorno'] != "Não possui deficiência/transtorno"]
    with pd.ExcelWriter(output_paralimpiada, engine='xlsxwriter') as writer:
        escolas = paralimpiada_df['Escola'].unique()
        for escola in escolas:
            escola_sheet_name = escola[:31] if pd.notna(escola) and escola != "" else "ESCOLA_DESCONHECIDA"
            df_escola = paralimpiada_df[paralimpiada_df['Escola'] == escola]
            df_escola.to_excel(writer, sheet_name=escola_sheet_name, index=False, startrow=1)
            worksheet = writer.sheets[escola_sheet_name]
            worksheet.merge_range('A1:G1', escola, writer.book.add_format({'align': 'center', 'bold': True}))
    
    output_olimpiada.seek(0)
    output_paralimpiada.seek(0)
    return output_olimpiada, output_paralimpiada

# Função principal do aplicativo Streamlit
def main():
    st.title("Tabulação: Gerador de Classificatória por Escola")


     # Exemplo de estrutura de dados
    st.write("Exemplo da estrutura de dados esperada para upload (Formulario de resposta):")
    exemplo_data = {
        "Nome do aluno?": ["ALUNO 1", "ALUNO 2", "ALUNO 3"],
        "Qual é o nome da sua escola?": ["ESCOLA 1", "ESCOLA 2", "ESCOLA 3"],
        "Escreva o nome da escola caso ela no esteja listada": ["", "", ""],
        "Ano escolar do aluno:": ["1ª ANO", "2ª ANO", "3ª ANO"],
        "Total de pontuação?": [45, 38, 50],
        "Quanto tempo de realização?": ["00:15:00", "00:12:00", "00:10:00"],
        "Se for aluno com deficiência/transtorno:": ["NÃO POSSUI DEFICIÊNCIA/TRANSTORNO", "DEFICIÊNCIA FÍSICA", "NÃO POSSUI DEFICIÊNCIA/TRANSTORNO"]
    }
    exemplo_df = pd.DataFrame(exemplo_data)
    st.table(exemplo_df)


    uploaded_file = st.file_uploader("Envie o arquivo do Formulário de Resposta", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        formulario_df = pd.read_excel(uploaded_file)
        
        # Seleção da Etapa
        etapa_selecionada = st.selectbox("Selecione a Etapa", ["1° CLASSIFICATÓRIA", "2° CLASSIFICATÓRIA", "OUTROS"])
        if etapa_selecionada == "OUTROS":
            etapa_customizada = st.text_input("Digite o nome da Etapa")
            etapa = etapa_customizada if etapa_customizada else "OUTROS"
        else:
            etapa = etapa_selecionada
        
        classificatoria_df = gerar_classificatoria(formulario_df, etapa)
        st.write("Dados filtrados e ordenados para o arquivo classificatória:")
        st.dataframe(classificatoria_df)
        
        if st.button("Gerar Arquivos"):
            output_olimpiada, output_paralimpiada = salvar_excel_por_categoria(classificatoria_df)
            col1, col2 = st.columns(2)  # Cria duas colunas
    
            with col1:
                st.download_button(
                    label="Baixar Alunos Olimpíada",
                    data=output_olimpiada,
                    file_name="classificatoria_olimpiada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                st.download_button(
                    label="Baixar Alunos Paralimpíada",
                    data=output_paralimpiada,
                    file_name="classificatoria_paralimpiada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        

if __name__ == '__main__':
    main()
