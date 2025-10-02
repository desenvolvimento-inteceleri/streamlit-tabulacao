# Tabulação Olimpíada e Paralimpíada
# Gera 3 arquivos (Olimpíada, Paralimpíada e JUNÇÃO)
# Cada arquivo possui a aba "GERAL" + abas por escola, com:
# - Largura automática de colunas
# - Cabeçalho estilizado (fundo verde, fonte branca)
# - AutoFilter e Freeze Panes
import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import re

# --- Utilitários de padronização/ordenação ---
def padronizar_nome_escola(nome):
    if not isinstance(nome, str):
        return ""
    nome = re.sub(r'\b[Ee]\s*\.?\s*[Mm]\s*\.?\s*[Ee]\s*\.?\s*[Ff]\s*\.?\s*\b', '', nome)
    nome = re.sub(r'\s+', ' ', nome)
    return nome.strip().upper()

def padronizar_pontuacao(pontuacao):
    if isinstance(pontuacao, str):
        pontuacao = re.sub(r'[^\d]', '', pontuacao)
    try:
        return int(pontuacao)
    except ValueError:
        return 0

def obter_ordem_ano(ano):
    s = str(ano).lower()
    m = re.match(r'(\d+)[ªº]?\s*ano', s)
    if m:
        return int(m.group(1))
    m2 = re.match(r'ejai\s*(\d+)[ªº]?\s*etapa', s)
    if m2:
        return 100 + int(m2.group(1))
    return float('inf')

def ajustar_nome_aba(nome, usados):
    nome = nome[:31] if nome else "ESCOLA_DESCONHECIDA"
    nome = re.sub(r'[\\/*?:\[\]]', '_', nome).strip()
    original = nome
    c = 1
    while nome in usados:
        nome = f"{original[:28]}_{c}"
        c += 1
    usados.add(nome)
    return nome

# --- Formatação auxiliar (largura + header style) ---
def aplicar_formatacao_basica(writer, sheet_name, df, header_row_idx=0):
    """
    - Define largura automática de colunas
    - Aplica estilo ao cabeçalho na linha 'header_row_idx'
    - Mantém autofiltro e freeze panes (já configurados onde chamamos)
    """
    ws = writer.sheets[sheet_name]
    book = writer.book

    # Estilo do cabeçalho
    header_fmt = book.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#6AA84F',   # verde (pode ajustar)
        'font_color': 'white',
        'border': 1
    })

    # Aplica formatação de cabeçalho (linha do cabeçalho)
    for col_idx, col_name in enumerate(df.columns):
        ws.write(header_row_idx, col_idx, col_name, header_fmt)

    # Largura automática: máximo entre nome da coluna e dados (como string)
    for col_idx, col_name in enumerate(df.columns):
        series_as_str = df[col_name].astype(str)
        max_len_data = series_as_str.map(len).max() if len(series_as_str) > 0 else 0
        max_len_header = len(col_name)
        width = min(max(max_len_data, max_len_header) + 2, 60)  # limite máx 60
        ws.set_column(col_idx, col_idx, width)

# --- Pipeline principal de filtragem/ordenação ---
def gerar_classificatoria(formulario_df, etapa):
    colunas_mapeamento = {
        'Nome do aluno?': 'Nome do aluno?',
        'Selecione o nome da sua escola': 'Qual é o nome da sua escola?',
        'Caso a escola não esteja na lista acima, escreva o nome aqui:': 'Escreva o nome da escola caso ela no esteja listada',
        'Ano escolar do aluno:': 'Ano escolar do aluno:',
        'Total de pontuação ?': 'Total de pontuação?',
        'Quanto tempo de realização?': 'Quanto tempo de realização?',
        'Se for aluno com deficiência/transtorno:': 'Se for aluno com deficiência/transtorno:'
    }
    formulario_df.rename(columns=colunas_mapeamento, inplace=True)

    cols = [
        'Nome do aluno?',
        'Qual é o nome da sua escola?',
        'Escreva o nome da escola caso ela no esteja listada',
        'Ano escolar do aluno:',
        'Total de pontuação?',
        'Quanto tempo de realização?',
        'Se for aluno com deficiência/transtorno:'
    ]
    filtrado_df = formulario_df[cols].copy()
    mask_out = filtrado_df['Qual é o nome da sua escola?'] == "Escola não está na lista"
    filtrado_df.loc[mask_out, 'Qual é o nome da sua escola?'] = filtrado_df['Escreva o nome da escola caso ela no esteja listada']
    filtrado_df.drop(columns=['Escreva o nome da escola caso ela no esteja listada'], inplace=True)

    filtrado_df.columns = ['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Deficiência/Transtorno']  # <- reorganizamos ao renomear?
    # Atenção: a ordem acima precisa refletir corretamente. Vamos corrigir:
    filtrado_df.columns = ['Nome', 'Escola', 'Ano', 'Pontuação', 'Tempo', 'Deficiência/Transtorno']

    filtrado_df['Nome'] = filtrado_df['Nome'].str.upper()
    filtrado_df['Escola'] = filtrado_df['Escola'].apply(padronizar_nome_escola)
    filtrado_df['Pontuação'] = filtrado_df['Pontuação'].apply(padronizar_pontuacao)
    filtrado_df['Ordem_Ano'] = filtrado_df['Ano'].apply(obter_ordem_ano)

    filtrado_df.sort_values(by=['Ordem_Ano', 'Pontuação', 'Tempo'], ascending=[True, False, True], inplace=True)
    filtrado_df.drop(columns=['Ordem_Ano'], inplace=True)
    filtrado_df['ETAPA'] = etapa
    filtrado_df = filtrado_df[['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Deficiência/Transtorno', 'ETAPA']]
    return filtrado_df

# --- Escrita em Excel ---
def escrever_geral(writer, df):
    """Cria a aba 'GERAL' com o dataframe completo, com filtro e freeze panes, mais formatação."""
    sheet = 'GERAL'
    df.to_excel(writer, sheet_name=sheet, index=False)
    ws = writer.sheets[sheet]
    # Auto-filter e freeze da primeira linha
    ws.autofilter(0, 0, len(df), len(df.columns)-1)
    ws.freeze_panes(1, 0)
    # Formatação (cabeçalho está na linha 0)
    aplicar_formatacao_basica(writer, sheet, df, header_row_idx=0)

def escrever_por_escola(writer, df):
    """Cria abas por escola com cabeçalho mesclado, filtro, freeze panes e formatação."""
    escolas = df['Escola'].dropna().unique()
    usados = set(['GERAL'])  # reserva o nome GERAL
    for escola in escolas:
        nome_sheet = ajustar_nome_aba(escola, usados)
        df_esc = df[df['Escola'] == escola]
        # escrevemos a partir da linha 1 (linha 0 reserva o título mesclado)
        df_esc.to_excel(writer, sheet_name=nome_sheet, index=False, startrow=1)
        ws = writer.sheets[nome_sheet]

        # Título mesclado
        ws.merge_range('A1:G1', escola, writer.book.add_format({
            'align': 'center', 'bold': True, 'bg_color': '#D9EAD3', 'border': 1
        }))

        # Filtro e freeze panes (linha de cabeçalho está na linha 1)
        ws.autofilter(1, 0, len(df_esc)+1, df_esc.shape[1]-1)
        ws.freeze_panes(2, 0)

        # Formatação (cabeçalho na linha 1)
        aplicar_formatacao_basica(writer, nome_sheet, df_esc, header_row_idx=1)

def salvar_excels(classificatoria_df):
    output_olimpiada = BytesIO()
    output_paralimpiada = BytesIO()
    output_juncao = BytesIO()

    # 1) Olimpíada
    olimpiada_df = classificatoria_df[classificatoria_df['Deficiência/Transtorno'] == "Não possui deficiência/transtorno"]
    with pd.ExcelWriter(output_olimpiada, engine='xlsxwriter') as writer:
        escrever_geral(writer, olimpiada_df)
        escrever_por_escola(writer, olimpiada_df)

    # 2) Paralimpíada
    paralimpiada_df = classificatoria_df[classificatoria_df['Deficiência/Transtorno'] != "Não possui deficiência/transtorno"]
    with pd.ExcelWriter(output_paralimpiada, engine='xlsxwriter') as writer:
        escrever_geral(writer, paralimpiada_df)
        escrever_por_escola(writer, paralimpiada_df)

    # 3) JUNÇÃO (todos)
    with pd.ExcelWriter(output_juncao, engine='xlsxwriter') as writer:
        escrever_geral(writer, classificatoria_df)
        escrever_por_escola(writer, classificatoria_df)

    output_olimpiada.seek(0)
    output_paralimpiada.seek(0)
    output_juncao.seek(0)
    return output_olimpiada, output_paralimpiada, output_juncao

# --- Streamlit ---
def main():
    st.title("Tabulação: Gerador de Classificatória por Escola")
    st.write("Exemplo da estrutura de dados esperada para upload:")
    exemplo_data = {
        "Nome do aluno?": ["ALUNO 1", "ALUNO 2", "ALUNO 3"],
        "Qual é o nome da sua escola?": ["ESCOLA 1", "ESCOLA 2", "ESCOLA 3"],
        "Escreva o nome da escola caso ela no esteja listada": ["", "", ""],
        "Ano escolar do aluno:": ["1ª ANO", "2ª ANO", "3ª ANO"],
        "Total de pontuação?": [45, 38, 50],
        "Quanto tempo de realização?": ["00:15:00", "00:12:00", "00:10:00"],
        "Se for aluno com deficiência/transtorno:": ["NÃO POSSUI DEFICIÊNCIA/TRANSTORNO", "DEFICIÊNCIA FÍSICA", "NÃO POSSUI DEFICIÊNCIA/TRANSTORNO"]
    }
    st.table(pd.DataFrame(exemplo_data))

    uploaded_file = st.file_uploader("Envie o arquivo do Formulário de Resposta", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            formulario_df = pd.read_excel(uploaded_file)
            st.success("Arquivo carregado com sucesso!")
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
            return

        etapa_sel = st.selectbox("Selecione a Etapa", ["1° CLASSIFICATÓRIA", "2° CLASSIFICATÓRIA", "OUTROS"])
        etapa = st.text_input("Digite o nome da Etapa") if etapa_sel == "OUTROS" else etapa_sel

        classificatoria_df = gerar_classificatoria(formulario_df, etapa)
        st.write("Dados filtrados e ordenados:")
        st.dataframe(classificatoria_df)

        if st.button("Gerar Arquivos"):
            out_olimp, out_para, out_junc = salvar_excels(classificatoria_df)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    "Baixar Alunos Olimpíada",
                    out_olimp,
                    "classificatoria_olimpiada.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    "Baixar Alunos Paralimpíada",
                    out_para,
                    "classificatoria_paralimpiada.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col3:
                st.download_button(
                    "Baixar JUNÇÃO (Olimpíada + Paralimpíada)",
                    out_junc,
                    "classificatoria_juncao.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == '__main__':
    main()
