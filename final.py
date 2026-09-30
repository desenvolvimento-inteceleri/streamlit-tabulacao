# upload da listagem de respostas do formulário (etapa final),
# seleciona os 2 melhores (ou qtd que o usuario desejar) de cada ano escolar de cada escola
# criterios: maior pontuação e, em caso de empate, menor tempo
# gera um arquivo no padrão da semifinal (aba GERAL + uma aba por escola), com a coluna PROFESSOR

import re
import streamlit as st
import pandas as pd
from io import BytesIO
from tabulacaoOlimpiadasEParalimpada import inserir_banner

COLUNAS_SAIDA = ['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Deficiência/Transtorno', 'RESPONSÁVEL', 'PROFESSOR']
LARGURAS = [9, 38, 42, 11, 11, 32, 34, 34]


# Encontra a coluna do formulário que contém todos os trechos informados (ignorando maiúsculas)
def encontrar_coluna(df, *trechos, evitar=None):
    for coluna in df.columns:
        nome = str(coluna).lower()
        if all(t in nome for t in trechos) and not (evitar and evitar in nome):
            return coluna
    return None


# Converte o tempo digitado (ex.: "00:11: 52", "00:14;00", "02.41.60", "00:45") em segundos
# Retorna (segundos, formato_irregular)
def converter_tempo(valor):
    if pd.isna(valor):
        return None, True
    texto = re.sub(r"[;.'’]", ':', str(valor)).replace(' ', '')
    partes = texto.split(':')
    if not all(p.isdigit() for p in partes):
        return None, True
    numeros = [int(p) for p in partes]
    if len(numeros) == 3:
        h, m, s = numeros
        irregular = not re.fullmatch(r'\d{2}:\d{2}:\d{2}', str(valor).strip()) or m >= 60 or s >= 60
    elif len(numeros) == 2:
        # Sem as horas, considera MM:SS
        h, (m, s) = 0, numeros
        irregular = True
    else:
        return None, True
    return h * 3600 + m * 60 + s, irregular


def formatar_tempo(segundos):
    if segundos is None or pd.isna(segundos):
        return ''
    segundos = int(segundos)
    return f'{segundos // 3600:02d}:{segundos % 3600 // 60:02d}:{segundos % 60:02d}'


# Numero do ano escolar para ordenar (1º ano, 2º ano, ...)
def numero_do_ano(ano):
    encontrado = re.search(r'\d+', str(ano))
    return int(encontrado.group()) if encontrado else 99


# Monta a tabela no padrão da semifinal a partir das respostas do formulário
def organizar_respostas(df):
    col_nome = encontrar_coluna(df, 'nome do aluno')
    col_escola = encontrar_coluna(df, 'escola', evitar='caso')
    col_escola_alt = encontrar_coluna(df, 'escola', 'caso')
    col_ano = encontrar_coluna(df, 'ano escolar')
    col_pontos = encontrar_coluna(df, 'pontos') or encontrar_coluna(df, 'pontuação')
    col_tempo = encontrar_coluna(df, 'tempo')
    col_deficiencia = encontrar_coluna(df, 'deficiência')
    col_professor = encontrar_coluna(df, 'professor')
    col_responsavel = encontrar_coluna(df, 'nome do responsável')
    col_carimbo = encontrar_coluna(df, 'carimbo')

    obrigatorias = {'Nome do aluno': col_nome, 'Escola': col_escola, 'Ano escolar': col_ano,
                    'Pontos': col_pontos, 'Tempo': col_tempo}
    faltando = [nome for nome, coluna in obrigatorias.items() if coluna is None]
    if faltando:
        raise ValueError(f"Colunas não encontradas na planilha: {', '.join(faltando)}")

    def texto(coluna):
        if coluna is None:
            return pd.Series('', index=df.index)
        return df[coluna].fillna('').astype(str).str.strip()

    tabela = pd.DataFrame({
        'Ano': texto(col_ano),
        'Nome': texto(col_nome),
        'Escola': texto(col_escola),
        'Pontuação': pd.to_numeric(df[col_pontos], errors='coerce'),
        'Tempo original': df[col_tempo],
        'Deficiência/Transtorno': texto(col_deficiencia),
        'RESPONSÁVEL': texto(col_responsavel),
        'PROFESSOR': texto(col_professor),
        'Carimbo': df[col_carimbo] if col_carimbo else pd.NaT,
    })

    # Escola fora da lista: usa o nome escrito no campo alternativo
    escola_alt = texto(col_escola_alt)
    fora_da_lista = (tabela['Escola'] == '') | tabela['Escola'].str.lower().str.contains('não está na lista')
    tabela.loc[fora_da_lista, 'Escola'] = escola_alt[fora_da_lista]

    # Remove linhas vazias
    tabela = tabela[tabela['Nome'] != '']

    tempos = tabela['Tempo original'].apply(converter_tempo)
    tabela['Segundos'] = tempos.str[0]
    tabela['Tempo irregular'] = tempos.str[1]
    tabela['Tempo'] = tabela['Segundos'].apply(formatar_tempo)
    tabela['Ordem ano'] = tabela['Ano'].apply(numero_do_ano)

    # Aluno enviado mais de uma vez: mantém o envio mais recente
    tabela['Nome normalizado'] = tabela['Nome'].str.upper().str.split().str.join(' ')
    chave = ['Nome normalizado', 'Escola', 'Ano']
    repetidos = tabela[tabela.duplicated(subset=chave, keep=False)]
    tabela = tabela.sort_values('Carimbo').drop_duplicates(subset=chave, keep='last')

    return tabela, repetidos


# Seleciona os N melhores de cada ano de cada escola
def selecionar_melhores(tabela, quantidade):
    ordenada = tabela.sort_values(
        by=['Escola', 'Ordem ano', 'Pontuação', 'Segundos'],
        ascending=[True, True, False, True],
        na_position='last',
    )
    grupos = ordenada.groupby(['Escola', 'Ordem ano'], sort=False)
    selecionados = grupos.head(quantidade)

    # Empate exato (pontuação e tempo) entre o último classificado e o primeiro de fora
    empates = []
    for (escola, _), grupo in grupos:
        if len(grupo) > quantidade:
            ultimo, proximo = grupo.iloc[quantidade - 1], grupo.iloc[quantidade]
            if ultimo['Pontuação'] == proximo['Pontuação'] and ultimo['Segundos'] == proximo['Segundos']:
                empates.append(grupo.iloc[quantidade - 1:quantidade + 1])
    empates = pd.concat(empates) if empates else pd.DataFrame()

    return selecionados, empates


# Nome de aba valido no Excel (max 31 caracteres, sem caracteres proibidos e sem repetir)
def nome_da_aba(nome, usados):
    base = re.sub(r'[\[\]:*?/\\]', '', nome)[:31] or 'SEM ESCOLA'
    aba, contador = base, 2
    while aba.upper() in usados:
        sufixo = f' ({contador})'
        aba = base[:31 - len(sufixo)] + sufixo
        contador += 1
    usados.add(aba.upper())
    return aba


# Gera o Excel no padrão da semifinal: aba GERAL + uma aba por escola
# Com imagem, a logo ocupa as primeiras banner_rows linhas de todas as abas
def gerar_excel(selecionados, image_bytes=None, banner_rows=2, banner_h_px=110):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        book = writer.book
        fmt_cabecalho = book.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#2254C5',
                                         'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_dados = book.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        ultima_coluna = chr(ord('A') + len(COLUNAS_SAIDA) - 1)
        larguras_px = [int(largura * 7 + 5) for largura in LARGURAS]

        # Área do topo: logo (se houver) ou linhas em branco mescladas, como na semifinal
        def escrever_topo(worksheet, linhas_sem_logo):
            if image_bytes:
                inserir_banner(worksheet, image_bytes, larguras_px, len(COLUNAS_SAIDA),
                               banner_rows=banner_rows, target_height_px=banner_h_px)
                return banner_rows
            worksheet.merge_range(f'A1:{ultima_coluna}{linhas_sem_logo}', '', fmt_dados)
            return linhas_sem_logo

        def escrever_tabela(worksheet, df, linha_cabecalho):
            for col, titulo in enumerate(COLUNAS_SAIDA):
                worksheet.write(linha_cabecalho, col, titulo, fmt_cabecalho)
                worksheet.set_column(col, col, LARGURAS[col])
            for i, linha in enumerate(df[COLUNAS_SAIDA].itertuples(index=False), start=linha_cabecalho + 1):
                for col, valor in enumerate(linha):
                    worksheet.write(i, col, '' if pd.isna(valor) else valor, fmt_dados)
            worksheet.freeze_panes(linha_cabecalho + 1, 0)

        # Aba GERAL: todos os classificados, por ano e escola
        geral = selecionados.sort_values(by=['Ordem ano', 'Escola', 'Pontuação', 'Segundos'],
                                         ascending=[True, True, False, True], na_position='last')
        ws_geral = book.add_worksheet('GERAL')
        linhas_topo = escrever_topo(ws_geral, 1)
        escrever_tabela(ws_geral, geral, linhas_topo)

        # Uma aba por escola: topo, nome da escola e tabela
        usados = {'GERAL'}
        for escola, df_escola in selecionados.groupby('Escola', sort=True):
            ws = book.add_worksheet(nome_da_aba(escola, usados))
            linhas_topo = escrever_topo(ws, 2)
            ws.merge_range(linhas_topo, 0, linhas_topo, len(COLUNAS_SAIDA) - 1, escola, fmt_cabecalho)
            escrever_tabela(ws, df_escola, linhas_topo + 1)

    output.seek(0)
    return output, geral


# Função principal do aplicativo
def main():
    st.title('Final')
    st.write('Carregue a listagem de respostas do formulário (etapa final). '
             'Serão selecionados os melhores alunos de cada ano escolar de cada escola, '
             'por maior pontuação e, em caso de empate, menor tempo.')

    arquivo = st.file_uploader('Upload da listagem de respostas', type=['xlsx'])
    quantidade = st.number_input('Quantidade de alunos por ano escolar em cada escola', min_value=1, max_value=10, value=2)

    if not arquivo:
        return

    abas = pd.ExcelFile(arquivo).sheet_names
    aba = st.selectbox('Aba com as respostas', abas, index=0)
    df = pd.read_excel(arquivo, sheet_name=aba)

    try:
        tabela, repetidos = organizar_respostas(df)
    except ValueError as erro:
        st.error(str(erro))
        return

    usar_banner = st.checkbox('Adicionar imagem no topo (todas as abas)', value=True)
    banner_altura = st.slider('Altura do banner (px)', min_value=60, max_value=220, value=110, step=10)
    banner_linhas = st.slider('Linhas reservadas para o banner', min_value=2, max_value=5, value=2, step=1)

    image_bytes = None
    if usar_banner:
        img_file = st.file_uploader('Envie a imagem (PNG/JPG), opcional', type=['png', 'jpg', 'jpeg'])
        if img_file is not None:
            image_bytes = img_file.read()

    selecionados, empates = selecionar_melhores(tabela, quantidade)
    output, geral = gerar_excel(selecionados, image_bytes=image_bytes,
                                banner_rows=banner_linhas, banner_h_px=banner_altura)

    st.success(f'{len(selecionados)} alunos classificados de {tabela["Escola"].nunique()} escolas '
               f'({len(tabela)} respostas válidas).')

    if not repetidos.empty:
        st.warning('Alunos enviados mais de uma vez (mesmo nome, escola e ano). Foi mantido apenas o envio mais recente:')
        st.dataframe(repetidos[['Carimbo', 'Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo original']], hide_index=True)

    irregulares = tabela[tabela['Tempo irregular']]
    if not irregulares.empty:
        st.warning('Tempos digitados fora do padrão HH:MM:SS. Confira se a conversão está correta '
                   '(tempos com só duas partes, como "00:45", foram lidos como MM:SS):')
        st.dataframe(irregulares[['Ano', 'Nome', 'Escola', 'Tempo original', 'Tempo']], hide_index=True)

    if not empates.empty:
        st.warning('Empate exato (mesma pontuação e mesmo tempo) no limite da classificação. Confira manualmente:')
        st.dataframe(empates[['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo']], hide_index=True)

    st.subheader('Classificados')
    st.dataframe(geral[COLUNAS_SAIDA], hide_index=True)

    st.download_button(
        label='Baixar Classificação Final',
        data=output,
        file_name='classificacao_final.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    main()
