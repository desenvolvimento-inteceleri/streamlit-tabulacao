#classificação pontuação/tabulação
# - Lê a planilha gerada pela tabulação (logo, título e cabeçalho em qualquer linha)
# - Mantém os N primeiros alunos de cada ano em cada escola, na ordem em que já estão na planilha
# - Edita o próprio arquivo enviado (openpyxl): abas, logo, mesclagens, cores, larguras e filtros são preservados

import re
import unicodedata
from copy import copy
from io import BytesIO

import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter, range_boundaries

LINHAS_PARA_PROCURAR_CABECALHO = 30


def normalizar(valor):
    """Texto em minúsculas, sem acentos, sem º/°/ª e sem espaços extras."""
    if valor is None:
        return ''
    texto = re.sub(r'[º°ª]', '', str(valor))  # antes do NFKD, que transformaria º em "o"
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(c for c in texto if not unicodedata.combining(c)).lower()
    return re.sub(r'\s+', ' ', texto).strip(' :?')


def encontrar_cabecalho(ws):
    """Procura nas primeiras linhas a que tem a coluna Ano (junto com Nome ou Pontuação).
    Retorna (linha, {campo: coluna}) ou (None, {})."""
    for linha in range(1, min(ws.max_row, LINHAS_PARA_PROCURAR_CABECALHO) + 1):
        colunas = {}
        for col in range(1, ws.max_column + 1):
            t = normalizar(ws.cell(linha, col).value)
            if t in ('ano', 'ano escolar') or t.startswith('ano escolar'):
                colunas.setdefault('ano', col)
            elif t == 'escola':
                colunas.setdefault('escola', col)
            elif t == 'nome' or 'nome do aluno' in t:
                colunas.setdefault('nome', col)
            elif 'pontua' in t or 'pontos' in t:
                colunas.setdefault('pontuacao', col)
        if 'ano' in colunas and ('nome' in colunas or 'pontuacao' in colunas):
            return linha, colunas
    return None, {}


def selecionar_linhas(ws, linha_cabecalho, colunas, top_n):
    """Percorre os alunos na ordem da planilha e guarda os top_n primeiros de cada escola + ano.
    Retorna (linhas_mantidas, anos_dos_mantidos)."""
    contagem = {}
    manter, anos = [], []
    for linha in range(linha_cabecalho + 1, ws.max_row + 1):
        valores = [ws.cell(linha, c).value for c in range(1, ws.max_column + 1)]
        if all(v is None or str(v).strip() == '' for v in valores):
            continue  # linha vazia

        ano = ws.cell(linha, colunas['ano']).value
        if ano is None or str(ano).strip() == '':
            manter.append(linha)  # linha que não é de aluno: mantém como está
            continue

        escola = normalizar(ws.cell(linha, colunas['escola']).value) if 'escola' in colunas else ''
        chave = (escola, normalizar(ano))
        contagem[chave] = contagem.get(chave, 0) + 1
        if contagem[chave] <= top_n:
            manter.append(linha)
            anos.append(str(ano).strip())
    return manter, anos


def compactar_aba(ws, linha_cabecalho, manter):
    """Sobe as linhas mantidas para logo abaixo do cabeçalho (com valores, estilos e altura)
    e apaga o restante. O que fica acima do cabeçalho não é tocado."""
    ultima_original = ws.max_row
    max_col = ws.max_column
    destino = linha_cabecalho + 1

    for origem in manter:
        if origem != destino:
            for col in range(1, max_col + 1):
                o = ws.cell(origem, col)
                d = ws.cell(destino, col)
                if isinstance(o, MergedCell) or isinstance(d, MergedCell):
                    continue
                d.value = o.value
                d._style = copy(o._style)
                d.number_format = o.number_format
            ws.row_dimensions[destino].height = ws.row_dimensions[origem].height
        destino += 1

    nova_ultima = destino - 1
    if nova_ultima < ultima_original:
        # Desfaz mesclagens que estavam na área apagada
        for faixa in list(ws.merged_cells.ranges):
            if faixa.min_row > nova_ultima:
                ws.merged_cells.remove(faixa)
        ws.delete_rows(nova_ultima + 1, ultima_original - nova_ultima)

    # Filtro (setas do cabeçalho) e tabelas passam a cobrir só as linhas que ficaram
    ultima_linha = max(nova_ultima, linha_cabecalho)
    if ws.auto_filter.ref:
        min_col, min_row, max_col_f, _ = range_boundaries(ws.auto_filter.ref)
        ws.auto_filter.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col_f)}{ultima_linha}"
    for tabela in ws.tables.values():
        min_col, min_row, max_col_t, _ = range_boundaries(tabela.ref)
        tabela.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col_t)}{max(ultima_linha, min_row + 1)}"
        if tabela.autoFilter is not None:
            tabela.autoFilter.ref = tabela.ref


def ordem_do_ano(ano):
    """1º ano, 2º ano, ... e EJAI depois dos anos regulares."""
    t = normalizar(ano)
    numero = re.search(r'\d+', t)
    n = int(numero.group()) if numero else 99
    return n + 100 if 'ejai' in t else n


def processar_planilha(arquivo_bytes, top_n):
    wb = load_workbook(BytesIO(arquivo_bytes))
    anos_por_aba, erros = {}, []

    for ws in wb.worksheets:
        linha_cabecalho, colunas = encontrar_cabecalho(ws)
        if linha_cabecalho is None:
            erros.append(ws.title)
            continue
        manter, anos = selecionar_linhas(ws, linha_cabecalho, colunas, top_n)
        compactar_aba(ws, linha_cabecalho, manter)
        anos_por_aba[ws.title] = anos

    saida = BytesIO()
    wb.save(saida)
    saida.seek(0)
    return saida, anos_por_aba, erros


# Função principal do aplicativo
def main():
    st.title("Classificação dos Melhores Alunos por escola")

    # Passo 1: Upload do arquivo
    uploaded_file = st.file_uploader("Carregue o arquivo Excel com os dados dos alunos", type=["xlsx"])

    if uploaded_file is not None:
        # Seleção do número de melhores alunos por ano para exibir
        top_n = st.selectbox("Escolha o número de melhores alunos por ano para exibir", [1, 2, 3, 4, 5])

        try:
            output, anos_por_aba, erros = processar_planilha(uploaded_file.getvalue(), top_n)
        except Exception as e:
            st.error(f"Não foi possível ler o arquivo Excel: {e}")
            return

        for aba in erros:
            st.error(f"Erro na aba {aba}: Coluna 'Ano' não encontrada. A aba foi mantida sem alterações.")

        # Botão para download do arquivo Excel gerado
        st.download_button(
            label="Baixar Classificação dos Melhores Alunos",
            data=output,
            file_name="melhores_alunos_classificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Contagem por ano: a aba GERAL repete os alunos das escolas, então só entra se for a única
        abas_escolas = [a for a in anos_por_aba if normalizar(a) != 'geral']
        abas_contadas = abas_escolas or list(anos_por_aba)
        anos = [ano for aba in abas_contadas for ano in anos_por_aba[aba]]

        if not anos:
            st.info("Nenhum aluno encontrado para montar o gráfico.")
            return

        # "4º ano" e "4° ano" contam como o mesmo ano; o gráfico mostra a primeira grafia encontrada
        rotulos = {}
        for ano in anos:
            rotulos.setdefault(normalizar(ano), ano)
        contagem = pd.Series([normalizar(a) for a in anos]).value_counts()
        ordem = sorted(contagem.index, key=ordem_do_ano)
        table_data = pd.DataFrame({'Ano': [rotulos[a] for a in ordem],
                                   'Quantidade': [int(contagem[a]) for a in ordem]})

        # Exibir gráfico de quantidade de alunos por ano interativo
        st.write("Quantidade total de alunos por Ano:")
        fig = px.bar(table_data, x='Ano', y='Quantidade', title="Quantidade por Ano")
        st.plotly_chart(fig)

        # Exibir a tabela de contagem de alunos por ano
        st.write("Tabela de Quantidade de Alunos por Ano:")
        st.dataframe(table_data, hide_index=True)

if __name__ == '__main__':
    main()
