#combinar abas sheets
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# Função para gerar a tabela dinâmica
def criar_tabela_dinamica(dados):
    # Gerar a tabela dinâmica com 'Escola' como índice, 'Ano' como colunas
    # E a contagem de alunos como valores
    tabela_dinamica = pd.pivot_table(
        dados,
        index='Escola',        # Escola será a linha
        columns='Ano',         # Ano será a coluna
        values='Nome',         # Vamos contar os alunos (coluna 'Nome')
        aggfunc='count',       # Contar o número de alunos por combinação
        fill_value=0           # Preencher valores ausentes com 0
    )
    
    return tabela_dinamica

def gerar_grafico(dados):
    # Agrupar os dados para contar a quantidade total de anos escolares
    total_anos_escolares = dados.groupby(['Ano']).size().reset_index(name="Quantidade de Anos")

    # Criar o gráfico de barras
    fig = px.bar(
        total_anos_escolares,
        x="Ano",
        y="Quantidade de Anos",
        title="Quantidade Total de Anos Escolares por Ano",
        labels={"Quantidade de Anos": "Total de Anos Escolares", "Ano": "Ano Escolar"}
    )
    fig.update_layout(barmode='group')
    return fig

def converter_para_xlsx(dataframe):
    xlsx_buffer = BytesIO()
    with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=True)
    xlsx_buffer.seek(0)
    return xlsx_buffer



