# Tabulação Olimpíada e Paralimpíada (versão unificada, banner dentro da célula)
# - Gera 3 arquivos (Olimpíada, Paralimpíada e JUNÇÃO)
# - Cada arquivo: aba GERAL + abas por escola
# - Cabeçalho estilizado + largura automática + AutoFilter + Freeze panes
# - Banner opcional no topo (área mesclada A1:..), sem sobreposição

import streamlit as st
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
from io import BytesIO
import re
import numpy as np
import unicodedata
from PIL import Image  # para dimensionar a imagem do banner com precisão

# ------------------ Utilitários ------------------
def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = _strip_accents(s).lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[:?;.!]+$", "", s)
    return s

def padronizar_nome_escola(nome):
    if not isinstance(nome, str):
        return ""
    # remove E.M.E.F. / EMEF etc.
    nome = re.sub(r'\b[Ee]\s*\.?\s*[Mm]\s*\.?\s*[Ee]\s*\.?\s*[Ff]\s*\.?\s*\b', '', nome)
    nome = re.sub(r'\s+', ' ', nome)
    return nome.strip().upper()

def padronizar_pontuacao(p):
    if isinstance(p, str):
        p = re.sub(r'[^\d]', '', p)
    try:
        return int(p)
    except Exception:
        try:
            return int(float(p))
        except Exception:
            return 0

def _parse_tempo(x):
    """Converte 'hh:mm:ss', 'mm:ss' ou 'ss' (str ou num) para segundos (float)."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if re.fullmatch(r"\d+(\.\d+)?", s):
        try:
            return float(s)
        except:
            return np.nan
    parts = s.split(":")
    try:
        parts = [float(p) for p in parts]
    except:
        return np.nan
    if len(parts) == 3:
        h, m, s2 = parts
        return h*3600 + m*60 + s2
    if len(parts) == 2:
        m, s2 = parts
        return m*60 + s2
    if len(parts) == 1:
        return parts[0]
    return np.nan

def obter_ordem_ano(ano):
    """Extrai número do ano; aceita 1º/1ª/1° ANO. EJAI* vai para 100+etapa."""
    s = str(ano).lower()
    m = re.match(r'(\d+)[ªº°]?\s*ano', s)
    if m:
        return int(m.group(1))
    m2 = re.match(r'ejai\s*(\d+)[ªº°]?\s*etapa', s)
    if m2:
        return 100 + int(m2.group(1))
    m3 = re.search(r'(\d+)', s)
    if m3:
        return int(m3.group(1))
    return float('inf')

def ajustar_nome_aba(nome, usados):
    nome = nome[:31] if nome else "ESCOLA_DESCONHECIDA"
    nome = re.sub(r'[\\/*?:\[\]]', '_', nome).strip()
    base = nome
    c = 1
    while nome in usados:
        nome = f"{base[:28]}_{c}"
        c += 1
    usados.add(nome)
    return nome

# ------------------ Mapeamento de Colunas (flexível) ------------------
# Canoniza para estes nomes:
# 'Data/Hora','Email','Nome','EscolaSel','EscolaLivre','Ano','Pontuacao','Tempo','DefTran','Mensagem'
_COL_CANON = {
    'carimbo de data/hora': 'Data/Hora',
    'endereco de e-mail': 'Email',
    'endereço de e-mail': 'Email',
    'nome do aluno': 'Nome',
    'nome do aluno?': 'Nome',
    'selecione o nome da sua escola': 'EscolaSel',
    'selecione o nome da sua escola?': 'EscolaSel',
    'qual e o nome da sua escola': 'EscolaSel',
    'qual é o nome da sua escola': 'EscolaSel',
    'escreva o nome da escola caso ela no esteja listada': 'EscolaLivre',
    'ano escolar do aluno': 'Ano',
    'ano escolar do aluno:': 'Ano',
    'quantos pontos o aluno fez': 'Pontuacao',
    'total de pontuacao': 'Pontuacao',
    'total de pontuação': 'Pontuacao',
    'quanto tempo de realizacao': 'Tempo',
    'quanto tempo de realização': 'Tempo',
    'se for aluno com deficiencia/transtorno, escolha a categoria da olimpiada que o(a) aluno(a) se encaixa': 'DefTran',
    'se for aluno com deficiencia/transtorno': 'DefTran',
    'se for aluno com deficiência/transtorno': 'DefTran',
    'quer deixar uma mensagem? pode usar o espaco abaixo': 'Mensagem',
    'quer deixar uma mensagem? pode usar o espaço abaixo': 'Mensagem',
}

def mapear_colunas(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for c in df.columns:
        key = _norm(c)
        mapped = None
        if key in _COL_CANON:
            mapped = _COL_CANON[key]
        else:
            for k, v in _COL_CANON.items():
                if k in key:
                    mapped = v
                    break
        renamed[c] = mapped if mapped else c
    return df.rename(columns=renamed)

def aplicar_formatacao_basica(writer, sheet_name, df, header_row_idx=0, col_width_chars=18):
    """
    - Aplica estilo ao cabeçalho
    - Define largura **uniforme** das colunas (padrão 18 chars) para deixar o banner consistente
    - Retorna lista com larguras em pixels (aprox) para posicionar o banner
    """
    ws = writer.sheets[sheet_name]
    book = writer.book

    header_fmt = book.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#6AA84F',
        'font_color': 'white',
        'border': 1
    })

    # escreve cabeçalho com estilo
    for col_idx, col_name in enumerate(df.columns):
        ws.write(header_row_idx, col_idx, col_name, header_fmt)

    # largura **fixa** para todas as colunas -> consistência do banner entre abas
    for col_idx in range(df.shape[1]):
        ws.set_column(col_idx, col_idx, col_width_chars)

    # conversão aproximada chars -> pixels (fator ~7)
    col_pixels = [int(col_width_chars * 7) for _ in range(df.shape[1])]
    return col_pixels

from xlsxwriter.utility import xl_col_to_name

def inserir_banner(ws, image_bytes, col_widths_px, cols, banner_rows=3, target_height_px=110):
    """
    Insere a imagem centralizada dentro da área MESCLADA A1:.. (banner_rows linhas),
    escalando para caber e centralizando **horizontal e verticalmente**.
    Usa object_position=1 (move e redimensiona com as células) para o Google Sheets.
    """
    if not image_bytes:
        return

    # 1) Mescla A1:LastCol<banner_rows>
    last_col_name = xl_col_to_name(cols - 1)
    merge_range = f"A1:{last_col_name}{banner_rows}"
    ws.merge_range(merge_range, "", None)

    # 2) Ajusta altura das linhas da faixa do banner (total = target_height_px)
    for r in range(banner_rows):
        ws.set_row(r, target_height_px / banner_rows)

    # 3) Largura total da área em pixels
    total_px = sum(col_widths_px[:cols]) if col_widths_px else 600

    # 4) Escala e offsets (centro)
    with Image.open(BytesIO(image_bytes)) as im:
        img_w, img_h = im.size

    if img_w <= 0 or img_h <= 0:
        x_scale = y_scale = 1.0
        x_offset = y_offset = 0
    else:
        scale_w = total_px / img_w
        scale_h = target_height_px / img_h
        scale = min(scale_w, scale_h, 1.0)  # não amplia

        out_w = img_w * scale
        out_h = img_h * scale
        x_offset = max(int((total_px - out_w) / 2), 0)
        y_offset = max(int((target_height_px - out_h) / 2), 0)

        x_scale = y_scale = scale

    # 5) Insere na A1, “dentro” da área mesclada e seguindo as células
    ws.insert_image(
        "A1",
        "banner.png",
        {
            "image_data": BytesIO(image_bytes),
            "x_scale": x_scale,
            "y_scale": y_scale,
            "x_offset": x_offset,
            "y_offset": y_offset,
            "object_position": 1,  # **move e redimensiona com as células** (Sheets fica estável)
        },
    )


# ------------------ Pipeline de dados ------------------
def gerar_classificatoria(formulario_df: pd.DataFrame, etapa: str) -> pd.DataFrame:
    df = mapear_colunas(formulario_df).copy()

    # Precisamos destas (canon):
    req = ['Nome', 'EscolaSel', 'EscolaLivre', 'Ano', 'Pontuacao', 'Tempo', 'DefTran']
    faltando = [c for c in req if c not in df.columns]
    if faltando:
        raise KeyError(f"Colunas ausentes: {faltando}. Recebidas: {list(df.columns)}")

    work = df[req].copy()

    # "Escola não está na lista" -> usa EscolaLivre
    escola_sel_norm = work['EscolaSel'].astype(str).str.strip().str.lower()
    marcadores_out = {
        'escola não está na lista', 'escola nao esta na lista', 'escola nao está na lista',
        'outros', 'outro'
    }
    mask_out = escola_sel_norm.isin(marcadores_out)
    work.loc[mask_out, 'EscolaSel'] = work['EscolaLivre']
    work.drop(columns=['EscolaLivre'], inplace=True)

    # Renomeia finais
    work = work.rename(columns={
        'Nome': 'Nome',
        'EscolaSel': 'Escola',
        'Ano': 'Ano',
        'Pontuacao': 'Pontuação',
        'Tempo': 'Tempo',
        'DefTran': 'Deficiência/Transtorno'
    })

    # Normalizações
    work['Nome'] = work['Nome'].astype(str).str.upper()
    work['Escola'] = work['Escola'].apply(padronizar_nome_escola)
    work['Pontuação'] = work['Pontuação'].apply(padronizar_pontuacao)
    work['Tempo_seg'] = work['Tempo'].apply(_parse_tempo)

    # Ordenação
    work['Ordem_Ano'] = work['Ano'].apply(obter_ordem_ano)
    work = work.sort_values(by=['Ordem_Ano', 'Pontuação', 'Tempo_seg'],
                            ascending=[True, False, True]).drop(columns=['Ordem_Ano'])

    work['ETAPA'] = etapa
    # Mantém Tempo_seg para ordenação/debug; removemos na exportação
    work = work[['Ano', 'Nome', 'Escola', 'Pontuação', 'Tempo', 'Deficiência/Transtorno', 'ETAPA', 'Tempo_seg']]
    return work

# ------------------ Escrita em Excel ------------------
def escrever_geral(writer, df, image_bytes=None, banner_rows=3, banner_h_px=110):
    sheet = 'GERAL'
    header_row = banner_rows if image_bytes else 0
    df_export = df.drop(columns=['Tempo_seg'], errors='ignore')
    df_export.to_excel(writer, sheet_name=sheet, index=False, startrow=header_row)
    ws = writer.sheets[sheet]

    # Formatação + larguras reais
    col_pixels = aplicar_formatacao_basica(writer, sheet, df_export, header_row_idx=header_row)

    # Banner dentro da célula mesclada A1:.. (se houver imagem)
    if image_bytes:
        inserir_banner(ws, image_bytes, col_widths_px=col_pixels, cols=df_export.shape[1],
                       banner_rows=banner_rows, target_height_px=banner_h_px)

    # Filtros e freeze
    ws.autofilter(header_row, 0, header_row + len(df_export), df_export.shape[1] - 1)
    ws.freeze_panes(header_row + 1, 0)

def escrever_por_escola(writer, df, image_bytes=None, banner_rows=3, banner_h_px=110):
    # ↓↓↓ ORDEM ALFABÉTICA nas abas por escola
    escolas = sorted(df['Escola'].dropna().unique())

    usados = set(['GERAL'])  # mantém GERAL reservado; evita conflito de nome
    book = writer.book

    for escola in escolas:
        sheet = ajustar_nome_aba(escola, usados)
        title_row = banner_rows if image_bytes else 0
        header_row = title_row + 1

        df_esc = df[df['Escola'] == escola].drop(columns=['Tempo_seg'], errors='ignore')
        df_esc.to_excel(writer, sheet_name=sheet, index=False, startrow=header_row)
        ws = writer.sheets[sheet]

        col_pixels = aplicar_formatacao_basica(writer, sheet, df_esc, header_row_idx=header_row)

        if image_bytes:
            inserir_banner(ws, image_bytes, col_widths_px=col_pixels, cols=df_esc.shape[1],
                           banner_rows=banner_rows, target_height_px=banner_h_px)

        ws.merge_range(title_row, 0, title_row, df_esc.shape[1]-1, escola, book.add_format({
            'align': 'center', 'bold': True, 'bg_color': '#2E7D32', 'font_color': 'white', 'border': 1
        }))

        ws.autofilter(header_row, 0, header_row + len(df_esc), df_esc.shape[1]-1)
        ws.freeze_panes(header_row + 1, 0)


def salvar_excels(classificatoria_df, image_bytes=None, banner_rows=3, banner_h_px=110):
    out_olimpiada, out_paralimpiada, out_juncao = BytesIO(), BytesIO(), BytesIO()

    # Normaliza campo de deficiência
    def _norm_low(s): return _norm(s)

    sem_def_flags = {
        'nao possui deficiencia/transtorno',
        'não possui deficiência/transtorno',
        'sem deficiencia',
        'sem deficiência',
        'nao', 'não', 'n'
    }

    base = classificatoria_df.copy()
    base['__def_norm'] = base['Deficiência/Transtorno'].map(_norm_low)

    olimpiada_df = base[base['__def_norm'].isin(sem_def_flags)].drop(columns=['__def_norm'])
    paralimpiada_df = base[~base.index.isin(olimpiada_df.index)].drop(columns=['__def_norm'])

    # 1) Olimpíada
    with pd.ExcelWriter(out_olimpiada, engine='xlsxwriter') as writer:
        escrever_geral(writer, olimpiada_df, image_bytes=image_bytes, banner_rows=banner_rows, banner_h_px=banner_h_px)
        escrever_por_escola(writer, olimpiada_df, image_bytes=image_bytes, banner_rows=banner_rows, banner_h_px=banner_h_px)

    # 2) Paralimpíada
    with pd.ExcelWriter(out_paralimpiada, engine='xlsxwriter') as writer:
        escrever_geral(writer, paralimpiada_df, image_bytes=image_bytes, banner_rows=banner_rows, banner_h_px=banner_h_px)
        escrever_por_escola(writer, paralimpiada_df, image_bytes=image_bytes, banner_rows=banner_rows, banner_h_px=banner_h_px)

    # 3) JUNÇÃO
    base_export = base.drop(columns=['__def_norm'], errors='ignore')
    with pd.ExcelWriter(out_juncao, engine='xlsxwriter') as writer:
        escrever_geral(writer, base_export, image_bytes=image_bytes, banner_rows=banner_rows, banner_h_px=banner_h_px)
        escrever_por_escola(writer, base_export, image_bytes=image_bytes, banner_rows=banner_rows, banner_h_px=banner_h_px)

    out_olimpiada.seek(0); out_paralimpiada.seek(0); out_juncao.seek(0)
    return out_olimpiada, out_paralimpiada, out_juncao

# ------------------ App ------------------
def main():
    st.title("Tabulação: Gerador de Classificatória por Escola")

    st.write("Exemplo da estrutura de dados esperada (podem existir variações):")
    exemplo_data = {
        "Nome do aluno?": ["Aluno 1", "Aluno 2", "Aluno 3"],
        "Qual é o nome da sua escola?": ["EMEF Exemplo", "Escola não está na lista", "EMEF Central"],
        "Escreva o nome da escola caso ela no esteja listada": ["EMEF Nova Esperança", "", ""],
        "Ano escolar do aluno:": ["1° ANO", "2° ANO", "EJAI 2ª ETAPA"],
        "Total de pontuação?": [45, 38, 50],
        "Quanto tempo de realização?": ["00:15:00", "12:30", "540"],
        "Se for aluno com deficiência/transtorno:": [
            "Não possui deficiência/transtorno", "Deficiência física", "N"
    ]
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

        usar_banner = st.checkbox("Adicionar imagem no topo (todas as abas)", value=True)
        banner_altura = st.slider("Altura do banner (px)", min_value=60, max_value=220, value=110, step=10)
        banner_linhas = st.slider("Linhas reservadas para o banner", min_value=2, max_value=5, value=3, step=1)

        image_bytes = None
        if usar_banner:
            img_file = st.file_uploader("Envie a imagem (PNG/JPG), opcional", type=["png", "jpg", "jpeg"])
            if img_file is not None:
                image_bytes = img_file.read()

        try:
            classificatoria_df = gerar_classificatoria(formulario_df, etapa)
        except KeyError as e:
            st.error(f"Planilha não está no formato esperado: {e}")
            return

        st.write("Dados filtrados e ordenados:")
        st.dataframe(classificatoria_df)

        if st.button("Gerar Arquivos"):
            out_olimp, out_para, out_junc = salvar_excels(
                classificatoria_df,
                image_bytes=image_bytes,
                banner_rows=banner_linhas,
                banner_h_px=banner_altura
            )
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button("Baixar Alunos Olimpíada", out_olimp, "classificatoria_olimpiada.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                st.download_button("Baixar Alunos Paralimpíada", out_para, "classificatoria_paralimpiada.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col3:
                st.download_button("Baixar JUNÇÃO (Olimpíada + Paralimpíada)", out_junc, "classificatoria_juncao.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    main()
