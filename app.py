# app.py
import streamlit as st
from juntarSheets import main as juntar_sheets_main
from tabulacaoOlimpiadasEParalimpada import main as tabulacao_main
from classificacaoMelhoresColocados import main as classificacao_tabulacao_main
from semifinal import main as seminifinal_main
from final import main as final_main

# Configuração da página
st.set_page_config(
    page_title="Inteceleri - Pedagógico",  
    page_icon="📘"
)

# Inicialização ou reinicialização da tela inicial
def set_initial_state():
    st.session_state.current_page = 'home'

if 'current_page' not in st.session_state:
    set_initial_state()

# Adicionando os botões na barra lateral
st.sidebar.title("Menu")

if st.sidebar.button('Combinar Abas Sheets'):
    st.session_state.current_page = 'juntar_sheets' #juntarSheets.py

if st.sidebar.button('Tabulação Olimpiada e Paralimpiada '): #tabulacaoOlimpiadasEParalimpada.py
    st.session_state.current_page = 'tabulacao'

if st.sidebar.button('Classificação Pontuação/Tabulação'):
    st.session_state.current_page = 'classificacaoMelhoresColocados' #classificacaoMelhoresColocados.py

if st.sidebar.button('Semifinal '): #semifinal.py
    st.session_state.current_page = 'semifinal'

if st.sidebar.button('Final '): #final.py
    st.session_state.current_page = 'final'


# Mostrando conteúdos baseados no estado
if st.session_state.current_page == 'juntar_sheets': #juntarSheets.py
    juntar_sheets_main()
elif st.session_state.current_page == 'tabulacao': #tabulacaoOlimpiadasEParalimpada.py
    tabulacao_main()
elif st.session_state.current_page == 'classificacaoMelhoresColocados': #classificacaoMelhoresColocados.py
    classificacao_tabulacao_main()
elif st.session_state.current_page == 'semifinal': #semifinal.py
    seminifinal_main()
elif st.session_state.current_page == 'final': #final.py
    final_main()
elif st.session_state.current_page == 'home':
    st.title("Inteleceleri - Pedagógico")
    st.write("Por favor, escolha uma opção ao lado para visualizar os dados.")

# Espaços adicionais na barra lateral, para estética
for _ in range(20):
    st.sidebar.write("")

# Assinatura no final da barra lateral
st.sidebar.markdown("---")
st.sidebar.markdown("*Desenvolvido por: Inteceleri*", unsafe_allow_html=True)
