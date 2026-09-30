# Inteceleri - Área Pedagogica

Este projeto é uma aplicação web desenvolvida com **Streamlit** e **Pandas**, projetada para processar dados e automatizar processos, reduzindo demandas manuais no setor pedagógico.

## Funcionalidades

Abaixo estão as funcionalidades disponíveis na aplicação:

### Combinar Abas Sheets

- **Descrição**: Permite fazer upload de um arquivo Excel com múltiplas abas. A aplicação combina os dados de todas as abas em um único DataFrame e fornece uma tabela dinâmica para análise.
- **Recursos**:
  - Exibe um exemplo da estrutura de dados esperada.
  - Combina os dados de todas as abas do arquivo.
  - Gera uma tabela dinâmica para contar a quantidade de alunos por escola e ano.
  - Permite o download dos dados combinados e da tabela dinâmica em formatos CSV e XLSX.
- **Download**: Arquivo consolidado dos dados combinados e da tabela dinâmica.
- **Estrutura de Dados Necessária**:
  - Ano: O ano escolar do aluno (por exemplo, "1ª ANO").
  - Nome: Nome do aluno.
  - Escola: Nome da escola.
  - Pontuação: Pontuação obtida pelo aluno.
  - Tempo: Tempo de realização (formato HH:MM:SS).
  - Se for aluno com deficiência/transtorno: Indicação se o aluno possui alguma deficiência ou transtorno.
  - Etapa de Classificação: Indicação da etapa (por exemplo, "1º CLASSIFICATÓRIA").

### Tabulação Olimpíada e Paralimpíada

- **Descrição**: Processa os dados do formulário geral de respostas enviado pelos professores, separando-os entre alunos participantes da Olimpíada e da Paralimpíada.
- **Recursos**:
  - Filtra os alunos com e sem deficiência/transtorno.
  - Organiza as respostas em abas separadas por escola, ordenando por pontuação (decrescente) e tempo (crescente).
  - Gera arquivos Excel separados para os alunos da Olimpíada e da Paralimpíada.
- **Download**: Dois arquivos Excel – um para a Olimpíada e outro para a Paralimpíada, com uma aba para cada escola.
- **Estrutura de Dados Necessária**:
  - Nome do aluno?: Nome do aluno.
  - Qual é o nome da sua escola?: Nome da escola do aluno ou "Escola não está na lista" se não estiver na lista.
  - Escreva o nome da escola caso ela não esteja listada: Nome alternativo da escola, caso não esteja na lista.
  - Ano escolar do aluno: Ano escolar do aluno.
  - Total de pontuação?: Pontuação total obtida pelo aluno.
  - Quanto tempo de realização?: Tempo de realização (formato HH:MM:SS).
  - Se for aluno com deficiência/transtorno: Indicação se o aluno possui alguma deficiência ou transtorno.


### Classificação Pontuação/Tabulação

- **Descrição**: Permite a seleção dos melhores alunos de cada ano em uma aba específica, com base na pontuação e no tempo de realização. É possivel especificar se é para separar 1 a 5 alunos. Ex: "quero separar os 2 melhores dessa listagem", ele separa 2 alunos.
- **Recursos**:
  - Classifica os alunos de cada ano por pontuação (decrescente) e tempo (crescente).
  - Filtra os melhores alunos de cada ano, conforme o número selecionado pelo usuário.
  - Gera um arquivo Excel com a classificação dos melhores alunos por escola.
  - Exibe um gráfico interativo e uma tabela com a quantidade de alunos por ano escolar.
- **Download**: Arquivo Excel com a classificação dos melhores alunos.
- **Estrutura de Dados Necessária**:
  - Ano Escolar: Ano escolar do aluno.
  - Pontuação: Pontuação do aluno.
  - Tempo: Tempo de realização.

### Semifinal

- **Descrição**: Combina os dados das 1ª e 2ª classificatórias em um único arquivo, classifica-os e permite selecionar os melhores alunos para a fase semifinal.
- **Recursos**:
  - Permite o upload de dois arquivos (1ª e 2ª classificatórias).
  - Combina os dados das duas etapas em uma única tabela, organizando-os por ano, pontuação e tempo.
  - Permite que o usuário selecione o número de alunos a serem classificados para a fase semifinal.
  - Gera um arquivo Excel consolidado com a classificação organizada para cada escola.
- **Download**: Arquivo Excel com os dados combinados e organizados das duas classificatórias.
- **Estrutura de Dados Necessária**:
  - Mesmas colunas das etapas anteriores, dependendo da funcionalidade desejada.

### Final

- **Descrição**: Processa a listagem de respostas do formulário da etapa final e seleciona os melhores alunos de cada ano escolar de cada escola.
- **Recursos**:
  - Seleciona os **2 melhores alunos de cada ano escolar, de cada escola** (a quantidade pode ser alterada na tela).
  - Critérios de classificação: **maior pontuação** e, em caso de empate, **menor tempo de realização**.
  - Organiza a listagem pela ordem dos anos escolares (1º ano → 2º ano → 3º ano → ...).
  - Inclui a coluna **PROFESSOR** (professor que acompanhou o aluno), além do **RESPONSÁVEL**.
  - Normaliza os tempos digitados em formatos diferentes (ex.: `00:11: 52`, `00:14;00`, `02.41.60`) para `HH:MM:SS`.
  - Quando o mesmo aluno é enviado mais de uma vez (mesmo nome, escola e ano), mantém apenas o envio mais recente.
  - Mostra avisos para conferência manual: alunos repetidos, tempos fora do padrão e empates exatos no limite da classificação.
  - Logo do cliente opcional no topo de todas as abas, com altura e número de linhas ajustáveis (mesmo recurso da Tabulação Olimpíada e Paralimpíada).
- **Download**: Arquivo Excel no padrão da semifinal, com a aba **GERAL** (todos os classificados) e uma aba para cada escola.
- **Estrutura de Dados Necessária** (colunas do formulário):
  - Nome do aluno
  - Nome da escola onde você atua (ou o campo "Escreva o nome da escola caso ela não esteja listada")
  - Ano escolar do aluno
  - Quantos pontos o aluno fez?
  - Quanto tempo de realização?
  - Se for aluno com deficiência/transtorno
  - Qual o nome do professor que acompanhou o aluno durante a olimpíada?
  - Escreva o nome do responsável

## Requisitos

Este projeto requer **Python 3.13** e as seguintes bibliotecas para ser executado corretamente:

- **pandas==3.0.6**
- **seaborn==0.13.2** (para visualizações estatísticas)
- **matplotlib==3.11.2**
- **streamlit==1.64.0**
- **plotly==7.1.0**
- **openpyxl==3.1.5** (para leitura e escrita de arquivos Excel)
- **xlsxwriter==3.2.9** (para gerar arquivos Excel com múltiplas abas)

Para instalar todas as dependências necessárias, execute o seguinte comando:

```bash
pip install -r requirements.txt
```

## Como executar

Na pasta do projeto, execute:

```bash
streamlit run app.py
```

Se o comando `streamlit` não for reconhecido, use:

```bash
python -m streamlit run app.py
```

A aplicação abre no navegador em `http://localhost:8501`.

## 📝 Desenvolvido por
<table>
  <tr>
    <td align="center">
      <a href="https://inteceleri.com.br/" target="_blank" rel="external">
        <img src="https://avatars.githubusercontent.com/timedesenvolvimento-inteceleri" width="150px;" alt="Inteceleri Github Photo"/><br>
        <sub> 
          <b>Inteceleri </b><br>
          <b>Tecnologia para Educação</b><br>
        </sub>
      </a>
    </td>
  </tr>
</table>