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

- **Descrição**: [Funcionalidade em desenvolvimento]
- **Objetivo**: Esta opção será destinada à geração dos dados finais para a última etapa da competição.

## Requisitos

Este projeto requer **Python 3.12.3** e as seguintes bibliotecas para ser executado corretamente:

- **pandas==2.2.2**
- **seaborn** (para visualizações estatísticas, pode ser instalado com `pip install seaborn`)
- **matplotlib==3.9.0**
- **streamlit==1.34.0**
- **plotly==5.22.0**
- **openpyxl==3.1.2** (para leitura e escrita de arquivos Excel)
- **xlsxwriter==3.2.0** (para gerar arquivos Excel com múltiplas abas)

Para instalar todas as dependências necessárias, execute o seguinte comando:

```bash
pip install -r requirements.txt
```

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