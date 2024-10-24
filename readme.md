# Gerador de Classificatória por Escola

Este projeto é uma aplicação web construída com **Streamlit** e **Pandas** que permite processar os dados de um formulário de respostas, filtrá-los por escola e gerar um arquivo Excel com múltiplas planilhas, cada uma representando uma escola. A ferramenta também permite o download do arquivo Excel finalizado.

## Funcionalidades

- **Upload de arquivo**: O usuário pode enviar um arquivo Excel contendo o formulário de respostas.
- **Filtragem por escola**: A aplicação filtra os dados com base na coluna "Escola" e exibe uma tabela separada para cada escola.
- **Exibição dinâmica**: As tabelas de cada escola são exibidas diretamente na página, com o nome da escola em destaque.
- **Geração de arquivo Excel**: O sistema cria um arquivo Excel com uma aba (sheet) para cada escola, incluindo uma linha inicial com o nome da escola.
- **Download do arquivo gerado**: O arquivo Excel final pode ser baixado diretamente da interface da aplicação.

## Requisitos

Antes de rodar o projeto, você precisará instalar os seguintes pacotes:

- **Streamlit**: Para a criação da interface web.
- **Pandas**: Para manipulação de dados tabulares.
- **XlsxWriter**: Para salvar o arquivo Excel com múltiplas abas.

Instale as dependências com o seguinte comando:

```bash
pip install streamlit pandas xlsxwriter
```

## Como rodar a aplicação
```bash
streamlit run app.py
```

## Como usar

1. Upload do arquivo
Acesse a aplicação no seu navegador.

Faça o upload do arquivo Excel que contém o Formulário de Respostas. Certifique-se de que o arquivo tenha as seguintes colunas:

- Nome do aluno?
- Qual é o nome da sua escola?
- Escreva o nome da escola caso ela no esteja listada
- Ano escolar do aluno:
- Total de pontuação?
- Quanto tempo de realização?
- Se for aluno com deficiência/transtorno:

2. Exibição dos dados: 
Após o upload, a aplicação processará os dados e exibirá uma tabela para cada escola, com o nome da escola em destaque.

A tabela contém as seguintes colunas:

- Ano
- Nome
- Escola
- Pontuação
- Tempo
- Se for aluno com deficiência/transtorno
- ETAPA (preenchida automaticamente como "2° CLASSIFICATÓRIA")

3. Download do arquivo gerado: 
Após a exibição dos dados, você poderá clicar no botão para gerar e baixar o arquivo Excel.

A estrutura da planilha será:

- Ano
- Nome
- Escola
- Pontuação
- Tempo
- Se for aluno com deficiência/transtorno
- ETAPA (com o valor "2° CLASSIFICATÓRIA")



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