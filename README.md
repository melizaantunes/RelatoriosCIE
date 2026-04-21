# Relatório de análises do CIE com Streamlit

A ideia foi desenvolver uma aplicação web simples em Streamlit para gerar relatórios a partir das planilhas de solicitações de análises do CIE.

A aplicação permite filtrar os dados por ano, isótopo, status de conclusão e gerar um gráfico de barras, além de exibir um mini-relatório em Markdown com os resultados.

## Funcionalidades

- Seleção de um ou mais anos
- Seleção de um ou mais isótopos
- Filtro para contabilizar:
  - todas as solicitações
  - somente as concluídas
- Geração opcional de gráfico de barras
- Mini-relatório automático em Markdown
- Suporte para*um ou mais arquivos Excel com a mesma estrutura
- Tratamento dos isótopos:
  - `CT` conta como `C`
  - `NT` conta como `N`
  - `NCT` conta como `NC`
- Regra de contagem:
  - `HO` conta como 1
  - os demais contam pelo número de elementos do isótopo
  - exemplo: `CO = 2`, `HCO = 3`, `NC = 2`

## Como executar o projeto

### 1. Clonar o repositório

```bash
git clone https://github.com/SEU_USUARIO/NOME_DO_REPO.git
cd NOME_DO_REPO

### 2. Criar um ambiente virtual
No Windows:

python -m venv .venv
.venv\Scripts\Activate.ps1

### 3. Instalar as dependências
pip install -r requirements.txt

### 4. Rodar a aplicação
streamlit run app.py
