# 🚀 Fast Control — Controladoria CLI

Automação confiável e rápida de rotinas da Controladoria via terminal interativo, com módulos para extração de APs (SQL Server), processamento de benefícios (concatenação e rateio) e extração de transitórias com mapeamento contábil.

<p align="left">
  <a href="https://www.python.org/"><img alt="Python" src="https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white"></a>
  <img alt="OS" src="https://img.shields.io/badge/OS-Windows-blue?logo=windows">
  <img alt="License" src="https://img.shields.io/badge/License-Interno-informational">
  <a href="#-instala%C3%A7%C3%A3o"><img alt="Installer" src="https://img.shields.io/badge/Install-Poetry%2Fpip-success"></a>
</p>

---


## 🧭 Visão Geral
- 1️⃣ Base de Dados: extração de APs direto do SQL Server por período (competência).
- 2️⃣ Benefícios: concatenação de planilhas e geração de planilha de rateio por benefício.
- 3️⃣ Transitórias: extração por APS a partir de base local com mapeamento custo/despesa → conta contábil.
- 4️⃣ Provisões: geração de lançamentos de provisão e reversão para ERP a partir de planilha horizontal.

---

- [📦 Requisitos](#-requisitos)
- [⚙️ Instalação](#️-instalação)
- [🔐 Configuração (.env)](#-configuração-env)
- [⚡ Quickstart](#-quickstart)
- [🧰 Referência de Comandos](#-referência-de-comandos)
- [🗂️ Estrutura de Pastas](#️-estrutura-de-pastas)
- [🛠️ Dicas e Solução de Problemas](#️-dicas-e-solução-de-problemas)
- [❓ FAQ](#-faq)
- [📬 Suporte](#-suporte)
- [📄 Licença](#-licença)

---

## 📦 Requisitos
- Python 3.9+
- Windows com ODBC Driver do SQL Server instalado (ODBC Driver 17+ recomendado)
- Acesso ao SQL Server (para o módulo “Base de Dados”)
- Pacotes Python: `pandas`, `pyodbc`, `pydantic-settings`, `xlsxwriter`, `colorama`, `openpyxl`

---

## ⚙️ Instalação

### Usando Poetry (recomendado)
```powershell
# Na raiz do projeto
poetry install

# Garante libs usadas pelo código
poetry add colorama openpyxl
```

### Usando pip (alternativa)
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -U pip
pip install -r requirements.txt
```

> Dica: Para `pyodbc` funcionar, instale o ODBC Driver do SQL Server (Painel de Controle → Ferramentas Administrativas → Fontes de Dados ODBC) ou baixe do site da Microsoft.

---

## 🔐 Configuração (.env)
O módulo “Base de Dados” lê variáveis via `pydantic-settings` (`controladoria_cli/core/config/settings.py`). Crie um arquivo `.env` na raiz:

```env
DB_DRIVER=SQL Server
DB_SERVER=SEU_SERVIDOR_SQL
DB_NAME=NOME_DO_BANCO
DB_USER=USUARIO
DB_PASSWORD=SENHA
```

- `DB_DRIVER` padrão: "SQL Server". Ajuste se necessário, por exemplo: `ODBC Driver 17 for SQL Server`.

---

## ⚡ Quickstart

### Executar via Poetry
```powershell
poetry run fast-control
```

### Executar via Python
```powershell
python controladoria_cli/main.py
# ou
python -c "from controladoria_cli.main import main; main()"
```


Menu interativo esperado:
```
--- FAST CONTROL ---
1 - Base de Dados: Extrair APs
2 - Benefícios
3 - Transitórias: Extrair APs com Mapeamento
4 - Provisões
0 - Sair
```

---

## 🧰 Referência de Comandos

### 1️⃣ Base de Dados: Extrair APs
- Arquivo: `controladoria_cli/commands/base_dados.py`
- Serviço: `controladoria_cli/core/services/base_dados_service.py`
- O que faz:
  - Solicita período (ano/mês inicial e final)
  - Conecta ao SQL Server e busca dados “superior” e “inferior” (rateios)
  - Faz merge, normaliza textos e exporta para Excel
- Configuração de saída: edite `CAMINHO_PASTA_SAIDA` em `base_dados.py`.
- Saída exemplo:
  - `controladoria_cli/dados_aps/output_dados_base_aps/base_aps_YYYYMMDD_HHMMSS.xlsx`

Pré‑requisitos:
- `.env` configurado e conectividade com o SQL Server.


### 2️⃣ Benefícios
- Menu: `controladoria_cli/commands/beneficios.py`
- Serviço: `controladoria_cli/core/services/beneficios_service.py`

Ferramentas:
- 🔗 Concatenação de planilhas
  - Entradas: `CAMINHO_CONCATENACAO_ENTRADA`
  - Saída: `CAMINHO_CONCATENACAO_SAIDA/arquivo_concatenado_gerado.xlsx`
  - Ajuste no topo do arquivo de serviço.

- 📊 Rateio de benefícios
  - Entrada: `ARQUIVO_RATEIO_ENTRADA`
  - Saída: arquivo Excel em `ARQUIVO_RATEIO_SAIDA` com abas por benefício (VR, CB, VA, OB, AJ, VT, PJ, AUX MORADIA, DESPESAS ACIONISTAS, D‑BÔNUS, Taxa)
  - Ajuste no topo do arquivo de serviço.

Observações:
- Exija as colunas esperadas na base; leitura de `.xlsx` usa `openpyxl`.


### 3️⃣ Transitórias: Extrair APs com Mapeamento
- Menu: `controladoria_cli/commands/transitorias.py`
- Serviço: `controladoria_cli/core/services/transitorias_services.py`
- Mapeamento: `controladoria_cli/core/config/account_mapping.py`

O que faz:
- Lê base local `data_base_de_aps.xlsx`
- Filtra por lista de valores “APS” informados no terminal
- Gera coluna `conta contabil` a partir de `NATUREZA` e `CUSTO OU DESPESA`
- Oferece salvar o resultado (Excel) na pasta configurada


Configurações importantes no serviço:
- `CAMINHO_ARQUIVO_ENTRADA`: caminho para `data_base_de_aps.xlsx`
- `CAMINHO_PASTA_SAIDA`: pasta onde os resultados serão gravados

### 4️⃣ Provisões: Geração de Lançamentos ERP
- Menu: `controladoria_cli/commands/provisoes.py`
- Serviço: `controladoria_cli/core/services/provisoes_service.py`

O que faz:
- Lê planilha horizontal de provisões (clientes x valores).
- Aplica de-para contábil e gera lançamentos verticais para ERP (provisão e reversão).
- Salva arquivo Excel com abas por empresa/filial.

Configurações importantes no serviço:
- Caminho do arquivo de entrada e saída ajustável no topo do script.

---

## 🗂️ Estrutura de Pastas
```text
controladoria_cli/
  __init__.py                    # Marca o pacote Python
  main.py                        # Entrypoint e menu principal da CLI
  commands/
    __init__.py                  # Marca o subpacote de comandos
    base_dados.py                # Automatiza extração de APs via SQL Server
    beneficios.py                # Submenu de benefícios e chamada de serviços
    transitorias.py              # Extração de APs transitórias com mapeamento
    provisoes.py                 # Geração de lançamentos de provisão/reversão para ERP
  core/
    __init__.py                  # Marca o subpacote core
    config/
      __init__.py                # Marca o subpacote de configuração
      settings.py                # Carrega configuração de banco via .env
      account_mapping.py         # Dicionários e mapeamentos de contas contábeis
    services/
      __init__.py                # Marca o subpacote de serviços
      base_dados_service.py      # Conexões SQL, queries e processamento das APs
      beneficios_service.py      # Concatenação e geração de rateio de benefícios
      transitorias_services.py   # Filtragem de APS e criação da conta contábil
      provisoes_service.py       # Processamento de provisões e geração de lançamentos ERP
  dados_aps/
    data_base_de_aps.xlsx         # Base local de APs usada pelo módulo de Transitórias
    output_dados_aps/             # Saída dos arquivos de Transitórias
    output_dados_base_aps/        # Saída dos arquivos de Base de Dados
  dados_beneficios/
    concatenacao/
      CONCAT/                     # Pasta de saída da concatenação de planilhas
    rateio/                      # Pasta de entrada e saída do rateio de benefícios
pyproject.toml
README.md
requirements.txt
```

---

## 🧩 Arquitetura e Função de Cada Arquivo
### `pyproject.toml`
- Configura o projeto Python com Poetry.
- Declara dependências essenciais: `pandas`, `pyodbc`, `pydantic-settings`, `xlsxwriter`, `colorama`, `openpyxl`.
- Define o script CLI `fast-control` apontando para `controladoria_cli.main:main`.

### `requirements.txt`
- Lista de dependências para instalação via `pip`.

### `controladoria_cli/__init__.py`
- Arquivo vazio que torna `controladoria_cli` um pacote Python.

### `controladoria_cli/main.py`
- Ponto de entrada da aplicação.
- Exibe o menu principal e roteia para os módulos:
  - Base de Dados
  - Benefícios
  - Transitórias
- Captura exceções de cada módulo e mantém o loop interativo.

### `controladoria_cli/commands/__init__.py`
- Arquivo vazio que torna `commands` um subpacote.

### `controladoria_cli/commands/base_dados.py`
- Interface interativa para extrair APs do SQL Server.
- Solicita período de competência inicial/final ao usuário.
- Usa `BaseDadosService` para obter e processar dados.
- Grava o resultado em Excel no diretório configurado por `CAMINHO_PASTA_SAIDA`.

### `controladoria_cli/commands/beneficios.py`
- Submenu de Benefícios com duas opções:
  - Concatenar planilhas de entrada
  - Executar rateio de benefícios
- Chama funções de `beneficios_service`.

### `controladoria_cli/commands/transitorias.py`
- Coleta valores de `APS` do usuário.
- Processa a extração de APs usando `transitorias_services`.
- Mostra pré-visualização dos resultados e pergunta se deseja salvar em Excel.

### `controladoria_cli/core/__init__.py`
- Arquivo vazio que torna `core` um subpacote.

### `controladoria_cli/core/config/settings.py`
- Usa `pydantic-settings` para carregar variáveis de ambiente de `.env`.
- Define parâmetros de conexão ao banco de dados:
  - `DB_DRIVER`
  - `DB_SERVER`
  - `DB_NAME`
  - `DB_USER`
  - `DB_PASSWORD`

### `controladoria_cli/core/config/account_mapping.py`
- Define dois dicionários principais:
  - `Conta_custo`: mapeia contas de custo por natureza
  - `conta_despesa`: mapeia contas de despesa por natureza
- Constrói `mapeamento_custo` e `mapeamento_despesa` para lookup rápido.
- Usado pelo módulo de Transitórias para criar a coluna `conta contabil`.

### `controladoria_cli/core/services/__init__.py`
- Marca o pacote de serviços.

### `controladoria_cli/core/services/base_dados_service.py`
- Serviço responsável por conectar ao SQL Server via `pyodbc`.
- Executa duas queries principais:
  - `sql_sup`: dados superiores das APs dentro do período
  - `sql_inf`: dados inferiores associados às APs filtradas
- Processa e une os resultados com `pandas`.
- Trata casos onde não há dados inferiores preenchendo valores a partir do superior.
- Limpa texto para remover caracteres especiais e normalizar colunas.
- Lança erros de conexão e quaisquer falhas inesperadas.

### `controladoria_cli/core/services/beneficios_service.py`
- `executar_concatenacao()`:
  - Lê todos os arquivos `.xlsx` em `CAMINHO_CONCATENACAO_ENTRADA`.
  - Adiciona coluna `Arquivo Original` e concatena os dados.
  - Salva arquivo único em `CAMINHO_CONCATENACAO_SAIDA`.
- `executar_rateio()`:
  - Lê base de rateio em `ARQUIVO_RATEIO_ENTRADA`.
  - Agrupa e soma valores para cada benefício listado.
  - Gera planilhas separadas no arquivo de saída `ARQUIVO_RATEIO_SAIDA`.
- Inclui tratamento de arquivos inválidos e avisos de colunas faltantes.

### `controladoria_cli/core/services/transitorias_services.py`
### `controladoria_cli/core/services/provisoes_service.py`
- `processar_planilha_provisoes(caminho_arquivo, ano, mes)`:
  - Lê planilha horizontal de provisões.
  - Aplica de-para contábil e gera lançamentos de provisão e reversão.
  - Retorna DataFrame pronto para exportação.
- `salvar_layout_erp_por_empresa(df, filename)`:
  - Salva o DataFrame em Excel, criando uma aba para cada empresa/filial.
  - Aplica formatação automática e destaca cabeçalhos.
- `processar_extracao_aps_transitarias(valores_aps)`:
  - Lê arquivo Excel de entrada `CAMINHO_ARQUIVO_ENTRADA`.
  - Filtra linhas pela coluna `APS`.
  - Aplica mapeamento de contas por `NATUREZA` e `CUSTO OU DESPESA`.
  - Retorna o DataFrame filtrado e a lista de APS não encontradas.
- `_determinar_conta_contabil(row)`:
  - Decide se a linha é custo ou despesa.
  - Mapeia o valor de natureza para conta contábil apropriada.
- `salvar_dataframe_transitarias(df, file_name)`:
  - Salva o resultado em Excel no diretório `CAMINHO_PASTA_SAIDA`.

### `controladoria_cli/dados_aps/data_base_de_aps.xlsx`
- Base local de dados de APs usada pelo módulo de Transitórias.
- Deve conter pelo menos as colunas `APS`, `NATUREZA` e `CUSTO OU DESPESA`.

### `controladoria_cli/dados_aps/output_dados_aps/`
- Local de saída padrão para arquivos gerados pelo módulo de Transitórias.

### `controladoria_cli/dados_aps/output_dados_base_aps/`
- Local de saída padrão para arquivos gerados pelo módulo de Base de Dados.

### `controladoria_cli/dados_beneficios/concatenacao/`
- Diretório de entrada para arquivos de benefícios a serem concatenados.
- Subpasta `CONCAT/` é usada para salvar o arquivo concatenado.

### `controladoria_cli/dados_beneficios/rateio/`
- Diretório de entrada e saída para o processo de rateio.
- Deve conter `base_rateio.xlsx` como fonte dos dados.

---

## 🛠️ Dicas e Solução de Problemas
- ❗ Conexão `pyodbc`/SQL Server
  - Verifique ODBC Driver instalado e visível em “Fontes de Dados ODBC”.
  - Confira `DB_SERVER`, credenciais e firewall.
  - A string de conexão usa `TrustServerCertificate=yes;` por padrão.
- 📁 Caminhos absolutos
  - Ajuste as constantes `CAMINHO_*` nos serviços conforme seu ambiente local.
- 📄 Excel
  - Instale `openpyxl` para leitura/escrita `.xlsx`.
- 🌐 Acentos/encoding
  - O serviço normaliza textos para evitar caracteres problemáticos.
- 🎨 Cores no terminal
  - Instale `colorama` se as cores não aparecerem.

Erros comuns:
- `pyodbc.Error`: checar driver ODBC, rede e autenticação.
- `FileNotFoundError` em planilhas/caminhos: revisar `CAMINHO_*` nos serviços.
- Colunas ausentes: garantir que a base contenha os nomes exatos esperados.

---

## ❓ FAQ
- Posso executar sem Poetry?
  - Sim, usando `pip` e `python controladoria_cli/main.py`.
- Onde ajusto os caminhos de entrada/saída?
  - No topo de `beneficios_service.py`, `transitorias_services.py` e em `commands/base_dados.py`.
- Como rodar com um comando curto?
  - Com Poetry: `poetry run fast-control` (entrypoint definido em `[tool.poetry.scripts]`).
- Como fixar versões das libs?
  - Use `requirements.txt` (já incluso) ou `poetry.lock` quando usar Poetry.

---

## 📬 Suporte
Em caso de dúvidas ou sugestões, registre um issue no repositório interno ou contate o responsável pelo projeto.

---

## 📄 Licença
Uso interno. Todos os direitos reservados.



