# Conversor Excel -> CNAB 240 (Banco do Brasil)

Aplicacao desktop em Python para converter planilhas Excel em arquivos de remessa CNAB 240.

## Funcionalidades

- Conversao de Excel para CNAB 240 (segmentos A e B).
- Suporte a dois perfis de pagador no `config/config.json` (ex.: PGA e BBJUSMP).
- Leitura de planilhas de pagamentos e de aplicacoes.
- Geracao de arquivo `.txt` com linhas de 240 caracteres.
- Interface grafica em PyQt5.

## Requisitos

- Python 3.10+
- Dependencias em `requirements.txt`

## Instalacao

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Execucao

```bash
python main.py
```

## Build do executavel (opcional)

Ja existe o arquivo `CNAB Generator.spec` no projeto.

```bash
pyinstaller "CNAB Generator.spec"
```

O executavel sera gerado em `dist/CNAB Generator/`.

## Estrutura do projeto

```text
cnab_converter/
|-- main.py
|-- requirements.txt
|-- CNAB Generator.spec
|-- config/
|   `-- config.json
|-- gui/
|   |-- __init__.py
|   `-- app.py
|-- logic/
|   |-- __init__.py
|   |-- config_loader.py
|   |-- excel_reader.py
|   `-- cnab_generator.py
`-- visual/
```

## Configuracao (`config/config.json`)

O arquivo de configuracao contem:

- `arquivo.codigo_banco`
- `pagadores.PAGADOR_BPGA`
- `pagadores.PAGADOR_BBJUSMP`
- `aplicacoes` (contas por fundo)

Observacoes:

- `conta` deve ser informada sem mascara e sem separadores.
- `dv_conta` deve conter apenas o digito verificador.
- Nunca versione dados sensiveis reais em repositorios publicos.

## Formato esperado do Excel (pagamentos)

Os nomes das colunas podem variar. O sistema reconhece aliases para os campos abaixo:

- `nome`
- `cpf_cnpj`
- `banco`
- `agencia`
- `conta`
- `valor`
- `data_pagamento` (opcional)
- `identificador` (opcional)

Exemplo:

| nome | cpf_cnpj | banco | agencia | conta | valor | data_pagamento |
|------|----------|-------|---------|-------|-------|----------------|
| JOAO DA SILVA | 12345678901 | 001 | 1234-5 | 12345-6 | 1500,00 | 10/01/2025 |
| EMPRESA LTDA | 12345678000190 | 341 | 4321-0 | 98765-4 | 8750,50 | 10/01/2025 |

## Saida gerada

- O aplicativo solicita a pasta de saida.
- Para cada arquivo de entrada valido, gera um TXT no formato:

```text
CNAB_<nome_base>_<yyyymmdd_hhmmss>.txt
```

- Cada registro possui 240 caracteres.

## Erros comuns

| Erro | Causa comum | Acao recomendada |
|------|-------------|------------------|
| `config.json nao encontrado` | Arquivo ausente em `config/` | Criar/copiar `config/config.json` |
| `Colunas ausentes` | Cabecalhos diferentes do esperado | Revisar nomes de colunas da planilha |
| `Nenhuma linha com valor > 0` | Valores invalidos ou zerados | Corrigir coluna de valor no Excel |
| `Linha N chars != 240` | Inconsistencia de layout | Revisar dados de entrada e configuracao |
