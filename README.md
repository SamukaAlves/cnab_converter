# Conversor Excel -> CNAB 240 (Banco do Brasil)
<img width="892" height="944" alt="image" src="https://github.com/user-attachments/assets/f2fa1abf-be11-4230-9860-ae19cb3db31c" />


Aplicacao desktop em Python para converter planilhas Excel em arquivos de remessa CNAB 240, com suporte a perfis de pagador, leitura automatica de layouts e geracao de TXT validado.

## Funcionalidades

- Converte planilhas de pagamento em arquivos CNAB 240 para o Banco do Brasil.
- Gera registros dos segmentos A e B, com validacao de tamanho fixo em 240 caracteres.
- Suporta dois perfis de pagador no `config/config.json`, como PGA e BBJUSMP.
- Lê planilhas de pagamentos e de aplicacoes, identificando colunas mesmo com aliases diferentes.
- Normaliza agencia, conta, documento e valor antes da geracao do arquivo.
- Permite informar data de pagamento para compor o header do lote e os segmentos.
- Exporta o resultado em arquivo `.txt` com nome padronizado por data e hora.
- Oferece interface grafica em PyQt5 para uso direto no desktop.

## O que o sistema faz

1. Carrega a configuracao em `config/config.json`.
2. Identifica se a planilha enviada e de pagamentos ou aplicacoes.
3. Processa os dados, ajusta campos bancarios e monta o layout CNAB.
4. Gera o arquivo final para envio ao Banco do Brasil.

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
