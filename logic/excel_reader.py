"""
excel_reader.py
Lê dois formatos de Excel:
  data(5).xlsx                  → aba 'Export'   → pagamentos
  PROGRAMACAO_DE_APLICACAO.xlsx → qualquer aba   → aplicações por fundo
"""
import re, unicodedata, logging
import pandas as pd

logger = logging.getLogger(__name__)


# ── helpers ───────────────────────────────────────────────────────
def _only_numbers(v) -> str:
    return "".join(c for c in str(v or "") if c.isdigit())


def _split_agencia_bb(ag_raw) -> tuple:
    """
    Separa agência e DV para o BB.
    BB Segmento A: campo agência = 4 dígitos, DV agência = 1 dígito (campos separados).
    A planilha pode trazer o DV embutido no final do número de agência:
      - 5 dígitos → primeiros 4 = agência, último = DV
      - 4 dígitos → agência pura, DV = '0'
      - 6+ dígitos → últimos 5 divididos da mesma forma
    Retorna: (ag4: str, dv_ag: str)
    """
    s = str(ag_raw or "").strip()
    if not s:
        return "0000", "0"
    # aceita 'X' ou 'x' como DV final
    last = s[-1]
    has_x = True if last.upper() == "X" else False
    digits = "".join(c for c in s if c.isdigit())
    if has_x:
        if len(digits) == 5:
            return digits[:4], "X"
        elif len(digits) == 4:
            return digits, "X"
        elif len(digits) > 5:
            return digits[-5:-1], "X"
        else:
            return digits.zfill(4), "X"
    else:
        if len(digits) == 5:
            return digits[:4], digits[4]
        elif len(digits) == 4:
            return digits, "0"
        elif len(digits) > 5:
            return digits[-5:-1], digits[-1]
        else:
            return digits.zfill(4), "0"


def _split_conta_bb(ct_raw) -> tuple:
    """
    Separa conta e DV para o BB.
    BB Segmento A: campo conta = 12 dígitos, DV conta = 1 dígito (campos separados).
    Regra: o último dígito é SEMPRE o DV, independente do tamanho.
      - 2+ dígitos → tudo exceto o último = conta, último = DV
      - 1 dígito   → conta = '0' * 12, DV = esse dígito
      - vazio      → conta = '0' * 12, DV = '0'
    Retorna: (ct12: str zfilled 12, dv_ct: str)
    """
    s = str(ct_raw or "").strip()
    if not s:
        return "0" * 12, "0"
    last = s[-1]
    has_x = True if last.upper() == "X" else False
    digits = "".join(c for c in s if c.isdigit())
    if has_x:
        # quando o DV é 'X', os dígitos restantes formam a conta
        if len(digits) >= 1:
            return digits.zfill(12)[-12:], "X"
        else:
            return "0" * 12, "X"
    else:
        if len(digits) >= 2:
            return digits[:-1].zfill(12)[-12:], digits[-1]
        elif len(digits) == 1:
            return "0" * 12, digits[0]
        else:
            return "0" * 12, "0"


def _normalizar_codigo_barras(raw) -> str:
    """
    Corrige o artefato mais comum quando a célula de código de barras/linha
    digitável não está formatada como Texto no Excel: o valor é lido como
    float e vira notação científica (ex.: '1.234567890123e+46') ou ganha
    sufixo '.0'. Reconstrói a string numérica original nesses casos.
    Não tenta "adivinhar" zeros à esquerda perdidos — isso é ambíguo e é
    tratado como erro de validação mais à frente (ver boleto_generator).
    """
    s = str(raw or "").strip()
    if not s or s.lower() == "nan":
        return ""
    if re.search(r"e[+-]?\d+", s.lower()):
        try:
            s = format(int(float(s)), "f")
        except Exception:
            pass
    elif s.endswith(".0"):
        s = s[:-2]
    return s


def _parse_valor(v) -> float:
    raw = str(v or "").strip().replace("R$","").replace(" ","")
    if "," in raw and "." in raw:
        raw = raw.replace(".","").replace(",",".")
    elif "," in raw:
        raw = raw.replace(",",".")
    try:
        return abs(float(raw))
    except Exception:
        return 0.0


def _norm_col(col: str) -> str:
    s = str(col or "").strip().upper()
    s = " ".join(s.split())
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def _norm_name(name: str) -> str:
    s = str(name or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s)


# ── detecção de tipo ───────────────────────────────────────────────
_KW_FUNDOS = ["PGA","FCBE","HORIZONTE","HORIZONE","PROTEGIDO","PASSIVO"]
_KW_PAGTO  = ["BANCO","AGENCIA","CONTA","CNPJ","CPF"]


def _is_aplicacoes(cols: list) -> bool:
    j = " ".join(cols)
    if any(k in j for k in _KW_PAGTO):
        return False
    return any(k in j for k in _KW_FUNDOS)


# ── classificação de pagamentos ────────────────────────────────────
def _classify_pagamento(row) -> tuple:
    """
    Classifica uma linha em (tipo_pagamento, forma_lancamento).

    Fonte primária: coluna DOCUMENTO_CONTABIL (aliás 'doc_contabil'), que
    já indica explicitamente se o pagamento é BOLETO, TRANSFERÊNCIA ou TED.

    Regras:
      1. Se DOCUMENTO_CONTABIL contém 'BOLETO':
           código de barras começa com '8' (produto = arrecadação/convênio,
           ex.: Claro, Sabesp, Cemig, tributos) e tem 44 ou 48 dígitos
             → BOLETO_ARRECADACAO (forma 32 — usa Segmento O, não J!)
           banco do código de barras (3 primeiros dígitos) == '001'
             → BOLETO_BB      (forma 30)
           caso contrário
             → BOLETO_OUTROS  (forma 31)
      2. Se contém 'TRANSF' → TRANSFERENCIA_BB (forma 01)
      3. Se contém 'TED'    → TED_DOC          (forma 03)
      4. Fallback (coluna ausente/vazia/valor não reconhecido):
         mantém a regra antiga baseada em codigo_barras/banco, para
         garantir retrocompatibilidade com planilhas sem essa coluna.

    IMPORTANTE: contas de arrecadação (água, luz, telefone, tributos) usam
    um padrão de código de barras totalmente diferente do boleto bancário
    (produto '8' na 1ª posição, valor nas posições 5-15) e precisam do
    Segmento O do CNAB 240 — nunca do Segmento J. Ver TributoGenerator em
    boleto_generator.py.
    """
    doc_contabil = _norm_col(row.get("doc_contabil", ""))
    cod = str(row.get("codigo_barras", "")).strip()
    digits = _only_numbers(cod)
    banco_cod = digits[:3] if digits else ""
    is_arrecadacao = bool(digits) and digits[0] == "8" and len(digits) in (44, 48)

    if "BOLETO" in doc_contabil:
        if is_arrecadacao:
            return "BOLETO_ARRECADACAO", "32"
        if banco_cod == "001":
            return "BOLETO_BB", "30"
        return "BOLETO_OUTROS", "31"

    if "TRANSF" in doc_contabil:
        return "TRANSFERENCIA_BB", "01"

    if "TED" in doc_contabil:
        return "TED_DOC", "03"

    # ── Fallback: DOCUMENTO_CONTABIL ausente ou não reconhecido ──────
    if cod:
        if is_arrecadacao:
            return "BOLETO_ARRECADACAO", "32"
        if banco_cod == "001":
            return "BOLETO_BB", "30"
        return "BOLETO_OUTROS", "31"

    banco = _only_numbers(row.get("banco", ""))
    if banco == "001":
        return "TRANSFERENCIA_BB", "01"
    return "TED_DOC", "03"


def classificar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        df = df.copy()
        df["tipo_pagamento"] = []
        df["forma_lancamento"] = []
        return df

    tipos, formas = [], []
    for idx, row in df.iterrows():
        tipo, forma = _classify_pagamento(row)
        tipos.append(tipo)
        formas.append(forma)

        doc_contabil = str(row.get("doc_contabil", "")).strip()
        cod = str(row.get("codigo_barras", "")).strip()
        ref = doc_contabil or (_only_numbers(cod)[:3] if cod else _only_numbers(row.get("banco", "")))
        logger.info("Linha %s classificada como %s (%s)", idx + 1, tipo, ref or forma)

        # validação suave (não interrompe a execução)
        if tipo in ("BOLETO_BB", "BOLETO_OUTROS", "BOLETO_ARRECADACAO"):
            digits = _only_numbers(cod)
            if len(digits) not in (44, 47, 48):
                logger.warning(
                    "Linha %s: código de barras com tamanho suspeito (%s dígitos)",
                    idx + 1, len(digits)
                )
        else:
            if not (_only_numbers(row.get("banco", "")) and str(row.get("agencia", "")).strip()
                     and str(row.get("conta", "")).strip()):
                logger.warning(
                    "Linha %s: dados bancários incompletos para %s", idx + 1, tipo
                )

    df = df.copy()
    df["tipo_pagamento"] = tipos
    df["forma_lancamento"] = formas
    return df


def agrupar_pagamentos(df: pd.DataFrame) -> dict:
    """
    Agrupa um DataFrame de pagamentos por (tipo_pagamento, forma_lancamento).
    Classifica automaticamente se o DataFrame ainda não tiver sido classificado.
    Retorna: { (tipo_pagamento, forma_lancamento): DataFrame }
    """
    if df.empty:
        return {}

    if "tipo_pagamento" not in df.columns or "forma_lancamento" not in df.columns:
        df = classificar_dataframe(df)

    grupos = {}
    for chave, sub in df.groupby(["tipo_pagamento", "forma_lancamento"]):
        grupos[chave] = sub.reset_index(drop=True)
    return grupos


# ── leitura principal ─────────────────────────────────────────────
def read_excel(file_path: str, aba: str = None):
    """Retorna (DataFrame, lista_erros). aba: 'pagamentos'|'aplicacoes'|None"""
    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception:
        try:
            xls = pd.ExcelFile(file_path, engine="xlrd")
        except Exception as e:
            return pd.DataFrame(), [f"Não foi possível abrir: {e}"]

    sheet = _find_sheet(xls, aba)

    try:
        df_raw = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=None)
    except Exception as e:
        return pd.DataFrame(), [f"Erro ao ler aba '{sheet}': {e}"]

    df_raw = df_raw.dropna(how="all").reset_index(drop=True)
    if df_raw.empty:
        return pd.DataFrame(), ["Planilha vazia."]

    hdr = _find_header(df_raw, aba)
    df_raw.columns = [str(c).strip() for c in df_raw.iloc[hdr]]
    df = df_raw.iloc[hdr+1:].reset_index(drop=True).dropna(how="all").reset_index(drop=True)

    cols_norm = [_norm_col(c) for c in df.columns]

    if aba == "aplicacoes" or _is_aplicacoes(cols_norm):
        return _parse_aplicacoes(df)
    return _parse_pagamentos(df)


def _find_sheet(xls, aba):
    if aba:
        an = _norm_name(aba)
        for s in xls.sheet_names:
            if _norm_name(s) == an:
                return s
    for s in xls.sheet_names:
        try:
            probe = pd.read_excel(xls, sheet_name=s, nrows=6, dtype=str, header=None)
        except Exception:
            continue
        for ri in range(min(6, len(probe))):
            vals = [_norm_col(str(v)) for v in probe.iloc[ri] if pd.notna(v)]
            if aba == "aplicacoes" and _is_aplicacoes(vals):
                return s
            if aba == "pagamentos" and any(k in " ".join(vals) for k in _KW_PAGTO):
                return s
    return xls.sheet_names[0]


def _find_header(df, aba):
    for i in range(min(10, len(df))):
        vals = [_norm_col(str(v)) for v in df.iloc[i] if pd.notna(v) and str(v).strip()]
        if aba == "aplicacoes" and _is_aplicacoes(vals):
            return i
        if aba != "aplicacoes" and any(k in " ".join(vals) for k in _KW_PAGTO):
            return i
        if len(vals) >= 3:
            return i
    return 0


# ── parser: pagamentos ─────────────────────────────────────────────
_ALIAS = {
    "CONTRAPARTE":"nome","NOME":"nome",
    "CNPJ/CPF":"cpf_cnpj","CPF/CNPJ":"cpf_cnpj","CNPJ":"cpf_cnpj","DOCUMENTO":"cpf_cnpj",
    "BANCO":"banco",
    "AGENCIA":"agencia",
    "AF":"identificador","NUMERO":"numero",
    "CONTA":"conta",
    "VALOR LIQUIDO":"valor","VALOR":"valor",
    "DATA":"data_pagamento",
    "CODIGO DE BARRAS":"codigo_barras",
    "DOCUMENTO CONTABIL":"doc_contabil",    
}


def _parse_pagamentos(df_raw):
    rename = {c: _ALIAS[_norm_col(c)] for c in df_raw.columns if _norm_col(c) in _ALIAS}
    df = df_raw.rename(columns=rename)
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    if "codigo_barras" in df.columns:
        df["codigo_barras"] = df["codigo_barras"].apply(_normalizar_codigo_barras)
    required = ["nome","cpf_cnpj","banco","agencia","conta","valor"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        return pd.DataFrame(), [f"Colunas ausentes: {missing}. Encontradas: {list(df.columns)}"]
    df = df[df["valor"].apply(lambda v: _parse_valor(v) > 0)].reset_index(drop=True)
    if df.empty:
        return pd.DataFrame(), ["Nenhuma linha com valor > 0."]

    df = classificar_dataframe(df)
    return df, []


# ── parser: aplicações ─────────────────────────────────────────────
_FUNDO_MAP = {
    "PGA":"PGA","FCBE":"FCBE",
    "PASSIVO 40":"PASSIVO 2040","PASSIVO 2040":"PASSIVO 2040",
    "PASSIVO 50":"PASSIVO 2050","PASSIVO 2050":"PASSIVO 2050",
    "HORIZONTE 40":"HORIZONE 2040","HORIZONTE 2040":"HORIZONE 2040",
    "HORIZONTE 50":"HORIZONE 2050","HORIZONTE 2050":"HORIZONE 2050",
    "PROTEGIDO":"HORIZONE PROTEGIDO","HORIZONTE PROTEGIDO":"HORIZONE PROTEGIDO",
}


def _match_fundo(col_norm):
    for pat, key in _FUNDO_MAP.items():
        if pat in col_norm:
            return key
    return None


def _parse_aplicacoes(df_raw):
    errors, rows = [], []
    # linha TOTAL
    total_row = None
    for _, row in df_raw.iterrows():
        if "TOTAL" in " ".join(str(x).upper() for x in row.values):
            total_row = row; break
    if total_row is None:
        errors.append("Linha TOTAL não encontrada — usando última linha.")
        total_row = df_raw.iloc[-1]

    for col in df_raw.columns:
        fundo = _match_fundo(_norm_col(col))
        if not fundo:
            continue
        raw_val = total_row[col]
        if pd.isna(raw_val) or str(raw_val).strip() in ("","nan"):
            continue
        valor = _parse_valor(raw_val)
        if valor > 0:
            rows.append({"fundo": fundo, "valor": valor})

    if not rows:
        return pd.DataFrame(), ["Nenhum fundo com valor encontrado."]
    return pd.DataFrame(rows), errors


# ── ExcelReader (compatível com app.py) ───────────────────────────
class ExcelReader:
    def __init__(self, config: dict):
        self.config = config

    def read_pagamentos(self, file_path: str):
        df, _ = read_excel(file_path, aba="pagamentos")
        result = []
        for _, r in df.iterrows():
            banco = _only_numbers(r.get("banco","")).zfill(3)[:3]
            ag4, dv_ag = _split_agencia_bb(r.get("agencia",""))
            ct12, dv_ct = _split_conta_bb(r.get("conta",""))
            doc   = _only_numbers(r.get("cpf_cnpj",""))
            nome  = str(r.get("nome","")).strip()
            valor = _parse_valor(r.get("valor",""))
            if not (banco and nome and doc and valor > 0): continue
            result.append({
                "nome":nome,"documento":doc,"banco":banco,
                "agencia5":ag4.zfill(5),"dv_agencia":dv_ag,
                "conta12":ct12,"dv_conta":dv_ct,"dv_ag_conta":" ",
                "valor":valor,
                "seu_numero":str(r.get("identificador","")).strip(),
                "data_pagamento":str(r.get("data_pagamento","")).strip(),
            })
        return result

    def read_aplicacoes(self, file_path: str):
        df, _ = read_excel(file_path, aba="aplicacoes")
        ap  = self.config.get("aplicacoes",{})
        cts = ap.get("contas",{})
        banco = str(ap.get("banco","208")).zfill(3)
        ag5   = str(ap.get("agencia","00001")).zfill(5)[:5]
        dv_ag = str(ap.get("dv_agencia"," "))[:1]
        result = []
        for _, r in df.iterrows():
            fundo = r["fundo"]; valor = _parse_valor(r["valor"])
            cc = cts.get(fundo)
            if not cc: raise ValueError(f"Fundo '{fundo}' não encontrado no config.json")
            result.append({
                "nome":str(cc.get("nome",fundo))[:30].ljust(30),
                "documento":_only_numbers(cc.get("cnpj","")),
                "banco":banco,"agencia5":ag5,"dv_agencia":dv_ag,
                "conta12":str(cc.get("conta","")).zfill(12)[:12],
                "dv_conta":str(cc.get("dv_conta"," "))[:1],"dv_ag_conta":" ",
                "valor":valor,"seu_numero":str(cc.get("finalidade",fundo)),
            })
        return result