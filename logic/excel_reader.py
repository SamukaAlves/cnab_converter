"""
excel_reader.py
Lê dois formatos de Excel:
  data(5).xlsx                  → aba 'Export'   → pagamentos
  PROGRAMACAO_DE_APLICACAO.xlsx → qualquer aba   → aplicações por fundo
"""
import re, unicodedata
import pandas as pd


# ── helpers ───────────────────────────────────────────────────────
def _only_numbers(v) -> str:
    return "".join(c for c in str(v or "") if c.isdigit())


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
_KW_FUNDOS = ["PGA","FCBE","HORIZONTE","HORIZONE","PROTEGIDO"]
_KW_PAGTO  = ["BANCO","AGENCIA","CONTA","CNPJ","CPF"]


def _is_aplicacoes(cols: list) -> bool:
    j = " ".join(cols)
    if any(k in j for k in _KW_PAGTO):
        return False
    return any(k in j for k in _KW_FUNDOS)


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
    "BANCO":"banco","AGENCIA":"agencia",
    "AF":"identificador","NUMERO":"numero",
    "CONTA":"conta",
    "VALOR LIQUIDO":"valor","VALOR":"valor",
    "DATA":"data_pagamento",
    "CODIGO DE BARRAS":"codigo_barras","DOCUMENTO CONTABIL":"doc_contabil",
}


def _parse_pagamentos(df_raw):
    rename = {c: _ALIAS[_norm_col(c)] for c in df_raw.columns if _norm_col(c) in _ALIAS}
    df = df_raw.rename(columns=rename)
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    required = ["nome","cpf_cnpj","banco","agencia","conta","valor"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        return pd.DataFrame(), [f"Colunas ausentes: {missing}. Encontradas: {list(df.columns)}"]
    df = df[df["valor"].apply(lambda v: _parse_valor(v) > 0)].reset_index(drop=True)
    if df.empty:
        return pd.DataFrame(), ["Nenhuma linha com valor > 0."]
    return df, []


# ── parser: aplicações ─────────────────────────────────────────────
_FUNDO_MAP = {
    "PGA":"PGA","FCBE":"FCBE",
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
            ag5   = _only_numbers(r.get("agencia","")).zfill(5)[-5:]
            ct12  = _only_numbers(r.get("conta","")).zfill(12)[-12:]
            doc   = _only_numbers(r.get("cpf_cnpj","")).zfill(14)
            nome  = str(r.get("nome","")).strip()
            valor = _parse_valor(r.get("valor",""))
            if not (banco and nome and doc and valor > 0): continue
            result.append({
                "nome":nome,"documento":doc,"banco":banco,
                "agencia5":ag5,"dv_agencia":" ","conta12":ct12,
                "dv_conta":" ","dv_ag_conta":" ","valor":valor,
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
                "documento":_only_numbers(cc.get("cnpj","")).zfill(14),
                "banco":banco,"agencia5":ag5,"dv_agencia":dv_ag,
                "conta12":str(cc.get("conta","")).zfill(12)[:12],
                "dv_conta":str(cc.get("dv_conta"," "))[:1],"dv_ag_conta":" ",
                "valor":valor,"seu_numero":str(cc.get("finalidade",fundo)),
            })
        return result
