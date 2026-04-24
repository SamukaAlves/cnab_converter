"""
cnab_generator.py
Gera CNAB 240 BB — layout verificado contra o validador oficial de leiautes.
"""
from datetime import datetime


# ── formatação ────────────────────────────────────────────────────
def _fn(v, n): return str(v or 0).zfill(n)[:n]
def _fa(v, n): return str(v or "").upper().ljust(n)[:n]
def _vc(f):    return str(int(round(float(f or 0) * 100))).zfill(15)[:15]

def _pv(v) -> float:
    """
    Suporta:
      - Formato US:  39378.98  ou  -39378.98
      - Formato BR:  39.378,98 ou  -39.378,98
      - Inteiro:     37540
    """
    raw = str(v or "").strip().replace("R$", "").replace(" ", "")
    if not raw or raw == "-":
        return 0.0
    # Formato BR: tem vírgula E ponto → ponto é separador de milhar, vírgula é decimal
    if "," in raw and "." in raw:
        raw = raw.replace(".", "").replace(",", ".")
    # Formato BR simples: só vírgula → vírgula é decimal
    elif "," in raw:
        raw = raw.replace(",", ".")
    # Formato US: só ponto → ponto já é decimal, não mexer pelo amor de Jesus
    try:
        return abs(float(raw))
    except Exception:
        return 0.0

def _nums(v): return "".join(c for c in str(v or "") if c.isdigit())

def _date(raw, now):
    s = str(raw or "").strip()
    if len(s) == 10 and s[2] == "/" and s[5] == "/": return s[:2] + s[3:5] + s[6:]
    if len(s) == 8 and s.isdigit(): return s
    return now.strftime("%d%m%Y")

def _chk(l, ctx=""):
    if len(l) != 240:
        raise ValueError(f"Linha {len(l)} chars ≠ 240{' ['+ctx+']' if ctx else ''}")


# ── gerador principal ─────────────────────────────────────────────
class CNAB240Generator:

    def __init__(self, cfg):
        self.cfg   = cfg
        self.p     = cfg.pagador
        self.lt    = cfg.lote
        self.banco = cfg.codigo_banco   # '001'

    def gerar(self, pagamentos: list) -> str:
        now    = datetime.now()
        linhas = []

        linhas.append(self._h_arq(now))
        linhas.append(self._h_lote())

        total = 0.0
        for i, pg in enumerate(pagamentos, 1):
            total += _pv(pg.get("valor", 0))
            linhas.append(self._seg_a(pg, i * 2 - 1, now))
            linhas.append(self._seg_b(pg, i * 2))

        qtd_lote  = 1 + len(pagamentos) * 2 + 1   # hdr_lote + segs + trl_lote
        qtd_total = 1 + qtd_lote + 1               # hdr_arq + lote + trl_arq

        linhas.append(self._t_lote(qtd_lote, total))
        linhas.append(self._t_arq(qtd_total))

        for i, l in enumerate(linhas): _chk(l, f"linha {i}")

        # Terminador \n (LF simples) — compatível com validador BB
        return "\n".join(linhas) + "\n"

    # ── Header Arquivo ────────────────────────────────────────────
    # [000:171] campos fixos (171 chars)
    # [171:240] = 20sp + 20sp + 13sp + 16zeros = 69 chars → [230:240]='0000000000'
    def _h_arq(self, now):
        p = self.p
        return (
            self.banco + "0000" + "0" + " "*9 + "2" +                    # 000-017 (18)
            p["cnpj"] + p["convenio"] + p["codigo_bb2"] + " "*7 +        # 018-051 (34)
            p["agencia"] + "0" + p["dv_agencia"] +                       # 052-057 (6)
            p["conta"] + p["dv_conta"] + p["dv_ag_conta"] +              # 058-071 (14)
            _fa(p["nome_empresa"], 30) + _fa("BANCO DO BRASIL", 30) +    # 072-131 (60)
            " "*10 + "1" +                                                # 132-142 (11)
            now.strftime("%d%m%Y") + now.strftime("%H%M%S") + _fn(1, 6) + # 143-162 (20)
            p["versao_layout_arquivo"] + "00000" +                        # 163-170 (8)
            " "*20 + " "*20 + " "*13 + "0"*16                            # 171-239 (69): [230:240]='0000000000'
        )

    # ── Header Lote ───────────────────────────────────────────────
    # [000:172] campos fixos (172 chars)
    # [172:177] = "00000"
    # [177:240] = 53sp + "0000000000" → [230:240]='0000000000'
    def _h_lote(self):
        p, lt = self.p, self.lt
        return (
            self.banco + "0001" + "1" + "C" +                             # 000-007 (8)
            _fn(lt["tipo_servico"], 2) + _fn(lt["forma_lancamento"], 2) + # 009-012 (4)
            p["versao_layout_lote"] + " " + "2" +                         # 013-017 (5)
            p["cnpj"] + p["convenio"] + p["codigo_bb2"] + " "*7 +         # 018-051 (34)
            p["agencia"] + "0" + p["dv_agencia"] +                        # 052-057 (6)
            p["conta"] + p["dv_conta"] + p["dv_ag_conta"] +               # 058-071 (14)
            _fa(p["nome_empresa"], 30) +                                  # 072-101 (30)
            " "*40 + " "*30 +                                             # 102-171 (70)
            "00000" +                                                     # 172-176 (5)
            " "*53 + "0000000000"                                         # 177-239 (63): [230:240]='0000000000'
        )

    # ── Segmento A ────────────────────────────────────────────────
    def _seg_a(self, pg, seq, now):
        cam   = str(self.lt.get("camara", "018")).zfill(3)[:3]
        banco = str(pg.get("banco", "000")).zfill(3)[:3]
        ag5   = str(pg.get("agencia5", "00000")).zfill(5)[:5]
        dv_ag = str(pg.get("dv_agencia", " "))[:1]
        ct12  = str(pg.get("conta12", "0"*12)).zfill(12)[:12]
        dv_c  = str(pg.get("dv_conta", " "))[:1]
        dv_ac = str(pg.get("dv_ag_conta", " "))[:1]
        nome  = _fa(pg.get("nome", ""), 30)
        snm   = _fa(pg.get("seu_numero", ""), 20)
        data  = _date(pg.get("data_pagamento", ""), now)
        val   = _pv(pg.get("valor", 0))
        return (
            self.banco + "0001" + "3" + _fn(seq, 5) + "A" + "0" + "00" +
            cam + banco + ag5 + dv_ag + ct12 + dv_c + dv_ac + nome + snm +
            data + "BRL" + "0"*15 + _vc(val) +
            " "*20 + "00000000" + "0"*15 +
            " "*20 + " "*5 + " "*8 + " "*19 + "0"*11
        )

    # ── Segmento B ────────────────────────────────────────────────
    # Template exato extraído dos refs:
    # [032:062] 30sp | [062:067] '00000' | [067:117] 50sp
    # [117:122] '00000' | [122:127] 5sp | [127:210] 83zeros
    # [210:225] 15sp | [225:232] '0000000' | [232:240] 8sp
    def _seg_b(self, pg, seq):
        doc = _nums(pg.get("documento", ""))
        ti  = "1" if len(doc) == 11 else "2"
        doc = doc.zfill(14)[:14]
        return (
            self.banco + "0001" + "3" + _fn(seq, 5) + "B" + " " + "  " +
            ti + doc +
            " "*30 + "00000" + " "*50 + "00000" + " "*5 +
            "0"*83 + " "*15 + "0000000" + " "*8
        )

    # ── Trailer Lote ──────────────────────────────────────────────
    def _t_lote(self, qtd, total):
        soma = str(int(round(total * 100))).zfill(18)[:18]
        return (
            self.banco + "0001" + "5" + " "*9 +
            _fn(qtd, 6) + soma + "0"*18 + "0"*6 +
            " "*165 + "0"*10
        )

    # ── Trailer Arquivo ───────────────────────────────────────────
    def _t_arq(self, qtd):
        return (
            self.banco + "9999" + "9" + " "*9 +
            "000001" + _fn(qtd, 6) + "000000" + " "*205
        )


#  CNABGenerator (compatível com app.py) 
class CNABGenerator:
    def __init__(self, cfg):
        self.cfg = cfg

    def generate(self, df, tipo_pagamento="PGA", num_lote=1):
        gen  = CNAB240Generator(self.cfg)
        tipo = str(tipo_pagamento).upper()
        pags = self._apl(df) if tipo in ("APLICACAO", "APLICACOES") else self._pag(df)
        return gen.gerar(pags)

    def _pag(self, df):
        res = []
        for _, r in df.iterrows():
            banco = _nums(r.get("banco", "")).zfill(3)[:3]
            ag5   = _nums(r.get("agencia", "")).zfill(5)[-5:]
            ct12  = _nums(r.get("conta", "")).zfill(12)[-12:]
            doc   = _nums(r.get("cpf_cnpj", "")).zfill(14)
            nome  = str(r.get("nome", "")).strip()
            val   = _pv(r.get("valor", 0))
            if not (banco and nome and doc and val > 0): continue
            res.append({
                "nome": nome, "documento": doc, "banco": banco,
                "agencia5": ag5, "dv_agencia": " ", "conta12": ct12,
                "dv_conta": " ", "dv_ag_conta": " ", "valor": val,
                "seu_numero": str(r.get("identificador", "")).strip(),
                "data_pagamento": str(r.get("data_pagamento", "")).strip(),
            })
        return res

    def _apl(self, df):
        ap  = self.cfg.aplicacoes_config
        cts = ap.get("contas", {})
        banco = str(ap.get("banco", "208")).zfill(3)
        ag5   = str(ap.get("agencia", "00001")).zfill(5)[:5]
        dv_ag = str(ap.get("dv_agencia", " "))[:1]
        res   = []
        for _, r in df.iterrows():
            fundo = str(r.get("fundo", "")).strip()
            val   = _pv(r.get("valor", 0))
            if val <= 0: continue
            cc = cts.get(fundo)
            if not cc: raise ValueError(f"Fundo '{fundo}' não encontrado no config.json")
            res.append({
                "nome":        str(cc.get("nome", fundo))[:30].ljust(30),
                "documento":   _nums(cc.get("cnpj", "")).zfill(14),
                "banco":       banco, "agencia5": ag5, "dv_agencia": dv_ag,
                "conta12":     str(cc.get("conta", "")).zfill(12)[:12],
                "dv_conta":    str(cc.get("dv_conta", " "))[:1], "dv_ag_conta": " ",
                "valor":       val,
                "seu_numero":  str(cc.get("finalidade", fundo)),
            })
        return res
