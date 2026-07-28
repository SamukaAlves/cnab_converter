"""
boleto_generator.py
Gera segmentos CNAB 240 BB para pagamento de boletos (Segmento J + J-52).
Reutiliza Header/Trailer de Arquivo do CNAB240Generator existente.
Autor: Samuel Alves
"""
from datetime import datetime

from logic.cnab_generator import (
    CNAB240Generator, _fn, _fa, _vc, _pv, _nums, _date, _chk, _tipo_inscricao,
)


# ── Código de barras / Linha digitável ────────────────────────────

def _limpar_digitavel(raw: str) -> str:
    """Remove pontos, espaços e hífens de uma linha digitável."""
    return "".join(c for c in str(raw or "") if c.isdigit())


def linha_digitavel_to_barcode(ld: str) -> str:
    """
    Converte linha digitável (47 dígitos) para código de barras (44 dígitos).

    Linha digitável (47):
      Campo1(10) Campo2(11) Campo3(11) Campo4(1) Campo5(14)
      AAABC.CCCCX DDDDD.DDDDDY EEEEE.EEEEEZ K UUUUVVVVVVVVVV

    Código de barras (44):
      AAAB K UUUUVVVVVVVVVV CCCCC DDDDDDDDD EEEEEEEEE

    Onde X, Y, Z são DVs dos campos (descartados) e K é o DV geral.
    """
    d = _limpar_digitavel(ld)
    if len(d) not in (47, 48):
        raise ValueError(f"Linha digitável deve ter 47 dígitos, recebeu {len(d)}")

    # Normalizar para 47 se veio 48 (alguns sistemas adicionam dígito extra)
    if len(d) == 48:
        d = d[:47]

    # Campos da linha digitável (posições 0-based)
    campo1 = d[0:10]     # AAABCCCCC + X (DV)
    campo2 = d[10:21]    # DDDDDDDDDD + Y (DV)
    campo3 = d[21:32]    # EEEEEEEEEE + Z (DV)
    campo4 = d[32:33]    # K (DV geral do barcode)
    campo5 = d[33:47]    # UUUUVVVVVVVVVV (fator venc + valor)

    # Montar código de barras (44 dígitos):
    # Pos 1-3:   Banco (AAA) = campo1[0:3]
    # Pos 4:     Moeda (B)   = campo1[3]
    # Pos 5:     DV geral (K)= campo4
    # Pos 6-19:  Fator+Valor = campo5
    # Pos 20-24: campo1[4:9] (CCCCC, sem o DV X)
    # Pos 25-34: campo2[0:10] (DDDDDDDDDD, sem o DV Y)
    # Pos 35-44: campo3[0:10] (EEEEEEEEEE, sem o DV Z)
    barcode = (
        campo1[0:4] +      # AAAB (banco + moeda)
        campo4 +            # K (DV geral)
        campo5 +            # fator vencimento + valor
        campo1[4:9] +       # parte livre campo1 (sem DV)
        campo2[0:10] +      # parte livre campo2 (sem DV)
        campo3[0:10]         # parte livre campo3 (sem DV)
    )

    if len(barcode) != 44:
        raise ValueError(f"Erro na conversão: barcode resultante com {len(barcode)} dígitos")

    return barcode


def _mod11_barcode(barcode: str) -> int:
    """
    Calcula o DV do código de barras (módulo 11, pesos 2-9).
    Usado para validar a posição 5 do código de barras.
    """
    # Remove a posição 5 (DV) para calcular — posições 0-based: índice 4
    sem_dv = barcode[:4] + barcode[5:]
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    soma = 0
    for i, c in enumerate(reversed(sem_dv)):
        soma += int(c) * pesos[i % 8]
    resto = soma % 11
    dv = 11 - resto
    if dv in (0, 10, 11):
        return 1
    return dv


def validar_codigo_barras(cod_raw: str) -> str:
    """
    Valida e retorna código de barras com 44 dígitos.
    Aceita:
      - 44 dígitos (código de barras direto)
      - 47/48 dígitos (linha digitável → converte)
    Valida DV (posição 5, módulo 11).
    Retorna o código de barras limpo (44 dígitos).
    Lança ValueError se inválido.
    """
    d = _limpar_digitavel(cod_raw)

    if not d:
        raise ValueError("Código de barras vazio")

    if len(d) in (47, 48):
        d = linha_digitavel_to_barcode(d)
    elif len(d) != 44:
        raise ValueError(
            f"Código de barras deve ter 44 dígitos (ou linha digitável 47/48). "
            f"Recebido: {len(d)} dígitos"
        )

    # Validar DV (posição 5, índice 4)
    dv_informado = int(d[4])
    dv_calculado = _mod11_barcode(d)
    if dv_informado != dv_calculado:
        raise ValueError(
            f"DV do código de barras inválido: esperado {dv_calculado}, "
            f"encontrado {dv_informado}"
        )

    return d


# ── Gerador de boletos ────────────────────────────────────────────

def valor_do_codigo_barras(barcode: str) -> float:
    """
    Decodifica o valor do título embutido nas posições 10-19 (índice 9:19)
    do código de barras (44 dígitos), em reais. Retorna 0.0 quando o campo
    vem zerado (boleto "sem valor fixo" / valor em aberto).
    """
    campo = barcode[9:19]
    if not campo.isdigit():
        return 0.0
    return int(campo) / 100.0


class BoletoGenerator:
    """
    Gera segmentos de lote CNAB 240 BB para pagamento de boletos.
    Segmentos: Header Lote J, Segmento J, Segmento J-52, Trailer Lote.
    Header e Trailer de Arquivo são gerados pelo CNAB240Generator (reutilização).
    """

    def __init__(self, cfg):
        self.cfg   = cfg
        self.p     = cfg.pagador
        self.lt    = cfg.lote
        self.banco = cfg.codigo_banco   # '001'

    def gerar(self, pagamentos: list, payment_date=None) -> str:
        now    = datetime.now()
        linhas = []

        # Normaliza payment_date (mesma lógica do CNAB240Generator)
        pd_dt = None
        if payment_date:
            if isinstance(payment_date, datetime):
                pd_dt = payment_date
            else:
                s = str(payment_date).strip()
                try:
                    if len(s) == 8 and s.isdigit():
                        pd_dt = datetime.strptime(s, "%d%m%Y")
                    elif "/" in s:
                        pd_dt = datetime.strptime(s, "%d/%m/%Y")
                    else:
                        pd_dt = datetime.fromisoformat(s)
                except Exception:
                    pd_dt = None

        # Header Arquivo — reutiliza CNAB240Generator
        base_gen = CNAB240Generator(self.cfg)
        linhas.append(base_gen._h_arq(now))

        # Header Lote para boletos
        linhas.append(self._h_lote_j(pd_dt or now))

        total = 0.0
        seq = 0
        self.erros = []  # boletos descartados por código de barras inválido
        for pg in pagamentos:
            # Validar código de barras — um boleto malformado (ex.: código
            # que veio como número do Excel e perdeu zeros à esquerda, ou
            # em notação científica) não deve abortar o lote inteiro.
            cod_raw = str(pg.get("codigo_barras", "")).strip()
            try:
                barcode = validar_codigo_barras(cod_raw)
            except ValueError as e:
                self.erros.append(
                    f"{pg.get('nome', '?')} (doc. {pg.get('seu_numero', '?')}): {e}"
                )
                continue
            pg["_barcode_44"] = barcode

            # O layout do BB valida que "valor do título"/"valor do
            # pagamento" (Segmento J) sejam coerentes com o valor
            # codificado no próprio código de barras. Para boletos com
            # valor fixo (campo não-zerado), esse valor é a fonte da
            # verdade — a planilha pode estar desatualizada ou trazer
            # um valor de outro boleto por engano.
            valor_bc = valor_do_codigo_barras(barcode)
            valor_excel = _pv(pg.get("valor", 0))
            if valor_bc > 0:
                if abs(valor_bc - valor_excel) > 0.01:
                    self.erros.append(
                        f"{pg.get('nome', '?')} (doc. {pg.get('seu_numero', '?')}): "
                        f"valor da planilha (R$ {valor_excel:.2f}) difere do valor "
                        f"codificado no código de barras (R$ {valor_bc:.2f}); "
                        f"usando o valor do código de barras."
                    )
                pg["valor"] = valor_bc
            # se valor_bc == 0 (boleto em aberto), mantém o valor da planilha

            val = _pv(pg.get("valor", 0))
            total += val
            seq += 1

            linhas.append(self._seg_j(pg, seq * 2 - 1, now, pd_dt))
            linhas.append(self._seg_j52(pg, seq * 2))

        if seq == 0:
            raise ValueError(
                "Nenhum boleto com código de barras válido neste lote:\n"
                + "\n".join(self.erros)
            )

        qtd_lote  = 1 + seq * 2 + 1   # hdr_lote + segs + trl_lote
        qtd_total = 1 + qtd_lote + 1  # hdr_arq + lote + trl_arq

        linhas.append(self._t_lote(qtd_lote, total))
        linhas.append(base_gen._t_arq(qtd_total))

        for i, l in enumerate(linhas):
            _chk(l, f"linha {i}")

        return "\n".join(linhas) + "\n"

    # ── Header Lote J ─────────────────────────────────────────────
    def _h_lote_j(self, data_pagamento: datetime = None):
        p, lt = self.p, self.lt
        dt_str = (data_pagamento or datetime.now()).strftime("%d%m%Y")
        return (
            self.banco + "0001" + "1" + "C" +                             # 000-007 (8)
            _fn(lt["tipo_servico"], 2) + _fn(lt["forma_lancamento"], 2) + # 008-011 (4)
            p["versao_layout_lote"] + " " + "2" +                         # 012-017 (6)
            p["cnpj"] + p["convenio"] + p["codigo_bb2"] + " "*7 +         # 018-051 (34)
            p["agencia"] + "0" + p["dv_agencia"] +                        # 052-057 (6)
            p["conta"] + p["dv_conta"] + p["dv_ag_conta"] +               # 058-071 (14)
            _fa(p["nome_empresa"], 30) +                                  # 072-101 (30)
            " "*40 + " "*30 +                                             # 102-171 (70)
            "00000" +                                                     # 172-176 (5)
            dt_str +                                                      # 177-184 (8) data pagamento
            "00000000" + "0"*8 +                                          # 185-200 (16) zeros
            " "*29 + "0000000000"                                         # 201-239 (39)
        )

    # ── Segmento J ────────────────────────────────────────────────
    def _seg_j(self, pg, seq, now, payment_now=None):
        barcode = pg.get("_barcode_44", "0" * 44)
        nome    = _fa(pg.get("nome", ""), 30)
        val     = _pv(pg.get("valor", 0))
        data_pg = _date(pg.get("data_pagamento", ""), payment_now or now)
        data_vc = _date(pg.get("data_vencimento", ""), payment_now or now)
        snm     = _fa(pg.get("seu_numero", ""), 20)

        return (
            self.banco + "0001" + "3" + _fn(seq, 5) + "J" +     # 001-014 (14)
            "0" + "00" +                                          # 015-017 (3) tipo mov + cód instrução
            barcode +                                             # 018-061 (44) código de barras
            nome +                                                # 062-091 (30) nome beneficiário
            data_vc +                                             # 092-099 (8) data vencimento
            _vc(val) +                                            # 100-114 (15) valor título
            "0"*15 +                                              # 115-129 (15) desconto/abatimento
            "0"*15 +                                              # 130-144 (15) acréscimos/multa
            data_pg +                                             # 145-152 (8) data pagamento
            _vc(val) +                                            # 153-167 (15) valor pagamento
            "0"*15 +                                              # 168-182 (15) qtde moeda
            snm +                                                 # 183-202 (20) nº documento empresa
            " "*20 +                                              # 203-222 (20) nosso número
            "09" +                                                # 223-224 (2) código moeda (Real)
            " "*6 +                                               # 225-230 (6) brancos
            " "*10                                                # 231-240 (10) ocorrências
        )

    # ── Segmento J-52 ─────────────────────────────────────────────
    def _seg_j52(self, pg, seq):
        p = self.p

        # Dados do pagador (empresa — vêm do config)
        ti_pag  = "2"                                     # CNPJ da empresa
        doc_pag = p["cnpj"].zfill(15)[:15]                # 15 chars

        # Dados do beneficiário (favorecido — vêm do pagamento)
        doc_ben = _nums(pg.get("documento", ""))
        ti_ben  = _tipo_inscricao(doc_ben)
        doc_ben = doc_ben.zfill(15)[:15]

        return (
            self.banco + "0001" + "3" + _fn(seq, 5) + "J" +       # 001-014 (14)
            "0" + "00" +                                          # 015-017 (3) tipo mov + cód instrução
            "52" +                                                # 018-019 (2) código registro opcional
            ti_pag + doc_pag +                                    # 020-035 (16) inscrição pagador
            _fa(p["nome_empresa"], 40) +                          # 036-075 (40) nome pagador
            ti_ben + doc_ben +                                    # 076-091 (16) inscrição beneficiário
            _fa(pg.get("nome", ""), 40) +                         # 092-131 (40) nome beneficiário
            "0" + "0"*15 +                                        # 132-147 (16) sacador/avalista (não aplicável)
            " "*40 +                                              # 148-187 (40) nome sacador
            " "*53                                                # 188-240 (53) brancos
        )

    # ── Trailer Lote ──────────────────────────────────────────────
    def _t_lote(self, qtd, total):
        soma = str(int(round(total * 100))).zfill(18)[:18]
        return (
            self.banco + "0001" + "5" + " "*9 +
            _fn(qtd, 6) + soma + "0"*18 + "0"*6 +
            " "*165 + "0"*10
        )


# ── Arrecadação / Convênio (contas de consumo, tributos) ──────────
#
# ATENÇÃO: este é um padrão TOTALMENTE DIFERENTE do boleto bancário acima.
# Contas como Claro, Sabesp, Cemig, IPTU, DARF etc. usam código de barras
# de "arrecadação" (1ª posição = '8'), cuja estrutura e regras de DV não
# têm nada a ver com o boleto bbancário tradicional (1ª posição = código
# do banco, ex. '001', '341'). Usar o algoritmo de boleto bancário nesses
# códigos produz um valor decodificado errado (bug identificado em
# produção: código da Claro decodificado como R$ 2.607.252,01 em vez do
# valor real, e o segundo boleto era descartado por falha de DV).
#
# Estrutura do código de barras de arrecadação (44 dígitos):
#   posição 1     : '8' (produto = arrecadação)
#   posição 2     : segmento (2=água/esgoto, 3=energia/gás, 4=telecom,
#                    5=órgãos governamentais, 6=carnês, 7=multas trânsito,
#                    9=uso exclusivo do banco)
#   posição 3     : identificador do valor efetivo/referência
#                    (6 ou 7 → DV geral em módulo 10; 8 ou 9 → módulo 11)
#   posição 4     : DV geral do código de barras
#   posições 5-15 : valor do documento (11 dígitos, 2 casas decimais)
#   posições 16-44: identificação da empresa/convênio + campo livre
#
# Linha digitável de arrecadação (48 dígitos) = 4 blocos de 12 dígitos
# (11 dígitos + 1 DV de bloco cada), SEM o reagrupamento usado no boleto
# bancário. Para obter o código de barras (44), basta remover o DV de
# cada um dos 4 blocos.

def _mod10_arrecadacao(digits: str) -> int:
    """DV módulo 10 (febraban, contas de arrecadação/convênio)."""
    total = 0
    pesos = [2, 1]
    for i, c in enumerate(reversed(digits)):
        p = pesos[i % 2]
        prod = int(c) * p
        if prod > 9:
            prod -= 9
        total += prod
    resto = total % 10
    return 0 if resto == 0 else 10 - resto


def _mod11_arrecadacao(digits: str) -> int:
    """
    DV módulo 11 para contas de arrecadação/convênio (identificador de
    valor 8 ou 9). Regra FEBRABAN específica deste produto — diferente
    do módulo 11 do boleto bancário: quando o resto da divisão por 11
    for 0 ou 1, o DV é 0 (não há substituição por 1).
    """
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    total = 0
    for i, c in enumerate(reversed(digits)):
        total += int(c) * pesos[i % 8]
    resto = total % 11
    if resto in (0, 1):
        return 0
    return 11 - resto


def is_codigo_barras_arrecadacao(cod_raw: str) -> bool:
    """True se o código/linha digitável for de arrecadação (1ª posição '8')."""
    d = _limpar_digitavel(cod_raw)
    return bool(d) and d[0] == "8" and len(d) in (44, 48)


def linha_digitavel_arrecadacao_to_barcode(ld: str) -> str:
    """Converte linha digitável de arrecadação (48) para código de barras (44)."""
    d = _limpar_digitavel(ld)
    if len(d) != 48:
        raise ValueError(
            f"Linha digitável de arrecadação deve ter 48 dígitos, recebeu {len(d)}"
        )
    blocos = [d[i:i + 12] for i in range(0, 48, 12)]
    return "".join(b[:11] for b in blocos)


def validar_codigo_barras_arrecadacao(cod_raw: str) -> str:
    """
    Valida e retorna código de barras de arrecadação (44 dígitos).
    Aceita 44 (direto) ou 48 (linha digitável) dígitos. Valida o DV geral
    (posição 4) em módulo 10 ou 11, conforme a posição 3.
    """
    d = _limpar_digitavel(cod_raw)
    if not d:
        raise ValueError("Código de barras de arrecadação vazio")
    if len(d) == 48:
        d = linha_digitavel_arrecadacao_to_barcode(d)
    elif len(d) != 44:
        raise ValueError(
            f"Código de barras de arrecadação deve ter 44 dígitos "
            f"(ou linha digitável de 48). Recebido: {len(d)} dígitos"
        )

    ident_valor = d[2]
    dv_informado = int(d[3])
    aux = d[0:3] + d[4:]  # 43 posições, excluindo o próprio DV
    if ident_valor in ("6", "7"):
        dv_calculado = _mod10_arrecadacao(aux)
    elif ident_valor in ("8", "9"):
        dv_calculado = _mod11_arrecadacao(aux)  # regra própria de arrecadação (resto 0/1 -> DV 0)
    else:
        raise ValueError(
            f"Identificador de valor '{ident_valor}' (posição 3) não reconhecido "
            f"em código de barras de arrecadação"
        )

    if dv_informado != dv_calculado:
        raise ValueError(
            f"DV do código de barras de arrecadação inválido: esperado "
            f"{dv_calculado}, encontrado {dv_informado}"
        )
    return d


def valor_arrecadacao(barcode: str) -> float:
    """Decodifica o valor do documento (posições 5-15) em reais."""
    campo = barcode[4:15]
    if not campo.isdigit():
        return 0.0
    return int(campo) / 100.0


class TributoGenerator:
    """
    Gera Segmento O (Pagamento de Contas e Tributos com Código de Barras)
    do CNAB 240 BB — usado para contas de arrecadação/convênio (água, luz,
    telefone, tributos), que são estruturalmente diferentes do boleto
    bancário (Segmento J) e não podem ser geradas por BoletoGenerator.

    Forma de lançamento '11' e tipo de serviço '98' conforme manual de
    particularidades CNAB 240 do Banco do Brasil.
    """

    FORMA_LANCAMENTO = "11"
    TIPO_SERVICO = "98"
    LAYOUT_LOTE = "012"

    def __init__(self, cfg):
        self.cfg = cfg
        self.p = cfg.pagador
        self.banco = cfg.codigo_banco
        self.erros = []

    def gerar(self, pagamentos: list, payment_date=None) -> str:
        now = datetime.now()
        linhas = []

        pd_dt = None
        if payment_date:
            if isinstance(payment_date, datetime):
                pd_dt = payment_date
            else:
                s = str(payment_date).strip()
                try:
                    if len(s) == 8 and s.isdigit():
                        pd_dt = datetime.strptime(s, "%d%m%Y")
                    elif "/" in s:
                        pd_dt = datetime.strptime(s, "%d/%m/%Y")
                    else:
                        pd_dt = datetime.fromisoformat(s)
                except Exception:
                    pd_dt = None

        base_gen = CNAB240Generator(self.cfg)
        linhas.append(base_gen._h_arq(now))
        linhas.append(self._h_lote_o(pd_dt or now))

        total = 0.0
        seq = 0
        for pg in pagamentos:
            cod_raw = str(pg.get("codigo_barras", "")).strip()
            try:
                barcode = validar_codigo_barras_arrecadacao(cod_raw)
            except ValueError as e:
                self.erros.append(
                    f"{pg.get('nome', '?')} (doc. {pg.get('seu_numero', '?')}): {e}"
                )
                continue

            valor_bc = valor_arrecadacao(barcode)
            valor_excel = _pv(pg.get("valor", 0))
            if valor_bc > 0:
                if abs(valor_bc - valor_excel) > 0.01:
                    self.erros.append(
                        f"{pg.get('nome', '?')} (doc. {pg.get('seu_numero', '?')}): "
                        f"valor da planilha (R$ {valor_excel:.2f}) difere do valor "
                        f"codificado no código de barras (R$ {valor_bc:.2f}); "
                        f"usando o valor do código de barras."
                    )
                pg["valor"] = valor_bc

            seq += 1
            total += _pv(pg.get("valor", 0))
            linhas.append(self._seg_o(pg, barcode, seq, now, pd_dt))

        if seq == 0:
            raise ValueError(
                "Nenhuma conta de arrecadação com código de barras válido neste lote:\n"
                + "\n".join(self.erros)
            )

        qtd_lote = 1 + seq + 1
        qtd_total = 1 + qtd_lote + 1

        linhas.append(self._t_lote(qtd_lote, total))
        linhas.append(base_gen._t_arq(qtd_total))

        for i, l in enumerate(linhas):
            _chk(l, f"linha {i}")

        return "\n".join(linhas) + "\n"

    # ── Header Lote O ─────────────────────────────────────────────
    def _h_lote_o(self, data_pagamento: datetime = None):
        p = self.p
        return (
            self.banco + "0001" + "1" + "C" +                             # 001-009 (9)
            self.TIPO_SERVICO + self.FORMA_LANCAMENTO + self.LAYOUT_LOTE + # 010-016 (7)
            " " +                                                          # 017 (1)
            _tipo_inscricao(p["cnpj"]) + p["cnpj"].zfill(14) +             # 018-032 (15)
            p["convenio"] + p["codigo_bb2"] + " "*7 +                     # 033-052 (20)
            p["agencia"] + "0" + p["dv_agencia"] +                        # 053-058 (6)
            p["conta"] + p["dv_conta"] + p["dv_ag_conta"] +                # 059-072 (14)
            _fa(p["nome_empresa"], 30) +                                  # 073-102 (30)
            " "*40 +                                                      # 103-142 (40) informação/mensagem
            " "*30 + " "*5 + " "*15 +                                     # 143-192 (50) endereço (não usado)
            " "*20 + " "*5 + " "*3 + " "*2 +                              # 193-222 (30) cidade/CEP/UF (não usado)
            "00" +                                                        # 223-224 (2) indicativo forma pagto
            " "*6 +                                                       # 225-230 (6) brancos
            " "*10                                                        # 231-240 (10) ocorrências
        )

    # ── Segmento O ────────────────────────────────────────────────
    def _seg_o(self, pg, barcode, seq, now, payment_now=None):
        nome = _fa(pg.get("nome", ""), 30)
        data_vc = _date(pg.get("data_vencimento", ""), payment_now or now)
        data_pg = _date(pg.get("data_pagamento", ""), payment_now or now)
        val = _pv(pg.get("valor", 0))
        snm = _fa(pg.get("seu_numero", ""), 20)

        return (
            self.banco + "0001" + "3" + _fn(seq, 5) + "O" +   # 001-014 (14)
            "0" + "00" +                                       # 015-017 (3) tipo mov + cód instrução
            barcode +                                          # 018-061 (44) código de barras
            nome +                                             # 062-091 (30) nome da concessionária
            data_vc +                                          # 092-099 (8) data vencimento
            data_pg +                                          # 100-107 (8) data pagamento
            _vc(val) +                                         # 108-122 (15) valor pagamento
            snm +                                              # 123-142 (20) seu número
            " "*20 +                                           # 143-162 (20) nosso número
            " "*68 +                                           # 163-230 (68) uso exclusivo FEBRABAN/CNAB
            " "*10                                             # 231-240 (10) ocorrências
        )

    # ── Trailer Lote ──────────────────────────────────────────────
    def _t_lote(self, qtd, total):
        soma = str(int(round(total * 100))).zfill(18)[:18]
        return (
            self.banco + "0001" + "5" + " "*9 +
            _fn(qtd, 6) + soma + "0"*18 + "0"*6 +
            " "*165 + "0"*10
        )
