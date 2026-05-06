"""
config_loader.py  –  Carrega, valida e normaliza o config.json.
"""
import json, os, re


class ConfigError(Exception):
    pass


class ConfigLoader:

    REQUIRED_PAGADOR = [
        "cnpj", "convenio", "codigo_bb2",
        "agencia", "dv_agencia", "conta", "dv_conta",
        "nome_empresa", "versao_layout_arquivo", "versao_layout_lote",
    ]
    REQUIRED_LOTE    = ["tipo_servico", "forma_lancamento", "camara"]
    REQUIRED_ARQUIVO = ["codigo_banco"]

    def __init__(self, config_path: str = None, tipo_pagamento: str = None):
        if config_path is None:
            base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(base, "config", "config.json")

        self.config_path   = config_path
        self._tipo         = (tipo_pagamento or "").upper().strip()
        self._profiles     = {}
        self._profile_name = ""
        self._data         = {}
        self._raw          = {}
        self._load()
        self._select_profile()
        self._validate()
        self._normalize()

    # ── load ──────────────────────────────────────────────────────
    def _load(self):
        if not os.path.exists(self.config_path):
            raise ConfigError(f"config.json não encontrado: {self.config_path}")
        try:
            with open(self.config_path, encoding="utf-8") as f:
                self._raw = json.load(f)
        except json.JSONDecodeError as e:
            raise ConfigError(f"config.json inválido (JSON): {e}")

        raw          = self._raw
        base_arquivo = raw.get("arquivo", {})

        if isinstance(raw.get("pagadores"), dict) and raw["pagadores"]:
            self._profiles = {
                k: self._build(v, base_arquivo)
                for k, v in raw["pagadores"].items()
                if isinstance(v, dict)
            }
            if not self._profiles:
                raise ConfigError("Nenhum perfil válido em 'pagadores'.")
            return

        raise ConfigError("config.json: esperado bloco 'pagadores' com perfis.")

    def _build(self, p: dict, base_arquivo: dict) -> dict:
        lote = p.get("lote", {}) if isinstance(p.get("lote"), dict) else {}
        return {
            "pagador": {
                "cnpj":                  str(p.get("cnpj", "")),
                "convenio":              str(p.get("convenio", "")),
                "codigo_bb2":            str(p.get("codigo_bb2", "")),
                "agencia":               str(p.get("agencia", "")),
                "dv_agencia":            str(p.get("dv_agencia", "5")),
                "conta":                 str(p.get("conta", "")),
                "dv_conta":              str(p.get("dv_conta", " ")),
                "dv_ag_conta":           str(p.get("dv_ag_conta", " ")),
                "nome_empresa":          str(p.get("nome_empresa", "")),
                "versao_layout_arquivo": str(p.get("versao_layout_arquivo", "084")),
                "versao_layout_lote":    str(p.get("versao_layout_lote", "045")),
            },
            "lote": {
                "tipo_servico":     str(lote.get("tipo_servico", "98")),
                "forma_lancamento": str(lote.get("forma_lancamento", "03")),
                "camara":           str(lote.get("camara", "018")),
            },
            "arquivo": {
                "codigo_banco": str(base_arquivo.get("codigo_banco", "001")),
            },
        }

    # ── select profile ────────────────────────────────────────────
    def _select_profile(self):
        if not self._profiles:
            raise ConfigError("Nenhum perfil carregado.")

        t = self._tipo
        if t in ("JUSMP", "BBJUSMP", "BBJUMP", "APLICACAO", "APLICACOES"):
            pref = "PAGADOR_BBJUSMP"
        elif t in ("PGA", ""):
            pref = "PAGADOR_BPGA"
        else:
            pref = ""

        if pref and pref in self._profiles:
            self._profile_name = pref
        else:
            self._profile_name = next(iter(self._profiles))

        self._data = self._profiles[self._profile_name]

    # ── validate ──────────────────────────────────────────────────
    def _validate(self):
        errors = []
        for section, fields in [
            ("pagador", self.REQUIRED_PAGADOR),
            ("lote",    self.REQUIRED_LOTE),
            ("arquivo", self.REQUIRED_ARQUIVO),
        ]:
            if section not in self._data:
                errors.append(f"Seção '{section}' ausente.")
                continue
            for field in fields:
                if not str(self._data[section].get(field, "")).strip():
                    errors.append(f"Campo vazio: [{section}][{field}]")
        if errors:
            raise ConfigError("Erros no config.json:\n" + "\n".join(errors))

    # ── normalize ─────────────────────────────────────────────────
    def _normalize(self):
        p = self._data["pagador"]
        p["cnpj"]    = re.sub(r"\D", "", p["cnpj"]).zfill(14)[:14]
        p["convenio"]= re.sub(r"\D", "", p["convenio"]).zfill(9)[:9]
        p["codigo_bb2"]  = str(p["codigo_bb2"]).zfill(4)[:4]
        p["agencia"]     = re.sub(r"\D", "", p["agencia"]).zfill(4)[:4]
        p["dv_agencia"]  = str(p["dv_agencia"])[:1]
        p["conta"]       = re.sub(r"\D", "", p["conta"]).zfill(12)[:12]
        p["dv_conta"]    = str(p["dv_conta"])[:1]
        p["dv_ag_conta"] = str(p["dv_ag_conta"])[:1]
        p["nome_empresa"]= str(p["nome_empresa"]).ljust(30)[:30]
        p["versao_layout_arquivo"] = str(p["versao_layout_arquivo"]).zfill(3)[:3]
        p["versao_layout_lote"]    = str(p["versao_layout_lote"]).zfill(3)[:3]

        # Normaliza forma_lancamento: garante 2 chars numéricos
        lt = self._data["lote"]
        lt["forma_lancamento"] = str(lt["forma_lancamento"]).zfill(2)[:2]
        lt["tipo_servico"]     = str(lt["tipo_servico"]).zfill(2)[:2]
        lt["camara"]           = str(lt["camara"]).zfill(3)[:3]

    # ── properties ────────────────────────────────────────────────
    @property
    def pagador(self)       -> dict: return self._data["pagador"]
    @property
    def lote(self)          -> dict: return self._data["lote"]
    @property
    def arquivo(self)       -> dict: return self._data["arquivo"]
    @property
    def codigo_banco(self)  -> str:  return str(self._data["arquivo"]["codigo_banco"]).zfill(3)
    @property
    def profile_name(self)  -> str:  return self._profile_name
    @property
    def aplicacoes_config(self) -> dict: return self._raw.get("aplicacoes", {})
