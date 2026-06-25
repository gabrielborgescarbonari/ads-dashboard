# -*- coding: utf-8 -*-
"""Deriva Publico e Formato a partir do nome da campanha.

Regras validadas contra a aba Procv do Excel GROWTH (133 mapeamentos):
Formato 100%, Publico 99,2% (o unico furo e a campanha PF-PME, ambigua de proposito).
Quando nao souber classificar, devolve um valor sentinela ('?') para revisao manual.
"""

REDE_MAP = {"SCH": "Search", "PMAX": "Pmax", "DGEN": "Demand Gen",
            "GDN": "Demand Gen", "YT": "Youtube"}

PUBLICOS_VALIDOS = {"PF", "PME", "PPO", "Awareness", "Venda Serviço", "Odonto"}
FORMATOS_GOOGLE = {"Search", "Pmax", "Demand Gen", "Youtube"}
FORMATOS_META = {"Clique pro Whatsapp", "Leadform", "Clique pro site"}

# Nova nomenclatura (rollout jun/2026): o nome do ativo no gerenciador passou a ser
# o SLUG minusculo com underscore (por causa das UTMs dinamicas do Meta), no lugar do
# CAPSLOCK com " | ". Mesmo conteudo da taxonomia, mesma ordem de 8 campos:
# frente_funil_temperatura_rede_local_objetivo_data_descricao...
# A descricao pode trazer underscores extras (pipes/espacos da origem), entao os 7
# primeiros tokens sao fixos e o resto e descricao.
NEW_FRENTES = {"pf", "pme", "all", "ppo"}
NEW_FUNIS = {"perpetuo", "distribuicao", "awareness", "fullfunnel"}


def _is_tax(ca: str) -> bool:
    return " | " in ca


def _is_new_slug(ca: str) -> bool:
    # Detecta pelo par frente+funil nas 2 primeiras posicoes: nao casa com os legados
    # (e_google..., e_meta..., h_meta..., nomes de YT com espaco no inicio).
    if " | " in ca:
        return False
    toks = ca.lower().split("_")
    return len(toks) >= 7 and toks[0] in NEW_FRENTES and toks[1] in NEW_FUNIS


def derive_publico(ca: str) -> str:
    ca = str(ca)
    ca_l = ca.lower()
    # PPO = PME premium, tratado como publico proprio. Mesma regra do gerador do dash:
    # "PPO" no nome -> PPO. (O caso SPRJ JUN26, cujo nome cru nao tem PPO, e resolvido
    # pelo de-para do gerar_dashboard_midia.py na hora de montar o dash.)
    if "ppo" in ca_l:
        return "PPO"
    if _is_new_slug(ca):
        toks = ca_l.split("_")
        frente, funil = toks[0], toks[1]
        if funil == "awareness":
            return "Awareness"
        if frente in ("pf", "pme"):
            return frente.upper()
        if frente == "all":
            desc = "_".join(toks[7:])
            return "PME" if desc.endswith("pme") else "PF"  # mesma regra do ALL no pipe
        return "?"
    if _is_tax(ca):
        parts = [p.strip() for p in ca.split(" | ")]
        funil = parts[1].upper() if len(parts) > 1 else ""
        pub = parts[0].upper()
        desc = parts[-1].upper()
        if funil == "AWARENESS":
            return "Awareness"
        if pub in ("PF", "PME"):
            return pub
        if pub == "ALL":
            return "PME" if desc.endswith("PME") else "PF"
        return "?"
    # legado minusculo
    if "vendaservico" in ca_l or "vendacheckup" in ca_l:
        return "Venda Serviço"
    if "_pme_" in ca_l or "_pme-" in ca_l:
        return "PME"
    if "_pf_" in ca_l or "_pf-" in ca_l:
        return "PF"
    if ca_l.startswith("e_meta") or ca_l.startswith("h_meta"):
        return "PF"  # mar/aberto e afins sem token de publico -> PF
    return "?"


def derive_formato(ca: str, plataforma: str) -> str:
    ca = str(ca)
    ca_l = ca.lower()
    if plataforma == "Meta":
        if "whatsapp" in ca_l:
            return "Clique pro Whatsapp"
        if "leadform" in ca_l or "form nativo" in ca_l or "native" in ca_l:
            return "Leadform"
        return "Clique pro site"
    # Google
    if _is_new_slug(ca):
        toks = ca_l.split("_")
        return REDE_MAP.get(toks[3].upper(), "?") if len(toks) > 3 else "?"
    if _is_tax(ca):
        parts = [p.strip().upper() for p in ca.split(" | ")]
        rede = parts[3] if len(parts) > 3 else ""
        return REDE_MAP.get(rede, "?")
    toks = ca_l.split("_")
    if len(toks) > 2:
        return REDE_MAP.get(toks[2].upper(), "?")
    return "?"


def precisa_revisao(ca: str, plataforma: str) -> bool:
    """True quando o codigo nao classificou com seguranca (entra na lista de revisao)."""
    pb = derive_publico(ca)
    fm = derive_formato(ca, plataforma)
    return pb == "?" or fm == "?"
