"""Extra slides for the Part 4 deck: why the group's charges are double charges, and one summary slide per study
session (S1-S4) plus the rest of the SEI. Figures are read from each session's master sheet (read-only)."""

from __future__ import annotations

from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[4]
EMP = ROOT / "02_trabalho/empresas"


def br(v, dec=2):
    """Brazilian number format: 9.661.414 -> 9,66."""
    return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _rows(path, sheet):
    return list(
        openpyxl.load_workbook(path, data_only=True)[sheet].iter_rows(values_only=True)
    )


def _col_sum(path, sheet, header, key_filter=None):
    rows = _rows(path, sheet)
    h = rows[0]
    j = h.index(header)
    k = h.index("Chave") if key_filter else None
    return sum(
        r[j]
        for r in rows[1:]
        if isinstance(r[j], (int, float))
        and (not key_filter or key_filter(str(r[k] or "")))
    )


def _lookup(rows, label_start, col=1):
    for r in rows:
        if r and isinstance(r[0], str) and r[0].startswith(label_start):
            return r[col]
    raise KeyError(label_start)


def session_figures():
    """[(title, source, grupo, rafael, proposta, proposta_label, comments)]"""
    out = []
    # S1 — 02 São Crispin (Agropecuária) and 05 Hotel Princesa dos Vales
    f = EMP / "02-05-sao-crispin-hotel-princesa/planilha-mestre-02-05.xlsx"
    res = _rows(f, "Resumo")
    mest = _rows(f, "Mestre")
    ia = mest[0].index("Valor adotado (R$)")
    decided = [r for r in mest[1:] if isinstance(r[ia], (int, float))]
    out.append(
        (
            "S1 — São Crispin e Hotel Princesa dos Vales (02 e 05)",
            "planilha-mestre-02-05.xlsx (sessão S1), aba Resumo",
            float(_lookup(res, "Bens 2025 do Hotel pelo grupo")),
            float(_lookup(res, "Bens 2025 do Hotel por Rafael")),
            float(_lookup(res, "Soma pelo critério C0")),
            "Proposta (critério C0, referência)",
            [
                f"Só {len(decided)} de {len(mest) - 1} itens já decididos (R$ {sum(r[ia] for r in decided) / 1e6:,.1f} mi): "
                "a proposta usa o critério C0 como referência.".replace(".", ",", 1),
                (
                    "Agropecuária São Crispin: R$ 0 em 2025. Créditos em soja sem valor (Rigo 37.694 sacas; Lorenzet 9.403 + "
                    "3.682 sacas) vão para a Parte 5."
                ),
                "Em andamento: quanto a São Crispin recebeu com a venda de áreas, quanto já saiu e quanto já foi pago.",
            ],
        )
    )
    # S2 — 06 Hotel Araçá
    f = EMP / "06-hotel-araca/planilha-mestre-06.xlsx"
    tot = _rows(f, "Totais")
    out.append(
        (
            "S2 — Hotel Araçá (06)",
            "planilha-mestre-06.xlsx (sessão S2), aba Totais",
            float(_lookup(tot, "Total grupo 2025")),
            float(_lookup(tot, "Soma dos itens de Rafael")),
            float(_lookup(tot, "Valores adotados (decisões aprovadas)")),
            "Proposta (decisões aprovadas)",
            [
                f"A diferença está em 2 itens: R$ {float(_lookup(tot, 'Diferença explicada por itens')) / 1e6:,.1f} mi "
                "(Rafael acima do grupo).".replace(".", ",", 1),
                "O total declarado por Rafael é R$ 360 mil maior que a soma dos itens dele, sem explicação.",
                "Pendente: área construída do hotel caiu de 12.000 m² (2018) para 10.500 m² (2025).",
            ],
        )
    )
    # S3 — 01 Abastecedora and 08 Transportadora
    f = EMP / "01-08-abastecedora-transportadora/planilha-mestre-01-08.xlsx"
    res = _rows(f, "Resumo")
    g = float(_lookup(res, "Valor grupo 2025", 3))
    out.append(
        (
            "S3 — Abastecedora e Transportadora (01 e 08)",
            "planilha-mestre-01-08.xlsx (sessão S3), abas Resumo e Mestre",
            g,
            float(_lookup(res, "Valor Rafael", 3)),
            float(_col_sum(f, "Mestre", "Valor adotado (R$)")),
            "Proposta (decisões aprovadas)",
            [
                (
                    f"Abastecedora: grupo R$ {br(float(_lookup(res, 'Valor grupo 2025', 1)) / 1e6)} mi × Rafael "
                    f"R$ {br(float(_lookup(res, 'Valor Rafael', 1)) / 1e6)} mi."
                ),
                (
                    f"Transportadora: {_lookup(res, 'Concorda (quantidade)', 2)} de {_lookup(res, 'Itens 2025', 2)} itens com o mesmo "
                    "valor nas duas avaliações."
                ),
                "Bens de 2018 sem par em 2025, não deduzidos: R$ 324 mil.",
            ],
        )
    )
    # S4 — 03, 04, 07, 09, 10
    f = EMP / "03-04-07-09-10-radio-ceval-baixadas/planilha-mestre-03-04-07-09-10.xlsx"
    rs = [r for r in _rows(f, "Resumo por empresa")[1:] if r and r[0]]
    gs, rf = (
        sum(float(r[3] or 0) for r in rs),
        sum(float(r[7] or r[3] or 0) for r in rs),
    )
    by = {r[0]: r for r in rs}
    out.append(
        (
            "S4 — Rádio, Ceval e baixadas (03, 04, 07, 09, 10)",
            "planilha-mestre-03-04-07-09-10.xlsx (sessão S4)",
            gs,
            rf,
            float(_col_sum(f, "Mestre", "Valor adotado")),
            "Proposta (decisões aprovadas)",
            [
                (
                    f"Rádio: R$ {br(by['07 Rádio'][3] / 1e6)} mi × R$ {br(by['07 Rádio'][7] / 1e6)} mi. Ceval: "
                    f"R$ {br(by['04 Ceval'][3] / 1e3, 0)} mil × R$ {br(by['04 Ceval'][7] / 1e3, 0)} mil."
                ),
                "Associação Cultural, Turismo e Siviero Transportes: sem bens em 2025.",
                "A Ceval aparece no estudo de Rafael com uma escavadeira que também está na SEI (possível dupla contagem).",
            ],
        )
    )
    # rest of the SEI — assets (11.1), without the 11.2 topic of this study
    f = ROOT / "02_trabalho/planilha-mestre-sei.xlsx"
    only_assets = lambda k: "2025" in k and "CONTR" not in k
    mest = _rows(f, "Mestre")
    h = mest[0]
    ig, ir, ik = (
        h.index("Valor grupo 2025 (R$)"),
        h.index("Valor avaliado por Rafael (R$)"),
        h.index("Chave"),
    )
    raf = sum(
        (r[ir] if isinstance(r[ir], (int, float)) else (r[ig] or 0))
        for r in mest[1:]
        if only_assets(str(r[ik] or ""))
    )
    out.append(
        (
            "SEI — bens (restante, sem o tópico 11.2)",
            "planilha-mestre-sei.xlsx (sessão SEI), aba Mestre",
            float(_col_sum(f, "Mestre", "Valor grupo 2025 (R$)", only_assets)),
            float(raf),
            float(_col_sum(f, "Mestre", "Valor adotado (R$)", only_assets)),
            "Proposta (decisões aprovadas)",
            [
                "Lotes, áreas, imóveis, veículos e máquinas. Onde Rafael não avalia, vale o valor do grupo.",
                "Critério C0: bem que virá para nós fica com o menor valor; bem que fica na empresa, com o maior.",
                "O tópico 11.2 e o valor sem explicação estão nos slides anteriores (Parte 4).",
            ],
        )
    )
    return out
