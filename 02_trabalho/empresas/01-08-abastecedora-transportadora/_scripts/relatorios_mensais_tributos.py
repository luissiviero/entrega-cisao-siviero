"""Taxes, works costs and other findings in the monthly paid-accounts reports (jul/2023 to oct/2024).

Answers Luís's question of 28/10/2026: are INSS, INCC and similar charges
already in the monthly reports, so that a deduction in the patrimonial value
would count them twice? Reads the 'Consolidado' sheet of
`01_fontes/outros/Contas_Pagas_2023-07_a_2024-10.xlsx` (notes are cut at
about 45 characters, so patterns tolerate truncated values), the sold lots and
the recurring/extraordinary classification, and writes
`_extracao/relatorios-mensais-tributos.xlsx`.

Usage: python -I relatorios_mensais_tributos.py
"""

from __future__ import annotations

import collections
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

sys.path.insert(0, str(Path(__file__).resolve().parent))
import extract_fase0 as fase0

WORK = Path(__file__).resolve().parents[1]
OUTROS = fase0.BASE / "01_fontes" / "outros"
SRC = OUTROS / "Contas_Pagas_2023-07_a_2024-10.xlsx"
SRC_LOTS = OUTROS / "Contas_Pagas_e_Lotes_Vendidos_2023-07_a_2024-10.xlsx"
SRC_CLASS = OUTROS / "Contas_Pagas_Recorrentes_x_Extraordinarias_2023-07_a_2024-10.xlsx"
OUT = WORK / "_extracao" / "relatorios-mensais-tributos.xlsx"

MONTHS = {
    "JAN": "JANEIRO", "FEV": "FEVEREIRO", "MAR": "MARCO", "ABR": "ABRIL", "MAI": "MAIO", "JUN": "JUNHO",
    "JUL": "JULHO", "AGO": "AGOSTO", "SET": "SETEMBRO", "OUT": "OUTUBRO", "NOV": "NOVEMBRO", "DEZ": "DEZEMBRO",
}  # fmt: skip
REVENUE_TAX = re.compile(
    r"\bPIS\b|COFINS|IRPJ|CSLL|C\.S\.L\.L|CONT\.? SOCIAL|CONTRIBUICAO SOCIAL",
    re.IGNORECASE,
)
# (label, company prefix, pattern on supplier + notes) for the findings listed in the document
FINDINGS = [
    (
        "01: ISSQN sobre mão de obra da ampliação do posto",
        "ABASTECEDORA",
        re.compile(r"ISSQN.*AMPLIACAO", re.IGNORECASE),
    ),
    (
        "01: topografia da área do posto",
        "ABASTECEDORA",
        re.compile(r"TOPOGRAFIA", re.IGNORECASE),
    ),
    (
        "01: seguro Ford Ranger (veículo fora do inventário)",
        "ABASTECEDORA",
        re.compile(r"RANGER", re.IGNORECASE),
    ),
    ("01: títulos de capitalização", "ABASTECEDORA", re.compile(r"CAPITALIZ", re.IGNORECASE)),
    (
        "01: devolução de aluguel à Borracharia Valverde",
        "ABASTECEDORA",
        re.compile(r"BORRACHARIA.*ALUGU|DEVOLUCAO.*ALUGU", re.IGNORECASE),
    ),
    (
        "01: empréstimo Caixa da Rádio (última parcela)",
        "ABASTECEDORA",
        re.compile(r"EMPRESTIMO CAIXA", re.IGNORECASE),
    ),
    (
        "08: INSS de folha",
        "TRANSPORTADO",
        re.compile(r"INSS.*(SALAR|FUNC|13|FERIAS)", re.IGNORECASE),
    ),
    (
        "08: SIMPLES sobre o faturamento",
        "TRANSPORTADO",
        re.compile(r"\bSIMPLES\b", re.IGNORECASE),
    ),
    (
        "08: licença de operação da oficina (Boff) e alvará",
        "TRANSPORTADO",
        re.compile(r"LICENCA OPERACAO|ALVARA", re.IGNORECASE),
    ),
    ("08: taxa IBAMA", "TRANSPORTADO", re.compile(r"IBAMA", re.IGNORECASE)),
    (
        "08: anuidade CREA do engenheiro",
        "TRANSPORTADO",
        re.compile(r"CREA", re.IGNORECASE),
    ),
    (
        "08: empréstimo Banrisul placas solares",
        "TRANSPORTADO",
        re.compile(r"EMPRESTIMO PLACAS SOLAR", re.IGNORECASE),
    ),
    (
        "08: manutenção das placas solares",
        "TRANSPORTADO",
        re.compile(r"MANUTENCAO.*PLACAS SOLAR", re.IGNORECASE),
    ),
    (
        "08: garantia 'conta TIS' (registro, averbação, avaliação)",
        "TRANSPORTADO",
        re.compile(r"GARANTIA", re.IGNORECASE),
    ),
    (
        "08: iluminação e cabos do pavilhão da oficina (Enertec)",
        "TRANSPORTADO",
        re.compile(r"ENERTEC", re.IGNORECASE),
    ),
    (
        "08: vidros, tanque e soleiras da oficina",
        "TRANSPORTADO",
        re.compile(r"VIDRO|SOLEIRA", re.IGNORECASE),
    ),
    (
        "08: compressor e central telefônica",
        "TRANSPORTADO",
        re.compile(r"COMPRESSOR|CENTRAL TELEF", re.IGNORECASE),
    ),
    ("08: títulos de capitalização (texto cortado: 'PARCELA' / 'PARCELA CA…' da conta TIS)", "TRANSPORTADO", re.compile(r"BANCO ESTADO RIO GRANDE DO SUL - TIS.*IRMAOS SIVIERO.{0,10}- PARCELA", re.IGNORECASE)),
    ("08: consórcio Banrisul", "TRANSPORTADO", re.compile(r"CONSORCIO", re.IGNORECASE)),
    ("08: IPVA de veículos", "TRANSPORTADO", re.compile(r"^IPVA", re.IGNORECASE)),
    ("08: IPVA pago como vale a funcionário", "TRANSPORTADO", re.compile(r"VALE.*IPVA", re.IGNORECASE)),
    (
        "08: TAC por infração ambiental (Promotoria)",
        "TRANSPORTADO",
        re.compile(r"INFRACAO AMBIENTAL|TERMO ACORDO", re.IGNORECASE),
    ),
    (
        "08: água oficina/posto/lancheria (CORSAN)",
        "TRANSPORTADO",
        re.compile(r"CORSAN", re.IGNORECASE),
    ),
    (
        "SEI: comissões de corretagem",
        "SIVIERO EMPR",
        re.compile(r"COMISS|CORRETAG|INTERMEDIA", re.IGNORECASE),
    ),
]


def money(text: str) -> float | None:
    """'R$ 1.256.000,0' / '604.833,3' / '297.145,O0' (truncated or typed with O) -> float."""
    m = re.search(r"R\$\s*=?\s*([\d.]+),([\dO]{0,2})", text)
    if not m:
        return None
    cents = m.group(2).replace("O", "0").ljust(2, "0")
    return float(m.group(1).replace(".", "") + "." + cents)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    f = fase0.fmt
    ws = openpyxl.load_workbook(SRC, read_only=True, data_only=True)["Consolidado"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    rows = [dict(zip(header, r, strict=False)) for r in rows[1:]]
    months = sorted({str(r["Mês"])[:7] for r in rows})
    by_co = collections.Counter(str(r["Empresa"])[:12] for r in rows)
    print(
        f"fonte {SRC.name} md5 {fase0.md5_10(SRC)}: {len(rows)} lançamentos ({dict(by_co)}), {len(months)} meses ({months[0]} a {months[-1]})"
    )

    def text(r: dict) -> str:
        return f"{r['Fornecedor'] or ''} {r['Observações'] or ''}"

    def val(r: dict) -> float:
        return float(r["Valor Dupl."] or 0)

    # -- SEI: taxes on monthly revenue ("FATURAMENTO MES <month>/<year> = R$ ...")
    sei_rev_tax = [
        r
        for r in rows
        if str(r["Empresa"]).startswith("SIVIERO EMPR") and REVENUE_TAX.search(text(r))
    ]
    on_revenue = [
        r
        for r in sei_rev_tax
        if "FATURAMENTO MES" in str(r["Observações"] or "").upper()
    ]
    others = [r for r in sei_rev_tax if r not in on_revenue]
    revenue, tax = {}, collections.defaultdict(float)
    for r in on_revenue:
        obs = str(r["Observações"]).upper()
        m = re.search(r"FATURAMENTO MES ([A-Z]+)\s*/\s*(\d{4})", obs)
        key = f"{MONTHS.get(m.group(1)[:3], m.group(1))}/{m.group(2)}"
        tax[key] += val(r)
        v = money(obs)
        if v and v > revenue.get(key, 0):
            revenue[key] = v
    total_rev, total_tax = sum(revenue.values()), sum(tax.values())
    print(
        f"SEI impostos sobre o faturamento mensal: {len(on_revenue)} lançamentos, {len(revenue)} meses; "
        f"faturamento {f(total_rev)}; impostos {f(total_tax)}; alíquota efetiva {fase0.fmt(total_tax / total_rev * 100)} %"
    )
    for k, rev in revenue.items():
        print(
            f"   {k}: faturamento {f(rev)}; impostos {f(tax[k])} ({fase0.fmt(tax[k] / rev * 100)} %)"
        )
    print(
        f"SEI: outros {len(others)} lançamentos de PIS/COFINS/IRPJ/CSLL que não são do faturamento mensal (Ceval, soja, notas de serviço): {f(sum(val(r) for r in others))}"
    )
    for r in others:
        print(
            f"   {str(r['Mês'])[:7]} | {r['Fornecedor']} | {f(val(r))} | {r['Observações']}"
        )

    # -- 01: taxes on its own revenue, and the revenue itself
    ab = [
        r
        for r in rows
        if str(r["Empresa"]).startswith("ABASTECEDORA") and REVENUE_TAX.search(text(r))
    ]
    ab_rev = {}
    for r in ab:
        m = re.search(r"(\w+)/(\d{4})", str(r["Observações"] or ""))
        v = money(str(r["Observações"] or ""))
        if m and v:
            ab_rev[f"{m.group(1)}/{m.group(2)}"] = v
    print(
        f"01 impostos sobre o faturamento (PIS, COFINS, IRPJ, CSLL): {len(ab)} lançamentos, {f(sum(val(r) for r in ab))}"
    )
    print(
        f"01 faturamento citado nas observações: {sorted({f(v) for v in ab_rev.values()})} ({len(ab_rev)} meses)"
    )

    # -- other findings
    found = collections.OrderedDict()
    for label, co, pat in FINDINGS:
        hits = [
            r for r in rows if str(r["Empresa"]).startswith(co) and pat.search(text(r))
        ]
        found[label] = hits
        nf = {}
        for r in hits:
            nf[(r["Fornecedor"], r["Valor NF"])] = float(r["Valor NF"] or 0)
        print(
            f"{label}: {len(hits)} lançamentos, pago {f(sum(val(r) for r in hits))}; soma das NF distintas {f(sum(nf.values()))}"
        )

    loan = found["08: empréstimo Banrisul placas solares"]
    last = max(loan, key=lambda r: str(r["Mês"]))
    m = re.search(r"P\.\s*(\d+)/(\d+)", str(last["Observações"]))
    paid_last, total_n = int(m.group(1)), int(m.group(2))
    remaining = total_n - (paid_last + (12 - int(str(last["Mês"])[5:7])))
    print(
        f"placas solares: última no período P.{paid_last}/{total_n} ({f(val(last))}); em 31.12.2024 restam cerca de {remaining} parcelas, soma ≈ {f(remaining * val(last))} com juros (não é o saldo devedor)"
    )
    plates = sorted(
        {
            p.replace("O", "0") if p[3] == "O" else p
            for r in found["08: IPVA de veículos"]
            for p in re.findall(r"-\s*([A-Z]{3}\s?[0-9O][A-Z0-9][0-9O]{2})\b", str(r["Observações"] or "").upper())
        }
    )
    print(f"08: placas citadas nos IPVA: {plates}")

    # -- lots sold and commission classification
    lots = list(
        openpyxl.load_workbook(SRC_LOTS, read_only=True, data_only=True)[
            "Lotes Vendidos (todos)"
        ].iter_rows(values_only=True)
    )
    lots = [dict(zip(lots[0], r, strict=False)) for r in lots[1:]]
    in_period = [
        r
        for r in lots
        if str(r.get("No período dos relatórios") or "").lower().startswith("s")
    ]
    cited = sum(
        float(r["Valor citado"])
        for r in in_period
        if isinstance(r.get("Valor citado"), (int, float))
    )
    print(
        f"lotes vendidos no período: {len(in_period)}; soma dos valores citados {f(cited)}"
    )
    rec = list(
        openpyxl.load_workbook(SRC_CLASS, read_only=True, data_only=True)[
            "Recorrentes"
        ].iter_rows(values_only=True)
    )
    rec = [dict(zip(rec[0], r, strict=False)) for r in rec[1:]]
    com = [r for r in rec if "comiss" in str(r.get("Categoria") or "").lower()]
    print(
        f"classificação Recorrentes: categoria de comissões {len(com)} lançamentos, {f(sum(float(r['Valor Dupl.'] or 0) for r in com))}"
    )
    comm = found["SEI: comissões de corretagem"]
    print(
        f"comissões SEI / faturamento: {fase0.fmt(sum(val(r) for r in comm) / total_rev * 100)} %"
    )

    out = openpyxl.Workbook()
    sh = out.active
    sh.title = "SEI faturamento"
    sh.append(
        [
            "Mês",
            "Faturamento citado (R$)",
            "PIS+COFINS+IRPJ+CSLL (R$)",
            "Alíquota efetiva",
        ]
    )
    for k, rev in revenue.items():
        sh.append([k, rev, tax[k], tax[k] / rev])
    sh.append(["Total", total_rev, total_tax, total_tax / total_rev])
    for name, data in [("SEI outros impostos", others), ("01 impostos", ab)] + [
        (k[:31], v) for k, v in found.items()
    ]:
        s = out.create_sheet(re.sub(r"[\\/*?:\[\]']", "", name)[:31])
        s.append(
            [
                "Mês",
                "Empresa",
                "Fornecedor",
                "Valor NF (R$)",
                "Valor Dupl. (R$)",
                "Observações",
            ]
        )
        for r in data:
            s.append(
                [
                    str(r["Mês"])[:7],
                    r["Empresa"],
                    r["Fornecedor"],
                    r["Valor NF"],
                    r["Valor Dupl."],
                    r["Observações"],
                ]
            )
    for s in out.worksheets:
        for c in s[1]:
            c.font = Font(bold=True)
        s.freeze_panes = "A2"
    out.save(OUT)
    print(f"saved {OUT}")


if __name__ == "__main__":
    main()
