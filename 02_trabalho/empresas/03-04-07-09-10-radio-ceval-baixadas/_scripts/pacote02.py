"""Packet 2 (exceptions and company-level questions) for companies 04 and 07: values on
both sides, C0 direction, effect on the 1/6 quota, and the totals each option leads to.
Numbers come from the master spreadsheet and the Phase 0 extraction; the narrative packet
(pacote-02.md) quotes pacotes/pacote-02-tabelas.md."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import openpyxl

BASE = Path(
    r"C:\entrega-cisao-siviero\02_trabalho\empresas\03-04-07-09-10-radio-ceval-baixadas"
)
MASTER = BASE / "planilha-mestre-03-04-07-09-10.xlsx"
EXTRACT = BASE / "_extracao" / "fase0-extracao.json"
OUT = BASE / "pacotes" / "pacote-02-tabelas.md"
EXCEPTIONS = ["07-2025-01", "07-2025-06", "04-2025-01", "04-2025-02"]
REDUCTION = 0.08  # Rádio inventory 2025, par. 17: "redução de 8,0% (OITO POR CENTO)"


def brl(v) -> str:
    if v is None:
        return "—"
    s = f"{v:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def doc00_totals() -> tuple[float, float]:
    """TOTAL of doc 00 (2025) and Telmo de Toni's figure, read from the source."""
    import sys

    sys.path.insert(0, str(Path(__file__).parent))
    from fase0_extracao import SRC_DOC00, docx_blocks, moneys

    total = telmo = None
    for kind, block in docx_blocks(SRC_DOC00):
        if kind == "t" and "E M P R E S A" in " ".join(block[0]):
            for row in block:
                if row[0].upper().startswith("TOTAL"):
                    total = moneys(row[-1].replace("100.874,027,60", "100.874.027,60"))[
                        0
                    ]
        if kind == "p" and "TELMO" in block.upper():
            telmo = moneys(block)[0]
        if total is not None and telmo is not None:
            break
    return total, telmo


def main():
    ws = openpyxl.load_workbook(MASTER, data_only=True)["Mestre"]
    it = ws.iter_rows(values_only=True)
    head = next(it)
    rows = {r[0]: dict(zip(head, r)) for r in it if r[0]}
    data = json.loads(EXTRACT.read_text(encoding="utf-8"))
    out = [
        f"<!-- gerado por _scripts/pacote02.py; planilha-mestre md5 {hashlib.md5(MASTER.read_bytes()).hexdigest()} -->",
        "",
    ]

    out += [
        "## Exceções (grupo × Rafael)",
        "",
        "| ID | Bem | Grupo | Rafael | Diferença | Lista C0? | Valor pelo C0 | Faixa de Rafael | Efeito no 1/6 se adotado o C0 |",
        "|---|---|---|---|---|---|---|---|---|",
    ]
    c0 = {}
    for rid in EXCEPTIONS:
        r = rows[rid]
        g, rv = r["Valor grupo 2025"], r["Valor Rafael"]
        pick = min(g, rv) if r["Lista de Rafael (C0)?"] == "Sim" else max(g, rv)
        c0[rid] = pick
        out.append(
            f"| {rid} | {r['Descrição'][:55]} | {brl(g)} | {brl(rv)} | {brl(rv - g)} | {r['Lista de Rafael (C0)?']} | "
            f"{brl(pick)} | {r['Faixa de Rafael (C1)'] or ''}; outras: {r['Outras referências de Rafael'] or ''} | "
            f"{brl((pick - g) / 6)} |"
        )

    # Company totals under the adopted values (packet 1) + C0 on the exceptions + LK 1313 as is.
    adopted = {
        k: v["Valor adotado"]
        for k, v in rows.items()
        if v["Valor adotado"] not in (None, "")
    }
    radio_ids = [k for k, v in rows.items() if v["Empresa (2025)"] == "07 Rádio"]
    ceval_ids = [k for k, v in rows.items() if v["Empresa (2025)"] == "04 Ceval"]

    def total(ids, lk=None):
        t = 0.0
        for k in ids:
            if k in adopted:
                t += adopted[k]
            elif k in c0:
                t += c0[k]
            elif k == "04-2025-03":
                t += lk if lk is not None else rows[k]["Valor grupo 2025"]
            else:
                t += rows[k]["Valor grupo 2025"]
        return t

    radio_g = sum(rows[k]["Valor grupo 2025"] for k in radio_ids)
    radio_c0 = total(radio_ids)
    ceval_g = sum(rows[k]["Valor grupo 2025"] for k in ceval_ids)
    ceval_c0 = total(ceval_ids)
    out += [
        "",
        "## Totais das empresas",
        "",
        "| Empresa | Grupo 2025 | Rafael | Com pacote 1 + C0 nas exceções (LK 1313 = R$ 45.000) |",
        "|---|---|---|---|",
        f"| 07 Rádio | {brl(radio_g)} | {brl(sum(rows[k]['Valor Rafael'] for k in radio_ids))} | {brl(radio_c0)} |",
        f"| 04 Ceval | {brl(ceval_g)} | {brl(sum(rows[k]['Valor Rafael'] for k in ceval_ids))} | {brl(ceval_c0)} |",
        "",
    ]

    # LK 1313 options.
    lk = rows["04-2025-03"]
    out += [
        "## Caminhão LK 1313 (04-2025-03): opções",
        "",
        "| Opção | Valor | Efeito no 1/6 frente a R$ 45.000 |",
        "|---|---|---|",
    ]
    for label, v in [
        ("Manter grupo = Rafael", lk["Valor grupo 2025"]),
        ("Piso da faixa de Rafael (caso extremo)", 90000.0),
        ("Teto da faixa de Rafael (caso extremo)", 150000.0),
        ("FIPE citada por Rafael", 30600.0),
    ]:
        out.append(f"| {label} | {brl(v)} | {brl((v - lk['Valor grupo 2025']) / 6)} |")

    # 8 % reduction of the Rádio.
    out += [
        "",
        "## Redução de 8 % da Rádio",
        "",
        "| Base | Valor | 8 % | Efeito no 1/6 |",
        "|---|---|---|---|",
    ]
    for label, v in [
        ("Total do grupo (inventário 2025)", radio_g),
        ("Total com pacote 1 + C0", radio_c0),
        (
            "Total de Rafael (balanço de Valter)",
            sum(rows[k]["Valor Rafael"] for k in radio_ids),
        ),
    ]:
        out.append(
            f"| {label} | {brl(v)} | {brl(v * REDUCTION)} | {brl(v * REDUCTION / 6)} |"
        )

    # Telmo de Toni: base of the 1/6.
    doc00_total, telmo = doc00_totals()
    out += [
        "",
        "## Base do 1/6 na Rádio (Telmo de Toni, 1 %)",
        "",
        "| Base | 1/6 do total (Rafael, p. 10) | 16,50 % (inventário e doc 00) | Diferença por sócio |",
        "|---|---|---|---|",
    ]
    for label, v in [
        ("Total do grupo", radio_g),
        ("Total com pacote 1 + C0", radio_c0),
    ]:
        out.append(
            f"| {label} ({brl(v)}) | {brl(v / 6)} | {brl(v * 0.165)} | {brl(v / 6 - v * 0.165)} |"
        )
    out += [
        "",
        (
            f"Doc 00 (todas as empresas): (total {brl(doc00_total)} − Telmo {brl(telmo)}) / 6 = "
            f"{brl((doc00_total - telmo) / 6)}; total / 6 = {brl(doc00_total / 6)}; diferença {brl(telmo / 6)}."
        ),
        f"Telmo = 1 % da Rádio pelo grupo: {brl(radio_g * 0.01)} (doc 00: {brl(telmo)}).",
        "",
    ]

    # NF 4946 annotation: 1/3 of the excavator.
    e = rows["04-2025-01"]
    out += [
        '## Anotação "66,66 Siviero / 33,33 Celito" na NF da 320D',
        "",
        (
            f"Se 1/3 da máquina fosse de terceiro: grupo {brl(e['Valor grupo 2025'] / 3)}; C0 {brl(c0['04-2025-01'] / 3)}; "
            f"efeito no 1/6 {brl(c0['04-2025-01'] / 3 / 6)}."
        ),
        "",
    ]
    # Cross-check doc 00 total against the extraction.
    out += [
        f"Conferência: doc 00 da Rádio na extração = {brl(data['doc00']['07'])}; Ceval = {brl(data['doc00']['04'])}.",
        "",
    ]
    OUT.write_text("\n".join(out) + "\n", encoding="utf-8")
    print("\n".join(out))


if __name__ == "__main__":
    main()
