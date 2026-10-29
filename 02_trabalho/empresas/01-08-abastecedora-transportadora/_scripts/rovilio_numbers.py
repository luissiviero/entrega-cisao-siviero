"""Figures for rovilio-quotas.md (sale of Rovílio Siviero's quotas, 28/08/2025).

Reads the payment items from the contract transcription (checked on the page
images), the company totals from the 2025 inventories and the SEI liabilities
list, and prints the comparison table used in rovilio-quotas.md.

Usage: python -I rovilio_numbers.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import extract_fase0 as fase0

WORK = Path(__file__).resolve().parents[1]


def money_lines_after(text: str, anchor: str) -> list[float]:
    """Stand-alone 'R$ x' lines after the anchor paragraph (the SEI 'observações' list)."""
    lines = text.splitlines()
    start = next(i for i, ln in enumerate(lines) if ln.startswith(anchor))
    out = []
    for ln in lines[start:]:
        m = re.fullmatch(r"\s*R\$\s*([\d.]+,\d{2})\s*(\(-\))?\s*", ln)
        if m:
            out.append(fase0.brl(m.group(1)))
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    f = fase0.fmt
    tr = (WORK / "_extracao" / "arq10-rovilio-transcricao.md").read_text(
        encoding="utf-8"
    )
    pay = {
        int(n): fase0.brl(v)
        for v, n in re.findall(
            r"^\| R\$\s*([\d.]+,\d{2}) \| Item (\d+)", tr, re.MULTILINE
        )
    }
    price = fase0.brl(
        re.search(r"importância de \*\*<u>R\$\s*([\d.]+,\d{2})", tr).group(1)
    )
    assert abs(sum(pay.values()) - price) < 0.005, (
        "payment items do not add up to the price"
    )

    commission = pay[7] + pay[8]
    apartment = pay[9]
    before_cutoff = (
        pay[1] + pay[5] + pay[6] + pay[9]
    )  # dated 24/03/2023, 01/11/2012, 25/10/2013, maio/2022
    to_deliver = pay[3] + pay[4] + pay[8]

    def table_total(prefix: str, label: str) -> float:
        """Value in the row whose first cell starts with the label (summary rows of an inventory)."""
        for _, _, cells in fase0.docx_rows(fase0.one(fase0.INV25, prefix)):
            if cells and " ".join(cells).upper().startswith(label):
                return fase0.brl(fase0.MONEY.findall(" ".join(cells))[-1])
        raise SystemExit(f"row {label!r} not found in {prefix}")

    sei_gross = table_total("11. ", "TOTAL GERAL DOS BENS MÓVEIS E IMÓVEIS")
    sei_net = table_total("11. ", "TOTAL LÍQUIDO PATRIMONIAL")
    others = table_total("11. ", "VALOR R$")
    # the 'observações' are plain paragraphs after the last table; table cells are not part of the list
    sei_doc = fase0.docx.Document(str(fase0.one(fase0.INV25, "11. ")))
    paragraphs = "\n".join(p.text for p in sei_doc.paragraphs)
    obs = money_lines_after(paragraphs, "As cotas de capital do sócio Rovílio")
    total01 = table_total("01. ", "TOTAL GERAL DA EMPRESA")
    total08 = table_total("08. ", "TOTAL GERAL DA EMPRESA")
    total09 = total10 = 0.0  # companies 09 and 10: R$ 0,00 in doc 00 (2025)

    print(f"preço (cl. 3ª): {f(price)}; soma itens 1–10: {f(sum(pay.values()))}")
    for n in sorted(pay):
        print(f"  item {n}: {f(pay[n])}")
    print(
        f"comissão (itens 7+8): {f(commission)}; preço sem comissão: {f(price - commission)}"
    )
    print(
        f"apartamento do vendedor (item 9): {f(apartment)}; preço sem comissão e sem item 9: {f(price - commission - apartment)}"
    )
    print(
        f"pagos antes de 31.12.2024 (itens 1, 5, 6, 9): {f(before_cutoff)}; item 2 (fev/2025): {f(pay[2])}; a entregar (3, 4, 8): {f(to_deliver)}"
    )
    print(f"SEI bens (bruto): {f(sei_gross)}; SEI líquido: {f(sei_net)}")
    print(
        f"SEI 'demais compromissos': {f(others)}; soma das {len(obs)} linhas das observações: {f(sum(obs))}; primeira linha (Rovílio a pagar): {f(obs[0])}"
    )
    base_net = sei_net + total01 + total08
    base_gross = sei_gross + total01 + total08
    print(
        f"2 % (SEI líquido + 01 + 08) = 2 % × {f(base_net)} = {f(0.02 * base_net)}  [SEI {f(0.02 * sei_net)}; 01 {f(0.02 * total01)}; 08 {f(0.02 * total08)}]"
    )
    print(f"  + 1 % de 09 e 10 (R$ 0,00 em 2025) = {f(0.01 * (total09 + total10))}")
    print(
        f"2 % (SEI bens brutos + 01 + 08) = 2 % × {f(base_gross)} = {f(0.02 * base_gross)}"
    )
    print(
        f"SEI líquido sem o 'a pagar' do Rovílio: {f(sei_net + obs[0])}; 2 % com essa base: {f(0.02 * (sei_net + obs[0] + total01 + total08))}"
    )
    for label, value in [
        ("preço", price),
        ("preço sem comissão", price - commission),
        ("preço sem comissão e sem item 9", price - commission - apartment),
    ]:
        print(
            f"{label}: {f(value)} = {value / (0.02 * base_net) * 100:.2f} % de 2 % (líquido); valor implícito de 100 % = {f(value / 0.02)}"
        )
    print(
        f"Dr. Pablo: (3.440.082 / 2.538.208,55 − 1) = {(price / (0.02 * base_net) - 1) * 100:.2f} %; diferença {f(price - 0.02 * base_net)}"
    )
    print(f"base 2 % abaixo do preço: {(1 - 0.02 * base_net / price) * 100:.2f} %")

    # sensitivity of the base (the SEI net value already deducts part of the price itself)
    bases = {
        "SEI líquido": base_net,
        "SEI líquido + comissão do item 8 a pagar": base_net + pay[8],
        "SEI líquido + todo o 'a pagar' do Rovílio": base_net + obs[0],
        "SEI bens brutos": base_gross,
    }
    for label, base in bases.items():
        two = 0.02 * base
        print(
            f"base {label}: 2 % = {f(two)}; preço = {price / two * 100:.2f} %; sem comissão = {(price - commission) / two * 100:.2f} %"
        )

    # 2018 benchmark: Rovílio's share of the 2018 totals (doc 27), 2 % in 01, 08 and SEI, 1 % in 09 and 10
    resumo18 = {}
    for _, _, cells in fase0.docx_rows(fase0.one(fase0.DRIVE2018, "27. 1. ")):
        if len(cells) >= 2 and fase0.MONEY.search(cells[-1]):
            resumo18[cells[0].upper()] = fase0.brl(fase0.MONEY.findall(cells[-1])[-1])
    share18 = {
        "ABASTECEDORA": 0.02,
        "TRANSPORTADORA": 0.02,
        "SIVIERO EMPREENDIMENTOS": 0.02,
        "TURISMO": 0.01,
        "SIVIERO TRANSPORTES": 0.01,
    }
    total18 = 0.0
    for name, pct in share18.items():
        value = next(v for k, v in resumo18.items() if name in k)
        total18 += pct * value
        print(f"2018 {name}: {f(value)} × {pct:.0%} = {f(pct * value)}")
    paid = sum(pay[n] for n in (1, 2, 5, 6, 7, 9, 10))
    print(
        f"2018: parcela de Rovílio = {f(total18)}; itens já entregues ou pagos (1, 2, 5, 6, 7, 9, 10) = {f(paid)}; diferença {f(paid - total18)}; preço / 2018 − 1 = {(price / total18 - 1) * 100:.2f} %"
    )


if __name__ == "__main__":
    main()
