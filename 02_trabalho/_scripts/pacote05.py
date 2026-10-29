"""Packet 5: one rule for every "possible APP" case in the SEI.

The value of a lot with an environmental doubt is a share of the full lot,
set by the level of the case (decision rule under C2, Luís 28/10/2026):
  A  0 %  APP recognised on the lot by court decision, expert report or
          recorded encumbrance; or >= 50 % of the lot inside APP or a
          non-building strip on the official municipal map;
  B 50 %  10 % to 50 % of the lot inside APP or a strip on the map; or a
          pending process at the base date over the area (map layer
          "Revisão MP" with a process before 31.12.2024); or, where there
          is no map, a documented restriction on the lot (TAC);
  C 100 % less than 10 % of the lot inside APP or a strip, and no process.
For third-party lots listed as indemnifiable the liability is the
complement: A 100 %, B 50 %, C 0 % of the listed value.

Inputs: sei-2025-itens.xlsx (full-lot values), Rafael's workbooks,
geofarroupilha-app-lotes.xlsx (municipal APP map, geofarroupilha_app.py)
and planilha-mestre-sei.xlsx (adopted totals).
Usage: python pacote05.py   (prints the tables used in pacotes/pacote-05.md)
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import openpyxl

BASE = Path("C:/entrega-cisao-siviero")
W = BASE / "02_trabalho"
RAF = BASE / "01_fontes/estudo-rafael/empreendimentos-xlsx"
DATA_BASE_YEAR = 2024

SHARE = {"A": 0.0, "B": 0.5, "C": 1.0}
PCT_COL = "% do lote em APP ou faixa não edificante (união das camadas 2 a 5)"

# Cases outside the Farroupilha map, decided by documents
DOC_LEVEL = {
    "Lago Negro": (
        "A",
        "ACP 5000454-62.2016 com perícia que confirma APP nos lotes 15–21 (arq. 1 p.25); R$ 0 já aprovado",
    ),
    "Fracionamento Simioni": (
        "B",
        "Veranópolis, sem mapa municipal nas fontes; TAC com o MP (arq. 17 p.2) reconhece nascente e restrição para construir no entorno; doc 11 marca 'ÁREA DE APP'; sem perícia nem decisão por lote",
    ),
}

# Lots taken out of the third-party liability before the rule: SEI stock at
# 31.12.2024 (ACERTOS packet 1, item S1, approved by Luís on 28/10/2026).
S1 = {
    ("Jardim Itália I – II – III", "1747", "001"),
    ("Primavera", "1199", "010"),
    ("Primavera", "1199", "011"),
    ("Industrial Vicentino", "1301", "003"),
}


def rows(path: Path, sheet: str):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    head = [str(h) for h in next(it)]
    return [dict(zip(head, r)) for r in it if any(v is not None for v in r)]


def brl(v: float) -> str:
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


SENS_COL = "% do lote em APP, faixa ou sensibilidade (união das camadas 2 a 5 e 10)"
# How to read the "Sensibilidade_Ambiental" layer (30 m strip along every
# watercourse), the central choice of the packet:
#   estrita        only APP 30/50 m and non-building strips count
#   intermediaria  as estrita, but a lot in the 30 m strip is at least B
#   ampla          the 30 m strip counts as APP
READINGS = ("estrita", "intermediaria", "ampla")
READING = sys.argv[1] if len(sys.argv) > 1 else "intermediaria"
assert READING in READINGS, READING


def map_level(g: dict, reading: str = READING) -> tuple[str, str]:
    pct = g[PCT_COL] or 0.0
    pct_s = g[SENS_COL] or 0.0
    use = pct_s if reading == "ampla" else pct
    what = "APP, faixa não edificante ou faixa de 30 m" if reading == "ampla" else "APP ou faixa não edificante"
    if use >= 50:
        return "A", f"{use:.1f} % do lote em {what} (mapa municipal)"
    if use >= 10:
        return "B", f"{use:.1f} % do lote em {what} (mapa municipal)"
    if reading == "intermediaria" and pct_s >= 10:
        return "B", f"{pct:.1f} % em APP/faixa não edificante, mas {pct_s:.1f} % na faixa de 30 m (sensibilidade)"
    return "C", f"{use:.1f} % do lote em {what} (mapa municipal)"


def main() -> None:
    lotes = rows(W / "sei-2025-itens.xlsx", "Lotes")
    terc_bloco = "Lotes de terceiros passíveis de indenização"
    poss = [r for r in lotes if r["observação"] == "ÁREA DE POSSÍVEL APP"]
    terc = [r for r in lotes if str(r["bloco"]).startswith(terc_bloco)]
    simioni = [r for r in lotes if r["loteamento"] == "Fracionamento Simioni"]
    assert len(poss) == 7 and len(terc) == 26 and len(simioni) == 4

    geo = {
        (g["Loteamento"], str(g["Quadra"]), str(g["Lote"])): g
        for g in rows(W / "geofarroupilha-app-lotes.xlsx", "Lotes")
    }

    # Rafael's values, lot by lot
    def nq(s):
        return str(s).split("[")[0].strip().lstrip("0")

    def nl(s):
        return str(s).upper().replace("O", "0").lstrip("0")

    raf_exec = {}
    for r in rows(
        RAF / "LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.xlsx", "Lotes"
    ):
        m = re.search(r"R\$ ([\d.]+,\d\d) escrito ao lado", str(r["Quadra"]))
        if m:
            raf_exec[(nq(r["Quadra"]), nl(r["Lote"]))] = float(
                m.group(1).replace(".", "").replace(",", ".")
            )
    raf_terc = {
        (nq(r["Quadra"]), nl(r["Lote"])): r["Valor (R$)"]
        for r in rows(
            RAF / "LOTES DE TERCEIROS INDENIZAVEIS, BELISE LTDA.xlsx", "Lotes"
        )
    }
    raf_sim = next(
        r["Valor avaliado (R$)"]
        for r in rows(
            RAF / "LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.xlsx",
            "Avaliacao",
        )
        if "Simioni" in str(r["Descrição"])
    )
    for r in poss:
        assert raf_exec[(nq(r["quadra"]), nl(r["lote"]))] == r["valor (R$)"], r
    for r in terc:
        assert raf_terc[(nq(r["quadra"]), nl(r["lote"]))] == r["valor (R$)"], r

    resumo = {
        r["Situação"]: r["Quantidade"]
        for r in rows(W / "planilha-mestre-sei.xlsx", "Resumo")
    }
    adotado = resumo["Bens 2025 — soma dos valores adotados (R$)"]

    def level(r):
        if r["loteamento"] in DOC_LEVEL:
            return DOC_LEVEL[r["loteamento"]]
        g = geo[(r["loteamento"], str(r["quadra"]), str(r["lote"]))]
        assert g["Situação da busca"] == "ok", g
        return map_level(g)

    ativos = []  # (loteamento, quadra, lote, nível, grupo, rafael, recomendado, motivo)
    for r in poss:
        lvl, why = level(r)
        v = r["valor (R$)"]
        ativos.append(
            (r["loteamento"], r["quadra"], r["lote"], lvl, v, v, v * SHARE[lvl], why)
        )
    for r in simioni:
        lvl, why = level(r)
        v = raf_sim / len(simioni)
        ativos.append(
            (r["loteamento"], r["quadra"], r["lote"], lvl, 0.0, v, v * SHARE[lvl], why)
        )

    passivos = []  # (loteamento, quadra, lote, nível, grupo, recomendado, motivo)
    for r in terc:
        k = (r["loteamento"], r["quadra"], r["lote"])
        if k in S1:
            g = geo[(r["loteamento"], str(r["quadra"]), str(r["lote"]))]
            passivos.append(
                (
                    *k,
                    "S1",
                    r["valor (R$)"],
                    0.0,
                    f"estoque da SEI em 31.12.2024 (S1); no mapa: {map_level(g)[0]}, {map_level(g)[1]}",
                )
            )
            continue
        lvl, why = level(r)
        passivos.append(
            (*k, lvl, r["valor (R$)"], r["valor (R$)"] * (1 - SHARE[lvl]), why)
        )

    print("## Lotes próprios (ativo)\n")
    print(
        "| Loteamento | Quadra | Lote | Nível | Grupo | Rafael | Recomendado | Motivo |"
    )
    print("|---|---|---|---|---|---|---|---|")
    for a in ativos:
        print(
            f"| {a[0]} | {a[1]} | {a[2]} | {a[3]} | {brl(a[4])} | {brl(a[5])} | {brl(a[6])} | {a[7]} |"
        )
    print("\n## Lotes de terceiros (passivo)\n")
    print("| Loteamento | Quadra | Lote | Nível | Grupo | Recomendado | Motivo |")
    print("|---|---|---|---|---|---|---|")
    for p in passivos:
        print(
            f"| {p[0]} | {p[1]} | {p[2]} | {p[3]} | {brl(p[4])} | {brl(p[5])} | {p[6]} |"
        )

    # Exceptions per development and side, largest value first
    groups: dict = {}
    for lt, _q, _l, lvl, vg, vr, va, _w in ativos:
        g = groups.setdefault(("ativo", lt), [0, 0.0, 0.0, 0.0, set()])
        g[0] += 1
        g[1] += vg
        g[2] += vr
        g[3] += va
        g[4].add(lvl)
    for lt, _q, _l, lvl, vg, va, _w in passivos:
        if lvl == "S1":
            continue
        g = groups.setdefault(("passivo", lt), [0, 0.0, 0.0, 0.0, set()])
        g[0] += 1
        g[1] += vg
        g[2] += vg
        g[3] += va
        g[4].add(lvl)
    print("\n## Exceções por loteamento\n")
    print("| # | Lado | Loteamento | Lotes | Níveis | Grupo | Rafael | Recomendado |")
    print("|---|---|---|---|---|---|---|---|")
    for i, ((lado, lt), (n, vg, vr, va, lv)) in enumerate(
        sorted(groups.items(), key=lambda kv: -max(kv[1][1], kv[1][2])), 1
    ):
        print(
            f"| {i} | {lado} | {lt} | {n} | {', '.join(sorted(lv))} | {brl(vg)} | {brl(vr)} | {brl(va)} |"
        )

    at = sum(a[6] for a in ativos)
    ps = sum(p[5] for p in passivos)
    s1 = sum(p[4] for p in passivos if p[3] == "S1")
    terc_total = sum(p[4] for p in passivos)
    regra = at + (terc_total - s1 - ps)
    print("\n## Totais")
    print(
        f"- Ativos: valor cheio {brl(sum(a[5] for a in ativos))}; recomendado {brl(at)}"
    )
    print(
        f"- Passivo de terceiros: grupo {brl(terc_total)}; S1 {brl(s1)}; recomendado {brl(ps)}"
    )
    print(
        f"- Soma adotada dos bens: hoje {brl(adotado)}; com o pacote 5 {brl(adotado + at)}"
    )
    print(f"- Efeito da regra R6: + {brl(regra)}; 1/6 = {brl(regra / 6)}")
    print(f"- Efeito do S1 (já aprovado): + {brl(s1)}; 1/6 = {brl(s1 / 6)}")
    print(f"- Total: + {brl(regra + s1)}; 1/6 = {brl((regra + s1) / 6)}")
    c0 = sum(a[5] - a[6] for a in ativos)
    print(f"- Leitura: {READING}. Custo frente ao C0 nos lotes próprios (C0 daria o valor cheio): {brl(c0)}; 1/6 = {brl(c0 / 6)}")
    for nivel in "ABC":
        n_at = [a for a in ativos if a[3] == nivel]
        n_ps = [p for p in passivos if p[3] == nivel]
        print(
            f"- Nível {nivel}: {len(n_at)} lotes próprios ({brl(sum(a[5] for a in n_at))} cheio), {len(n_ps)} de terceiros ({brl(sum(p[4] for p in n_ps))} listado)"
        )
    return ativos, passivos


if __name__ == "__main__":
    main()
