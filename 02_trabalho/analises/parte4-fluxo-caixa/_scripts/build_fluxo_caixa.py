"""Part 4 (SEI) v0.1: cash-flow reconciliation of lots sold 2018-2024 against explained uses.

Reads only sources (never recalled figures) and writes:
  - fluxo-caixa-sei-v0.1.xlsx   (Portuguese workbook for the user)
  - numeros-v0.1.json           (key figures for the presentation script)

v0.1 applies the independent review in _diario/revisao-v0.md (findings A1-A8, M1-M8, B1-B5).
The v0 script is frozen in build_fluxo_caixa_v0.py.
Run: python -I _scripts/build_fluxo_caixa.py   (from the parte4-fluxo-caixa folder)
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

VERSION = "v0.2"
HERE = Path(__file__).resolve().parent.parent
ROOT = HERE.parents[2]  # C:\entrega-cisao-siviero
SRC = {
    "gsheet": ROOT
    / "01_fontes/planilhas-google/Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx",
    # final version of the classified bills (decisions of 28/10/2026, saved 18:53), with the IPCA tab
    "contas": HERE
    / "_entradas/contas-pagas-classificadas_snapshot-2026-10-28_1853.xlsx",
    "contas_lotes": ROOT
    / "01_fontes/outros/Contas_Pagas_e_Lotes_Vendidos_2023-07_a_2024-10.xlsx",
    "sei2018": ROOT / "02_trabalho/sei-2018-itens.xlsx",
    "sei2025": ROOT / "02_trabalho/sei-2025-itens.xlsx",
    "mestre": ROOT / "02_trabalho/planilha-mestre-sei.xlsx",
}
FILE17 = ROOT / "01_fontes/respostas-docs-2026-08-13"

# ---- premises (to be approved by the user) ----
PERIOD_END = 2024  # data-base 31.12.2024 (P2)
# user's decision (28/10/2026, aba Extrapolação of the classified bills): recurring costs are extrapolated
# over jan/2019-dez/2024, each month = 15-month average x IPCA(month) / IPCA(mar/2024); extraordinary costs
# are not extrapolated. The bridge uses the same period, so 2018 exits stay out of it.
BASE_START = 2019
REPORT_MONTHS = 15
IPCA_BASE = (2024, 3)  # middle of the 15-month sample
IPCA_LOTS = (
    2024,
    12,
)  # the group's 2025 lot values are brought back to each sale month from dec/2024
# reference price scenario, approved by the user on 28/10/2026 (supported by the taxes and commissions paid)
REF = "teto"
CONCESSION_MAX_LOTS = {
    "Família Falavigna": 50,
    "Colina do Retiro II": 100,
}  # v0 upper bound, now a sensitivity
COST_PER_LOT_GROUP = (
    50_000  # group's own unit cost in the 11.2.1 texts (12.100.000 / 242 lotes)
)

# group's own values for the lots that actually left, where the subdivision average misleads (review A3, M1, B1)
VALUE_OVERRIDES = {
    "Perini": {
        "v18": 95_000,
        "v25": 160_000,
        "src": "lotes pequenos: 2018 = inventário seq. 66 (6 lotes a 95.000); 2025 = GERAL 11.2.2 (lotes 1469 a 160.000)",
    },
    "Lago Negro": {
        "v18": 8_650_000 / 35,
        "v25": None,
        "src": "2018 = inventário seq. 77 (35 lotes com preço: 5 × 350.000 + 30 × 230.000)",
    },
    "Lotes urbanos (item 11.1.3)": {
        "v18": (70_000 + 100_000 + 100_000) / 3,
        "v25": None,
        "src": "2018 = GERAL C58–C60 (lotes A.2 a A.4; o A.1 saiu em 2025)",
    },
}

# 15-month categories -> class. M = keep the company running (extrapolated); INV = works on subdivisions
# (evidence, not extrapolated); DIV = principal repaid (handled by the debt change); BENS = asset purchases
# (handled by the new 2025 assets); LOTE = paid with lots (already in the lot exits); TRANSF = paid to another
# group company (not a use of the SEI); DOC = documented debt counted at its total, not extrapolated.
CLASS_EXTRA = {
    "Multas": "M",
    "Processos, acordos e serviços jurídicos pontuais": "M",
    "Distribuição anual de lucros": "M",
    "Rescisões trabalhistas": "M",
    "Despesas funerárias": "M",
    "Doações, campanhas eleitorais e eventos": "M",
    "Impostos sobre venda pontual de soja": "M",
    "ITBI, escrituras e registros de imóveis": "M",
    "Reformas, benfeitorias e despesas pontuais": "M",
    "Compra de equipamentos / sistemas": "M",
    "Empréstimos concedidos a familiares": "M",
    "Devolução de principal de empréstimos": "DIV",
    "Obras de infraestrutura de loteamentos": "INV",
    "Projetos, licenças e regularização de loteamentos": "INV",
    "Compra de imóveis / áreas / lotes": "BENS",
    "Compra e venda de máquinas e veículos": "BENS",
    "Permutas / dação em pagamento / veículos recebidos em troca": "LOTE",
    "Compra de bens pagos com lotes": "LOTE",
    "Acerto parcelado de serviços passados (Flávia Manica)": "DOC",
    "Manutenção parcelada (João Batista F. Oliveira)": "M",
    "Materiais de construção diversos (Simonetto)": "M",
    "Materiais/serviços via Antoninho Zanella (Passo Fundo)": "M",
    "Antecipação contratual (Dorval Boff)": "INV",
    "Pagamento referente a lotes (Rogério Binda)": "BENS",
}
CLASS_LABEL = {
    "M": "Manter a empresa (extrapolado)",
    "INV": "Investimento em loteamentos (evidência; não extrapolado)",
    "DIV": "Devolução de empréstimo (tratada pela variação da dívida)",
    "BENS": "Compra de bens (tratada pelos bens novos de 2025)",
    "LOTE": "Pago com lotes (já está na saída de lotes)",
    "TRANSF": "Pago a outra empresa do grupo (não é gasto da SEI)",
    "DOC": "Dívida documentada (entra pelo total, não extrapolada)",
}
# pró-labore paid to 4 of the 6 branches: Válter, Adélia, Pedro, and Raul (paid through Hotel Araçá Ltda.)
PARTNER_PROLABORE = re.compile(r"valter|adelia|pedro siviero|hotel araca")
INCLUDE_FILE17_EXTRA = True  # user's decision (28/10/2026): lots listed in file 17 beyond the reduction count

# subdivisions contracted but not finished (R$ 0 + fines in 2025) and finished ones, for tagging works
UNFINISHED = [
    "caminhos do sol",
    "vale das colinas",
    "vista da colina",
    "pedrotti",
    "andreis",
    "lorandi",
    "cecconello",
    "bonetto",
    "pasqual",
    "letti",
    "sol nascente iii",
    "champagnat",
    "santa corona",
    "alto das gracas ii",
    "altos das gracas ii",
    "congregacao",
    "sao carlos",
    "bella visione",
    "colombo",
    "simioni",
    "morada dos ipes ii",
    "ipes ii",
    "basso",
    "zanandrea",
    "fioreze",
    "boff",
    "piccoli",
    "monte berico",
    "buriti",
    "zattera",
]
FINISHED = [
    "falavigna",
    "colina do retiro",
    "medianeira",
    "pedra lisa",
    "monte veneto",
    "barbiero",
    "carnevalli",
    "lago negro",
    "san marino",
    "jardim italia",
    "altos da julieta",
    "alto das gracas",
    "altos das gracas",
    "perini",
    "dalzochio",
    "vicentino",
    "sol nascente",
    "morada do ipe",
    "morada dos ipes",
    "solar norte",
    "salvador",
    "femaca",
    "baldasso",
    "basalto",
    "primavera",
    "vicenza",
    "dalla costa",
    "nossa senhora",
]
# settlements tied to contracted subdivisions that were never built or were terminated (review M6)
CONTRACTED_TIE = re.compile(
    r"distrato|atraso|nao execu|buriti|caminhos do sol|vale das colinas|zanandrea|monte berico|colombo|zattera|"
    r"fioreze|via nova|susin|ambitec|intermediacao"
)

# Ordered rules on the ASCII-folded "forma" text (or, where flagged, the recipient's name)
EXIT_RULES = [
    ("DIVL", r"devolucao o? ?de emprestimo", False),
    (
        "SOC",
        (
            r"consignad|irmas dos socios|luis artur|vagner siviero|familia adelia|quotas de capital|"
            r"adiantamento para acerto futuro|socio"
        ),
        False,
    ),
    # lot swapped for a lot that is in the 2025 stock (review N3); Pasqual is a contracted, unbuilt subdivision
    (
        "BEM",
        (
            r"permuta de outro imovel\W*lote|permuta(do)?, com (o|um) lote|permuta de um lote (?!.*pasqual)|"
            r"area de [\d.,]+ ?m\w?\W*junto com a area da familia"
        ),  # 3.100 m2 joined to the Sapopema area (row 76)
        False,
    ),
    (
        "OBRA",
        r"paisagismo|moveis para o mobiliario",
        False,
    ),  # works of the Morada dos Ipes II club house
    ("ACORDO", r"gratifica", False),  # gratuity to an employee (review, low finding)
    (
        "VENDA",
        r"^vendid[oa]s? a (vista|prazo)\b.*(permut|parte|especie|carro|veiculo|importancia)",
        False,
    ),
    ("SOC", r"\bsiviero\b", True),
    ("ACORDO", r"distrato|indeniza|acordo judicial|processo judicial|danos", False),
    (
        "OBRA",
        (
            r"paviment|prestacao (de )?servi|servicos|tubos|aduelas|detona|licenciamento|topografia|material|pvs|"
            r"mao de obra|percentual referente|rede eletr|honorarios referente"
        ),
        False,
    ),
    ("ACORDO", r"acordo|honorario|comiss|divida|multa|dacao|doacao", False),
    (
        # goods or land received for the lot that do not appear in the 2025 inventory (review A5, N4)
        "BEMX",
        (
            r"permut|troca|aquisicao|area terras|area de terras|apartamento|permuta de uma casa|casa \(unidade|"
            r"veiculo|caminhao|camionete|carro|pavilh|land rover|moveis|instalado um bar"
        ),
        False,
    ),
    ("VENDA", r"vendid[oa]s? a (vista|prazo)|comprador", False),
]
EXIT_LABEL = {
    "VENDA": "Venda (dinheiro, inclusive com parte em bem que não está em 2025)",
    "OBRA": "Pagamento de obras e serviços de loteamento",
    "ACORDO": "Acordos, indenizações, distratos e comissões",
    "DIVL": "Empréstimos de sócios pagos com lote",
    "BEM": "Troca por outro lote que está no estoque de 2025",
    "BEMX": "Troca por bem ou área que não aparece no inventário de 2025",
    "SOC": "Sócios e família",
    "SEM": "Sem informação sobre o destino",
}
CASH_CLASSES = ("VENDA",)


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def norm(s) -> str:
    s = (
        unicodedata.normalize("NFKD", str(s or ""))
        .encode("ascii", "ignore")
        .decode()
        .lower()
    )
    return re.sub(r"\s+", " ", s).strip()


def classify_exit(forma, dest) -> str:
    # sources mix composed and decomposed accents, so match on the ASCII-folded text
    f, d = norm(forma), norm(dest)
    for key, pat, use_dest in EXIT_RULES:
        if re.search(pat, d if use_dest else f):
            return key
    return "SEM"


def sheet(wb, prefix):
    return next(ws for ws in wb.worksheets if ws.title.startswith(prefix))


# --------------------------------------------------------------------------- inputs
def load_lots(gs, stock25):
    ws = sheet(gs, "Lotes que sa")
    summary = []
    for r in ws.iter_rows(min_row=6, max_row=46, max_col=9, values_only=True):
        if r[0]:
            summary.append(
                {
                    "lot": r[0],
                    "cidade": r[1],
                    "n2018": float(r[2] or 0),
                    "n2025": float(r[3] or 0),
                    "red": float(r[4] or 0),
                    "obs": r[8],
                }
            )
    total_row = next(ws.iter_rows(min_row=47, max_row=47, max_col=8, values_only=True))
    notes = {s["lot"]: norm(s["obs"]) for s in summary}
    exits = []
    for i, r in enumerate(
        ws.iter_rows(min_row=61, max_col=9, values_only=True), start=61
    ):
        if not r[0]:
            continue
        d = r[4] if isinstance(r[4], dt.datetime) else None
        year, how = (d.year, "data da fonte") if d else (None, "")
        f = norm(r[6])
        # undated records that the sheet itself marks as leaving after dec/2024 (review A2)
        after_2024 = (
            "nao consta no inventario 2025 atualizado" in f
            or "constava em dez/2024" in f
            or (
                norm(r[2]) == "o01"
                and "entre dez/2024 e out/2025" in notes.get(r[0], "")
            )
        )
        if not d and after_2024:
            year, how = 2025, "sem data: a aba indica saída depois de dez/2024"
        exits.append(
            {
                "row": i,
                "lot": r[0],
                "lote": r[2],
                "quadra": r[3],
                "data": d,
                "ano": year,
                "como": how,
                "dest": r[5],
                "forma": r[6],
                "fonte": r[7],
                "classe": swap_check(classify_exit(r[6], r[5]), f, stock25),
                "contratado": bool(CONTRACTED_TIE.search(f + " " + norm(r[5]))),
            }
        )
    return summary, total_row, exits


def load_stock_2025(path):
    """(words of the subdivision name, quadra, lote) of every lot listed in the 2025 inventory (doc 11)."""
    ws = openpyxl.load_workbook(path, data_only=True)["Lotes"]
    out = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] and str(r[2]).isdigit() and str(r[3]).isdigit():
            out.append(
                (
                    {w for w in re.findall(r"[a-z]+", norm(r[0])) if len(w) > 3},
                    int(r[2]),
                    int(r[3]),
                )
            )
    return out


def swap_check(cls, f, stock25):
    """A lot swapped for another lot is explained only if the lot received is in the 2025 stock (review V1)."""
    m = re.search(r"lote (\d+),? quadra (\d+)\W*loteamento ([a-z ]+)", f)
    if cls != "BEM" or not m:
        return cls
    words = {w for w in re.findall(r"[a-z]+", m.group(3)) if len(w) > 3}
    found = any(
        q == int(m.group(2)) and lt == int(m.group(1)) and words & ws
        for ws, q, lt in stock25
    )
    return "BEM" if found else "BEMX"


def load_avg_values(gs):
    out, total = {}, None
    for r in sheet(gs, "Valor m").iter_rows(min_row=4, max_col=12, values_only=True):
        if r[1] and isinstance(r[0], (int, float)):
            out[r[1]] = {
                "v18": float(r[5] or 0),
                "v25": float(r[8] or 0),
                "n25v": float(r[6] or 0),
            }
        elif r[1] == "Total / média geral":
            total = {"v18": float(r[5]), "v25": float(r[8])}
    return out, total


def load_mestre_2018(mestre):
    out = {}
    for r in mestre["Loteamentos 2018x2025"].iter_rows(min_row=2, values_only=True):
        if r[1] and r[4] and r[5]:
            out[norm(r[1])] = float(r[5]) / float(r[4])
    return out


def load_geral(gs):
    return {
        c.coordinate: c.value
        for row in gs["GERAL"].iter_rows()
        for c in row
        if c.value is not None
    }


def load_costs(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    lines = []
    for name, tipo in (
        ("Recorrentes", "Recorrente"),
        ("Extraordinários", "Extraordinário"),
        ("A determinar", "A determinar"),
    ):
        if name not in wb.sheetnames:
            continue
        for r in list(wb[name].iter_rows(values_only=True))[1:]:
            if not (r[2] and str(r[2]).startswith("SIVIERO EMP")):
                continue
            cat, forn = r[9], str(r[3] or "")
            if "FLAVIA MANICA" in forn.upper():
                cls = "DOC"  # a R$ 500.000 settlement in 50 instalments: counted by what was paid, not extrapolated
            elif tipo == "Recorrente":
                # user's rule (28/10/2026): costs of other group companies and of relatives paid by the SEI are
                # recurring, so the Hotel Araçá pró-labore stays here (review M3 kept as a sensitivity)
                cls = (
                    "INV" if cat == "Combustíveis" else "M"
                )  # Luís: fuel means machines working on works
            else:
                cls = CLASS_EXTRA.get(cat)
            if cls is None:
                raise SystemExit(f"categoria sem classe: {tipo} / {cat}")
            lines.append(
                {
                    "mes": r[0],
                    "tipo": tipo,
                    "cat": cat,
                    "forn": forn,
                    "valor": float(r[7] or 0),
                    "valor_nf": r[6],
                    "obs": r[8],
                    "classe": cls,
                }
            )
    return lines


class ClampedIndex(dict):
    """Months outside the table (it runs dec/2018-dec/2024) take the nearest month in it."""

    def __missing__(self, ym):
        return self[min(self)] if ym < min(self) else self[max(self)]


def load_ipca(path):
    """IPCA index by (year, month), from the IPCA tab of the classified bills (IBGE SIDRA table 1737)."""
    ws = openpyxl.load_workbook(path, data_only=True)["IPCA"]
    return ClampedIndex(
        {
            (r[0].year, r[0].month): float(r[1])
            for r in ws.iter_rows(min_row=3, values_only=True)
            if isinstance(r[0], dt.datetime) and isinstance(r[1], (int, float))
        }
    )


def load_2018(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    cred, invest, rents = [], [], []
    for r in list(wb["Compromissos"].iter_rows(values_only=True))[1:]:
        if str(r[0]).startswith("Credores"):
            cred.append(
                {
                    "n": r[2],
                    "credor": r[3],
                    "desc": r[4],
                    "valor": r[6],
                    "sit": r[7],
                    "saldo": r[8] or 0,
                }
            )
        elif str(r[0]).startswith("Planilha"):
            invest.append({"n": r[2], "desc": r[4], "valor": r[6]})
        elif str(r[0]).startswith("Alugu"):
            rents.append({"n": r[2], "tipo": r[1], "desc": r[4], "valor": r[6]})
    items = {r[0]: r for r in list(wb["Itens"].iter_rows(values_only=True))[1:]}
    return cred, invest, rents, items


def load_cited_prices(path):
    """Single lots whose file-17 record gives the value the group attributed to the lot itself (review M7)."""
    partial = re.compile(
        r"parte em|parte permuta|atraves do veiculo|permuta (de )?(um|uma)|em especie|importancia|"
        r"\+ o valor|moveis|honorarios"
    )
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for r in list(wb["Lotes Vendidos (todos)"].iter_rows(values_only=True))[1:]:
        d, v, conj, text = r[0], r[6], r[8], r[9]
        if not (
            isinstance(d, dt.datetime)
            and 2018 <= d.year <= 2025
            and isinstance(v, (int, float))
        ):
            continue
        if conj not in (None, "Não") or partial.search(norm(text)):
            continue
        out.append(
            {
                "data": d,
                "lot": r[1],
                "municipio": r[2],
                "lote": r[3],
                "quadra": r[4],
                "valor": float(v),
                "obs": text,
            }
        )
    return out


def match_lot(name, municipio, rows):
    words = {w for w in re.findall(r"[a-z]+", norm(name)) if len(w) > 2} - {
        "loteamento",
        "condominio",
    }
    best, score = None, 0
    for r in rows:
        rw = {w for w in re.findall(r"[a-z]+", norm(r["lot"])) if len(w) > 2}
        s = len(words & ({w.rstrip("s") for w in rw} | rw)) + (
            2 if norm(r["cidade"]) == norm(municipio) else 0
        )
        if s > score:
            best, score = r, s
    return best if score >= 3 else None


# --------------------------------------------------------------------------- model
def lot_values_table(summary, avg, total_avg, v18_mestre):
    ratio = total_avg["v25"] / total_avg["v18"]
    vals, unmatched = {}, []
    for s in summary:
        name = s["lot"]
        a = avg.get(name)
        if a and a["v18"] > 0:
            v18, src18 = a["v18"], "Valor médio dos lotes (2018)"
        elif norm(name) in v18_mestre:
            v18, src18 = v18_mestre[norm(name)], "planilha-mestre, inventário 2018"
        elif norm(name) == "femaca":
            nb = next(k for k in avg if norm(k).startswith("salvador"))
            v18, src18 = avg[nb]["v18"], f"vizinho: {nb} (2018)"
        elif s["red"] <= 0:
            v18, src18 = 0.0, "sem saída: valor não usado"
        elif name in VALUE_OVERRIDES:
            v18, src18 = 0.0, ""  # set from the override just below
        else:
            v18, src18 = 0.0, "sem valor"
            unmatched.append(name)
        o = VALUE_OVERRIDES.get(name)
        if o:
            v18, src18 = o["v18"], o["src"]
        if o and o["v25"]:
            v25, src25 = o["v25"], o["src"]
        elif a and a["v25"] > 0:
            v25, src25 = a["v25"], "Valor médio dos lotes (out/2025)"
        else:
            v25, src25 = (
                v18 * ratio,
                f"índice geral {ratio:.3f} × valor 2018 (loteamento sem lote em 2025)",
            )
        vals[name] = {"v18": v18, "v25": v25, "src18": src18, "src25": src25}
    return vals, unmatched, ratio


def interp(v18, v25, year):
    if not isinstance(year, int):
        year = (BASE_START + PERIOD_END) / 2
    t = min(max((year - 2018) / 7, 0), 1)
    return v18 + (v25 - v18) * t


def allocate(summary, exits, vals, start):
    """Spread each subdivision's inventory reduction (P3) over its records, by year and class.

    Records in the period, in 2025 and undated take the reduction first; records dated before `start` take only
    what is left and stay out of the bridge (review A1, N1); reduction left after that becomes SEM in the period.
    """
    by_lot = defaultdict(list)
    for e in exits:
        by_lot[e["lot"]].append(e)
    rows = []
    for s in summary:
        name = s["lot"]
        red = max(s["red"], 0)
        listed = by_lot.get(name, [])
        pre = [e for e in listed if e["ano"] and e["ano"] < start]
        inper = [e for e in listed if e["ano"] and start <= e["ano"] <= PERIOD_END]
        post = [e for e in listed if e["ano"] and e["ano"] > PERIOD_END]
        nodate = [e for e in listed if not e["ano"]]
        usable = post + inper + nodate
        pre_n = min(len(pre), max(0.0, red - len(usable)))
        red2 = red - pre_n
        # user's decision (28/10/2026): the lots that file 17 lists beyond the inventory reduction count as SEI
        # lots that left, so every record weighs one lot; the 2018 stock is raised by the same number
        scale = (
            1.0
            if INCLUDE_FILE17_EXTRA
            else (min(1.0, red2 / len(usable)) if usable else 0.0)
        )
        extra = max(0.0, red2 - len(usable))
        cells = [(e["ano"], e["classe"], e["contratado"], scale, e) for e in usable]
        if extra:
            cells.append((None, "SEM", False, extra, None))
        rows.append(
            {
                **s,
                **vals[name],
                "listed": len(listed),
                "pre": len(pre),
                "pre_n": pre_n,
                "inper": len(inper),
                "post": len(post),
                "nodate": len(nodate),
                "extra": extra,
                "cells": cells,
                "excess": max(0.0, len(inper) - red2),
                "added_2018": max(0.0, len(usable) - red2)
                if INCLUDE_FILE17_EXTRA
                else 0.0,
            }
        )
    return rows


SCEN = ("piso", "central", "teto", "teto_nominal")


def scenario_values(rows, ipca):
    """Value of the exits per price scenario. teto = the group's 2025 value brought back by the IPCA to the
    month of the exit (same footing as the IPCA-adjusted costs); teto_nominal = the 2025 value as it stands."""
    out = {k: defaultdict(float) for k in SCEN}
    post = dict.fromkeys(("n", *SCEN), 0.0)
    pre = dict.fromkeys(("n", *SCEN), 0.0)
    by_year = defaultdict(lambda: dict.fromkeys(("n", *SCEN), 0.0))
    n_cls = Counter()

    def deflator(e, y):
        if e and e["data"]:
            ym = (e["data"].year, e["data"].month)
        elif y:
            ym = (y, 7)
        else:
            ym = ((BASE_START + PERIOD_END) // 2, 12)  # undated: middle of the period
        return ipca[ym] / ipca[IPCA_LOTS]

    def price(r, k, y, e=None):
        if k == "piso":
            return r["v18"]
        if k == "central":
            return interp(r["v18"], r["v25"], y)
        return r["v25"] * (deflator(e, y) if k == "teto" else 1.0)

    for r in rows:
        pre["n"] += r["pre_n"]
        r["val_in"] = dict.fromkeys(SCEN, 0.0)
        for k in out:
            pre[k] += r["pre_n"] * price(r, k, 2018)
        for y, cls, contr, n, e in r["cells"]:
            if y and y > PERIOD_END:
                post["n"] += n
                for k in out:
                    post[k] += n * price(r, k, y, e)
                continue
            n_cls[cls] += n
            yk = y if y else "sem data"
            by_year[yk]["n"] += n
            for k, acc in out.items():
                v = n * price(r, k, y, e)
                acc[cls] += v
                if contr and cls in ("ACORDO", "OBRA"):
                    acc["CONTRATADO"] += v
                by_year[yk][k] += v
                r["val_in"][k] += v
    return out, post, pre, by_year, n_cls


def tag_lot(text):
    t = norm(text)
    for k in UNFINISHED:
        if k in t:
            return "não finalizado", k
    for k in FINISHED:
        if k in t:
            return "finalizado/estoque", k
    return "sem loteamento identificado", ""


# --------------------------------------------------------------------------- workbook helpers
HDR = Font(bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F3864")
BOLD = Font(bold=True)
NOTE = Font(italic=True, color="595959")
MONEY = "#,##0;[Red]-#,##0"


def write_table(ws, r0, headers, rows, money_cols=(), widths=None):
    for j, h in enumerate(headers, 1):
        c = ws.cell(r0, j, h)
        c.font, c.fill = HDR, HFILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(r0 + i, j, v)
            if j in money_cols and isinstance(v, (int, float)):
                c.number_format = MONEY
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return r0 + len(rows) + 1


def para(ws, r, text, font=None):
    c = ws.cell(r, 1, text)
    if font:
        c.font = font
    return r + 1


# --------------------------------------------------------------------------- main
def main():
    hashes = {k: md5(p) for k, p in SRC.items()}
    gs = openpyxl.load_workbook(SRC["gsheet"], data_only=True)
    mestre = openpyxl.load_workbook(SRC["mestre"], data_only=True)
    geral = load_geral(gs)
    summary, total_row, exits = load_lots(gs, load_stock_2025(SRC["sei2025"]))
    avg, total_avg = load_avg_values(gs)
    vals, unmatched, ratio = lot_values_table(
        summary, avg, total_avg, load_mestre_2018(mestre)
    )
    if unmatched:
        raise SystemExit(f"loteamentos sem valor de 2018: {unmatched}")
    costs = load_costs(SRC["contas"])
    cred18, inv18, rents18, items18 = load_2018(SRC["sei2018"])
    cited = load_cited_prices(SRC["contas_lotes"])
    ipca = load_ipca(SRC["contas"])

    rows = allocate(summary, exits, vals, BASE_START)
    scen, post, pre, by_year, n_cls = scenario_values(rows, ipca)
    months = 12 * (PERIOD_END - BASE_START + 1)

    def ipca_sum(start, base):
        """Sum over the bridge months of IPCA(month) / IPCA(base): turns a monthly amount at `base` prices into
        the nominal total of the period (the method of the aba Extrapolação)."""
        return sum(
            ipca[(y, m)] / ipca[base]
            for y in range(start, PERIOD_END + 1)
            for m in range(1, 13)
        )

    # --- checks against the source totals
    assert abs(sum(r["n2018"] for r in rows) - total_row[2]) < 1e-6
    assert abs(sum(r["n2025"] for r in rows) - total_row[3]) < 1e-6
    assert abs(sum(r["red"] for r in rows) - total_row[4]) < 1e-6
    added18 = sum(r["added_2018"] for r in rows)  # lots file 17 adds to the 2018 count
    red_pos = sum(max(r["red"], 0) for r in rows) + added18
    n_in = sum(v["n"] for v in by_year.values())
    assert abs(n_in + post["n"] + pre["n"] - red_pos) < 1e-6, (
        n_in,
        post["n"],
        pre["n"],
        red_pos,
    )
    sei_15m = sum(c["valor"] for c in costs)
    sei_months = [
        r[1]
        for r in sheet(gs, "Relat").iter_rows(min_row=2, max_row=16, values_only=True)
    ]
    assert abs(sum(sei_months) - sei_15m) < 1.0

    # --- costs
    by_cls, by_cat = defaultdict(float), defaultdict(lambda: [0, 0.0, ""])
    for c in costs:
        by_cls[c["classe"]] += c["valor"]
        k = (c["tipo"], c["cat"])
        by_cat[k][0] += 1
        by_cat[k][1] += c["valor"]
        by_cat[k][2] = c["classe"]
    m_rec = sum(
        c["valor"] for c in costs if c["classe"] == "M" and c["tipo"] == "Recorrente"
    )
    m_ext = by_cls["M"] - m_rec
    m_month = by_cls["M"] / REPORT_MONTHS
    partner_pl = sum(
        c["valor"]
        for c in costs
        if c["cat"] == "Pró-labore / retiradas mensais"
        and c["classe"] == "M"
        and PARTNER_PROLABORE.search(norm(c["forn"]))
    )
    # Flávia Manica: instalments paid up to dec/2024 = last instalment number seen + months left to dec/2024
    flavia_paid = 0.0
    for c in costs:
        pm = (
            re.search(r"p\.? ?(\d+)/(\d+)", norm(c["obs"]))
            if c["classe"] == "DOC"
            else None
        )
        if pm and isinstance(c["mes"], dt.datetime):
            left = (PERIOD_END - c["mes"].year) * 12 + 12 - c["mes"].month
            n_paid = min(int(pm.group(1)) + left, int(pm.group(2)))
            flavia_paid = max(flavia_paid, n_paid * c["valor"])
    flavia_total = flavia_paid
    other_cos = sum(
        c["valor"]
        for c in costs
        if c["classe"] == "M"
        and (
            c["cat"].startswith("Despesas de familiares")
            or re.search(r"ARA[CÇ]A", c["forn"])
        )
    )
    inv_lines = [c for c in costs if c["classe"] == "INV"]
    inv_tag, inv_by_key = defaultdict(float), defaultdict(float)
    for c in inv_lines:
        c["status"], c["chave"] = tag_lot(f"{c['forn']} {c['obs']}")
        inv_tag[c["status"]] += c["valor"]
        if c["chave"]:
            inv_by_key[(c["status"], c["chave"])] += c["valor"]
    dal_ponte_cash = sum(
        c["valor"]
        for c in costs
        if c["classe"] == "BENS" and "DAL PONTE" in c["forn"].upper()
    )

    # revenue test (review A8): PIS/COFINS (cumulative regime: 0,65% and 3%) and 6% commissions vs model cash sales
    cofins = sum(
        c["valor"]
        for c in costs
        if re.match(r"COFINS\b", c["forn"].strip(), re.IGNORECASE)
    )
    pis = sum(
        c["valor"]
        for c in costs
        if re.match(r"PIS\b", c["forn"].strip(), re.IGNORECASE)
    )
    comis = by_cat[("Recorrente", "Comissões de corretagem (variável com vendas)")][1]
    win = (
        dt.date(2023, 6, 1),
        dt.date(2024, 9, 30),
    )  # payment lags the sale by about a month
    win_sales = dict.fromkeys(SCEN, 0.0)
    for r in rows:
        for y, cls, _, n, e in r["cells"]:
            if (
                e
                and e["data"]
                and win[0] <= e["data"].date() <= win[1]
                and cls in CASH_CLASSES
            ):
                win_sales["piso"] += n * r["v18"]
                win_sales["central"] += n * interp(r["v18"], r["v25"], y)
                win_sales["teto"] += (
                    n
                    * r["v25"]
                    * ipca[(e["data"].year, e["data"].month)]
                    / ipca[IPCA_LOTS]
                )
                win_sales["teto_nominal"] += n * r["v25"]
    revenue_test = {
        "cofins": cofins,
        "pis": pis,
        "receita_cofins": cofins / 0.03,
        "receita_pis": pis / 0.0065,
        "comissoes": comis,
        "vendas_comissao": comis / 0.06,
        "modelo": win_sales,
    }

    # --- debts (review A6): jan/2019 balance without non-debts, other companies' limits and loans paid with lots
    excl18 = {
        "26": "estimativa de compromissos mensais, não é dívida (já está no custo de manter)",
        "30": "limite bancário da TIS (outra empresa)",
        "32": "limite bancário do HOTEL (outra empresa)",
    }
    for c in cred18:
        n = str(c["n"]).zfill(2)
        c["fora"] = excl18.get(n) or (
            "Irineu Siviero: pago com 7 lotes em 21/12/2020 (arquivo 17)"
            if n in {"08", "09", "10", "28"}
            else ""
        )
    debt18 = sum(c["saldo"] for c in cred18 if not c["fora"])
    m = re.search(
        r"quantidade de (\d+) .*?parcelas, no valor de r\$ ([\d.]+,\d{2})",
        norm(geral["B742"]),
    )
    cat_parcels, cat_value = (
        int(m.group(1)),
        float(m.group(2).replace(".", "").replace(",", ".")),
    )
    escav_2024 = (
        cat_parcels + 7
    ) * cat_value  # parcels left on 31/07/2025, plus the 7 paid from jan to jul/2025
    debt25_doc = -float(geral["E738"])
    debt25 = debt25_doc + float(geral["E743"]) + float(geral["E742"]) + escav_2024
    debt_paid = debt18 - debt25 + flavia_total

    # --- new assets and 2018 assets gone (review M2)
    plates25 = set()
    for rr in range(5, 29):
        e = geral.get(f"E{rr}")
        pm = re.search(r"placa \w+ \w+", norm(geral.get(f"B{rr}")))
        if isinstance(e, (int, float)) and e > 0 and pm:
            plates25.add(pm.group(0))
    new_assets, new_post, gone = [], [], []
    for rr in range(5, 92):
        c, e, a = geral.get(f"C{rr}"), geral.get(f"E{rr}"), geral.get(f"A{rr}")
        if isinstance(a, str) and a.startswith("11."):
            continue  # section headers carry text codes; items carry numbers
        desc = str(geral.get(f"B{rr}") or "")
        nd = norm(desc)
        if isinstance(e, (int, float)) and e > 0 and c in (0, 0.0):
            item = {
                "linha": rr,
                "desc": desc[:170],
                "valor": float(e),
                "valor_ponte": float(e),
                "nota": "",
            }
            if re.search(r"adquirid[oa] em \w+ de 2025", nd):
                item["nota"] = "comprado depois de 31.12.2024"
                new_post.append(item)
                continue
            if rr == 54:
                item["valor_ponte"] = dal_ponte_cash
                item["nota"] = (
                    "lote Dal Ponte: pago em 10/09/2025 com 3 lotes (arquivo 17); entra só o dinheiro pago em "
                    "2024 que aparece nos relatórios"
                )
            elif rr == 76:
                item["valor_ponte"] = 0.0
                item["nota"] = (
                    "Sapopema: mesma área do item A.2 de 2018 (comentário do Luís em K88) mais 3.100 m² pagos "
                    "com o lote 06/01 do Solar Norte; não é bem novo"
                )
            new_assets.append(item)
        elif (
            isinstance(c, (int, float))
            and c > 0
            and e in (0, 0.0, None)
            and a is not None  # rows without an item number are the block's lot counts
            and not 57 <= rr <= 60
            and rr != 88
        ):
            pm = re.search(r"placa \w+ \w+", nd)
            if pm and pm.group(0) in plates25:
                continue  # the same vehicle is still listed in 2025 in another row
            if (
                re.search(r"vendido|troca", nd) or rr == 65
            ):  # row 65: office furniture, small write-off
                kind = "com destino declarado"
            elif re.search(r"(pedro|valter|joao) (siviero|lunelli)\.?$", nd):
                kind = "em nome de pessoa (verificar)"
            else:
                kind = "sem destino documentado"
            gone.append(
                {"linha": rr, "desc": desc[:170], "valor": float(c), "tipo": kind}
            )
    new_assets_total = sum(a["valor_ponte"] for a in new_assets)
    gone_sem = sum(g["valor"] for g in gone if g["tipo"] == "sem destino documentado")
    rent_month = (
        sum(r["valor"] for r in rents18 if "mensal" in r["tipo"])
        + sum(r["valor"] for r in rents18 if "anual" in r["tipo"]) / 12
    )

    # concession (review A4): only what the group's own numbers document
    fal_to_finish = float(
        items18[111][15]
    )  # "CUSTO PARA CONCLUIR" in the 2018 inventory
    fal_left_2025 = -float(geral["E737"])
    # Falavigna works already in the extrapolated recurring costs (user's rule: works on finished subdivisions are
    # recurring) are part of the same documented amount, so they come off the concession (review V4)
    fal_in_c1 = (
        sum(
            c["valor"]
            for c in costs
            if c["classe"] == "M"
            and c["tipo"] == "Recorrente"
            and "falavigna" in norm(f"{c['forn']} {c['obs']}")
        )
        / REPORT_MONTHS
        * ipca_sum(BASE_START, IPCA_BASE)
    )
    concession = max(0.0, fal_to_finish - fal_left_2025 - fal_in_c1)
    concession_max = COST_PER_LOT_GROUP * sum(CONCESSION_MAX_LOTS.values())

    def bridge_for(v, start, conc):
        # recurring costs: 15-month average at mar/2024 prices, carried to each month by the IPCA (user's method);
        # extraordinary "keep the company" items: counted once, as paid in the 15 months (not extrapolated)
        rec = m_rec / REPORT_MONTHS * ipca_sum(start, IPCA_BASE)
        manter = rec + m_ext
        rents = rent_month * ipca_sum(
            start, (2018, 12)
        )  # 2018 rent roll carried by the IPCA
        cash = sum(v[c] for c in CASH_CLASSES)
        uses = manter + debt_paid + new_assets_total + conc
        gap = cash - uses + rents
        unexplained = gap + v["SEM"] + v["BEMX"] + gone_sem
        return {
            "total": sum(v[c] for c in EXIT_LABEL),
            "VENDA": cash,
            "OBRA": v["OBRA"],
            "ACORDO": v["ACORDO"],
            "DIVL": v["DIVL"],
            "BEM": v["BEM"],
            "BEMX": v["BEMX"],
            "SOC": v["SOC"],
            "SEM": v["SEM"],
            "contratado": v["CONTRATADO"],
            "manter": manter,
            "manter_rec": rec,
            "manter_ext": m_ext,
            "divida": debt_paid,
            "bens_novos": new_assets_total,
            "conclusao": conc,
            "usos": uses,
            "saldo_lotes": cash - uses,
            "alugueis": rents,
            "saldo_com_alugueis": gap,
            "bens_2018_sem_destino": gone_sem,
            "sem_explicacao": unexplained,
            "um_sexto": unexplained / 6,
            "investimento_a_demonstrar": unexplained + v["OBRA"] + v["CONTRATADO"],
        }

    bridge = {k: bridge_for(scen[k], BASE_START, concession) for k in scen}

    # sensitivity: a 2018 start (84 months, 2018 exits inside) against the 2019 base (review M4)
    scen18 = scenario_values(allocate(summary, exits, vals, 2018), ipca)[0]
    sens = []
    for start, sc in ((2019, scen), (2018, scen18)):
        mo = 12 * (PERIOD_END - start + 1)
        for conc_label, conc in (
            ("documentada", concession),
            ("máxima (v0)", concession_max),
        ):
            sens.append(
                {
                    "inicio": start,
                    "meses": mo,
                    "concessao": conc_label,
                    **{k: bridge_for(sc[k], start, conc)["sem_explicacao"] for k in sc},
                }
            )
    excess_n = sum(r["excess"] for r in rows)
    excess_val = {
        "piso": sum(r["excess"] * r["v18"] for r in rows),
        "central": sum(r["excess"] * (r["v18"] + r["v25"]) / 2 for r in rows),
        "teto": sum(r["excess"] * r["v25"] for r in rows),
    }
    f17 = openpyxl.load_workbook(next(FILE17.glob("17*.xlsx")), data_only=True)
    tot17 = [
        {"loteamento": r[0], "texto": r[3]}
        for r in f17["Totais por loteamento"].iter_rows(min_row=2, values_only=True)
        if r[0] and r[3] and "2018" in str(r[3])
    ]

    # --- yearly stock path at 2025 prices (chart); undated and pre-2018 exits spread over the bridge years
    stock_n = sum(r["n2018"] + r["added_2018"] for r in rows)
    stock_v = sum((r["n2018"] + r["added_2018"]) * r["v25"] for r in rows)
    label18 = (
        "fim 2018 (inventário + arquivo 17)" if added18 else "fim 2018 (inventário)"
    )
    path = [{"ano": label18, "lotes": stock_n, "valor25": stock_v}]
    spread_n = sum(n for r in rows for y, _, _, n, _ in r["cells"] if not y) + pre["n"]
    spread_v = sum(
        n * r["v25"] for r in rows for y, _, _, n, _ in r["cells"] if not y
    ) + sum(r["pre_n"] * r["v25"] for r in rows)
    nyears = PERIOD_END - BASE_START + 1
    for y in range(BASE_START, 2026):
        n_y = sum(
            n
            for r in rows
            for yy, _, _, n, _ in r["cells"]
            if yy == y or (y == 2025 and yy and yy > 2025)
        )
        v_y = sum(
            n * r["v25"]
            for r in rows
            for yy, _, _, n, _ in r["cells"]
            if yy == y or (y == 2025 and yy and yy > 2025)
        )
        if y <= PERIOD_END:
            n_y += spread_n / nyears
            v_y += spread_v / nyears
        stock_n -= n_y
        stock_v -= v_y
        path.append(
            {
                "ano": f"fim {y}" if y <= PERIOD_END else "out/2025",
                "lotes": stock_n,
                "valor25": stock_v,
                "saidas": n_y,
            }
        )
    expected_end = sum(
        r["n2018"] - max(r["red"], 0) for r in rows
    )  # the added lots go in and out
    assert abs(path[-1]["lotes"] - expected_end) < 1e-6, (path[-1], expected_end)

    # --- 11.2 comparison
    sec = [
        (
            "11.2.1",
            "Áreas contratadas: loteamentos em licenciamento/execução (multas rescisórias)",
            "E666",
            "F666",
        ),
        (
            "11.2.2",
            "Lotes de terceiros passíveis de indenização (Farroupilha)",
            "E677",
            "F677",
        ),
        ("11.2.3", "Obras finais a pagar (Colina do Retiro II)", "E719", "F719"),
        ("11.2.4", "Demais compromissos (itens 1 a 15 e 19)", "E721", "F721"),
        ("11.2.5", "Empréstimos e financiamentos", "E738", None),
    ]
    comp = [
        {
            "codigo": code,
            "item": label,
            "siviero": float(geral[ce]),
            "rafael": float(geral[cf])
            if cf
            else 0.0,  # Rafael's 11.2.4 (F721) already includes 11.2.5
            "comentario_luis": str(geral.get("K" + ce[1:], "") or ""),
        }
        for code, label, ce, cf in sec
    ]
    tot_siv, tot_raf = float(geral["E664"]), float(geral["F664"])
    assert abs(sum(c["siviero"] for c in comp) - tot_siv) < 1.0
    assert abs(sum(c["rafael"] for c in comp) - tot_raf) < 1.0
    # user's rule (28/10/2026): only future charges may stay in 11.2. Anything already paid, with lots, money or other
    # goods, is zeroed: the asset already left the inventory or the payment already left the cash. The abatement is
    # what was paid, capped at the balance charged (cases from cruzamento-compromissos-arq17.xlsx)
    cross = openpyxl.load_workbook(
        HERE / "cruzamento-compromissos-arq17.xlsx", data_only=True
    )["Avaliação"]
    cases = {
        r[0]: r
        for r in cross.iter_rows(min_row=2, values_only=True)
        if r[0] and str(r[0]).startswith("D")
    }
    abate = [  # (11.2 code, case id, abatement, reason)
        (
            "11.2.2",
            "D1",
            cases["D1"][6],
            "4 lotes da SEI que saíram em 2025 descontados como lotes de terceiros",
        ),
        (
            "11.2.3",
            "D4",
            min(-cases["D4"][2], cases["D4"][6]),
            "pavimentação do Colina do Retiro II já paga com lote",
        ),
        (
            "11.2.4",
            "D3",
            min(-cases["D3"][2], cases["D3"][6]),
            "obras do Falavigna já pagas com 5 lotes",
        ),
        (
            "11.2.4",
            "D5",
            -cases["D5"][2],
            "comissão Buriti de Susin: lotes do Termo A já entregues",
        ),
        (
            "11.2.4",
            "D2",
            -cases["D2"][2],
            "Solar Serra: lote 02/07 já entregue em abr/2024 (a dívida era do Hotel Araçá)",
        ),
    ]
    for c in comp:
        c["ajustes"] = [
            {"caso": cid, "valor": float(v), "motivo": why}
            for code, cid, v, why in abate
            if code == c["codigo"]
        ]
        c["proposta"] = (
            0.0
            if c["codigo"] == "11.2.1"
            else c["siviero"] + sum(a["valor"] for a in c["ajustes"])
        )

    cited_rows = []
    for c in cited:
        hit = match_lot(c["lot"], c["municipio"], rows)
        cited_rows.append(
            {
                **c,
                "v18": hit["v18"] if hit else None,
                "v25": hit["v25"] if hit else None,
                "lot_model": hit["lot"] if hit else None,
                "rel25": c["valor"] / hit["v25"] if hit and hit["v25"] else None,
            }
        )

    # ------------------------------------------------------------------ workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leia-me"
    ws.column_dimensions["A"].width = 150
    r = 1
    for line, f in [
        (
            f"Parte 4 da SEI — fluxo de caixa 2018–2024 (versão {VERSION}, 28/10/2026)",
            Font(bold=True, size=14),
        ),
        (
            (
                "Situação: RASCUNHO para apresentação. Não aprovado pelo Luís. Revisado por revisor independente "
                "(_diario/revisao-v0.md e revisao-v0.1.md); correções e mudanças de método na aba Pendências."
            ),
            BOLD,
        ),
        ("", None),
        (
            (
                "Pergunta: o valor dos lotes que saíram de 2018 a 2024 é explicado pelo que a empresa gastou para se manter, "
                "pelas dívidas que pagou e pelos bens que comprou?"
            ),
            None,
        ),
        (
            (
                "O que sobra, se a empresa diz não ter caixa, foi investido em algo — provavelmente nos loteamentos não "
                "finalizados, que o grupo avalia em R$ 0 e ainda desconta multas."
            ),
            None,
        ),
        ("", None),
        (
            (
                "Como ler: Ponte (resultado) → Teste de receita (qual preço usar) → Sensibilidade → abas 1 a 8 → Contratos "
                "11.2 → Pendências."
            ),
            None,
        ),
        (
            (
                "Preço dos lotes: piso = valor 2018 do inventário; teto = valor 2025 do grupo trazido pelo IPCA ao mês da "
                "venda (mesma base dos custos); central = interpolação pelo ano da saída. Cenário de referência: "
                f"{REF} (aba Teste de receita)."
            ),
            None,
        ),
        (
            (
                f"Período: {BASE_START}–{PERIOD_END} ({months} meses), o mesmo da extrapolação de custos decidida pelo "
                "Luís. Saídas de 2025 ficam fora (P2); saídas registradas até 2018 só contam se sobrar redução do "
                "inventário sem outro registro."
            ),
            None,
        ),
        (
            "Nenhum valor foi reavaliado (P5): todo preço é do grupo (inventário 2018, doc 11 ou o próprio passivo 11.2.2).",
            None,
        ),
        ("", None),
        ("Fontes (MD5):", BOLD),
    ]:
        r = para(ws, r, line, f)
    for k, p in SRC.items():
        r = para(ws, r, f"  {p.relative_to(ROOT)}  —  {hashes[k]}")
    para(ws, r + 1, "Gerado por _scripts/build_fluxo_caixa.py. Não editar à mão.", NOTE)

    # Ponte
    ws = wb.create_sheet("Ponte")
    lines = [
        ("A", "Valor dos lotes que saíram em 2018–2024", "total", "Abas 1 e 2"),
        (
            "A1",
            "   vendidos em dinheiro (inclui vendas com parte em bem que não aparece em 2025)",
            "VENDA",
            "entra caixa",
        ),
        (
            "A2",
            "   dados em pagamento de obras e serviços de loteamento",
            "OBRA",
            "investimento pago com lote",
        ),
        (
            "A3",
            "   dados em acordos, indenizações, distratos e comissões",
            "ACORDO",
            "dívida quitada com lote",
        ),
        (
            "A3c",
            "      dos quais ligados a loteamentos contratados ou distratados (A2 + A3)",
            "contratado",
            "gasto com loteamentos que não deram certo",
        ),
        (
            "A4",
            "   empréstimos de sócios pagos com lote",
            "DIVL",
            "dívida quitada com lote",
        ),
        (
            "A5",
            "   trocados por outro lote que está no estoque de 2025",
            "BEM",
            "destino explicado",
        ),
        (
            "A5x",
            "   trocados por bem ou área que não aparece no inventário de 2025",
            "BEMX",
            "sem destino: o bem recebido sumiu (revisão A5/N4)",
        ),
        (
            "A6",
            "   entregues a sócios e família",
            "SOC",
            "distribuição: verificar a compensação entre os ramos",
        ),
        (
            "A7",
            "   sem informação sobre o destino",
            "SEM",
            "P4: sem destino documentado",
        ),
        ("", "", None, ""),
        ("B", "Dinheiro que entrou com a venda de lotes (= A1)", "VENDA", ""),
        (
            "C1",
            f"(−) Manter a empresa: recorrentes, média dos 15 meses levada a cada um dos {months} meses pelo IPCA",
            "manter_rec",
            "método decidido pelo Luís (aba Extrapolação das contas pagas)",
        ),
        (
            "C1b",
            "(−) Manter a empresa: multas, acordos, lucros e afins pagos nos 15 meses (uma vez, sem extrapolar)",
            "manter_ext",
            "extraordinários não são extrapolados (mesma decisão)",
        ),
        (
            "C2",
            "(−) Dívidas de 2018 pagas em dinheiro (saldo jan/2019 − dívida em 31.12.2024 + parcelas Flávia até dez/2024)",
            "divida",
            "Aba 5",
        ),
        (
            "C3",
            "(−) Bens comprados no período que estão no inventário de 2025",
            "bens_novos",
            "Aba 6",
        ),
        (
            "C4",
            "(−) Conclusão de loteamentos documentada (Falavigna: a concluir em 2018 − saldo em 2025)",
            "conclusao",
            "aba Sensibilidade traz a concessão máxima da v0",
        ),
        ("C", "Total dos usos explicados", "usos", ""),
        ("D", "Saldo do dinheiro dos lotes não explicado (B − C)", "saldo_lotes", ""),
        (
            "E",
            f"(+) Outras entradas da SEI mostradas à parte: aluguéis (relação de 2018, {months} meses, corrigida pelo IPCA)",
            "alugueis",
            "Aba 7",
        ),
        ("F", "Saldo com outras entradas (D + E)", "saldo_com_alugueis", ""),
        ("G1", "(+) Lotes sem destino documentado (A7)", "SEM", ""),
        (
            "G2",
            "(+) Bens de 2018 que sumiram sem destino documentado",
            "bens_2018_sem_destino",
            "Aba 6",
        ),
        (
            "G3",
            "(+) Bens recebidos por lotes que não aparecem em 2025 (A5x)",
            "BEMX",
            "",
        ),
        ("H", "VALOR SEM EXPLICAÇÃO (F + G1 + G2 + G3)", "sem_explicacao", ""),
        ("H/6", "Parcela de 1/6", "um_sexto", ""),
        ("", "", None, ""),
        (
            "I",
            "INVESTIMENTO A DEMONSTRAR NOS LOTEAMENTOS (H + A2 + A3c)",
            "investimento_a_demonstrar",
            "dinheiro sem destino + lotes dados em obras e em loteamentos contratados",
        ),
    ]
    data = [
        [code, label]
        + (
            [bridge[s][key] for s in ("piso", "central", "teto")]
            if key
            else ["", "", ""]
        )
        + [note]
        for code, label, key, note in lines
    ]
    end = write_table(
        ws,
        1,
        [
            "Linha",
            "Descrição",
            "Piso (valor 2018)",
            "Central (interpolado)",
            "Teto (valor 2025 do grupo, trazido pelo IPCA à data da venda) — referência",
            "Observação",
        ],
        data,
        money_cols=(3, 4, 5),
        widths=[7, 84, 18, 18, 22, 60],
    )
    for rr in range(2, end + 1):
        if ws.cell(rr, 1).value in ("A", "B", "C", "D", "F", "H", "H/6", "I"):
            for j in range(1, 6):
                ws.cell(rr, j).font = BOLD
    r = end + 1
    for line in [
        "H é o valor que a empresa recebeu (ou deixou de ter) e que nenhum documento recebido explica.",
        (
            "Se o grupo diz que esse dinheiro foi investido nos loteamentos não finalizados, esse investimento existe e tem "
            "de aparecer no patrimônio, em vez de R$ 0 + multas."
        ),
        "Ou o grupo apresenta os valores reais e a diferença é recalculada, ou 1/6 de H deve ser devolvido.",
        (
            "Preço de referência = teto: os impostos (PIS/COFINS) e as comissões que a SEI pagou nos 15 meses indicam "
            "vendas iguais ou maiores que o teto (aba Teste de receita)."
        ),
    ]:
        r = para(ws, r, line, NOTE)

    # Teste de receita
    ws = wb.create_sheet("Teste de receita")
    rt = revenue_test
    data = [
        [
            "COFINS pago pela SEI nos 15 meses (regime cumulativo, 3%)",
            rt["cofins"],
            rt["receita_cofins"],
        ],
        ["PIS pago pela SEI nos 15 meses (0,65%)", rt["pis"], rt["receita_pis"]],
        [
            "Comissões de corretagem pagas (6% sobre a venda, item 12 das observações do grupo)",
            rt["comissoes"],
            rt["vendas_comissao"],
        ],
    ]
    end = write_table(
        ws,
        1,
        ["Indicador", "Valor pago", "Receita/venda implícita"],
        data,
        money_cols=(2, 3),
        widths=[80, 18, 22],
    )
    data = [
        [f"Lotes vendidos em dinheiro de jun/2023 a set/2024 — {k}", v]
        for k, v in rt["modelo"].items()
    ]
    end = write_table(
        ws, end + 1, ["Mesmo período no modelo", "Valor"], data, money_cols=(2,)
    )
    for line in [
        (
            "Leitura: os tributos e as comissões já estão no custo de manter, no nível real. Um cenário que avalia as vendas "
            "abaixo da receita implícita é incoerente com os custos que ele mesmo usa."
        ),
        (
            "Ressalvas: a base de PIS/COFINS inclui aluguéis (cerca de R$ 35 mil/mês), serviços, parcelas de vendas antigas, "
            "permutas e dações. O teste indica a ordem de grandeza, não o valor exato."
        ),
        (
            "A amostra de valores citados no arquivo 17 (aba 8) não serve como prova de preço: na maioria dos registros o "
            "valor é só a parte em dinheiro ou em bem (revisão, achado M7)."
        ),
    ]:
        end = para(ws, end + 1, line, NOTE)

    # Sensibilidade
    ws = wb.create_sheet("Sensibilidade")
    data = [
        [s["inicio"], s["meses"], s["concessao"], s["piso"], s["central"], s["teto"]]
        for s in sens
    ]
    end = write_table(
        ws,
        1,
        [
            "Início",
            "Meses",
            "Concessão de conclusão",
            "Sem explicação — piso",
            "central",
            "teto",
        ],
        data,
        money_cols=(4, 5, 6),
        widths=[8, 8, 22, 22, 18, 18],
    )
    r = (
        para(
            ws,
            end + 1,
            f"Base: início {BASE_START}, concessão documentada (R$ {concession:,.2f}). Concessão máxima da "
            f"v0: 150 lotes × R$ 50.000 = R$ {concession_max:,.2f}, mas o grupo ainda lista R$ 2,0 mi "
            "(Falavigna) e R$ 3,8 mi (Colina do Retiro II) a pagar em 2025: usar as duas coisas conta o "
            "mesmo gasto duas vezes.",
            NOTE,
        )
        + 1
    )
    data = [
        [
            "Lotes listados no arquivo 17 como saídos em 2018–2024 além da redução do inventário",
            excess_n,
            excess_val["piso"],
            excess_val["central"],
            excess_val["teto"],
        ]
    ]
    end = write_table(
        ws,
        r,
        [
            "Lotes do arquivo 17 além da redução do inventário: INCLUÍDOS na base (decisão do Luís, 28/10/2026)",
            "Lotes",
            "piso",
            "central",
            "teto",
        ],
        data,
        money_cols=(3, 4, 5),
    )
    r = para(
        ws,
        end + 1,
        "O próprio arquivo 17 declara quantidades de 2018 maiores que as do inventário de 2018 "
        "(lista abaixo). Por decisão do Luís, esses lotes contam como estoque da SEI vendido no período (valores ao lado: preço sem IPCA, só referência).",
        NOTE,
    )
    end = write_table(
        ws,
        r + 1,
        ["Loteamento (arquivo 17)", "Totais do PDF"],
        [[t["loteamento"], t["texto"]] for t in tot17],
    )
    data = [
        [
            "Pró-labore de 4 dos 6 ramos (Válter, Adélia, Pedro e Raul, este pago ao Hotel Araçá): custo da empresa (decisão de 28/10)",
            partner_pl,
            partner_pl / REPORT_MONTHS * ipca_sum(BASE_START, IPCA_BASE),
        ],
        [
            (
                "Despesas de familiares e de outras empresas (inclui pró-labore ao Hotel Araçá), mantidas no custo pela "
                "regra de 28/10"
            ),
            other_cos,
            other_cos / REPORT_MONTHS * ipca_sum(BASE_START, IPCA_BASE),
        ],
    ]
    write_table(
        ws,
        end + 1,
        ["Itens do custo de manter (referência)", "15 meses", f"× {months} meses"],
        data,
        money_cols=(2, 3),
    )

    # 1 Lotes por loteamento
    ws = wb.create_sheet("1 Lotes por loteamento")
    hdr = [
        "Loteamento",
        "Cidade",
        "Lotes 2018",
        "Lotes out/2025",
        "Redução",
        "Valor/lote 2018",
        "Fonte 2018",
        "Valor/lote 2025",
        "Fonte 2025",
        "Registros",
        "antes do período",
        "redução usada por eles",
        f"{BASE_START}–{PERIOD_END}",
        "2025",
        "sem data",
        "Redução sem registro",
        "Registros além da redução",
        "Valor 2018–2024 piso",
        "central",
        "teto",
        "Observação da aba",
    ]
    data = [
        [
            x["lot"],
            x["cidade"],
            x["n2018"],
            x["n2025"],
            x["red"],
            x["v18"],
            x["src18"],
            x["v25"],
            x["src25"],
            x["listed"],
            x["pre"],
            x["pre_n"],
            x["inper"],
            x["post"],
            x["nodate"],
            x["extra"],
            x["excess"],
            x["val_in"]["piso"],
            x["val_in"]["central"],
            x["val_in"]["teto"],
            x["obs"],
        ]
        for x in rows
    ]
    end = write_table(
        ws,
        1,
        hdr,
        data,
        money_cols=(6, 8, 18, 19, 20),
        widths=[
            30,
            16,
            9,
            9,
            9,
            13,
            34,
            13,
            34,
            9,
            9,
            10,
            9,
            7,
            8,
            10,
            10,
            15,
            15,
            15,
            60,
        ],
    )
    for j, v in (
        (1, "Total"),
        (3, sum(x["n2018"] for x in rows)),
        (4, sum(x["n2025"] for x in rows)),
        (5, sum(x["red"] for x in rows)),
    ):
        ws.cell(end, j, v).font = BOLD
    for j, k in ((18, "piso"), (19, "central"), (20, "teto")):
        c = ws.cell(end, j, sum(x["val_in"][k] for x in rows))
        c.font, c.number_format = BOLD, MONEY
    para(
        ws,
        end + 2,
        "Fonte: abas 'Lotes que saíram 2018-2025' e 'Valor médio dos lotes' da planilha Google. A redução "
        "2018→2025 manda (P3); os registros dão ano e forma. Quantidades fracionadas = redução distribuída "
        "pelos registros.",
        NOTE,
    )
    ws.freeze_panes = "B2"

    # 2 Saídas por ano e forma
    ws = wb.create_sheet("2 Saídas por ano e forma")
    keys = sorted(k for k in by_year if isinstance(k, int)) + [
        k for k in by_year if not isinstance(k, int)
    ]
    data = [
        [
            k,
            by_year[k]["n"],
            by_year[k]["piso"],
            by_year[k]["central"],
            by_year[k]["teto"],
        ]
        for k in keys
    ]
    data.append(
        [
            "2025 (após a data-base)",
            post["n"],
            post["piso"],
            post["central"],
            post["teto"],
        ]
    )
    data.append(
        [
            f"antes de {BASE_START} (vendas registradas que ainda estavam na contagem de 2018)",
            pre["n"],
            pre["piso"],
            pre["central"],
            pre["teto"],
        ]
    )
    end = write_table(
        ws,
        1,
        ["Ano da saída", "Lotes", "Valor piso", "Valor central", "Valor teto"],
        data,
        money_cols=(3, 4, 5),
        widths=[34, 10, 18, 18, 18, 40, 70, 34, 30, 10],
    )
    r = (
        para(
            ws,
            end + 1,
            "Lotes saídos em 2025 ainda eram da empresa em 31.12.2024: candidatos ao registro de "
            "data-base (voltam ao patrimônio de 31.12.2024 ou o grupo informa o que recebeu).",
            NOTE,
        )
        + 1
    )
    data = [
        [EXIT_LABEL[k], n_cls[k], scen["piso"][k], scen["central"][k], scen["teto"][k]]
        for k in EXIT_LABEL
    ]
    end = write_table(
        ws,
        r,
        ["Forma de saída (2018–2024)", "Lotes", "Piso", "Central", "Teto"],
        data,
        money_cols=(3, 4, 5),
    )
    r = (
        para(
            ws,
            end + 1,
            "Classificação por palavras-chave do campo 'Forma / observação da fonte' (arquivo 17). Lista "
            "lote a lote abaixo.",
            NOTE,
        )
        + 1
    )
    data = [
        [
            e["lot"],
            e["lote"],
            e["quadra"],
            e["data"].date() if e["data"] else "",
            e["ano"] or "",
            e["como"],
            e["dest"],
            e["forma"],
            EXIT_LABEL[e["classe"]],
            "sim" if e["contratado"] else "",
            e["fonte"],
        ]
        for e in exits
    ]
    write_table(
        ws,
        r,
        [
            "Loteamento",
            "Lote",
            "Quadra",
            "Data",
            "Ano usado",
            "Como o ano foi obtido",
            "Destinatário",
            "Forma (fonte)",
            "Classe",
            "Ligado a loteamento contratado?",
            "Fonte",
        ],
        data,
    )

    # 3 Gastos 15 meses
    ws = wb.create_sheet("3 Gastos 15 meses (SEI)")
    data = [
        [CLASS_LABEL[cls], tipo, cat, n, v, v / REPORT_MONTHS]
        for (tipo, cat), (n, v, cls) in sorted(
            by_cat.items(), key=lambda kv: (kv[1][2], -kv[1][1])
        )
    ]
    end = write_table(
        ws,
        1,
        [
            "Classe nesta análise",
            "Tipo na planilha classificada",
            "Categoria",
            "Lançamentos",
            "Total 15 meses",
            "Média mensal",
        ],
        data,
        money_cols=(5, 6),
        widths=[52, 22, 60, 12, 16, 14],
    )
    r = end + 1
    for cls, label in CLASS_LABEL.items():
        ws.cell(r, 1, label).font = BOLD
        ws.cell(r, 5, by_cls[cls]).number_format = MONEY
        ws.cell(r, 6, by_cls[cls] / REPORT_MONTHS).number_format = MONEY
        r += 1
    ws.cell(
        r,
        1,
        "Total SEI nos 15 meses (confere com a aba Relatórios Mensais da planilha Google)",
    ).font = BOLD
    ws.cell(r, 5, sei_15m).number_format = MONEY
    r += 2
    for line in [
        (
            f"Recorrentes: R$ {m_rec:,.2f} em 15 meses = R$ {m_rec / REPORT_MONTHS:,.2f}/mês a preços de mar/2024; × soma "
            f"dos fatores IPCA de {BASE_START} a {PERIOD_END} ({ipca_sum(BASE_START, IPCA_BASE):.2f}). Extraordinários de "
            f"manter (multas, acordos, lucros…): R$ {m_ext:,.2f}, contados uma vez."
        ),
        (
            "Regras do Luís: combustível não é recorrente (indica máquinas em obra) → Investimento; obras em loteamento "
            "finalizado, loteamento rescindido e despesas de familiares e outras empresas pagas pela SEI → recorrentes."
        ),
        (
            f"Acerto da Flávia Manica: entra pelas parcelas pagas até dez/2024 (R$ {flavia_total:,.2f}), na linha C2, em vez "
            "de extrapolado."
        ),
        "Devolução de principal não é extrapolada: a dívida é tratada pela variação do saldo (aba 5).",
        "Fonte: Contas_Pagas_Recorrentes_x_Extraordinarias_..._revisado.xlsx (versão final de 28/10/2026, 18:53).",
    ]:
        r = para(ws, r, line, NOTE)

    # 4 Investimentos 15 meses
    ws = wb.create_sheet("4 Investimentos 15 meses")
    data = [
        [s, k, v] for (s, k), v in sorted(inv_by_key.items(), key=lambda kv: -kv[1])
    ]
    data.append(
        ["sem loteamento identificado", "", inv_tag["sem loteamento identificado"]]
    )
    end = write_table(
        ws,
        1,
        ["Situação do loteamento", "Palavra encontrada", "Total 15 meses"],
        data,
        money_cols=(3,),
        widths=[30, 26, 16, 22, 40, 14, 60],
    )
    r = end + 1
    for s in ("não finalizado", "finalizado/estoque", "sem loteamento identificado"):
        ws.cell(r, 1, f"Total {s}").font = BOLD
        ws.cell(r, 3, inv_tag[s]).number_format = MONEY
        r += 1
    data = [
        [
            c["mes"].date() if isinstance(c["mes"], dt.datetime) else c["mes"],
            c["cat"],
            c["forn"],
            c["obs"],
            c["valor"],
            c["status"],
            c["chave"],
        ]
        for c in sorted(inv_lines, key=lambda c: -c["valor"])
    ]
    write_table(
        ws,
        r + 1,
        [
            "Mês",
            "Categoria",
            "Fornecedor",
            "Observação",
            "Valor",
            "Situação",
            "Palavra",
        ],
        data,
        money_cols=(5,),
    )

    # 5 Dívidas
    ws = wb.create_sheet("5 Dívidas 2018 x 2025")
    data = [
        [
            c["n"],
            c["credor"],
            str(c["desc"])[:140],
            c["valor"],
            c["sit"],
            c["saldo"],
            c["fora"],
        ]
        for c in cred18
    ]
    end = write_table(
        ws,
        1,
        [
            "Nº",
            "Credor",
            "Descrição (inventário 2018)",
            "Valor",
            "Situação",
            "Saldo em aberto",
            "Fora da conta porque",
        ],
        data,
        money_cols=(4, 6),
        widths=[6, 34, 90, 14, 18, 16, 60],
    )
    r = end
    for label, v in [
        ("Saldo das dívidas da SEI em jan/2019 (sem os itens fora da conta)", debt18),
        ("Empréstimos e financiamentos no doc 11 (11.2.5)", debt25_doc),
        ("(−) trator D6N, comprado em set/2025", float(geral["E743"])),
        (
            f"(±) escavadeira na posição de 31.12.2024: ({cat_parcels} + 7) parcelas × R$ {cat_value:,.2f}",
            escav_2024 + float(geral["E742"]),
        ),
        ("Dívida em 31.12.2024", debt25),
        (f"(+) Acerto Flávia Manica (acordo de R$ {flavia_total:,.2f})", flavia_total),
        ("Dívida paga em dinheiro no período (linha C2)", debt_paid),
    ]:
        ws.cell(r, 2, label).font = BOLD
        ws.cell(r, 6, v).number_format = MONEY
        r += 1
    para(
        ws,
        r + 1,
        "Banco SICREDI (item 24, R$ 5.400.000, a maior parte do saldo): a descrição não diz de qual empresa "
        "é a conta. Pedir o extrato.",
        NOTE,
    )

    # 6 Bens novos e baixados
    ws = wb.create_sheet("6 Bens novos e baixados")
    data = [
        [a["linha"], a["desc"], a["valor"], a["valor_ponte"], a["nota"]]
        for a in new_assets
    ]
    end = write_table(
        ws,
        1,
        [
            "Linha GERAL",
            "Bem novo (inventário 2025)",
            "Valor 2025",
            "Entra na ponte",
            "Nota",
        ],
        data,
        money_cols=(3, 4),
        widths=[12, 100, 16, 16, 80],
    )
    ws.cell(end, 2, "Total na ponte (linha C3)").font = BOLD
    ws.cell(end, 4, new_assets_total).number_format = MONEY
    r = end + 1
    for a in new_post:
        r = para(
            ws,
            r,
            f"Fora da ponte (comprado depois de 31.12.2024): linha {a['linha']}, {a['desc'][:90]}, "
            f"R$ {a['valor']:,.2f}",
            NOTE,
        )
    data = [[g["linha"], g["desc"], g["valor"], g["tipo"]] for g in gone]
    end = write_table(
        ws,
        r + 1,
        ["Linha GERAL", "Bem de 2018 que não está em 2025", "Valor 2018", "Situação"],
        data,
        money_cols=(3,),
    )
    ws.cell(end, 2, "Total sem destino documentado (linha G2)").font = BOLD
    ws.cell(end, 3, gone_sem).number_format = MONEY
    para(
        ws,
        end + 2,
        "Bens novos tratados como comprados em dinheiro: a favor do grupo, porque alguns vieram em permuta "
        "de lote. Os veículos recebidos em permuta (Jeep Commander, Amarok, S10, RAV4…) não aparecem no "
        "inventário de 2025: por isso essas vendas contam como dinheiro (linha A1).",
        NOTE,
    )

    # 7 Outras entradas
    ws = wb.create_sheet("7 Outras entradas")
    data = [[x["n"], x["tipo"], str(x["desc"])[:120], x["valor"]] for x in rents18]
    end = write_table(
        ws,
        1,
        ["Nº", "Tipo", "Aluguel (inventário 2018)", "Valor"],
        data,
        money_cols=(4,),
        widths=[6, 18, 100, 16],
    )
    r = end
    for label, v in [
        ("Aluguel mensal equivalente", rent_month),
        (f"× {months} meses", rent_month * months),
    ]:
        ws.cell(r, 3, label).font = BOLD
        ws.cell(r, 4, v).number_format = MONEY
        r += 1
    r += 1
    for line in [
        "Não somadas à SEI nesta versão (outras empresas ou sem valor; a confirmar):",
        (
            "• Venda de 190 ha da São Crispin à Família Rigo (2018): R$ 9.448.000 = 118.100 sacas; 20.150 sacas/ano de 2019 "
            "a 2022 (2021 e 2022 'faltam receber', segundo o grupo). Credor: Hotel Princesa dos Vales."
        ),
        "• Venda de 85,32 ha à Família Pellegrini (2018): R$ 2.100.000 (R$ 1.000.000 em jun/2019).",
        "• Arrendamento da São Crispin à Família Lorenzet (contratos 2017–2026): pago em sacas de soja.",
        (
            "• Indícios de que esse dinheiro passa pela SEI: Funrural/IRPJ/CSLL da Agropecuária São Crispin pagos pela SEI; "
            "'20.150 (−) 17.544 = 2.606 SC SOJA VENDIDOS' (mai/2024); sacas usadas na escavadeira; 'cheques soja Titton' no "
            "lote Dal Ponte."
        ),
        "• Receita de serviços da SEI (notas para CERAN, CEVAL e Prefeitura): impostos no custo, receita fora da ponte.",
        "• Contas a receber: vendas em 26 a 60 parcelas deixam saldo em 31.12.2024, ausente nos dois inventários.",
        "• Locação Oleoplan (arquivo 15): R$ 14.000/mês no 1º ano, segundo o OCR — conferir na imagem.",
    ]:
        r = para(ws, r, line, NOTE)

    # Contratos 11.2
    ws = wb.create_sheet("Contratos 11.2")
    gap = bridge[REF]["sem_explicacao"]
    data = []
    for c in comp:
        prop, why = (
            c["proposta"],
            "valor do grupo, menos o que já foi pago com lotes (cobrança dupla; casos "
            + ", ".join(a["caso"] for a in c["ajustes"])
            + ")"
            if c["ajustes"]
            else "valor do grupo nesta versão (decisão item a item pendente)",
        )
        if c["codigo"] == "11.2.1":
            prop, why = (
                0.0,
                (
                    "sem multa: contratos vigentes. Multa só com distrato e termo de nunca lotear a gleba; "
                    "nesse caso os investimentos nela não são cobrados, mas precisam ser demonstrados"
                ),
            )
        data.append(
            [
                c["codigo"],
                c["item"],
                c["siviero"],
                c["rafael"] or None,
                prop,
                why,
                c["comentario_luis"][:300],
            ]
        )
    data.append(
        [
            "novo",
            f"Investimento a demonstrar (valor sem explicação, cenário {REF})",
            0.0,
            None,
            gap,
            "ativo até o grupo demonstrar o valor real",
            "",
        ]
    )
    end = write_table(
        ws,
        1,
        [
            "Item",
            "Descrição",
            "Siviero (doc 11)",
            "Rafael (Estudo 2)",
            f"Proposta {VERSION}",
            "Motivo da proposta",
            "Comentário do Luís na planilha",
        ],
        data,
        money_cols=(3, 4, 5),
        widths=[8, 62, 16, 16, 16, 70, 70],
    )
    ws.cell(end, 2, "Total").font = BOLD
    for j, v in ((3, tot_siv), (4, tot_raf), (5, sum(d[4] for d in data))):
        c = ws.cell(end, j, v)
        c.font, c.number_format = BOLD, MONEY
    rows_adj = [
        [c["codigo"], a["caso"], a["motivo"], a["valor"]]
        for c in comp
        for a in c["ajustes"]
    ]
    end = write_table(
        ws,
        end + 6,
        [
            "Item",
            "Caso",
            "Cobrança dupla abatida (decisão do Luís, 28/10/2026)",
            "Abatimento (R$)",
        ],
        rows_adj,
        money_cols=(4,),
    )
    ws.cell(end, 3, "Total abatido").font = BOLD
    c = ws.cell(end, 4, sum(r[3] for r in rows_adj))
    c.font, c.number_format = BOLD, MONEY
    para(
        ws,
        end + 1,
        "Detalhe lote a lote: cruzamento-compromissos-arq17.xlsx, aba Avaliação.",
        NOTE,
    )
    end += 2
    para(
        ws,
        end + 2,
        "Rafael soma no 11.2.1 o valor de lotes de loteamentos não executados com as multas "
        "(contraditório). O 11.2.4 de Rafael já inclui o 11.2.5.",
        NOTE,
    )
    para(
        ws,
        end + 3,
        "2018: o grupo atribuía R$ 5.400.000 aos loteamentos em execução e R$ 1.250.000 aos contratados "
        f"(R$ 50.000 cada), além da planilha de investimentos de R$ {sum(x['valor'] for x in inv18):,.2f}. "
        "Em 2025: R$ 0 e multa.",
        NOTE,
    )

    # 8 Preços citados
    ws = wb.create_sheet("8 Preços citados (amostra)")
    data = [
        [
            c["data"].date(),
            c["lot"],
            c["lote"],
            c["quadra"],
            c["valor"],
            c["lot_model"],
            c["v18"],
            c["v25"],
            c["rel25"],
            str(c["obs"])[:160],
        ]
        for c in cited_rows
    ]
    end = write_table(
        ws,
        1,
        [
            "Data",
            "Loteamento (arq. 17)",
            "Lote",
            "Quadra",
            "Valor atribuído",
            "Loteamento no modelo",
            "Valor/lote 2018",
            "Valor/lote 2025",
            "÷ média 2025",
            "Texto",
        ],
        data,
        money_cols=(5, 7, 8),
        widths=[11, 30, 6, 7, 16, 26, 14, 14, 12, 90],
    )
    para(
        ws,
        end + 1,
        "Amostra pequena e não conclusiva (revisão, achado M7): mantida só para consulta.",
        NOTE,
    )

    # Pendências
    ws = wb.create_sheet("Pendências")
    ws.column_dimensions["A"].width = 170
    r = 1
    for line, f in [
        (f"Pendências e limites da versão {VERSION}", Font(bold=True, size=13)),
        ("Correções da revisão independente (revisao-v0.md e revisao-v0.1.md):", BOLD),
        (
            (
                "A1/N1 registros anteriores ao período só usam a redução que sobra · A2 saídas sem data marcadas como "
                "2025 · A3/M1/B1 valores do grupo para os lotes que saíram (Perini, Lago Negro, lotes urbanos); N2 sem "
                "troca da média de 2025 por índice · A4 concessão limitada ao documentado · A5/N3 vendas com parte em "
                "dinheiro contam como dinheiro, permuta lote por lote não · N4 bens recebidos que não aparecem em 2025 "
                "entram como sem explicação · A6 dívida de 2019 sem o que não é dívida da SEI e dívida de 2025 em "
                "31.12.2024 · A8 teste de receita · M2 bens novos e baixados · M6 lotes ligados a loteamentos "
                "contratados · M7 amostra de preços só para consulta · Flávia pelas parcelas até dez/2024."
            ),
            None,
        ),
        (
            (
                "Mudança de método na v0.2 (decisões do Luís de 28/10/2026 na planilha de contas pagas): período "
                "2019–2024; recorrentes extrapolados com IPCA; extraordinários não extrapolados; obras em loteamento "
                "finalizado, loteamento rescindido e despesas de familiares e outras empresas = recorrentes. Combustível "
                "segue como obra, pela instrução do Luís nesta análise (a planilha ainda deixa parte como recorrente)."
            ),
            None,
        ),
        ("", None),
        ("Aprovado pelo Luís (log-decisoes.md desta pasta):", BOLD),
        (
            (
                "• 28/10/2026 — Preço de referência: teto, com o valor de 2025 do grupo trazido pelo IPCA à data da venda "
                "(mesma base dos custos). O teto nominal fica só na aba Sensibilidade."
            ),
            None,
        ),
        (
            "• 28/10/2026 — Os lotes que o arquivo 17 lista além da redução do inventário entram na conta (o estoque de 2018 sobe na mesma quantidade; exceção à P3).",
            None,
        ),
        (
            "• 28/10/2026 — Pró-labore de 4 dos 6 ramos (Válter, Adélia, Pedro e Raul) é custo da empresa: fica no custo de manter.",
            None,
        ),
        ("", None),
        ("Decisões do Luís ainda pendentes e notas:", BOLD),
        (
            (
                "• Nota — IPCA de 2025: a tabela vai até dez/2024; o valor de out/2025 do grupo foi trazido a partir de dez/2024 "
                "(leve viés a favor do sócio que sai)."
            ),
            None,
        ),
        (
            "• Nenhuma premissa de método pendente; itens 11.2.2 a 11.2.5 seguem para decisão item a item.",
            None,
        ),
        ("", None),
        ("Materiais ainda em edição:", BOLD),
        (
            "• Contas pagas classificadas: usada a versão de 28/10/2026 18:53 (Dúvidas decididas). Se mudar, regerar.",
            None,
        ),
        (
            (
                "• Achados da sessão ACERTOS (aguardam aprovação) · São Crispin / Hotel (sessão S1) · levantamento ambiental "
                "(lotes de terceiros, 11.2.2)."
            ),
            None,
        ),
        ("", None),
        ("Limites conhecidos:", BOLD),
        (
            "• Os 15 meses representam todo o período; os juros de 2019–2020 eram maiores (a favor do sócio).",
            None,
        ),
        (
            "• Classificação das formas de saída por palavras-chave: conferir a lista na aba 2.",
            None,
        ),
        (
            (
                "• Banco SICREDI (R$ 5,4 mi): empresa da conta não identificada. Contas a receber e receita de serviços fora da "
                "ponte."
            ),
            None,
        ),
    ]:
        r = para(ws, r, line, f)

    out = HERE / f"fluxo-caixa-sei-{VERSION}.xlsx"
    wb.save(out)
    numbers = {
        "versao": VERSION,
        "gerado_em": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "fontes_md5": hashes,
        "meses": months,
        "referencia": REF,
        "lotes": {
            "n2018": sum(r["n2018"] for r in rows),
            "n2025": sum(r["n2025"] for r in rows),
            "reducao": sum(r["red"] for r in rows),
            "no_periodo": n_in,
            "apos_2024": post["n"],
            "antes_2018": pre["n"],
            "indice_2018_2025": ratio,
            "por_classe": dict(n_cls),
        },
        "ponte": bridge,
        "apos_2024": post,
        "antes_2018": pre,
        "custos": {
            "m_mes": m_month,
            "m_rec": m_rec,
            "m_ext": m_ext,
            "classes": dict(by_cls),
            "sei_15m": sei_15m,
            "inv_por_situacao": dict(inv_tag),
            "prolabore_socios": partner_pl,
        },
        "divida": {
            "jan2019": debt18,
            "dez2024": debt25,
            "flavia": flavia_total,
            "quitada": debt_paid,
        },
        "bens_novos": new_assets_total,
        "bens_2018_sem_destino": gone_sem,
        "alugueis_mes": rent_month,
        "concessao": {"documentada": concession, "maxima": concession_max},
        "teste_receita": revenue_test,
        "contratos": comp,
        "contratos_total": {"siviero": tot_siv, "rafael": tot_raf},
        "investimentos_2018": {
            "planilha": sum(x["valor"] for x in inv18),
            "em_execucao": 5_400_000,
            "contratados": 1_250_000,
        },
        "trajetoria": path,
        "sensibilidade": sens,
        "excedente_lista": {"n": excess_n, **excess_val},
    }
    (HERE / f"numeros-{VERSION}.json").write_text(
        json.dumps(numbers, ensure_ascii=False, indent=1, default=str), encoding="utf-8"
    )
    print(
        "lotes 2018/2025/red:",
        numbers["lotes"]["n2018"],
        numbers["lotes"]["n2025"],
        numbers["lotes"]["reducao"],
        "| período",
        round(n_in, 2),
        "| 2025",
        round(post["n"], 2),
        "| antes 2018",
        pre["n"],
    )
    print("por classe", {k: round(v, 1) for k, v in n_cls.items()})
    print(
        "manter/mês",
        round(m_month),
        "rec",
        round(m_rec),
        "ext",
        round(m_ext),
        "| dívida 2019",
        debt18,
        "2024",
        round(debt25, 2),
        "paga",
        round(debt_paid, 2),
        "| bens novos",
        round(new_assets_total),
        "| baixados",
        gone_sem,
        "| concessão",
        concession,
    )
    for k in ("piso", "central", "teto"):
        print(k, {kk: round(vv) for kk, vv in bridge[k].items()})
    print(
        "teste receita",
        {
            k: (
                round(v)
                if isinstance(v, float)
                else {a: round(b) for a, b in v.items()}
            )
            for k, v in revenue_test.items()
        },
    )
    print(
        "sens",
        [
            (
                s["inicio"],
                s["concessao"],
                round(s["piso"]),
                round(s["central"]),
                round(s["teto"]),
            )
            for s in sens
        ],
    )
    print("excedente", excess_n, {k: round(v) for k, v in excess_val.items()})
    print(
        "trajetória",
        [(p["ano"], round(p["lotes"], 1), round(p["valor25"])) for p in path],
    )
    print("gone", [(g["linha"], g["valor"], g["tipo"]) for g in gone])
    print("saved", out)


if __name__ == "__main__":
    main()
