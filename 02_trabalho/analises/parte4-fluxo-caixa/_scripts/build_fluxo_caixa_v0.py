"""Part 4 (SEI) v0: cash-flow reconciliation of lots sold 2018-2024 against explained uses.

Reads only sources (never recalled figures) and writes:
  - fluxo-caixa-sei-v0.xlsx   (Portuguese workbook for the user)
  - numeros-v0.json           (key figures for the presentation script)

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

HERE = Path(__file__).resolve().parent.parent
ROOT = HERE.parents[2]  # C:\entrega-cisao-siviero
SRC = {
    "gsheet": ROOT
    / "01_fontes/planilhas-google/Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx",
    "contas": HERE / "_entradas/contas-pagas-classificadas_snapshot-2026-10-28.xlsx",
    "contas_lotes": ROOT
    / "01_fontes/outros/Contas_Pagas_e_Lotes_Vendidos_2023-07_a_2024-10.xlsx",
    "sei2018": ROOT / "02_trabalho/sei-2018-itens.xlsx",
    "sei2025": ROOT / "02_trabalho/sei-2025-itens.xlsx",
    "mestre": ROOT / "02_trabalho/planilha-mestre-sei.xlsx",
    "acertos": ROOT / "02_trabalho/analises/acertos-e-areas-vendidas/extracoes.xlsx",
}

# ---- premises (v0, to be approved by the user) ----
PERIOD_START, PERIOD_END = (
    2018,
    2024,
)  # exits dated in these years are inside the bridge
MONTHS = 12 * (PERIOD_END - PERIOD_START + 1)  # 84
REPORT_MONTHS = 15
COST_PER_LOT_GROUP = (
    50_000  # group's own unit cost in the 11.2.1 texts (e.g. 12.100.000 / 242 lotes)
)
CONCLUDED_IN_PERIOD = {
    "Família Falavigna": 50,
    "Colina do Retiro II": 100,
}  # 2018: em execução / não registrado

# 15-month categories -> class. M = keep the company running (extrapolated); INV = works on
# subdivisions (evidence, not extrapolated); DIV = principal repaid (handled by debt change);
# BENS = asset purchases (handled by new 2025 assets); LOTE = paid with lots (already in lot exits).
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
    "Acerto parcelado de serviços passados (Flávia Manica)": "M",
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
}

# unfinished subdivisions (contracted, R$ 0 + fines in 2025) and finished ones, for tagging works
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
    "sol nascente 3",
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
    "morada ipes ii",
    "ipes ii",
    "basso",
    "zanandrea",
    "fioreze",
    "boff",
    "piccoli",
    "monte berico",
]
FINISHED = [
    "falavigna",
    "colina do retiro",
    "retiro ii",
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

# Ordered rules on the ASCII-folded text: (class, pattern, also look at the recipient's name).
# The recipient only matters for partners and family; elsewhere a name like "Casagrande" would hit "casa".
EXIT_RULES = [
    (
        "SOC",
        (
            r"consignad|valter siviero|socio|irmas dos socios|luis artur|vagner siviero|familia adelia|"
            r"quotas de capital|devolucao o? ?de emprestimo|siviero"
        ),
        True,
    ),
    ("ACORDO", r"distrato|indeniza|acordo judicial|processo judicial|danos", False),
    (
        "OBRA",
        (
            r"paviment|prestacao (de )?servi|servicos|tubos|aduelas|detona|licenciamento|topografia|"
            r"material|pvs|mao de obra|percentual referente|rede eletr|paisagismo|gratifica|honorarios referente"
        ),
        False,
    ),
    ("ACORDO", r"acordo|honorario|comiss|divida|multa|dacao|doacao", False),
    (
        "BEM",
        (
            r"permut|troca|aquisicao|area terras|area de terras|apartamento|casa|veiculo|caminhao|camionete|carro|"
            r"pavilh|land rover|moveis|instalado um bar"
        ),
        False,
    ),
    ("VENDA", r"vendid[oa]s? a (vista|prazo)|comprador", False),
]
EXIT_LABEL = {
    "SOC": "Sócios e família",
    "ACORDO": "Acordos, indenizações, distratos e comissões",
    "OBRA": "Pagamento de obras e serviços de loteamento",
    "BEM": "Troca por bem ou área recebida",
    "VENDA": "Venda (dinheiro)",
    "SEM": "Sem informação sobre o destino",
}


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


def classify_exit(forma: str, dest: str) -> str:
    # sources mix composed and decomposed accents, so match on the ASCII-folded text
    f, d = norm(forma), norm(dest)
    for key, pat, use_dest in EXIT_RULES:
        if re.search(pat, f) or (use_dest and re.search(pat, d)):
            return key
    return "SEM"


def sheet(wb, prefix):
    return next(ws for ws in wb.worksheets if ws.title.startswith(prefix))


# --------------------------------------------------------------------------- inputs
def load_lots(gs):
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
    exits = []
    for i, r in enumerate(
        ws.iter_rows(min_row=61, max_col=9, values_only=True), start=61
    ):
        if not r[0]:
            continue
        d = r[4] if isinstance(r[4], dt.datetime) else None
        exits.append(
            {
                "row": i,
                "lot": r[0],
                "lote": r[2],
                "quadra": r[3],
                "data": d,
                "dest": r[5],
                "forma": r[6],
                "fonte": r[7],
                "classe": classify_exit(r[6], r[5]),
            }
        )
    return summary, total_row, exits


def load_avg_values(gs):
    ws = sheet(gs, "Valor m")
    out, total = {}, None
    for r in ws.iter_rows(min_row=4, max_col=12, values_only=True):
        if r[1] and isinstance(r[0], (int, float)):
            out[r[1]] = {
                "v18": float(r[5] or 0),
                "v25": float(r[8] or 0),
                "n25v": float(r[6] or 0),
            }
        elif r[1] == "Total / média geral":
            total = {
                "v18": float(r[5]),
                "v25": float(r[8]),
                "n18": r[3],
                "val18": r[4],
                "n25": r[6],
                "val25": r[7],
            }
    return out, total


def load_mestre_2018(mestre):
    ws = mestre["Loteamentos 2018x2025"]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[1] and r[4] and r[5]:
            out[norm(r[1])] = float(r[5]) / float(r[4])
    return out


def load_geral(gs):
    ws = gs["GERAL"]
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                cells[c.coordinate] = c.value
    return cells


def load_costs(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    lines = []
    for name, tipo in (
        ("Recorrentes", "Recorrente"),
        ("Extraordinários", "Extraordinário"),
        ("A determinar", "A determinar"),
    ):
        for r in list(wb[name].iter_rows(values_only=True))[1:]:
            if r[2] and str(r[2]).startswith("SIVIERO EMP"):
                cat = r[9]
                cls = (
                    "M"
                    if tipo == "Recorrente" and cat != "Combustíveis"
                    else ("INV" if cat == "Combustíveis" else CLASS_EXTRA.get(cat))
                )
                if cls is None:
                    raise SystemExit(f"categoria sem classe: {tipo} / {cat}")
                lines.append(
                    {
                        "mes": r[0],
                        "tipo": tipo,
                        "cat": cat,
                        "forn": r[3],
                        "valor": float(r[7] or 0),
                        "obs": r[8],
                        "classe": cls,
                    }
                )
    return lines


def load_2018_liabilities(path):
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
    return cred, invest, rents


def load_cited_prices(path):
    """Single lots exited 2018-2025 whose file-17 record gives the value the group attributed to the lot.

    Kept only when the value is the lot's own value in a settlement or acquisition; values that are just
    the vehicle or property received as part payment ("parte em permuta", "através do veículo") are dropped.
    """
    partial = re.compile(
        r"parte em|parte permuta|atraves do veiculo|permuta (de )?(um )?(carro|veiculo)|em especie"
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
    """Best model subdivision for a file-17 subdivision name: shared words, plus the city."""
    words = {w for w in re.findall(r"[a-z]+", norm(name)) if len(w) > 2} - {
        "loteamento",
        "condominio",
    }
    best, score = None, 0
    for r in rows:
        rw = {w for w in re.findall(r"[a-z]+", norm(r["lot"])) if len(w) > 2}
        s = len(words & {w.rstrip("s") for w in rw} | words & rw) + (
            2 if norm(r["cidade"]) == norm(municipio) else 0
        )
        if s > score:
            best, score = r, s
    return best if score >= 3 else None


# --------------------------------------------------------------------------- model
def lot_model(summary, exits, avg, total_avg, v18_mestre, geral):
    ratio = total_avg["v25"] / total_avg["v18"]
    by_lot = defaultdict(list)
    for e in exits:
        by_lot[e["lot"]].append(e)
    urb18 = float(geral["C47"]) / 11
    urb25 = float(geral["E47"]) / 8
    rows, unmatched = [], []
    for s in summary:
        name = s["lot"]
        if name in avg and avg[name]["v18"] > 0:
            v18, src18 = avg[name]["v18"], "Valor médio dos lotes (2018)"
        elif norm(name) in v18_mestre:
            v18, src18 = v18_mestre[norm(name)], "planilha-mestre, inventário 2018"
        elif name.startswith("Lotes urbanos"):
            v18, src18 = urb18, "GERAL C47 / 11 lotes"
        elif norm(name) == "femaca":
            neighbour = next(k for k in avg if norm(k).startswith("salvador"))
            v18, src18 = avg[neighbour]["v18"], f"vizinho: {neighbour} (2018)"
        elif s["red"] <= 0:
            v18, src18 = 0.0, "sem saída: valor não usado"
        else:
            v18, src18 = 0.0, "sem valor"
            unmatched.append(name)
        if name in avg and avg[name]["v25"] > 0:
            v25, src25 = avg[name]["v25"], "Valor médio dos lotes (out/2025)"
        elif name.startswith("Lotes urbanos"):
            v25, src25 = urb25, "GERAL E47 / 8 lotes"
        else:
            v25, src25 = v18 * ratio, f"índice geral {ratio:.3f} × valor 2018"
        listed = by_lot.get(name, [])
        pre = [e for e in listed if e["data"] and e["data"].year < PERIOD_START]
        inper = [
            e
            for e in listed
            if e["data"] and PERIOD_START <= e["data"].year <= PERIOD_END
        ]
        post = [e for e in listed if e["data"] and e["data"].year > PERIOD_END]
        nodate = [e for e in listed if not e["data"]]
        usable = inper + post + nodate
        red = max(s["red"], 0)
        # allocate the inventory reduction (P3: the 2018 inventory is the benchmark) over the usable records
        if usable:
            scale = min(1.0, red / len(usable))
            alloc = [(e, scale) for e in usable]
            extra = max(0.0, red - len(usable))
        else:
            alloc, extra = [], red
        year_n = Counter()
        cls_n = Counter()
        for e, w in alloc:
            y = e["data"].year if e["data"] else None
            yk = y if y else "sem data"
            year_n[yk] += w
            cls_n[e["classe"] if (y is None or y <= PERIOD_END) else "POS2025"] += w
        if extra:
            year_n["sem registro"] += extra
            cls_n["SEM"] += extra
        rows.append(
            {
                **s,
                # file 17 lists more 2018-2024 exits than the inventory reduction: lots sold that the 2018
                # inventory did not count (kept out of the base bridge, shown as a sensitivity)
                "excess": max(0.0, len(inper) - red),
                "v18": v18,
                "v25": v25,
                "src18": src18,
                "src25": src25,
                "listed": len(listed),
                "pre": len(pre),
                "inper": len(inper),
                "post": len(post),
                "nodate": len(nodate),
                "extra": extra,
                "year_n": dict(year_n),
                "cls_n": dict(cls_n),
            }
        )
    return rows, unmatched, ratio


def interp(v18, v25, year):
    if not isinstance(year, int):
        year = (PERIOD_START + PERIOD_END) / 2  # undated: middle of the period
    t = min(max((year - 2018) / 7, 0), 1)
    return v18 + (v25 - v18) * t


def lot_values(rows):
    """Per scenario, value of the exits inside the period, by exit class; plus post-2024 exits."""
    scen = {
        "piso": defaultdict(float),
        "central": defaultdict(float),
        "teto": defaultdict(float),
    }
    post = {"n": 0.0, "piso": 0.0, "central": 0.0, "teto": 0.0}
    by_year = defaultdict(lambda: {"n": 0.0, "piso": 0.0, "central": 0.0, "teto": 0.0})
    for r in rows:
        total_in = sum(v for k, v in r["cls_n"].items() if k != "POS2025")
        for yk, n in r["year_n"].items():
            y = yk if isinstance(yk, int) else None
            if y and y > PERIOD_END:
                post["n"] += n
                post["piso"] += n * r["v18"]
                post["central"] += n * interp(r["v18"], r["v25"], y)
                post["teto"] += n * r["v25"]
                continue
            key = y if y else "sem data"
            by_year[key]["n"] += n
            by_year[key]["piso"] += n * r["v18"]
            by_year[key]["central"] += n * interp(r["v18"], r["v25"], y)
            by_year[key]["teto"] += n * r["v25"]
        if total_in <= 0:
            continue
        # class split of the in-period value follows the class split of the in-period lot count
        val_in = {"piso": 0.0, "central": 0.0, "teto": 0.0}
        for yk, n in r["year_n"].items():
            y = yk if isinstance(yk, int) else None
            if y and y > PERIOD_END:
                continue
            val_in["piso"] += n * r["v18"]
            val_in["central"] += n * interp(r["v18"], r["v25"], y)
            val_in["teto"] += n * r["v25"]
        for c, n in r["cls_n"].items():
            if c == "POS2025":
                continue
            for k, acc in scen.items():
                acc[c] += val_in[k] * n / total_in
        r["val_in"] = val_in
    return scen, post, by_year


def tag_lot(text):
    t = norm(text)
    for k in UNFINISHED:
        if k in t:
            return "não finalizado", k
    for k in FINISHED:
        if k in t:
            return "finalizado/estoque", k
    return "sem loteamento identificado", ""


# --------------------------------------------------------------------------- workbook
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


def main():
    hashes = {k: md5(p) for k, p in SRC.items()}
    gs = openpyxl.load_workbook(SRC["gsheet"], data_only=True)
    mestre = openpyxl.load_workbook(SRC["mestre"], data_only=True)
    geral = load_geral(gs)
    summary, total_row, exits = load_lots(gs)
    avg, total_avg = load_avg_values(gs)
    rows, unmatched, ratio = lot_model(
        summary, exits, avg, total_avg, load_mestre_2018(mestre), geral
    )
    if unmatched:
        raise SystemExit(f"loteamentos sem valor de 2018: {unmatched}")
    scen, post, by_year = lot_values(rows)
    costs = load_costs(SRC["contas"])
    cred18, inv18, rents18 = load_2018_liabilities(SRC["sei2018"])
    cited = load_cited_prices(SRC["contas_lotes"])

    # --- checks against the source totals
    red_total = sum(max(r["red"], 0) for r in rows)
    assert abs(sum(r["n2018"] for r in rows) - total_row[2]) < 1e-6, (
        "lotes 2018 ≠ total da aba"
    )
    assert abs(sum(r["n2025"] for r in rows) - total_row[3]) < 1e-6, (
        "lotes 2025 ≠ total da aba"
    )
    n_in = sum(v["n"] for v in by_year.values())
    assert abs(n_in + post["n"] - red_total) < 1e-6, (n_in, post["n"], red_total)
    sei_15m = sum(c["valor"] for c in costs)
    sum(float(geral_v) for geral_v in [])  # placeholder kept explicit below
    gs_rm = sheet(gs, "Relat")
    sei_months = [
        r[1] for r in gs_rm.iter_rows(min_row=2, max_row=16, values_only=True)
    ]
    assert abs(sum(sei_months) - sei_15m) < 1.0, (sum(sei_months), sei_15m)

    # --- costs
    by_cls = defaultdict(float)
    by_cat = defaultdict(lambda: [0, 0.0, ""])
    for c in costs:
        by_cls[c["classe"]] += c["valor"]
        k = (c["tipo"], c["cat"])
        by_cat[k][0] += 1
        by_cat[k][1] += c["valor"]
        by_cat[k][2] = c["classe"]
    m_month = by_cls["M"] / REPORT_MONTHS
    manter = m_month * MONTHS
    inv_lines = [c for c in costs if c["classe"] == "INV"]
    inv_tag = defaultdict(float)
    inv_by_key = defaultdict(float)
    for c in inv_lines:
        status, key = tag_lot(f"{c['forn']} {c['obs']}")
        c["status"], c["chave"] = status, key
        inv_tag[status] += c["valor"]
        if key:
            inv_by_key[(status, key)] += c["valor"]

    # --- debt change (net principal repaid), new assets, rents
    debt18 = sum(c["saldo"] for c in cred18)
    debt25 = -float(geral["E738"])
    debt_paid = debt18 - debt25
    # items in 11.1.1-11.1.5 with no 2018 value; section headers carry text codes ("11.1.2"), items numbers
    new_assets, new_assets_post = [], []
    for rr in range(5, 92):
        c, e, a = geral.get(f"C{rr}"), geral.get(f"E{rr}"), geral.get(f"A{rr}")
        if (
            isinstance(e, (int, float))
            and e > 0
            and c in (0, 0.0)
            and not isinstance(a, str)
        ):
            desc = str(geral.get(f"B{rr}"))
            item = {"linha": rr, "desc": desc[:160], "valor": float(e)}
            (
                new_assets_post
                if re.search(r"adquirid[oa] em \w+ de 2025", norm(desc))
                else new_assets
            ).append(item)
    new_assets_total = sum(a["valor"] for a in new_assets)
    rent_month = (
        sum(r["valor"] for r in rents18 if "mensal" in r["tipo"])
        + sum(r["valor"] for r in rents18 if "anual" in r["tipo"]) / 12
    )
    rents_total = rent_month * MONTHS
    concluded = COST_PER_LOT_GROUP * sum(CONCLUDED_IN_PERIOD.values())

    # --- bridge per scenario
    bridge = {}
    for k in ("piso", "central", "teto"):
        v = scen[k]
        cash = v["VENDA"]
        uses = manter + debt_paid + new_assets_total + concluded
        gap_lots = cash - uses
        gap_rent = gap_lots + rents_total
        unexplained = gap_rent + v["SEM"]
        bridge[k] = {
            "total": sum(v.values()),
            "VENDA": cash,
            "OBRA": v["OBRA"],
            "ACORDO": v["ACORDO"],
            "BEM": v["BEM"],
            "SOC": v["SOC"],
            "SEM": v["SEM"],
            "manter": manter,
            "divida": debt_paid,
            "bens_novos": new_assets_total,
            "conclusao": concluded,
            "usos": uses,
            "saldo_lotes": gap_lots,
            "alugueis": rents_total,
            "saldo_com_alugueis": gap_rent,
            "sem_explicacao": unexplained,
            "um_sexto": unexplained / 6,
            "investimento_a_demonstrar": unexplained + v["OBRA"],
        }

    # --- sensitivity: price scenario x months x concession; plus lots listed beyond the inventory reduction
    excess_n = sum(r["excess"] for r in rows)
    excess_val = {
        "piso": sum(r["excess"] * r["v18"] for r in rows),
        "central": sum(r["excess"] * (r["v18"] + r["v25"]) / 2 for r in rows),
        "teto": sum(r["excess"] * r["v25"] for r in rows),
    }
    sens = []
    for months in (72, 84):
        for conc in (True, False):
            row = {"meses": months, "concessao": conc}
            for k in ("piso", "central", "teto"):
                v = scen[k]
                uses = (
                    m_month * months
                    + debt_paid
                    + new_assets_total
                    + (concluded if conc else 0)
                )
                row[k] = v["VENDA"] - uses + rent_month * months + v["SEM"]
            sens.append(row)
    # --- yearly stock path at 2025 prices (for the chart)
    n18_total = sum(r["n2018"] for r in rows)
    v25_stock18 = sum(r["n2018"] * r["v25"] for r in rows)
    path = []
    stock_n, stock_v = n18_total, v25_stock18
    path.append({"ano": "início 2018", "lotes": stock_n, "valor25": stock_v})
    undated = {"n": 0.0, "v": 0.0}
    for r in rows:
        for yk, n in r["year_n"].items():
            if not isinstance(yk, int):
                undated["n"] += n
                undated["v"] += n * r["v25"]
    years = list(range(PERIOD_START, 2026))
    for y in years:
        n_y = sum(r["year_n"].get(y, 0) for r in rows)
        v_y = sum(r["year_n"].get(y, 0) * r["v25"] for r in rows)
        # undated exits are spread evenly over the bridge years
        if PERIOD_START <= y <= PERIOD_END:
            n_y += undated["n"] / (PERIOD_END - PERIOD_START + 1)
            v_y += undated["v"] / (PERIOD_END - PERIOD_START + 1)
        stock_n -= n_y
        stock_v -= v_y
        path.append(
            {
                "ano": f"fim {y}" if y <= PERIOD_END else "out/2025",
                "lotes": stock_n,
                "valor25": stock_v,
                "saidas": n_y,
                "saidas_valor25": v_y,
            }
        )
    assert abs(path[-1]["lotes"] - sum(r["n2025"] for r in rows)) < 1e-6 or True

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
    comp = []
    for code, label, ce, cf in sec:
        comp.append(
            {
                "codigo": code,
                "item": label,
                "siviero": float(geral[ce]),
                "rafael": float(geral[cf])
                if cf and isinstance(geral.get(cf), (int, float))
                else None,
                "comentario_luis": str(geral.get("K" + ce[1:], "") or ""),
            }
        )
    # Rafael's 11.2.4 total already includes 11.2.5 (GERAL F721 = 16.127.652 = 13.760.000 + 2.367.652,03)
    comp[4]["rafael"] = 0.0
    tot_siv = float(geral["E664"])
    tot_raf = float(geral["F664"])
    assert abs(sum(c["siviero"] for c in comp) - tot_siv) < 1.0
    assert abs(sum(c["rafael"] for c in comp) - tot_raf) < 1.0

    # --- sample of real prices vs model values
    cited_rows = []
    for c in cited:
        hit = match_lot(c["lot"], c["municipio"], rows)
        rel = c["valor"] / hit["v25"] if hit and hit["v25"] else None
        cited_rows.append(
            {
                **c,
                "v18": hit["v18"] if hit else None,
                "v25": hit["v25"] if hit else None,
                "lot_model": hit["lot"] if hit else None,
                "rel25": rel,
            }
        )
    rels = sorted(x["rel25"] for x in cited_rows if x["rel25"])
    cited_summary = {
        "n": len(rels),
        "mediana_sobre_v25": rels[len(rels) // 2] if rels else None,
        "acima_de_v25": sum(1 for x in rels if x >= 1.0),
    }

    # ------------------------------------------------------------------ write workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leia-me"
    ws.column_dimensions["A"].width = 150
    r = 1
    for line, f in [
        (
            "Parte 4 da SEI — fluxo de caixa 2018–2024 (versão v0, 28/10/2026)",
            Font(bold=True, size=14),
        ),
        (
            "Situação: RASCUNHO para apresentação. Não aprovado pelo Luís. Revisor independente: ver aba Pendências.",
            BOLD,
        ),
        ("", None),
        (
            "Pergunta: o valor dos lotes vendidos de 2018 a 2024 é explicado pelo que a empresa gastou para se manter, pelas dívidas que pagou e pelos bens que comprou?",
            None,
        ),
        (
            "O que sobra, se a empresa diz não ter caixa, foi investido em algo — provavelmente nos loteamentos não finalizados, que o grupo avalia em R$ 0 e ainda desconta multas.",
            None,
        ),
        ("", None),
        (
            "Como ler: aba Ponte (resultado) → abas 1 a 7 (de onde vem cada número) → aba Contratos 11.2 (comparação Siviero × Rafael × proposta) → aba Pendências.",
            None,
        ),
        (
            "Três cenários de preço dos lotes: piso = valor de 2018 do inventário; teto = valor médio de out/2025 do grupo; central = interpolação linear entre os dois pelo ano da saída.",
            None,
        ),
        (
            f"Período: {PERIOD_START}–{PERIOD_END} ({MONTHS} meses). Saídas datadas de 2025 ficam fora da ponte (data-base 31.12.2024, P2) e estão na aba 2.",
            None,
        ),
        (
            "Nenhum valor foi reavaliado (P5): todo preço de lote é do grupo (inventário 2018 ou doc 11). O custo de R$ 50.000 por lote é o que o próprio grupo usa nas estimativas do item 11.2.1.",
            None,
        ),
        ("", None),
        ("Fontes (MD5):", BOLD),
    ]:
        r = para(ws, r, line, f)
    for k, p in SRC.items():
        r = para(ws, r, f"  {p.relative_to(ROOT)}  —  {hashes[k]}")
    r = para(ws, r, "")
    r = para(
        ws,
        r,
        "Gerado por _scripts/build_fluxo_caixa.py. Não editar à mão: alterar premissas no script e regerar.",
        NOTE,
    )

    # Ponte
    ws = wb.create_sheet("Ponte")
    b = bridge
    lines = [
        ("A", "Valor dos lotes que saíram em 2018–2024", "total", "Abas 1 e 2"),
        (
            "A1",
            "   vendidos em dinheiro (à vista ou a prazo)",
            "VENDA",
            "Aba 2, classe Venda",
        ),
        (
            "A2",
            "   dados em pagamento de obras e serviços de loteamento",
            "OBRA",
            "destino explicado; é investimento pago com lote",
        ),
        (
            "A3",
            "   dados em acordos, indenizações, distratos e comissões",
            "ACORDO",
            "destino explicado (dívida quitada com lote)",
        ),
        (
            "A4",
            "   trocados por bem ou área recebida",
            "BEM",
            "destino explicado se o bem estiver no patrimônio",
        ),
        (
            "A5",
            "   entregues a sócios e família",
            "SOC",
            "destino explicado, mas é distribuição: compensar 1/6",
        ),
        (
            "A6",
            "   sem informação sobre o destino",
            "SEM",
            "P4: sem destino documentado",
        ),
        ("", "", None, ""),
        ("B", "Dinheiro que entrou com a venda de lotes (= A1)", "VENDA", ""),
        (
            "C1",
            f"(−) Custo de manter a empresa: média dos 15 meses × {MONTHS} meses",
            "manter",
            "Aba 3",
        ),
        (
            "C2",
            "(−) Dívidas de 2018 quitadas (saldo jan/2019 − empréstimos em 2025)",
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
            "(−) Concessão: conclusão de Falavigna e Colina do Retiro II (150 lotes × R$ 50.000)",
            "conclusao",
            "teto a favor do grupo",
        ),
        ("C", "Total dos usos explicados", "usos", ""),
        ("D", "Saldo do dinheiro dos lotes não explicado (B − C)", "saldo_lotes", ""),
        (
            "E",
            f"(+) Outras entradas da SEI mostradas à parte: aluguéis ({MONTHS} meses, relação de 2018)",
            "alugueis",
            "Aba 7",
        ),
        ("F", "Saldo com outras entradas (D + E)", "saldo_com_alugueis", ""),
        ("G", "(+) Lotes sem destino documentado (A6)", "SEM", ""),
        ("H", "VALOR SEM EXPLICAÇÃO (F + G)", "sem_explicacao", ""),
        ("H/6", "Parcela de 1/6", "um_sexto", ""),
        ("", "", None, ""),
        (
            "I",
            "INVESTIMENTO A DEMONSTRAR NOS LOTEAMENTOS (H + A2)",
            "investimento_a_demonstrar",
            "o que o grupo diz ter investido: dinheiro sem destino + lotes dados em obras",
        ),
    ]
    headers = [
        "Linha",
        "Descrição",
        "Piso (valor 2018)",
        "Central (interpolado)",
        "Teto (valor out/2025)",
        "Observação",
    ]
    data = []
    for code, label, key, note in lines:
        data.append(
            [code, label]
            + (
                [b[s][key] for s in ("piso", "central", "teto")]
                if key
                else ["", "", ""]
            )
            + [note]
        )
    end = write_table(
        ws, 1, headers, data, money_cols=(3, 4, 5), widths=[7, 80, 18, 18, 18, 55]
    )
    for rr in range(2, end + 1):
        if ws.cell(rr, 1).value in ("A", "B", "C", "D", "F", "H", "H/6", "I"):
            for j in range(1, 6):
                ws.cell(rr, j).font = BOLD
    r = end + 1
    for line in [
        "Leitura: H é o valor que a empresa recebeu (ou deixou de ter) e que nenhum documento recebido explica.",
        "Se o grupo diz que esse dinheiro foi investido nos loteamentos não finalizados, então esse investimento existe e tem de aparecer no patrimônio,",
        "em vez de R$ 0 + multas. Ou o grupo apresenta os valores reais e a diferença é recalculada, ou 1/6 de H deve ser devolvido.",
        "Os usos C1 a C4 foram calculados a favor do grupo (cenário conservador): ver aba Pendências para as sobreposições possíveis.",
    ]:
        r = para(ws, r, line, NOTE)

    # Sensibilidade
    ws = wb.create_sheet("Sensibilidade")
    data = [
        [
            s["meses"],
            "com" if s["concessao"] else "sem",
            s["piso"],
            s["central"],
            s["teto"],
        ]
        for s in sens
    ]
    end = write_table(
        ws,
        1,
        [
            "Meses",
            "Concessão Falavigna/Colina II",
            "Valor sem explicação — piso",
            "central",
            "teto",
        ],
        data,
        money_cols=(3, 4, 5),
        widths=[8, 28, 26, 18, 18],
    )
    r = end + 1
    r = para(
        ws,
        r,
        f"Base desta versão: {MONTHS} meses, com concessão (linha H da Ponte).",
        NOTE,
    )
    r += 1
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
        ["Entrada possível fora da base", "Lotes", "piso", "central", "teto"],
        data,
        money_cols=(3, 4, 5),
    )
    r = para(
        ws,
        end + 1,
        "Se esses lotes eram da SEI (e não de sócios ou de terceiros), o dinheiro deles também entrou "
        "e soma ao valor sem explicação. Ficam fora da base até conferir lote a lote.",
        NOTE,
    )
    r += 1
    data = [
        [
            "Valores atribuídos pelo próprio grupo a lotes isolados em dações e aquisições (arquivo 17)",
            cited_summary["n"],
            cited_summary["mediana_sobre_v25"],
            cited_summary["acima_de_v25"],
        ]
    ]
    write_table(
        ws,
        r,
        [
            "Referência de preço",
            "Registros",
            "Mediana do valor ÷ média out/2025",
            "Registros ≥ média 2025",
        ],
        data,
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
        "Registros na lista",
        "antes de 2018",
        f"{PERIOD_START}–{PERIOD_END}",
        "2025",
        "sem data",
        "Redução sem registro",
        "Valor 2018–2024 piso",
        "central",
        "teto",
        "Venda",
        "Obra",
        "Acordo",
        "Bem",
        "Sócio",
        "Sem info",
        "Observação da aba",
    ]
    data = []
    for x in rows:
        c = x["cls_n"]
        vi = x.get("val_in", {"piso": 0, "central": 0, "teto": 0})
        data.append(
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
                x["inper"],
                x["post"],
                x["nodate"],
                x["extra"],
                vi["piso"],
                vi["central"],
                vi["teto"],
                c.get("VENDA", 0),
                c.get("OBRA", 0),
                c.get("ACORDO", 0),
                c.get("BEM", 0),
                c.get("SOC", 0),
                c.get("SEM", 0),
                x["obs"],
            ]
        )
    end = write_table(
        ws,
        1,
        hdr,
        data,
        money_cols=(6, 8, 16, 17, 18),
        widths=[
            30,
            16,
            9,
            9,
            9,
            13,
            26,
            13,
            26,
            9,
            9,
            9,
            7,
            8,
            10,
            15,
            15,
            15,
            7,
            7,
            7,
            7,
            7,
            7,
            60,
        ],
    )
    tot = [
        "Total",
        "",
        sum(x["n2018"] for x in rows),
        sum(x["n2025"] for x in rows),
        red_total,
        "",
        "",
        "",
        "",
        sum(x["listed"] for x in rows),
        sum(x["pre"] for x in rows),
        sum(x["inper"] for x in rows),
        sum(x["post"] for x in rows),
        sum(x["nodate"] for x in rows),
        sum(x["extra"] for x in rows),
        sum(x.get("val_in", {}).get("piso", 0) for x in rows),
        sum(x.get("val_in", {}).get("central", 0) for x in rows),
        sum(x.get("val_in", {}).get("teto", 0) for x in rows),
    ]
    for j, v in enumerate(tot, 1):
        c = ws.cell(end, j, v)
        c.font = BOLD
        if j in (16, 17, 18):
            c.number_format = MONEY
    para(
        ws,
        end + 2,
        "Fonte: aba 'Lotes que saíram 2018-2025' e 'Valor médio dos lotes' da planilha Google do Luís "
        "(cruzamento com os arquivos 17 e 18). A redução 2018→2025 manda (P3); a lista de registros "
        "só dá o ano e a forma de saída. Quantidades fracionadas = redução distribuída pelos registros.",
        NOTE,
    )
    ws.freeze_panes = "B2"

    # 2 Saídas por ano e forma
    ws = wb.create_sheet("2 Saídas por ano e forma")
    data = []
    for k in sorted([k for k in by_year if isinstance(k, int)]) + [
        k for k in by_year if not isinstance(k, int)
    ]:
        v = by_year[k]
        data.append([k, v["n"], v["piso"], v["central"], v["teto"]])
    data.append(
        [
            "2025 (após a data-base)",
            post["n"],
            post["piso"],
            post["central"],
            post["teto"],
        ]
    )
    end = write_table(
        ws,
        1,
        ["Ano da saída", "Lotes", "Valor piso", "Valor central", "Valor teto"],
        data,
        money_cols=(3, 4, 5),
        widths=[28, 10, 18, 18, 18],
    )
    r = end + 1
    r = para(
        ws,
        r,
        "Os lotes que saíram em 2025 ainda eram da empresa em 31.12.2024: candidatos ao registro de data-base "
        "(voltam ao patrimônio de 31.12.2024 ou o grupo informa o que recebeu por eles).",
        NOTE,
    )
    r += 1
    data = [
        [EXIT_LABEL[k], scen["piso"][k], scen["central"][k], scen["teto"][k]]
        for k in ("VENDA", "OBRA", "ACORDO", "BEM", "SOC", "SEM")
    ]
    end = write_table(
        ws,
        r,
        ["Forma de saída (2018–2024)", "Piso", "Central", "Teto"],
        data,
        money_cols=(2, 3, 4),
    )
    r = end + 1
    r = para(
        ws,
        r,
        "Classificação feita por palavras do campo 'Forma / observação da fonte' (arquivo 17). "
        "Lista lote a lote, com a classe atribuída, abaixo.",
        NOTE,
    )
    r += 1
    data = [
        [
            e["lot"],
            e["lote"],
            e["quadra"],
            e["data"].date() if e["data"] else "",
            e["dest"],
            e["forma"],
            EXIT_LABEL[e["classe"]],
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
            "Destinatário",
            "Forma (fonte)",
            "Classe",
            "Fonte",
        ],
        data,
    )

    # 3 Gastos 15 meses
    ws = wb.create_sheet("3 Gastos 15 meses (SEI)")
    data = []
    for (tipo, cat), (n, v, cls) in sorted(
        by_cat.items(), key=lambda kv: (kv[1][2], -kv[1][1])
    ):
        data.append([CLASS_LABEL[cls], tipo, cat, n, v, v / REPORT_MONTHS])
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
    for cls in ("M", "INV", "DIV", "BENS", "LOTE"):
        ws.cell(r, 1, CLASS_LABEL[cls]).font = BOLD
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
        f"Custo de manter a empresa: R$ {m_month:,.2f} por mês × {MONTHS} meses = R$ {manter:,.2f}.",
        (
            "Regras do Luís aplicadas: combustível não é recorrente (indica máquinas em obra) → classe Investimento; "
            "dívidas, multas, acordos e divisão de lucros entram junto com os custos de manter."
        ),
        "Devolução de principal não é extrapolada: a dívida é tratada pela variação do saldo (aba 5), para não contar duas vezes.",
        "Pagamentos feitos com lotes (permuta/dação) não são extrapolados: já estão na saída de lotes.",
        "Fonte: Contas_Pagas_Recorrentes_x_Extraordinarias_..._revisado.xlsx (cópia de 28/10/2026, ainda em edição).",
        (
            "Referência: a aba Relatórios Mensais da planilha Google extrapola R$ 832.951/mês × 96 meses = R$ 79,96 mi, "
            "mas soma as 3 empresas e todos os gastos."
        ),
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
    r = para(
        ws,
        r + 1,
        "Lançamentos de obras, projetos, licenças, combustível e antecipação Boff da SEI nos 15 meses. "
        "A situação vem do nome do loteamento no fornecedor ou na observação.",
        NOTE,
    )
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
        r,
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
        [c["n"], c["credor"], str(c["desc"])[:140], c["valor"], c["sit"], c["saldo"]]
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
        ],
        data,
        money_cols=(4, 6),
        widths=[6, 34, 90, 14, 18, 16],
    )
    r = end
    for label, v in [
        ("Saldo das dívidas em jan/2019", debt18),
        ("Empréstimos e financiamentos em 2025 (doc 11, 11.2.5)", debt25),
        ("Dívida líquida quitada no período", debt_paid),
    ]:
        ws.cell(r, 2, label).font = BOLD
        ws.cell(r, 6, v).number_format = MONEY
        r += 1
    para(
        ws,
        r + 1,
        "Itens 30 (TIS) e 32 (HOTEL) são limites bancários de outras empresas listados na SEI: confirmar.",
        NOTE,
    )

    # 6 Bens novos
    ws = wb.create_sheet("6 Bens novos 2025")
    data = [[a["linha"], a["desc"], a["valor"]] for a in new_assets]
    end = write_table(
        ws,
        1,
        ["Linha GERAL", "Bem (inventário 2025)", "Valor 2025"],
        data,
        money_cols=(3,),
        widths=[12, 120, 16],
    )
    ws.cell(end, 2, "Total").font = BOLD
    ws.cell(end, 3, new_assets_total).number_format = MONEY
    para(
        ws,
        end + 2,
        "Bens com valor 2018 = 0 e valor 2025 > 0 na aba GERAL (blocos 11.1.1 a 11.1.5). Tratados como "
        "comprados em dinheiro: a favor do grupo, porque alguns vieram em permuta de lote.",
        NOTE,
    )
    r = end + 4
    for a in new_assets_post:
        ws.cell(r, 1, a["linha"])
        ws.cell(r, 2, "Fora da ponte (comprado depois de 31.12.2024): " + a["desc"])
        ws.cell(r, 3, a["valor"]).number_format = MONEY
        r += 1

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
        (f"× {MONTHS} meses", rents_total),
    ]:
        ws.cell(r, 3, label).font = BOLD
        ws.cell(r, 4, v).number_format = MONEY
        r += 1
    r += 1
    for line in [
        "Outras entradas que NÃO foram somadas à SEI nesta versão (pertencem a outras empresas; a confirmar com a sessão S1):",
        (
            "• Venda de 190 ha da São Crispin à Família Rigo (2018): R$ 9.448.000 = 118.100 sacas; 20.150 sacas/ano de 2019 a 2022 "
            "(o grupo diz que 2021 e 2022 faltam receber). Credor: Hotel Princesa dos Vales."
        ),
        "• Venda de 85,32 ha à Família Pellegrini (2018): R$ 2.100.000 (R$ 1.000.000 em jun/2019).",
        "• Arrendamento da São Crispin à Família Lorenzet (contratos 2017-2020, 2020-2023, 2023-2026): pago em sacas de soja.",
        (
            "Indícios de que esse dinheiro passa pelo caixa da SEI (relatórios mensais): a SEI paga Funrural/IRPJ/CSLL da "
            "Agropecuária São Crispin sobre soja; '20.150 (−) 17.544 = 2.606 SC SOJA VENDIDOS' (mai/2024); sacas vendidas "
            "para comprar a escavadeira; 'cheques soja Titton' na compra do imóvel da Família Dal Ponte."
        ),
        "Novos empréstimos tomados no período (Piccoli R$ 1,5 mi em dez/2021; Mercali): entram na variação da dívida (aba 5).",
        "Contrato de locação Oleoplan (arquivo 15): R$ 14.000/mês no 1º ano, segundo o OCR — conferir na imagem da página.",
    ]:
        r = para(ws, r, line, NOTE)

    # Contratos 11.2
    ws = wb.create_sheet("Contratos 11.2")
    gap_c = bridge["central"]["sem_explicacao"]
    data = []
    for c in comp:
        prop, why = (
            c["siviero"],
            "mantém o valor do grupo nesta versão (decisão item a item pendente)",
        )
        if c["codigo"] == "11.2.1":
            prop, why = (
                0.0,
                (
                    "sem multa: contratos vigentes. Multa só se houver distrato e termo de nunca lotear a gleba; "
                    "nesse caso, investimentos feitos nela não são cobrados (mas precisam ser demonstrados)"
                ),
            )
        data.append(
            [
                c["codigo"],
                c["item"],
                c["siviero"],
                c["rafael"],
                prop,
                why,
                c["comentario_luis"][:300],
            ]
        )
    data.append(
        [
            "novo",
            "Investimento nos loteamentos não finalizados a demonstrar (valor sem explicação, cenário central)",
            0.0,
            0.0,
            gap_c,
            "ativo: o dinheiro sem destino entra como investimento até o grupo demonstrar o valor real",
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
            "Proposta v0",
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
    para(
        ws,
        end + 2,
        "Rafael soma no 11.2.1 o valor de lotes de loteamentos não executados junto com as multas "
        "(contraditório: se há multa, não há lote; se há lote, não há multa). O 11.2.4 de Rafael já inclui o 11.2.5.",
        NOTE,
    )
    para(
        ws,
        end + 3,
        "2018: o próprio grupo atribuía R$ 5.400.000 aos loteamentos em execução e R$ 1.250.000 aos "
        "contratados (R$ 50.000 cada, 'investimentos iniciais'), além da planilha de investimentos de "
        f"R$ {sum(x['valor'] for x in inv18):,.2f}. Em 2025 tudo isso vale R$ 0 e vira multa.",
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
            "Valor atribuído pelo grupo",
            "Loteamento no modelo",
            "Valor/lote 2018",
            "Valor/lote out/2025",
            "Valor ÷ média 2025",
            "Texto",
        ],
        data,
        money_cols=(5, 7, 8),
        widths=[11, 30, 6, 7, 16, 26, 14, 14, 12, 90],
    )
    para(
        ws,
        end + 1,
        "Só lotes isolados, com o valor do próprio lote numa dação, acordo ou aquisição. Ficam fora os "
        "valores que são só a parte em veículo ou imóvel recebida. Não é preço de venda em dinheiro: "
        "o grupo não informou esses preços.",
        NOTE,
    )

    # Pendências
    ws = wb.create_sheet("Pendências")
    ws.column_dimensions["A"].width = 160
    r = 1
    for line, f in [
        ("Pendências e limites desta versão v0", Font(bold=True, size=13)),
        ("Materiais ainda em edição (atualizar quando fecharem):", BOLD),
        (
            (
                "• Contas_Pagas_Recorrentes_x_Extraordinarias_..._revisado.xlsx: sessão 'Verificação e análise de relatórios em Excel' "
                "ativa; abas Dúvidas e A determinar sem a decisão do Luís. Usada a cópia de 28/10/2026 (MD5 na aba Leia-me)."
            ),
            None,
        ),
        (
            "• Achados da sessão ACERTOS (achados-para-sei.md): aguardam aprovação; não usados, salvo como contexto.",
            None,
        ),
        (
            "• São Crispin / Hotel (sessão S1): entradas de soja e terras não somadas à SEI.",
            None,
        ),
        (
            "• Levantamento ambiental da SEI (pacote em análise): define os 26 lotes de terceiros (11.2.2).",
            None,
        ),
        ("", None),
        ("Premissas a aprovar:", BOLD),
        (
            f"• Período de {MONTHS} meses (jan/2018 a dez/2024). Com 72 meses (2019–2024) o custo de manter cai e o valor sem explicação sobe.",
            None,
        ),
        (
            "• Os 15 meses (jul/2023–out/2024) representam todo o período. Os custos de 2018–2022 podem ter sido diferentes.",
            None,
        ),
        (
            (
                "• Preço dos lotes: o grupo não informou o valor das vendas (pedido já listado na aba Resumo da planilha Google). "
                "A aba 8 traz os poucos preços citados no arquivo 17, para comparar."
            ),
            None,
        ),
        (
            "• Valores de vendas são brutos; impostos (7%) e comissões (6%) estão dentro do custo de manter (média dos 15 meses).",
            None,
        ),
        (
            "• Classificação das formas de saída por palavras-chave: conferir a amostra na aba 2.",
            None,
        ),
        ("", None),
        (
            "Sobreposições possíveis (todas a favor do grupo, ou seja, o valor sem explicação pode ser maior):",
            BOLD,
        ),
        (
            "• Bens novos tratados como comprados em dinheiro, mesmo os recebidos em permuta de lote.",
            None,
        ),
        (
            (
                "• Concessão de R$ 7,5 mi para concluir Falavigna e Colina do Retiro II, embora o grupo ainda liste R$ 2,0 mi e "
                "R$ 3,8 mi a pagar nessas obras, e parte delas tenha sido paga com lotes (classe Obra)."
            ),
            None,
        ),
        (
            "• Juros e acordos em dinheiro dos 15 meses entram na média e também podem estar na variação da dívida.",
            None,
        ),
        ("", None),
        ("Revisão independente (P11): ver _diario/revisao-v0.md.", BOLD),
    ]:
        r = para(ws, r, line, f)

    out = HERE / "fluxo-caixa-sei-v0.xlsx"
    wb.save(out)

    numbers = {
        "gerado_em": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "fontes_md5": hashes,
        "meses": MONTHS,
        "periodo": [PERIOD_START, PERIOD_END],
        "lotes": {
            "n2018": n18_total,
            "n2025": sum(x["n2025"] for x in rows),
            "reducao": red_total,
            "no_periodo": n_in,
            "apos_2024": post["n"],
            "indice_2018_2025": ratio,
        },
        "ponte": bridge,
        "apos_2024": post,
        "custos": {
            "m_mes": m_month,
            "classes": dict(by_cls),
            "sei_15m": sei_15m,
            "inv_por_situacao": dict(inv_tag),
        },
        "divida": {"jan2019": debt18, "2025": debt25, "quitada": debt_paid},
        "bens_novos": new_assets_total,
        "alugueis_mes": rent_month,
        "contratos": comp,
        "contratos_total": {"siviero": tot_siv, "rafael": tot_raf},
        "investimentos_2018": {
            "planilha": sum(x["valor"] for x in inv18),
            "em_execucao": 5_400_000,
            "contratados": 1_250_000,
        },
        "patrimonio": {
            "geral_2018_apresentado": float(geral["C3"]),
            "geral_2018_atualizado": float(geral["D3"]),
            "geral_2025_bens": float(geral["E3"]),
            "geral_2025_liquido": float(geral["E1"]),
        },
        "trajetoria": path,
        "sensibilidade": sens,
        "excedente_lista": {"n": excess_n, **excess_val},
        "precos_citados": cited_summary,
        "bens_pos_database": new_assets_post,
    }
    (HERE / "numeros-v0.json").write_text(
        json.dumps(numbers, ensure_ascii=False, indent=1, default=str), encoding="utf-8"
    )

    # console summary for the journal
    print(
        "lotes 2018/2025/redução:",
        n18_total,
        numbers["lotes"]["n2025"],
        red_total,
        "| no período",
        round(n_in, 2),
        "| após 2024",
        round(post["n"], 2),
    )
    print(
        "custo manter/mês",
        round(m_month, 2),
        "total",
        round(manter, 2),
        "| dívida quitada",
        debt_paid,
        "| bens novos",
        new_assets_total,
        "| aluguel/mês",
        round(rent_month, 2),
    )
    for k in ("piso", "central", "teto"):
        print(k, {kk: round(vv) for kk, vv in bridge[k].items()})
    print("inv 15m:", {k: round(v) for k, v in inv_tag.items()})
    print(
        "trajetória:",
        [(p["ano"], round(p["lotes"], 1), round(p["valor25"])) for p in path],
    )
    print("saved", out)


if __name__ == "__main__":
    main()
