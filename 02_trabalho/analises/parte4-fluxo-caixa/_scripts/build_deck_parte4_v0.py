"""Build the simplified Part 4 deck (python-pptx, native charts) from numeros-v0.json and the v0 workbook.

Every figure comes from those two files, which build_fluxo_caixa.py derives from the sources.
Style follows 03_entrega/apresentacao/_build/build_deck.py (v1 deck) so both decks read as one.
Usage: python -I _scripts/build_deck_parte4.py <output.pptx>
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent.parent
OUT = Path(sys.argv[1])
N = json.loads((HERE / "numeros-v0.json").read_text(encoding="utf-8"))
WB = openpyxl.load_workbook(HERE / "fluxo-caixa-sei-v0.xlsx", data_only=True)
REF = "teto"  # reference price scenario shown in the headline (premise to be approved)
VERSION = "Rascunho v0 · Parte 4 da SEI · 28/10/2026"

DARK = RGBColor(0x1F, 0x3A, 0x2E)
MID = RGBColor(0x4E, 0x7A, 0x5E)
TINT = RGBColor(0xEC, 0xF1, 0xED)
WARN = RGBColor(0xB8, 0x5A, 0x22)
WARN_TINT = RGBColor(0xFA, 0xEE, 0xE6)
TEXT = RGBColor(0x22, 0x2B, 0x27)
MUTED = RGBColor(0x5E, 0x6A, 0x64)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRID = RGBColor(0xD9, 0xDE, 0xDB)
HEAD, BODY = "Cambria", "Calibri"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
prs.core_properties.title = "SEI — para onde foi o dinheiro dos lotes (rascunho v0)"
prs.core_properties.author = "Luís Siviero (rascunho)"
state = {"n": 0}


# ------------------------------------------------------------------ formatting helpers
def mi(v: float, dec: int = 1) -> str:
    s = f"{abs(v) / 1e6:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if v < 0 else "") + f"R$ {s} mi"


def mil(v: float) -> str:
    s = f"{v / 1e3:,.0f}".replace(",", ".")
    return f"R$ {s} mil"


def num(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def nb(t: str) -> str:
    """Keep currency, figures and units together on one line."""
    t = t.replace("R$ ", "R$\u00a0").replace("− R$", "−R$")
    return re.sub(r"(\d) (mil|mi|ha|lotes|meses)\b", "\\1\u00a0\\2", t)


def text(
    slide,
    x,
    y,
    w,
    h,
    paras,
    size=16,
    color=TEXT,
    bold=False,
    font=BODY,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    space_after=6,
    margin=0.05,
):
    """paras: str | list of (str | list of (text, {opts}))"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    for i, p in enumerate([paras] if isinstance(paras, str) else paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_after = Pt(space_after)
        for t, o in p if isinstance(p, list) else [(p, {})]:
            r = para.add_run()
            r.text = nb(t)
            f = r.font
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", False)
            f.color.rgb = o.get("color", color)
    return tb


def box(slide, x, y, w, h, fill=TINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.25)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def badge(slide, x, y, label, d=0.55, fill=DARK, size=18):
    c = box(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    tf = c.text_frame
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = (
        HEAD,
        Pt(size),
        True,
        WHITE,
    )


def tag(slide, x, y, label="PROVISÓRIO", w=1.6):
    s = box(slide, x, y, w, 0.36, fill=WARN)
    s.adjustments[0] = 0.5
    tf = s.text_frame
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = BODY, Pt(11), True, WHITE


def new_slide(title, source=None, provisional=True, dark=False, notes=""):
    state["n"] += 1
    s = prs.slides.add_slide(BLANK)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    if dark:
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = DARK
        return s
    text(
        s,
        0.6,
        0.35,
        10.4 if provisional else 12.1,
        0.9,
        title,
        size=28,
        color=DARK,
        bold=True,
        font=HEAD,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    if provisional:
        tag(s, 11.13, 0.62)
    foot = VERSION + (" · Fonte: " + source if source else "")
    text(s, 0.6, 7.0, 11.4, 0.35, foot, size=10, color=MUTED)
    text(
        s,
        12.1,
        7.0,
        0.65,
        0.35,
        str(state["n"]),
        size=10,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )
    return s


def table(slide, x, y, col_w, rows, size=13, row_h=0.42, bold_rows=(), warn_rows=()):
    shape = slide.shapes.add_table(
        len(rows),
        len(col_w),
        Inches(x),
        Inches(y),
        Inches(sum(col_w)),
        Inches(row_h * len(rows)),
    )
    t = shape.table
    for j, w in enumerate(col_w):
        t.columns[j].width = Inches(w)
    for i, row in enumerate(rows):
        t.rows[i].height = Inches(row_h)
        for j, v in enumerate(row):
            cell = t.cell(i, j)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = (
                DARK
                if i == 0
                else (
                    WARN_TINT if i in warn_rows else (TINT if i in bold_rows else WHITE)
                )
            )
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = nb(str(v))
            r.font.name, r.font.size = BODY, Pt(size)
            r.font.bold = i == 0 or i in bold_rows
            r.font.color.rgb = WHITE if i == 0 else TEXT
    return t


def style_chart(chart, legend=False):
    chart.font.name = BODY
    chart.font.size = Pt(12)
    chart.font.color.rgb = MUTED
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(11)
    chart.category_axis.format.line.color.rgb = GRID
    chart.category_axis.tick_labels.font.size = Pt(11)


B = N["ponte"][REF]
BP, BC = N["ponte"]["piso"], N["ponte"]["central"]
months = N["meses"]
ws1 = WB["1 Lotes por loteamento"]
hdr = [c.value for c in ws1[1]]
cls_cols = {
    k: hdr.index(k) for k in ("Venda", "Obra", "Acordo", "Bem", "Sócio", "Sem info")
}
cls_n = dict.fromkeys(cls_cols, 0.0)
for row in ws1.iter_rows(min_row=2, values_only=True):
    if row[0] and row[0] != "Total" and isinstance(row[2], (int, float)):
        for k, j in cls_cols.items():
            cls_n[k] += row[j] or 0
lots_in = N["lotes"]["no_periodo"]
assert abs(sum(cls_n.values()) - lots_in) < 1e-6, (cls_n, lots_in)
cost = N["custos"]["classes"]

# ------------------------------------------------------------------ 1. cover
s = new_slide(
    "",
    dark=True,
    notes=(
        "Rascunho para discussão. Este estudo segue o dinheiro: quanto entrou com a venda de lotes de 2018 a 2024, "
        "quanto o grupo gastou para manter a empresa e pagar dívidas, e quanto sobra sem explicação."
    ),
)
text(
    s,
    0.9,
    1.7,
    11.5,
    1.2,
    "Para onde foi o dinheiro dos lotes vendidos?",
    size=40,
    color=WHITE,
    bold=True,
    font=HEAD,
)
text(
    s,
    0.9,
    2.95,
    11.5,
    0.8,
    "SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS · 2018 a 2024",
    size=22,
    color=TINT,
)
text(
    s,
    0.9,
    3.75,
    11.5,
    0.8,
    "Parte 4 — Contratos | Indenizações | Compromissos | Empréstimos",
    size=18,
    color=TINT,
)
text(
    s,
    0.9,
    5.9,
    11.5,
    0.5,
    "Rascunho v0 para discussão · 28/10/2026 · valores provisórios",
    size=14,
    color=TINT,
)

# ------------------------------------------------------------------ 2. headline
s = new_slide(
    "O patrimônio caiu e o dinheiro não aparece",
    source="fluxo-caixa-sei-v0.xlsx, aba Ponte",
    notes=(
        f"Cenário de referência: preço dos lotes igual à média que o próprio grupo usa em out/2025. "
        f"No preço de 2018 do inventário, a conta fecha negativa ({mi(BP['sem_explicacao'])}); "
        f"no cenário intermediário, {mi(BC['sem_explicacao'])}. A diferença vem do preço de venda, "
        "que o grupo ainda não informou."
    ),
)
cards = [
    (
        f"{num(lots_in)}",
        "lotes saíram do estoque de 2018 a 2024",
        f"mais {num(N['lotes']['apos_2024'])} em 2025",
    ),
    (
        mi(B["total"]),
        "valem esses lotes no preço do próprio grupo",
        f"{mi(B['VENDA'])} vendidos em dinheiro",
    ),
    (
        mi(B["usos"]),
        "explicam os gastos, dívidas e bens comprados",
        f"{months} meses de custo de manter a empresa",
    ),
    (mi(B["sem_explicacao"]), "ficam sem explicação", f"1/6 = {mi(B['um_sexto'])}"),
]
for i, (big, lab, sub) in enumerate(cards):
    x = 0.6 + i * 3.08
    fill = WARN_TINT if i == 3 else TINT
    box(s, x, 1.65, 2.85, 3.0, fill=fill)
    text(
        s,
        x + 0.2,
        1.85,
        2.5,
        0.9,
        big,
        size=30,
        color=WARN if i == 3 else DARK,
        bold=True,
        font=HEAD,
    )
    text(s, x + 0.2, 2.85, 2.5, 1.0, lab, size=15)
    text(s, x + 0.2, 3.85, 2.5, 0.7, sub, size=12, color=MUTED)
text(
    s,
    0.6,
    5.0,
    12.1,
    1.8,
    [
        [
            (
                "Se a empresa diz que não tem caixa, esse dinheiro foi investido em algo — ",
                {},
            ),
            (
                "provavelmente nos loteamentos que o grupo avalia em R$ 0 e dos quais ainda desconta multas.",
                {"bold": True},
            ),
        ],
        (
            "Ou o grupo mostra os valores reais e a diferença é recalculada, ou 1/6 do valor sem explicação volta ao sócio "
            "que sai."
        ),
    ],
    size=16,
)

# ------------------------------------------------------------------ 3. the contradiction
inv18 = N["investimentos_2018"]
c112 = {c["codigo"]: c for c in N["contratos"]}
s = new_slide(
    "“R$ 0 + multa” não fecha com o que o próprio grupo dizia em 2018",
    source="inventário 2018 (doc 18); doc 11; Estudo 2 de Rafael; aba GERAL",
    notes=(
        "Em 2018 o grupo atribuía valor de investimento aos loteamentos em execução e contratados. "
        "Em 2025 os mesmos empreendimentos valem zero e geram multa. Rafael mistura lotes executados "
        "com multas, o que é contraditório: se há multa por distrato, não há lote; se há lote, não há multa."
    ),
)
cols = [
    (
        "2018 · o grupo",
        [
            f"Loteamentos em execução: {mi(inv18['em_execucao'])}",
            f"Contratados: {mi(inv18['contratados'], 2)} (R$ 50 mil cada, “investimentos iniciais”)",
            f"Planilha de investimentos pagos: {mi(inv18['planilha'])}",
        ],
        TINT,
        DARK,
    ),
    (
        "2025 · o grupo",
        [
            "Os mesmos loteamentos: R$ 0",
            f"Multas rescisórias descontadas: {mi(c112['11.2.1']['siviero'])}",
            "Nenhum investimento do período aparece",
        ],
        WARN_TINT,
        WARN,
    ),
    (
        "2025 · Rafael",
        [
            f"Áreas contratadas: {mi(c112['11.2.1']['rafael'])}",
            "Soma valor de lotes executados com multas",
            "Contraditório: lote e multa se excluem",
        ],
        TINT,
        MID,
    ),
]
for i, (head, items, fill, col) in enumerate(cols):
    x = 0.6 + i * 4.1
    box(s, x, 1.6, 3.85, 3.6, fill=fill)
    text(s, x + 0.25, 1.75, 3.4, 0.5, head, size=20, color=col, bold=True, font=HEAD)
    text(s, x + 0.25, 2.4, 3.4, 2.7, items, size=15, space_after=10)
text(
    s,
    0.6,
    5.5,
    12.1,
    1.3,
    [
        [
            (
                "Em vez de avaliar cada loteamento (não sabemos quanto foi investido em cada um), ",
                {},
            ),
            ("seguimos o dinheiro: ", {"bold": True}),
            (
                "o que entrou com os lotes, o que saiu para manter a empresa e o que sobra.",
                {},
            ),
        ],
    ],
    size=16,
)

# ------------------------------------------------------------------ 4. patrimony over time
path = N["trajetoria"]
s = new_slide(
    "O estoque de lotes perdeu quase metade do valor de 2018 a 2024",
    source="aba 'Lotes que saíram 2018-2025' e arquivo 17; preços de out/2025 do grupo",
    notes=(
        "Estoque de lotes executados, todos avaliados no mesmo preço (média do grupo em out/2025), para "
        "que a queda mostre só os lotes que saíram e não a variação de preço. Saídas sem data foram "
        "distribuídas igualmente entre 2018 e 2024."
    ),
)
cd = CategoryChartData()
cd.categories = [p["ano"].replace("fim ", "") for p in path]
cd.add_series(
    "Estoque de lotes (R$ mi, preços de out/2025)",
    [round(p["valor25"] / 1e6, 1) for p in path],
)
gf = s.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS, Inches(0.6), Inches(1.5), Inches(8.3), Inches(5.3), cd
)
ch = gf.chart
style_chart(ch)
ser = ch.series[0]
ser.format.line.color.rgb = DARK
ser.format.line.width = Pt(3)
ser.marker.format.fill.solid()
ser.marker.format.fill.fore_color.rgb = DARK
ser.smooth = False
pl = ch.plots[0]
pl.has_data_labels = True
pl.data_labels.number_format = "#,##0.0"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.ABOVE
pl.data_labels.font.size = Pt(11)
ch.value_axis.minimum_scale = 0
ch.value_axis.tick_labels.number_format = "#,##0"
ch.value_axis.tick_labels.number_format_is_linked = False
first, end24, last = path[0], next(p for p in path if p["ano"] == "fim 2024"), path[-1]
box(s, 9.2, 1.6, 3.5, 5.1, fill=TINT)
text(
    s,
    9.45,
    1.8,
    3.1,
    4.8,
    [
        [("Início de 2018", {"bold": True, "color": DARK})],
        f"{num(first['lotes'])} lotes · {mi(first['valor25'])}",
        [("Fim de 2024 (data-base)", {"bold": True, "color": DARK})],
        f"{num(end24['lotes'])} lotes · {mi(end24['valor25'])}",
        [("Out/2025 (doc 11)", {"bold": True, "color": DARK})],
        f"{num(last['lotes'])} lotes · {mi(last['valor25'])}",
        [
            ("Queda até a data-base: ", {"bold": True}),
            (mi(first["valor25"] - end24["valor25"]), {"bold": True, "color": WARN}),
        ],
    ],
    size=15,
    space_after=8,
)

# ------------------------------------------------------------------ 5. how lots left
s = new_slide(
    f"Como saíram os {num(lots_in)} lotes de 2018 a 2024",
    source="arquivo 17 (forma de cada saída), classificado por palavras-chave",
    notes=(
        "Quantidade de lotes por forma de saída. Só a venda em dinheiro gera caixa. As demais formas têm "
        "destino declarado: pagaram obras e serviços, acordos e indenizações, foram trocadas por bens ou "
        "entregues a sócios e família. Lotes sem informação não têm destino documentado (premissa P4)."
    ),
)
labels = {
    "Venda": "Venda em dinheiro",
    "Obra": "Pagamento de obras e serviços",
    "Acordo": "Acordos, indenizações e distratos",
    "Bem": "Troca por bem ou área",
    "Sócio": "Sócios e família",
    "Sem info": "Sem destino documentado",
}
order = sorted(cls_n, key=lambda k: -cls_n[k])
cd = CategoryChartData()
cd.categories = [labels[k] for k in order]
cd.add_series("Lotes", [round(cls_n[k], 1) for k in order])
gf = s.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(1.5), Inches(8.0), Inches(5.3), cd
)
ch = gf.chart
style_chart(ch)
ch.category_axis.reverse_order = True
ch.value_axis.tick_labels.number_format = "0"
ch.value_axis.tick_labels.number_format_is_linked = False
pl = ch.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
pl.data_labels.number_format = "0"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
for i, k in enumerate(order):
    pt = ch.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = (
        DARK if k == "Venda" else (WARN if k == "Sem info" else MID)
    )
box(s, 8.9, 1.6, 3.8, 5.1, fill=TINT)
text(
    s,
    9.15,
    1.8,
    3.4,
    4.8,
    [
        [("No preço do grupo:", {"bold": True, "color": DARK})],
        f"Venda em dinheiro: {mi(B['VENDA'])}",
        f"Obras e serviços: {mi(B['OBRA'])}",
        f"Acordos e indenizações: {mi(B['ACORDO'])}",
        f"Troca por bens: {mi(B['BEM'])}",
        f"Sócios e família: {mi(B['SOC'])}",
        [
            ("Sem destino: ", {"bold": True}),
            (mi(B["SEM"]), {"bold": True, "color": WARN}),
        ],
    ],
    size=15,
    space_after=8,
)

# ------------------------------------------------------------------ 6. cost of keeping the company
s = new_slide(
    "Manter a empresa custa cerca de " + mil(N["custos"]["m_mes"]) + " por mês",
    source="relatórios mensais de contas pagas, jul/2023 a out/2024 (15 meses), só a SEI",
    notes=(
        "Média dos 15 meses de contas pagas da SEI. Entram os gastos recorrentes (folha, pró-labore, "
        "impostos, juros, comissões, divisão de aluguéis) e também dívidas, multas, acordos e divisão de "
        "lucros. Combustível não é recorrente: indica máquinas em obra, por isso vai para investimento. "
        "Devoluções de empréstimo são tratadas pela variação da dívida, e pagamentos com lotes já estão "
        "na saída de lotes."
    ),
)
cats = [
    ("Manter a empresa", cost["M"]),
    ("Obras, projetos e combustível", cost["INV"]),
    ("Compra de bens", cost["BENS"]),
    ("Pago com lotes", cost["LOTE"]),
    ("Devolução de empréstimos", cost["DIV"]),
]
cd = CategoryChartData()
cd.categories = [c for c, _ in cats]
cd.add_series("R$ mi em 15 meses", [round(v / 1e6, 2) for _, v in cats])
gf = s.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.6),
    Inches(1.5),
    Inches(7.6),
    Inches(5.3),
    cd,
)
ch = gf.chart
style_chart(ch)
ch.value_axis.tick_labels.number_format = "0.0"
ch.value_axis.tick_labels.number_format_is_linked = False
pl = ch.plots[0]
pl.gap_width = 70
pl.has_data_labels = True
pl.data_labels.number_format = "#,##0.00"
pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
for i in range(len(cats)):
    pt = ch.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = DARK if i == 0 else MID
box(s, 8.5, 1.6, 4.2, 5.1, fill=TINT)
text(
    s,
    8.75,
    1.8,
    3.8,
    4.8,
    [
        [("Total pago pela SEI em 15 meses", {"bold": True, "color": DARK})],
        mi(N["custos"]["sei_15m"], 2),
        [("Manter a empresa", {"bold": True, "color": DARK})],
        f"{mi(cost['M'], 2)} ÷ 15 = {mil(N['custos']['m_mes'])}/mês",
        [("Estendido ao período", {"bold": True, "color": DARK})],
        f"{mil(N['custos']['m_mes'])} × {months} meses = {mi(B['manter'])}",
        [
            ("Regra aplicada: ", {"bold": True}),
            ("combustível conta como obra, não como custo fixo.", {}),
        ],
    ],
    size=15,
    space_after=8,
)

# ------------------------------------------------------------------ 7. bridge (waterfall)
s = new_slide(
    "A conta, passo a passo, no preço do próprio grupo",
    source="fluxo-caixa-sei-v0.xlsx, aba Ponte (cenário teto = média do grupo em out/2025)",
    notes=(
        "Entradas em verde, saídas em terracota. A concessão para concluir Falavigna e Colina do Retiro II "
        "usa o custo de R$ 50 mil por lote que o próprio grupo adota nas estimativas dos loteamentos "
        "contratados; é um teto a favor do grupo. Dívidas: saldo de jan/2019 menos os empréstimos que "
        "aparecem em 2025."
    ),
)
steps = [
    ("Lotes vendidos em dinheiro", B["VENDA"]),
    ("Aluguéis", B["alugueis"]),
    ("Lotes sem destino", B["SEM"]),
    ("Manter a empresa", -B["manter"]),
    ("Dívidas de 2018 pagas", -B["divida"]),
    ("Bens comprados", -B["bens_novos"]),
    ("Conclusão de loteamentos", -B["conclusao"]),
]
base, bar, run = [], [], 0.0
for _, v in steps:
    if v >= 0:
        base.append(run)
        bar.append(v)
        run += v
    else:
        run += v
        base.append(run)
        bar.append(-v)
assert run >= 0 and abs(run - B["sem_explicacao"]) < 1.0, (run, B["sem_explicacao"])
base.append(0.0)
bar.append(run)
cd = CategoryChartData()
cd.categories = [c for c, _ in steps] + ["Sem explicação"]
cd.add_series("base", [round(v / 1e6, 1) for v in base])
cd.add_series("valor", [round(v / 1e6, 1) for v in bar])
gf = s.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED,
    Inches(0.6),
    Inches(1.45),
    Inches(12.1),
    Inches(4.4),
    cd,
)
ch = gf.chart
style_chart(ch)
ch.value_axis.tick_labels.number_format = "0"
ch.value_axis.tick_labels.number_format_is_linked = False
ch.category_axis.tick_labels.font.size = Pt(10)
ch.plots[0].gap_width = 40
ch.series[0].format.fill.background()
ch.series[0].format.line.fill.background()
vs = ch.series[1]
vs.data_labels.show_value = True
vs.data_labels.number_format = "#,##0.0"
vs.data_labels.number_format_is_linked = False
vs.data_labels.position = XL_LABEL_POSITION.INSIDE_END
vs.data_labels.font.size = Pt(11)
vs.data_labels.font.color.rgb = WHITE
for i, (_, v) in enumerate([*steps, ("", 0)]):
    pt = vs.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = (
        DARK if i == len(steps) else (MID if v >= 0 else WARN)
    )
text(
    s,
    0.6,
    5.95,
    12.1,
    0.9,
    [
        [
            ("R$ milhões. ", {"color": MUTED}),
            ("Resultado: ", {"bold": True}),
            (mi(B["sem_explicacao"]), {"bold": True, "color": WARN}),
            (
                (
                    f" sem explicação · 1/6 = {mi(B['um_sexto'])} · somando os lotes dados em obras, investimento a demonstrar de "
                    f"{mi(B['investimento_a_demonstrar'])}"
                ),
                {},
            ),
        ],
    ],
    size=15,
)

# ------------------------------------------------------------------ 8. which price
pc = N["precos_citados"]
sens = N["sensibilidade"]
s = new_slide(
    "O resultado depende do preço de venda, que o grupo não informou",
    source="abas Ponte, Sensibilidade e 8 Preços citados",
    notes=(
        "O grupo não informou por quanto vendeu cada lote. Nas dações e aquisições de lote isolado, o próprio "
        "grupo atribuiu aos lotes valores próximos ou acima da média de out/2025. Por isso o preço do grupo "
        "de 2025 é a referência deste rascunho; a decisão final depende dos valores reais."
    ),
)
rows = [["Preço dos lotes", "Valor sem explicação", "1/6"]]
for k, lab in (
    ("piso", "Valor de 2018 do inventário"),
    ("central", "Intermediário (pelo ano da venda)"),
    ("teto", "Média do grupo em out/2025"),
):
    v = N["ponte"][k]["sem_explicacao"]
    rows.append([lab, mi(v), mi(v / 6) if v > 0 else "—"])
table(s, 0.6, 1.6, [3.9, 2.3, 1.7], rows, size=14, row_h=0.55, bold_rows=(3,))
box(s, 8.9, 1.6, 3.8, 2.2, fill=WARN_TINT)
text(
    s,
    9.15,
    1.75,
    3.4,
    2.0,
    [
        [("Nas dações, o próprio grupo", {"bold": True, "color": WARN})],
        (
            f"atribuiu aos lotes {num(round(pc['mediana_sobre_v25'] * 100))}% da média de 2025 (mediana de "
            f"{pc['n']} lotes; {pc['acima_de_v25']} acima dela)."
        ),
    ],
    size=15,
)
r84 = next(x for x in sens if x["meses"] == 72 and not x["concessao"])
text(
    s,
    0.6,
    4.1,
    12.1,
    2.7,
    [
        [
            (
                "Premissas que mudam o número (no preço do grupo):",
                {"bold": True, "color": DARK},
            )
        ],
        f"Contando 72 meses (2019–2024) e sem a concessão de conclusão: {mi(r84['teto'])} sem explicação.",
        (
            f"{num(N['excedente_lista']['n'])} lotes do arquivo 17 saíram além da redução do inventário: até "
            f"{mi(N['excedente_lista']['teto'])} a mais, se eram da SEI (fora da conta até conferir)."
        ),
        (
            f"{num(N['lotes']['apos_2024'])} lotes saíram em 2025, depois da data-base: voltam ao patrimônio de 31/12/2024 "
            f"ou o grupo informa o que recebeu ({mi(N['apos_2024']['teto'])} no preço do grupo)."
        ),
    ],
    size=15,
    space_after=8,
)

# ------------------------------------------------------------------ 9. other income shown separately
s = new_slide(
    "Outras entradas, mostradas à parte",
    source="inventário 2018 (aluguéis); docx do Hotel; relatórios mensais; contratos Lorenzet",
    notes=(
        "Só os aluguéis da SEI entram na conta, pela relação de aluguéis de 2018. As vendas de terras e a "
        "soja são de outras empresas do grupo (Hotel Princesa dos Vales / São Crispin) e não foram somadas. "
        "Os relatórios mensais mostram que esse dinheiro passa pelo caixa da SEI, o que precisa ser "
        "esclarecido."
    ),
)
items = [
    (
        "1",
        "Aluguéis da SEI",
        (
            f"{mil(N['alugueis_mes'])}/mês na relação de 2018 → {mi(B['alugueis'])} em {months} meses. "
            "Entram na conta."
        ),
        DARK,
    ),
    (
        "2",
        "Venda de terras da São Crispin (2018)",
        (
            "Rigo: R$ 9,4 mi em 118.100 sacas de soja, 20.150 por ano até 2022. "
            "Pellegrini: R$ 2,1 mi. Não somadas: são do Hotel/São Crispin."
        ),
        MID,
    ),
    (
        "3",
        "Arrendamento da São Crispin",
        "Família Lorenzet, pago em sacas de soja (contratos 2017–2026). Não somado.",
        MID,
    ),
    (
        "4",
        "Soja no caixa da SEI",
        (
            "A SEI paga Funrural e IRPJ da São Crispin sobre soja e usa sacas para comprar a "
            "escavadeira. Precisa ser explicado."
        ),
        WARN,
    ),
]
for i, (n_, head, body, col) in enumerate(items):
    y = 1.6 + i * 1.3
    badge(s, 0.7, y + 0.1, n_, fill=col)
    text(s, 1.5, y, 11.0, 0.45, head, size=18, color=DARK, bold=True, font=HEAD)
    text(s, 1.5, y + 0.45, 11.0, 0.8, body, size=15)

# ------------------------------------------------------------------ 10. section 11.2 comparison
gap = B["sem_explicacao"]
s = new_slide(
    "Contratos | Indenizações | Compromissos | Empréstimos",
    source="aba GERAL (Siviero = doc 11, Rafael = Estudo 2); proposta v0",
    notes=(
        "Primeiro vêm os blocos em separado; só depois o total do tópico. Nesta versão, só o 11.2.1 muda: "
        "sem multa enquanto os contratos estiverem vigentes, e o investimento sem explicação entra como "
        "ativo a demonstrar. Os demais itens seguem o valor do grupo até a decisão item a item."
    ),
)
rows = [["Item", "Siviero", "Rafael", "Proposta v0"]]
labels112 = {
    "11.2.1": "11.2.1 Loteamentos contratados (multas)",
    "11.2.2": "11.2.2 Lotes de terceiros indenizáveis",
    "11.2.3": "11.2.3 Obras finais a pagar",
    "11.2.4": "11.2.4 Demais compromissos",
    "11.2.5": "11.2.5 Empréstimos e financiamentos",
}
prop_total = 0.0
for c in N["contratos"]:
    p = 0.0 if c["codigo"] == "11.2.1" else c["siviero"]
    prop_total += p
    rows.append(
        [
            labels112[c["codigo"]],
            mi(c["siviero"]),
            mi(c["rafael"]) if c["rafael"] else "incl. 11.2.4",
            mi(p),
        ]
    )
rows.append(["Investimento a demonstrar (sem explicação)", "R$ 0,0 mi", "—", mi(gap)])
prop_total += gap
rows.append(
    [
        "Total do tópico",
        mi(N["contratos_total"]["siviero"]),
        mi(N["contratos_total"]["rafael"]),
        mi(prop_total),
    ]
)
table(
    s,
    0.6,
    1.5,
    [5.0, 2.1, 2.1, 2.1],
    rows,
    size=14,
    row_h=0.5,
    bold_rows=(7,),
    warn_rows=(1, 6),
)
text(
    s,
    0.6,
    5.65,
    12.1,
    1.2,
    [
        [
            ("Multa rescisória: ", {"bold": True}),
            (
                (
                    "não aceita enquanto os contratos estiverem vigentes. Aceita só com distrato assinado e termo de que a empresa "
                    "nunca mais loteará a gleba; nesse caso, os investimentos nela não são cobrados, mas precisam ser demonstrados."
                ),
                {},
            ),
        ],
    ],
    size=14,
)

# ------------------------------------------------------------------ 11. what we ask
s = new_slide(
    "O que pedimos ao grupo",
    source="proposta v0, sujeita à aprovação do Luís",
    notes="Dois caminhos: os valores reais permitem recalcular; sem eles, vale a devolução de 1/6.",
)
asks = [
    (
        "1",
        "Valores reais",
        (
            "Preço e recebimento de cada lote vendido de 2018 a 2024, e o investimento feito em cada "
            "loteamento não finalizado."
        ),
    ),
    (
        "2",
        "Ou a devolução de 1/6",
        (
            f"Sem esses valores: 1/6 do que não tem explicação, {mi(B['um_sexto'])} no preço do "
            "próprio grupo."
        ),
    ),
    (
        "3",
        "Multas só com distrato",
        (
            "Distrato assinado e termo de nunca mais lotear a gleba. Novo contrato depois da "
            "cisão seria tratado como fraude."
        ),
    ),
    (
        "4",
        "Lotes de 2025 e lotes a mais",
        (
            f"Os {num(N['lotes']['apos_2024'])} lotes vendidos em 2025 voltam à data-base; os "
            f"{num(N['excedente_lista']['n'])} lotes a mais do arquivo 17 precisam de explicação."
        ),
    ),
]
for i, (n_, head, body) in enumerate(asks):
    x = 0.6 + (i % 2) * 6.15
    y = 1.6 + (i // 2) * 2.6
    box(s, x, y, 5.9, 2.35, fill=TINT)
    badge(s, x + 0.25, y + 0.25, n_)
    text(
        s, x + 1.0, y + 0.22, 4.7, 0.6, head, size=20, color=DARK, bold=True, font=HEAD
    )
    text(s, x + 1.0, y + 0.85, 4.7, 1.4, body, size=15)

# ------------------------------------------------------------------ 12. limits
s = new_slide(
    "Limites deste rascunho",
    source="aba Pendências",
    notes=(
        "Tudo aqui é provisório. Os números mudam quando os materiais em edição forem fechados e quando o Luís aprovar "
        "as premissas."
    ),
)
text(
    s,
    0.6,
    1.5,
    12.1,
    5.3,
    [
        [("Premissas a aprovar", {"bold": True, "color": DARK})],
        (
            f"Os 15 meses de relatórios representam os {months} meses do período; o preço de referência é a média do grupo "
            "de 2025."
        ),
        [("Materiais ainda em edição", {"bold": True, "color": DARK})],
        (
            "Classificação das contas pagas (dúvidas sem resposta); achados da análise do ACERTOS E ACORDOS; levantamento "
            "ambiental (lotes de terceiros); soja e terras da São Crispin."
        ),
        [("Escolhas a favor do grupo", {"bold": True, "color": DARK})],
        (
            "Concessão de R$ 7,5 mi para concluir loteamentos; bens novos tratados como comprados em dinheiro; dívidas de 2018 "
            "ausentes em 2025 tratadas como pagas. Corrigir isso só aumenta o valor sem explicação."
        ),
        [("Classificação das saídas", {"bold": True, "color": DARK})],
        "Feita por palavras-chave no texto do arquivo 17; a lista lote a lote está na planilha para conferência.",
    ],
    size=16,
    space_after=8,
)

prs.save(OUT)
print("saved", OUT, "slides:", state["n"])
