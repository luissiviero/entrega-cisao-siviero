"""Build the simplified Part 4 deck (python-pptx, native charts) from numeros-v0.2.json.

Every figure comes from that file, which build_fluxo_caixa.py derives from the sources.
Style follows 03_entrega/apresentacao/_build/build_deck.py (v1 deck) so both decks read as one.
The v0 generator is frozen in build_deck_parte4_v0.py.
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
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE / "_scripts"))
from deck_extra_slides import session_figures

OUT = Path(sys.argv[1])
N = json.loads((HERE / "numeros-v0.2.json").read_text(encoding="utf-8"))
REF = N["referencia"]
VERSION = "Rascunho v0.3 · Parte 4 da SEI · 28/10/2026"

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
prs.core_properties.title = (
    f"SEI — para onde foi o dinheiro dos lotes (rascunho {N['versao']})"
)
prs.core_properties.author = "Luís Siviero (rascunho)"
state = {"n": 0}


# ------------------------------------------------------------------ formatting helpers
def mi(v: float, dec: int = 1) -> str:
    s = f"{abs(v) / 1e6:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if v < 0 else "") + f"R$ {s} mi"


def mil(v: float) -> str:
    return "R$ " + f"{v / 1e3:,.0f}".replace(",", ".") + " mil"


def num(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def nb(t: str) -> str:
    """Keep currency, figures and units together on one line."""
    t = t.replace("R$ ", "R$\u00a0")
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


def box(slide, x, y, w, h, fill=TINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def label_shape(s, label, size, color=WHITE, font=HEAD):
    tf = s.text_frame
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = (
        font,
        Pt(size),
        True,
        color,
    )


def badge(slide, x, y, label, d=0.55, fill=DARK, size=18):
    label_shape(box(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL), label, size)


def tag(slide, x, y, label="PROVISÓRIO", w=1.6):
    s = box(slide, x, y, w, 0.36, fill=WARN)
    s.adjustments[0] = 0.5
    label_shape(s, label, 11, font=BODY)


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
    text(
        s,
        0.6,
        7.0,
        11.4,
        0.35,
        VERSION + (" · Fonte: " + source if source else ""),
        size=10,
        color=MUTED,
    )
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
                else WARN_TINT
                if i in warn_rows
                else TINT
                if i in bold_rows
                else WHITE
            )
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = nb(str(v))
            r.font.name, r.font.size = BODY, Pt(size)
            r.font.bold = i == 0 or i in bold_rows
            r.font.color.rgb = WHITE if i == 0 else TEXT


def style_chart(chart):
    chart.font.name = BODY
    chart.font.size = Pt(12)
    chart.font.color.rgb = MUTED
    chart.has_legend = False
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(11)
    va.tick_labels.number_format_is_linked = False
    chart.category_axis.format.line.color.rgb = GRID
    chart.category_axis.tick_labels.font.size = Pt(11)


def labels_on(plot, fmt, pos):
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format, dl.number_format_is_linked, dl.position = fmt, False, pos
    dl.font.size = Pt(11)


B = N["ponte"][REF]
BP, BC = N["ponte"]["piso"], N["ponte"]["central"]
months = N["meses"]
lots_in = N["lotes"]["no_periodo"]
cls_n = N["lotes"]["por_classe"]
cost = N["custos"]["classes"]
rt = N["teste_receita"]
sens = {(s["inicio"], s["concessao"]): s for s in N["sensibilidade"]}
assert abs(sum(cls_n.values()) - lots_in) < 1e-6

# ------------------------------------------------------------------ 1. cover
s = new_slide(
    "",
    dark=True,
    notes=(
        "Rascunho para discussão. O estudo segue o dinheiro: quanto entrou com os lotes que saíram de 2019 a 2024, quanto "
        "o grupo gastou para manter a empresa e pagar dívidas, e quanto sobra sem explicação."
    ),
)
text(
    s,
    0.9,
    1.7,
    11.5,
    1.2,
    "Para onde foi o dinheiro dos lotes vendidos?",
    size=36,
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
    "SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS · do inventário de 2018 a 31/12/2024",
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
    f"Rascunho {N['versao']} para discussão · 28/10/2026 · valores provisórios",
    size=14,
    color=TINT,
)

# ------------------------------------------------------------------ 2. headline
s = new_slide(
    "O patrimônio caiu e o dinheiro não aparece",
    source=f"fluxo-caixa-sei-{N['versao']}.xlsx, aba Ponte",
    notes=(
        f"Preço de referência: média do próprio grupo de 2025, trazida pelo IPCA à data da venda e sustentada pelos impostos e comissões "
        f"pagos (slide 8). No preço de 2018 do inventário a conta fecha em {mi(BP['sem_explicacao'])}; "
        f"no intermediário, {mi(BC['sem_explicacao'])}."
    ),
)
cards = [
    (
        num(lots_in),
        "lotes saíram do estoque de 2019 a 2024",
        f"mais {num(N['lotes']['apos_2024'])} em 2025, fora da conta",
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
    box(s, x, 1.65, 2.85, 3.0, fill=WARN_TINT if i == 3 else TINT)
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
            "Ou o grupo mostra os valores reais e a diferença é recalculada, ou 1/6 do valor sem explicação volta ao sócio que "
            "sai."
        ),
    ],
    size=16,
)

# ------------------------------------------------------------------ 3. the contradiction
inv18 = N["investimentos_2018"]
c112 = {c["codigo"]: c for c in N["contratos"]}
s = new_slide(
    "“R$ 0 + multa” contradiz o próprio grupo em 2018",
    source="inventário 2018 (doc 18); doc 11; Estudo 2 de Rafael; aba GERAL",
    notes=(
        "Em 2018 o grupo atribuía valor de investimento aos loteamentos em execução e contratados. Em 2025 "
        "os mesmos empreendimentos valem zero e geram multa. Rafael soma lotes executados com multas, o "
        "que é contraditório: se há multa por distrato, não há lote; se há lote, não há multa."
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
        ]
    ],
    size=16,
)

# ------------------------------------------------------------------ 4. stock over time
path = N["trajetoria"]
first, end24, last = path[0], next(p for p in path if p["ano"] == "fim 2024"), path[-1]
drop = 1 - end24["valor25"] / first["valor25"]
s = new_slide(
    f"O estoque de lotes perdeu {num(drop * 100)}% do valor até 2024",
    source="aba 'Lotes que saíram 2018-2025' e arquivo 17; preços de out/2025 do grupo",
    notes=(
        "Estoque de lotes executados, todos no mesmo preço (média do grupo em out/2025), para que a queda "
        "mostre só os lotes que saíram e não a variação de preço. Saídas sem data e vendas antigas que "
        "ainda estavam na contagem de 2018 foram distribuídas igualmente entre 2018 e 2024."
    ),
)
cd = CategoryChartData()
cd.categories = [p["ano"].replace("fim ", "") for p in path]
cd.add_series(
    "Estoque de lotes (R$ mi, preços de out/2025)",
    [round(p["valor25"] / 1e6, 1) for p in path],
)
ch = s.shapes.add_chart(
    XL_CHART_TYPE.LINE_MARKERS, Inches(0.6), Inches(1.5), Inches(8.3), Inches(5.3), cd
).chart
style_chart(ch)
ser = ch.series[0]
ser.format.line.color.rgb = DARK
ser.format.line.width = Pt(3)
ser.marker.format.fill.solid()
ser.marker.format.fill.fore_color.rgb = DARK
ser.smooth = False
labels_on(ch.plots[0], "#,##0.0", XL_LABEL_POSITION.ABOVE)
ch.value_axis.minimum_scale = 0
ch.value_axis.tick_labels.number_format = "#,##0"
box(s, 9.2, 1.6, 3.5, 5.1, fill=TINT)
text(
    s,
    9.45,
    1.8,
    3.1,
    4.8,
    [
        [(first["ano"].replace("fim", "Fim de"), {"bold": True, "color": DARK})],
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
    f"Como saíram os {num(lots_in)} lotes de 2019 a 2024",
    source="arquivo 17 (forma de cada saída), classificado por palavras-chave",
    notes=(
        "Lotes por forma de saída. Só a venda gera dinheiro; vendas com parte em veículo ou imóvel contam "
        "como dinheiro porque esses bens não aparecem no inventário de 2025. As demais formas têm destino "
        "declarado. Lotes sem informação não têm destino documentado (premissa P4)."
    ),
)
labels = {
    "VENDA": "Venda (dinheiro)",
    "ACORDO": "Acordos, indenizações e distratos",
    "OBRA": "Pagamento de obras",
    "SOC": "Sócios e família",
    "BEMX": "Troca por bem que não está em 2025",
    "DIVL": "Empréstimo de sócio pago com lote",
    "BEM": "Troca por outro lote do estoque",
    "SEM": "Sem destino documentado",
}
order = sorted(labels, key=lambda k: -cls_n.get(k, 0))
cd = CategoryChartData()
cd.categories = [labels[k] for k in order]
cd.add_series("Lotes", [round(cls_n.get(k, 0), 1) for k in order])
ch = s.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(1.5), Inches(8.0), Inches(5.3), cd
).chart
style_chart(ch)
ch.category_axis.reverse_order = True
ch.value_axis.tick_labels.number_format = "0"
ch.plots[0].gap_width = 60
labels_on(ch.plots[0], "0", XL_LABEL_POSITION.OUTSIDE_END)
for i, k in enumerate(order):
    pt = ch.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = (
        DARK if k == "VENDA" else WARN if k in ("SEM", "BEMX") else MID
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
        f"Acordos e distratos: {mi(B['ACORDO'])}",
        f"Obras e serviços: {mi(B['OBRA'])}",
        f"Sócios e família: {mi(B['SOC'])}",
        f"Empréstimos de sócios: {mi(B['DIVL'])}",
        f"Troca por outro lote: {mi(B['BEM'])}",
        [
            ("Sem destino ou bem sumido: ", {"bold": True}),
            (mi(B["SEM"] + B["BEMX"]), {"bold": True, "color": WARN}),
        ],
    ],
    size=15,
    space_after=7,
)

# ------------------------------------------------------------------ 6. cost of keeping the company
s = new_slide(
    "Manter a empresa custa " + mil(N["custos"]["m_rec"] / 15) + " por mês",
    source="relatórios mensais de contas pagas, jul/2023 a out/2024 (15 meses), só a SEI; IPCA (IBGE)",
    notes=(
        "Método decidido pelo Luís: a média mensal dos gastos recorrentes da SEI nos 15 meses (preços de mar/2024) é "
        "levada a cada mês de jan/2019 a dez/2024 pelo IPCA. Recorrentes incluem folha, pró-labore, impostos, juros, "
        "comissões, divisão de aluguéis, obras em loteamentos finalizados e despesas de familiares e de outras "
        "empresas. Gastos extraordinários (multas, acordos, lucros) entram uma vez, como pagos. Combustível não é "
        "recorrente: indica máquinas em obra."
    ),
)
cats = [
    ("Manter a empresa", cost["M"]),
    ("Obras, projetos e combustível", cost["INV"]),
    ("Compra de bens", cost["BENS"]),
    ("Pago com lotes", cost["LOTE"]),
    ("Dívidas e acertos", cost["DIV"] + cost.get("DOC", 0)),
]
cd = CategoryChartData()
cd.categories = [c for c, _ in cats]
cd.add_series("R$ mi em 15 meses", [round(v / 1e6, 2) for _, v in cats])
ch = s.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.6),
    Inches(1.5),
    Inches(7.6),
    Inches(5.3),
    cd,
).chart
style_chart(ch)
ch.value_axis.tick_labels.number_format = "0.0"
ch.category_axis.tick_labels.font.size = Pt(10)
ch.plots[0].gap_width = 70
labels_on(ch.plots[0], "#,##0.00", XL_LABEL_POSITION.OUTSIDE_END)
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
        [("Recorrentes", {"bold": True, "color": DARK})],
        f"{mi(N['custos']['m_rec'], 2)} ÷ 15 = {mil(N['custos']['m_rec'] / 15)}/mês",
        [("Estendido a 2019–2024 com IPCA", {"bold": True, "color": DARK})],
        f"{months} meses = {mi(B['manter_rec'])}, mais {mi(B['manter_ext'])} de multas, acordos e lucros",
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
    source=f"fluxo-caixa-sei-{N['versao']}.xlsx, aba Ponte (preço = média do grupo de 2025 trazida pelo IPCA)",
    notes=(
        "Entradas em verde, saídas em terracota. Dívidas: saldo de jan/2019 da SEI menos a dívida em "
        "31/12/2024, mais o acerto com a Flávia Manica. Conclusão de loteamentos: só o que os números do "
        "grupo documentam (Falavigna: R$ 2,45 mi a concluir em 2018, R$ 2,0 mi ainda a pagar em 2025)."
    ),
)
steps = [
    ("Lotes vendidos em dinheiro", B["VENDA"]),
    ("Aluguéis", B["alugueis"]),
    ("Lotes e bens sem destino", B["SEM"] + B["BEMX"] + B["bens_2018_sem_destino"]),
    ("Manter: gastos recorrentes", -B["manter_rec"]),
    ("Manter: multas, acordos, lucros", -B["manter_ext"]),
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


def signed(v):
    s = f"{abs(v) / 1e6:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("+" if v >= 0 else "−") + s


# values go in the category labels: bars under ~R$ 1 mi are too thin for a label inside them
cd.categories = [f"{c}\n{signed(v)}" for c, v in steps] + [
    f"Sem explicação\n{signed(run)}"
]
cd.add_series("base", [round(v / 1e6, 1) for v in base])
cd.add_series("valor", [round(v / 1e6, 1) for v in bar])
ch = s.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED,
    Inches(0.6),
    Inches(1.45),
    Inches(12.1),
    Inches(4.4),
    cd,
).chart
style_chart(ch)
ch.value_axis.tick_labels.number_format = "0"
ch.category_axis.tick_labels.font.size = Pt(10)
ch.plots[0].gap_width = 40
ch.series[0].format.fill.background()
ch.series[0].format.line.fill.background()
vs = ch.series[1]
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
                    f" sem explicação · 1/6 = {mi(B['um_sexto'])} · somando os lotes dados em obras e em loteamentos contratados, "
                    f"investimento a demonstrar de {mi(B['investimento_a_demonstrar'])}"
                ),
                {},
            ),
        ]
    ],
    size=15,
)

# ------------------------------------------------------------------ 8. which price
s = new_slide(
    "Os impostos pagos confirmam o preço do grupo",
    source="aba Teste de receita (PIS, COFINS e comissões dos 15 meses) e aba Sensibilidade",
    notes=(
        "O grupo não informou por quanto vendeu cada lote. Mas, nos 15 meses de relatórios, a SEI pagou PIS e "
        "COFINS sobre uma receita de cerca de R$ 11 mi e comissões de corretagem sobre cerca de R$ 6,8 mi "
        "de vendas. Nos mesmos meses, o preço do grupo dá vendas desse tamanho; o preço de 2018 dá metade. "
        "A base de PIS e COFINS também inclui aluguéis, serviços e parcelas de vendas antigas."
    ),
)
mod = rt["modelo"]
cd = CategoryChartData()
cd.categories = [
    "Receita pelo PIS/COFINS",
    "Vendas pelas comissões",
    "Modelo: preço de 2018",
    "Modelo: intermediário",
    "Modelo: preço do grupo",
]
cd.add_series(
    "R$ mi",
    [
        round(v / 1e6, 1)
        for v in (
            rt["receita_cofins"],
            rt["vendas_comissao"],
            mod["piso"],
            mod["central"],
            mod["teto"],
        )
    ],
)
ch = s.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(1.5), Inches(7.0), Inches(3.4), cd
).chart
style_chart(ch)
ch.category_axis.reverse_order = True
ch.value_axis.tick_labels.number_format = "0"
ch.plots[0].gap_width = 50
labels_on(ch.plots[0], "#,##0.0", XL_LABEL_POSITION.OUTSIDE_END)
for i in range(5):
    pt = ch.series[0].points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = WARN if i < 2 else (DARK if i == 4 else MID)
rows = [["Preço dos lotes", "Sem explicação", "1/6"]]
for k, lab in (
    ("piso", "Valor de 2018"),
    ("central", "Intermediário"),
    ("teto", "Média do grupo 2025"),
):
    v = N["ponte"][k]["sem_explicacao"]
    rows.append([lab, mi(v), mi(v / 6) if v > 0 else "—"])
table(s, 7.9, 1.6, [2.15, 1.6, 1.1], rows, size=13, row_h=0.5, bold_rows=(3,))
s18 = sens[(2018, "documentada")]
smax = sens[(2019, "máxima (v0)")]
nominal = N["ponte"]["teto_nominal"]["sem_explicacao"]
ex = N["excedente_lista"]
text(
    s,
    0.6,
    5.05,
    12.1,
    1.9,
    [
        [
            ("Jun/2023 a set/2024, R$ milhões. ", {"color": MUTED}),
            ("O que muda o número, no preço do grupo:", {"bold": True, "color": DARK}),
        ],
        (
            f"Preço de 2025 sem descontar a inflação: {mi(nominal)}. Contar também 2018: {mi(s18['teto'])}. Aceitar a "
            f"concessão máxima para concluir Falavigna e Colina do Retiro II: {mi(smax['teto'])}."
        ),
        (
            f"A conta já inclui os {num(ex['n'])} lotes que o arquivo 17 lista além da redução do inventário (o próprio "
            "arquivo 17 declara essas quantidades em 2018; decisão de 28/10)."
        ),
    ],
    size=14,
    space_after=6,
)

# ------------------------------------------------------------------ 9. other income shown separately
s = new_slide(
    "Outras entradas, mostradas à parte",
    source="inventário 2018 (aluguéis); docx do Hotel; relatórios mensais; contratos Lorenzet",
    notes=(
        "Só os aluguéis da SEI entram na conta, pela relação de aluguéis de 2018. As vendas de terras e a soja "
        "são de outras empresas do grupo (Hotel Princesa dos Vales / São Crispin) e não foram somadas. Os "
        "relatórios mensais mostram que esse dinheiro passa pelo caixa da SEI, o que precisa ser "
        "esclarecido."
    ),
)
items = [
    (
        "1",
        "Aluguéis da SEI",
        (
            f"{mil(N['alugueis_mes'])}/mês na relação de 2018 = {mi(B['alugueis'])} em {months} meses. "
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
        "Arrendamento da São Crispin e serviços",
        (
            "Família Lorenzet, pago em sacas de soja (2017–2026). A SEI também "
            "emite notas de serviço (CERAN, CEVAL). Não somados."
        ),
        MID,
    ),
    (
        "4",
        "Soja e contas a receber",
        (
            "A SEI paga Funrural e IRPJ da São Crispin sobre soja e usa sacas para comprar a "
            "escavadeira. Vendas a prazo deixam saldo a receber que nenhum inventário mostra."
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
    source="aba GERAL (Siviero = doc 11, Rafael = Estudo 2); proposta desta versão",
    notes=(
        "Primeiro vêm os blocos em separado; só depois o total do tópico. O 11.2.1 vai a zero: sem multa enquanto "
        "os contratos estiverem vigentes, e o valor sem explicação entra como investimento a demonstrar. Do 11.2.2, "
        "11.2.3 e 11.2.4 sai o que já foi pago com lotes que saíram do inventário (cobrança dupla): 4 lotes da SEI "
        "contados como de terceiros, a pavimentação do Colina do Retiro II, as obras do Falavigna, a comissão de "
        "Susin e o lote da Solar Serra. Regra: só cobranças futuras ficam no 11.2."
    ),
)
rows = [["Item", "Siviero", "Rafael", "Proposta"]]
labels112 = {
    "11.2.1": "11.2.1 Loteamentos contratados (multas)",
    "11.2.2": "11.2.2 Lotes de terceiros indenizáveis",
    "11.2.3": "11.2.3 Obras finais a pagar",
    "11.2.4": "11.2.4 Demais compromissos",
    "11.2.5": "11.2.5 Empréstimos e financiamentos",
}
prop_total = gap
for c in N["contratos"]:
    p = c["proposta"]
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
        [
            ("Cobrança dupla: ", {"bold": True}),
            (
                (
                    f"{mi(sum(a['valor'] for c in N['contratos'] for a in c.get('ajustes', [])))} de dívidas já pagas com "
                    "lotes que saíram do inventário foram tirados do 11.2.2, 11.2.3 e 11.2.4: no 11.2 ficam só cobranças futuras."
                ),
                {},
            ),
        ],
    ],
    size=14,
)

# ------------------------------------------------------------------ 10b. why it is a double charge
cross = openpyxl.load_workbook(
    HERE / "cruzamento-compromissos-arq17.xlsx", data_only=True
)["Avaliação"]
case = {
    r[0]: r
    for r in cross.iter_rows(min_row=2, values_only=True)
    if r[0] and str(r[0]).startswith("D")
}
fal_lots, fal_debt = case["D3"][6], -case["D3"][2]
adj = [(c["codigo"], a) for c in N["contratos"] for a in c.get("ajustes", [])]
adj_total = sum(a["valor"] for _, a in adj)
s = new_slide(
    "Por que é cobrança em duplicidade",
    source="arquivo 17 (destino dos lotes) × doc 11 (passivos); cruzamento-compromissos-arq17.xlsx",
    notes=(
        "Quando um lote paga uma dívida, a dívida acaba. Se o lote sai do inventário e a dívida continua na lista de "
        "compromissos, o patrimônio é reduzido duas vezes pelo mesmo pagamento: uma pela saída do lote e outra pelo "
        "passivo. Exemplo: os lotes do Falavigna entregues à Florense e à Bett em 2025."
    ),
)
steps3 = [
    (
        "1",
        "O lote paga a dívida",
        (
            "Em 2025 a SEI entregou 5 lotes do Falavigna à Florense e à Bett para pagar a galeria "
            "e a pavimentação do loteamento."
        ),
    ),
    (
        "2",
        "O lote sai do inventário",
        f"O patrimônio cai pelo valor dos lotes: {mi(fal_lots)} pelo preço do próprio grupo.",
    ),
    (
        "3",
        "A dívida continua na lista",
        f"O doc 11 ainda desconta {mi(fal_debt)} para concluir as mesmas obras.",
    ),
]
for i, (n_, head, body) in enumerate(steps3):
    x = 0.6 + i * 4.1
    box(s, x, 1.55, 3.85, 2.35, fill=TINT if i < 2 else WARN_TINT)
    badge(s, x + 0.25, 1.75, n_, fill=DARK if i < 2 else WARN)
    text(
        s,
        x + 1.0,
        1.78,
        2.75,
        0.55,
        head,
        size=17,
        color=DARK if i < 2 else WARN,
        bold=True,
        font=HEAD,
    )
    text(s, x + 0.3, 2.5, 3.3, 1.35, body, size=14)
text(
    s,
    0.6,
    4.05,
    6.2,
    1.1,
    [
        [
            (
                "O mesmo pagamento é descontado duas vezes: ",
                {"bold": True, "color": WARN},
            ),
            (
                "uma no ativo, quando o lote sai, e outra no passivo, quando a dívida continua.",
                {},
            ),
        ],
        [
            ("Regra: ", {"bold": True}),
            (
                "no 11.2 só ficam cobranças futuras. O que já foi pago com lote ou dinheiro sai.",
                {},
            ),
        ],
    ],
    size=15,
    space_after=8,
)
labels_adj = {
    "D1": "4 lotes da SEI contados como de terceiros",
    "D2": "Solar Serra (lote já entregue)",
    "D3": "Obras do Falavigna",
    "D4": "Pavimentação do Colina do Retiro II",
    "D5": "Comissão Buriti (Susin)",
}
rows = [["Cobrança passada retirada", "Valor"]]
rows += [
    [f"{code} · {labels_adj.get(a['caso'], a['motivo'])}", mi(a["valor"], 2)]
    for code, a in adj
]
rows.append(["Total", mi(adj_total, 2)])
table(s, 7.1, 4.05, [4.3, 1.35], rows, size=12, row_h=0.36, bold_rows=(len(rows) - 1,))
text(
    s,
    0.6,
    5.35,
    6.2,
    0.9,
    f"Efeito para quem sai: 1/6 × {mi(adj_total)} = {mi(adj_total / 6)}.",
    size=15,
    color=DARK,
    bold=True,
)

# ------------------------------------------------------------------ 11. what we ask
s = new_slide(
    "O que pedimos ao grupo",
    source="proposta desta versão, sujeita à aprovação do Luís",
    notes="Dois caminhos: os valores reais permitem recalcular; sem eles, vale a devolução de 1/6.",
)
asks = [
    (
        "1",
        "Valores reais",
        (
            "Preço e recebimento de cada lote vendido de 2019 a 2024, e o investimento feito em cada "
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
            f"Os {num(N['lotes']['apos_2024'])} lotes que saíram em 2025 voltam à "
            f"data-base; os {num(ex['n'])} lotes a mais do arquivo 17 já contam como "
            "vendidos pela SEI."
        ),
    ),
]
for i, (n_, head, body) in enumerate(asks):
    x, y = 0.6 + (i % 2) * 6.15, 1.6 + (i // 2) * 2.6
    box(s, x, y, 5.9, 2.35, fill=TINT)
    badge(s, x + 0.25, y + 0.25, n_)
    text(
        s, x + 1.0, y + 0.22, 4.7, 0.6, head, size=20, color=DARK, bold=True, font=HEAD
    )
    text(s, x + 1.0, y + 0.85, 4.7, 1.4, body, size=15)

# ------------------------------------------------------------------ 12-16. summaries per study session
for title, source, g, r_, prop, prop_label, comments in session_figures():
    s = new_slide(
        title,
        source=source,
        notes=(
            "Resumo da sessão que estuda esta empresa: valor do grupo Siviero (inventário 2025), valor de Rafael (Estudo 2) "
            "e a proposta com as decisões aprovadas até 28/10/2026. Valores em andamento."
        ),
    )
    cards3 = [
        ("Grupo Siviero", g, TINT, DARK),
        ("Rafael", r_, TINT, MID),
        (prop_label, prop, WARN_TINT, WARN),
    ]
    for i, (lab, v, fill, col) in enumerate(cards3):
        x = 0.6 + i * 4.1
        box(s, x, 1.6, 3.85, 2.3, fill=fill)
        text(s, x + 0.25, 1.8, 3.4, 0.5, lab, size=16, color=col, bold=True)
        text(
            s, x + 0.25, 2.35, 3.4, 0.9, mi(v), size=32, color=col, bold=True, font=HEAD
        )
        if i:
            d = v - g
            text(
                s,
                x + 0.25,
                3.25,
                3.4,
                0.5,
                ("+" if d >= 0 else "") + mi(d) + " em relação ao grupo",
                size=13,
                color=MUTED,
            )
    text(
        s,
        0.6,
        4.25,
        12.1,
        2.5,
        [[("• ", {"bold": True, "color": DARK}), (c, {})] for c in comments],
        size=15,
        space_after=8,
    )

# ------------------------------------------------------------------ 12. limits
s = new_slide(
    "Limites deste rascunho",
    source="aba Pendências",
    notes=(
        "Tudo aqui é provisório. Os números mudam quando os materiais em edição forem fechados e quando o Luís aprovar as "
        "premissas."
    ),
)
text(
    s,
    0.6,
    1.5,
    12.1,
    5.3,
    [
        [("Premissas", {"bold": True, "color": DARK})],
        (
            f"Aprovadas: custos de 2019 a 2024 ({months} meses) extrapolados com IPCA; preço de referência = média do "
            "grupo de 2025 trazida pelo IPCA; lotes a mais do arquivo 17 incluídos; pró-labore de 4 dos 6 ramos (Válter, "
            "Adélia, Pedro e Raul) como custo da empresa. A decidir: itens 11.2.2 a 11.2.5, um a um."
        ),
        [("Materiais ainda em edição", {"bold": True, "color": DARK})],
        (
            "Achados da análise do ACERTOS E ACORDOS; levantamento ambiental (lotes de terceiros); soja e terras da São "
            "Crispin. A classificação das contas pagas usada é a versão final de 28/10."
        ),
        [("Escolhas a favor do grupo", {"bold": True, "color": DARK})],
        (
            "Bens novos tratados como comprados em dinheiro; dívidas de 2018 que não aparecem em 2025 tratadas como pagas; "
            "conta do Sicredi (R$ 5,4 mi) tratada como da SEI."
        ),
        [("Conferência", {"bold": True, "color": DARK})],
        (
            "Revisor independente em três rodadas (v0, v0.1 e v0.2); os achados foram corrigidos nesta versão. Classificação "
            "das saídas por palavras-chave; a lista lote a lote está na planilha."
        ),
    ],
    size=16,
    space_after=8,
)

prs.save(OUT)
print("saved", OUT, "slides:", state["n"])
