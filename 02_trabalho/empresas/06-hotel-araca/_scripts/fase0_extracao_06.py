"""Phase 0 extraction for company 06 (Hotel Araçá).

Reads the 2018, 2025 (current and old) inventories, Rafael's PDF (group
transcription, valuation ranges and the "balanço executado"), Rafael's
selection list, the Google-sheet tabs and the 2018/2025 summary documents,
parses every item, checks each total against the total written in the
document and writes _extracoes/fase0-06.xlsx.

Every number comes from the sources. Each check needs both sides parsed
(assert). Known, documented divergences are listed in EXPECTED_DIVERGE; any
other divergence, or an expected one that disappears, stops the script.

Usage: python fase0_extracao_06.py
"""

from __future__ import annotations

import hashlib
import re
from itertools import pairwise
from pathlib import Path

import docx
import openpyxl
import pymupdf
from docx.table import Table

BASE = Path(__file__).resolve().parents[4]
OUT_DIR = BASE / "02_trabalho" / "empresas" / "06-hotel-araca" / "_extracoes"
DRIVE = Path("G:/Meu Drive/02_Patrimonio_e_Sociedades/Cisao_Siviero/Material Siviero")

SRC = {
    "inv2025": BASE
    / "01_fontes/inventario-2025/06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx",
    "inv2025_antigo": BASE
    / "01_fontes/inventario-2025-antigo/ANTIGO 06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx",
    "inv2018": DRIVE
    / "2018/inventário-2018/16. 1. HOTEL ARAÇÁ LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.docx",
    "resumo2018": DRIVE
    / "2018/inventário-2018/27. 1. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - RESUMO   -   Dezembro de 2018.docx",
    "doc00_2025": BASE
    / "01_fontes/inventario-2025/00. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - 31.12.2024.docx",
    "rafael_pdf": BASE / "01_fontes/estudo-rafael/HOTEL ARAÇÁ LTDA - BELISE LTDA.pdf",
    "diretrizes_pdf": BASE
    / "01_fontes/estudo-rafael/DIRETRIZES PARA SELEÇÃO DE ATIVOS NA SEPARAÇÃO PATRIMONIAL E SOCIETÁRIA , BENTA E V SIVIERO BELISE LTDA.pdf",
    "lista_rafael": BASE / "01_fontes/outros/ativos-avaliados-valter.xlsx",
    "gsheet_2_0": BASE
    / "01_fontes/planilhas-google/2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx",
    "gsheet_rafael": BASE
    / "01_fontes/planilhas-google/Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx",
    "observacoes": BASE / "02_trabalho/observacoes-planilhas-google.xlsx",
    "conferencia_md": BASE / "01_fontes/outros/conferencia-discrepancias.md",
    "conferencia_xlsx": BASE
    / "01_fontes/outros/inventario-2025-avaliacoes-corrigido.xlsx",
    "pendencias_md": BASE / "02_trabalho/pendencias-conferencia-anterior.md",
    "diff_md": BASE / "02_trabalho/diff-inventario-2025-antigo-x-atual.md",
    "resp_arq1": BASE
    / "01_fontes/respostas-docs-2026-08-13/1 - grupo siviero - encaminhamento esclarecimentos ao Dr. Pablo - 13.08.2026.pdf",
    "resp_arq12": BASE
    / "01_fontes/respostas-docs-2026-08-13/12 - OUTORGA_POCO_HOTEL_ARACA.pdf",
    "resp_arq17": BASE
    / "01_fontes/respostas-docs-2026-08-13/17 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - LOTES VENDIDOS DE JANEIRO DE 2018 A OUTUBRO DE 2025 - REV.13.08.2026.xlsx",
}

MONEY = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
PARTNERS = [
    "VÁLTER SIVIERO",
    "IRINEU SIVIERO",
    "RAUL SIVIERO",
    "ELCIO SIVIERO",
    "ADÉLIA F. SIVIERO",
    "PEDRO SIVIERO",
]
RAFAEL_ITEM20_INSCRICAO = "645.0027.015.000"

# Divergences found and documented in the journal (2026-10-28). A new one, or
# one of these turning into a match, stops the script so it is looked at.
EXPECTED_DIVERGE = {
    "Rafael – balanço executado (p. 16–20): item 01 + 21 lotes",
    "Rafael executado = grupo, item 01",
    "Rafael executado = grupo, item 20",
    "GERAL C70 (2018) × total do inventário 2018",
    "GERAL C80 (lotes 2018) × lotes do inventário 2018",
    "GERAL D70 (2018 atualizado) × 2018 com o lote 28/q38 em 350.000",
    "GERAL G70 (diferença) × lotes 2025 sem par com valor em 2018",
    "Aba 6.Hotel Araçá (2_0) C2 × total do inventário 2018",
    "Aba 6.Hotel Araçá (2_0) H2 × lotes 2025 sem par com valor em 2018",
}


def md5_10(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()[:10]


def brl(s: str) -> float:
    return float(s.replace(".", "").replace(",", "."))


def first(pattern: str, text: str):
    m = re.search(pattern, text)
    return m.group(1) if m else None


def norm_inscricao(s):
    """'360.0037.015.000.000' and '360.0037.015.000' are the same parcel."""
    return ".".join(s.split(".")[:4]) if s else None


# ---------------------------------------------------------------- inventories


def doc_blocks(path: Path):
    d = docx.Document(str(path))
    for child in d.element.body.iterchildren():
        if child.tag.endswith("}tbl"):
            yield "tbl", Table(child, d)
        elif child.tag.endswith("}p"):
            yield (
                "p",
                "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t")),
            )


def table_rows(path: Path):
    """Yield (table_no, cells) for every row of every table, merged cells once."""
    tno = 0
    for kind, b in doc_blocks(path):
        if kind != "tbl":
            continue
        tno += 1
        for r in b.rows:
            cells, prev = [], None
            for c in r.cells:
                if c._tc is prev:
                    continue
                prev = c._tc
                cells.append(c.text.strip())
            yield tno, cells


def observation_paragraphs(path: Path):
    """Paragraphs from 'OBSERVAÇÕES' up to the place/date line."""
    out, on = [], False
    for kind, b in doc_blocks(path):
        if kind != "p":
            continue
        t = b.strip()
        if t.startswith("OBSERVAÇÕES"):
            on = True
        if on and re.match(r"Veranópolis – RS, \d", t):
            break
        if on and t and not set(t) <= set(". "):
            out.append(t)
    return out


def parse_inventory(path: Path, label: str):
    items, total = [], None
    for tno, cells in table_rows(path):
        if not cells:
            continue
        if cells[0].startswith("TOTAL GERAL"):
            total = brl(first(MONEY, cells[1]))
            continue
        if not re.fullmatch(r"\d{2}", cells[0]):
            continue
        desc, val_txt = cells[1], cells[-1]
        m = re.search(MONEY, val_txt)
        partner = next((p for p in PARTNERS if p in val_txt.upper()), None)
        insc = first(r"Inscrição Imobiliária nº ([\d\.]+\d)", desc)
        cod = first(r"Código de Imóvel nº (\d+)", desc)
        items.append(
            {
                "fonte": label,
                "tabela": tno,
                "item": cells[0],
                "lote": first(
                    r"lote (?:urbano )?(?:denominado de lote )?nº (\d+)", desc
                ),
                "quadra": first(r"[Qq]uadra (?:nº )?(\d+)", desc),
                "area_m2": first(
                    r"área (?:total )?de (\d{1,3}(?:\.\d{3})*,\d{2})m²", desc
                ),
                "matricula": first(r"Matrícula nº ([\d\.]+\d)", desc)
                if "Matrículas nº" not in desc
                else "12 matrículas",
                "matricula_antiga": first(r"antiga matrícula ([\d\.]+\d)", desc),
                "comarca": first(r"Comarca de ([^–]+?) –", desc),
                "inscricao": insc,
                "inscricao_norm": norm_inscricao(insc),
                "codigo_imovel": cod.zfill(6) if cod else None,
                "valor": brl(m.group(1)) if m else None,
                "valor_texto": val_txt.replace("\n", " / "),
                "nome_socio_no_valor": partner,
                "descricao": desc.replace("\n", " / "),
            }
        )
    return items, total


COMPONENTS = [
    ("1.1 Terreno", r"^Valor do terreno onde está construído"),
    ("1.2 Móveis 36 aptos novos", r"^Móveis e Equipamentos – novos"),
    ("1.3 Móveis 108 aptos", r"^Móveis e Equipamentos, instalados nos 108"),
    ("1.4 Obra civil", r"^Valor do Edifício Construído – obra civil = 04 andares, com"),
    ("1.5 Equipamentos diversos", r"^Móveis e Equipamentos Diversos"),
    (
        "1.6 Vantagem da coisa feita",
        r"^Valor do Edifício Construído – obra civil = 04 andares = \d",
    ),
]


def item01_components(desc: str, label: str):
    paras = [p.strip() for p in desc.split(" / ")]
    unit_m2 = first(r"valor comercial de R\$ ([\d\.,]+\d), ao metro", desc)
    rows = []
    for name, pat in COMPONENTS:
        hits = [p for p in paras if re.search(pat, p)]
        assert len(hits) == 1, (label, name, len(hits))
        p = hits[0]
        amounts = re.findall(r"R\$ ([\d\.,]+\d)", p)
        area = first(r"(\d{1,3}(?:\.\d{3})*,\d{2})m² de (?:terreno|área construída)", p)
        unit = (
            unit_m2
            if name.startswith("1.1")
            else first(r"X R\$ ([\d\.,]+\d)", p)
            or first(r"valor unitário de R\$ ([\d\.,]+\d)", p)
        )
        rows.append(
            {
                "fonte": label,
                "componente": name,
                "valor": brl(amounts[-1]),
                "area_m2": area,
                "valor_unitario": unit,
                "base_cub": first(r"(CUB/[A-ZÇ]+ DE \d{4})", p),
                "depreciacao_pct": first(r"depreciação de (\d+,\d{2})%", p),
                "percentual_vantagem": first(r"= (\d+,\d{2})% \(", p)
                if name.startswith("1.6")
                else None,
                "texto": p,
            }
        )
    return rows


# ---------------------------------------------------------------- Rafael PDF


def pdf_pages(path: Path):
    return [p.get_text() for p in pymupdf.open(str(path))]


def balance_section(pages, first_page: int, last_page: int):
    """Item 01 value and the 21 lots of one balance section, anchored per lot.

    The section text is split at each 'Um lote urbano' (start of a lot
    description). Each block must hold exactly one matrícula, one inscrição
    and one lot value; the 'valor total global de R$ 750.000,00' quoted inside
    item 20 is not a lot value. Item 01's value is the single amount between
    the 12-matrícula list and the first lot.
    """
    marks = []
    parts = []
    for i in range(first_page, last_page + 1):
        marks.append((sum(len(x) for x in parts), i))
        parts.append(pages[i - 1])
    text = "".join(parts)

    def page_of(pos):
        return [p for start, p in marks if start <= pos][-1]

    head = text.index("Dezoito (18) terrenos")
    total_pos = text.index("TOTAL GERAL DA", head)
    lot_starts = [
        m.start()
        for m in re.finditer(r"Um lote urbano", text[:total_pos])
        if m.start() > head
    ]
    assert len(lot_starts) == 21, len(lot_starts)
    it01 = list(re.finditer(r"R\$\s*" + MONEY, text[head : lot_starts[0]]))
    it01 = [
        m for m in it01 if not text[head + m.start() - 30 : head + m.start()].count("=")
    ]
    # the item 01 value is the last amount before the first lot that is not part of a '=' calculation
    item01_val = brl(it01[-1].group(1))
    item01_page = page_of(head + it01[-1].start())
    rows = []
    bounds = lot_starts + [total_pos]
    for a, b in pairwise(bounds):
        block = text[a:b]
        mats = re.findall(r"Matrícula\s+nº\s+([\d\.]+\d)", block)
        insc = re.findall(r"Inscrição\s+Imobiliária\s+nº\s+([\d\.]+\d)", block)
        vals = [
            m
            for m in re.finditer(r"R\$\s*" + MONEY, block)
            if not re.search(
                r"global\s+de\s*$", block[max(0, m.start() - 15) : m.start()]
            )
        ]
        assert len(mats) == len(insc) == len(vals) == 1, (
            block[:80],
            mats,
            insc,
            len(vals),
        )
        rows.append(
            {
                "inscricao": insc[0],
                "matricula": mats[0],
                "valor": brl(vals[0].group(1)),
                "pagina": page_of(a + vals[0].start()),
            }
        )
    tail = text[total_pos:]
    # Rafael's p. 20 prints the total with one decimal digit ("R$76.230.000,0").
    total = brl(first(r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{1,2})", tail))
    return item01_val, item01_page, rows, total


AMOUNT = r"R\$ ?(\d{1,3}(?:\.\d{3})+(?:,\d{2})?|\d+(?:,\d+)? milhões)"
RANGE_BEM = {
    7: "01 Hotel",
    8: "01 Hotel",
    10: "Lotes q37/q38",
    11: "Lotes q37/q38",
    13: "20 Condado",
    14: "20 Condado",
    15: "20 Condado",
    24: "Casas Rua da Divisa",
    25: "Casas Rua da Divisa",
}


def amount(s: str) -> float:
    if "milhões" in s:
        return float(s.split()[0].replace(",", ".")) * 1_000_000
    return brl(s) if "," in s else float(s.replace(".", ""))


def rafael_ranges(pages):
    rows = []
    for i in range(5, 26):
        flat = re.sub(r"\s+", " ", pages[i - 1])
        for m in re.finditer(
            r"(?i:entre|de|faixa de|faixa dos) " + AMOUNT + r" (?:e|a) " + AMOUNT, flat
        ):
            ctx = flat[max(0, m.start() - 160) : m.end() + 20]
            rows.append(
                [
                    i,
                    RANGE_BEM.get(i),
                    "faixa",
                    amount(m.group(1)),
                    amount(m.group(2)),
                    ctx,
                ]
            )
        for m in re.finditer(
            r"(Valor (?:Mínimo|Máximo) (?:Físico|por Renda)): Avaliado em " + AMOUNT,
            flat,
        ):
            rows.append(
                [
                    i,
                    RANGE_BEM.get(i),
                    m.group(1),
                    amount(m.group(2)),
                    None,
                    flat[m.start() : m.end() + 160],
                ]
            )
        for m in re.finditer(r"(?:R\$ ?[\d\.]+/m²)", flat):
            rows.append(
                [
                    i,
                    RANGE_BEM.get(i),
                    "R$/m²",
                    None,
                    None,
                    flat[max(0, m.start() - 120) : m.end() + 40],
                ]
            )
    return rows


def sentences_with_money(pages, page_list):
    out = []
    for i in page_list:
        flat = re.sub(r"\s+", " ", pages[i - 1])
        for s in re.split(r"(?<=\.)\s+(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9(\[])", flat):
            if "R$" in s:
                out.append([i, s.strip()])
    return out


# ---------------------------------------------------------------- main


def main():
    wb = openpyxl.Workbook()
    checks = []

    def check(name, got, expected, src):
        assert got is not None and expected is not None, (name, got, expected)
        ok = abs(got - expected) < 0.005
        checks.append([name, got, expected, "CONFERE" if ok else "DIVERGE", src])
        return ok

    ws = wb.active
    ws.title = "Fontes"
    ws.append(["chave", "caminho", "md5 (10)"])
    for k, p in SRC.items():
        ws.append([k, str(p), md5_10(p)])

    # inventories
    inv, totals = {}, {}
    for key, label in [
        ("inv2018", "2018"),
        ("inv2025", "2025 atual"),
        ("inv2025_antigo", "2025 antigo"),
    ]:
        items, total = parse_inventory(SRC[key], label)
        inv[label], totals[label] = items, total
        check(
            f"Soma dos itens com valor – {label} ({len(items)} itens)",
            sum(i["valor"] or 0 for i in items),
            total,
            f"{key}: TOTAL GERAL DA EMPRESA",
        )
    assert (
        len(inv["2018"]) == 25
        and len(inv["2025 atual"]) == 22
        and len(inv["2025 antigo"]) == 22
    )
    sh = wb.create_sheet("Itens")
    cols = list(inv["2018"][0].keys())
    sh.append(cols)
    for items in inv.values():
        for it in items:
            sh.append([it[c] for c in cols])

    obs = wb.create_sheet("Observacoes_inventario")
    obs.append(["fonte", "nº", "parágrafo"])
    for key, label in [
        ("inv2018", "2018"),
        ("inv2025", "2025 atual"),
        ("inv2025_antigo", "2025 antigo"),
    ]:
        for n, t in enumerate(observation_paragraphs(SRC[key]), 1):
            obs.append([label, n, t])
    obs25 = " ".join(observation_paragraphs(SRC["inv2025"]))
    inss = brl(first(r"aproximadamente R\$ " + MONEY, obs25))
    reducao = brl(first(r"redução de aproximadamente 4,00% = R\$ " + MONEY, obs25))
    check(
        "Obs. 2025: INSS da obra (não deduzido do total)",
        inss,
        1601357.90,
        "inventário 2025, observações",
    )
    check(
        "Obs. 2025: redução estimada do item 01 (não deduzida)",
        reducao,
        2000000.00,
        "inventário 2025, observações",
    )

    # item 01 components
    comp = wb.create_sheet("Item01_componentes")
    comp_rows = {}
    for label, items in inv.items():
        rows = item01_components(items[0]["descricao"], label)
        comp_rows[label] = rows
        check(
            f"Componentes do item 01 – {label}",
            sum(r["valor"] for r in rows),
            items[0]["valor"],
            "descrição do item 01",
        )
    comp.append(list(comp_rows["2018"][0].keys()))
    for rows in comp_rows.values():
        for r in rows:
            comp.append(list(r.values()))

    # Rafael PDF
    pages = pdf_pages(SRC["rafael_pdf"])
    t01, t01p, transc, transc_total = balance_section(pages, 1, 5)
    e01, e01p, execu, exec_total = balance_section(pages, 16, 20)
    check(
        "Rafael – transcrição (p. 1–5), item 01 = inventário 2025",
        t01,
        inv["2025 atual"][0]["valor"],
        f"PDF Rafael p. {t01p}",
    )
    check(
        "Rafael – transcrição do balanço do grupo (p. 1–5): item 01 + 21 lotes",
        t01 + sum(r["valor"] for r in transc),
        transc_total,
        "PDF Rafael p. 5",
    )
    check(
        "Rafael – balanço executado (p. 16–20): item 01 + 21 lotes",
        e01 + sum(r["valor"] for r in execu),
        exec_total,
        "PDF Rafael p. 20",
    )
    raf = wb.create_sheet("Rafael_itens")
    raf.append(
        [
            "item 2025 (por inscrição)",
            "inscricao",
            "matricula",
            "valor grupo (transcrição Rafael)",
            "pág. transcrição",
            "valor Rafael (balanço executado)",
            "pág. executado",
            "valor grupo (inventário 2025)",
        ]
    )
    g2025 = {i["inscricao"]: i for i in inv["2025 atual"] if i["inscricao"]}
    raf.append(
        [
            "01",
            None,
            "12 matrículas",
            t01,
            t01p,
            e01,
            e01p,
            inv["2025 atual"][0]["valor"],
        ]
    )
    check(
        "Rafael executado = grupo, item 01",
        e01,
        inv["2025 atual"][0]["valor"],
        f"PDF Rafael p. {e01p}",
    )
    for t, e in zip(transc, execu):
        assert t["inscricao"] == e["inscricao"] and t["matricula"] == e["matricula"]
        g = g2025[t["inscricao"]]
        assert g["matricula"] == t["matricula"]
        raf.append(
            [
                g["item"],
                t["inscricao"],
                t["matricula"],
                t["valor"],
                t["pagina"],
                e["valor"],
                e["pagina"],
                g["valor"],
            ]
        )
        check(
            f"Transcrição de Rafael = inventário 2025, item {g['item']}",
            t["valor"],
            g["valor"],
            f"PDF Rafael p. {t['pagina']}",
        )
        check(
            f"Rafael executado = grupo, item {g['item']}",
            e["valor"],
            g["valor"],
            f"PDF Rafael p. {e['pagina']}",
        )
    item20 = next(e for e in execu if e["inscricao"] == RAFAEL_ITEM20_INSCRICAO)
    check(
        "Rafael executado, item 20 (por inscrição) = 2.400.000 da p. 19",
        item20["valor"],
        2400000.00,
        "PDF Rafael p. 19",
    )

    fx = wb.create_sheet("Rafael_faixas")
    fx.append(["página", "bem", "tipo", "mínimo / valor", "máximo", "trecho"])
    for r in rafael_ranges(pages):
        fx.append(r)
    tx = wb.create_sheet("Rafael_textos")
    tx.append(["página", "frase com valor"])
    for r in sentences_with_money(pages, range(5, 27)):
        tx.append(r)

    selected = brl(
        first(r"\(02\.,11\.,12\., 18\. e o 19\)\s*R\$\s*" + MONEY, pages[20])
    )
    credit = brl(first(r"R\$\s*" + MONEY, pages[20].split("selecionados de")[1]))
    sixth = brl(first(r"R\$\s*" + MONEY, pages[20].split("1/6 do patrimônio total")[1]))
    check("Rafael – 1/6 do total declarado", exec_total / 6, sixth, "PDF Rafael p. 21")
    check(
        "Rafael – crédito = 1/6 − selecionados",
        sixth - selected,
        credit,
        "PDF Rafael p. 21",
    )

    # Rafael's list (C0)
    lr = openpyxl.load_workbook(SRC["lista_rafael"], data_only=True)["Ativos"]
    lst = wb.create_sheet("Lista_Rafael")
    lst.append(
        [
            "item 2025",
            "lote",
            "quadra",
            "matrícula",
            "valor na lista",
            "pág. Diretrizes",
            "observação da lista",
        ]
    )
    lsum = 0.0
    for r in lr.iter_rows(min_row=2, values_only=True):
        if r[1] and str(r[1]).startswith("06 -"):
            lst.append([r[3], r[6], r[7], r[8], r[10], r[13], r[14]])
            lsum += r[10]
    check(
        "Lista de Rafael – soma dos 5 itens da empresa 06",
        lsum,
        selected,
        "PDF Rafael p. 21–22",
    )

    # 2018 and 2025 summary documents
    t27 = " ".join(c for _, cells in table_rows(SRC["resumo2018"]) for c in cells)
    t00 = " ".join(c for _, cells in table_rows(SRC["doc00_2025"]) for c in cells)
    check(
        "Doc 27 (resumo 2018) × inventário 2018",
        brl(first(r"HOTEL ARAÇA LTDA\. R\$\s+" + MONEY, t27)),
        totals["2018"],
        "doc 27 de 2018",
    )
    check(
        "Doc 00 (2025) × inventário 2025",
        brl(first(r"06\. HOTEL ARAÇA LTDA\. R\$\s+" + MONEY, t00)),
        totals["2025 atual"],
        "doc 00 de 2025",
    )
    check(
        "Doc 00 (2025), cota por sócio × total / 6",
        brl(first(r"06\. HOTEL ARAÇA LTDA\.\s+COTA .*?R\$\s+" + MONEY, t00)),
        round(totals["2025 atual"] / 6, 2),
        "doc 00 de 2025",
    )

    # Google sheets
    gs = wb.create_sheet("Planilhas_Google")
    gs.append(
        [
            "planilha",
            "aba",
            "linha",
            "item",
            "2018 apresentado",
            "2018 atualizado",
            "col. F sem cabeçalho (= 2025 antigo)",
            "outubro 2025",
            "'Estudo Rafael' (só os 5 lotes escolhidos, não a avaliação)",
            "diferença",
        ]
    )
    g = openpyxl.load_workbook(SRC["gsheet_2_0"], data_only=True)
    for r in g["6.Hotel Araçá"].iter_rows(min_row=2, max_row=34):
        v = [c.value for c in r]
        if v[0] is not None or v[1] == "Lotes":
            gs.append(
                [
                    "2_0",
                    "6.Hotel Araçá",
                    r[0].row,
                    v[0] if v[0] is not None else "subtotal Lotes",
                    v[2],
                    v[3],
                    v[5],
                    v[6],
                    None,
                    v[7],
                ]
            )
    for row in range(70, 103):
        v = [c.value for c in g["GERAL"][row]]
        if v[0] is not None or v[1] == "Lotes":
            gs.append(
                [
                    "2_0",
                    "GERAL",
                    row,
                    v[0] if v[0] is not None else "subtotal Lotes",
                    v[2],
                    v[3],
                    None,
                    v[4],
                    None,
                    v[6],
                ]
            )
    g2 = openpyxl.load_workbook(SRC["gsheet_rafael"], data_only=True)
    for r in g2["6.Hotel Araçá"].iter_rows(min_row=2, max_row=34):
        v = [c.value for c in r]
        if v[0] is not None or v[1] == "Lotes":
            gs.append(
                [
                    "empreend. e Rafael",
                    "6.Hotel Araçá",
                    r[0].row,
                    v[0] if v[0] is not None else "subtotal Lotes",
                    v[2],
                    v[3],
                    v[5],
                    v[6],
                    v[7],
                    v[8],
                ]
            )

    total18 = totals["2018"]
    lots18 = sum(i["valor"] or 0 for i in inv["2018"][1:])
    paired = {
        norm_inscricao(i["inscricao"])
        for i in inv["2018"]
        if i["inscricao"] and i["valor"]
    }
    new_2025 = sum(
        i["valor"] for i in inv["2025 atual"][1:] if i["inscricao_norm"] not in paired
    )
    # "2018 atualizado" = the items valued in 2018, at their 2025 values
    updated_2018 = inv["2025 atual"][0]["valor"] + sum(
        i["valor"] for i in inv["2025 atual"][1:] if i["inscricao_norm"] in paired
    )
    check(
        "GERAL C70 (2018) × total do inventário 2018",
        g["GERAL"]["C70"].value,
        total18,
        "GERAL C70",
    )
    check(
        "GERAL C80 (lotes 2018) × lotes do inventário 2018",
        g["GERAL"]["C80"].value,
        lots18,
        "GERAL C80",
    )
    check(
        "GERAL D70 (2018 atualizado) × 2018 com o lote 28/q38 em 350.000",
        g["GERAL"]["D70"].value,
        updated_2018,
        "GERAL D70, D99",
    )
    check(
        "GERAL G70 (diferença) × lotes 2025 sem par com valor em 2018",
        g["GERAL"]["G70"].value,
        new_2025,
        "GERAL G70",
    )
    check(
        "GERAL E70 (out/2025) × inventário 2025",
        g["GERAL"]["E70"].value,
        totals["2025 atual"],
        "GERAL E70",
    )
    a = g["6.Hotel Araçá"]
    check(
        "Aba 6.Hotel Araçá (2_0) C2 × total do inventário 2018",
        a["C2"].value,
        total18,
        "6.Hotel Araçá C2",
    )
    check(
        "Aba 6.Hotel Araçá (2_0) H2 × lotes 2025 sem par com valor em 2018",
        a["H2"].value,
        new_2025,
        "6.Hotel Araçá H2",
    )
    check(
        "Aba 6.Hotel Araçá (2_0) F2 × inventário 2025 antigo",
        a["F2"].value,
        totals["2025 antigo"],
        "6.Hotel Araçá F2",
    )
    check(
        "Planilha empreend. e Rafael, H2 'Estudo Rafael' × lista de Rafael",
        g2["6.Hotel Araçá"]["H2"].value,
        lsum,
        "6.Hotel Araçá H2",
    )

    ck = wb.create_sheet("Conferencias", 1)
    ck.append(
        [
            "conferência",
            "valor extraído",
            "valor esperado (documento)",
            "resultado",
            "fonte",
        ]
    )
    for c in checks:
        ck.append(c)

    out = OUT_DIR / "fase0-06.xlsx"
    wb.save(out)
    for c in checks:
        if c[3] == "DIVERGE" or not c[0].startswith(
            ("Transcrição", "Rafael executado = grupo")
        ):
            print(f"{c[3]:8} {c[0]}: {c[1]:,.2f} x {c[2]:,.2f}")
    diverge = {c[0] for c in checks if c[3] == "DIVERGE"}
    assert diverge == EXPECTED_DIVERGE, (
        "new:",
        diverge - EXPECTED_DIVERGE,
        "gone:",
        EXPECTED_DIVERGE - diverge,
    )
    print(f"{len(checks)} checks; divergences as documented: {len(diverge)}")
    print("saved", out, md5_10(out))


if __name__ == "__main__":
    main()
