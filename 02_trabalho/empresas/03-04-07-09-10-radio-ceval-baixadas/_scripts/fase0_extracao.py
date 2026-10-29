"""Phase 0 extraction for companies 03, 04, 07, 09 and 10 of the Siviero split study.

Reads the 2018 inventories (Drive), the current and old 2025 inventories, Rafael's
Estudo 2 PDFs (every section), the Valter selection list and the Google sheet 2_0,
and writes one workbook with every item, every declared total and the check of
each extracted sum against the document total. Labels are in Portuguese
(deliverable language); code is in English.

Usage: python fase0_extracao.py   (paths are fixed below; nothing is edited by hand)
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

import docx
import openpyxl
import pymupdf
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(r"C:\entrega-cisao-siviero")
F = ROOT / "01_fontes"
DRIVE18 = Path(
    r"G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018"
)
OUT_DIR = (
    ROOT
    / "02_trabalho"
    / "empresas"
    / "03-04-07-09-10-radio-ceval-baixadas"
    / "_extracao"
)

COMPANIES = {
    "03": "Associação Cultural Princesa dos Vales",
    "04": "Ceval Terraplenagem e Escavações",
    "07": "Rádio Comunidade Princesa dos Vales FM",
    "09": "Turismo Siviero",
    "10": "Siviero Transportes",
}


def one(pattern: str, folder: Path) -> Path:
    hits = sorted(folder.glob(pattern))
    if len(hits) != 1:
        sys.exit(f"expected exactly one file for {pattern!r} in {folder}, got {hits}")
    return hits[0]


SRC_2018 = {
    "03": one("14. 1. ASSOCIAÇÃO*.docx", DRIVE18),
    "04": one("11. 1. CEVAL*.docx", DRIVE18),
    "07": one("9. 1. RÁDIO*.docx", DRIVE18),
    "09": one("8. 1. TURISMO*.docx", DRIVE18),
    "10": one("7 .1. SIVIERO TRANSPORTES*.docx", DRIVE18),
}
SRC_2025 = {c: one(f"{c}. INVENTÁRIO*.docx", F / "inventario-2025") for c in COMPANIES}
SRC_2025_OLD = {
    c: one(f"ANTIGO {c}. INVENTÁRIO*.docx", F / "inventario-2025-antigo")
    for c in COMPANIES
}
SRC_RAFAEL = {
    "03": F
    / "estudo-rafael"
    / "ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES LTDA - BELISE LTDA.pdf",
    "04": F
    / "estudo-rafael"
    / "CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA, BELISE LTDA.pdf",
    "07": F
    / "estudo-rafael"
    / "RADIO COMUNIDADE PRINCESA DOS VALES FM LTDA, BELISE LTDA.pdf",
    "09": F / "estudo-rafael" / "TURISMO SIVIERO LTDA - BELISE LTDA.pdf",
    "10": F / "estudo-rafael" / "SIVIERO TRANSPORTES LTDA - BELISE LTDA.pdf",
}
SRC_DIRETRIZES = one("DIRETRIZES*.pdf", F / "estudo-rafael")
SRC_INTRO = one("INTRODUÇÃO*.pdf", F / "estudo-rafael")
SRC_DOC27 = one("27. 1. *.docx", F / "inventario-2018")
SRC_DOC00 = one("00. *.docx", F / "inventario-2025")
SRC_DOC00_OLD = one("ANTIGO 00. *.docx", F / "inventario-2025-antigo")
SRC_DOC23 = one("23. 1. *.docx", DRIVE18)
SRC_VALTER = F / "outros" / "ativos-avaliados-valter.xlsx"
SRC_GSHEET = (
    F / "planilhas-google" / "2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx"
)
SRC_GSHEET_RAF = (
    F
    / "planilhas-google"
    / "Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx"
)
SRC_OBS = ROOT / "02_trabalho" / "observacoes-planilhas-google.xlsx"
# Cross-reference sources (reviewer, Phase 0): SEI 2018 and 2025, lot division 2018, partners' meeting.
SRC_DOC18 = one("18. 1. *.docx", F / "inventario-2018")
SRC_DOC25 = one("25. 1. *.docx", F / "inventario-2018")
SRC_DOC11 = one("11. INVENTÁRIO*.docx", F / "inventario-2025")
# Doc 22 has a path longer than 260 characters: open it through the \\?\ prefix.
SRC_DOC22 = Path("\\\\?\\" + str(one("22. 1. *.docx", DRIVE18)))
SRC_RAF_XLSX = F / "outros" / "inventário-estudo-rafael.xlsx"
SRC_ACERTOS = F / "inventario-2025" / "ACERTOS E ACORDOS.pdf"

# Other sources used in later steps: hashed now so the journal has them from Phase 0.
RESP = F / "respostas-docs-2026-08-13"
SRC_OTHER = [
    one("1 - *.pdf", RESP),
    one("5 - *.pdf", RESP),
    one("6 - *.pdf", RESP),
    one("10 - *.pdf", RESP),
    one("13 - *.pdf", RESP),
    one("14 - *.pdf", RESP),
    F / "outros" / "conferencia-discrepancias.md",
    F / "outros" / "inventario-2025-avaliacoes-corrigido.xlsx",
    ROOT / "02_trabalho" / "diff-inventario-2025-antigo-x-atual.md",
    ROOT / "02_trabalho" / "parte6-candidatos-duplicidade.md",
    ROOT
    / "02_trabalho"
    / "analises"
    / "acertos-e-areas-vendidas"
    / "achados-para-empresas.md",
]

# Google sheet 2_0: rows of each company in GERAL and the company tabs (brief, section 3).
GERAL_ROWS = {
    "03": (16, 19),
    "04": (20, 27),
    "07": (105, 115),
    "09": (135, 138),
    "10": (139, 152),
}
GSHEET_TABS = {"04": "4.Ceval", "07": "7.Rádio", "10": "10.Transportes"}

MONEY = re.compile(r"R\$\s*\(?-?\)?\s*([\d]{1,3}(?:\.\d{3})*(?:,\d{2})?)")
PCT = re.compile(r"(\d{1,3},\d{1,3})\s*%")


def md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def brl(s: str) -> float:
    return float(s.replace(".", "").replace(",", "."))


def moneys(text: str) -> list[float]:
    return [brl(m) for m in MONEY.findall(text)]


def keys_of(text: str) -> str:
    """Identity keys (matrícula, placa, chassi, NF) found in a description."""
    found = []
    for m in re.finditer(r"Matr[íi]cula\s*n?º?\s*([\d.]+\d)", text, re.IGNORECASE):
        found.append(f"mat. {m.group(1)}")
    for m in re.finditer(r"\b(?:placa|PLACA)[:\s]+([A-Z]{3}\s?-?\d{4})", text):
        found.append("placa " + re.sub(r"[\s-]", "", m.group(1)))
    for m in re.finditer(r"chassi\s*n?º?\s*([A-Z0-9]{8,})", text, re.IGNORECASE):
        found.append(f"chassi {m.group(1)}")
    for m in re.finditer(r"Nota Fiscal\s*n?º?\s*([\d.]+)", text, re.IGNORECASE):
        found.append(f"NF {m.group(1)}")
    return "; ".join(dict.fromkeys(found))


def docx_blocks(path: Path):
    """Body in order: ('p', text) or ('t', rows) with merged cells collapsed."""
    d = docx.Document(path)
    for ch in d.element.body.iterchildren():
        tag = ch.tag.split("}")[1]
        if tag == "p":
            t = Paragraph(ch, d).text.strip()
            if t:
                yield "p", t
        elif tag == "tbl":
            rows = []
            for r in Table(ch, d).rows:
                cells, prev = [], None
                for c in r.cells:
                    if c._tc is prev:
                        continue
                    prev = c._tc
                    cells.append(c.text.strip())
                rows.append(cells)
            yield "t", rows


def parse_inventory(path: Path, company: str, version: str):
    """Items, declared total, partner split table and free-text notes of one inventory."""
    items, notes, partners, quadros = [], [], [], []
    declared = None
    declared_gross = None
    t_idx = 0
    for kind, block in docx_blocks(path):
        if kind == "p":
            notes.append(block)
            continue
        t_idx += 1
        header = " ".join(block[0]).upper()
        if header.startswith("ITEM") and "DESCRI" in header:
            for row in block[1:]:
                first = row[0]
                joined = " | ".join(row)
                if first.upper().startswith("TOTAL"):
                    vals = moneys(joined)
                    if vals:
                        declared_gross = vals[0]
                        declared = vals[-1]
                    continue
                desc = row[1] if len(row) > 1 else ""
                val_cell = row[2] if len(row) > 2 else ""
                vals = moneys(val_cell)
                status = ""
                low = val_cell.lower()
                if "vendido" in low:
                    status = re.sub(
                        r"\s+", " ", val_cell[val_cell.lower().index("vendido") :]
                    ).strip()
                empty = (not desc.strip(".- ")) and not vals
                in_desc = moneys(desc)
                alert = ""
                if in_desc and vals and all(abs(x - vals[0]) >= 0.005 for x in in_desc):
                    alert = "valor citado na descrição difere da coluna de valor"
                items.append(
                    {
                        "empresa": company,
                        "versao": version,
                        "tabela": t_idx,
                        "item": first,
                        "descricao": re.sub(r"\s+", " ", desc).strip(),
                        "valor": vals[0] if vals else None,
                        "valor_descricao": "; ".join(f"{x:,.2f}" for x in in_desc),
                        "alerta": alert,
                        "situacao_no_doc": status,
                        "vazio": empty,
                        "chaves": keys_of(desc),
                    }
                )
        elif (
            "NOME DO SÓCIO" in header
            and "VALOR" in header
            and "COTAS DE CAPITAL E" in header.replace("\n", " ")
        ):
            for row in block[1:]:
                pct = PCT.findall(row[-1])
                vals = moneys(row[-1])
                partners.append(
                    {
                        "empresa": company,
                        "versao": version,
                        "socio": row[0],
                        "pct": pct[0] if pct else "",
                        "valor": vals[0] if vals else None,
                    }
                )
        elif "NOME DO SÓCIO" in header:
            quadros.append(
                {
                    "empresa": company,
                    "versao": version,
                    "tabela": t_idx,
                    "linhas": [" | ".join(r) for r in block[1:]],
                }
            )
    return {
        "items": items,
        "declared": declared,
        "declared_gross": declared_gross,
        "partners": partners,
        "notes": notes,
        "quadros": quadros,
    }


# --- Rafael PDFs ---------------------------------------------------------------------------

SECTIONS = [
    ("apresentado_grupo", r"BALANÇO PATRIMONIAL APRESENTADO PELO GRUPO"),
    ("avaliacao_valter", r"AVALIAÇÃO PATRIMONIAL EXECUTADA POR VALTER"),
    ("balanco_valter", r"BALANÇO PATRIMONIAL EXECUTADO POR VALTER"),
    ("quadro_resumo", r"QUADRO-RESUMO"),
    ("bens_escolhidos", r"BENS MÓVEIS ESCOLHIDOS|BENS IMÓVEIS ESCOLHIDOS"),
    ("ausencia_interesse", r"DA AUSÊNCIA DE INTERESSE"),
    ("baixa", r"DA BAIXA E DA REGULARIDADE"),
    ("selecao", r"DA SELEÇÃO DE BENS"),
    ("escopo", r"DECLARAÇÃO DE ESCOPO"),
]
ITEM_LINE = re.compile(r"^\s*0?(\d{1,2})\s*\.?\s*$")
RANGE = re.compile(r"R\$\s*([\d.]+(?:,\d{2})?)\s*(?:e|a)\s*R\$\s*([\d.]+(?:,\d{2})?)")


def pdf_lines(path: Path):
    doc = pymupdf.open(path)
    out = []
    for pno, page in enumerate(doc, 1):
        for line in page.get_text().splitlines():
            if line.strip():
                out.append((pno, line.strip()))
    return out, len(doc)


def parse_rafael(path: Path, company: str):
    lines, npages = pdf_lines(path)
    # Section boundaries on the single-line joined stream (headers can wrap across lines).
    joined = []
    pos = []
    for i, (p, t) in enumerate(lines):
        pos.append(sum(len(x) + 1 for x in joined))
        joined.append(t)
    text = "\n".join(joined)
    marks = []
    for name, pat in SECTIONS:
        for m in re.finditer(pat, text.replace("\n", " ")):
            marks.append((m.start(), name))
    marks.sort()

    def section_at(offset: int) -> str:
        cur = "cabecalho"
        for start, name in marks:
            if start <= offset:
                cur = name
        return cur

    def new_item(sec, number, p):
        return {
            "empresa": company,
            "secao": sec,
            "item": number,
            "pagina": p,
            "descricao": "",
            "valor": None,
            "pagina_valor": None,
            "faixa_min": None,
            "faixa_max": None,
            "pagina_faixa": None,
            "texto_faixa": "",
            "bloco": [],
        }

    items = []
    totals = []
    current = None
    for i, (p, t) in enumerate(lines):
        sec = section_at(pos[i])
        if current is not None and current["secao"] != sec:
            current = None
        m = ITEM_LINE.match(t)
        if m and sec in ("apresentado_grupo", "avaliacao_valter", "balanco_valter"):
            current = new_item(sec, int(m.group(1)), p)
            items.append(current)
            continue
        # The chosen-assets section has no item numbers: an item starts at "Um/Uma ...".
        if (
            sec == "bens_escolhidos"
            and current is None
            and re.match(r"^(Um|Uma|Lote|Terreno)\b", t)
        ):
            current = new_item(sec, None, p)
            items.append(current)
        # "TOTAL GERAL" may be split over lines ("TOTAL" / "GERAL" / "DA" / "EMPRESA").
        head = " ".join(x[1] for x in lines[i : i + 2]).upper()
        if t.upper().startswith("TOTAL") and "TOTAL GERAL" in re.sub(r"\s+", " ", head):
            for j in range(i, min(i + 6, len(lines))):
                v = moneys(lines[j][1]) or (
                    [brl(lines[j][1])]
                    if re.fullmatch(r"[\d.]+,\d{2}", lines[j][1])
                    else []
                )
                if v:
                    totals.append(
                        {
                            "empresa": company,
                            "secao": sec,
                            "valor": v[0],
                            "pagina": lines[j][0],
                        }
                    )
                    break
            current = None
            continue
        if current is None:
            continue
        if current["valor"] is None:
            v = moneys(t)
            if (
                not v
                and re.fullmatch(r"[\d.]+,\d{2}", t)
                and i > 0
                and lines[i - 1][1].strip() == "R$"
            ):
                v = [brl(t)]
            if v and (t.startswith("R$") or re.fullmatch(r"[\d.]+,\d{2}", t)):
                current["valor"] = v[0]
                current["pagina_valor"] = p
            elif t != "R$":
                current["descricao"] = (current["descricao"] + " " + t).strip()
            continue
        current["bloco"].append((p, t))
        if current["faixa_min"] is None and sec == "avaliacao_valter":
            window = " ".join(x[1] for x in lines[i : i + 3])
            r = RANGE.search(window)
            if r:
                current["faixa_min"], current["faixa_max"] = (
                    brl(r.group(1)),
                    brl(r.group(2)),
                )
                current["pagina_faixa"] = p
                s = window[max(0, r.start() - 160) : r.end() + 20]
                current["texto_faixa"] = re.sub(r"\s+", " ", s)
    for it in items:
        it["chaves"] = keys_of(it["descricao"])
        # Every other R$ figure in the item's block (FIPE, per-version prices, box range...).
        refs = []
        blk = it.pop("bloco")
        btxt = " ".join(t for _, t in blk)
        for k, (p, t) in enumerate(blk):
            # "R$" can end a line with the number on the next one: search line + next line,
            # keep only matches that start on this line.
            nxt = blk[k + 1][1] if k + 1 < len(blk) else ""
            for mm in MONEY.finditer(t + " " + nxt):
                if mm.start() >= len(t):
                    continue
                ctx = re.sub(
                    r"\s+", " ", " ".join(x[1] for x in blk[max(0, k - 2) : k + 2])
                )
                refs.append(f"p. {p}: R$ {mm.group(1)} — …{ctx[-180:]}")
        it["outras_referencias"] = "\n".join(dict.fromkeys(refs))
        if it["secao"] == "avaliacao_valter":
            if it["faixa_min"] is not None:
                it["criterio"] = "faixa de mercado"
            elif "CONCORDÂNCIA" in btxt.upper():
                it["criterio"] = "concordância expressa com o valor do grupo"
            else:
                it["criterio"] = "sem avaliação própria"
        else:
            it["criterio"] = ""
    # Quadro-resumo: money figures in order, up to the "Leitura do quadro" paragraph.
    quadro, quadro_vals, fracs = [], [], []
    reading = False
    for i, (p, t) in enumerate(lines):
        s = section_at(pos[i])
        if s in ("quadro_resumo", "bens_escolhidos") and (
            moneys(t) or re.search(r"escolhid|selecionad|bens", t, re.IGNORECASE)
        ):
            quadro.append({"empresa": company, "pagina": p, "linha": t})
        if s == "quadro_resumo":
            if t.startswith("Leitura"):
                reading = True
            if not reading:
                quadro_vals += [(v, p) for v in moneys(t)]
    qtext = " ".join(
        t for i, (p, t) in enumerate(lines) if section_at(pos[i]) == "quadro_resumo"
    )
    for mm in re.finditer(r"(\d+,\d+)\s*bens\s+(\w+)", qtext):
        fracs.append(f"{mm.group(1)} bens {mm.group(2)}")
    sec_pages = {}
    for i, (p, t) in enumerate(lines):
        sec_pages.setdefault(section_at(pos[i]), set()).add(p)
    return {
        "items": items,
        "totals": totals,
        "quadro": quadro,
        "quadro_vals": quadro_vals,
        "fracoes_bens": fracs,
        "texto": re.sub(r"\s+", " ", text),
        "npages": npages,
        "secoes": {k: sorted(v) for k, v in sec_pages.items()},
    }


def diretrizes_hits():
    lines, _ = pdf_lines(SRC_DIRETRIZES)
    hits = []
    pat = re.compile(
        r"R[ÁA]DIO|CEVAL|TURISMO|SIVIERO TRANSPORTES|ASSOCIA[ÇC][ÃA]O|BANDEIRANTE|IHW|320\s?D|IKI|IHM|IRO 2562|11\.382|11\.324",
        re.IGNORECASE,
    )
    for i, (p, t) in enumerate(lines):
        if pat.search(t):
            ctx = " ".join(x[1] for x in lines[i : i + 6])
            hits.append(
                {
                    "fonte": "Diretrizes",
                    "pagina": p,
                    "linha": t,
                    "contexto": re.sub(r"\s+", " ", ctx)[:300],
                }
            )
    lines, _ = pdf_lines(SRC_INTRO)
    for i, (p, t) in enumerate(lines):
        if pat.search(t):
            ctx = " ".join(x[1] for x in lines[i : i + 6])
            hits.append(
                {
                    "fonte": "Introdução",
                    "pagina": p,
                    "linha": t,
                    "contexto": re.sub(r"\s+", " ", ctx)[:300],
                }
            )
    return hits


def summary_table(path: Path):
    """Company totals from doc 27 (2018) or doc 00 (2025)."""
    out = {}
    for kind, block in docx_blocks(path):
        if kind != "t":
            continue
        if "E M P R E S A" not in " ".join(block[0]):
            continue
        for row in block[1:]:
            name = row[0].split("\n")[0].upper()
            vals = moneys(row[-1].replace("100.874,027,60", "100.874.027,60"))
            for code, label in [
                ("03", "ASSOCIA"),
                ("04", "CEVAL"),
                ("07", "RÁDIO"),
                ("09", "TURISMO"),
                ("10", "SIVIERO TRANSPORTES"),
            ]:
                if label in name and vals and code not in out:
                    out[code] = vals[0]
        if out:
            break
    return out


def header_map(ws) -> dict[str, int]:
    """Column index by header text (row 1), so columns are never read by position."""
    out = {}
    for c in ws[1]:
        label = re.sub(r"\s+", " ", str(c.value or "")).strip().lower()
        if not label:
            continue
        for key, pat in [
            ("2018_apresentado", r"^2018: apresentado"),
            ("2018_atualizado", r"^2018: atualizado"),
            ("dez2024", r"^dezembro 2024"),
            ("out2025", r"^outubro 2025"),
            ("divergencia", r"^diverg"),
            ("diferenca", r"^diferen"),
            ("estudo_rafael", r"estudo rafael"),
            ("observacao", r"^observa"),
        ]:
            if re.search(pat, label) and key not in out:
                out[key] = c.column
    return out


GS_KEYS = [
    "2018_apresentado",
    "2018_atualizado",
    "dez2024",
    "auxiliar",
    "out2025",
    "estudo_rafael",
    "divergencia",
    "diferenca",
    "observacao",
]


def gsheet_rows():
    rows = []
    for wb_path, wb_label in [(SRC_GSHEET, "2_0"), (SRC_GSHEET_RAF, "emp. e Rafael")]:
        wb = openpyxl.load_workbook(wb_path, data_only=True)
        targets = []
        if wb_label == "2_0":
            targets += [("GERAL", code, a, b) for code, (a, b) in GERAL_ROWS.items()]
        for code, tab in GSHEET_TABS.items():
            ws = wb[tab]
            start = 2
            if (
                code == "07"
            ):  # rows 2-33 of the tab are a copy of Hotel Araçá; Rádio starts at its header
                for r in range(1, ws.max_row + 1):
                    if str(ws.cell(r, 2).value or "").upper().startswith("RÁDIO"):
                        start = r
                        break
            targets.append((tab, code, start, ws.max_row))
        for tab, code, a, b in targets:
            ws = wb[tab]
            hm = header_map(ws)
            # Unlabelled column between "dezembro 2024" and "outubro 2025": helper sums.
            if (
                "dez2024" in hm
                and "out2025" in hm
                and hm["out2025"] - hm["dez2024"] == 2
            ):
                hm["auxiliar"] = hm["dez2024"] + 1
            for r in range(a, b + 1):
                item, desc = ws.cell(r, 1).value, ws.cell(r, 2).value
                vals = {
                    k: (ws.cell(r, hm[k]).value if k in hm else None) for k in GS_KEYS
                }
                if (
                    item in (None, "")
                    and desc in (None, "")
                    and all(v in (None, "") for v in vals.values())
                ):
                    continue
                rows.append(
                    {
                        "planilha": wb_label,
                        "empresa": code,
                        "aba": tab,
                        "linha": r,
                        "item": item,
                        "descricao": re.sub(r"\s+", " ", str(desc or ""))[:300],
                        **vals,
                    }
                )
    ws = openpyxl.load_workbook(SRC_GSHEET, data_only=True)["7.Rádio"]
    araca_copy = None
    if "MIGUEL MAIN FELO" in str(ws.cell(2, 2).value or "").upper():
        araca_copy = "Aba 7.Rádio, linhas 2 a 33: cópia do bloco do Hotel Araçá antes do bloco da Rádio (linha 35)."
    return rows, araca_copy


# Rows that concern these companies or their assets (names, registrations, plates).
ASSET_PAT = re.compile(
    r"R[ÁA]DIO|CEVAL|TURISMO SIVIERO|SIVIERO TRANSPORTES|ASSOCIA[ÇC][ÃA]O CULTURAL|TELMO DE TONI|"
    r"3\.751|11\.382|11\.324|IHW ?0285|IRO ?2562|IKI ?4325|IHM ?0033|ILD ?3959|IUY ?8068|IAZ ?0281|"
    r"ISS ?7019|IXO ?6000|IMY ?5000|IOM ?9093|BANDEIRANTE|NOTA FISCAL N?º? ?4946|ROV[ÍI]LIO",
    re.IGNORECASE,
)


DATA_TABS = {
    "Lotes que saíram 2018-2025",
    "Valor médio dos lotes",
    "Rafael × Inventário",
}


def gsheet_notes():
    """Observation-column texts and cell comments on rows about these companies, read
    straight from both Google sheets; flags the ones missing from the Part 7 extract."""
    extract = set()
    wbo = openpyxl.load_workbook(SRC_OBS, data_only=True)
    for r in wbo["Observacoes"].iter_rows(min_row=2, values_only=True):
        extract.add((str(r[0])[:10], r[1], r[2]))
    hits = []
    for wb_path, wb_label, key in [
        (SRC_GSHEET, "2_0", "2_0_Compar"),
        (SRC_GSHEET_RAF, "emp. e Rafael", "Comparativ"),
    ]:
        wb = openpyxl.load_workbook(
            wb_path
        )  # keep formulas off: comments need the normal load
        wbv = openpyxl.load_workbook(wb_path, data_only=True)
        for ws in wb.worksheets:
            wsv = wbv[ws.title]
            for row in wsv.iter_rows():
                txt = " ".join(str(c.value) for c in row if c.value not in (None, ""))
                if not ASSET_PAT.search(txt):
                    continue
                r = row[0].row
                if ws.title == "7.Rádio" and r < 35:
                    continue  # Hotel Araçá copy
                desc = re.sub(r"\s+", " ", str(wsv.cell(r, 2).value or ""))[:120]
                # Not every tab has a header row (GERAL of the second sheet has none): take any
                # text from column C onwards; A and B are item number and inventory description.
                for col in range(3, wsv.max_column + 1):
                    v = wsv.cell(r, col).value
                    if isinstance(v, str) and v.strip():
                        cell = f"{get_column_letter(col)}{r}"
                        hits.append(
                            {
                                "planilha": wb_label,
                                "aba": ws.title,
                                "celula": cell,
                                # Data tabs (one sold lot per row) carry row data, not notes.
                                "tipo": "dado da tabela"
                                if ws.title in DATA_TABS
                                else "texto",
                                "texto": v.strip(),
                                "linha_desc": desc,
                                "no_extrato_parte7": "sim"
                                if (key, ws.title, cell) in extract
                                else "NÃO",
                            }
                        )
                for c in ws[r]:
                    if c.comment:
                        hits.append(
                            {
                                "planilha": wb_label,
                                "aba": ws.title,
                                "celula": c.coordinate,
                                "tipo": "comentário",
                                "texto": c.comment.text.strip(),
                                "linha_desc": desc,
                                "no_extrato_parte7": "sim"
                                if (key, ws.title, c.coordinate) in extract
                                else "NÃO",
                            }
                        )
    return hits


# --- other sources: cross references with the SEI and history ----------------------------

CROSS_PAT = re.compile(
    r"R[ÁA]DIO|CEVAL|TURISMO SIVIERO|SIVIERO TRANSPORTES?\b|ASSOCIA[ÇC][ÃA]O CULTURAL|TELMO|CELITO|"
    r"3\.751|11\.382|11\.324|320\s?D\b|580\s?L|IKI ?\d{4}|IHW ?0285|BANDEIRANTE|TOYOTA(?!\s*(?:HILUX|RAV))|IXO|IMY|IOM|IUY|IAZ|ISS 7019|"
    r"MIL+ENIUM|ROV[ÍI]LIO",
    re.IGNORECASE,
)
CROSS_COMPANY = [
    (
        "07",
        r"R[ÁA]DIO|3\.751|11\.382|11\.324|TELMO|IHW ?0285|BANDEIRANTE|TOYOTA(?!\s*(?:HILUX|RAV))",
    ),
    ("04", r"CEVAL|CELITO|320\s?D\b|580\s?L|IKI ?\d{4}"),
    ("10", r"SIVIERO TRANSPORTES?\b|IXO|IMY|IOM|IUY|IAZ|ISS 7019|MIL+ENIUM|ROV[ÍI]LIO"),
    ("09", r"TURISMO|ROV[ÍI]LIO"),
    ("03", r"ASSOCIA[ÇC][ÃA]O CULTURAL"),
]


def cross_refs():
    out = []
    for label, path in [
        ("Doc 18 de 2018 (SEI)", SRC_DOC18),
        ("Doc 11 de 2025 (SEI)", SRC_DOC11),
        ("Doc 25 de 2018 (divisão de lotes)", SRC_DOC25),
        ("Doc 22 de 2018/2019 (reunião)", SRC_DOC22),
        ("Doc 23 de 2018 (dívidas 2011)", SRC_DOC23),
    ]:
        t_idx = p_idx = 0
        for kind, block in docx_blocks(path):
            if kind == "p":
                p_idx += 1
                units = [(f"parágrafo {p_idx}", block)]
            else:
                t_idx += 1
                units = [
                    (f"tabela {t_idx}, linha {n}", " | ".join(r))
                    for n, r in enumerate(block, 1)
                ]
            for where, txt in units:
                if not CROSS_PAT.search(txt):
                    continue
                txt1 = re.sub(r"\s+", " ", txt).strip()
                # Skip the long partner-list header that names every company.
                if len(re.findall(r"LTDA", txt1, re.IGNORECASE)) >= 6:
                    continue
                comps = [
                    c for c, pat in CROSS_COMPANY if re.search(pat, txt1, re.IGNORECASE)
                ]
                out.append(
                    {
                        "fonte": label,
                        "local": where,
                        "empresas": ", ".join(comps),
                        "valores": "; ".join(f"{x:,.2f}" for x in moneys(txt1)),
                        "texto": txt1[:600],
                    }
                )
    return out


def doc23_sums():
    """Running sums of the 2011 lists in doc 23, closed at each TOTAL line."""
    res = []
    run, n, pending = 0.0, 0, None
    section = "início"
    for kind, block in docx_blocks(SRC_DOC23):
        if kind != "p":
            continue
        t = re.sub(r"\s+", " ", block).strip()
        if re.match(
            r"^(Levantamento feito|D[ií]vidas|D í v i d a s|Capital|Previs[ãa]o|Financiamento|L u c i a n o|REUNI)",
            t,
        ):
            if not t.startswith("Levantamento"):
                section = t[:40]
            run, n = 0.0, 0
            continue
        vals = moneys(t)
        if t.upper().startswith("TOTAL") and not vals:
            pending = t
            continue
        if t.upper().startswith("TOTAL") or pending:
            if vals:
                res.append(
                    {
                        "secao": section,
                        "soma": round(run, 2),
                        "itens": n,
                        "total_declarado": vals[-1],
                    }
                )
            run, n, pending = 0.0, 0, None
            continue
        if vals:
            run += vals[-1]
            n += 1
    return res


NOTE_KIND = [
    ("dívida", r"D[ÍI]VIDA|dívida|parcelad|Nota Promiss"),
    ("crédito/processo", r"receber|acordo|PROCESSO|acordad|restitui"),
    ("venda de quotas (28/08/2025, após a data-base)", r"COMPRA E VENDA DE QUOTAS|rescisão do Contrato"),
    ("venda", r"vendid|(?<!E )VENDA(?! das quotas)"),
    ("redução %", r"redução de \d"),
    ("ônus da empresa", r"CUSTO PARA O DESLOCAMENTO"),
    ("intragrupo", r"subsidiad|NÃO paga|Siviero Empreendimentos Imobiliários Ltda\., ainda"),
]
# Paragraphs kept even without a money figure: dissolution, lawsuits, burdens, intra-group, quotas.
NOTE_KEEP = r"BAIXADA|R\$\s*\.{3,}|PROCESSO JUDICIAL|CUSTO PARA O DESLOCAMENTO|subsidiad|NÃO paga|COMPRA E VENDA DE QUOTAS|rescisão do Contrato"


def note_values(inv):
    """Money figures and percentages that sit in the free text of each inventory."""
    out = []
    for code in COMPANIES:
        for key in ("2018", "2025", "2025a"):
            d = inv[code][key]
            ver = {"2018": "2018", "2025": "2025 atual", "2025a": "2025 antigo"}[key]
            for n, t in enumerate(d["notes"], 1):
                vals = moneys(t)
                # A 0,00 % reduction is boilerplate; keep only non-zero ones.
                pct = [
                    x for x in re.findall(r"redução de (\d+,\d+)%", t) if brl(x) != 0
                ]
                if (
                    not vals
                    and not pct
                    and not re.search(NOTE_KEEP, t)
                ):
                    continue
                kinds = [k for k, pat in NOTE_KIND if re.search(pat, t)]
                out.append(
                    {
                        "empresa": code,
                        "versao": ver,
                        "paragrafo": n,
                        "valores": "; ".join(f"{x:,.2f}" for x in vals)
                        + (f" | {pct[0]}%" if pct else ""),
                        "natureza": ", ".join(kinds)
                        or ("baixa" if "BAIXADA" in t else ""),
                        "texto": re.sub(r"\s+", " ", t)[:500],
                    }
                )
    return out


def doc00_quotas(path: Path):
    """Per-partner quota of each company in the third table of doc 00 (current version)."""
    out = {}
    for kind, block in docx_blocks(path):
        if kind != "t" or "E M P R E S A" not in " ".join(block[0]):
            continue
        for row in block[1:]:
            cell = re.sub(r"\s+", " ", row[0])
            if "COTA" not in cell.upper():
                continue
            for code, label in [
                ("03", "ASSOCIA"),
                ("04", "CEVAL"),
                ("07", "RÁDIO"),
                ("09", "TURISMO"),
                ("10", "SIVIERO TRANSPORTES"),
            ]:
                if label in cell.upper():
                    vals = moneys(cell)
                    out[code] = {
                        "socio": vals[0] if vals else None,
                        "telmo": vals[1] if len(vals) > 1 else None,
                    }
    return out


def valter_list():
    wb = openpyxl.load_workbook(SRC_VALTER, data_only=True)
    ws = wb["Ativos"]
    out = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        txt = " | ".join(str(v) for v in row if v not in (None, ""))
        # Rafael's numbering differs ("10 -" is the SEI there): match by company name only.
        if re.search(
            r"Rádio|Ceval|Turismo|Siviero Transportes|Associação Cultural", txt
        ):
            out.append(txt)
    return out


def norm(t: str) -> str:
    return re.sub(r"[^0-9A-Z]", "", t.upper())


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    sources = []

    def add_src(role, path, company=""):
        info = {
            "papel": role,
            "empresa": company,
            "caminho": str(path).replace("\\\\?\\", ""),
            "md5": md5(path),
            "bytes": path.stat().st_size,
            "paginas": "",
        }
        if path.suffix.lower() == ".pdf":
            d = pymupdf.open(path)
            info["paginas"] = len(d)
            info["texto"] = (
                "sim" if any(p.get_text().strip() for p in d) else "não (escaneado)"
            )
        sources.append(info)

    inv = {}
    for code in COMPANIES:
        add_src("Inventário 2018 (Drive)", SRC_2018[code], code)
        add_src("Inventário 2025 atual", SRC_2025[code], code)
        add_src("Inventário 2025 antigo", SRC_2025_OLD[code], code)
        add_src("Estudo 2 Rafael", SRC_RAFAEL[code], code)
        inv[code] = {
            "2018": parse_inventory(SRC_2018[code], code, "2018"),
            "2025": parse_inventory(SRC_2025[code], code, "2025 atual"),
            "2025a": parse_inventory(SRC_2025_OLD[code], code, "2025 antigo"),
            "raf": parse_rafael(SRC_RAFAEL[code], code),
        }
    for role, p in [
        ("Doc 27 de 2018 (resumo)", SRC_DOC27),
        ("Doc 00 de 2025 (resumo)", SRC_DOC00),
        ("Doc 00 de 2025 antigo", SRC_DOC00_OLD),
        ("Doc 23 de 2018 (dívidas TS+ST, 2011)", SRC_DOC23),
        ("Doc 18 de 2018 (SEI): cruzamentos", SRC_DOC18),
        ("Doc 25 de 2018 (divisão de lotes): histórico", SRC_DOC25),
        ("Doc 22 de 2018/2019 (reunião de sócios): histórico", SRC_DOC22),
        ("Doc 11 de 2025 (SEI): cruzamentos", SRC_DOC11),
        ("Rafael: Diretrizes", SRC_DIRETRIZES),
        ("Rafael: Introdução", SRC_INTRO),
        ("Lista Valter (xlsx)", SRC_VALTER),
        ("Rafael em xlsx (repete a lista Valter)", SRC_RAF_XLSX),
        ("Planilha Google 2_0", SRC_GSHEET),
        ("Planilha Google empreendimentos e Rafael", SRC_GSHEET_RAF),
        ("Observações das planilhas Google (extrato da Parte 7)", SRC_OBS),
        (
            "ACERTOS E ACORDOS (escaneado; abas das planilhas não citam estas empresas)",
            SRC_ACERTOS,
        ),
    ]:
        add_src(role, p)
    for p in SRC_OTHER:
        add_src("Apoio (etapas seguintes)", p)

    doc27 = summary_table(SRC_DOC27)
    doc00 = summary_table(SRC_DOC00)
    doc00a = summary_table(SRC_DOC00_OLD)
    quotas = doc00_quotas(SRC_DOC00)
    graw, araca_copy = gsheet_rows()
    gnotes = gsheet_notes()
    dhits = diretrizes_hits()
    vlist = valter_list()
    xrefs = cross_refs()
    d23 = doc23_sums()
    nvals = note_values(inv)

    # --- checks -------------------------------------------------------------------------
    checks = []

    def chk(code, what, extracted, declared, src, note="", status=None):
        ok = (
            declared is not None
            and extracted is not None
            and abs(extracted - declared) < 0.005
        )
        checks.append(
            {
                "empresa": code,
                "conferencia": what,
                "soma_extraida": extracted,
                "total_declarado": declared,
                "confere": status or ("sim" if ok else "NÃO"),
                "fonte": src,
                "obs": note,
            }
        )

    def item_by_key(code, key, pattern):
        for i in inv[code][key]["items"]:
            if re.search(pattern, i["chaves"] + " " + i["descricao"], re.IGNORECASE):
                return i
        return None

    for code in COMPANIES:
        for key, label in [
            ("2018", "2018"),
            ("2025", "2025 atual"),
            ("2025a", "2025 antigo"),
        ]:
            d = inv[code][key]
            vals = [i["valor"] for i in d["items"] if i["valor"] is not None]
            s = sum(vals)
            if key == "2018" and code == "10":
                sold = sum(i["valor"] for i in d["items"] if i["situacao_no_doc"])
                chk(
                    code,
                    f"{label}: soma bruta dos itens",
                    s,
                    d["declared_gross"],
                    "tabela de bens",
                    "total bruto na célula de total",
                )
                chk(
                    code,
                    f"{label}: soma sem os ônibus vendidos a prazo",
                    s - sold,
                    d["declared"],
                    "tabela de bens",
                    f"vendidos a prazo: {sold:,.2f}",
                )
                continue
            note = ""
            if d["declared"] is None:
                note = "documento sem total preenchido"
            elif not vals and d["declared"]:
                note = "total declarado sem nenhum item com valor"
            chk(
                code,
                f"{label}: soma dos itens",
                s,
                d["declared"],
                "tabela de bens",
                note,
            )
            # partner split table must add up to the total (when it has values)
            parts = [p["valor"] for p in d["partners"][:-1] if p["valor"] is not None]
            if parts:
                chk(
                    code,
                    f"{label}: tabela 'valor real de cada sócio' soma o total",
                    sum(parts),
                    d["declared"],
                    "tabela de sócios",
                )
        chk(
            code,
            "2018: total do inventário × doc 27",
            inv[code]["2018"]["declared"],
            doc27.get(code),
            "doc 27",
            "" if inv[code]["2018"]["declared"] is not None else "inventário sem total",
        )
        chk(
            code,
            "2025: total do inventário × doc 00",
            inv[code]["2025"]["declared"],
            doc00.get(code),
            "doc 00",
        )
        chk(
            code,
            "2025 antigo: total × doc 00 antigo",
            inv[code]["2025a"]["declared"],
            doc00a.get(code),
            "doc 00 antigo",
        )
        if code in quotas:
            first = (
                inv[code]["2025"]["partners"][0]["valor"]
                if inv[code]["2025"]["partners"]
                else None
            )
            chk(
                code,
                "Doc 00, tabela 3: cota por sócio × inventário 2025",
                quotas[code]["socio"],
                first,
                "doc 00 × tabela de sócios",
            )
            if quotas[code]["telmo"] is not None:
                telmo = next(
                    (
                        p["valor"]
                        for p in inv[code]["2025"]["partners"]
                        if "TELMO" in p["socio"].upper()
                    ),
                    None,
                )
                chk(
                    code,
                    "Doc 00, tabela 3: cota de Telmo de Toni × inventário 2025",
                    quotas[code]["telmo"],
                    telmo,
                    "doc 00 × tabela de sócios",
                )
        # Rafael
        raf = inv[code]["raf"]
        for sec, label in [
            ("apresentado_grupo", "Rafael: balanço apresentado pelo grupo"),
            ("balanco_valter", "Rafael: balanço executado por Valter"),
        ]:
            its = [i for i in raf["items"] if i["secao"] == sec]
            tot = [t for t in raf["totals"] if t["secao"] == sec]
            if its or tot:
                chk(
                    code,
                    f"{label}: soma dos itens",
                    sum(i["valor"] or 0 for i in its),
                    tot[0]["valor"] if tot else None,
                    f"PDF p. {tot[0]['pagina'] if tot else '?'}",
                )
        g = [r for r in raf["items"] if r["secao"] == "apresentado_grupo"]
        inv_items = inv[code]["2025"]["items"]
        if g:
            same_n = len(g) == len(inv_items)
            same_v = same_n and all(
                abs((a["valor"] or 0) - (b["valor"] or 0)) < 0.005
                for a, b in zip(g, inv_items)
            )
            chk(
                code,
                "Rafael: transcrição do grupo — nº de itens e valores iguais ao inventário 2025",
                None,
                None,
                "PDF × docx",
                f"{len(g)} itens no PDF, {len(inv_items)} no inventário",
                status="sim" if same_v else "NÃO",
            )
        # Paragraphs with money, the dissolution date or Rovílio: present on both sides?
        ptxt = norm(raf["texto"])
        miss, extra = [], []
        for t in inv[code]["2025"]["notes"]:
            if (
                re.search(r"R\$\s*[\d]|BAIXADA|ROV[ÍI]LIO", t)
                and norm(t)[:60] not in ptxt
            ):
                miss.append(re.sub(r"\s+", " ", t)[:120])
        inv_txt = norm(" ".join(inv[code]["2025"]["notes"]))
        # norm() drops accented letters, so "ROVÍLIO" becomes "ROVLIO" on both sides.
        if "ROVLIO" in ptxt and "ROVLIO" not in inv_txt:
            extra.append("parágrafo da venda das quotas de Rovílio")
        chk(
            code,
            "Rafael: notas do inventário 2025 (com valor, baixa ou Rovílio) transcritas no PDF",
            None,
            None,
            "PDF × docx",
            ("faltam no PDF: " + " / ".join(miss) if miss else "")
            + (" | só no PDF: " + "; ".join(extra) if extra else ""),
            status="sim" if not miss and not extra else "NÃO",
        )
        # Rafael's summary table
        qv = [v for v, _ in raf["quadro_vals"]]
        tot_v = next(
            (t["valor"] for t in raf["totals"] if t["secao"] == "balanco_valter"), None
        )
        if qv and tot_v is not None:
            pg = raf["quadro_vals"][0][1]
            chk(
                code,
                "Rafael quadro-resumo: patrimônio total = total do balanço de Valter",
                qv[0],
                tot_v,
                f"PDF p. {pg}",
            )
            chk(
                code,
                "Rafael quadro-resumo: 1/6 = patrimônio / 6",
                qv[1],
                round(qv[0] / 6, 2),
                f"PDF p. {pg}",
                "PDF trunca os centavos"
                if abs(qv[1] - round(qv[0] / 6, 2)) < 1 and qv[1] != round(qv[0] / 6, 2)
                else "Rafael usa 1/6; o inventário dá 16,50 % por sócio + 1 % a Telmo de Toni"
                if code == "07"
                else "",
            )
            chosen = sum(
                i["valor"] or 0 for i in raf["items"] if i["secao"] == "bens_escolhidos"
            )
            if len(qv) >= 4:
                chk(
                    code,
                    "Rafael quadro-resumo: selecionados = bens escolhidos",
                    qv[2],
                    chosen,
                    f"PDF p. {pg}",
                )
                chk(
                    code,
                    "Rafael quadro-resumo: crédito = 1/6 − selecionados",
                    qv[3],
                    round(qv[1] - qv[2], 2),
                    f"PDF p. {pg}",
                )
            elif len(qv) == 3:
                chk(
                    code,
                    "Rafael quadro-resumo: crédito = 1/6 (nenhum bem selecionado)",
                    qv[2],
                    qv[1],
                    f"PDF p. {pg}",
                )
            if raf["fracoes_bens"]:
                checks.append(
                    {
                        "empresa": code,
                        "conferencia": "Rafael quadro-resumo: frações de 'bens' citadas",
                        "soma_extraida": None,
                        "total_declarado": None,
                        "confere": "NÃO"
                        if len({f.split()[0] for f in raf["fracoes_bens"]}) > 1
                        else "sim",
                        "fonte": f"PDF p. {pg}",
                        "obs": "; ".join(raf["fracoes_bens"]),
                    }
                )
        # Old × current: items, then partner tables and notes separately
        a = [(i["item"], i["valor"]) for i in inv[code]["2025a"]["items"]]
        b = [(i["item"], i["valor"]) for i in inv[code]["2025"]["items"]]
        chk(
            code,
            "2025 antigo × atual: itens e valores iguais",
            None,
            None,
            "docx × docx",
            status="sim" if a == b else "NÃO",
        )
        pa = [
            (p["socio"], p["pct"], p["valor"]) for p in inv[code]["2025a"]["partners"]
        ]
        pb = [(p["socio"], p["pct"], p["valor"]) for p in inv[code]["2025"]["partners"]]
        na = {norm(t) for t in inv[code]["2025a"]["notes"]}
        nb = {norm(t) for t in inv[code]["2025"]["notes"]}
        diffs = []
        if pa != pb:
            diffs.append(
                f"tabela de sócios: antiga {len(pa)} linhas "
                f"({', '.join(sorted({x[1] for x in pa if x[1]}))}), atual {len(pb)} linhas "
                f"({', '.join(sorted({x[1] for x in pb if x[1]}))})"
            )
        only_new = [
            t for t in inv[code]["2025"]["notes"] if norm(t) not in na and len(t) > 30
        ]
        only_old = [
            t for t in inv[code]["2025a"]["notes"] if norm(t) not in nb and len(t) > 30
        ]
        if only_new:
            diffs.append(
                "só na atual: "
                + " / ".join(re.sub(r"\s+", " ", t)[:90] for t in only_new)
            )
        if only_old:
            diffs.append(
                "só na antiga: "
                + " / ".join(re.sub(r"\s+", " ", t)[:90] for t in only_old)
            )
        chk(
            code,
            "2025 antigo × atual: tabela de sócios e notas",
            None,
            None,
            "docx × docx",
            "; ".join(diffs),
            status="sim" if not diffs else "difere",
        )

    # Rafael's chosen asset × Valter list × inventory (C0 list)
    toy_raf = [i for i in inv["07"]["raf"]["items"] if i["secao"] == "bens_escolhidos"]
    toy_inv = item_by_key("07", "2025", r"IHW0285")
    vl_vals = [
        float(re.search(r"\| (\d+(?:\.\d+)?) \| 1 \|", v).group(1))
        for v in vlist
        if re.search(r"\| (\d+(?:\.\d+)?) \| 1 \|", v)
    ]
    if toy_raf and toy_inv:
        chk(
            "07",
            "Bem escolhido por Rafael (p. 11) × inventário 2025 item 07",
            toy_raf[0]["valor"],
            toy_inv["valor"],
            f"PDF p. {toy_raf[0]['pagina_valor']}",
            "rótulo do quadro p. 10: 'Conjunto de lotes escolhidos (07.)' para um veículo",
        )
    if vl_vals and toy_inv:
        chk(
            "07",
            "Lista Valter (xlsx, nº 10 da lista) × inventário 2025 item 07",
            sum(vl_vals),
            toy_inv["valor"],
            "ativos-avaliados-valter.xlsx",
            f"{len(vlist)} linha(s) destas 5 empresas na lista",
        )

    # Google sheet: headers vs inventories, with reconciliation of 2018 figures
    terr = item_by_key("07", "2018", r"mat\. 3\.751")
    toy18 = item_by_key("10", "2018", r"IHW0285")
    for r in graw:
        if (
            r["planilha"] != "2_0"
            or r["aba"] != "GERAL"
            or r["linha"] != GERAL_ROWS[r["empresa"]][0]
        ):
            continue
        code = r["empresa"]
        chk(
            code,
            "Planilha 2_0 GERAL: coluna 'outubro 2025' × inventário 2025",
            r["out2025"],
            inv[code]["2025"]["declared"],
            f"GERAL linha {r['linha']}",
        )
        expected, note = doc27.get(code), ""
        if code == "07" and terr and toy18:
            mov = next(
                (
                    x
                    for x in graw
                    if x["planilha"] == "2_0"
                    and x["aba"] == "GERAL"
                    and x["empresa"] == "07"
                    and str(x["descricao"]).startswith("Móveis")
                ),
                None,
            )
            mv = mov["2018_apresentado"] if mov else 0
            expected = doc27["07"] - terr["valor"] + toy18["valor"] + (mv or 0)
            note = (
                f"reconciliação: doc 27 {doc27['07']:,.2f} − terreno mat. 3.751 {terr['valor']:,.2f} (foi para a SEI) "
                f"+ Toyota {toy18['valor']:,.2f} (veio da Siviero Transportes) + móveis {mv:,.2f} "
                f"(GERAL linha {mov['linha'] if mov else '?'}): item que NÃO existe no inventário 2018 "
                "(os móveis estavam dentro do item 02, R$ 1.700.000)"
            )
        elif code == "10" and toy18:
            expected = doc27["10"] - toy18["valor"]
            note = f"reconciliação: doc 27 {doc27['10']:,.2f} − Toyota {toy18['valor']:,.2f} (passada à Rádio)"
        elif code == "03":
            note = "a planilha zerou o total de 2018 (R$ 10.000 no inventário, sem nenhum item)"
        chk(
            code,
            "Planilha 2_0 GERAL: '2018 apresentado' × doc 27 (reconciliado)",
            r["2018_apresentado"],
            expected,
            f"GERAL linha {r['linha']}",
            note,
        )
    for r in graw:
        if r["aba"] in GSHEET_TABS.values() and str(r["descricao"]).upper().startswith(
            ("RÁDIO", "CEVAL", "SIVIERO TRANSPORTES")
        ):
            code = r["empresa"]
            chk(
                code,
                f"Planilha {r['planilha']}, aba {r['aba']}: 'outubro 2025' × inventário 2025",
                r["out2025"],
                inv[code]["2025"]["declared"],
                f"linha {r['linha']}",
                f"coluna 'Estudo Rafael' = {r['estudo_rafael']:,.2f} (só o bem escolhido, não o balanço de Valter)"
                if isinstance(r["estudo_rafael"], (int, float))
                else "",
            )
    for s in d23:
        chk(
            "09/10",
            f"Doc 23 (2011): soma da lista '{s['secao']}'",
            s["soma"],
            s["total_declarado"],
            "doc 23",
            f"{s['itens']} linhas",
        )

    # --- workbook -----------------------------------------------------------------------
    wb = openpyxl.Workbook()
    bold = Font(bold=True)
    hdr_fill = PatternFill("solid", fgColor="DDEBF7")
    bad_fill = PatternFill("solid", fgColor="F8CBAD")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")

    def sheet(title, header, rows, widths=None):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font = bold
            c.fill = hdr_fill
        for r in rows:
            ws.append(r)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, w in enumerate(widths or [], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if isinstance(c.value, float):
                    c.number_format = "#,##0.00"
        return ws

    wb.remove(wb.active)
    sheet(
        "Fontes",
        ["Papel", "Empresa", "Caminho", "MD5", "Bytes", "Páginas", "Tem texto?"],
        [
            [
                s["papel"],
                s["empresa"],
                s["caminho"],
                s["md5"],
                s["bytes"],
                s["paginas"],
                s.get("texto", ""),
            ]
            for s in sources
        ],
        [34, 8, 90, 34, 10, 8, 14],
    )
    ws = sheet(
        "Conferências",
        [
            "Empresa",
            "Conferência",
            "Valor extraído",
            "Valor de referência",
            "Confere?",
            "Fonte",
            "Obs.",
        ],
        [
            [
                c["empresa"],
                c["conferencia"],
                c["soma_extraida"],
                c["total_declarado"],
                c["confere"],
                c["fonte"],
                c["obs"],
            ]
            for c in checks
        ],
        [8, 60, 16, 16, 9, 24, 70],
    )
    for row in ws.iter_rows(min_row=2):
        fill = (
            bad_fill
            if row[4].value == "NÃO"
            else warn_fill
            if row[4].value == "difere"
            else None
        )
        if fill:
            for c in row:
                c.fill = fill
    rows = []
    for code in COMPANIES:
        for key in ("2018", "2025", "2025a"):
            for i in inv[code][key]["items"]:
                rows.append(
                    [
                        code,
                        i["versao"],
                        i["tabela"],
                        i["item"],
                        i["descricao"],
                        i["valor"],
                        i["valor_descricao"],
                        i["alerta"],
                        i["situacao_no_doc"],
                        "sim" if i["vazio"] else "",
                        i["chaves"],
                    ]
                )
    sheet(
        "Itens inventários",
        [
            "Empresa",
            "Versão",
            "Tabela",
            "Item",
            "Descrição",
            "Valor (coluna)",
            "Valor citado na descrição",
            "Alerta",
            "Situação no doc",
            "Linha vazia?",
            "Chaves (matrícula, placa, chassi, NF)",
        ],
        rows,
        [8, 12, 7, 6, 80, 14, 16, 24, 26, 8, 36],
    )
    rows = []
    for code in COMPANIES:
        for i in inv[code]["raf"]["items"]:
            rows.append(
                [
                    code,
                    i["secao"],
                    i["item"],
                    i["pagina"],
                    i["descricao"][:400],
                    i["valor"],
                    i["pagina_valor"],
                    i["criterio"],
                    i["faixa_min"],
                    i["faixa_max"],
                    i["pagina_faixa"],
                    i["texto_faixa"],
                    i["outras_referencias"],
                    i["chaves"],
                ]
            )
    sheet(
        "Rafael itens",
        [
            "Empresa",
            "Seção do PDF",
            "Item",
            "Página",
            "Descrição",
            "Valor",
            "Pág. valor",
            "Critério de Rafael",
            "Faixa mín.",
            "Faixa máx.",
            "Pág. faixa",
            "Texto da faixa",
            "Todas as referências R$ do bloco (pág., valor, contexto)",
            "Chaves",
        ],
        rows,
        [8, 18, 5, 6, 50, 14, 7, 22, 12, 12, 7, 45, 80, 26],
    )
    rows = []
    for code in COMPANIES:
        r = inv[code]["raf"]
        for t in r["totals"]:
            rows.append([code, "total", t["secao"], t["pagina"], t["valor"], ""])
        for v, p in r["quadro_vals"]:
            rows.append([code, "quadro-resumo: valor", "quadro_resumo", p, v, ""])
        for q in r["quadro"]:
            rows.append(
                [code, "quadro/escolhidos: linha", "", q["pagina"], None, q["linha"]]
            )
        for f in r["fracoes_bens"]:
            rows.append(
                [code, "quadro-resumo: fração de bens", "quadro_resumo", "", None, f]
            )
        for s, pages in r["secoes"].items():
            rows.append(
                [code, "seção presente", s, ", ".join(map(str, pages)), None, ""]
            )
    for h in dhits:
        rows.append(["", f"{h['fonte']}: menção", "", h["pagina"], None, h["contexto"]])
    for v in vlist:
        rows.append(["", "lista Valter (xlsx)", "", "", None, "nº " + v])
    sheet(
        "Rafael totais e seções",
        ["Empresa", "Tipo", "Seção", "Página(s)", "Valor", "Texto"],
        rows,
        [8, 26, 20, 12, 16, 100],
    )
    rows = []
    for code in COMPANIES:
        for key in ("2018", "2025", "2025a"):
            d = inv[code][key]
            for p in d["partners"]:
                rows.append(
                    [
                        code,
                        p["versao"],
                        "valor real de cada sócio",
                        p["socio"],
                        p["pct"],
                        p["valor"],
                    ]
                )
            for q in d["quadros"]:
                for ln in q["linhas"]:
                    rows.append(
                        [
                            code,
                            q["versao"],
                            f"quadro societário (tabela {q['tabela']})",
                            ln,
                            "",
                            None,
                        ]
                    )
    sheet(
        "Sócios",
        ["Empresa", "Versão", "Tabela", "Sócio / linha", "%", "Valor"],
        rows,
        [8, 12, 30, 80, 8, 14],
    )
    sheet(
        "Valores em notas",
        [
            "Empresa",
            "Versão",
            "Parágrafo nº",
            "Valores (R$ | %)",
            "Natureza (classificação automática)",
            "Texto",
        ],
        [
            [
                n["empresa"],
                n["versao"],
                n["paragrafo"],
                n["valores"],
                n["natureza"],
                n["texto"],
            ]
            for n in nvals
        ],
        [8, 12, 8, 30, 22, 120],
    )
    rows = []
    for code in COMPANIES:
        for key in ("2018", "2025", "2025a"):
            ver = {"2018": "2018", "2025": "2025 atual", "2025a": "2025 antigo"}[key]
            for n, t in enumerate(inv[code][key]["notes"], 1):
                if len(t) > 40 and not re.match(
                    r"^(Empresa|INVENTÁRIO|CNPJ|Fone|Avenida|Rua|Edifício|Veranópolis)",
                    t,
                ):
                    rows.append([code, ver, n, t])
    sheet(
        "Textos dos inventários",
        ["Empresa", "Versão", "Parágrafo nº", "Texto"],
        rows,
        [8, 12, 8, 140],
    )
    sheet(
        "Outras fontes (cruzamentos)",
        ["Fonte", "Local", "Empresas (das 5)", "Valores R$", "Texto"],
        [
            [x["fonte"], x["local"], x["empresas"], x["valores"], x["texto"]]
            for x in xrefs
        ],
        [26, 18, 12, 30, 120],
    )
    sheet(
        "Planilhas Google",
        [
            "Planilha",
            "Empresa",
            "Aba",
            "Linha",
            "Item",
            "Descrição",
            "2018 apresentado",
            "2018 atualizado",
            "dezembro 2024",
            "auxiliar (sem rótulo)",
            "outubro 2025",
            "Estudo Rafael",
            "Divergência",
            "Diferença",
            "Observação",
        ],
        [
            [
                r["planilha"],
                r["empresa"],
                r["aba"],
                r["linha"],
                r["item"],
                r["descricao"],
            ]
            + [r[k] for k in GS_KEYS]
            for r in graw
        ],
        [12, 8, 14, 6, 6, 60, 14, 14, 14, 14, 14, 14, 12, 12, 30],
    )
    sheet(
        "Observações Google",
        [
            "Planilha",
            "Aba",
            "Célula",
            "Tipo",
            "Texto",
            "Linha (descrição)",
            "Está no extrato da Parte 7?",
        ],
        [
            [
                h["planilha"],
                h["aba"],
                h["celula"],
                h["tipo"],
                h["texto"],
                h["linha_desc"],
                h["no_extrato_parte7"],
            ]
            for h in gnotes
        ]
        or [["(nenhuma)", "", "", "", "", "", ""]],
        [14, 18, 8, 10, 70, 50, 12],
    )
    sheet(
        "Resumos 27 e 00",
        [
            "Empresa",
            "Nome",
            "Doc 27 (2018)",
            "Doc 00 (2025)",
            "Doc 00 antigo",
            "Doc 00: cota por sócio",
            "Doc 00: cota Telmo de Toni",
        ],
        [
            [
                k,
                COMPANIES[k],
                doc27.get(k),
                doc00.get(k),
                doc00a.get(k),
                quotas.get(k, {}).get("socio"),
                quotas.get(k, {}).get("telmo"),
            ]
            for k in COMPANIES
        ],
        [8, 40, 16, 16, 16, 16, 16],
    )

    out_xlsx = OUT_DIR / "fase0-extracao.xlsx"
    wb.save(out_xlsx)
    for code in inv:
        inv[code]["raf"].pop("texto", None)
    payload = {
        "sources": sources,
        "inventories": inv,
        "doc27": doc27,
        "doc00": doc00,
        "doc00_old": doc00a,
        "doc00_quotas": quotas,
        "gsheet": graw,
        "gsheet_notes": gnotes,
        "gsheet_radio_tab_note": araca_copy,
        "diretrizes_hits": dhits,
        "valter_list": vlist,
        "cross_refs": xrefs,
        "doc23_sums": d23,
        "note_values": nvals,
        "checks": checks,
    }
    (OUT_DIR / "fase0-extracao.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=1, default=str), encoding="utf-8"
    )
    bad = [c for c in checks if c["confere"] != "sim"]
    print(f"sources={len(sources)} checks={len(checks)} not_sim={len(bad)}")
    for c in bad:
        print(
            f"  {c['confere']}: {c['empresa']} {c['conferencia']} | {c['soma_extraida']} × {c['total_declarado']} | {c['obs'][:220]}"
        )
    print(
        "gsheet notes:",
        [
            (
                h["planilha"],
                h["aba"],
                h["celula"],
                h["texto"][:60],
                h["no_extrato_parte7"],
            )
            for h in gnotes
        ],
    )
    print(
        "cross refs:",
        len(xrefs),
        "| doc23:",
        d23,
        "| note values:",
        len(nvals),
        "| quotas:",
        quotas,
    )
    print("wrote", out_xlsx)


if __name__ == "__main__":
    main()
