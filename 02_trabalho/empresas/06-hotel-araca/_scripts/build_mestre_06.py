"""Master sheet for company 06 (Hotel Araçá): planilha-mestre-06.xlsx.

One row per asset: 2025 items, 2018 items with no 2025 match, assets cited
only by Rafael and assets found in other sources but valued by nobody.
Columns follow section 0.7 of 00_PLANO.md, plus Rafael's list (C0), the C0
direction and Rafael's market range. 2018 x 2025 pairs come only from the
normalised municipal registration (inscrição imobiliária), cross-checked by
matrícula and property code when both years show one (method lesson 3).
Item 01 is broken into components for both sides (lesson 2); the decision is
taken at item level, never by component.

Decisions are never typed here: they are read from ../decisoes.json and
applied on every rebuild. Every number is read from the sources through the
parsers of fase0_extracao_06.py. Notes are keyed by inscrição, not by item
number.

Usage: python build_mestre_06.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fase0_extracao_06 as f0

HERE = Path(__file__).resolve().parents[1]
OUT = HERE / "planilha-mestre-06.xlsx"
DECISOES = HERE / "decisoes.json"
EMPRESA = "06 - Hotel Araçá Ltda."
TOL = 0.005
DRIVE18 = f0.DRIVE / "2018/inventário-2018"
EXTRA_SRC = {
    "inv11_2025": f0.BASE
    / "01_fontes/inventario-2025/11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx",
    "doc25_2018": next(DRIVE18.glob("25. 1. PLANILHA = DIVISÃO DE LOTES*.docx")),
    "doc6_2018": next(DRIVE18.glob("6. 1. VALTER SIVIERO*.docx")),
    "doc1_2018": next(DRIVE18.glob("1. 1. RAUL SIVIERO*.docx")),
    "doc3_2018": next(DRIVE18.glob("3. 1. IRINEU SIVIERO*.docx")),
    "doc4_2018": next(DRIVE18.glob("4. 1. PEDRO SIVIERO*.docx")),
    "doc5_2018": next(DRIVE18.glob("5. 1. JOSE VAL*RIO SIVIERO*.docx")),
    "draft_2018": next(
        (f0.DRIVE / "Material antigo/2018_desatualizado").glob("4.HOTEL ARA*.docx")
    ),
    "doc22_trechos": HERE / "_extracoes/doc22-2018-trechos.txt",
    "reuniao_2017": next(
        (f0.DRIVE / "Material antigo/2018_desatualizado").glob("G-SÓCIOS*.docx")
    ),
    "arq18_pdf": next(
        (f0.BASE / "01_fontes/respostas-docs-2026-08-13").glob("18 - *.pdf")
    ),
}

# notes keyed by normalised inscrição (2025 items) -------------------------
INSC_ITEM20 = "645.0027.015.000"
INSC_LOT28_Q38 = "360.0038.028.000"
INSC_LOT27_Q38 = "360.0038.027.000"

SITUACOES = [
    ("Concorda", "mesmo valor no inventário 2025 e no balanço executado de Rafael"),
    ("Discorda", "valores diferentes nas duas fontes"),
    (
        "Só 2018",
        "extensão do plano: bem do inventário 2018 sem par em 2025 (destino a documentar, P4)",
    ),
    ("Só Rafael", "citado só por Rafael, fora do inventário"),
    ("Não abordado", "achado em outra fonte; nem o grupo nem Rafael avaliam"),
]


def br(v: float) -> str:
    """Brazilian number format: 6.900,00."""
    return f"{v:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")


def matriculas_item01(desc: str) -> list[str]:
    seg = desc.split("Matrículas nº.:")[1].split(", onde")[0]
    return [m.strip() for m in re.split(r",| e ", seg) if m.strip()]


def bloco(it) -> str:
    if it["item"] == "01":
        return "Hotel (quadra 37)"
    if it["inscricao_norm"] == INSC_ITEM20:
        return "Condado de Capão"
    return "Lotes urbanos próximos ao hotel"


def chave(it, mats01=None) -> str:
    if it["item"] == "01":
        return f"HOTEL-Q37 ({len(mats01)} matrículas: {', '.join(mats01)})"
    mat = it["matricula"] or "sem matrícula no texto"
    return f"insc. {it['inscricao_norm']} | mat. {mat} | cód. {it['codigo_imovel']}"


def desc_curta(it) -> str:
    d = it["descricao"]
    if it["item"] == "01":
        return (
            "Hotel: 18 terrenos (6.900 m²) com prédio de 4 andares, 144 apartamentos"
            + (" (2018)" if it["fonte"] == "2018" else "")
        )
    tag = re.search(r"Hotel Araçá Ltda\.\s*(?:–\s*)?(.*)$", d)
    extra = tag.group(1).strip(" .") if tag else ""
    extra = extra.split(" / ")[0].strip(" .")
    parts = [
        x
        for x in [
            it["lote"] and f"Lote {it['lote']}",
            it["quadra"] and f"q{it['quadra']}",
            it["area_m2"] and f"{it['area_m2']} m²",
        ]
        if x
    ]
    if "CONDADO" in d:
        parts.append("Condado de Capão, casa 158,18 m²")
    elif extra:
        parts.append(extra)
    if it["nome_socio_no_valor"]:
        parts.append(f"no lugar do valor: {it['nome_socio_no_valor']}")
    return ", ".join(parts)


def text_of(path: Path) -> str:
    out = []
    for kind, b in f0.doc_blocks(path):
        if kind == "p":
            out.append(b)
        else:
            for r in b.rows:
                out.append(" || ".join(c.text.strip() for c in r.cells))
    return re.sub(r"[ \t]+", " ", "\n".join(out))


def partner_capao_lines(path: Path) -> list[str]:
    """IPTU lists of 2018 (docs 1, 3–6): lines located in Capão da Canoa."""
    lines = []
    for _, cells in f0.table_rows(path):
        if any("Capão da Canoa" in c for c in cells):
            lines.append(
                " | ".join(dict.fromkeys(c.replace("\n", " ") for c in cells if c))
            )
    return lines


def main():
    dec = json.loads(DECISOES.read_text(encoding="utf-8"))
    inv = {}
    for key, label in [
        ("inv2018", "2018"),
        ("inv2025", "2025 atual"),
        ("inv2025_antigo", "2025 antigo"),
    ]:
        items, total = f0.parse_inventory(f0.SRC[key], label)
        assert abs(sum(i["valor"] or 0 for i in items) - total) < TOL
        inv[label] = items
    i18, i25, i25a = inv["2018"], inv["2025 atual"], inv["2025 antigo"]
    n_valued_18 = sum(1 for i in i18 if i["valor"])

    pages = f0.pdf_pages(f0.SRC["rafael_pdf"])
    r01, r01p, rlots, rtotal = f0.balance_section(pages, 16, 20)
    raf = {f0.norm_inscricao(r["inscricao"]): r for r in rlots}
    ranges = f0.rafael_ranges(pages)

    def rng(bem):
        hits = [r for r in ranges if r[1] == bem and r[2] == "faixa"]
        return hits[0] if hits else None

    lista = {}
    for r in openpyxl.load_workbook(f0.SRC["lista_rafael"], data_only=True)[
        "Ativos"
    ].iter_rows(min_row=2, values_only=True):
        if r[1] and str(r[1]).startswith("06 -"):
            insc = re.search(r"Inscrição ([\d\.]+\d)", str(r[14])).group(1)
            lista[f0.norm_inscricao(insc)] = {
                "item": str(r[3]).zfill(2),
                "valor": r[10],
                "matricula": str(r[8]).split(" ")[0],
                "pag": r[13],
            }

    # item 01: the same 12 matrículas in every version
    mats01 = matriculas_item01(i25[0]["descricao"])
    assert len(mats01) == 12
    for it in (i18[0], i25a[0]):
        assert matriculas_item01(it["descricao"]) == mats01

    # 2018 x 2025 by normalised inscrição; matrícula and code must agree when both exist
    by18 = {}
    for it in i18[1:]:
        k = it["inscricao_norm"]
        assert k and k not in by18, k
        by18[k] = it
    pairs, used18 = {"01": i18[0]}, {"01"}
    for it in i25[1:]:
        p = by18.get(it["inscricao_norm"])
        if p:
            if p["matricula"]:
                assert p["matricula"] == it["matricula"], (
                    it["inscricao_norm"],
                    p["matricula"],
                    it["matricula"],
                )
            assert p["codigo_imovel"] == it["codigo_imovel"], (
                it["inscricao_norm"],
                p["codigo_imovel"],
                it["codigo_imovel"],
            )
            pairs[it["item"]] = p
            used18.add(p["item"])
    antigo = {it["inscricao_norm"] or "01": it for it in i25a}

    # other sources, read from the documents
    t11 = text_of(EXTRA_SRC["inv11_2025"])
    solar_obs = re.search(r"SOLAR SERRA PLACAS SOLARES[^\n]*", t11).group(0)
    solar_val = f0.brl(re.search(r"importância de R\$ " + f0.MONEY, solar_obs).group(1))
    a17 = openpyxl.load_workbook(f0.SRC["resp_arq17"], data_only=True).worksheets[0]
    solar17 = [
        (i, r)
        for i, r in enumerate(a17.iter_rows(values_only=True), 1)
        if any("placas solares Hotel Araçá" in str(x) for x in r)
    ]
    assert len(solar17) == 1
    s17row, s17 = solar17[0]
    t25 = text_of(EXTRA_SRC["doc25_2018"])
    troca = (
        re.search(r"Lote 10 – Quadra 03 – Loteamento Aditivo Madianeira[^|\n]*", t25)
        .group(0)
        .strip()
    )
    aptos = re.findall(
        r"(Um apartamento adquirido no ano de 2011[^|\n]*pag[oa][^|\n]*Hotel Araçá[^|\n]*)\|\| R\$ "
        + f0.MONEY,
        t25,
    )
    draft_insc = {}
    for _, cells in f0.table_rows(EXTRA_SRC["draft_2018"]):
        if cells and re.fullmatch(r"\d{2}", cells[0]) and len(cells) > 1:
            ins = f0.first(r"Inscrição Imobiliária nº ([\d\.]+\d)", cells[1])
            if ins:
                draft_insc[f0.norm_inscricao(ins)] = cells[0]
    iptu = {
        k: partner_capao_lines(EXTRA_SRC[k])
        for k in ("doc1_2018", "doc3_2018", "doc4_2018", "doc5_2018", "doc6_2018")
    }
    assert any("360.0019.004.000" in x for x in iptu["doc6_2018"])
    obs_wb = openpyxl.load_workbook(f0.SRC["observacoes"], data_only=True)[
        "Observacoes"
    ]
    luis_obs = [
        r
        for r in obs_wb.iter_rows(min_row=2, values_only=True)
        if r[1]
        and (
            "Ara" in str(r[1])
            or (
                r[1] == "GERAL"
                and (m := re.match(r"[A-Z]+(\d+)$", str(r[2])))
                and 70 <= int(m.group(1)) <= 104
                and str(r[0]).startswith("2_0")
            )
        )
    ]
    obs_txt = (
        "; ".join(f"{r[0][:30]} {r[1]} {r[2]}: {r[5]}" for r in luis_obs)
        or "nenhuma (2_0 GERAL linhas 70–104 e abas 6.Hotel Araçá das duas planilhas)"
    )
    obs_sit = "Pendente" if luis_obs else "—"
    obs25 = " ".join(f0.observation_paragraphs(f0.SRC["inv2025"]))
    flat7, flat8, flat9 = (re.sub(r"\s+", " ", pages[i]) for i in (6, 7, 8))
    flat24, flat25 = re.sub(r"\s+", " ", pages[23]), re.sub(r"\s+", " ", pages[24])
    m2 = re.search(
        r"oscila entre R\$ ([\d\.]+)/m² e R\$ ([\d\.]+)/m²",
        re.sub(r"\s+", " ", pages[10]),
    )
    m2_lo, m2_hi = f0.amount(m2.group(1)), f0.amount(m2.group(2))

    hdr = [
        "Empresa",
        "Bloco",
        "Chave",
        "Descrição",
        "Existe em 2018?",
        "Ref. 2018",
        "Valor 2018 (referência, P3)",
        "Existe em 2025 (grupo)?",
        "Ref. 2025",
        "Valor 2025 antigo",
        "Existe no Estudo 2 (Rafael)?",
        "Ref. Rafael",
        "Valor grupo 2025",
        "Valor Rafael",
        "Diferença R$",
        "Diferença %",
        "Situação",
        "Lista de Rafael (C0)",
        "Direção C0",
        "Faixa de mercado de Rafael",
        "Decisão (P5)",
        "Valor adotado",
        "Motivo da decisão",
        "Destino 2018→2025",
        "Registro de data-base",
        "Observação do Luís",
        "Situação da observação",
        "Concordância de Rafael",
        "Nome de sócio em 2018",
        "Notas de identificação",
    ]
    C = {h: i for i, h in enumerate(hdr)}
    rows = []

    def new_row(**kw):
        r = [None] * len(hdr)
        r[C["Empresa"]] = EMPRESA
        r[C["Observação do Luís"]] = obs_txt
        r[C["Situação da observação"]] = obs_sit
        r[C["Decisão (P5)"]] = "—"
        for k, v in kw.items():
            r[C[k]] = v
        rows.append(r)
        return r

    def decision(key):
        d = dec["decisoes"].get(key)
        return (
            (d["decisao"], d.get("valor_adotado"), d.get("motivo"))
            if d
            else ("Pendente", None, None)
        )

    datebase = {
        "01": "Registro 1 (Rafael: fontes de mercado de 2025); registro 2 (grupo: CUB de dez/2022)",
        INSC_ITEM20: "Registro 1 (Rafael: fontes de mercado de 2025)",
    }

    for it in i25:
        k = it["inscricao_norm"]
        p = pairs.get(it["item"])
        a = antigo[k or "01"]
        if it["item"] == "01":
            rv, rp, fx = r01, r01p, rng("01 Hotel")
        else:
            rr = raf[k]
            rv, rp = rr["valor"], rr["pagina"]
            fx = (
                rng("20 Condado")
                if k == INSC_ITEM20
                else (rng("Lotes q37/q38") if it["quadra"] in ("37", "38") else None)
            )
        g = it["valor"]
        diff = rv - g
        situ = "Concorda" if abs(diff) < TOL else "Discorda"
        lst = lista.get(k)
        if lst:
            assert (
                lst["item"] == it["item"]
                and lst["matricula"] == it["matricula"]
                and abs(lst["valor"] - g) < TOL
            ), (k, lst)
        c0 = (
            ""
            if situ == "Concorda"
            else ("menor (vem para nós)" if lst else "maior (fica na empresa)")
        )
        dd, dv, dm = decision(f"2025-{it['item']}")
        notas = []
        if it["item"] == "01":
            notas.append(
                "composição e conta de Rafael na aba Item01_componentes; decisão no nível do item"
            )
            notas.append(
                "área construída 12.000 m² em 2018 e 10.500 m² em 2025, a esclarecer"
            )
            notas.append(
                "nesga de 420 m² é 'parte do lote 12' (mat. 33.101), que também entra entre os 13 lotes de 360 m²: conferir área da matrícula"
            )
            notas.append(
                "lotes 13 e 14 da quadra 37 não aparecem em nenhum item (linha própria)"
            )
            notas.append(
                f"grupo estima, sem deduzir, redução de R$ {br(f0.brl(re.search(r'4,00% = R\$ ' + f0.MONEY, obs25).group(1)))} e INSS de R$ {br(f0.brl(re.search(r'aproximadamente R\$ ' + f0.MONEY, obs25).group(1)))} (aba Observacoes)"
            )
        if it["quadra"] == "37" and k and ".0071." in k:
            notas.append("texto diz quadra 37; inscrição é da quadra 0071")
        if it["matricula"] == "73.482":
            notas.append("matrícula 73.482 repetida nos lotes 27 e 28/q38")
        if it["comarca"] and it["comarca"] != "Capão da Canoa":
            notas.append(f"registro na comarca de {it['comarca']}")
        if it["matricula_antiga"]:
            notas.append(f"antiga matrícula {it['matricula_antiga']}")
        if p and p["nome_socio_no_valor"] and k in draft_insc:
            notas.append(
                f"no rascunho do inventário (pasta '2018_desatualizado') é o item {draft_insc[k]} do Hotel, sem nome de sócio"
            )
        if p and p["nome_socio_no_valor"]:
            notas.append(
                f"em 2018 era o lote '{p['nome_socio_no_valor']}', sem valor; em 2025 valorado na empresa; não consta da lista de IPTU do sócio em 2018 (doc 3)"
            )
        if k == INSC_ITEM20:
            notas.append(
                "contrato particular de compra e venda (maio/2014) com matrícula 69.488 citada; o grupo afirma que todos os imóveis estão em nome da empresa (arq. 1, p. 11); Rafael avalia casa de 160 m² (inventário 158,18 m²)"
            )
        if situ == "Concorda" and it["item"] != "01":
            expr = (
                "concordância expressa de Rafael (p. 12–13 e 16, itens 02 a 19)"
                if k in {f0.norm_inscricao(i["inscricao"]) for i in i25[1:19]}
                else "concorda por valor (p. 15 e 20); sem declaração expressa de Rafael (a da p. 16 cita 'itens 02 à 19')"
            )
            notas.append(expr)
            conc_kind = (
                "expressa (p. 12–13 e 16)"
                if expr.startswith("concordância expressa")
                else "por valor (p. 15 e 20)"
            )
            if it["area_m2"] and it["quadra"] in ("37", "38"):
                per_m2 = g / f0.brl(it["area_m2"])
                if m2_lo <= per_m2 <= m2_hi:
                    notas.append(
                        f"R$ {br(per_m2)}/m², dentro da faixa de Rafael por m² (R$ {br(m2_lo)} a {br(m2_hi)}/m², p. 11)"
                    )
                elif per_m2 < m2_lo:
                    notas.append(
                        f"R$ {br(per_m2)}/m², abaixo da faixa de Rafael por m² (R$ {br(m2_lo)} a {br(m2_hi)}/m², p. 11)"
                    )
            if it["quadra"] == "37" and it["lote"] in ("15", "16", "17"):
                notas.append(
                    "lote dentro do quarteirão do hotel (quadra 37, ao lado dos lotes do item 01): conferir se parte do hotel está sobre ele"
                )
        else:
            conc_kind = "—"
        if fx and g < fx[3] - TOL:
            notas.append(
                "valor abaixo do piso da faixa por lote de Rafael, que mesmo assim o adota"
                if situ == "Concorda"
                else "valor do grupo abaixo do piso da faixa de Rafael"
            )
        if not p:
            dest = (
                "— (sem par no inventário 2018; data e forma de aquisição a documentar)"
            )
        else:
            dest = "—"
        new_row(
            **{
                "Bloco": bloco(it),
                "Chave": chave(it, mats01),
                "Descrição": desc_curta(it),
                "Existe em 2018?": "Sim" if p else "Não",
                "Ref. 2018": f"item {p['item']} (tabela {p['tabela']} do docx)"
                if p
                else "",
                "Valor 2018 (referência, P3)": p["valor"] if p else None,
                "Existe em 2025 (grupo)?": "Sim",
                "Ref. 2025": f"item {it['item']} (tabela {it['tabela']} do docx)",
                "Valor 2025 antigo": a["valor"],
                "Existe no Estudo 2 (Rafael)?": "Sim",
                "Ref. Rafael": f"PDF Rafael, balanço executado, p. {rp}",
                "Valor grupo 2025": g,
                "Valor Rafael": rv,
                "Diferença R$": diff,
                "Diferença %": diff / g,
                "Situação": situ,
                "Lista de Rafael (C0)": "Sim" if lst else "Não",
                "Direção C0": c0,
                "Faixa de mercado de Rafael": f"R$ {br(fx[3])} a R$ {br(fx[4])} (p. {fx[0]})"
                if fx
                else "",
                "Decisão (P5)": dd,
                "Valor adotado": dv,
                "Motivo da decisão": dm,
                "Destino 2018→2025": dest,
                "Registro de data-base": datebase.get(
                    "01" if it["item"] == "01" else k, ""
                ),
                "Concordância de Rafael": conc_kind,
                "Nome de sócio em 2018": (
                    p["nome_socio_no_valor"] if p and p["nome_socio_no_valor"] else ""
                ),
                "Notas de identificação": "; ".join(notas),
            }
        )

    # 2018 items with no 2025 match
    iptu_by_insc = {
        "360.0019.004.000": (
            "doc6_2018",
            "doc 6 de 2018 (relação de imóveis para IPTU de Válter e família)",
        ),
    }
    for it in i18[1:]:
        if it["item"] in used18:
            continue
        k = it["inscricao_norm"]
        notas = []
        if it["nome_socio_no_valor"]:
            notas.append(
                f"no lugar do valor: {it['nome_socio_no_valor']}; bem da empresa ou do sócio a esclarecer"
            )
        if k in iptu_by_insc:
            src, lab = iptu_by_insc[k]
            line = next(x for x in iptu[src] if k in x)
            notas.append(f"{lab}: '{line}'")
        if "TROCOU" in it["descricao"]:
            notas.append(
                "nota de 2018: 'trocou este lote/casa por um lote no Loteamento Aditivo Medianeira'; valor '(Avaliar Lote e Casa)'; o total de 2018 ressalva '+ o valor da casa e lote – Élcio Siviero'"
            )
            notas.append(
                f"doc 25 de 2018 (seção Élcio): '{troca}' (lote da SEI: candidato à Parte 6)"
            )
        if k == "360.0019.004.000":
            a18 = re.sub(r"\s+", " ", f0.pdf_pages(EXTRA_SRC["arq18_pdf"])[2])
            m = re.search(
                r"Lote Capão da Canoa Lote na esquina próximo ao Hotel Araçá Ltda\., com valor diferenciado aos demais lotes dados aos sócios\.",
                a18,
            )
            assert m
            notas.append(
                f"arquivo 18, p. 3 (planilha de Válter 'para acerto futuro', bloco de 10.06.2019): '{m.group(0)}' — sem inscrição; pode ser este lote, mas não está provado; risco de cruzar com o item 02 (lote 03/q19, lista de Rafael)"
            )
        casa_doc25 = {
            "360.0057.001.000": r"CASA DA ESQUINA: [^.]*?\)",
            "360.0057.002.000": r"CASA AO LADO: Raul Siviero",
        }.get(k)
        if casa_doc25:
            notas.append(
                f"doc 25 de 2018, seção Hotel: '{re.search(casa_doc25, t25).group(0)}', entre as '02 casas de veraneio dos Sócios' mantidas pelo Hotel"
            )
        if "CASA" in it["descricao"]:
            notas.append(
                "lote com casa; possível correspondência com as duas residências da Rua da Divisa citadas por Rafael (p. 24): a confirmar"
            )
        if k in draft_insc:
            notas.append(
                f"no rascunho do inventário (pasta '2018_desatualizado', arquivo datado de dez/2015) é o item {draft_insc[k]} do Hotel, sem nome de sócio e sem valor"
            )
        socio_doc = {
            "RAUL SIVIERO": "doc1_2018",
            "PEDRO SIVIERO": "doc4_2018",
            "ADÉLIA F. SIVIERO": "doc5_2018",
        }.get(it["nome_socio_no_valor"])
        # neighbours of 2025 items (same block, lot ±1) and sequential matrículas
        for n in i25[1:]:
            a_, b_ = (k or "").split("."), (n["inscricao_norm"] or "").split(".")
            if (
                len(a_) == 4
                and len(b_) == 4
                and a_[:2] == b_[:2]
                and abs(int(a_[2]) - int(b_[2])) == 1
            ):
                notas.append(
                    f"lote vizinho do item {n['item']} de 2025 (inscrição {n['inscricao_norm']})"
                    + (
                        ", que está na lista de Rafael"
                        if n["inscricao_norm"] in lista
                        else ""
                    )
                )
            if (
                it["matricula"]
                and n["matricula"]
                and abs(
                    int(it["matricula"].replace(".", ""))
                    - int(n["matricula"].replace(".", ""))
                )
                == 1
            ):
                notas.append(
                    f"matrícula {it['matricula']} sequencial à {n['matricula']} do item {n['item']} de 2025"
                    + (" (lista de Rafael)" if n["inscricao_norm"] in lista else "")
                )
        if socio_doc and iptu[socio_doc]:
            notas.append(
                f"lista de IPTU de 2018 do sócio ({socio_doc.split('_')[0].replace('doc', 'doc ')}) tem imóvel em Capão da Canoa sem inscrição: '{iptu[socio_doc][0]}'"
            )
        new_row(
            **{
                "Bloco": "Lotes de 2018 sem par em 2025",
                "Chave": chave(it),
                "Descrição": desc_curta(it),
                "Existe em 2018?": "Sim",
                "Ref. 2018": f"item {it['item']} (tabela {it['tabela']} do docx)",
                "Existe em 2025 (grupo)?": "Não",
                "Existe no Estudo 2 (Rafael)?": "a verificar (possíveis residências da p. 24)"
                if "CASA" in it["descricao"]
                else "Não",
                "Situação": "Só 2018",
                "Lista de Rafael (C0)": "Não",
                "Destino 2018→2025": dec["destinos_2018"].get(
                    f"2018-{it['item']}", "Pendente (Parte 4)"
                ),
                "Concordância de Rafael": "—",
                "Nome de sócio em 2018": it["nome_socio_no_valor"] or "",
                "Notas de identificação": "; ".join(notas),
            }
        )

    # cited only by Rafael
    casas = re.search(r"fica entre R\$ ([\d\.]+) e R\$ ([\d\.]+)", flat24)
    assert casas and "Fiat Toro" in flat25
    new_row(
        **{
            "Bloco": "Citados só por Rafael",
            "Chave": "sem chave (Rafael não dá matrícula nem inscrição)",
            "Descrição": "Duas residências de ~200 m² em dois lotes em sequência, esquina da Rua da Divisa com a Rua Três, bairro Araçá",
            "Existe em 2018?": "a verificar (itens 22 e 23 de 2018?)",
            "Existe em 2025 (grupo)?": "Não",
            "Existe no Estudo 2 (Rafael)?": "Sim",
            "Ref. Rafael": "PDF Rafael p. 24–25 ('Observação importante')",
            "Situação": "Só Rafael",
            "Lista de Rafael (C0)": "Não",
            "Faixa de mercado de Rafael": f"R$ {casas.group(1)} a R$ {casas.group(2)} pelo par (p. 24)",
            "Decisão (P5)": "Pendente",
            "Destino 2018→2025": "—",
            "Registro de data-base": "Registro 1 (Rafael: fontes de mercado de 2025)",
            "Notas de identificação": "não somadas ao total de Rafael; titular desconhecido",
        }
    )
    new_row(
        **{
            "Bloco": "Citados só por Rafael",
            "Chave": "sem chave (Rafael não dá placa)",
            "Descrição": "Fiat Toro, automóvel de serviço do Hotel Araçá",
            "Existe em 2018?": "Não",
            "Existe em 2025 (grupo)?": "Não",
            "Existe no Estudo 2 (Rafael)?": "Sim",
            "Ref. Rafael": "PDF Rafael p. 25",
            "Situação": "Só Rafael",
            "Lista de Rafael (C0)": "Não",
            "Decisão (P5)": "Pendente",
            "Destino 2018→2025": "—",
            "Notas de identificação": "sem valor; não somada ao total de Rafael; titular e placa desconhecidos",
        }
    )

    # found in other sources, valued by nobody
    d22 = text_of(
        EXTRA_SRC["reuniao_2017"]
    )  # same sheet as doc 22 of 2018 (read via the Drive connector)
    assert "37.698,00" in EXTRA_SRC["doc22_trechos"].read_text(encoding="utf-8")
    cofre = re.search(
        r"U\$ ([\d\.]+,\d{2}) – TOTAL NO COFRE EM VERANÓPOLIS", d22
    ).group(1)
    new_row(
        **{
            "Bloco": "Não abordados",
            "Chave": "sem chave",
            "Descrição": f"Dólares do Hotel Araçá guardados no cofre em Veranópolis (U$ {cofre} em 2017)",
            "Existe em 2018?": "Sim, fora do inventário: anexo 'Planilha – resumo dos dólares' da convocação de 26/05/2017, repetido na ata de 14/01/2019 (doc 22 de 2018)",
            "Existe em 2025 (grupo)?": "Não",
            "Existe no Estudo 2 (Rafael)?": "Não",
            "Ref. 2018": "Material antigo/2018_desatualizado/G-SÓCIOS… Reunião no dia 26 de maio de 2017.docx; doc 22 de 2018 (lido pelo conector do Drive)",
            "Situação": "Não abordado",
            "Lista de Rafael (C0)": "Não",
            "Decisão (P5)": "Pendente",
            "Destino 2018→2025": "Pendente (Parte 4)",
            "Notas de identificação": "saldo de U$ 37.698,00 no cofre em 25/05/2017 (2006–2017); o mesmo anexo registra U$ 9.000,00 'Pago Coletto–Gramado – 02 lotes' (2009) e retiradas para Válter e Luciane (2012–2013); doc lido pelo conector do Drive (arquivo não baixado); destino depois de 2017 a perguntar",
        }
    )

    new_row(
        **{
            "Bloco": "Não abordados",
            "Chave": "sem chave",
            "Descrição": "Placas solares instaladas no hotel (Solar Serra – Mário Bordignon)",
            "Existe em 2018?": "Não",
            "Existe em 2025 (grupo)?": "Não (não citadas no inventário 06)",
            "Existe no Estudo 2 (Rafael)?": "Não",
            "Ref. 2025": f"inventário 11 (SEI) 2025, observação 'SOLAR SERRA'; arquivo 17, linha {s17row}",
            "Situação": "Não abordado",
            "Lista de Rafael (C0)": "Não",
            "Decisão (P5)": "Pendente",
            "Destino 2018→2025": "—",
            "Notas de identificação": (
                f"arq. 17: {s17[0]}, lote {s17[2]}/q{s17[3]}, {s17[1]:%d/%m/%Y}, {s17[4]}: '{s17[5]}'; inventário 11: compromisso de mais R$ {br(solar_val)} "
                "em 1 lote 'além do valor pago'; pagas pela SEI e fora do ativo do Hotel (equipamentos diversos R$ 8.000.000 brutos em 2018 e em 2025); valor total a perguntar; candidato à Parte 6"
            ),
        }
    )
    new_row(
        **{
            "Bloco": "Não abordados",
            "Chave": "quadra 37, lotes 13 e 14",
            "Descrição": "Lotes 13 e 14 da quadra 37 (dentro do quarteirão do hotel)",
            "Existe em 2018?": "Não",
            "Existe em 2025 (grupo)?": "Não",
            "Existe no Estudo 2 (Rafael)?": "Não",
            "Situação": "Não abordado",
            "Lista de Rafael (C0)": "Não",
            "Decisão (P5)": "—",
            "Destino 2018→2025": "—",
            "Notas de identificação": "o item 01 cobre os lotes 01–12 e 18–22 + nesga; os lotes 15, 16 e 17 são os itens 21, 03 e 22; titularidade dos lotes 13 e 14 a perguntar (pesa na unificação da quadra)",
        }
    )

    # approved rules without value (R3): Decisão and Motivo from decisoes.json
    # generic rules first; rules with a description filter (later packets) override them
    rules = sorted(dec.get("regras", {}).values(), key=lambda x: "filtro" in x)
    for rule in rules:
        for r in rows:
            if r[C["Situação"]] not in rule["situacoes"]:
                continue
            if "filtro" in rule:
                if rule["filtro"] not in (r[C["Descrição"]] or ""):
                    continue
            elif r[C["Decisão (P5)"]] not in ("—", "Pendente"):
                continue
            elif (
                "descricoes" in rule
                and (r[C["Descrição"]] or "") not in rule["descricoes"]
            ):
                continue  # rule approved for a fixed set of rows; later rows need their own decision
            r[C["Decisão (P5)"]] = "Sem valor nesta etapa"
            r[C["Motivo da decisão"]] = rule["texto"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mestre"
    ws.append(hdr)
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for h in (
            "Valor 2018 (referência, P3)",
            "Valor 2025 antigo",
            "Valor grupo 2025",
            "Valor Rafael",
            "Diferença R$",
            "Valor adotado",
        ):
            row[C[h]].number_format = "#,##0.00"
        row[C["Diferença %"]].number_format = "0.0%"
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9D9D9")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions

    lg = wb.create_sheet("Legenda")
    lg.append(["Situação", "Significado"])
    for s in SITUACOES:
        lg.append(list(s))
    lg.append([])
    lg.append(
        [
            "Ref. 2018 / Ref. 2025",
            "número do item e número da tabela no docx (o docx não tem página fixa)",
        ]
    )
    lg.append(
        [
            "Direção C0",
            "para onde o critério C0 aponta quando há divergência; não é decisão",
        ]
    )
    lg.append(
        [
            "Chave",
            "inscrição imobiliária normalizada (4 grupos) | matrícula | código de imóvel",
        ]
    )

    # item 01, component by component
    ic = wb.create_sheet("Item01_componentes")
    ic.append(
        [
            "Componente",
            "Grupo 2018",
            "Grupo 2025 antigo",
            "Grupo 2025 atual",
            "Rafael (derivado da p. 7)",
            "Diferença Rafael − grupo 2025",
            "Base de cálculo",
            "Nota",
        ]
    )
    comp = {lab: f0.item01_components(inv[lab][0]["descricao"], lab) for lab in inv}
    land = f0.amount(re.search(r"terreno a R\$ ([\d\.]+)/m²", flat7).group(1))
    built_area = f0.amount(re.search(r"construída de ([\d\.]+) m²", flat7).group(1))
    built_rate = f0.amount(re.search(r"m² a R\$ ([\d\.]+)/m² pelo CUB", flat7).group(1))
    land_area = f0.brl(comp["2025 atual"][0]["area_m2"])
    raf_land, raf_built = land_area * land, built_area * built_rate
    assert abs(raf_land + raf_built - r01) < TOL, (raf_land, raf_built, r01)
    vant = comp["2025 atual"][5]["texto"]
    vant_base = re.search(r"obra feita: R\$ " + f0.MONEY, vant).group(1)
    nao_sep = "não avaliado em separado por Rafael (incluído ou omitido nos 66 milhões: a definir)"
    raf_comp = {
        "1.1 Terreno": (
            raf_land,
            f"{br(land_area)} m² × R$ {br(land)}/m²",
            "nesga de 420 m² = 'parte do lote 12' (mat. 33.101), que também conta entre os 13 lotes de 360 m²: conferir a área da matrícula",
        ),
        "1.2 Móveis 36 aptos novos": (None, nao_sep, ""),
        "1.3 Móveis 108 aptos": (None, nao_sep, ""),
        "1.4 Obra civil": (
            raf_built,
            f"{br(built_area)} m² × R$ {br(built_rate)}/m² ('depreciação técnica … pelo CUB-RS modificado'); grupo líquido = R$ {br(f0.brl(comp['2025 atual'][3]['valor_unitario']) * 0.8)}/m²",
            "área construída 12.000 m² em 2018 e 10.500 m² em 2025 (grupo e Rafael): a esclarecer; CUB do grupo é de dez/2022",
        ),
        "1.5 Equipamentos diversos": (
            None,
            nao_sep,
            "inclui o poço com outorga, sem valor próprio",
        ),
        "1.6 Vantagem da coisa feita": (
            None,
            nao_sep,
            f"base R$ {vant_base} = obra civil depreciada + móveis + equipamentos (não só a obra)",
        ),
    }
    for idx, (name, _) in enumerate(f0.COMPONENTS):
        g18, ga, g25 = (comp[lab][idx] for lab in ("2018", "2025 antigo", "2025 atual"))
        rv, base_r, nota = raf_comp[name]
        base_g = ", ".join(
            x
            for x in [
                g25["area_m2"] and f"{g25['area_m2']} m²",
                g25["valor_unitario"] and f"R$ {g25['valor_unitario']}",
                g25["base_cub"],
                g25["depreciacao_pct"] and f"depreciação {g25['depreciacao_pct']} %",
                g25["percentual_vantagem"]
                and f"{g25['percentual_vantagem']} % sobre R$ {vant_base}",
            ]
            if x
        )
        g18base = ", ".join(
            x
            for x in [
                g18["area_m2"] and f"{g18['area_m2']} m²",
                g18["valor_unitario"] and f"R$ {g18['valor_unitario']}",
                g18["base_cub"],
                g18["percentual_vantagem"] and f"{g18['percentual_vantagem']} %",
            ]
            if x
        )
        ic.append(
            [
                name,
                g18["valor"],
                ga["valor"],
                g25["valor"],
                rv,
                (rv - g25["valor"]) if rv is not None else None,
                f"grupo 2025: {base_g}; Rafael: {base_r}",
                "; ".join(x for x in [g18base and f"2018: {g18base}", nota] if x),
            ]
        )
    tot = [
        sum(c["valor"] for c in comp[lab])
        for lab in ("2018", "2025 antigo", "2025 atual")
    ]
    rooms = int(re.search(r"capacidade estimada de (\d+) quartos", flat9).group(1))
    adr = float(re.search(r"Estipulada em R\$ (\d+)", flat9).group(1))
    occ = float(re.search(r"Média conservadora de (\d+)%", flat9).group(1)) / 100
    ebitda = float(re.search(r"Fixada em (\d+)% para a operação", flat9).group(1)) / 100
    caps = [
        float(x.replace(",", ".")) / 100
        for x in re.findall(
            r"Cap Rate\)? (?:mais comprimido )?de (\d+,\d)% ao ano", flat8
        )
    ]
    noi = rooms * adr * 365 * occ * ebitda
    renda = " e ".join(
        f"R$ {br(noi / c)} a {str(c * 100).replace('.', ',')} %" for c in caps
    )
    ic.append(
        [
            "Total item 01",
            *tot,
            r01,
            r01 - tot[2],
            f"Rafael: balanço executado p. {r01p}; faixa de mercado R$ 55 a 85 milhões; renda R$ 41,0 a 50,6 milhões (p. 7–8)",
            f"decisão no nível do item, não por componente; a aritmética das premissas de renda de Rafael (p. 8–9: {rooms} quartos, R$ {adr:.0f}, {occ:.0%}, margem {ebitda:.0%}) dá lucro de R$ {br(noi)}/ano e {renda}, não os 41,0 e 50,6 milhões que ele escreve",
        ]
    )
    for row in ic.iter_rows(min_row=2):
        for c in row[1:6]:
            c.number_format = "#,##0.00"

    # totals and reconciliation
    tt = wb.create_sheet("Totais")
    g_tot = sum(i["valor"] for i in i25)
    r_sum = r01 + sum(r["valor"] for r in rlots)
    S, G, R, V = (
        C["Situação"],
        C["Valor grupo 2025"],
        C["Valor Rafael"],
        C["Valor adotado"],
    )
    n_conc = sum(1 for r in rows if r[S] == "Concorda")
    n_disc = sum(1 for r in rows if r[S] == "Discorda")
    conc = sum(r[G] for r in rows if r[S] == "Concorda")
    disc_g = sum(r[G] for r in rows if r[S] == "Discorda")
    disc_r = sum(r[R] for r in rows if r[S] == "Discorda")
    sel = sum(v["valor"] for v in lista.values())
    lot_vals = sorted({r["valor"] for r in rlots})
    same = [i["item"] for i in i25[1:] if abs(i["valor"] - (rtotal - r_sum)) < TOL]
    tt.append(["Linha", "Valor", "Fonte / conta"])
    for line in (
        [
            [
                f"Total grupo 2025 ({len(i25)} itens)",
                g_tot,
                "inventário 2025, TOTAL GERAL",
            ],
            [
                f"Soma dos itens de Rafael ({len(rlots) + 1} itens)",
                r_sum,
                "PDF Rafael p. 16–20",
            ],
            ["Total declarado por Rafael", rtotal, "PDF Rafael p. 20"],
            [
                "Diferença de soma de Rafael (não identificada)",
                rtotal - r_sum,
                f"total declarado − soma dos itens; igual ao valor de cada um dos itens {', '.join(same)} (R$ {br(rtotal - r_sum)}); valores de lote de Rafael: {', '.join(br(v) for v in lot_vals)}",
            ],
            ["1/6 declarado por Rafael", rtotal / 6, "PDF Rafael p. 21"],
            ["1/6 pela soma dos itens de Rafael", r_sum / 6, "soma dos itens / 6"],
            ["Crédito declarado por Rafael", rtotal / 6 - sel, "PDF Rafael p. 21"],
            [
                "Crédito pela soma dos itens de Rafael",
                r_sum / 6 - sel,
                "1/6 pela soma − lotes escolhidos",
            ],
            [f"Itens que concordam: valor ({n_conc} itens)", conc, ""],
            [f"Itens que discordam: grupo ({n_disc} itens)", disc_g, ""],
            ["Itens que discordam: Rafael", disc_r, ""],
            [
                "Diferença explicada por itens",
                disc_r - disc_g,
                "Rafael − grupo nos itens que discordam",
            ],
            [
                "Diferença total Rafael declarado − grupo",
                rtotal - g_tot,
                "= explicada por itens + diferença de soma",
            ],
            [
                f"Total 2018 ({n_valued_18} itens com valor)",
                sum(i["valor"] or 0 for i in i18),
                "inventário 2018; o próprio documento ressalva '+ o valor da casa e lote – Élcio Siviero', não somado",
            ],
            [
                f"Lista de Rafael ({len(lista)} itens)",
                sel,
                "ativos-avaliados-valter.xlsx",
            ],
            [
                "Valores adotados (decisões aprovadas)",
                sum(r[V] for r in rows if r[V] is not None),
                f"{sum(1 for r in rows if r[V] is not None)} itens com decisão",
            ],
            [
                "Valores adotados: 1/6",
                sum(r[V] for r in rows if r[V] is not None) / 6,
                "valores adotados / 6",
            ],
        ]
        + [
            [f"Decisão sem item: {k}", 0.0, v["texto"]]
            for k, v in dec.get("excecoes_sem_item", {}).items()
        ]
        + [
            [f"Pendência de componente: {k}", None, v["texto"]]
            for k, v in dec.get("pendencias_componentes", {}).items()
        ]
    ):
        tt.append(line)
    for row in tt.iter_rows(min_row=2):
        row[1].number_format = "#,##0.00"
    assert abs(conc + disc_g - g_tot) < TOL
    assert abs((disc_r - disc_g) + (rtotal - r_sum) - (rtotal - g_tot)) < TOL

    # 2018 x 2025 pairs
    pr = wb.create_sheet("Pares_2018_2025")
    pr.append(
        [
            "Item 2025",
            "Item 2018",
            "Inscrição (normalizada)",
            "Matrícula 2025",
            "Matrícula 2018",
            "Código 2025",
            "Código 2018",
            "Critério",
            "Valor 2018",
            "Valor 2025",
        ]
    )
    for it in i25:
        p = pairs.get(it["item"])
        if it["item"] == "01":
            crit = "item 01: as mesmas 12 matrículas nas três versões"
        elif p:
            crit = (
                "inscrição + código + matrícula"
                if p["matricula"]
                else "inscrição + código (2018 sem matrícula)"
            )
        else:
            crit = "sem par"
        pr.append(
            [
                it["item"],
                p["item"] if p else "",
                it["inscricao_norm"],
                it["matricula"],
                p["matricula"] if p else "",
                it["codigo_imovel"],
                p["codigo_imovel"] if p else "",
                crit,
                p["valor"] if p else None,
                it["valor"],
            ]
        )
    for it in i18:
        if it["item"] not in used18:
            pr.append(
                [
                    "",
                    it["item"],
                    it["inscricao_norm"],
                    "",
                    it["matricula"],
                    "",
                    it["codigo_imovel"],
                    "sem par em 2025",
                    it["valor"],
                    None,
                ]
            )

    obs = wb.create_sheet("Observacoes")
    obs.append(["Fonte", "nº", "Parágrafo"])
    for key, label in [
        ("inv2025", "inventário 2025 atual"),
        ("inv2018", "inventário 2018"),
    ]:
        for n, t in enumerate(f0.observation_paragraphs(f0.SRC[key]), 1):
            obs.append([label, n, t])
    obs.append(["inventário 11 (SEI) 2025", 1, solar_obs])
    for n, (txt, val) in enumerate(aptos, 1):
        obs.append(
            [
                "doc 25 de 2018 (pagos pelo Hotel a favor de sócio)",
                n,
                f"{txt.strip()} – R$ {val}",
            ]
        )
    for k, lines in iptu.items():
        for n, x in enumerate(lines, 1):
            obs.append(
                [f"{k.replace('_2018', '')} de 2018 (IPTU, Capão da Canoa)", n, x]
            )

    for sheet in wb.worksheets:
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = min(
                60, max(10, max(len(str(c.value or "")) for c in col[:40]) * 0.9)
            )
    wb.save(OUT)
    cnt = {s: sum(1 for r in rows if r[S] == s) for s, _ in SITUACOES}
    print(f"rows {len(rows)}: {cnt}")
    print(
        f"concorda {n_conc} = {conc:,.2f}; discorda {n_disc}: grupo {disc_g:,.2f} x Rafael {disc_r:,.2f}"
    )
    print(
        f"Rafael soma {r_sum:,.2f} declarado {rtotal:,.2f} dif {rtotal - r_sum:,.2f} (= itens {same})"
    )
    print(
        f"item 01 Rafael: terreno {raf_land:,.2f} construção {raf_built:,.2f}; renda NOI {noi:,.2f} caps {caps}"
    )
    print(
        f"pares: {len(pairs) - 1} lotes + item 01; sem par 2025: {[i['item'] for i in i25[1:] if i['item'] not in pairs]}; sem par 2018: {[i['item'] for i in i18 if i['item'] not in used18]}"
    )
    print(
        f"lista C0 por inscrição: {sorted(v['item'] for v in lista.values())}; obs Luís: {len(luis_obs)}; aptos doc25: {len(aptos)}"
    )
    print("saved", OUT, f0.md5_10(OUT))


if __name__ == "__main__":
    main()
