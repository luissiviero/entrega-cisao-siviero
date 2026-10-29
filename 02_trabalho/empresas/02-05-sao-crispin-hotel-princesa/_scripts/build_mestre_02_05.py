"""Build planilha-mestre-02-05.xlsx (companies 02 and 05).

Regenerates the Phase 0 extraction first (fase0_extracao.main), then reads
`fase0-extracoes-02-05.xlsx` and builds one row per asset:
  - every 2025 asset of doc 05 (doc 02 has none);
  - every 2018 asset of docs 13/17 with no pair among them;
  - assets cited by the group but valued by nobody (soy credits, Parte 5).

2018 x 2025 pairs are made by matrícula plus area, seller or plate, never by
text similarity (lesson 3 of the plan). For 2018 assets now in another
company, Rafael's value is looked up in the document of that company
(lesson 1). Approved decisions are read from `decisoes.json`.

v2 (28/10/2026) applies the review in `_diario/revisao-mestre.md`.

Run: PYTHONIOENCODING=utf-8 python -W ignore build_mestre_02_05.py
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import re
import sys
from decimal import Decimal

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fase0_extracao as f0

DIR = f0.OUT_DIR
F0 = f0.OUT_XLSX
OUT = os.path.join(DIR, "planilha-mestre-02-05.xlsx")
DECISIONS = os.path.join(DIR, "decisoes.json")
RAFAEL_DIR = os.path.join(f0.ROOT, r"01_fontes\estudo-rafael")
RAFAEL_OTHER = {
    # 2025 inventory key of the other company -> Rafael's source for it
    "inv2025_11": (
        "xlsx",
        f0.one(os.path.join(RAFAEL_DIR, r"empreendimentos-xlsx\ÁREAS RURAIS*.xlsx")),
    ),
    "inv2025_08": ("pdf", f0.one(os.path.join(RAFAEL_DIR, "TRANSPORTADORA*.pdf"))),
}

# Observation cells of the Google sheet (Phase 0, sheet "Planilhas Google")
# mapped to master keys. Every observation cell must be in OBS_MAP, OBS_DUP,
# OBS_COMPANY or OBS_SKIP, or the build stops.
OBS_MAP = {
    ("GERAL", "H30"): ["05-2025-01"],
    ("GERAL", "H34"): ["05-2025-01"],
    ("GERAL", "H35"): ["05-2025-01"],
    ("GERAL", "H38"): ["05-2025-01"],
    ("GERAL", "H46"): ["05-2025-04"],
    ("GERAL", "H52"): ["05-2025-05"],
    ("GERAL", "H63"): ["05-2025-06"],
    ("GERAL", "H882"): ["05-2025-05"],
    ("Resumo", "B80"): ["05-2025-05"],
    ("Resumo", "E80"): ["05-2025-05"],
    ("Resumo", "B87"): ["05-2025-05"],
    ("Resumo", "B88"): ["05-2025-05"],
    ("Resumo", "E88"): ["05-2025-05"],
    ("Resumo", "F88"): ["05-2025-05"],
    ("Resumo", "F91"): ["05-2025-05"],
    ("Resumo", "B92"): ["05-2025-05"],
    ("Resumo", "E92"): ["05-2025-05"],
    ("Resumo", "C94"): ["05-2025-05"],
    ("Resumo", "B96"): ["05-2025-05"],
    ("Resumo", "E96"): ["05-2025-05"],
    ("Resumo", "B100"): ["05-2025-05"],
    ("Resumo", "B191"): ["05-2025-05"],
    ("GERAL", "B230"): ["02-2018-s10"],
    ("GERAL", "H230"): ["02-2018-s10"],
    ("GERAL", "B235"): ["02-2018-s12"],
    ("GERAL", "F943"): [
        "02-2018-s03",
        "02-2018-s04",
        "02-2018-s05",
        "02-2018-s06",
        "05-2025-05",
    ],
    ("GERAL", "F947"): ["02-2018-s03"],
    ("GERAL", "F948"): ["02-2018-s04"],
    ("GERAL", "F949"): ["02-2018-s05"],
    ("GERAL", "F950"): ["02-2018-s06"],
    ("GERAL", "H14"): ["NA-SOJA-LOR-2021-2023", "NA-SOJA-LOR-2024", "05-2025-05"],
    ("GERAL", "F986"): ["NA-SOJA-RIGO"],
    ("GERAL", "F987"): ["NA-SOJA-LOR-2021-2023", "NA-SOJA-LOR-2024"],
}
# same text as a mapped cell (the company tabs repeat column H of GERAL)
OBS_DUP = {
    ("2.São Crispin", "I5"): ("GERAL", "H14"),
    ("5.Hotel Princesa", "I4"): ("GERAL", "H30"),
    ("5.Hotel Princesa", "I8"): ("GERAL", "H34"),
    ("5.Hotel Princesa", "I9"): ("GERAL", "H35"),
    ("5.Hotel Princesa", "I12"): ("GERAL", "H38"),
    ("5.Hotel Princesa", "I20"): ("GERAL", "H46"),
    ("5.Hotel Princesa", "I26"): ("GERAL", "H52"),
    ("5.Hotel Princesa", "I37"): ("GERAL", "H63"),
    ("11.Empreendimentos", "I731"): ("GERAL", "H882"),
}
# notes about the company as a whole (go to the Resumo sheet)
OBS_COMPANY = {("Resumo", "B20"): "Lista 'VALUATIONS NECESSÁRIAS' (Resumo B16–B20)"}
# copies of the inventory's own OBSERVAÇÕES text, not notes of Luís
OBS_SKIP = {("GERAL", "H69"), ("5.Hotel Princesa", "I43")}

# data-base register (registro-data-base.md): RDB-1 covers every Hotel item
RDB_EXTRA = {
    "05-2025-05": ["RDB-3"],
    "05-2025-08": ["RDB-2"],
    "05-2025-01": ["RDB-4"],
    "NA-SOJA-LOR-2024": ["RDB-3"],
}
PARTE6 = {
    "05-2025-04": "P6-01",
    "05-2025-07": "P6-02",
    "02-2018-s10": "P6-03",
    "02-2018-s11": "P6-04",
    "02-2018-s12": "P6-04",
    "02-2018-s14": "P6-05",
    "02-2018-s15": "P6-05",
    "05-2025-05": "P6-06, P6-08",
    "NA-SOJA-RIGO": "P6-06, P6-07",
}


def dec(v) -> Decimal | None:
    if v is None or v == "":
        return None
    return Decimal(str(v)).quantize(Decimal("0.01"))


def read(wb, name: str) -> list[dict]:
    ws = wb[name]
    hr = 3 if ws["A2"].value is None and ws["A3"].value is not None else 1
    hdr = [c.value for c in ws[hr]]
    return [
        dict(zip(hdr, r))
        for r in ws.iter_rows(min_row=hr + 1, values_only=True)
        if any(v is not None for v in r)
    ]


def rafael_other(doc_key: str, mats: set[str]) -> dict | None:
    """Rafael's value for an asset that sits in another company's 2025 inventory."""
    kind, path = RAFAEL_OTHER[doc_key]
    name = os.path.basename(path)
    if kind == "xlsx":
        wb = openpyxl.load_workbook(path, data_only=True)
        grp = [
            r
            for r in wb["Balanco_Grupo"].iter_rows(min_row=2, values_only=True)
            if r[0]
        ]
        hit = [r for r in grp if mats & set(re.split(r",\s*", str(r[3] or "")))]
        if len(hit) != 1:
            return None
        item = hit[0][0]
        qr = [
            r
            for r in wb["Quadro_Resumo"].iter_rows(min_row=2, values_only=True)
            if str(r[1] or "").startswith(f"Item {item} ")
        ]
        if len(qr) != 1:
            return None
        # evaluation pages (first Avaliacao row of the item) and matrícula
        # misspellings on those pages of the primary PDF
        ava = [
            r
            for r in wb["Avaliacao"].iter_rows(min_row=2, values_only=True)
            if r[0] == item
        ]
        pages = sorted({int(r[9]) for r in ava if r[9]})
        pdf = f0.one(
            os.path.join(
                RAFAEL_DIR, "empreendimentos-pdf", os.path.splitext(name)[0] + ".pdf"
            )
        )
        doc = f0.pymupdf.open(pdf)
        typos = []
        for p in pages:
            for m in re.findall(
                r"Matr[íi]cula\s*(?:nº\s*)?([\d.]{3,})",
                " ".join(doc[p - 1].get_text().split()),
            ):
                m = m.rstrip(".")
                for want in mats:
                    if m != want and sorted(m) == sorted(want):
                        typos.append(f"p. {p} grafa a matrícula como {m}")
        return {
            "doc": f"{os.path.basename(pdf)} (valores pelo xlsx de apoio {name})",
            "item": item,
            "grupo": dec(hit[0][4]),
            "pagina_grupo": hit[0][6],
            "rafael": dec(qr[0][3]),
            "pagina": qr[0][4],
            "pagina_ava": f"{pages[0]}–{pages[-1]}" if pages else "",
            "concordancia": "; ".join(sorted(set(typos))),
        }
    values, _, notes, blocks = f0.rafael_sections(path)
    markers = [
        m
        for (sec, m), text in blocks.items()
        if sec == "Balanço apresentado pelo grupo"
        and m
        and any(
            re.search(r"(?<![\d.])" + re.escape(x) + r"(?![\d])", text) for x in mats
        )
    ]
    if len(markers) != 1:
        return None
    item = f"{markers[0]:02d}"

    def first(sec):
        return next(
            (v for v in values if v["secao"] == sec and v["item"] == item), None
        )

    g, e = (
        first("Balanço apresentado pelo grupo"),
        first("Balanço executado por Valter"),
    )
    conc = [
        n for n in notes if n["tipo"] == "concordância expressa" and n["item"] == item
    ]
    return {
        "doc": name,
        "item": item,
        "grupo": dec(g["valor"]),
        "pagina_grupo": g["pagina"],
        "rafael": dec(e["valor"]),
        "pagina": e["pagina"],
        "concordancia": f"concordância expressa p. {conc[0]['pagina']}" if conc else "",
    }


def main() -> None:
    with contextlib.redirect_stdout(io.StringIO()) as buf:
        f0.main()
    tail = buf.getvalue().strip().splitlines()[-1]
    wb = openpyxl.load_workbook(F0, data_only=True)

    i25 = [
        r
        for r in read(wb, "Itens 2025")
        if r["Versão"] == "2025 atual" and r["Empresa"].startswith("05")
    ]
    old25 = {r["Item"]: r for r in read(wb, "2025 antigo x atual")}
    i18 = read(wb, "Itens 2018")
    raf_vals = read(wb, "Rafael 05 valores")
    comp = {r["Item"]: r for r in read(wb, "Grupo x Rafael 05")}
    faixas = read(wb, "Rafael 05 faixas")
    lista = {
        r["Item"]: r
        for r in read(wb, "Lista de Rafael")
        if str(r["Empresa"]).startswith("05")
    }
    soy = read(wb, "Créditos de soja")
    where = read(wb, "2018 → 2025 por matrícula")
    gs = read(wb, "Planilhas Google")
    bases = {r["Linha"]: r["Valor (R$)"] for r in read(wb, "Base 2018")}
    vendas = [r for r in read(wb, "Vendas 2018") if r["Fonte"] == "doc 17 2018"]

    def section(name):
        out: dict[str, tuple] = {}
        for r in raf_vals:
            if r["Seção do PDF"] == name and r["Item"] not in out:
                out[r["Item"]] = (r["Valor (R$)"], r["Página"], r.get("Observação"))
        return out

    grp_rafael = section("Balanço apresentado pelo grupo")
    exe_rafael = section("Balanço executado por Valter")
    ava_rafael = section("Avaliação executada por Valter")
    esc_rafael = section("Bens escolhidos")
    dup_exe = {
        r["Item"]: r["Página"]
        for r in raf_vals
        if r["Seção do PDF"] == "Balanço executado por Valter" and r.get("Observação")
    }

    def pair_2018(x25: dict) -> tuple[dict | None, str]:
        mats25 = {m for m in (x25["Matrículas"] or "").split(", ") if m}
        if "JTB 8585" in x25["Descrição"]:
            hit = [y for y in i18 if "JTB 8585" in (y["Descrição"] or "")]
            return (hit[0], "placa JTB 8585") if len(hit) == 1 else (None, "")
        cands = [
            y
            for y in i18
            if mats25 & {m for m in (y["Matrículas"] or "").split(", ") if m}
        ]
        if len(cands) == 1:
            return cands[0], "matrícula"
        same_area = [
            y
            for y in cands
            if (y["Área total"] or "").split(" (")[0] == (x25["Área total"] or "")
        ]
        if len(same_area) == 1:
            return same_area[0], "matrícula + área"
        same_seller = [
            y
            for y in cands
            if y["Adquirido de"] and y["Adquirido de"] == x25["Adquirido de"]
        ]
        if len(same_seller) == 1:
            return same_seller[0], "matrícula + vendedor"
        return None, f"ambíguo: {len(cands)} candidatos"

    with open(DECISIONS, encoding="utf-8") as fh:
        decisions = json.load(fh)
    approved = decisions.get("itens", {})

    obs_by_key: dict[str, list[str]] = {}
    obs_company: list[str] = []
    seen_obs = set()
    for g in gs:
        if g["É observação?"] != "sim":
            continue
        cell = (g["Aba"], g["Célula"])
        text = " ".join(str(g["Valor / texto"] or "").split())
        if g["Comentário"]:
            text = f"{text[:80]}… [comentário: {' '.join(str(g['Comentário']).split())[:200]}]"
        if cell in OBS_SKIP:
            continue
        if cell in OBS_DUP:
            seen_obs.add(OBS_DUP[cell])
            continue
        if cell in OBS_COMPANY:
            obs_company.append(f"{cell[0]}!{cell[1]} ({OBS_COMPANY[cell]}): {text}")
            continue
        if cell not in OBS_MAP:
            raise SystemExit(f"observation cell not mapped: {cell}")
        seen_obs.add(cell)
        for key in OBS_MAP[cell]:
            obs_by_key.setdefault(key, []).append(f"{cell[0]}!{cell[1]}: {text}")
    missing = set(OBS_MAP) - seen_obs
    if missing:
        raise SystemExit(f"mapped cells not found in Phase 0: {missing}")

    def rdb(key: str) -> str | None:
        refs = (["RDB-1"] if key.startswith("05-2025") else []) + RDB_EXTRA.get(key, [])
        return ", ".join(refs) or None

    rows: list[dict] = []
    used18 = set()
    blocks = {}
    for x in i25:
        blocks.setdefault(x["Tabela/linha"].split("R")[0], len(blocks) + 1)
    block_name = {
        1: "Bens – imóveis urbanos",
        2: "Bens – imóveis rurais",
        3: "Bens – veículos",
    }
    sale_rigo = next(v for v in vendas if v["Comprador"] == "RIGO")
    farm_caveat = any(
        "avaliação independente" in str(c.get("Notas") or "") for c in comp.values()
    )
    reduction05 = f0.brl(
        f0.need(
            r"estima numa redução de R\$ (" + f0.MONEY + ")",
            " ".join(f0.docx_paragraphs(f0.SRC["inv2025_05"])),
            "redução",
        ).group(1)
    )

    for x in i25:
        k = x["Item"]
        key = f"05-2025-{k}"
        y, basis = pair_2018(x)
        if y:
            used18.add((y["Empresa"], y["Seq."]))
        g = dec(x["Valor (R$)"])
        r_exe = dec(exe_rafael.get(k, (None,))[0])
        r_grp = dec(grp_rafael.get(k, (None,))[0])
        r_ava = dec(ava_rafael.get(k, (None,))[0])
        diff = (r_exe - g) if r_exe is not None else None
        on_list = k in lista
        c0 = None
        if r_exe is not None:
            c0 = min(g, r_exe) if on_list else max(g, r_exe)
        cr = comp.get(k, {})
        has_conc = (
            bool(cr.get("Concordância expressa?"))
            and cr["Concordância expressa?"] != "não"
        )
        notes = []
        if has_conc:
            notes.append(
                f"concordância expressa ({cr['Concordância expressa?'].replace('sim, ', '')})"
            )
        if cr.get("Notas"):
            notes.append(cr["Notas"])
        notes += [f"p. {f['Página']}: {f['Texto']}" for f in faixas if f["Item"] == k]
        if diff is None:
            situ = "Não abordado por Rafael"
        elif diff != 0:
            situ = "Discorda"
        elif not has_conc and "avaliação independente" in str(cr.get("Notas") or ""):
            situ = "Concorda com ressalva"
        elif "FAZENDA SÃO CRISPIN" in x["Descrição"] and farm_caveat:
            # part of the same farm: Rafael's caveat (p. 13-14) and the
            # R$ 400.000 regularization cover it too (review of packet 1, A1)
            situ = "Concorda com ressalva"
        else:
            situ = "Concorda"
        extra = []
        if situ == "Concorda com ressalva" and has_conc:
            extra.append(
                f"Parte da Fazenda São Crispin: segue a decisão do item 05 (avaliação independente pedida por Rafael, p. 13–14; "
                f"redução de R$ {f0.br(reduction05)} de regularização dos docs 02 e 05). A concordância de Rafael ({cr['Concordância expressa?'].replace('sim, ', '')}) "
                "remete aos itens anteriores, inclusive o 05."
            )
        area = x["Área avaliada"] or x["Área total"]
        if k == "01":
            built = f0.need(r"de terreno: ([\d.,]+)m²", x["Descrição"], "720 m²").group(
                1
            )
            area = f"{x['Área total']} na matrícula; {built} m² avaliados neste item (o restante é o item 02)"
            extra.append(
                f"Matrícula 1.412 com {x['Área total']}: o item 01 avalia {built} m² de terreno e o item 02 avalia o restante (arq. 1 p. 8–9). "
                "A faixa de Rafael da p. 8 trata a 'área comercial premium de 1.800 metros quadrados' (matrícula inteira, inclui o terreno do item 02); "
                f"o valor executado de R$ {f0.br(r_exe)} não tem memória de cálculo."
            )
        if k == "02":
            extra.append(
                "Mesma matrícula 1.412 do item 01 (arq. 1 p. 8–9: 720 m² no item 01, 1.080 m² aqui)."
            )
            rate = f0.need(
                r"X R\$ (" + f0.MONEY + r"), ao metro", x["Descrição"], "02 R$/m²"
            ).group(1)
            m2_range = [
                f["Texto"]
                for f in faixas
                if f["Item"] == "01"
                and "por m²" in f["Texto"]
                and "esquina" in f["Texto"]
            ]
            if m2_range:
                extra.append(
                    f"Risco (C1): o grupo usa R$ {rate}/m² e Rafael concorda (p. 10), "
                    f"mas na avaliação do item 01 ele escreve: '{m2_range[0]}'"
                )
        if k == "05":
            sold = f0.need(r"\(-\)\s*([\d,]+)\s*ha", x["Descrição"], "83,47").group(1)
            extra.append(
                f"{sold} ha da mat. 00.760 vendidos aos Rigo em 2018 (doc 17 2018 T3R2; venda total de {f0.br(dec(sale_rigo['Área (ha)']))} ha "
                f"por R$ {f0.br(dec(sale_rigo['Valor (R$)']))}, T6R1). Rafael: valores 'parâmetros preliminares', recomenda avaliação independente (p. 13–14)."
            )
            p13 = " ".join(f0.pymupdf.open(f0.SRC["rafael_05"])[12].get_text().split())
            phrase = f0.need(
                r"tendo como resultado valores inferiores[^.]*\.", p13, "Rafael p. 13"
            ).group(0)
            extra.append(
                f"Rafael, p. 13: a avaliação por uso do solo teve '{phrase[len('tendo como resultado ') :]}'"
            )
        if k in dup_exe:
            extra.append(
                f"Rafael imprime este item duas vezes no balanço executado (p. {dup_exe[k]}); o total conta uma vez."
            )
        dec_ = approved.get(key, {})
        rows.append(
            {
                "empresa": "05 Hotel Princesa",
                "bloco": block_name[blocks[x["Tabela/linha"].split("R")[0]]],
                "item": k,
                "chave": key,
                "descricao": x["Descrição"],
                "matricula": x["Matrículas"],
                "area": area,
                "valor_grupo": g,
                "valor_antigo": dec(old25.get(k, {}).get("Antigo (R$)")),
                "existe_2018": "Sim" if y else "Não",
                "item_2018": f"{y['Empresa'][:2]} 2018 item {y['Item'] or '—'} (seq. {y['Seq.']})"
                if y
                else None,
                "base_par": basis,
                "descricao_2018": (y["Descrição"] or "")[:200] if y else None,
                "valor_2018": dec(y["Valor (R$)"]) if y else None,
                "area_2018": y["Área total"] if y else None,
                "sei_2018": y["Também listado pela SEI em 2018?"] if y else None,
                "existe_2025": f"Sim (doc 05 {x['Tabela/linha']})",
                "existe_rafael": "Sim",
                "doc_rafael": "HOTEL PRINCESA DOS VALES LTDA- BELISE LTDA.pdf",
                "valor_grupo_rafael": r_grp,
                "valor_ava": r_ava,
                "valor_rafael": r_exe,
                "obs_rafael": "; ".join(notes),
                "diferenca": diff,
                "diferenca_pct": (diff / g) if diff is not None and g else None,
                "situacao": situ,
                "na_lista": "Sim" if on_list else "Não",
                "valor_c0": c0,
                "valor_adotado": dec(dec_.get("valor_adotado")),
                "decisao": dec_.get("decisao", "Pendente"),
                "motivo": dec_.get("motivo"),
                "rdb": rdb(key),
                "obs_luis": "\n".join(obs_by_key.get(key, [])) or None,
                "situacao_obs": "Pendente" if obs_by_key.get(key) else None,
                "parte6": PARTE6.get(key),
                "nota_par": " ".join(extra) or None,
                "fonte_2025": f"doc 05 2025 {x['Tabela/linha']}",
                "fonte_2018": f"doc {'13' if y['Empresa'].startswith('05') else '17'} 2018 {y['Tabela/linha']}"
                if y
                else None,
                "fonte_rafael": "; ".join(
                    p
                    for p in (
                        f"grupo p. {grp_rafael[k][1]}" if k in grp_rafael else "",
                        f"avaliação executada p. {ava_rafael[k][1]}"
                        if k in ava_rafael
                        else "",
                        f"balanço executado p. {exe_rafael[k][1]}"
                        if k in exe_rafael
                        else "",
                        f"bens escolhidos p. {esc_rafael[k][1]}"
                        if k in esc_rafael
                        else "",
                        f"Diretrizes p. {lista[k]['Página Diretrizes']}"
                        if k in lista
                        else "",
                    )
                    if p
                ),
            }
        )

    # ---- 2018 rows with no pair in 2025 (doc 05)
    for y in i18:
        if (y["Empresa"], y["Seq."]) in used18:
            continue
        key = f"{y['Empresa'][:2]}-2018-s{int(y['Seq.']):02d}"
        loc_rows = [
            w
            for w in where
            if w["Empresa 2018"] == y["Empresa"] and w["Seq."] == y["Seq."]
        ]
        atual = next(w for w in loc_rows if w["Versão 2025"] == "atual")
        antigo = next(w for w in loc_rows if w["Versão 2025"] == "antigo")
        elsewhere = atual["Onde aparece em 2025 (doc, tabela/linha)"]
        shared = str(atual.get("Nota") or "").startswith("matrícula compartilhada")
        sold = (y["Situação"] or "").startswith("VENDIDA")
        if sold and y["Comprador"]:
            sit = "Saída em 2018 por venda (registrada no doc 17); destino (P4) a confirmar na etapa 6"
        elif elsewhere != "não encontrado":
            docs = sorted({e.split()[0] for e in elsewhere.split("; ")})
            sit = f"Existia em 2018 na 02; em 2025 no inventário de outra empresa (doc {', '.join(docs)})"
        else:
            sit = "Existia em 2018 e não consta em 2025 (destino a apurar)"
        if (
            elsewhere == "não encontrado"
            and antigo["Onde aparece em 2025 (doc, tabela/linha)"] == "não encontrado"
        ):
            loc_txt = "não encontrado nos inventários 2025 (atual e antigo)"
        else:
            loc_txt = f"atual: {elsewhere} ({atual['Área 2025'] or '—'}; R$ {atual['Valor 2025 (R$)'] or '—'})"
            if (
                antigo["Onde aparece em 2025 (doc, tabela/linha)"] != elsewhere
                or antigo["Valor 2025 (R$)"] != atual["Valor 2025 (R$)"]
            ):
                loc_txt += f" | antigo: {antigo['Onde aparece em 2025 (doc, tabela/linha)']} (R$ {antigo['Valor 2025 (R$)'] or '—'})"
        if shared:
            loc_txt = f"parcela vendida: não está em 2025 (a mesma matrícula aparece em {loc_txt}, que é outra parcela)"
        if sold and y["Comprador"]:
            sale = next(
                (
                    v
                    for v in vendas
                    if v["Comprador"].upper() in (y["Comprador"] or "").upper()
                ),
                None,
            )
            if sale:
                loc_txt += f" | venda 2018 a {sale['Comprador']}: R$ {f0.br(dec(sale['Valor (R$)']))} por {f0.br(dec(sale['Área (ha)']))} ha (doc 17 {sale['Tabela/linha']})"
        # Rafael, in the document of the company that holds the asset in 2025
        raf = None
        if elsewhere != "não encontrado" and not shared:
            doc_key = "inv2025_" + elsewhere.split()[0]
            if doc_key in RAFAEL_OTHER:
                mats = {m for m in (y["Matrículas"] or "").split(", ") if m}
                raf = rafael_other(doc_key, mats)
                if raf is None:
                    raise SystemExit(f"Rafael value not found for {key} in {doc_key}")
        owner = {"11": "sessão SEI", "08": "sessão S3 (Abastecedora e Transportadora)"}
        if y["Bloco da fazenda?"] == "sim":
            bloco = (
                "[2018] Agropecuária – fazenda (benfeitorias)"
                if not y["Matrículas"] and "Benfeitorias" in (y["Descrição"] or "")
                else "[2018] Agropecuária – Fazenda São Crispin (parcelas)"
            )
        elif not y["Matrículas"]:
            bloco = "[2018] Agropecuária – máquinas e veículos"
        else:
            bloco = "[2018] Agropecuária – outras áreas rurais"
        doc_other = (
            elsewhere.split()[0]
            if elsewhere != "não encontrado" and not shared
            else None
        )
        rows.append(
            {
                "empresa": y["Empresa"],
                "bloco": bloco,
                "item": y["Item"],
                "chave": key,
                "matricula": y["Matrículas"],
                "existe_2018": "Sim",
                "item_2018": f"{y['Empresa'][:2]} 2018 item {y['Item'] or '—'} (seq. {y['Seq.']})",
                "base_par": "sem par no doc 05 2025",
                "descricao_2018": (y["Descrição"] or "")[:200],
                "valor_2018": dec(y["Valor (R$)"]),
                "area_2018": y["Área total"],
                "sei_2018": y["Também listado pela SEI em 2018?"],
                "existe_2025": f"Não no doc 05; Sim no doc {doc_other} ({elsewhere})"
                if doc_other
                else "Não",
                "existe_rafael": f"Sim, no documento da outra empresa (item {raf['item']})"
                if raf
                else "Não",
                "doc_rafael": raf["doc"] if raf else None,
                "valor_grupo_rafael": raf["grupo"] if raf else None,
                "valor_rafael": raf["rafael"] if raf else None,
                "diferenca": (raf["rafael"] - raf["grupo"]) if raf else None,
                "diferenca_pct": ((raf["rafael"] - raf["grupo"]) / raf["grupo"])
                if raf and raf["grupo"]
                else None,
                "obs_rafael": (
                    f"Valor do bem em 2025 na outra empresa; a decisão é da {owner.get(doc_other, 'sessão dona do bem')}. "
                    + raf.get("concordancia", "")
                ).strip()
                if raf
                else None,
                "fonte_rafael": (
                    f"grupo p. {raf['pagina_grupo']}; "
                    + (
                        f"avaliação p. {raf['pagina_ava']}; "
                        if raf.get("pagina_ava")
                        else ""
                    )
                    + f"executado p. {raf['pagina']}"
                )
                if raf
                else None,
                "situacao": sit,
                "na_lista": "Não",
                "decisao": approved.get(key, {}).get(
                    "decisao", "Pendente (destino, etapa 6)"
                ),
                "motivo": approved.get(key, {}).get("motivo"),
                "localizacao_2025": loc_txt,
                "destino": approved.get(key, {}).get("destino"),
                "rdb": rdb(key),
                "obs_luis": "\n".join(obs_by_key.get(key, [])) or None,
                "situacao_obs": "Pendente" if obs_by_key.get(key) else None,
                "parte6": PARTE6.get(key),
                "fonte_2018": f"doc {'13' if y['Empresa'].startswith('05') else '17'} 2018 {y['Tabela/linha']}",
                "nota_par": " ".join(
                    n
                    for n in (
                        atual.get("Nota"),
                        (
                            f"Rafael avalia só a área da matrícula na outra empresa ({atual['Área 2025']}, item {raf['item']}); "
                            "a diferença de área está em P6-03."
                        )
                        if raf
                        and atual.get("Nota")
                        and "área" in str(atual.get("Nota"))
                        else None,
                    )
                    if n
                )
                or None,
            }
        )

    # ---- not valued by anyone (Parte 5): soy credits
    def sacks(origin: str, kind: str, cond=lambda r: True) -> Decimal:
        return sum(
            (
                Decimal(str(r["Sacas"]))
                for r in soy
                if r["Origem"] == origin and r["Tipo"] == kind and cond(r)
            ),
            Decimal(0),
        )

    def years(pred) -> str:
        return ", ".join(
            str(r["Ano"])
            for r in soy
            if r["Origem"] == "LORENZET" and r["Tipo"] == "safra" and pred(r)
        )

    def is_hotel(r):
        return str(r["Empresa anotada"]).startswith("HOTEL")

    def is_agro(r):
        return str(r["Empresa anotada"]).startswith("AGRO")

    rigo_cred = sacks("RIGO", "crédito declarado")
    rigo_parcel = re.search(
        r"([\d.]+) SACAS DE SOJA em MAIO DE 2019", sale_rigo["Texto"]
    )
    na_rows = [
        (
            "NA-SOJA-RIGO",
            "05 Hotel (docx 2025) / venda feita pela 02 em 2018 (doc 17)",
            f"Soja da venda Rigo (2018), safras 2021 e 2022: {f0.br(rigo_cred, 0)} sacas declaradas como crédito, sem valor em R$.",
            f"Sim (doc 17 2018 {sale_rigo['Tabela/linha']}: parcelas de {rigo_parcel.group(1) if rigo_parcel else '?'} sacas de 2019 a 2022 dentro do 'TOTAL RECEBIDO')",
            (
                "Tabela do docx marca 'FALTA RECEBER'; as OBSERVAÇÕES dizem 'ainda não foram vendidos' e registram venda de "
                f"{f0.br(sacks('RIGO', 'dedução declarada'), 0)} sacas em 2024 'para aquisição da escavadeira hidráulica'. "
                "Em 2018 essas parcelas foram abatidas do bruto como recebidas (ver P6-06). Achado E1 aprovado."
            ),
        ),
        (
            "NA-SOJA-LOR-2021-2023",
            "05 Hotel Princesa",
            f"Soja do arrendamento Lorenzet, safras {years(is_hotel)}: {f0.br(sacks('LORENZET', 'safra', is_hotel), 0)} sacas não vendidas ('HOTEL/OLEOPLAN'), sem valor em R$.",
            "Não",
            "Achado E1 aprovado.",
        ),
        (
            "NA-SOJA-LOR-2024",
            "02 São Crispin (docx) / 05 Hotel (contrato 9)",
            f"Soja do arrendamento Lorenzet, safra {years(is_agro)}: {f0.br(sacks('LORENZET', 'safra', is_agro), 0)} sacas não vendidas ('AGROPECUÁRIA/OLEPOPLAN'), sem valor em R$.",
            "Não",
            (
                "O docx atribui à Agropecuária (inventário R$ 0,00); o contrato 9 (2023–2026) tem o Hotel como arrendador (achado E1 aprovado). "
                "O grupo diz que o arrendamento Hotel × Agropecuária é 'pro forma' e a venda dos grãos é tributada no CNPJ da Agropecuária (arq. 1 p. 6)."
            ),
        ),
    ]
    for key, emp, desc, in18, note in na_rows:
        rows.append(
            {
                "empresa": emp,
                "bloco": "Não abordado (Parte 5) – créditos em soja",
                "chave": key,
                "descricao": desc,
                "existe_2018": in18,
                "existe_2025": "Citado sem valor (docx 'áreas vendidas', OBSERVAÇÕES)",
                "existe_rafael": "Não",
                "obs_rafael": "Rafael não trata a soja.",
                "situacao": "Não abordado (sem valor nas duas avaliações)",
                "na_lista": "Não",
                "decisao": approved.get(key, {}).get("decisao", "Pendente (Parte 5)"),
                "motivo": approved.get(key, {}).get("motivo"),
                "rdb": rdb(key),
                "obs_luis": "\n".join(obs_by_key.get(key, [])) or None,
                "situacao_obs": "Pendente" if obs_by_key.get(key) else None,
                "parte6": PARTE6.get(key),
                "fonte_2025": "docx 'HOTEL PRINCESA … ÁREA DE TERRAS VENDIDAS', OBSERVAÇÕES",
                "nota_par": note,
            }
        )

    # ---- Phase 0 findings about a Hotel item (DIVERGE checks and warnings)
    confs = [r for r in read(wb, "Conferências") if r["Resultado"] == "DIVERGE"]
    warns = read(wb, "Avisos")
    for r in rows:
        if not r["chave"].startswith("05-2025"):
            continue
        k = r["item"]
        found = [
            f"Fase 0: {c['Conferência']} — {c['Nota'] or ''}".rstrip(" —")
            for c in confs
            if re.search(rf"\b2025 item {k}\b", c["Conferência"])
        ]
        found += [
            f"Fase 0 ({w['Assunto']}): {w['Texto']}"
            for w in warns
            if re.search(rf"\bitem {k}\b", str(w["Fonte"]))
            or f"item {k} do Hotel" in str(w["Texto"])
        ]
        if found:
            r["nota_par"] = " ".join(x for x in [r.get("nota_par"), *found] if x)

    # ---- checks
    hotel_rows = [r for r in rows if r["chave"].startswith("05-2025")]
    total25 = sum(r["valor_grupo"] for r in hotel_rows)
    total_raf = sum(r["valor_rafael"] for r in hotel_rows)
    problems = []
    if total25 != dec(grp_rafael["TOTAL"][0]):
        problems.append(f"total grupo {total25}")
    if total_raf != dec(exe_rafael["TOTAL"][0]):
        problems.append(f"total Rafael {total_raf}")
    if len(used18) + sum(1 for r in rows if "-2018-s" in r["chave"]) != len(i18):
        problems.append("2018 rows not all accounted for")
    if sorted(lista) != sorted(r["item"] for r in rows if r.get("na_lista") == "Sim"):
        problems.append("Rafael list mismatch")
    if problems:
        raise SystemExit("; ".join(problems))

    t11 = " ".join(f0.docx_paragraphs(f0.OTHER_2025["inv2025_11"]))
    sei_reg = f0.need(
        r"(Neste rol de custas[^.]*?adquiridas por contrato.*?custo aproximado de R\$ ("
        + f0.MONEY
        + r"))",
        t11,
        "SEI regularização",
    )
    write(rows, bases, total25, total_raf, tail, decisions, obs_company, sei_reg)
    by_sit: dict[str, int] = {}
    for r in rows:
        by_sit[r["situacao"]] = by_sit.get(r["situacao"], 0) + 1
    print(f"Fase 0: {tail}")
    print(
        f"linhas: {len(rows)}; 2025: {len(i25)}; 2018 casados: {len(used18)}; 2018 sem par: {sum(1 for r in rows if '-2018-s' in r['chave'])}; não abordados: {len(na_rows)}"
    )
    for k, v in by_sit.items():
        print(f"  {v:3} {k}")
    print(
        f"grupo 2025: {total25}; Rafael: {total_raf}; diferença: {total_raf - total25}"
    )
    for r in rows:
        if r.get("valor_rafael") is not None:
            print(
                f"  {r['chave']}: grupo {r.get('valor_grupo') or r.get('valor_grupo_rafael')} Rafael {r['valor_rafael']} [{r['situacao'][:40]}] C0 {r.get('valor_c0')}"
            )


COLS = [
    ("empresa", "Empresa", 16),
    ("bloco", "Bloco", 26),
    ("item", "Nº item", 7),
    ("chave", "Chave", 20),
    ("descricao", "Descrição (2025)", 50),
    ("matricula", "Matrícula", 18),
    ("area", "Área (2025)", 22),
    ("valor_grupo", "Valor grupo 2025 (R$)", 16),
    ("valor_antigo", "Valor grupo 2025, versão antiga (R$)", 16),
    ("existe_2018", "Existe em 2018?", 12),
    ("item_2018", "Item 2018", 22),
    ("base_par", "Base do par 2018", 18),
    ("area_2018", "Área 2018", 14),
    ("valor_2018", "Valor 2018 (R$)", 16),
    ("descricao_2018", "Descrição 2018", 50),
    ("sei_2018", "Listado pela SEI em 2018?", 10),
    ("existe_2025", "Existe em 2025 (grupo)?", 24),
    ("existe_rafael", "Existe no Estudo 2 (Rafael)?", 14),
    ("doc_rafael", "Documento de Rafael", 24),
    ("valor_grupo_rafael", "Valor do grupo segundo Rafael (R$)", 16),
    ("valor_ava", "Valor na 'avaliação executada' de Rafael (R$)", 16),
    ("valor_rafael", "Valor avaliado por Rafael, balanço executado (R$)", 16),
    ("obs_rafael", "Observação de Rafael", 50),
    ("diferenca", "Diferença Rafael − grupo (R$)", 16),
    ("diferenca_pct", "Diferença %", 9),
    ("situacao", "Situação", 30),
    ("na_lista", "Na lista de Rafael (virá para nós)?", 10),
    ("valor_c0", "Valor pelo C0 (referência)", 16),
    ("valor_adotado", "Valor adotado (R$)", 16),
    ("decisao", "Decisão (P5)", 22),
    ("motivo", "Motivo da decisão", 40),
    ("localizacao_2025", "Localização em 2025 (fato, por matrícula)", 50),
    ("destino", "Destino 2018→2025", 30),
    ("rdb", "Registro data-base", 12),
    ("obs_luis", "Observação do Luís (planilha 2_0)", 50),
    ("situacao_obs", "Situação da observação", 12),
    ("parte6", "Candidato Parte 6", 12),
    ("nota_par", "Nota", 50),
    ("fonte_2025", "Fonte 2025", 22),
    ("fonte_2018", "Fonte 2018", 22),
    ("fonte_rafael", "Fonte Rafael", 30),
]
WRAP = {
    "descricao",
    "descricao_2018",
    "obs_rafael",
    "obs_luis",
    "localizacao_2025",
    "nota_par",
    "area",
    "existe_2025",
    "fonte_rafael",
}


def write(
    rows, bases, total25, total_raf, tail, decisions, obs_company, sei_reg
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mestre"
    for j, (_, title, width) in enumerate(COLS, 1):
        c = ws.cell(1, j, title)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[get_column_letter(j)].width = width
    fills = {
        "Concorda": "E2EFDA",
        "Concorda com ressalva": "FFF2CC",
        "Discorda": "FCE4D6",
    }
    sit_col = [c[0] for c in COLS].index("situacao") + 1
    for i, r in enumerate(rows, 2):
        for j, (key, _, _) in enumerate(COLS, 1):
            v = r.get(key)
            cell = ws.cell(i, j)
            if isinstance(v, Decimal):
                cell.value = float(v)
                cell.number_format = "0.0%" if key == "diferenca_pct" else "#,##0.00"
            else:
                cell.value = v
            cell.alignment = Alignment(wrap_text=key in WRAP, vertical="top")
        if r["situacao"] in fills:
            ws.cell(i, sit_col).fill = PatternFill(
                "solid", fgColor=fills[r["situacao"]]
            )
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{len(rows) + 1}"

    rs = wb.create_sheet("Resumo")
    c0_total = sum(r["valor_c0"] for r in rows if r.get("valor_c0") is not None)
    lines = [
        ("Totais", None),
        ("Bens 2025 do Hotel pelo grupo (doc 05)", total25),
        ("Bens 2025 do Hotel por Rafael (balanço executado)", total_raf),
        ("Diferença Rafael − grupo", total_raf - total25),
        ("Soma pelo critério C0 (só referência; decisões pendentes)", c0_total),
        ("Agropecuária 2025 (doc 02)", Decimal(0)),
        (
            next(
                (
                    f"Benchmark 2018 das duas empresas — decisão do Luís ({r['data']}): {r['texto']}"
                    for r in decisions.get("regras", [])
                    if r.get("regra") == "BASE-2018"
                ),
                "Benchmark 2018 das duas empresas (escolha pendente do Luís)",
            ),
            None,
        ),
    ]
    for k, v in bases.items():
        if k.startswith(("A.", "B.", "   B", "B − A")):
            lines.append((k.strip(), dec(v)))
    lines.append(("Linhas da aba Mestre por situação", None))
    counts: dict[str, int] = {}
    for r in rows:
        counts[r["situacao"]] = counts.get(r["situacao"], 0) + 1
    lines += [(k, v) for k, v in counts.items()]
    lines.append(("Observações do Luís sobre a empresa como um todo", None))
    lines += [(o, None) for o in obs_company] or [("(nenhuma)", None)]
    lines += [
        ("Conferências da Fase 0", None),
        (tail, None),
        ("Decisões aprovadas em decisoes.json", len(decisions.get("itens", {}))),
    ]
    for i, (a, b) in enumerate(lines, 1):
        rs.cell(i, 1, a).font = Font(
            bold=b is None and not a.startswith(("Resumo!", "total", "(nenhuma)"))
        )
        if isinstance(b, Decimal):
            rs.cell(i, 2, float(b)).number_format = "#,##0.00"
        elif b is not None:
            rs.cell(i, 2, b)
    rs.column_dimensions["A"].width = 90
    rs.column_dimensions["B"].width = 20

    ps = wb.create_sheet("Passivos e deduções")
    hdr = [
        "Item",
        "Natureza",
        "Valor (R$)",
        "Tratamento no documento",
        "Fonte",
        "Tratamento de Rafael",
    ]
    for j, h in enumerate(hdr, 1):
        ps.cell(1, j, h).font = Font(bold=True)
    t05 = " ".join(f0.docx_paragraphs(f0.SRC["inv2025_05"]))
    reduction = f0.brl(
        f0.need(r"estima numa redução de R\$ (" + f0.MONEY + ")", t05, "redução").group(
            1
        )
    )
    ps.append(
        [
            "Fazenda São Crispin (itens 05 e 06)",
            (
                "Regularização: georreferenciamento, ITBI, escrituras e registro (estimativa do grupo). Escrituras 'que estão em nome da empresa "
                "HOTEL PRINCESA DOS VALES LTDA. e de TERCEIROS, POR CONTRATO', a transmitir para a Agropecuária."
            ),
            float(reduction),
            "Descrita no texto dos docs 02 e 05 como redução do valor da fazenda; não abatida do TOTAL GERAL do doc 05.",
            "doc 05 2025, OBSERVAÇÕES; doc 02 2025, mesmo texto",
            "Repete o texto do grupo (PDF Hotel p. 4–6; PDF Agropecuária p. 1–3), sem valor próprio.",
        ]
    )
    ps.append(
        [
            "Referência: SEI (doc 11), custas de regularização de áreas rurais",
            " ".join(sei_reg.group(1).split()),
            float(f0.brl(sei_reg.group(2))),
            (
                "Rol de custas da SEI, para escriturar em nome da própria SEI; não diz quais áreas. Nota do Luís (GERAL H882): 'Quais áreas: "
                "Fazenda São Crispin? Áreas de compensação?'. A planilha 2_0 já abate este valor (GERAL E882). O doc 11 tem duas áreas "
                "'por contrato', não registradas e com compensação ambiental (lote 47 Piccolo e Scalco, T9R4–T9R5). Ver P6-08 e Q11."
            ),
            "doc 11 2025, OBSERVAÇÕES",
            "—",
        ]
    )
    for rr in (2, 3):
        ps.cell(rr, 3).number_format = "#,##0.00"
    for j, w in enumerate((34, 60, 16, 60, 40, 50), 1):
        ps.column_dimensions[get_column_letter(j)].width = w

    lg = wb.create_sheet("Legenda")
    legend = [
        (
            "Mestre",
            "Uma linha por bem do inventário 2025 do Hotel (8; a Agropecuária não tem bens em 2025), mais os bens dos inventários 2018 do Hotel e da Agropecuária sem par em 2025 e os créditos em soja que ninguém avaliou.",
        ),
        (
            "Bloco",
            "Bloco do inventário 2025 (linhas 2025) ou do inventário 2018 (linhas [2018]).",
        ),
        (
            "Base do par 2018",
            "Par 2018 × 2025 feito por matrícula + área, vendedor ou placa, nunca por semelhança de texto. A matrícula 1.412 tem dois itens (01 e 02) e a 11.090 duas parcelas em 2018: nesses casos a área decide.",
        ),
        (
            "Valor 2018",
            "Valor do item no inventário 2018 (referência; 2018 é benchmark de existência, não de valor). Em 2018 a fazenda era avaliada pelos 491 ha inteiros, inclusive os 83,47 ha vendidos.",
        ),
        (
            "Existe em 2025 (grupo)?",
            "Para linhas 2018 sem par: se o bem aparece em outro inventário 2025 (doc 08 Transportadora, doc 11 SEI), com tabela/linha.",
        ),
        (
            "Valor do grupo segundo Rafael",
            "Valor que Rafael transcreve no 'balanço apresentado pelo grupo'. Para bens de 2018 que estão em outra empresa em 2025, vem do documento de Rafael dessa empresa.",
        ),
        (
            "Valor na 'avaliação executada'",
            "Seção do PDF do Hotel (p. 7–15) em que Rafael repete o valor do grupo e cita faixas de mercado.",
        ),
        (
            "Valor avaliado por Rafael",
            "Valor do 'balanço patrimonial executado por Valter' (PDF do Hotel p. 16–19; Quadro_Resumo das áreas rurais; balanço executado da Transportadora).",
        ),
        (
            "Situação",
            "Concorda = valor de Rafael igual ao do grupo, com concordância expressa; Concorda com ressalva = valor igual, sem bloco de concordância logo após o item (pode ser citado nas remissões de outros itens) e com ressalva escrita de Rafael; Discorda = diferente. Para 2018 sem par: saída por venda em 2018, em outra empresa em 2025, ou sem destino.",
        ),
        (
            "Valor pelo C0 (referência)",
            "Menor valor se o bem está na lista de Rafael, maior se fica na empresa (P10). Só referência: a decisão é tomada nos pacotes. Bens em outra empresa: a decisão é da sessão dona do bem.",
        ),
        (
            "Decisão (P5) / Motivo / Valor adotado / Destino",
            "Preenchidos a partir de decisoes.json, só com decisões aprovadas pelo Luís.",
        ),
        (
            "Localização em 2025",
            "Onde a mesma matrícula aparece nos inventários 2025 (atual e antigo). Fato de localização; o destino (P4) é decidido na etapa 6.",
        ),
        (
            "Registro data-base",
            "Linha de registro-data-base.md desta pasta. RDB-1 vale para todos os bens do Hotel.",
        ),
        ("Candidato Parte 6", "Linha de parte6-candidatos.md desta pasta."),
        (
            "Observação do Luís",
            "Células de observação da planilha 2_0 (abas GERAL, Resumo, 2.São Crispin, 5.Hotel Princesa, 11.Empreendimentos). As abas das empresas repetem a coluna H da GERAL: entram uma vez. Notas sobre a empresa como um todo estão na aba Resumo. A autoria das notas em vermelho da coluna F da GERAL (F943–F987) não está identificada.",
        ),
        (
            "Premissas",
            "1/6 em todas as instâncias; data-base 31.12.2024; 2018 é benchmark de existência; sem reavaliação própria; critério C0 (P10).",
        ),
    ]
    for i, (a, b) in enumerate(legend, 1):
        lg.cell(i, 1, a).font = Font(bold=True)
        lg.cell(i, 2, b).alignment = Alignment(wrap_text=True, vertical="top")
    lg.column_dimensions["A"].width = 32
    lg.column_dimensions["B"].width = 120
    wb.save(OUT)


if __name__ == "__main__":
    main()
