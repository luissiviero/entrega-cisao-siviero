"""Money from the 2018 sales of São Crispin land: contracted, received,
withdrawn, paid. Writes `levantamento-vendas-sao-crispin.md`.

Sources (all read here):
  - 2018 inventory of the Agropecuária (doc 17), sales table and items;
  - 2025 docx "áreas vendidas" (FALTA RECEBER marks, soy observations);
  - SEI paid bills Jul/2023-Oct/2024 (`Contas_Pagas_2023-07_a_2024-10.xlsx`,
    transcribed from the monthly PDF reports, which are scans);
  - answers to Dr. Pablo, file 1 (p. 5-6).

v2 (28/10/2026) applies `_diario/revisao-vendas-sao-crispin.md` (A1-A11).

Run: PYTHONIOENCODING=utf-8 python -W ignore vendas_sao_crispin.py
"""

from __future__ import annotations

import os
import re
import sys
from collections import defaultdict
from decimal import Decimal

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fase0_extracao as f0

OUT = os.path.join(f0.OUT_DIR, "levantamento-vendas-sao-crispin.md")
CONTAS = os.path.join(f0.ROOT, r"01_fontes\outros\Contas_Pagas_2023-07_a_2024-10.xlsx")
M = f0.MONEY
PARTNERS = ("VALTER", "PEDRO", "ADELIA", "IRINEU", "RAUL", "ELCIO")


def money(v) -> str:
    return "—" if v is None else f"R$ {f0.br(Decimal(str(v)))}"


def sacks(v) -> str:
    return f0.br(Decimal(str(v)), 0)


def pdf_text(path: str, page: int) -> str:
    return " ".join(f0.pymupdf.open(path)[page - 1].get_text().split())


def main() -> None:
    sales, totals = f0.sale_rows(f0.SRC["inv2018_17"])
    sales25, _ = f0.sale_rows(f0.SRC["inv2025_areas"])
    falta = {s["comprador"]: s["falta_receber"] for s in sales25}
    received_2018 = next(t["valor"] for t in totals if "RECEBIDO" in t["rotulo"])
    by = {s["comprador"]: s for s in sales}

    # ---- 1. what was sold, installment by installment
    inst = []  # (buyer, when, what, value, sacks)
    rigo = by["RIGO"]["texto"]
    price = f0.brl(
        f0.need(r"considerado foi de R\$ (" + M + ")", rigo, "preço").group(1)
    )
    inst.append(
        (
            "Rigo",
            "2018",
            "à vista",
            f0.brl(
                f0.need(r"R\$ (" + M + r") – À VISTA", rigo, "Rigo à vista").group(1)
            ),
            None,
        )
    )
    for q, year in re.findall(r"([\d.]+) SACAS DE SOJA em MAIO DE (\d{4})", rigo):
        inst.append(
            ("Rigo", f"maio/{year}", f"{q} sacas de soja", f0.num(q) * price, f0.num(q))
        )
    pel = by["PELLEGRINI"]["texto"]
    # the price = land + 41 head of cattle sold with it; paid "à vista" + "JUNHO/2019"
    for v, what in re.findall(r"R\$ (" + M + r") – (À VISTA|JUNHO/2019)", pel):
        inst.append(
            (
                "Pellegrini",
                "jun/2019" if what == "JUNHO/2019" else "2018",
                "à vista" if what == "À VISTA" else "parcela",
                f0.brl(v),
                None,
            )
        )
    cattle = f0.brl(
        f0.need(r"R\$ (" + M + r") – 41 cabeças de gado", pel, "gado").group(1)
    )
    land_pel = f0.brl(
        f0.need(r"R\$ (" + M + r") – [\d,]+ ha", pel, "terra Pellegrini").group(1)
    )
    ang = by["ANGELI"]
    if "À VISTA" in ang["texto"]:
        inst.append(("Angeli", "2018", "à vista", ang["valor"], None))
    contracted = sum(i[3] for i in inst)
    for b, s in by.items():
        got = sum(i[3] for i in inst if i[0].upper() == b)
        if abs(got - s["valor"]) > Decimal("0.01"):
            raise SystemExit(f"installments of {b} = {got}, sale = {s['valor']}")
    if contracted != received_2018 or land_pel + cattle != by["PELLEGRINI"]["valor"]:
        raise SystemExit("sales do not add up")
    angeli_place = "Cotiporã" if "Cotiporã" in ang["texto"] else "?"
    land_crispin = by["RIGO"]["valor"] + land_pel
    rigo25 = next(s for s in sales25 if s["comprador"] == "RIGO")["texto"]
    total_sacks = f0.need(r"= ([\d.]+) SC SOJA", rigo25, "118.100").group(1)
    vista_sacks = f0.need(r"equivalente a ([\d.]+) SACAS", rigo25, "37.500").group(1)
    cash_2018 = sum(i[3] for i in inst if i[1] == "2018")
    cash_2019 = sum(i[3] for i in inst if i[0] == "Pellegrini" and i[1] == "jun/2019")
    soy_inst = [i for i in inst if i[4]]
    falta_years = [y for y in falta.get("RIGO", "").split(", ") if y]
    soy_open = [i for i in soy_inst if i[1][-4:] in falta_years]
    soy_done = [i for i in soy_inst if i[1][-4:] not in falta_years]

    soy = f0.soy_credits(f0.SRC["inv2025_areas"])
    sold = next(r for r in soy if r["tipo"] == "vendida")
    credit_rigo = next(
        r for r in soy if r["origem"] == "RIGO" and r["tipo"] == "crédito declarado"
    )
    lor = [r for r in soy if r["origem"] == "LORENZET" and r["tipo"] == "safra"]
    lor_same = [r for r in lor if r["sacas"] == -sold["sacas"]]

    # inventory 2018: is any cash or receivable listed as an asset?
    items18, _ = f0.item_rows(f0.SRC["inv2018_17"], "02", "2018")
    cash_items = [
        x
        for x in items18
        if re.search(
            r"dinheiro|caixa|aplica|a receber|crédito|sacas",
            x["descricao"],
            re.IGNORECASE,
        )
    ]

    p5 = pdf_text(f0.SRC_LATER["respostas_arq01"], 5)
    pablo = f0.need(
        r"ainda estamos recebendo o pagamento de áreas vendidas por R\$ [\d.,]+/hectare",
        p5,
        "Dr. Pablo",
    ).group(0)
    p6 = pdf_text(f0.SRC_LATER["respostas_arq01"], 6)
    fluxo = f0.need(
        r"depositadas e saldadas conforme fluxo financeiro necessário para a empresa Siviero Empreendimentos Imobiliários Ltda\.",
        p6,
        "fluxo",
    ).group(0)
    cnpj = f0.need(
        r"tributando a venda dos grãos na inscrição estadual pelo CNPJ da Agropecuária",
        p6,
        "CNPJ",
    ).group(0)

    # ---- SEI paid bills
    wb = openpyxl.load_workbook(CONTAS, data_only=True)
    rows = [r for r in wb["Consolidado"].iter_rows(min_row=2, values_only=True) if r[0]]

    def hist(r) -> str:
        return f"{r[3]} | {r[8]}"

    tax = [
        r
        for r in rows
        if "SOJA" in str(r[8])
        and re.search(r"FUNRURAL|IRPJ|CSLL|CONTRIBUICAO|IMPOSTO DE RENDA", str(r[3]))
    ]
    soy_sales: dict[Decimal, dict] = {}
    for r in tax:
        obs = str(r[8])
        b = re.search(r"R\$ (" + M + r") X", obs)
        q = re.search(r"([\d.]+) SC SOJA VENDIDOS X", obs)
        if b:
            base = f0.brl(b.group(1))
        elif q and "1,50%" in str(r[3]):
            base = (Decimal(str(r[7])) / Decimal("0.015")).quantize(Decimal("0.01"))
        else:
            continue
        s = soy_sales.setdefault(
            base,
            {
                "base": base,
                "taxes": [],
                "obs": obs,
                "mes": r[0].strftime("%m/%Y"),
                "sacas": f0.num(q.group(1)) if q else None,
            },
        )
        s["taxes"].append(r)
    for s in soy_sales.values():
        s["pf"] = all("1,50%" in str(t[3]) for t in s["taxes"])
        fun = sum(Decimal(str(t[7])) for t in s["taxes"] if "FUNRURAL" in str(t[3]))
        s["rate"] = (fun / s["base"] * 100).quantize(Decimal("0.01"))
        if "17.544" in s["obs"]:
            s["origem"] = "Rigo (parcela de 2021)"
        elif any("LORENZET" in str(t[8]) for t in s["taxes"]):
            s["origem"] = "aluguel Lorenzet"
        else:
            s["origem"] = "não identificada"
        if s["sacas"]:
            s["preco"] = (s["base"] / s["sacas"]).quantize(Decimal("0.01"))
    rigo_sale = next(s for s in soy_sales.values() if s["origem"].startswith("Rigo"))
    lorenzet_named = [
        s for s in soy_sales.values() if s["origem"] == "aluguel Lorenzet"
    ]
    pj_other = [
        s
        for s in soy_sales.values()
        if not s["pf"] and s["origem"] == "não identificada"
    ]

    exc = [r for r in rows if "PESA/CAT" in str(r[3])]
    exc_nf = Decimal(str(exc[0][6]))
    fin = [r for r in exc if re.search(r"P\.\d+/48", str(r[8]))]
    fin_total = (
        f0.brl(f0.need(r"R\$ (" + M + ")", str(fin[0][8]), "financiamento").group(1))
        if fin
        else None
    )
    imovel = [r for r in rows if "MARIA TEREZA" in str(r[3])]
    imovel_nf = Decimal(str(imovel[0][6]))
    imovel_paid = sum(Decimal(str(r[7])) for r in imovel)
    imovel_soy = sum(Decimal(str(r[7])) for r in imovel if "SOJA TITTON" in str(r[8]))
    farm = [r for r in rows if re.search(r"CRISP", hist(r)) and "SOJA" not in str(r[8])]
    itr_ha = next(
        (m.group(1) for r in farm if (m := re.search(r"CRISPIN-(\d+) HA", str(r[8])))),
        None,
    )

    # partner withdrawals (classification by the text of each bill)
    wd: dict[str, dict[str, Decimal]] = defaultdict(lambda: defaultdict(Decimal))
    for r in rows:
        h, v = hist(r), Decimal(str(r[7]))
        who = next(
            (p for p in PARTNERS if re.search(p, str(r[3]), re.IGNORECASE)), None
        )
        if re.match(r"PRO LABORE", str(r[3])):
            wd["Pró-labore"][who] += v
        elif who and "DIVISAO DE ALUGUEIS" in h:
            wd["Divisão de aluguéis"][who] += v
        elif who and "DIVISAO DE LUCROS" in h:
            wd["Divisão de lucros (ano 2023)"][who] += v
        elif who and re.search(r"DEV(OLUCAO)? EMPRESTIMO", h):
            wd["Devolução de empréstimo ao sócio"][who] += v
        elif "RAUL SIVIERO+PAGTO CONTAS INDIVIDUAIS" in h:
            wd['Hotel Araçá: "Raul Siviero + pagto contas individuais"']["RAUL"] += v
    wd_total = sum(sum(d.values()) for d in wd.values())
    first, last = min(r[0] for r in rows), max(r[0] for r in rows)

    # ---- write
    L: list[str] = []
    w = L.append
    w(
        "# Vendas de terras da Fazenda São Crispin (2018): o que entrou, o que saiu e para onde foi"
    )
    w("")
    w(
        "Pedido do Luís em 28/10/2026. Números gerados por `_scripts/vendas_sao_crispin.py` a partir das fontes; revisado por revisor independente (`_diario/revisao-vendas-sao-crispin.md`)."
    )
    w("")
    w("## Resumo")
    w("")
    w("| Pergunta | O que as fontes mostram | O que falta |")
    w("|---|---|---|")
    w(
        f"| Quanto a empresa recebeu | Vendas de {money(contracted)} em 2018. Em dinheiro: {money(cash_2018)} à vista em 2018 e {money(cash_2019)} em jun/2019. "
        f"Em soja (Rigo): {sacks(sum(i[4] for i in soy_done))} sacas de {' e '.join(i[1][-4:] for i in soy_done)} tratadas como entregues; "
        f"das {sacks(sum(i[4] for i in soy_open))} de {' e '.join(falta_years)}, {sacks(-sold['sacas'])} foram vendidas em 2024 e {sacks(credit_rigo['sacas'])} continuam em aberto. "
        "| Nenhum recibo, extrato ou lançamento das entradas. |"
    )
    w(
        f"| Quanto já foi retirado | Nenhuma fonte liga o dinheiro das vendas a retiradas. No único período com registros (contas pagas da SEI, {first.strftime('%m/%Y')} a {last.strftime('%m/%Y')}), "
        f"os sócios receberam {money(wd_total)} em pró-labore, aluguéis, lucros e devolução de empréstimos, pagos pelo caixa comum da SEI, para onde o grupo diz que vai a receita da fazenda. "
        "| Registros de 2018 a meados de 2023; em nome de quem saíram as notas de venda de soja. |"
    )
    w(
        f"| Quanto já foi pago com esse dinheiro | Só um uso tem ligação escrita: {sacks(-sold['sacas'])} sacas dos Rigo (cerca de {money(rigo_sale['base'])}) usadas na compra da escavadeira da SEI ({money(exc_nf)}). "
        f"| Destino dos {money(cash_2018 + cash_2019)} em dinheiro e da soja de {' e '.join(i[1][-4:] for i in soy_done)}. |"
    )
    w("")
    w("## 1. O que foi vendido")
    w("")
    w(
        f"Fonte: inventário 2018 da Agropecuária, quadro de vendas (repetido no documento de 2025 das áreas vendidas). O total de **{money(contracted)}** se divide assim:"
    )
    w("")
    w(
        f"- terra da Fazenda São Crispin: **{money(land_crispin)}** (Rigo {money(by['RIGO']['valor'])} + Pellegrini {money(land_pel)});"
    )
    w(
        f"- 41 cabeças de gado vendidas junto com a terra aos Pellegrini: {money(cattle)};"
    )
    w(
        f"- terra em {angeli_place} (lote 24, fora da fazenda) vendida a Angeli: {money(ang['valor'])}."
    )
    w("")
    w(
        f"A venda aos Rigo foi fechada em soja: {total_sacks} sacas no total, das quais o pagamento à vista equivale a {vista_sacks} (documento de 2025). "
        f"O valor em reais usa R$ {f0.br(price)} a saca, cotação de out/2018; é referência, não dinheiro recebido."
    )
    w("")
    w(
        "| Comprador | Quando | O quê | Valor de referência | O que as fontes dizem depois |"
    )
    w("|---|---|---|---|---|")
    for buyer, when, what, value, q in inst:
        if not q:
            note = "Nenhuma fonte posterior cita este pagamento."
        elif when[-4:] in falta_years:
            note = 'Documento de 2025: "(FALTA RECEBER)"; as observações do mesmo documento dizem que a soja "ainda não foi vendida".'
            if when[-4:] == "2021":
                note += f" Em 2024 o grupo vendeu {sacks(-sold['sacas'])} sacas deste lote (seção 4)."
        else:
            note = 'Documento de 2025 não marca "falta receber"; as observações dizem que só duas safras não foram vendidas, então esta foi. Destino não informado.'
        w(f"| {buyer} | {when} | {what} | {money(value)} | {note} |")
    w(f"| **Total** | | | **{money(contracted)}** | |")
    w("")
    w("## 2. Quanto já foi recebido")
    w("")
    w(
        f'- **Em dinheiro:** {money(cash_2018)} à vista em 2018 (Rigo, Pellegrini e Angeli) e {money(cash_2019)} da Pellegrini em jun/2019, ao todo {money(cash_2018 + cash_2019)}. Nenhuma fonte comprova a entrada; o inventário de 2018 só abate o total como "recebido".'
    )
    w(
        f"- **Em soja (Rigo):** as parcelas de {' e '.join(i[1][-4:] for i in soy_done)} ({sacks(sum(i[4] for i in soy_done))} sacas) não aparecem como pendentes. "
        f"Das parcelas de {' e '.join(falta_years)} ({sacks(sum(i[4] for i in soy_open))} sacas), pelo menos {sacks(-sold['sacas'])} foram vendidas pelo grupo em 2024; "
        f"as outras {sacks(credit_rigo['sacas'])} continuam declaradas como crédito, sem valor. Não se sabe se essa soja foi entregue e está em depósito ou se a Rigo ainda deve: o documento diz as duas coisas."
    )
    if lor_same:
        w(
            f"- Coincidência a esclarecer: {sacks(-sold['sacas'])} sacas é também o número da safra Lorenzet de {lor_same[0]['ano']} no mesmo documento."
        )
    w(
        f"- **O inventário de 31/12/2018 não lista dinheiro, aplicação nem parcelas a receber das vendas como bem da Agropecuária** "
        f"({len(cash_items)} item desse tipo entre os {len(items18)} itens); o total de {money(received_2018)} é só abatido. Já no fim de 2018 o dinheiro das vendas não aparecia na empresa."
    )
    w(
        f'- Na pergunta do Dr. Pablo ao grupo (arq. 1, p. 5) está escrito: "{pablo}", o preço por hectare da venda aos Rigo. O grupo não respondeu esse ponto (p. 6). O Luís informou em 28/10/2026 que o lado dele não recebe nada dos Rigo diretamente; portanto, se parcelas ainda estão sendo recebidas, são recebidas pelo grupo (empresa ou outros sócios).'
    )
    w("")
    w("## 3. Quanto já foi retirado pelos sócios")
    w("")
    w("- **Nenhuma fonte liga o dinheiro das vendas a retiradas de sócios.**")
    w(
        f'- O grupo diz que as receitas da fazenda são "{fluxo}" (arq. 1, p. 6). Ou seja, o dinheiro da fazenda vai para o caixa comum da SEI, de onde saem as retiradas.'
    )
    w(
        f"- No único período com registros, as contas pagas da SEI ({first.strftime('%m/%Y')} a {last.strftime('%m/%Y')}) mostram estes pagamentos a sócios (classificação pelo texto de cada lançamento):"
    )
    w("")
    w("| Tipo | " + " | ".join(p.title() for p in PARTNERS) + " | Total |")
    w("|---|" + "---|" * (len(PARTNERS) + 1))
    order = ["Pró-labore", "Divisão de aluguéis", "Divisão de lucros (ano 2023)", "Devolução de empréstimo ao sócio"]
    for kind, d in sorted(wd.items(), key=lambda kv: order.index(kv[0]) if kv[0] in order else len(order)):
        w(
            f"| {kind} | "
            + " | ".join(money(d[p]) if d.get(p) else "—" for p in PARTNERS)
            + f" | {money(sum(d.values()))} |"
        )
    w(
        "| **Total** | "
        + " | ".join(money(sum(d.get(p, 0) for d in wd.values())) for p in PARTNERS)
        + f" | **{money(wd_total)}** |"
    )
    w("")
    w(
        "  Nenhum desses lançamentos cita a fazenda. O pró-labore do Válter é pago a Vagner Siviero."
    )
    w(
        "- Em 2018/2019, logo depois das vendas, há dois documentos sobre retiradas: o relatório de divisão de 31/12/2018 (moradia, estudo dos filhos e pró-labore de cada sócio de 2012 a 2018) "
        'e a anotação da reunião de 14/01/2019 (`Documentos/retiradas_socios_2018.pdf`), que registra "Irineu/Raul – têm dinheiro a juros na empresa = 1,3% a 1,5%" e "Devolver capital – empresa para pagar juros". '
        "Nenhum dos dois menciona o dinheiro das vendas."
    )
    w("")
    w("## 4. Quanto já foi pago com esse dinheiro")
    w("")
    w(
        "As contas pagas da SEI mostram os impostos de vendas de soja em 2024. O valor de cada venda foi calculado a partir do imposto pago (o FUNRURAL de pessoa física, por exemplo, é 1,5 % do valor vendido). Detalhe no anexo."
    )
    w("")
    w("| Venda de soja (2024) | Valor | Origem | Como foi tributada |")
    w("|---|---|---|---|")
    for s in sorted(soy_sales.values(), key=lambda s: s["mes"]):
        how = (
            f"pessoa física (FUNRURAL {f0.br(s['rate'])} %, sem IRPJ/CSLL)"
            if s["pf"]
            else f"Agropecuária (FUNRURAL {f0.br(s['rate'])} % + IRPJ + CSLL)"
        )
        detail = (
            f"{sacks(s['sacas'])} sacas a {money(s['preco'])}"
            if s.get("sacas")
            else "soja (quantidade não informada)"
        )
        origem = s["origem"] + (
            " (provável: mesma tributação da venda Lorenzet)" if s in pj_other else ""
        )
        w(f"| {s['mes']}: {detail} | {money(s['base'])} | {origem} | {how} |")
    w("")
    w(
        f"- **Só a venda de {sacks(-sold['sacas'])} sacas ({money(rigo_sale['base'])}) vem da venda das terras** (parcela Rigo de 2021). "
        'As vendas da Agropecuária são aluguel da fazenda (Lorenzet); a de 1.300 sacas tem origem não identificada (o histórico do relatório está cortado em "COLONIA SA").'
    )
    w(
        f"- As duas vendas de pessoa física (1.300 e {sacks(-sold['sacas'])} sacas) não pagaram IRPJ nem CSLL, ao contrário das da Agropecuária. O grupo diz que vende os grãos "
        f'"{cnpj.replace("tributando a venda dos grãos ", "")}" (arq. 1, p. 6). **Não se sabe em nome de quem saíram essas duas notas nem em que conta entrou o dinheiro**; se foi em nome de um sócio, é retirada. O lado do Luís não recebe nada dos Rigo diretamente (informação do Luís, 28/10/2026).'
    )
    w(
        f'- **Escavadeira CAT 320 da SEI** ({money(exc_nf)}): o documento de 2025 diz que as {sacks(-sold["sacas"])} sacas foram vendidas "para aquisição da escavadeira hidráulica". '
        f"Elas renderam cerca de {money(rigo_sale['base'])}, um terço do preço; o restante foi pago pela SEI (entrada e duas parcelas em 03–05/2024 e um financiamento de {money(fin_total)} em 48 parcelas, "
        f"das quais {len(fin)} pagas no período). A escavadeira está no inventário 2025 da SEI e nenhum inventário registra crédito do Hotel ou da Agropecuária contra a SEI."
    )
    w(
        f'- **A confirmar:** a SEI comprou um imóvel na Rua Flores da Cunha ({money(imovel_nf)}) e pagou no período {money(imovel_paid)}, dos quais {money(imovel_soy)} com "CHEQUES SOJA TITTON" (05/2024). '
        "Titton é compradora de soja; nada liga essa soja à fazenda."
    )
    if farm:
        w(
            "- A SEI também pagou despesas da fazenda: "
            + "; ".join(
                f"{str(r[8]).strip(' -')} ({money(r[7])}, {r[0].strftime('%m/%Y')})"
                for r in sorted(farm, key=lambda r: r[0])
            )
            + (
                f". O ITR declara {itr_ha} ha, contra 411 ha da fazenda."
                if itr_ha
                else "."
            )
        )
    w("")
    w("## 5. Divergências e o que pedir")
    w("")
    if lorenzet_named:
        w(
            f'- O documento de 2025 diz que a soja das 4 safras Lorenzet "não foi vendida" ({sacks(sum(r["sacas"] for r in lor))} sacas), mas a SEI pagou impostos sobre "SOJA LORENZET" vendida em 2024 ({money(lorenzet_named[0]["base"])}).'
        )
    w(
        "- As contas pagas cobrem só a SEI, a Transportadora e a Abastecedora, de jul/2023 a out/2024; não há registro de entradas, nem extratos do Hotel ou da Agropecuária, nem nada de 2018 a meados de 2023. Sem isso não dá para fechar recebido − retirado − pago."
    )
    w(
        "- Pedir ao grupo: (1) extratos ou livro-caixa do Hotel e da Agropecuária de 2018 em diante, com a entrada de cada parcela; (2) a quitação da parcela Pellegrini de jun/2019; "
        "(3) em nome de quem saíram as notas de 1.300 e 2.606 sacas de 2024 e para qual conta foi o dinheiro; (4) de quem era a soja vendida à Titton; "
        "(5) onde estão as 37.694 sacas dos Rigo e as sacas Lorenzet; (6) o destino dos R$ 12.058.000 (pergunta Q6, já em rascunho)."
    )
    w("")
    w(
        "## Anexo: impostos de venda de soja pagos pela SEI (como estão no relatório mensal)"
    )
    w("")
    w("| Mês | Imposto | Valor | Histórico | Valor da venda |")
    w("|---|---|---|---|---|")
    for r in sorted(tax, key=lambda r: r[0]):
        s = next(s for s in soy_sales.values() if r in s["taxes"])
        w(
            f"| {r[0].strftime('%m/%Y')} | {r[3]} | {money(r[7])} | {r[8]} | {money(s['base'])} |"
        )
    w("")
    w(
        f'O preço de {money(rigo_sale["preco"])} das {sacks(-sold["sacas"])} sacas é dedução: o histórico está cortado depois de "VENDIDOS X"; '
        f"a conta fecha ({sacks(-sold['sacas'])} × {f0.br(rigo_sale['preco'])} × 1,5 % = {money(Decimal(str(rigo_sale['taxes'][0][7])))})."
    )
    w("")
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))
    print(
        f"contratado {contracted}; terra SC {land_crispin}; dinheiro {cash_2018}+{cash_2019}; soja 2024 {sum(soy_sales)}; Rigo {rigo_sale['base']}; "
        f"retiradas {wd_total}; imóvel {imovel_paid} (soja {imovel_soy})"
    )


if __name__ == "__main__":
    main()
