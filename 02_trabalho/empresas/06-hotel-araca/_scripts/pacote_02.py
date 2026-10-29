"""Packet 2 for company 06: exceptions (items where the sides disagree, items
Rafael cites outside the inventory, Rafael's unexplained total gap) and the
treatment of the open cut-off-date records. Largest value first.

Reads planilha-mestre-06.xlsx (read only) for values, ranges, components and
totals; Rafael's PDF and the 2018/2025 inventories for the figures and
quotes used in the arguments. Writes pacotes/pacote-02-excecoes.md.
Recommendations follow C0 (P10) with the C1 argument written out; where the
recommendation departs from C0, the packet says so and gives both options.

Usage: python pacote_02.py
"""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fase0_extracao_06 as f0

HERE = Path(__file__).resolve().parents[1]
MESTRE = HERE / "planilha-mestre-06.xlsx"
OUT = HERE / "pacotes" / "pacote-02-excecoes.md"
TOL = 0.005


def br(v: float) -> str:
    return f"{v:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")


def flat(s: str) -> str:
    return re.sub(r"\s+", " ", s)


def main():
    md5 = hashlib.md5(MESTRE.read_bytes()).hexdigest()[:10]
    wb = openpyxl.load_workbook(MESTRE, data_only=True)
    hdr = [c.value for c in wb["Mestre"][1]]
    rows = [
        dict(zip(hdr, r)) for r in wb["Mestre"].iter_rows(min_row=2, values_only=True)
    ]
    by_item = {
        re.match(r"item (\d{2})", r["Ref. 2025"]).group(1): r
        for r in rows
        if (r["Ref. 2025"] or "").startswith("item")
    }
    h01, c20 = by_item["01"], by_item["20"]
    casas = next(r for r in rows if "Rua da Divisa" in (r["Descrição"] or ""))
    toro = next(r for r in rows if "Fiat Toro" in (r["Descrição"] or ""))
    assert [r for r in rows if r["Situação"] == "Discorda"] == [h01, c20]
    comp = {
        r[0]: r for r in wb["Item01_componentes"].iter_rows(min_row=2, values_only=True)
    }
    tot = {r[0]: r[1] for r in wb["Totais"].iter_rows(min_row=2, values_only=True)}

    def T(prefix):
        return next(v for k, v in tot.items() if k.startswith(prefix))

    gap, adopted, g_total = (
        T("Diferença de soma"),
        T("Valores adotados"),
        T("Total grupo 2025"),
    )
    r_sum, r_decl, sel = (
        T("Soma dos itens de Rafael"),
        T("Total declarado"),
        T("Lista de Rafael"),
    )
    pares = {
        r[0]: r
        for r in wb["Pares_2018_2025"].iter_rows(min_row=2, values_only=True)
        if r[0]
    }

    pages = f0.pdf_pages(f0.SRC["rafael_pdf"])
    P = [flat(p) for p in pages]
    ranges = f0.rafael_ranges(pages)
    mk = next(r for r in ranges if r[0] == 7 and r[2] == "faixa")
    phys_min = next(r for r in ranges if r[2] == "Valor Mínimo Físico")[3]
    phys_max = next(r for r in ranges if r[2] == "Valor Máximo Físico")[3]
    renda = [r[3] for r in ranges if r[2].startswith("Valor") and "Renda" in r[2]]
    r20 = next(r for r in ranges if r[0] == 13 and r[2] == "faixa")
    land20, house20 = [r for r in ranges if r[0] == 14 and r[2] == "faixa"][:2]
    m2 = re.search(r"oscila entre R\$ ([\d\.]+)/m² e R\$ ([\d\.]+)/m²", P[10])
    m2_lo, m2_hi = f0.amount(m2.group(1)), f0.amount(m2.group(2))
    t01_aval = list(re.finditer(r"R\$ ([\d\.]+,\d{2})", P[5]))[-1].group(
        1
    )  # last amount on p. 6
    t20_aval = re.search(
        r"global de R\$ 750\.000,00, totalmente quitada\. R\$ ([\d\.]+,\d{2})", P[12]
    ).group(1)
    assert (
        f0.brl(t01_aval) == h01["Valor grupo 2025"]
        and f0.brl(t20_aval) == c20["Valor grupo 2025"]
    )
    renda_txt = "É o método mais realista para fundos hoteleiros e investidores"
    obj_txt = "o Método da Renda dita o preço"
    unif_txt = "supervalorização de terrenos unificados frente ao mar"
    piso20_txt = "Se a casa for uma construção mais antiga, necessitando de reformas/modernização"
    conc_txt = "itens anteriores de 02., à 19."
    for page, txt in (
        (8, renda_txt),
        (7, obj_txt),
        (7, unif_txt),
        (15, piso20_txt),
        (16, conc_txt),
    ):
        assert txt in P[page - 1], (page, txt)
    fontes7 = re.findall(r"https?://(?:www\.)?([a-z0-9\.-]+)", " ".join(P[6:9]))
    renda_calc = re.search(
        r"dá lucro de (R\$ [\d\.,]+)/ano e (R\$ [\d\.,]+) a 10,5 % e (R\$ [\d\.,]+) a 8,5 %",
        comp["Total item 01"][7],
    )

    inv25 = f0.parse_inventory(f0.SRC["inv2025"], "2025 atual")[0]
    inv25a = f0.parse_inventory(f0.SRC["inv2025_antigo"], "2025 antigo")[0]
    inv18 = f0.parse_inventory(f0.SRC["inv2018"], "2018")[0]
    obs25 = " ".join(f0.observation_paragraphs(f0.SRC["inv2025"]))
    red = f0.brl(re.search(r"4,00% = R\$ " + f0.MONEY, obs25).group(1))
    inss = f0.brl(re.search(r"aproximadamente R\$ " + f0.MONEY, obs25).group(1))
    c25 = f0.item01_components(inv25[0]["descricao"], "2025 atual")
    c18 = f0.item01_components(inv18[0]["descricao"], "2018")
    area25, area18 = c25[3]["area_m2"], c18[3]["area_m2"]
    v18_01 = inv18[0]["valor"]
    va_01 = inv25a[0]["valor"]
    pct_a = next(
        c["percentual_vantagem"]
        for c in f0.item01_components(inv25a[0]["descricao"], "a")
        if c["percentual_vantagem"]
    )
    pct_18 = c18[5]["percentual_vantagem"]
    item20_18 = pares["20"][8]
    lot_ups = [
        (r[9] / r[8] - 1) for k, r in pares.items() if k not in ("01", "20") and r[8]
    ]
    up_lo, up_hi = min(lot_ups), max(lot_ups)
    up_01 = h01["Valor grupo 2025"] / v18_01 - 1

    t25 = "\n".join(
        b if k == "p" else ""
        for k, b in f0.doc_blocks(
            next((f0.DRIVE / "2018/inventário-2018").glob("25. 1.*.docx"))
        )
    )
    veraneio = (
        re.search(r"manutenção das 02 casas de veraneio dos Sócios[^\n]*", t25)
        .group(0)
        .strip()
        .rstrip(".")
    )
    esquina = re.search(r"CASA DA ESQUINA: [^.]*?\)", t25).group(0)
    ao_lado = re.search(r"CASA AO LADO: Raul Siviero", t25).group(0)

    land_g, land_r = comp["1.1 Terreno"][3], comp["1.1 Terreno"][4]
    obra_g, obra_r = comp["1.4 Obra civil"][3], comp["1.4 Obra civil"][4]
    vant_g = comp["1.6 Vantagem da coisa feita"][3]
    moveis_g = sum(
        comp[k][3]
        for k in (
            "1.2 Móveis 36 aptos novos",
            "1.3 Móveis 108 aptos",
            "1.5 Equipamentos diversos",
        )
    )
    land_area = f0.brl(c25[0]["area_m2"])
    rate_g, rate_r = land_g / land_area, land_r / land_area
    nesga_g, nesga_r = 420 * rate_g, 420 * rate_r
    lot_fundos = by_item["03"]["Valor grupo 2025"] / f0.brl(
        next(i for i in inv25 if i["item"] == "03")["area_m2"]
    )
    lot_q38 = by_item["11"]["Valor grupo 2025"] / 360.0
    d01 = h01["Valor Rafael"] - h01["Valor grupo 2025"]
    d20 = c20["Valor Rafael"] - c20["Valor grupo 2025"]
    after = adopted + h01["Valor Rafael"] + c20["Valor Rafael"]
    assert abs(after - r_sum) < TOL

    L = []
    w = L.append
    w("# Hotel Araçá (empresa 06) — Pacote 2: exceções")
    w("")
    w(
        f"Data: 28/10/2026. Gerado por `_scripts/pacote_02.py` a partir de `planilha-mestre-06.xlsx` (md5 `{md5}`), do PDF de Rafael e dos inventários. Maior valor primeiro. Recomendação pelo C0 (P10), com o argumento técnico (C1) escrito; onde a recomendação não segue o C0, as duas opções estão lado a lado. Revisado por revisor independente (`_diario/revisao-pacote-02.md`) antes desta versão. Nada está aplicado."
    )
    w("")
    w(
        f"Já adotado no Pacote 1: R$ {br(adopted)} (20 itens). Total do grupo: R$ {br(g_total)}. Rafael: soma dos itens R$ {br(r_sum)}; total declarado R$ {br(r_decl)}."
    )
    w("")
    w(
        "| # | Item | Grupo | Rafael | Na lista de Rafael? | C0 aponta | Recomendação | Efeito no seu 1/6 |"
    )
    w("|---|---|---|---|---|---|---|---|")
    w(
        f"| E1 | 01 Hotel, item inteiro (terreno, prédio, móveis e equipamentos) | R$ {br(h01['Valor grupo 2025'])} | R$ {br(h01['Valor Rafael'])} (p. 17) | não | maior (Rafael) | **Adoto Rafael**, com ressalvas para você decidir | +R$ {br(d01 / 6)} |"
    )
    w(
        f"| E2 | Duas casas na Rua da Divisa com a Rua Três (fora do inventário) | — | {casas['Faixa de mercado de Rafael']}, fora do total dele | não | — | **Sem valor agora**; G8; destino na Parte 4 | — |"
    )
    w(
        f"| E3 | 20 Lote 15/q02, Condado de Capão, com casa | R$ {br(c20['Valor grupo 2025'])} | R$ {br(c20['Valor Rafael'])} (p. 19) | não | maior (Rafael) | **Adoto Rafael** | +R$ {br(d20 / 6)} |"
    )
    w(
        f"| E4 | Diferença não identificada no total de Rafael | — | R$ {br(gap)} (p. 20) | — | maior (adotar) | **Não adotar** (contra o C0; duas opções abaixo) | R$ 0,00 (ou +R$ {br(gap / 6)}) |"
    )
    w(
        "| E5 | Fiat Toro de serviço do hotel (fora do inventário) | — | sem valor (p. 25) | não | — | **Sem valor agora**; G8; Parte 5 | — |"
    )
    w(
        "| E6 | Registros de data-base nº 1 e 2 | — | — | — | — | **Manter sem ajuste**, só registrados | — |"
    )
    w("")
    w(
        f"Se as recomendações forem aprovadas, os bens da empresa somam R$ {br(after)}, igual à soma dos itens de Rafael: +R$ {br(after - g_total)} contra o grupo, ou +R$ {br((after - g_total) / 6)} no seu 1/6."
    )
    w("")
    w("---")
    w("")
    w(
        f"## E1 — Item 01, Hotel: grupo R$ {br(h01['Valor grupo 2025'])} × Rafael R$ {br(h01['Valor Rafael'])}"
    )
    w("")
    w(
        f"**Recomendação:** adoto Rafael, R$ {br(h01['Valor Rafael'])}, para o **item inteiro, inclusive móveis e equipamentos** (C0: o hotel fica na empresa). Efeito: +R$ {br(d01)} na empresa, +R$ {br(d01 / 6)} no seu 1/6. A base de Rafael é fraca em parte; as ressalvas estão abaixo para você decidir, como pede o brief."
    )
    w("")
    w(
        f'**O que entra nos 66 milhões.** Na p. 7, Rafael diz que a avaliação física é "terreno + área construída": terreno a R$ {br(rate_r)}/m² ({br(land_area)} m² = R$ {br(land_r)}) e construção a R$ 3.000/m² ({area25} m² = R$ {br(obra_r)}). Os móveis e equipamentos do grupo (R$ {br(moveis_g)}) não entram nessa conta. Na p. 17, porém, ele põe os 66 milhões ao lado da descrição completa do item, com móveis e equipamentos. Adotar Rafael é adotar 66 milhões pelo item inteiro, sem somar os móveis à parte.'
    )
    w("")
    w("**Parcela a parcela** (aba Item01_componentes da mestre):")
    w("")
    w("| Parcela | Grupo 2025 | Rafael (p. 7) | Diferença |")
    w("|---|---|---|---|")
    w(
        f"| Terreno, {br(land_area)} m² | R$ {br(land_g)} (R$ {br(rate_g)}/m²) | R$ {br(land_r)} (R$ {br(rate_r)}/m²) | +R$ {br(land_r - land_g)} |"
    )
    w(
        f"| Obra civil, {area25} m² | R$ {br(obra_g)} (CUB dez/2022, menos 20 %) | R$ {br(obra_r)} (R$ 3.000/m², já depreciado) | +R$ {br(obra_r - obra_g)} (+{(obra_r / obra_g - 1) * 100:.0f} %) |"
    )
    w(
        f"| Vantagem de 15 % (calculada sobre obra, móveis e equipamentos) | R$ {br(vant_g)} | não há | −R$ {br(vant_g)} |"
    )
    w(
        f"| Móveis e equipamentos | R$ {br(moveis_g)} | fora da conta física | −R$ {br(moveis_g)} |"
    )
    w(
        f"| **Total** | **R$ {br(h01['Valor grupo 2025'])}** | **R$ {br(h01['Valor Rafael'])}** | **+R$ {br(d01)}** |"
    )
    w("")
    w("**A favor de Rafael (C1):**")
    w(
        f"- O valor do grupo fica abaixo do piso da faixa de mercado de Rafael para o hotel, de R$ {br(mk[3])} a R$ {br(mk[4])} (p. 7). Os 66 milhões são o piso da avaliação física dele, cujo teto é R$ {br(phys_max)}."
    )
    w(
        f"- O grupo dá ao terreno do hotel, {br(land_area)} m² de frente para o mar, R$ {br(rate_g)}/m². É o mesmo que dá aos lotes de fundos (item 03, R$ {br(lot_fundos)}/m²), menos que aos lotes da quadra 38 (R$ {br(lot_q38)}/m²) e menos que a faixa de Rafael para terreno residencial comum no Araçá (R$ {br(m2_lo)} a {br(m2_hi)}/m², p. 11). O próprio grupo justifica os 15 % de vantagem pela localização em frente ao mar e pelos índices urbanísticos difíceis de aprovar hoje (arquivo 1, p. 11), mas não leva isso ao preço do terreno."
    )
    w(
        f"- Versões do próprio grupo: o inventário 2025 antigo dava ao hotel R$ {br(va_01)} (vantagem de {pct_a} %); em 2018, a vantagem era de {pct_18} %. Entre 2018 e 2025 o grupo subiu os lotes vizinhos de {up_lo * 100:.0f} % a {up_hi * 100:.0f} %, e o hotel, {f'{up_01 * 100:.1f}'.replace('.', ',')} %. São fatos das versões do grupo, não reavaliação."
    )
    w(
        f"- Na obra civil, Rafael dá R$ {br(obra_r - obra_g)} a mais. O CUB do grupo é de dezembro de 2022, dois anos antes da data-base (registro nº 2)."
    )
    w("")
    w("**Contra, ou ressalvas (o outro lado vai usar):**")
    w(
        f'- **O método que o próprio Rafael chama de mais realista dá menos.** Na p. 7 ele diz que, se o hotel continuar operando, "{obj_txt}"; na p. 8, sobre a renda: "{renda_txt}". A renda dá R$ {br(renda[0])} a {br(renda[1])}, faixa que inclui o valor do grupo.'
        + (
            f" Pela aritmética das premissas dele (p. 8–9), o lucro seria {renda_calc.group(1)}/ano, e o valor, {renda_calc.group(2)} a {renda_calc.group(3)}."
            if renda_calc
            else ""
        )
    )
    w(
        f'- **A faixa de R$ {br(mk[3])} a {br(mk[4])} não tem conta no PDF.** Ela "integra" o físico ({br(phys_min)} a {br(phys_max)}) e a renda ({br(renda[0])} a {br(renda[1])}); a média simples das duas daria {br((phys_min + renda[0]) / 2)} a {br((phys_max + renda[1]) / 2)}.'
    )
    w(
        f'- **O teto físico supõe terreno unificado** ("{unif_txt}", p. 7), que não existe: o grupo diz que não há plano de unificação definido, nem projeto, nem habite-se (arquivo 1, p. 11). Os lotes 13 e 14 da quadra 37 nem se sabe de quem são (G1).'
    )
    w(
        f"- **Base de Rafael.** Ele não cita venda de terreno de frente para o mar a R$ {br(rate_r)}/m². As fontes das p. 7–9 são {', '.join(sorted(set(fontes7)))}; a do Jornal do Comércio é de dezembro de 2025, depois da data-base (registro nº 1). É custo de reposição, não comparação de vendas."
    )
    w(
        f'- **Duas leituras de Rafael.** Na seção "avaliação patrimonial executada" ele repete o valor do grupo, R$ {t01_aval} (p. 6); os 66 milhões estão no balanço executado (p. 17). Contraponto: a concordância expressa dele cobre só os "{conc_txt}" (p. 16), não o item 01.'
    )
    w(
        f'- **Área do terreno.** A nesga de 420 m² é "parte do lote 12", que também é contado com 360 m² (G2). Se houver área contada a mais, o terreno cai R$ {br(nesga_g)} pelo grupo e R$ {br(nesga_r)} por Rafael.'
    )
    w(
        f"- **Área construída** de {area18} m² em 2018 e {area25} m² em 2025, nos dois lados (G3)."
    )
    w(
        f"- **Regularização.** O grupo estima, sem deduzir do total, redução de R$ {br(red)} no item 01 e INSS da obra de R$ {br(inss)}; não há projeto, habite-se nem profissional contratado. O PPCI está aprovado e em vigor, e o poço com outorga não é usado para abastecimento (arquivo 1, p. 11). Rafael não desconta nada disso. Não há desconto em dobro hoje; o tema vai para os passivos (Parte 4) e a Parte 6."
    )
    w(
        "- **Sobreposição no quarteirão.** Os itens 03, 21 e 22 (lotes 16, 15 e 17 da quadra 37) são avaliados à parte; se parte do hotel estiver sobre eles, há sobreposição com o item 01."
    )
    w("")
    w(
        f"**Opções:** (a) Rafael, R$ {br(h01['Valor Rafael'])} pelo item inteiro, como recomendo; (b) grupo, R$ {br(h01['Valor grupo 2025'])}, se você considerar que a base de Rafael não sustenta o valor. Não há terceira opção sem reavaliação própria (P5)."
    )
    w("")
    w("---")
    w("")
    w("## E2 — Duas casas na Rua da Divisa com a Rua Três (citadas só por Rafael)")
    w("")
    w(
        f"**Recomendação:** sem valor agora; destino na Parte 4, junto com os itens 22 e 23 de 2018. Rafael estima {casas['Faixa de mercado de Rafael']}, não soma ao total dele e diz que devem entrar na negociação ou ter o destino comprovado (p. 24–26)."
    )
    w("")
    w(
        f'- **Identidade provável.** O doc 25 de 2018, na seção do Hotel Araçá, diz que foram agregados ao capital do Hotel a "{veraneio}", e que "{esquina}" e "{ao_lado}". Isso bate com os itens 22 e 23 de 2018: lotes vizinhos com casa, o 23 de esquina, com os nomes de Raul e Élcio no lugar do valor. O de Élcio foi trocado por um lote da SEI (doc 25, seção Élcio; candidato 2 da Parte 6).'
    )
    w(
        '- **Pergunta G8, refeita:** de quem são hoje as duas casas ("pertence à empresa Siviero": qual empresa?), e quem paga a manutenção.'
    )
    w(
        "- A manutenção das casas pelo Hotel entra nos candidatos da Parte 4 e 6 (despesa da empresa a favor de sócios)."
    )
    w(
        "- Pelo C0, se fossem da empresa, levariam o maior valor. Hoje não há valor do grupo, e o de Rafael é uma faixa sem bem identificado."
    )
    w("")
    w("---")
    w("")
    w(
        f"## E3 — Item 20, Lote 15/q02, Condado de Capão: grupo R$ {br(c20['Valor grupo 2025'])} × Rafael R$ {br(c20['Valor Rafael'])}"
    )
    w("")
    w(
        f"**Recomendação:** adoto Rafael, R$ {br(c20['Valor Rafael'])} (C0: fica na empresa). Efeito: +R$ {br(d20)} na empresa, +R$ {br(d20 / 6)} no seu 1/6."
    )
    w("")
    w("**A favor de Rafael (C1):**")
    w(
        f'- Rafael adota o piso da própria faixa, R$ {br(r20[3])} a {br(r20[4])} (p. 13; adoção na p. 19). Ele liga o piso a uma casa "mais antiga, necessitando de reformas" (p. 15): é o valor conservador da faixa.'
    )
    w(
        f"- Só o terreno, 725,39 m², Rafael estima em R$ {br(land20[3])} a {br(land20[4])}; a casa, em R$ {br(house20[3])} a {br(house20[4])} (p. 14). O valor do grupo fica abaixo até do piso só do terreno."
    )
    w(
        f"- O grupo manteve R$ {br(item20_18)} de 2018 a 2025, enquanto subiu os lotes de Capão de {up_lo * 100:.0f} % a {up_hi * 100:.0f} %. O imóvel custou R$ 750.000,00 em maio de 2014, com móveis."
    )
    w("")
    w("**Ressalvas:**")
    w(
        "- Rafael aplica o preço por m² de lotes de cerca de 300 m² a um lote de 725,39 m², sem fator de área (p. 13–14)."
    )
    w(
        "- Compra por contrato particular, com matrícula 69.488 citada. O grupo afirma que todos os imóveis estão em nome da empresa (arquivo 1, p. 11). Pergunta G6: se o registro não for da empresa, o item sai (Parte 4)."
    )
    w(
        f'- Na "avaliação patrimonial executada" Rafael repete o valor do grupo, R$ {t20_aval} (p. 13); os 2,4 milhões estão no balanço executado (p. 19). A concordância expressa dele (p. 16) não cobre o item 20.'
    )
    w(
        "- Fontes de Rafael de 2025 (registro nº 1). Ele avalia uma casa de 160 m²; o inventário diz 158,18 m²."
    )
    w("")
    w("---")
    w("")
    w(f"## E4 — Diferença não identificada no total de Rafael: R$ {br(gap)}")
    w("")
    w(
        f"Os 22 itens do balanço executado de Rafael somam R$ {br(r_sum)}; ele declara R$ {br(r_decl)} (p. 20). O 1/6 que Rafael calcula para o seu lado (p. 21) usa o total declarado: R$ {br(r_decl / 6)} em vez de R$ {br(r_sum / 6)}, e o crédito, R$ {br(r_decl / 6 - sel)} em vez de R$ {br(r_sum / 6 - sel)}."
    )
    w("")
    w("| Opção | Valor na empresa | Efeito no seu 1/6 | Argumento | Risco |")
    w("|---|---|---|---|---|")
    w(
        f"| (a) Não adotar (recomendo) | R$ 0,00 | R$ 0,00 | Não há bem que corresponda; a decisão é item a item (P5). | Abre mão de R$ {br(gap / 6)} que o C0 favoreceria. |"
    )
    w(
        f"| (b) Adotar como diferença do total de Rafael | +R$ {br(gap)} | +R$ {br(gap / 6)} | Segue o C0 e o total que Rafael declara. | Sem bem identificado, o outro lado aponta o erro de soma e a posição perde credibilidade nos itens maiores. |"
    )
    w("")
    w(
        f"- O valor é igual ao de cada um dos itens 03 a 10 (lotes 16 e 23 a 29 da quadra 37, R$ {br(gap)} cada). Rafael pode ter contado um lote não listado; os lotes 13 e 14 da quadra 37 (G1) são candidatos, mas é só hipótese. Se um deles aparecer como bem da empresa, volta como item, com valor."
    )
    w(
        "- **Diferença do precedente da SEI.** No imóvel Oleoplan, o valor de Rafael existia para o item e era o favorecido pelo C0 (bem da lista); manteve-se o valor dele e só se registrou o risco de alegação de erro. Aqui não há item: o C0 não tem bem onde se aplicar."
    )
    w("")
    w("---")
    w("")
    w("## E5 — Fiat Toro de serviço do hotel (citada só por Rafael)")
    w("")
    w(
        f'**Recomendação:** sem valor agora. Rafael diz que não consta do inventário e que deve entrar na negociação ou ter o destino comprovado (p. 25–26). Não há placa nem valor de nenhum lado. Pergunta G8. Se for da empresa em 31/12/2024, volta com valor; se não, vai para a Parte 5 com o destino documentado. Situação na mestre: "{toro["Situação"]}".'
    )
    w("")
    w("---")
    w("")
    w(
        "## E6 — Tratamento dos registros de data-base nº 1 e 2 (`registro-data-base.md`)"
    )
    w("")
    w(
        "**Recomendação:** manter sem ajuste nos dois; ficam só registrados (P2), como argumento."
    )
    w(
        "- **Nº 1 — Rafael usa fontes de mercado de 2025** (itens 01 e 20, casas da Rua da Divisa). Os valores de Rafael são adotados como estão: sem reavaliação própria (P5), não há como trazê-los a 31/12/2024. A fuga fica registrada e é ressalva em E1 e E3."
    )
    w(
        "- **Nº 2 — o grupo usa o CUB de dezembro de 2022** na obra civil do hotel. Mantido sem ajuste: o CUB antigo subavalia a obra na data-base, o que reforça o C1 a favor do valor maior em E1."
    )
    w("")
    w("---")
    w("")
    w(
        "Perguntas citadas: `perguntas-ao-grupo.md` (G1 a G11), ainda não enviadas; G8 será refeita como em E2. Candidatos à Parte 6: `parte6-candidatos.md` (entra a manutenção das casas)."
    )
    OUT.write_text("\n".join(L), encoding="utf-8")
    print(
        f"mestre {md5}; E1 {d01:,.2f}; E3 {d20:,.2f}; gap {gap:,.2f}; after {after:,.2f}; obra +{obra_r - obra_g:,.2f}; ups {up_lo:.3f}-{up_hi:.3f} hotel {up_01:.3f}; fontes {sorted(set(fontes7))}"
    )
    print("saved", OUT)


if __name__ == "__main__":
    main()
