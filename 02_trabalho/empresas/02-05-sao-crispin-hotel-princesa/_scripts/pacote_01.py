"""Write pacotes/pacote-01.md (block rules) from planilha-mestre-02-05.xlsx.

Every number in the packet comes from the master spreadsheet (rebuilt from
the sources first) or is read here from the sources. Labels of assets are
typed; values, pages and quotes are not.

v2 (28/10/2026) applies `_diario/revisao-pacote-01.md` (A1-A9).

Run: PYTHONIOENCODING=utf-8 python -W ignore pacote_01.py
"""

from __future__ import annotations

import contextlib
import io
import os
import re
import sys
from decimal import Decimal

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_mestre_02_05 as bm
import fase0_extracao as f0

OUT = os.path.join(bm.DIR, "pacotes", "pacote-01.md")
SEI_PACOTE_04 = os.path.join(f0.ROOT, r"02_trabalho\pacotes\pacote-04.md")

# short names for the reader (labels only)
NAME_2025 = {
    "01": "Prédio do hotel e terreno sob o prédio",
    "02": "Terreno ao lado do hotel (esquina)",
    "03": "Terreno nos fundos do hotel",
    "04": "Sede Campestre",
    "05": "Fazenda São Crispin (407,53 ha)",
    "06": "Fazenda São Crispin, 4 ha (Dirceu/Lairton)",
    "07": "Lote rural 52, Cotiporã (Bassani)",
    "08": "Caminhão Mercedinha MB 709",
}
NAME_2018 = {
    2: "Fazenda, 106,53 ha (Alzir), vendidos aos Rigo",
    3: "Fazenda, 4,9 ha (Anastácio), vendidos aos Pellegrini",
    4: "Fazenda, 10 ha (Aito), vendidos aos Pellegrini",
    5: "Fazenda, 26,16 ha (Ademar), vendidos aos Pellegrini",
    6: "Fazenda Três Porteiras, 44,16 ha (Garbin), vendidos aos Pellegrini",
    8: "Benfeitorias da fazenda (cercas, galpão, implementos)",
    9: "Lote rural 24, Cotiporã, 10,2 ha, vendido a Angeli",
    10: "Lote rural 12, Veranópolis (Sangali)",
    11: "Lote rural 89, Fagundes Varela",
    12: "Lote 33, Cotiporã (mato nativo)",
    14: "Terras de São Brás, Cotiporã, 69,8 ha (lotes 63, 65, 66 e N)",
    15: "Parte do lote 63, Cotiporã, 4,5 ha (Tercílio Dal Más)",
    16: "Trator Valtra BM100",
}
OWNER = {
    "11": ("da SEI", "áreas rurais", "sessão da SEI"),
    "08": ("da Transportadora", "bens", "sessão da Abastecedora e Transportadora"),
}


def money(v) -> str:
    return "—" if v in (None, "") else f"R$ {f0.br(Decimal(str(v)))}"


def pdf_sentence(path: str, page: int, pattern: str, what: str) -> str:
    text = " ".join(f0.pymupdf.open(path)[page - 1].get_text().split())
    return f0.need(pattern, text, what).group(0)


def main() -> None:
    with contextlib.redirect_stdout(io.StringIO()):
        bm.main()
    wbm = openpyxl.load_workbook(bm.OUT, data_only=True)
    ws = wbm["Mestre"]
    hdr = [c.value for c in ws[1]]
    rows = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    col = {key: title for key, title, _ in bm.COLS}
    resumo = {r[0]: r[1] for r in wbm["Resumo"].iter_rows(values_only=True)}

    def g(r, key):
        return r[col[key]]

    def seq(r) -> int:
        return int(re.search(r"seq\. (\d+)", str(g(r, "item_2018"))).group(1))

    hotel = [r for r in rows if str(g(r, "chave")).startswith("05-2025")]
    r1 = [r for r in hotel if g(r, "situacao") == "Concorda"]
    exc = sorted(
        (r for r in hotel if r not in r1),
        key=lambda r: -max(g(r, "valor_grupo"), g(r, "valor_rafael")),
    )
    other = [r for r in rows if "outra empresa" in str(g(r, "situacao"))]
    soy = [r for r in rows if str(g(r, "chave")).startswith("NA-")]
    dest = [r for r in rows if "-2018-s" in str(g(r, "chave")) and r not in other]
    total_hotel = sum(Decimal(str(g(r, "valor_grupo"))) for r in hotel)
    total_r1 = sum(Decimal(str(g(r, "valor_grupo"))) for r in r1)
    assert all(g(r, "valor_grupo") == g(r, "valor_rafael") for r in r1)
    assert len(hotel) + len(other) + len(dest) + len(soy) == len(rows)

    # facts read from the sources for the caveats
    raf_pdf = f0.SRC["rafael_05"]
    by_item = {g(r, "item"): r for r in hotel}
    cond_p11 = pdf_sentence(
        raf_pdf,
        11,
        r"fica sujeita à validação pelo escritório de advocacia contratado[^.]*\.",
        "p. 11",
    )
    cond_p21 = pdf_sentence(
        raf_pdf, 21, r"A presente seleção fica condicionada[^.]*\.", "p. 21"
    )
    averb = pdf_sentence(
        f0.SRC_LATER["respostas_arq01"],
        9,
        r"No que tange ao imóvel conhecido como SEDE CAMPESTRE[^.]*\.",
        "arq. 1 p. 9",
    )
    faixas: dict[str, list[tuple[int, str]]] = {}
    for f in openpyxl.load_workbook(f0.OUT_XLSX, data_only=True)[
        "Rafael 05 faixas"
    ].iter_rows(min_row=4, values_only=True):
        faixas.setdefault(f[1], []).append((f[2], f[3]))

    def faixa(item: str, needle: str) -> tuple[int, str]:
        return next((p, t) for p, t in faixas[item] if needle in t)

    caveats: dict[str, list[str]] = {}
    d02 = g(by_item["02"], "descricao")
    rate02 = f0.need(r"X R\$ (" + f0.MONEY + r"), ao metro", d02, "02").group(1)
    p8, t8 = faixa("01", "por m²")
    caveats["02"] = [
        "O terreno é parte da matrícula 1.412 (1.800 m²): 720 m² estão no item 01 e 1.080 m² aqui (resposta do grupo ao Dr. Pablo, arq. 1 p. 8–9). Nenhuma parte é contada duas vezes.",
        (
            f'Risco: o grupo usa R$ {rate02}/m² e Rafael concorda (p. 10), mas na avaliação do item 01 (p. {p8}) ele escreve: "{t8}" '
            "Essa faixa vale para a matrícula inteira, incluindo este terreno, e volta com o item 01 no pacote 2. "
            "O bem fica na empresa, então o C0 pede o maior valor; os dois lados já dão o mesmo."
        ),
    ]
    d03 = g(by_item["03"], "descricao")
    rate03 = f0.need(r"= R\$ (" + f0.MONEY + r") ao metro", d03, "03").group(1)
    m2_03 = f0.need(r"área de ([\d.,]+)m²", d03, "03 área").group(1)
    p9, t9 = faixa("01", "450 m²")
    caveats["03"] = [
        f"O documento usa R$ {rate03}/m²: {m2_03} m² × {rate03} = {money(f0.num(m2_03) * f0.brl(rate03))}, arredondado para {money(g(by_item['03'], 'valor_grupo'))}.",
        f'Risco: na avaliação do item 01 (p. {p9}) Rafael escreve: "{t9}" O grupo usa R$ {rate03}/m². Mesmo argumento do item 02.',
    ]
    d04 = g(by_item["04"], "descricao")
    casa = re.search(
        r"Sede Campestre = (\.+)m² de área construída: R\$ (" + f0.MONEY + ")", d04
    )
    caveats["04"] = [
        f'Benfeitorias sem averbação: "{averb}" (arq. 1 p. 9).',
        f"A área da casa de alvenaria está em branco no documento; a casa vale R$ {casa.group(2)}."
        if casa
        else "",
        'O texto diz "25,00% (Vinte por cento)"; o cálculo usa 25 %.',
        'Em 2018 a SEI também listava a Sede Campestre como bem dela (lista "bens relacionados nos inventários de outras empresas"); em 2025 só está no Hotel. Candidato nº 1 a duplicidade (Parte 6).',
        "Está na lista de Rafael (vem para nós): custo e risco da averbação ficam com quem recebe. Vale pergunta ao grupo; não muda o valor (P5).",
        'Nota sua na planilha: "Por que CUB/AGOSTO DE 2017".',
    ]
    d07 = g(by_item["07"], "descricao")
    rate07 = f0.need(r"X R\$ (" + f0.MONEY + ")", d07, "07").group(1)
    s10 = next(r for r in other if seq(r) == 10)
    s14 = next(r for r in other if seq(r) == 14)
    caveats["07"] = [
        "Na versão antiga do inventário 2025 o lote 52 aparecia também na SEI; a versão atual corrigiu (só no Hotel). Candidato nº 2 (Parte 6), só para registro.",
        (
            f"Informação: o grupo usa R$ {rate07}/ha; para terras próximas Rafael adota valores maiores por hectare "
            f"(São Brás: {money(g(s14, 'valor_rafael'))} por 69,8 ha; lote 12: {money(g(s10, 'valor_rafael'))} por 13,0 ha). "
            'Não há segundo valor para este item, então pela P5 o valor fica; seria argumento só para um "caso extremo".'
        ),
    ]
    caveats = {k: [c for c in v if c] for k, v in caveats.items()}
    with open(SEI_PACOTE_04, encoding="utf-8") as fh:
        sei7 = any(ln.startswith("| 7 |") and "3.062" in ln and "Pendente" in ln for ln in fh)

    lines: list[str] = []
    w = lines.append
    w(
        "# Pacote 1 — Agropecuária São Crispin (02) e Hotel Princesa dos Vales (05): regras em bloco"
    )
    w("")
    w(
        "Rascunho revisado (revisão independente em `_diario/revisao-pacote-01.md`). Números gerados por `_scripts/pacote_01.py` a partir da planilha-mestre, refeita das fontes a cada execução. **Nada aqui está aprovado.**"
    )
    w("")
    w(
        f"O Hotel tem {len(hotel)} bens em 2025, {money(total_hotel)} pelo grupo. A Agropecuária não tem bens no inventário 2025 (R$ 0,00). "
        f"Este pacote propõe o valor de {len(r1)} bens do Hotel ({money(total_r1)}) e trata em duas regras os bens de 2018 que hoje estão em outra empresa e os créditos em soja. "
        "Os outros bens do Hotel vão para o pacote 2."
    )
    w("")
    w(
        f"## Regra 1 — {len(r1)} bens do Hotel com o mesmo valor nas duas avaliações e concordância expressa de Rafael"
    )
    w("")
    w(
        "**Proposta: adotar o valor do grupo, que é igual ao de Rafael.** O critério C0 não muda nada: com valores iguais, o menor e o maior coincidem. "
        f"Como são só {len(r1)} bens, a amostra é a regra inteira."
    )
    w("")
    w(
        "| Item | Bem | Par em 2018 (valor 2018) | Grupo 2025 | Rafael | Na lista de Rafael? | Concordância de Rafael |"
    )
    w("|---|---|---|---|---|---|---|")
    for r in r1:
        k = g(r, "item")
        conc = (
            str(g(r, "obs_rafael"))
            .split(";")[0]
            .replace("concordância expressa (", "")
            .rstrip(")")
        )
        par = (
            str(g(r, "item_2018"))
            .replace("05 2018", "Hotel 2018")
            .replace("02 2018", "Agropecuária 2018")
        )
        w(
            f"| {k} | {NAME_2025[k]} | {par} ({money(g(r, 'valor_2018'))}) | {money(g(r, 'valor_grupo'))} | {money(g(r, 'valor_rafael'))} | {g(r, 'na_lista')} | {conc} |"
        )
    w(
        f"| | **Total** | | **{money(total_r1)}** | **{money(sum(Decimal(str(g(r, 'valor_rafael'))) for r in r1))}** | | |"
    )
    w("")
    w(
        "Ressalvas por item (nenhuma muda o valor; ficam para a ficha de fechamento e para as perguntas ao grupo):"
    )
    w("")
    for r in r1:
        k = g(r, "item")
        w(f"- **Item {k} — {NAME_2025[k]}.**")
        for c in caveats.get(k, []):
            w(f"  - {c}")
    w("")
    w(f"Ressalvas que valem para os {len(r1)}:")
    w("")
    w(
        f'- A concordância de Rafael é condicionada: "{cond_p11}" (p. 11). A escolha dos bens também: "{cond_p21}" (p. 21).'
    )
    w(
        '- Data dos valores: o grupo diz que avaliou pelo mercado "na data da elaboração do inventário" e que o termo final combinado é outubro/2025 (arq. 1 p. 1 e 9), não 31.12.2024. Registrado como RDB-1; não muda número.'
    )
    w("")
    w(
        f"## Regra 2 — {len(other)} bens de 2018 da Agropecuária que estão em 2025 no inventário de outra empresa"
    )
    w("")
    w(
        "**Proposta: não entram no total das empresas 02 e 05; o valor é decidido pela sessão da empresa onde o bem está em 2025.** "
        "Aqui fica só a localização. O destino 2018 → 2025 continua pendente até haver documento da transferência (perguntas Q9 e Q13 ao grupo) e será confirmado na etapa 6. "
        "Os valores de Rafael são só referência."
    )
    w("")
    w(
        "| Bem (2018) | Valor 2018 | Onde está em 2025 | Grupo 2025 | Rafael | Quem decide |"
    )
    w("|---|---|---|---|---|---|")
    for r in other:
        doc = re.search(r"Sim no doc (\d+)", str(g(r, "existe_2025"))).group(1)
        item = re.search(r"item (\d+)", str(g(r, "existe_rafael"))).group(1)
        who, part, owner = OWNER[doc]
        v18 = (
            money(g(r, "valor_2018"))
            if g(r, "valor_2018") not in (None, "")
            else 'sem valor ("não divisível")'
        )
        w(
            f"| {NAME_2018[seq(r)]} (mat. {g(r, 'matricula')}) | {v18} | inventário {who}, {part}, item {item} | {money(g(r, 'valor_grupo_rafael'))} | {money(g(r, 'valor_rafael'))} | {owner} |"
        )
    w("")
    w("Notas:")
    w("")
    w(
        "- Os três lotes rurais (12, 89 e 33) já foram pareados pela SEI (pacote 4 da SEI, aprovado), que adotou o valor de Rafael para os três (pacote 3 da SEI). Em 2018 a SEI já os listava como bens dela, embora estivessem no inventário da Agropecuária."
    )
    if sei7:
        w(
            '- Lote 12: em 2018 tinha 20,5 ha (13,0 + 7,5); em 2025 tem 13,0 ha. Os 7,5 ha coincidem em área e vendedor com o lote 14 da SEI (outra matrícula, 4.194). A SEI deixou o caso pendente (pacote 4 da SEI, caso 7: "Pendente: pedir a matrícula 3.062 atualizada"). Para não tratar em dobro, o caso fica com a SEI; aqui só a pergunta Q9 e o candidato nº 3 (Parte 6).'
        )
    w(
        '- Terras de São Brás: em 2018 estavam na Agropecuária sem valor ("não divisível"); em 2025 estão na Transportadora. A sessão da Abastecedora e Transportadora liga a área de 69,8 ha aos "créditos hereditários e/ou creditórios" da venda das quotas do Rovílio. Nenhuma fonte traz o documento da transferência (Q13).'
    )
    w("")
    w(f"## Regra 3 — {len(soy)} créditos em soja que nenhuma avaliação valoriza")
    w("")
    w(
        "**Proposta: vão para a Parte 5 (assuntos não abordados), sem valor agora.** É o achado E1 da análise de ACERTOS, já aprovado por você. Perguntas Q4 e Q5 ao grupo."
    )
    w("")
    w("| Crédito | Empresa |")
    w("|---|---|")
    for r in soy:
        w(f"| {g(r, 'descricao')} | {g(r, 'empresa')} |")
    w("")
    w(
        'Pontos em aberto: a soja dos Rigo está marcada "falta receber" numa tabela e "não vendida" em outra parte do mesmo documento (pode ser soja em depósito); as sacas de 2024 estão anotadas para a Agropecuária, mas o contrato tem o Hotel como arrendador.'
    )
    w("")
    w("## O que não está neste pacote")
    w("")
    w("**Pacote 2 (exceções, maior valor primeiro):**")
    w("")
    for r in exc:
        k = g(r, "item")
        why = {
            "05": "Rafael pede avaliação independente da fazenda",
            "06": "parte da mesma fazenda; segue o item 05",
        }.get(k, "valores diferentes")
        w(
            f"- Item {k} — {NAME_2025[k]}: grupo {money(g(r, 'valor_grupo'))}, Rafael {money(g(r, 'valor_rafael'))}, na lista de Rafael: {g(r, 'na_lista')} ({why})."
        )
    passivo = next(
        wbm["Passivos e deduções"].iter_rows(min_row=2, max_row=2, values_only=True)
    )
    w(
        f"- Junto com os itens 05 e 06: a redução de {money(passivo[2])} para regularizar a fazenda, descrita nos inventários 02 e 05 e não abatida de nenhum total."
    )
    w("")
    w(f"**Etapa 6 (destino 2018 → 2025), {len(dest)} bens de 2018:**")
    w("")
    for r in dest:
        w(
            f"- {NAME_2018[seq(r)]} ({money(g(r, 'valor_2018'))} em 2018): {g(r, 'situacao')}."
        )
    w("")
    paired = [r for r in hotel if g(r, "existe_2018") == "Sim"]
    w(
        f"**Conferência dos bens de 2018:** {len(paired)} têm par num bem do Hotel 2025 ({sum(1 for r in paired if r in r1)} na Regra 1 e {sum(1 for r in paired if r in exc)} no pacote 2), "
        f"{len(other)} estão na Regra 2 e {len(dest)} na etapa 6: {len(paired) + len(other) + len(dest)} no total, igual aos bens dos inventários 2018 do Hotel e da Agropecuária."
    )
    w("")
    w("**Pendente com você, a base do benchmark 2018:**")
    w("")
    for k, v in resumo.items():
        if str(k).startswith(("A.", "B.")):
            w(f"- {k}: {money(v)}")
    b1 = next(v for k, v in resumo.items() if str(k).startswith("B1."))
    w(
        f"- Atenção: quando você aprovou a Fase 0, a base B mostrada era {money(b1)} (só a seção do Hotel). A revisão da planilha-mestre incluiu a seção da Agropecuária (o trator)."
    )
    w("")
    w("## Legenda")
    w("")
    w(
        "- **C0**: critério de valor — bem da lista de Rafael (vem para nós) fica com o menor valor; bem que fica na empresa, com o maior."
    )
    w("- **P5**: sem reavaliação própria; adota-se um dos dois valores.")
    w(
        "- **Parte 6, candidato nº N**: suspeita de valor contado ou descontado duas vezes (`parte6-candidatos.md`)."
    )
    w("- **Q1–Q13**: perguntas ao grupo em rascunho (`perguntas-ao-grupo.md`).")
    w("- **RDB-1**: registro de data-base (`registro-data-base.md`).")
    w("- **arq. 1**: respostas do grupo ao Dr. Pablo, 13.08.2026.")
    w("")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    print(
        f"R1 {len(r1)} {total_r1}; R2 {len(other)}; R3 {len(soy)}; exceções {len(exc)}; destino {len(dest)}; pares {len(paired)}"
    )


if __name__ == "__main__":
    main()
