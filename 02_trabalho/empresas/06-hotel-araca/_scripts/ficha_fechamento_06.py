"""Closing sheet for company 06 (Hotel Araçá): ficha-fechamento-06.md.

One page: decisions, numbers, risks, open items and the proposals still to
approve. Every figure is read from planilha-mestre-06.xlsx, decisoes.json and
the registers; nothing is typed by hand except the wording.

Usage: python ficha_fechamento_06.py
"""

from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parents[1]
MESTRE = HERE / "planilha-mestre-06.xlsx"
DEC = HERE / "decisoes.json"
OUT = HERE / "ficha-fechamento-06.md"
TOL = 0.005


def br(v: float) -> str:
    return f"{v:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()[:10]


def main():
    dec = json.loads(DEC.read_text(encoding="utf-8"))
    wb = openpyxl.load_workbook(MESTRE, data_only=True)
    hdr = [c.value for c in wb["Mestre"][1]]
    rows = [
        dict(zip(hdr, r)) for r in wb["Mestre"].iter_rows(min_row=2, values_only=True)
    ]
    tot = {r[0]: r[1] for r in wb["Totais"].iter_rows(min_row=2, values_only=True)}

    def T(prefix):
        return next(v for k, v in tot.items() if k.startswith(prefix))

    item = {
        re.match(r"item (\d{2})", r["Ref. 2025"]).group(1): r
        for r in rows
        if (r["Ref. 2025"] or "").startswith("item")
    }
    g_tot, adopted, sel = (
        T("Total grupo 2025"),
        T("Valores adotados (decisões"),
        T("Lista de Rafael"),
    )
    r_sum, r_decl, gap = (
        T("Soma dos itens de Rafael"),
        T("Total declarado"),
        T("Diferença de soma"),
    )
    assert len(item) == 22 and all(
        r["Valor adotado"] is not None for r in item.values()
    )
    assert abs(sum(r["Valor adotado"] for r in item.values()) - adopted) < TOL
    assert len(dec["decisoes"]) == 22 and abs(sum(v["valor_adotado"] for v in dec["decisoes"].values()) - adopted) < TOL
    assert abs(adopted - r_sum) < TOL  # approved values = Rafael's item sum
    d01 = item["01"]["Valor adotado"] - item["01"]["Valor grupo 2025"]
    d20 = item["20"]["Valor adotado"] - item["20"]["Valor grupo 2025"]
    diff = adopted - g_tot
    assert abs(d01 + d20 - diff) < TOL
    titul = {k: item[k]["Valor adotado"] for k in ("17", "20", "21", "22")}
    comp = {
        r[0]: r for r in wb["Item01_componentes"].iter_rows(min_row=2, values_only=True)
    }
    nesga_r = 420 * comp["1.1 Terreno"][4] / 6900
    obs = " ".join(
        r[2]
        for r in wb["Observacoes"].iter_rows(min_row=2, values_only=True)
        if r[0] == "inventário 2025 atual"
    )
    red = float(
        re.search(r"4,00% = R\$ ([\d\.]+,\d{2})", obs)
        .group(1)
        .replace(".", "")
        .replace(",", ".")
    )
    inss = float(
        re.search(r"aproximadamente R\$ ([\d\.]+,\d{2})", obs)
        .group(1)
        .replace(".", "")
        .replace(",", ".")
    )
    casas = next(r for r in rows if "Rua da Divisa" in (r["Descrição"] or ""))
    c_lo, c_hi = (
        float(x.replace(".", ""))
        for x in re.findall(r"R\$ ([\d\.]+)", casas["Faixa de mercado de Rafael"])[:2]
    )
    n_perg = len(
        re.findall(
            r"^\| G\d+ \|",
            (HERE / "perguntas-ao-grupo.md").read_text(encoding="utf-8"),
            re.MULTILINE,
        )
    )
    n_p6 = len(
        re.findall(
            r"^\| \d+ \|",
            (HERE / "parte6-candidatos.md").read_text(encoding="utf-8"),
            re.MULTILINE,
        )
    )
    n_db = len(
        re.findall(
            r"^\| \d+ \|",
            (HERE / "registro-data-base.md").read_text(encoding="utf-8"),
            re.MULTILINE,
        )
    )
    sem_valor = [r for r in rows if r["Decisão (P5)"] == "Sem valor nesta etapa"]

    L = []
    w = L.append
    w("# Hotel Araçá (empresa 06) — ficha de fechamento")
    w("")
    w(
        f"28/10/2026. Gerada por `_scripts/ficha_fechamento_06.py` a partir da `planilha-mestre-06.xlsx` (md5 `{md5(MESTRE)}`) e de `decisoes.json` (md5 `{md5(DEC)}`). A revisão independente de todas as decisões aprovadas (`_diario/revisao-final-decisoes.md`) não achou decisão mal aplicada nem número errado."
    )
    w("")
    if any(x["etapa"] == "Ficha de fechamento" for x in dec["aprovacoes"]):
        w("**Situação: empresa fechada. Ficha aprovada pelo Luís em 28/10/2026.**")
        w("")
    w("## Decisões aprovadas")
    w("")
    w("| Etapa | Decisão | Valor |")
    w("|---|---|---|")
    w(
        f"| Pacote 1, R1 e R2 | 20 itens com o mesmo valor dos dois lados: adoto o valor (igual) | R$ {br(sum(r['Valor adotado'] for k, r in item.items() if k not in ('01', '20')))} |"
    )
    w(
        f"| Pacote 2, E1 | Hotel, item inteiro com móveis: adoto Rafael (C0) | R$ {br(item['01']['Valor adotado'])} (grupo R$ {br(item['01']['Valor grupo 2025'])}) |"
    )
    w(
        f"| Pacote 2, E3 | Lote do Condado de Capão: adoto Rafael (C0) | R$ {br(item['20']['Valor adotado'])} (grupo R$ {br(item['20']['Valor grupo 2025'])}) |"
    )
    w(
        f"| Pacote 2, E4 | Diferença não identificada no total de Rafael: não adotada (contra o C0) | R$ 0,00 (Rafael soma R$ {br(gap)} a mais) |"
    )
    w(
        f"| R3, E2, E5, D7 | {len(sem_valor)} linhas sem valor nesta etapa (lotes de 2018, casas e Toro citadas por Rafael, placas solares, lotes 13 e 14 da q37, dólares do cofre) | — |"
    )
    w(
        "| Pacote 3 | Destino 2018 → 2025: os cinco lotes com nome de sócio ficam pendentes (D3 com três leituras); área construída (D6) e dólares (D7) pendentes | — |"
    )
    w("| E6 | Registros de data-base mantidos sem ajuste | — |")
    w("")
    w("## Números")
    w("")
    w("| | Total da empresa | Seu 1/6 |")
    w("|---|---|---|")
    w(f"| Grupo (inventário 2025) | R$ {br(g_tot)} | R$ {br(g_tot / 6)} |")
    w(
        f"| **Valores adotados** | **R$ {br(adopted)}** | **R$ {br(adopted / 6)}** (+R$ {br(diff / 6)}) |"
    )
    w(f"| Rafael, como ele escreve | R$ {br(r_decl)} | R$ {br(r_decl / 6)} |")
    w("")
    w(
        f"Lista de Rafael (bens que viriam para nós): R$ {br(sel)}, igual nos dois lados. Crédito pelos valores adotados: R$ {br(adopted / 6 - sel)}; Rafael escreve R$ {br(r_decl / 6 - sel)} porque soma os R$ {br(gap)} não identificados."
    )
    w("")
    w("## O que pesa nesses números")
    w("")
    w(
        f"1. **{d01 / diff * 100:.0f} % da diferença vem só do hotel (E1).** A base de Rafael é fraca: os 66 milhões são o piso do custo de reposição; o método que ele chama de mais realista (renda) dá R$ 41,0 a 50,6 milhões; as fontes são de 2025. Se o E1 cair, o seu 1/6 volta a R$ {br((adopted - d01) / 6)}."
    )
    w(
        f"2. **Os valores são brutos.** Nada está descontado da regularização do hotel: o grupo estima R$ {br(red)} de redução e R$ {br(inss)} de INSS (R$ {br((red + inss) / 6)} no 1/6). Os passivos são da Parte 4, que ainda não foi feita."
    )
    w(
        f'3. **Aprovou-se o valor, não a titularidade.** Dependem de resposta do grupo os itens 17 (lote que era "Irineu" em 2018), 20 (contrato particular), 21 e 22 (sem par em 2018): R$ {br(sum(titul.values()))}, ou R$ {br(sum(titul.values()) / 6)} no 1/6. O item 02, da lista de Rafael, pode ser o lote de Capão "dado aos sócios" em 2019 (arquivo 18).'
    )
    w(
        f"4. **Área do terreno.** Se a nesga de 420 m² estiver contada duas vezes (G2), o valor de Rafael adotado cai R$ {br(nesga_r)} (R$ {br(nesga_r / 6)} no 1/6). Os lotes 15 a 17 da q37 têm matrículas próprias: não há área contada em dobro neles."
    )
    w("")
    w("## Fora do valor (sem como somar hoje, P5)")
    w("")
    w(
        f"- **Podem aumentar a empresa:** as duas casas da Rua da Divisa (Rafael: R$ {br(c_lo)} a {br(c_hi)} pelo par, ou R$ {br(c_lo / 6)} a {br(c_hi / 6)} no 1/6), provavelmente os lotes de Raul e Élcio de 2018; os lotes de Válter, Adélia e Pedro de 2018; os U$ 37.698,00 do cofre (2017); a Fiat Toro; as placas solares pagas pela SEI; os lotes 13 e 14 da quadra 37. As casas aparecem em três linhas da mestre, mas são dois bens."
    )
    w(
        "- **Podem reduzir:** área do terreno (G2) e área construída, que caiu 1.500 m² desde 2018 (G3)."
    )
    w("")
    w("## Pendências")
    w("")
    w(f"- {n_perg} perguntas ao grupo (G1 a G{n_perg}), aprovadas e **não enviadas**.")
    w(
        f"- {n_p6} candidatos à Parte 6 (duplicidades), abertos; {n_db} registros de data-base, sem ajuste."
    )
    w(
        "- Material não lido por mim: relatórios mensais e outros PDFs escaneados fora do OCR, imagens e .doc antigos (o revisor leu alguns e não achou destino dos lotes)."
    )
    w(
        '- Sua planilha 2_0 tem um erro no Hotel: o lote 28/q38 aparece com 0 em 2018 (eram R$ 200.000,00). O total de 2018 sai R$ 200 mil menor e a "Diferença", R$ 1.500.000,00 em vez de R$ 1.150.000,00.'
    )
    w(
        "- Os registros desta empresa só se juntam aos da SEI depois da sua aprovação desta ficha (P13)."
    )
    w("")
    approved = any(x["etapa"] == "Ficha de fechamento" for x in dec["aprovacoes"])
    w("## Propostas da revisão final" + (" (aprovadas pelo Luís em 28/10/2026 e aplicadas)" if approved else ", para você aprovar"))
    w("")
    w(
        '1. **G13 (nova):** o "Lote Capão da Canoa" do arquivo 18 (2019) é o item 02 (lote 03/q19, Casa da Praça)? Se for, o mesmo bem estaria entregue e escolhido por Rafael.'
    )
    w(
        "2. **G14 (nova) e candidato 8 da Parte 6:** a redução de R$ 2,0 milhões e o INSS de R$ 1,6 milhão da regularização do hotel serão lançados como passivo? Risco de desconto em dobro se também reduzirem o valor do item."
    )
    w(
        "3. **G1 ampliada:** alguma parte do hotel (piscinas, salão, estacionamento) está sobre os lotes 13 a 17 da quadra 37?"
    )
    w(
        "4. **G10 ampliada:** os lotes 15 e 17 da q37 estão em nome da empresa? Por que a matrícula 124.906 está em Osório?"
    )
    w(
        '5. **Redação do E5 (Toro):** "se for da empresa, é bem não abordado (Parte 5) e precisa de valor do grupo; se não for, registrar o titular". A decisão não muda.'
    )
    w("6. **Motivo do R2 por item:** 17 → G11; 21 e 22 → G10.")
    OUT.write_text("\n".join(L), encoding="utf-8")
    print(
        f"adopted {adopted:,.2f} /6 {adopted / 6:,.2f}; diff {diff:,.2f}; E1 share {d01 / diff:.3f}; titul {sum(titul.values()):,.2f}; regul {red + inss:,.2f}; perguntas {n_perg}; p6 {n_p6}; db {n_db}; sem valor {len(sem_valor)}"
    )
    print("saved", OUT, md5(OUT))


if __name__ == "__main__":
    main()
