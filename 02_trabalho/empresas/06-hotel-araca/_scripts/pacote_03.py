"""Packet 3 for company 06: destination 2018 -> 2025 of every asset that
existed in 2018 and is not in the 2025 inventory (P4), plus the related
findings of the search (assets outside both inventories).

Reads the master sheet (read only), the search output busca-destinos-06.xlsx,
the 2015 inventory, the inventory version kept in '2018_desatualizado',
doc 25 of 2018, file 18 and the 2017 meeting notice. Writes
pacotes/pacote-03-destinos.md. No value is proposed: 2018 is a benchmark of
existence (P3); an asset without a documented destination is an open item,
not a duplicate (P4). Evidence comes from the master sheet notes (all of
them, not a filtered subset); the proposal per asset is written here and
keyed by inscrição.

Usage: python pacote_03.py
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
BUSCA = HERE / "_extracoes/busca-destinos-06.xlsx"
OUT = HERE / "pacotes/pacote-03-destinos.md"
ANTIGO = f0.DRIVE / "Material antigo"

PROPOSTA = {
    "360.0019.004.000": 'Pendente. Indício de bem do sócio: na lista de IPTU de Válter e família (doc 6); o arquivo 18 registra um "Lote Capão da Canoa" na esquina perto do Hotel entre os "lotes dados aos sócios" (2019), sem inscrição. Sem ato de transferência.',
    "360.0057.002.000": 'Pendente. O doc 25 atribui a "CASA AO LADO" a Raul por acordo entre os sócios, com manutenção paga pelo Hotel; sem ato de transferência. Provável uma das duas casas citadas por Rafael (E2; G8; candidato 5 da Parte 6).',
    "360.0057.001.000": "Pendente, com leituras em conflito (abaixo). Provável a outra casa citada por Rafael (E2; G8; candidatos 2 e 5 da Parte 6).",
    "360.0074.030.000": "Pendente. Nenhum documento além do inventário 2018.",
    "360.0074.031.000": "Pendente. Nenhum documento além do inventário 2018.",
}


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()[:10]


def inv_rows(path: Path):
    out = []
    for _, cells in f0.table_rows(path):
        if cells and re.fullmatch(r"\d{2}", cells[0]) and len(cells) > 1:
            d = cells[1]
            out.append(
                {
                    "item": cells[0],
                    "insc": f0.norm_inscricao(
                        f0.first(r"Inscrição Imobiliária nº ([\d\.]+\d)", d)
                    ),
                    "mat": f0.first(r"Matrícula nº ([\d\.]+\d)", d),
                }
            )
    return out


def cell(s: str) -> str:
    return s.replace(" | ", "; ").replace("|", ";")


def main():
    wb = openpyxl.load_workbook(MESTRE, data_only=True)
    hdr = [c.value for c in wb["Mestre"][1]]
    rows = [
        dict(zip(hdr, r)) for r in wb["Mestre"].iter_rows(min_row=2, values_only=True)
    ]
    so18 = [r for r in rows if r["Situação"] == "Só 2018"]
    assert len(so18) == 5
    cofre_row = next(r for r in rows if "cofre" in (r["Descrição"] or ""))
    inv18 = f0.parse_inventory(f0.SRC["inv2018"], "2018")[0]
    by_insc18 = {i["inscricao_norm"]: i for i in inv18 if i["inscricao_norm"]}
    assert set(PROPOSTA) == {
        re.search(r"insc\. ([\d\.]+)", r["Chave"]).group(1) for r in so18
    }

    p2015 = next((ANTIGO / "inventário-2015").glob("HOTEL ARA*.docx"))
    pdraft = next((ANTIGO / "2018_desatualizado").glob("4.HOTEL ARA*.docx"))
    preun = next((ANTIGO / "2018_desatualizado").glob("G-SÓCIOS*.docx"))
    r2015, rdraft = inv_rows(p2015), inv_rows(pdraft)
    mats2015 = {r["mat"] for r in r2015 if r["mat"]}
    # 2015 lists items 02-19 of 2018 by matrícula (item 02 under its old number)
    assert mats2015 == {i["matricula_antiga"] or i["matricula"] for i in inv18[1:19]}
    draft_by_insc = {r["insc"]: r for r in rdraft if r["insc"]}

    bw = openpyxl.load_workbook(BUSCA, read_only=True)
    hits = list(bw["Achados"].iter_rows(min_row=2, values_only=True))
    counts = {
        r[0]: (r[1], r[2])
        for r in bw["Contagem"].iter_rows(min_row=2, values_only=True)
    }
    errs = list(bw["Erros"].iter_rows(min_row=2, values_only=True))
    n_unique = len({h[2] for h in hits if h[2] != "ocr"})
    roots = [f0.DRIVE.parent, f0.BASE / "01_fontes"]
    doc_old = sorted({p.name for root in roots for p in root.rglob("*.doc")})
    img = sorted(
        {
            p.name
            for root in roots
            for p in root.rglob("*")
            if p.suffix.lower() in (".jpg", ".jpeg", ".png")
        }
    )

    def files_for(prefix):
        return sorted({Path(h[1]).name[:60] for h in hits if h[0].startswith(prefix)})

    inv16 = (
        files_for("mat 77.833") + files_for("mat 74.857") + files_for("insc 360.0074")
    )
    assert len(set(inv16)) == 1  # Adélia's and Pedro's lots appear only in doc 16

    reun = "\n".join(
        b
        if k == "p"
        else "\n".join(" || ".join(c.text for c in r.cells) for r in b.rows)
        for k, b in f0.doc_blocks(preun)
    )
    reun = re.sub(r"[ \t]+", " ", reun)
    casa_txt = re.search(
        r"pode ser agregado a casa, EXCETO os lotes adquiridos adjacentes", reun
    ).group(0)
    coletto = re.search(
        r"U\$ ([\d\.]+,\d{2}) – Pago Coletto–Gramado – 02 lotes, em (\d{2}\.\d{2}\.\d{4})",
        reun,
    )
    cruzeiro = re.search(
        r"U\$ ([\d\.]+,\d{2}) – Válter Siviero – Viagem Cruzeiro, em (\d{2}\.\d{2}\.\d{4})",
        reun,
    )
    peru = re.search(
        r"U\$ ([\d\.]+,\d{2}) – Luciane Siviero – Viagem Peru, em (\d{2}\.\d{2}\.\d{4})",
        reun,
    )
    cofre = re.search(
        r"U\$ ([\d\.]+,\d{2}) – TOTAL NO COFRE EM VERANÓPOLIS", reun
    ).group(1)
    parts = [
        f0.brl(x) for x in re.findall(r"U\$ ([\d\.]+,\d{2}) – COFRE VERANÓPOLIS", reun)
    ]
    assert abs(sum(parts) - f0.brl(cofre)) < 0.005
    gramado6 = [
        ln
        for ln in re.split(
            r"\n",
            "\n".join(
                " | ".join(x.replace("\n", " ") for x in c)
                for _, c in f0.table_rows(
                    next((f0.DRIVE / "2018/inventário-2018").glob("6. 1. VALTER*.docx"))
                )
            ),
        )
        if "Lago Negro" in ln
    ]

    t25 = "\n".join(
        b if k == "p" else ""
        for k, b in f0.doc_blocks(
            next((f0.DRIVE / "2018/inventário-2018").glob("25. 1.*.docx"))
        )
    )
    pa = next(
        (ANTIGO / "2018_desatualizado").glob(
            "A-PLANILHA = DIVISÃO DE LOTES=ANO  2015*.docx"
        )
    )
    t15 = " ".join(" || ".join(c) for _, c in f0.table_rows(pa))
    busatto15 = re.search(
        r"Lote 10 – Quadra 03 – Loteamento Aditivo Madianeira, que conforme contrato foi dado para o Paulo Busatto\. \|\| R\$\s+([\d\.]+,\d{2})",
        t15,
    )
    assert busatto15 and "CASA DA ESQUINA" in t25

    L = []
    w = L.append
    w("# Hotel Araçá (empresa 06) — Pacote 3: destino 2018 → 2025")
    w("")
    w(
        f'Data: 28/10/2026. Gerado por `_scripts/pacote_03.py` a partir de `planilha-mestre-06.xlsx` (md5 `{md5(MESTRE)}`) e da busca `_extracoes/busca-destinos-06.xlsx` (md5 `{md5(BUSCA)}`). Revisado por revisor independente (`_diario/revisao-pacote-03.md`) antes desta versão. Nada está aprovado; a linha dos dólares já foi gravada na mestre como "Pendente", sem decisão, e sai se você não aprovar.'
    )
    w("")
    w(
        "Regra (P3 e P4): o inventário de 2018 é referência de existência, não de valor. Todo bem de 2018 que não está em 2025 precisa de destino documentado; sem documento, é pendência, não duplicidade."
    )
    w("")
    w("## Busca feita")
    w("")
    w(
        f'- **Lido por texto:** os .docx, .xlsx, .md e as páginas com texto dos .pdf da pasta do Drive `Cisao_Siviero` e de `01_fontes` ({n_unique} arquivos distintos com algum achado), e o OCR revisado das respostas ao Dr. Pablo. Termos: as inscrições, matrículas e códigos dos cinco lotes, Busatto, casas de veraneio, "casa da esquina"/"casa ao lado", Rua da Divisa, Fiat Toro, esquina, consignado, Coletto, dólar/cofre, Capão da Canoa e Hotel Araçá (script `_scripts/busca_destinos_06.py`). Lote e quadra escritos por extenso: {counts.get("lote/quadra das inscrições (04/q19, 01-02/q57, 30-31/q74)", (0, 0))[0]} achados.'
    )
    w(
        f"- **Lido pelo conector do Drive:** o doc 22 de 2018 (ata de 14/01/2019), que não está baixado no disco ({len(errs)} erro de leitura). O mesmo texto está na convocação de 26/05/2017, que está no disco e foi lida (`{preun.name[:60]}…`)."
    )
    w(
        f"- **Não lidos por mim:** PDFs escaneados fora do OCR (por exemplo, ACERTOS E ACORDOS, considerações de 2019, retiradas dos sócios de 2018, relatórios mensais), {len(img)} imagens (.jpg/.png) e {len(doc_old)} .doc antigos (cessões de quotas e contrato de compromisso mútuo de 2003). O revisor leu como imagem a ata de 14/01/2019, o relatório do hotel de 2025 e as retiradas de 2018, e pelo conector dois dos .doc; nenhum traz destino dos cinco lotes."
    )
    w(
        '- **Resultado:** nenhuma escritura, venda ou permuta dos cinco lotes depois de 2018. Há uma possível entrega a sócio: o arquivo 18 (planilha de Válter "para acerto futuro") registra, no bloco de 10.06.2019, um "Lote Capão da Canoa", "lote na esquina próximo ao Hotel Araçá Ltda., com valor diferenciado aos demais lotes dados aos sócios" (p. 3, sem inscrição e sem valor). Os lotes de Adélia e Pedro (inscrições 360.0074.030 e .031, matrículas 77.833 e 74.857) só aparecem no inventário de 2018.'
    )
    w("")
    w("## Os cinco lotes de 2018 sem par em 2025")
    w("")
    w(
        '| # | Lote em 2018 | Nome no lugar do valor | Inventário 2015 | Versão na pasta "2018_desatualizado" | Evidências (planilha-mestre) | Proposta |'
    )
    w("|---|---|---|---|---|---|---|")
    for n, r in enumerate(so18, 1):
        k = re.search(r"insc\. ([\d\.]+)", r["Chave"]).group(1)
        it = by_insc18[k]
        mat = it["matricula"] or "sem matrícula"
        in2015 = (
            "não" if not (it["matricula"] and it["matricula"] in mats2015) else "sim"
        )
        dr = draft_by_insc.get(k)
        dr_txt = f"item {dr['item']} do Hotel, sem nome e sem valor" if dr else "não"
        notes = [
            x
            for x in r["Notas de identificação"].split("; ")
            if not x.startswith(
                ("no lugar do valor", "bem da empresa ou do sócio", "no rascunho")
            )
        ]
        w(
            f"| D{n} | {it['item']}: insc. {k}, {it['area_m2']} m², {mat}{', com casa' if 'CASA' in it['descricao'] else ''} | {it['nome_socio_no_valor']} | {in2015} | {dr_txt} | {cell(' / '.join(notes))} | {PROPOSTA[k]} |"
        )
    w("")
    w(
        f'Inventário de 2015: tem {len(r2015)} itens, os mesmos 01 a 19 de 2018 (pela matrícula); nenhum dos cinco lotes. A versão da pasta "2018_desatualizado" se chama "… Dezembro de 2015", mas a data real é incerta: tem os lotes de Válter, Irineu, Élcio e Raul como itens comuns do Hotel, e a versão final de 2018 põe o nome do sócio no lugar do valor.'
    )
    w("")
    w("**D3, as leituras possíveis da troca de Élcio:**")
    w(
        f'1. **A casa é bem de empresa e falta em 2025.** O inventário 2018 manda "(Avaliar Lote e Casa)" e o total diz "+ o valor da casa e lote – Élcio Siviero"; o doc 25 diz "CASA DA ESQUINA: Élcio Siviero (que pertence à empresa Siviero)"; a divisão de lotes de 2015 lançava R$ {busatto15.group(1)} contra Élcio pelo lote dado a Paulo Busatto, e em 2018 esse valor fica em branco, com a nota "Trocado com o lote de Capão da Canoa – CASA". Nessa leitura, Élcio entregou a casa à empresa em troca do lote da SEI, e é isso que Rafael aponta na p. 24. Pelo C0, se a casa ficar na empresa, entra com o maior valor.'
    )
    w(
        '2. **A casa é de Élcio.** Ele recebeu a casa de Busatto, e o lote da SEI foi o pagamento, lançado como capital pago pelas empresas a ele (doc 25, seção Élcio). Contra: o doc 25 diz que a casa "pertence à empresa Siviero".'
    )
    w(
        "3. **Troca a favor do sócio, com a casa saindo da empresa.** Era a leitura da versão anterior deste pacote. Não se sustenta sem documento."
    )
    w(
        'Proposta: não escolher agora. Pergunta G7: qual é a "empresa Siviero" dona da casa e quem é o titular hoje.'
    )
    w("")
    w(
        '**Lotes que os sócios têm em Capão por conta própria.** As listas de IPTU de 2018 de Élcio (360.0059.005.000) e de Irineu (360.0058.001.000 e "Circo") trazem outros imóveis em Capão, perto da quadra 0057. São bens dos sócios, fora do Hotel, e não destino dos lotes acima. O doc 5 é a lista de José Valério, que figura com Adélia nos quadros societários.'
    )
    w("")
    w(
        "**Efeito no valor.** Nenhum agora (P3). Se algum desses lotes, ou as casas, for da empresa em 31/12/2024, tem de entrar no inventário 2025 com valor; pelo C0, um bem que fica na empresa leva o maior valor. Se for do sócio, o destino é a entrega ao sócio, que precisa ser documentada e lançada na conta dele (Parte 4)."
    )
    w("")
    w("## Outros achados")
    w("")
    a18 = f0.item01_components(inv18[0]["descricao"], "2018")[3]["area_m2"]
    a25 = f0.item01_components(
        f0.parse_inventory(f0.SRC["inv2025"], "2025")[0][0]["descricao"], "2025"
    )[3]["area_m2"]
    w(
        f"- **D6. Área construída do hotel.** Era {a18} m² em 2018 e é {a25} m² em 2025. A diferença de 1.500 m² não tem destino documentado (demolição, erro de medição ou área equivalente). Pergunta G3. Pendente."
    )
    w(
        f'- **D7. Dólares do Hotel no cofre em Veranópolis.** O anexo "Planilha – resumo dos dólares" da convocação de 26/05/2017, repetido na ata de 14/01/2019 (doc 22 de 2018), registra U$ {cofre} "total no cofre em Veranópolis" em 25/05/2017. Não estão no inventário 2018 nem no de 2025. Por estarem num documento do conjunto de 2018, trato como bem existente em 2018 que precisa de destino (P4), sem valor por ora. O mesmo anexo registra saídas:'
    )
    w(
        f'  - U$ {coletto.group(1)} "Pago Coletto–Gramado – 02 lotes" ({coletto.group(2)}). A lista de IPTU de Válter de 2018 (doc 6) traz lotes no Lago Negro, em Gramado: {"; ".join(cell(x) for x in gramado6)}. Não está provado que sejam os mesmos.'
    )
    w(
        f'  - U$ {cruzeiro.group(1)} para Válter Siviero, "Viagem Cruzeiro" ({cruzeiro.group(2)}), e U$ {peru.group(1)} para Luciane Siviero, "Viagem Peru" ({peru.group(2)}).'
    )
    w(
        f'  Situação na mestre: "{cofre_row["Situação"]}", destino "{cofre_row["Destino 2018→2025"]}".'
    )
    w(
        f'- **D8. Proposta de divisão de 2017.** A convocação de 26/05/2017 (repetida em 2019) diz, para o sócio que ficar com o Hotel: "{casa_txt}". Não diz qual casa. A exclusão dos "lotes adquiridos adjacentes" interessa à lista de Rafael, que escolheu cinco lotes vizinhos ao hotel. Só registro.'
    )
    w("")
    w("## Perguntas e candidatos propostos")
    w("")
    w(
        '- **G7 (refeita, junta os cinco lotes e as duas casas; a G8 fica só para a Fiat Toro):** de quem são hoje os lotes das inscrições 360.0019.004.000, 360.0057.001.000, 360.0057.002.000, 360.0074.030.000 e 360.0074.031.000 (matrículas 77.833 e 74.857)? Qual foi o destino de cada um desde 2018? O "Lote Capão da Canoa" do arquivo 18 (2019) é qual deles? As duas casas da esquina da Rua da Divisa com a Rua Três são os lotes 360.0057.001 e .002? Qual é a "empresa Siviero" dona da casa da esquina? Quem paga a manutenção das casas?'
    )
    w(
        f"- **G12 (nova):** qual foi o destino dos U$ {cofre} do Hotel guardados no cofre em Veranópolis em 25/05/2017? Os 2 lotes de Gramado pagos em 2009 com U$ {coletto.group(1)} do Hotel estão em nome de quem?"
    )
    w(
        "- **Candidato 6 da Parte 6:** dólares do Hotel usados para lotes em Gramado (2009) e entregues a Válter e Luciane (2012–2013): saídas do Hotel sem contrapartida registrada."
    )
    w(
        '- **Candidato 7 da Parte 6:** o lote de Capão "dado aos sócios" no arquivo 18 (2019), lançado na conta de Válter. Se for um dos lotes do Hotel que ainda estão no inventário 2025 (por exemplo o item 02, lote 03/q19, da lista de Rafael, vizinho do lote D1), o mesmo bem estaria entregue e escolhido duas vezes.'
    )
    w("")
    w("## O que peço para aprovar")
    w("")
    w(
        '1. Registrar D1 a D5 como **pendentes**, com as evidências acima na coluna "Destino 2018→2025" da mestre. D3 fica com as três leituras, sem escolher.'
    )
    w("2. Registrar D6 como pendente (G3).")
    w(
        "3. D7: manter a linha dos dólares na mestre, sem valor, com destino pendente (P4)."
    )
    w(
        "4. Refazer a G7, mudar a G8 e acrescentar a G12; acrescentar os candidatos 6 e 7 da Parte 6."
    )
    OUT.write_text("\n".join(L), encoding="utf-8")
    print(
        f"unique {n_unique}; doc_old {len(doc_old)}; img {len(img)}; cofre {cofre}; gramado6 {len(gramado6)}; busatto15 {busatto15.group(1)}"
    )
    print("saved", OUT)


if __name__ == "__main__":
    main()
