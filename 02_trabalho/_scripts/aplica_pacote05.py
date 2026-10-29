"""Write the decisions of packet 5 (approved by Luís on 28/10/2026) into decisoes-sei.json.

Approval, in Luís's words: "Aprovo como recomendado". That covers: rule R6,
its extension to the third-party liability, the intermediate reading of the
30 m strip, Simioni at level B (R$ 260.000), and the possible-APP lots as
assets staying in the company (Rafael's list = PDF Diretrizes).
Values come from pacote05.py (reading "intermediaria"); nothing is typed by hand.

Usage: python aplica_pacote05.py
"""

from __future__ import annotations

import contextlib
import io
import json
import sys
from pathlib import Path

sys.argv = [sys.argv[0], "intermediaria"]
import pacote05

W = Path("C:/entrega-cisao-siviero/02_trabalho")
DEC = W / "decisoes-sei.json"
DATE = "2026-10-28"
ITEM_OF = {  # development -> item key in the master sheet
    "Fracionamento Simioni": "SEI-2025-LOT-04",
    "San Marino": "SEI-2025-LOT-20",
    "Firminio Dalzóchio": "SEI-2025-LOT-21",
    "Primavera": "SEI-2025-LOT-22",
    "Industrial Vicentino": "SEI-2025-LOT-24",
}
NIVEL = {"A": "A (0 %)", "B": "B (50 %)", "C": "C (100 %)"}


def main() -> None:
    with contextlib.redirect_stdout(io.StringIO()):
        ativos, passivos = pacote05.main()
    dec = json.loads(DEC.read_text(encoding="utf-8"))
    mestre = {
        r["Chave"]: r for r in pacote05.rows(W / "planilha-mestre-sei.xlsx", "Mestre")
    }

    dec["regras"]["R6"] = {
        "aprovada_em": DATE,
        "pacote": 5,
        "decisao": "Regra única de 'possível APP' por nível (A 0 %, B 50 %, C 100 % do lote cheio; indenização a terceiro pelo complemento)",
        "motivo": (
            "Pacote 5 (aprovado como recomendado). Regra de decisão sob o C2, não reavaliação (P5: lote cheio = valor do grupo ou de Rafael). "
            "A: APP reconhecida por decisão, perícia ou gravame, ou >= 50 % do lote em APP ou faixa não edificante no mapa municipal (Lei 4.814/2023, GEOFarroupilha). "
            "B: 10 % a 50 % nessas camadas; ou lote na faixa de 30 m em volta dos cursos d'água ('Sensibilidade ambiental') com >= 10 % (leitura intermediária); ou, sem mapa, restrição documentada (TAC). "
            "C: menos de 10 % e nenhum documento. Estendida ao passivo de lotes de terceiros (A 100 %, B 50 %, C 0 %) com aprovação expressa. "
            "Mapa consultado em 28/10/2026 (versão atual); texto da lei não obtido; se a lei confirmar que a faixa de 30 m não restringe, rever para a leitura estrita."
        ),
        "leitura": "intermediaria",
    }

    lots = dec["lotes"]
    totals: dict[str, float] = {}
    for lt, q, l_, lvl, vg, vr, va, why in ativos:
        lots[f"{lt}|{q}|{l_}"] = {
            "aprovada_em": DATE,
            "pacote": 5,
            "decisao": f"R6, nível {NIVEL[lvl]}",
            "motivo": f"{why}. Lote cheio: {'Rafael' if lt == 'Fracionamento Simioni' else 'grupo = Rafael'} R$ {vr:,.0f}. Lote fica na empresa (lista de Rafael = PDF Diretrizes).",
            "valor_adotado": va,
        }
        totals[lt] = totals.get(lt, 0.0) + va
    for lt, q, l_, lvl, vg, va, why in passivos:
        if lvl == "S1":
            dec_txt, motivo = (
                "Retirado do passivo de terceiros (S1)",
                f"{why}. S1 do ACERTOS (pacote 1, aprovado em {DATE}), aplicado no pacote 5; não voltar na Parte 6. Lado do ativo em aberto (registro de data-base).",
            )
        else:
            dec_txt, motivo = (
                f"R6 (passivo), nível {NIVEL[lvl]}: indenização de {int((1 - pacote05.SHARE[lvl]) * 100)} %",
                f"{why}. Valor listado R$ {vg:,.0f}.",
            )
        lots[f"{lt}|{q}|{l_}"] = {
            "aprovada_em": DATE,
            "pacote": 5,
            "decisao": dec_txt,
            "motivo": motivo + " Valor adotado = indenização (passivo), não bem.",
            "valor_adotado": va,
        }

    for lt, extra in totals.items():
        key = ITEM_OF[lt]
        base = (
            0.0
            if lt == "Fracionamento Simioni"
            else float(mestre[key]["Valor grupo 2025 (R$)"] or 0)
        )
        dec["itens"][key] = {
            "aprovada_em": DATE,
            "pacote": 5,
            "decisao": "Lotes pelo grupo + lotes de possível APP pela R6"
            if base
            else "R6, nível B (TAC)",
            "motivo": (
                f"Pacote 5. Subtotal do grupo R$ {base:,.2f} (lotes sem restrição, R1) + lotes de possível APP pela R6 R$ {extra:,.2f}."
                if base
                else "Pacote 5. Simioni: TAC com o MP reconhece nascente e restrição de construir; sem perícia nem mapa; 4 × 50 % de R$ 130.000 (Rafael). Alternativas apresentadas: R$ 0 (C2) e R$ 520.000 (valor cheio; 2018: R$ 1.000.000)."
            ),
            "valor_adotado": base + extra,
            **({"lotes": "valor do grupo"} if base else {}),
        }

    terc = sum(p[5] for p in passivos)
    dec["passivos"] = {
        "_sobre": "Passivos do doc 11 com valor adotado aprovado. Chave = início do texto do item na aba Passivos.",
        "Lotes de terceiros passíveis de indenização": {
            "aprovada_em": DATE,
            "pacote": 5,
            "decisao": "Indenização pela R6 (leitura intermediária); S1 retirado",
            "motivo": f"Pacote 5. Doc 11: R$ 16.240.000. S1 (4 lotes, R$ 1.887.000) retirado. Demais 22 lotes pela R6: R$ {terc:,.2f}. Detalhe por lote na aba Lotes.",
            "valor_adotado": terc,
        },
    }
    dec["criterios"]["C2"]["texto"] += " Aplicação: regra R6 (pacote 5, 28/10/2026)."
    DEC.write_text(json.dumps(dec, ensure_ascii=False, indent=2), encoding="utf-8")
    print("itens:", {k: dec["itens"][ITEM_OF[k]]["valor_adotado"] for k in totals})
    print(
        "passivo de terceiros adotado:",
        terc,
        "| lotes gravados:",
        len(ativos) + len(passivos),
    )


if __name__ == "__main__":
    main()
