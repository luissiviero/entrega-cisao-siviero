# Revisão independente — Parte 4 SEI, verificação da v0.1

Revisor: agente de contexto novo (P11). Data: 28/10/2026.
Objeto: `_scripts/build_fluxo_caixa.py` (v0.1), `fluxo-caixa-sei-v0.1.xlsx` e `numeros-v0.1.json`. A comparação foi feita contra `build_fluxo_caixa_v0.py` e contra as fontes.
Método: importei as funções da v0.1 e as rodei com variações; comparei a classe de cada saída entre a v0 e a v0.1; refiz à mão as contas das abas 5 e 6, da Ponte e da Sensibilidade.
Efeitos em R$, na ordem piso / central / teto.

## Resultado por achado da v0

| Achado | Situação | Comentário |
|---|---|---|
| A1 | **Corrigido em excesso** | ver N1 |
| A2 | OK | Exatamente os 5 registros pedidos foram marcados como 2025: Morada O01; Colina do Retiro II 1/7, 12/8, 15/8; lote urbano A.1. Os 8 que continuam sem data estão certos. |
| A3, M1, B1 | OK | Perini 95.000 / 160.000; Lago Negro 247.143; lotes urbanos 90.000 (A.2 a A.4). |
| A3 (nova regra THIN_2025) | **Erro novo** | ver N2 |
| A4 | OK | Concessão = 2.450.000 (Itens seq. 111, coluna custo) − 2.000.000 (E737) = 450.000. A máxima da v0 continua na Sensibilidade. |
| A5 | Parcial | ver N3 e N4 |
| A6 | OK | Ver o bloco "Dívida" abaixo. |
| M2 | OK, com detalhe | Bens novos de 1.670.043,76 (linha 54 entra por 303.044, que é o dinheiro dos 15 meses; linha 76 entra por 0). Bens de 2018 sem destino: 1.256.500, recalculado. O item "móveis" (10.000, linha 65) tem destino declarado ("permaneceram como forma de pagamento pelas obras de conservação da sala"), então deveria sair de G2. Efeito: −10 mil. |
| M3 | OK | M de 15 meses = 4.998.032,98 − 150.000 (Flávia) − 186.000 (Hotel Araçá) = 4.662.032,98, ou 310.802,20 por mês. Pró-labore dos sócios: 558.000. Ver N5 sobre a Flávia. |
| M4 | OK | O início em 2019 também retira as saídas de 2018, e custos e aluguéis passam a 72 meses. |
| M5 | Parcial | Valdir Angest foi para ACORDO, Attolini para VENDA e os empréstimos para DIVL. Ver N6 para o que sobrou. |
| M6 | OK | A linha A3c só soma ACORDO ligado a loteamentos contratados. Na prática não há OBRA ligado (conferi: 0), então a linha I não conta duas vezes. |
| M7 | Parcial | Restam 6 registros, dos quais 3 não são o valor do próprio lote (ver N7). A amostra foi rebaixada a "consulta", o que basta. |
| B5 | OK | A asserção da trajetória está ativa e o placeholder foi removido. A trajetória termina em 425 lotes, não em 426, porque Reserva Paradiso tem redução −1. É só estético. |

### Dívida (aba 5): conferida
- Saldo de jan/2019 sem os itens 26, 30 e 32 e sem os 4 itens do Irineu: 7.320.900.
- Dívida em 31.12.2024: 2.367.652,03 − 320.000 (D6N) − 347.652,03 + 40 × 10.534,91 = 2.121.396,40.
  As 40 parcelas também batem com os relatórios: a parcela P.06/48 foi paga em out/2024, então em 31/12/2024 a P.08 estava paga e faltavam 40.
- C2 = 7.320.900 − 2.121.396,40 + 500.000 = 5.699.503,60.

### Ponte e Sensibilidade: conferidas
- Os três cenários conferem linha a linha: B − C = D, D + E = F, F + G1 + G2 = H, H/6, e I = H + A2 + A3c.
- O total A é igual à soma das 7 classes.
- Os lotes fecham: 273,67 no período + 31,33 depois de 2024 + 31 antes de 2018 = 336.
- Os 4 cenários da Sensibilidade foram recalculados.

---

## Pontos que ainda estão errados ou que surgiram na v0.1

### N1 (alta). A1 foi aplicado como "anteriores a 2018 primeiro", e isso tira da ponte vendas do período que estão documentadas
- **A culpa é do meu texto da v0.** A frase dava a entender que os registros anteriores a 2018 consomem a redução antes de tudo. A medição que fiz na v0, porém, usava outra regra: esses registros só ficam com a parte da redução que os registros de 2018–2025 não cobrem, isto é, `min(registros anteriores, max(0, redução − registros utilizáveis))`.
- **O que a v0.1 faz.** Põe 31 lotes "antes de 2018", em vez de 16. Os 15 a mais saem de loteamentos onde os registros do período já cobrem a redução: Lago Negro (2), Alto das Graças (7), Salvador (2), Monte Vêneto I–II (2) e Colina do Retiro II (2). No caso do Salvador, a própria aba diz que as vendas de 1992 e 2006 "provavelmente estão entre os 4 lotes vendidos já descontados em 2018".
- **Efeito.** VENDA sobe 0,91 / 1,46 / 2,41 mi, e ACORDO, OBRA e A3c sobem 0,25 / 0,38 / 0,65 mi. H sobe pelo menos o valor de VENDA.
- **Efeito colateral.** O excedente da Sensibilidade fica inflado: 58 lotes, em vez de cerca de 46.
- **Correção.** Em `allocate`, usar `pre_n = min(len(pre), max(0, red − len(usable)))`.

### N2 (alta). Erro novo: a regra THIN_2025 substitui valores de 2025 do próprio grupo por um índice
- **O que a regra faz.** Uma média de 2025 com menos de 3 lotes passa a ser trocada por 2,032 × o valor de 2018.
- **Por que está errado no Morada do Ipê I.** O grupo dá dois valores em 2025: O14/2 a 629.000 e O01/2 a 538.000 (GERAL, linhas 117 e 118). O modelo passa a usar 304.829. Isso é uma reavaliação que não vem do grupo, contra a premissa P5. A revisão não pediu essa regra: pediu só para conferir o Morada.
- **Efeito.** Comparando com a regra desligada, VENDA cai 0 / 3,01 / 6,15 mi e o total cai 0 / 3,89 / 7,91 mi. O teto é o cenário de referência.
- **Correção.** Retirar a regra. Se o Luís quiser um ajuste para o Morada, usar a média dos dois valores do grupo (583.500) e registrar como decisão dele.

### N3 (média). A5: a regra "vendido à vista/à prazo … permuta/parte/espécie" põe o registro inteiro em VENDA, inclusive bens que estão em 2025 ou que não são dinheiro
- Monte Vêneto 09 e 10/01 (linhas 332 e 333) foram trocados pelo lote 03/04 do Femaçã II, que **está** no estoque de 2025, a R$ 189.000 (GERAL linha 187). É troca, não venda: deve voltar para BEM. Efeito: cerca de −0,14 / −0,17 / −0,18 mi.
- Partes que não entraram como dinheiro:
  - paisagismo do condomínio (linha 82): 150 mil, que é OBRA;
  - móveis da sede (linha 97): 80,7 mil, que é OBRA;
  - caminhão entregue ao fornecedor Gilberto Portella (linhas 280 e 281): 110 mil;
  - SW4 "entregue para o sócio Raul Siviero" (linhas 449 e 450): 130 mil por registro, que é SOC.
- Efeito somado no central: cerca de −0,6 a −0,75 mi em VENDA.
- **Correção.** Dividir o registro: em VENDA, só a parte em dinheiro e os bens que não estão em 2025; o resto na classe certa.

### N4 (média). A5: as terras e casas recebidas por lotes continuam como BEM "explicado", mas não estão em 2025
São estes registros: Vila Azul (linhas 120 e 448, R$ 714.000), 0,67 ha do Ivo Frainer (linha 135, R$ 388.000), 7 ha do Antônio Ansolin (linha 143, R$ 586.666), casas da Arquidea (linhas 402 e 405), lote do Pasqual (linha 313) e móveis da Dhejota para a sede do Morada II (linha 80, que é OBRA).

A classe BEM ainda soma 1,41 / 2,09 / 2,92 mi, quase toda nesses registros. A nota da Ponte ("explicado se o bem estiver no patrimônio") continua sem verificação. **Correção:** listar esses itens na aba 6 como "bem recebido que não está em 2025". Depois, pedir ao grupo ou decidir se viram dinheiro em G2.

### N5 (baixa). A Flávia Manica entra por 500.000, o valor total do acordo
- Em jul/2023 foi paga a parcela 26 de 50, então até dez/2024 foram pagas 43 parcelas, ou 430.000. As outras 7 parcelas (70.000) são de 2025.
- **Correção.** Usar 430.000. Efeito em H: +70 mil.

### N6 (baixa). SOC, OBRA e DIVL: resíduos
- Linha 315, Carraro (gratificação a funcionário por 40 anos): agora está em **OBRA**. Isso infla A2 e a linha I em cerca de 0,07 / 0,10 / 0,15 mi. Deveria ser outra classe, ou SOC.
- Linha 85, comissão da Structura (Rovílio): continua em SOC.
- Linhas 386, 387 e 389, troca por área em Monte Bérico (loteamento contratado): continuam em SOC e fora de A3c.
- Linha 172, André (DIVL, jan/2024): o empréstimo de 100.000 do André continua na dívida de 2019. Os relatórios mostram uma devolução de 100.000 em dinheiro em nov/2023. Se o lote pagou esse mesmo empréstimo, os 100 mil são contados duas vezes. Confirmar com o grupo.
- A regra de "contratado" pega "intermediacao", que é a comissão do Zanella para o Morada I e II (o I está concluído), e "susin", que pega a comissão do Viero 05/1469. Nenhum dos dois é distrato.

### N7 (baixa). A amostra de preços citados ainda tem registros que não são o valor do lote
- Sol Nascente I 24/D: os 67.000 são o veículo ("sendo R$ 67.000,00 em permuta veículo").
- Alto das Graças 15/6371: os 465.000 são o contrato inteiro de pavimentação.
- Solar Norte 06/01 e Pedra Lisa 03/03: houve também parcelas em dinheiro nos 15 meses.

### N8 (baixa, informação)
- A concessão documentada cobre só o Falavigna. As obras de conclusão pagas nos 15 meses em loteamentos que estavam "executados e não registrados" em 2018 (Monte Vêneto, 289 mil; Barbiero, 57 mil) não entram nos usos. Isso favorece o sócio. Deixar escrito.
- Linha 76 (Sapopema): entra por 0. A parte em dinheiro dos 3.100 m² (parcelas ao Lindolfo Lago nos 15 meses) também não entra nos usos. Também favorece o sócio, e o valor é pequeno.

## Verificações sem problema
- Marcação de 2025 dos registros sem data.
- Todas as mudanças de classe entre a v0 e a v0.1: 35 registros do período, listados e conferidos um a um.
- Contas das abas 5 e 6.
- Teste de receita: COFINS de 330.650,31 (inclui uma linha "COFINS" de 13.316) e PIS de 71.640,91, ambos dando 11,02 mi; comissões dando 6,79 mi.
- Aritmética da Ponte e da Sensibilidade.
- Contratos 11.2 não mudaram.
- O total de redução no JSON é 335.
