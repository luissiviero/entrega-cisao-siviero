# Hotel Araçá (empresa 06) — resumo da Fase 0

Data: 28/10/2026. Números extraídos por script (`_scripts/fase0_extracao_06.py` → `_extracoes/fase0-06.xlsx`, 70 conferências). Fontes com hash no diário desta pasta. Revisão independente feita antes deste resumo (`_diario/revisao-fase0.md`): nenhum número extraído errado nem bem perdido; os 18 achados de risco e de omissão foram corrigidos no script e neste texto.

## 1. Totais conferidos

| Fonte | Itens | Total | Conferência |
|---|---|---|---|
| Inventário 2018 | 25 (19 com valor, 6 com nome de sócio no lugar do valor) | R$ 49.684.243,20 | igual ao total do documento e ao resumo de 2018 (doc 27) |
| Inventário 2025 antigo | 22 | R$ 61.790.460,00 | confere |
| Inventário 2025 atual | 22 | R$ 58.076.823,20 | confere, e confere com o doc 00 (R$ 9.679.470,53 por sócio = 1/6) |
| Rafael, transcrição do balanço do grupo (p. 1–5) | 22 | R$ 58.076.823,20 | igual ao inventário, item a item |
| Rafael, balanço executado (p. 16–20) | 22 | itens somam R$ 75.870.000,00; total declarado R$ 76.230.000,00 | **diverge R$ 360.000,00** |
| Lista de Rafael | 5 (itens 02, 11, 12, 18, 19) | R$ 1.880.000,00 | igual ao PDF (p. 21–22) e ao Diretrizes (p. 6–7) |

## 2. Onde grupo e Rafael divergem

Só em dois itens. Os outros 20 têm o mesmo valor dos dois lados. O valor de Rafael é o do balanço executado; nas p. 6 e 13 ele ainda repete o valor do grupo.

| Item | Grupo | Rafael | Diferença | Páginas de Rafael |
|---|---|---|---|---|
| 01 Hotel (18 terrenos, quadra 37) | R$ 49.606.823,20 | R$ 66.000.000,00 | +R$ 16.393.176,80 | 7–9 e 17 |
| 20 Lote 15/q02, Condado de Capão, com casa | R$ 1.000.000,00 | R$ 2.400.000,00 | +R$ 1.400.000,00 | 13–15 e 19 |
| Erro de soma de Rafael (sem item) | — | R$ 360.000,00 | +R$ 360.000,00 | 20 |
| **Total** | **R$ 58.076.823,20** | **R$ 76.230.000,00** | **+R$ 18.153.176,80** | |

O erro de soma passa para o 1/6 de Rafael e para o crédito que ele calcula (p. 21). Com a soma correta, 1/6 seria R$ 12.645.000,00 (ele escreve R$ 12.705.000,00) e o crédito seria R$ 10.765.000,00 (ele escreve R$ 10.825.000,00).

## 3. Pontos que já aparecem

1. **Item 01.**
   - Os R$ 66 milhões são o "valor mínimo físico" de Rafael: terreno a R$ 5.000/m² mais 10.500 m² construídos a R$ 3.000/m² (p. 7). Pela conta, só terreno e prédio. Os móveis e equipamentos (R$ 10.504.000,00 no grupo) e a vantagem de 15 % (R$ 5.570.455,20) não aparecem na conta dele.
   - Faixas de Rafael: mercado de R$ 55 a 85 milhões; físico de R$ 66 a 102,45 milhões; renda de R$ 41 a 50,6 milhões.
   - O grupo mudou os parâmetros desde 2018: área construída de 12.000 para 10.500 m²; CUB de ago/2017 para dez/2022 (dois anos antes da data-base); terreno de R$ 550,72 para R$ 1.000/m²; móveis de R$ 50 mil e 35 mil para R$ 42 mil e 30 mil por apartamento; vantagem de 30 % (2018) para 25 % (2025 antigo) e 15 % (atual).
   - O próprio grupo estima, sem descontar do total, cerca de R$ 2.000.000,00 de redução no item 01 para regularizar o prédio e cerca de R$ 1.601.357,90 de INSS da obra. Não há projeto, habite-se nem profissional contratado (arquivo 1, p. 11).
   - Pelo C0, o hotel fica na empresa e pende para Rafael. A comparação parcela a parcela vai no pacote de exceções.
2. **Lotes de 2018 em nome de sócio.**
   - Pela inscrição imobiliária, o lote "IRINEU SIVIERO" de 2018 (360.0038.015.000) é o item 17 de 2025: lote 15/q38, R$ 390.000,00, descrito como "Adquirido do Senhor Hugo Ferri".
   - Os lotes de Válter (q19), Raul e Élcio (q57, com casa) e de Adélia e Pedro (q74) não aparecem em 2025.
   - O lote de Élcio traz em 2018 a nota "trocou este lote/casa por um lote no Loteamento Aditivo Medianeira" (SEI). Não achei essa troca no arquivo 17.
3. **Bens que Rafael diz faltar no inventário** (p. 24–25): duas residências na esquina da Rua da Divisa com a Rua Três, de R$ 2,4 a 3,5 milhões o par, e **uma Fiat Toro de serviço do hotel**. Ele não soma nenhum dos dois ao total. Hipótese a confirmar: as casas seriam os lotes de Raul e Élcio de 2018, vizinhos e com casa, o de Élcio na esquina. Se for isso, a de Élcio teria sido trocada.
4. **Identificação dos lotes.**
   - Os itens 04 a 10 dizem "Quadra 37", mas a inscrição é da quadra 0071.
   - Os lotes 27 e 28 se repetem nas duas quadras. A planilha-mestre vai casar os anos só por inscrição e matrícula.
   - Os lotes 27 e 28/q38 têm a mesma matrícula, 73.482.
   - O item 22 está registrado em Osório; os demais, em Capão da Canoa.
   - Os lotes 13 e 14 da quadra do hotel não aparecem em nenhum item.
5. **Planilhas Google.**
   - O lote 28/q38 aparece com 0 em 2018; no inventário de 2018 vale R$ 200.000,00. O erro passa para o total de 2018 (R$ 49.484.243,20 em vez de R$ 49.684.243,20), para o "2018 atualizado" (falta o mesmo lote a R$ 350.000) e para a "Diferença". A diferença aparece como R$ 1.500.000,00; o acréscimo real de lotes é R$ 1.150.000,00: o lote de Irineu e os dois lotes laterais novos.
   - Isso está na GERAL (linhas 70, 80 e 99), nas abas Hotel Araçá das duas planilhas e num bloco repetido do hotel dentro da aba 7.Rádio.
   - A coluna "Estudo Rafael" traz só os 5 lotes escolhidos (R$ 1.880.000,00), não a avaliação dele.
6. **Placas solares do hotel pagas pela SEI.** Um lote do Aditivo Medianeira foi entregue à Solar Serra em 24/04/2024 (arquivo 17). O inventário da SEI traz, além disso, um compromisso de mais um lote de R$ 150.000,00 com a mesma empresa. Vai para os candidatos à Parte 6.
7. **Poço.** Tem outorga (Portaria DRH 1398/2012) para incêndio, jardinagem e piscina, em caráter precário. O grupo diz que não é usado para abastecimento. Não tem valor próprio: está dentro dos "equipamentos diversos" (R$ 6.400.000,00).
8. **Data-base e Rovílio.** Nada no inventário 06 nem no PDF de Rafael desta empresa muda por "outubro/2025" ou pela venda das quotas de Rovílio.

## 4. Próximo passo

Planilha-mestre `planilha-mestre-06.xlsx`, gerada por script, com:
- os pares 2018 × 2025 por inscrição e matrícula;
- o item 01 decomposto em componentes, lado a lado com a conta de Rafael;
- a marcação da lista de Rafael (C0).
