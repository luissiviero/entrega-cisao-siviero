# Diário: cobranças do 11.2 do doc 11 (quitação e duplicidade)

Pasta própria desta sessão (P13): `02_trabalho/analises/cobrancas-11.2/`. Não altera arquivos de `parte4-fluxo-caixa/` nem da raiz de `02_trabalho`. A fusão com o log da SEI fica para depois da aprovação do Luís.

## 2026-10-28 — lista de todas as cobranças, com quitação
- Pedido do Luís: listar todas as cobranças, dizer se já foram quitadas e como, e, quando a informação faltar, perguntar ao grupo para descartar outras duplicidades. O ponto de partida foi o relatório de duplicidades da sessão "Fila de tarefas SEI" (D1 a D6).
- Escopo: as 57 linhas do 11.2 (11.2.1: 9 multas; 11.2.2: 26 lotes; 11.2.3: 1 obra; 11.2.4: 16 itens, 6 sem valor; 11.2.5: 5 itens). Total R$ 51.922.233,03, que confere com GERAL E664. As observações citam "itens 1 a 24", mas o doc 11 só tem os itens 1 a 21; a soma dos 21 bate com os R$ 16.127.652,00 da tabela 50.
- Fontes cruzadas: arq. 17 (390 registros); contas pagas de jul/2023 a out/2024 (aba Consolidado, só SEI); arq. 1 (respostas ao Dr. Pablo, p. 18-30); arq. 16; memorando Buriti; achados S1-S13 do ACERTOS; revisão do cruzamento D1-D7; diff do doc 11 antigo × atual.
- Situação (57 linhas):
  - duplicidade já abatida: 4 linhas, R$ 1.887.000;
  - quitada: 1, R$ 150.000;
  - em pagamento comprovado: 4, R$ 2.047.652,03;
  - paga em parte: 3, R$ 5.900.000;
  - contingência sem pagamento: 32, R$ 37.307.581;
  - em aberto sem comprovação: 7, R$ 4.630.000;
  - sem valor: 6.
- Fatos novos:
  - O "Contrato Buriti" de 01_fontes/outros é o memorando de entendimentos, sem assinatura e com anotações à mão, datado de 14/04/2022. É depois do Termo A de Susin (13/07/2020). A comissão Buriti não está quitada pelo Termo A, e a cl. 8 exige documento escrito. Muda a base da decisão do Luís (Susin abatido como duplicidade): o abatimento continua defensável, mas por falta de documento, não por duplicidade. Pergunta P09.
  - O memorando Buriti, cl. 12, prevê um empréstimo de R$ 1.000.000 da Buriti à SEI. Se foi feito, é um passivo que não está no 11.2.5. Favorece o sócio que sai; deve ser informado ao Luís.
  - Dr. Letti: o Luís o tem como finalizado, mas o doc 11 cobra multa de R$ 1,65 mi por não execução.
  - Sol Nascente III: a Sustenseg recebeu o lote SN II 11/C (25/04/2024) pelo licenciamento, aprovação e registro; o doc 11 diz "sem projeto".
  - Pasqual: Darcy Pasqual recebeu o lote Altos da Julieta 03/1730 (19/11/2024) em permuta por um lote do Industrial Pasqual.
  - Vale das Colinas e Vista da Colina: pagamentos e lotes já dados aos donos e aos compradores "pelo atraso".
  - Lago Negro: R$ 580.000 em dinheiro a Zucoloto (set/2025) e processo encerrado em jun/2026, com gravame perpétuo.
  - Empréstimos Piccoli: as parcelas de R$ 100 mil (Inês, desde ago/2023) e de R$ 150 mil (Samir, desde fev/2024) são posteriores ao "desde dez/2021" do doc 11. Os juros estão em dia.
  - Escavadeira: preço de R$ 887 mil, 6 parcelas pagas até out/2024. Na data-base o saldo era R$ 421.396,40, maior que os R$ 347.652 do doc 11 (diferença contra o sócio que sai, não questionada).
  - Trator D6N: nenhuma parcela até out/2024; pode ter sido comprado depois da data-base.
- 17 perguntas ao grupo (P01-P17), por valor, com a coluna "Já perguntado?".
- Correções feitas antes da revisão: o filtro da Bett pegava só 2 de 7 lançamentos (descrições truncadas em "FALAVIGN"), corrigido para R$ 83.700; o filtro do trator pegava um D51 de outra máquina, corrigido para D6N; nos contratos de Zanrosso (R$ 150 mil) e Lori Fioreze (R$ 170 mil), o texto passou a separar o valor do contrato do valor pago no período.
- MD5: cobrancas-11.2-quitacao.xlsx 0de05e97d79f6d3e929e03eb08ead50e · _scripts/build_cobrancas.py a69286d814d3193ebbbfe0150d8220a5 · Contas_Pagas_2023-07_a_2024-10.xlsx a29fc2f2ebd108ae3e262233ad1d54a5 · Comparativo ... 2018 x 2025.xlsx c499486a2622eb81325b2201034adc75.
- Revisor independente disparado: `_diario/revisao-cobrancas.md`.

## 2026-10-28 — revisão independente aplicada
- Revisor: `_diario/revisao-cobrancas.md`. Encontrou 13 erros e uma lista do que favorece o grupo. Conferi nas fontes o trator ("adquirido em setembro de 2025", doc 11, máquinas item 14), o arq. 10 de Rovílio (cl. 3ª, itens 1-10), os lotes do Falavigna no arq. 17 (15/6371 Bett 2022, 18/6373 AM Cimento 2022, 028/322 Florense 06/2026), a topografia de áreas (Stuani R$ 4.500 + área Nestor R$ 15.000) e os pagamentos a Dorval Boff (contrato novo de 20 parcelas em jul/2024).
- Correções:
  - E1: trator passa a "Posterior à data-base".
  - E2: Rovílio passa a "Em aberto, documentada", pelo arq. 10. Deixa de haver inconsistência: o 07/04 é comissão, os lotes da N. S. da Saúde são de 2012/2013, as datas foram corrigidas e o lote de R$ 550 mil é comissão.
  - E3: Susin, os 6 lotes do Termo B passam a ser dívida documentada.
  - E4: Falavigna, duplicidade só até os 5 lotes de 2025; fatos de 2022 e de 2026 incluídos.
  - E5: áreas rurais, R$ 19.500.
  - E6: Fioreze, R$ 55.000, tratado como compra de área.
  - E7: Solar Serra não é duplicidade.
  - E8: Letti, Sol Nascente III, Pasqual, Bonetto e Cecconello passam a "multa contestável".
  - E9: valor ligado a cada pergunta agora calculado (P01 = R$ 13.473.000, 16 lotes).
  - E10: coluna "Já perguntado?" corrigida.
  - E11: contradição do Boff incluída na P05.
  - E12: 2 linhas sem valor (itens 22-24 e parágrafo final).
  - E13: P17, empréstimo Buriti.
  - Novas: P18 (itens sem valor), aba "A favor do grupo" (12 fatos), coluna "Já retirado pelo Luís".
- Resultado: 59 linhas, total R$ 51.922.233,03 (confere). Já retirado pelo Luís: R$ 21.047.247,67 (11.2.1 R$ 15.754.581 + R$ 5.292.666,67 de duplicidades). Situações: duplicidade abatida 4; quitada 1; paga em parte 3; em aberto documentada 5; em aberto sem comprovação 5; contingência 32; posterior à data-base 1; sem valor 8.
- Ponto para o Luís: a decisão sobre Susin (abatido como duplicidade) continua defensável, mas a base muda: falta de documento (cl. 8), não duplicidade. E os 6 lotes do Termo B são dívida real que pode estar fora do passivo.
- Não houve segunda revisão depois das correções.
- MD5: cobrancas-11.2-quitacao.xlsx 1ffa7a928e67109bef217b2fd9a3e944 · _scripts/build_cobrancas.py 80c143d4b24fd2dbb67eefcbc645aa63.
