# Diário — Parte 4 da SEI pelo fluxo de caixa (sessão própria)

Pasta própria: `02_trabalho/analises/parte4-fluxo-caixa/` (regra P13: esta sessão não escreve na raiz de `02_trabalho`, em `empresas/` nem em `03_entrega/`). Só acréscimos.

## 2026-10-28 — abertura e método do Luís
- O Luís pediu para acelerar o tópico "CONTRATOS | INDENIZAÇÕES | COMPROMISSOS | EMPRÉSTIMOS" (Parte 4), que estava em 4º lugar na fila da SEI, e explicou a abordagem (pré-requisito do plano, seção 4):
  - a Siviero avalia os loteamentos não finalizados em R$ 0 e ainda desconta multas rescisórias; Rafael mistura lotes executados com multas (contraditório);
  - em vez de avaliar cada loteamento, seguir o fluxo de caixa: lotes que saíram de 2018 a 2025, outras entradas (terras, soja), valores quitados (dívidas, multas, prestadores, juros) e custo recorrente dos 15 meses de relatórios estendido ao período;
  - combustível não é recorrente (indica máquinas em obra); dívidas, multas e afins quitados entram junto com os custos de manter;
  - parte do investimento pode ter ido para concluir loteamentos inacabados em 2018 e prontos em 2025;
  - objetivos: valor do investimento do período; queda do patrimônio sem explicação → (a) valores reais ou (b) devolução de 1/6; gráfico patrimônio × tempo; outras entradas em separado antes do total do tópico; comparação com Siviero e Rafael;
  - multa rescisória: não paga com contrato vigente; aceita só com distrato e termo de nunca mais lotear a gleba (novo contrato após a cisão = estelionato, nas palavras dele); com distrato, investimentos na gleba não são cobrados, mas precisam ser demonstrados.
  - Pediu planilhas e um PowerPoint simplificado, versão inicial com o que já existe. Exceção autorizada à P8 (apresentação antes da aprovação técnica).

## 2026-10-28 — entradas (MD5)
- Planilha Google (export de 28/10 11:18; Drive modificado pela última vez em 27/10 22:56 UTC, logo o export está atual): `01_fontes/planilhas-google/Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx` c499486a2622eb81325b2201034adc75.
- Contas pagas classificadas, **em edição** (lock `~$` às 17:21; sessão "Verificação e análise de relatórios em Excel" ativa): cópia em `_entradas/contas-pagas-classificadas_snapshot-2026-10-28.xlsx` 668473c991eadc76df54177ada68b5e6.
- `01_fontes/outros/Contas_Pagas_e_Lotes_Vendidos_2023-07_a_2024-10.xlsx` 80e181d0e1d2999fce4cffea6df13f71; `Contas_Pagas_2023-07_a_2024-10.xlsx` a29fc2f2ebd108ae3e262233ad1d54a5.
- `02_trabalho/sei-2018-itens.xlsx` dcabdf2209e918d39e1ef81f547c779e; `sei-2025-itens.xlsx` 62ed740e2dca3c7ee6e041d4936ce330; `planilha-mestre-sei.xlsx` a4c54c1c373e19ce8ae576959309449c; arquivo 17 xlsx 7ef0dc35dcc123bd1d356b0e1fa45164.
- Lidos só como contexto (aguardam aprovação, não usados em números): `analises/acertos-e-areas-vendidas/achados-para-sei.md` e `extracoes.xlsx`.
- OCR das respostas ao Dr. Pablo terminado pela sessão SEI (`02_trabalho/ocr-respostas-docs/`); locação Oleoplan R$ 14.000/mês no 1º ano segundo o OCR (arq. 15, a conferir na imagem).

## 2026-10-28 — v0 gerada
- Script `_scripts/build_fluxo_caixa.py` → `fluxo-caixa-sei-v0.xlsx` + `numeros-v0.json`.
- Premissas v0: período 2018–2024 (84 meses; o inventário de 2018 ainda conta lotes vendidos em 2018); saídas de 2025 fora da ponte (P2); redução do inventário manda na quantidade (P3), a lista do arq. 17 dá ano e forma; preço piso = valor 2018, teto = média out/2025 do grupo, central = interpolação pelo ano; custo de manter = média dos 15 meses (Recorrente sem combustível + multas, acordos, lucros etc.) × 84; dívida = saldo jan/2019 − empréstimos 2025; bens novos = itens GERAL com 2018 = 0; concessão Falavigna + Colina do Retiro II = 150 lotes × R$ 50.000 (custo/lote do próprio grupo no 11.2.1).
- Conferências do script: lotes 2018 = 761 e 2025 = 426 batem com a linha Total da aba; SEI 15 meses = soma da aba Relatórios Mensais (R$ 9.877.224); totais 11.2 Siviero (−51.922.233,03) e Rafael (−17.778.633) batem com GERAL E664/F664.
- Erros encontrados e corrigidos antes do revisor: (1) acentos decompostos no texto do arq. 17 → nenhuma "venda à vista" reconhecida; classificação passou a usar texto em ASCII; (2) nome do destinatário disparava classes ("Casagrande" → casa) → só usado para sócios/família; (3) filtro de bens novos excluía a Hilux (item 11.0 lido como seção) e incluía o trator comprado em set/2025 (pós data-base).
- Resultado v0 (84 meses, com concessão): sem explicação piso −17,3 mi; central 0,4 mi; teto 26,7 mi (1/6 = 4,4 mi). Investimento a demonstrar (teto, com lotes dados em obras) 33,1 mi. 72 meses sem concessão, teto: 37,7 mi. 46 lotes listados além da redução: até 24,3 mi (fora da base). 30 lotes saídos em 2025: 11,9 mi no teto (candidatos ao registro de data-base).
- Amostra de valores que o próprio grupo atribuiu a lotes isolados em dações e aquisições (n = 12): mediana 102% da média out/2025; 6 acima. Base para usar o teto como referência (premissa a aprovar).
- Deck: `_scripts/build_deck_parte4.py` → `parte4-sei-fluxo-caixa-v0.pptx` (12 slides, gráficos nativos, estilo da v1). A renderização pelo PowerPoint (`_scripts/render.ps1`, cópia do script da sessão de apresentação) foi **negada pelo classificador de permissões**; conferência de conteúdo feita pelo texto extraído; conferência visual pendente.
- Revisor independente (P11) disparado sobre script, planilha e JSON, com as fontes, sem as conclusões. Saída: `_diario/revisao-v0.md`.

## 2026-10-28 — revisão da v0 e v0.1
- Revisor: `_diario/revisao-v0.md`, 8 altos, 8 médios e 8 baixos. Conferidos por mim nas fontes: Perini 2018 = 4 × 500.000 + 6 × 95.000 (inventário seq. 66); Lago Negro 35 lotes com preço (seq. 77); Falavigna 2018 "custo para concluir" 2.450.000 (seq. 111) e GERAL E737 −2.000.000 em 2025; K88 "Mesma área do item 2?" (Sapopema); registros sem data de Colina do Retiro II marcados "Não consta no inventário 2025 atualizado"; pró-labore ao Hotel Araçá (R$ 186 mil com impostos); PIS/COFINS na razão do regime cumulativo.
- Consequência: a mediana de 102% da amostra de dações (v0) não vale como prova de preço (M7). O argumento do teto passou a ser o teste de receita (A8): PIS/COFINS dos 15 meses → R$ 11,0 mi de receita; comissões a 6% → R$ 6,8 mi (v0.1) / 6,2 mi (v0.2, após a reclassificação de 18:53).
- v0.1 (script congelado em `_scripts/build_fluxo_caixa_v0_1.py`): teto 25,8 mi sem explicação. Segunda revisão `_diario/revisao-v0.1.md`: N1 (pré-2018 consumindo redução primeiro — erro meu ao aplicar A1), N2 (regra THIN_2025 criada por mim, contraria P5), N3 (permuta lote por lote contada como dinheiro), N4 (bens recebidos que não estão em 2025 contados como explicados), e baixos (Flávia por parcelas, móveis, Carraro, amostra de preços).

## 2026-10-28 18:53 — versão final das contas pagas e v0.2
- A sessão "Verificação e análise de relatórios em Excel" fechou `Contas_Pagas_Recorrentes_x_Extraordinarias_..._revisado.xlsx` (18:53, MD5 84a615ce565cc4303112113e961e50ac); cópia em `_entradas/contas-pagas-classificadas_snapshot-2026-10-28_1853.xlsx`. Contém decisões do Luís de 28/10/2026 (abas Critérios, Dúvidas, Extrapolação): obras em loteamento finalizado, loteamento rescindido e despesas de familiares e de outras empresas = recorrentes; só recorrentes são extrapolados, jan/2019–dez/2024, cada mês = média × IPCA(mês)/IPCA(mar/2024); aba IPCA (IBGE SIDRA 1737).
- v0.2 adota esse método: período 2019–2024 (72 meses); extraordinários de manter contados uma vez; aluguéis 2018 × IPCA; teto = valor 2025 do grupo × IPCA(mês da venda)/IPCA(dez/2024) (a tabela termina em dez/2024: leve viés a favor do sócio; teto nominal na sensibilidade). Combustível mantido como obra pela instrução direta do Luís nesta sessão (a planilha deixa R$ 75 mil como recorrente).
- Correções N1–N4 e baixas aplicadas. Classe nova BEMX (troca por bem que não aparece em 2025) entra no valor sem explicação.
- Resultado v0.2: sem explicação piso −1,5 mi; central 16,8 mi; teto 28,5 mi (1/6 = 4,74 mi); investimento a demonstrar (teto) 40,9 mi. Sensibilidade (teto): nominal 38,2 mi; início 2018 26,8 mi; concessão máxima 21,4 mi; 26 lotes listados além da redução: até 12,5 mi fora da base; 34 lotes saídos em 2025 (registro de data-base).
- Saídas (MD5): build_fluxo_caixa.py 7b5a96181381b57107a2872c3aa6f2d6 · fluxo-caixa-sei-v0.2.xlsx 3a575fa04e8f9406a25eff1930c80ebf · numeros-v0.2.json 25970ce5ee672166ee6d764da3405d2c · build_deck_parte4.py (antes do ajuste de títulos) 323321b14b2c4d4df5255a6978b673aa.
- Deck v0.2: 12 slides; títulos encurtados por checagem aproximada de largura (renderização pelo PowerPoint continua negada). Terceira revisão (só das mudanças da v0.2) disparada: `_diario/revisao-v0.2.md`.

## 2026-10-28 — terceira revisão (v0.2) e ajustes finais
- Revisor: `_diario/revisao-v0.2.md`. Conferiu somas do IPCA (63,5962 base mar/2024; 85,6469 base dez/2018), C1, C1b, aluguéis, Flávia (430 mil), C2, deflator, alocação N1 e Ponte nos 4 cenários. Apontou V1 (travessão apagado pelo norm na regra de troca lote por lote), V2 (3.100 m² do Sapopema em BEMX), V3 (paisagismo e Attolini sem separar a parte em dinheiro, ~0,3 mi, não corrigido: registrado), V4 (obra do Falavigna nos recorrentes + concessão), V5 (deflator do teto é proposta minha, não decisão do Luís), V6/V7 (rótulos).
- Corrigidos: V1 (`imovel\W*lote` + conferência do lote recebido contra o estoque de 2025 do doc 11: Monte Vêneto 09 e 10/01 → BEM; 05/05 e 09/06 → BEMX), V2 (regra da área de 3.100 m²), V4 (concessão = 2.450.000 − 2.000.000 − obra do Falavigna já nos recorrentes = 370.716,70), V6 (rótulos). V5 levado ao Luís como premissa (aba Pendências); teto nominal na sensibilidade.
- Resultado final v0.2 (sem explicação / 1/6 / investimento a demonstrar): piso −1,66 / — / 4,33 mi; central 16,52 / 2,75 / 26,36 mi; teto (IPCA, referência) 28,19 / 4,70 / 40,63 mi; teto nominal 37,92 / 6,32 / 53,16 mi.
- MD5 finais: fluxo-caixa-sei-v0.2.xlsx f1f91761b14a13cba032010b8dcf674c · parte4-sei-fluxo-caixa-v0.2.pptx 0ead91a132c464e74dbb4370a7bab946 · numeros-v0.2.json c6a51d79d36bb985e5f1e9c40e28ccbe · build_fluxo_caixa.py eccfb7e709181c69f702afc16f89d1e3 · build_deck_parte4.py 323321b14b2c4d4df5255a6978b673aa.
- Pendente: conferência visual do deck (o Luís vai liberar a regra de permissão do render.ps1 desta pasta); aprovação das premissas; avisar as sessões SEI e de apresentação.

## 2026-10-28 — conferência visual do deck v0.2
- O Luís liberou a regra de permissão do `_scripts/render.ps1` desta pasta. Renderizados os 12 slides pelo PowerPoint (`_png/parte4-sei-fluxo-caixa-v0.2/`) e conferidos um a um.
- Defeitos corrigidos: slide 7 (valores das barras pequenas da cascata cortados → valores passaram para os rótulos das categorias), slide 8 (coluna 1/6 da tabela quebrava "R$ 4,7 mi"), slide 4 ("Início de 2018" → "Fim de 2018 (inventário)"), slide 12 (três rodadas de revisão). Rerenderizado e conferido.
- Deck final: parte4-sei-fluxo-caixa-v0.2.pptx MD5 a538387de45bf4622cef48d270f10965.

## 2026-10-28 — aprovação do preço de referência
- Luís: "Aprovar com IPCA". O preço de referência passa a ser a média do grupo em 2025, trazida pelo IPCA até o mês da venda (cenário teto). Resultado: R$ 28.186.245 sem explicação; 1/6 = R$ 4.697.708.
- Registrado em `log-decisoes.md` desta pasta, junto com as decisões de método da sessão. A sessão SEI copia para `02_trabalho/log-decisoes.md` e `decisoes-sei.json` na fusão (P13).
- Planilha (aba Pendências) e deck (slide 12) atualizados. Deck renderizado de novo e slide 12 conferido.
- MD5: fluxo-caixa-sei-v0.2.xlsx 1123cdd4e41a4198361bab074a10fe75 · parte4-sei-fluxo-caixa-v0.2.pptx dee071f95ce8cc6771eef6f29867ab3a.

## 2026-10-28 — lotes do arquivo 17 incluídos; pró-labore por ramo
- Luís: "os 26 lotes entram na conta". Implementado com `INCLUDE_FILE17_EXTRA = True`: cada registro do arquivo 17 conta como um lote; o estoque de 2018 sobe na mesma quantidade (761 → 802, contando também registros de 2025 e sem data). Lotes no período: 310; em 2025: 37.
- Pró-labore da SEI nos 15 meses (contas pagas, versão 18:53): R$ 186.000 para cada um de 4 recebedores, Válter (depósito para Vagner), Adélia (em mãos, Luciane), Pedro (em mãos, Pedro Elias) e HOTEL ARAÇÁ LTDA, este com a observação "Raul Siviero + pagto contas individuais". Total R$ 744.000. São 4 dos 6 ramos, e não 3 como estava escrito: corrigido na planilha, no deck e no log.
- Resultado (sem explicação / 1/6 / investimento a demonstrar): piso 2,01 / 0,33 / 9,04 mi; central 23,23 / 3,87 / 34,73 mi; teto (referência) 37,59 / 6,27 / 52,21 mi; teto nominal 49,45 / 8,24 / 67,44 mi.
- Deck renderizado de novo; slides 2, 4, 7, 8, 10, 11 e 12 conferidos.
- MD5: fluxo-caixa-sei-v0.2.xlsx 4be83d994a4c30a499ebcfd891526848 · parte4-sei-fluxo-caixa-v0.2.pptx 3d0e202244c2a5fb889469fa790bda8a · numeros-v0.2.json 2c9b9440e417c4ccd6ff09c05918fb0b · build_fluxo_caixa.py 91f6cb55488ea1ec5e0c6e7498f44f01 · build_deck_parte4.py 96ee21eed613868ab9a787f5350b6331.

## 2026-10-28 — pró-labore como custo da empresa
- Luís: "É custo da empresa". O pró-labore dos 4 ramos (Válter, Adélia, Pedro e Raul via Hotel Araçá) continua no custo de manter. Os números não mudam: sem explicação R$ 37.592.641, 1/6 = R$ 6.265.440.
- Registrado em log-decisoes.md. Pendências da planilha e slide 12 do deck atualizados; slide 12 renderizado e conferido. Premissas de método: nenhuma pendente. Seguem os itens 11.2.2 a 11.2.5, um a um.
- MD5: fluxo-caixa-sei-v0.2.xlsx 6f4e0c5140b57aa163f1b800250ce4a5 · parte4-sei-fluxo-caixa-v0.2.pptx 3e8e791f152b23dd8dff92db3196d91d · 

## 2026-10-28 — passivos do 11.2 já pagos com lotes (pergunta do Luís)
- Pergunta: os itens do 11.2.2 ao 11.2.5 aparecem na lista de destinos do arquivo 17? Se aparecem, o lote sai do inventário e a dívida continua cobrada (forma 1 da Parte 6).
- Script `_scripts/cruzar_compromissos_arq17.py` → `cruzamento-compromissos-arq17.xlsx`: 57 cruzamentos por palavra-chave (aba Cruzamento) e 6 casos lidos um a um (aba Avaliação).
- Revisor independente: `_diario/revisao-cruzamento-compromissos.md`. Correções aplicadas: Susin não confirmado (o Termo A só quita intermediações até 13/07/2020; cl. 8); Falavigna e Colina II só possíveis (dependem da composição dos saldos); 11.2.2 sobre lotes já de terceiros não é cobrança dupla (contingência); confirmados só os 4 lotes que saíram em 2025 (incluído Primavera 10/1199, arquivo 1 p. 26), R$ 1.887.000; Solar Serra confirmado (estoque e passivo se compensam; dívida do Hotel Araçá); Rovílio é inconsistência (lote 07/04 em dois pagamentos; N. S. da Saúde 12/237 e 15/252 fora do arquivo 17).
- Outros achados: lotes já entregues que seguem no estoque de 2025 (Medianeira 02/07, Alto das Graças 11/6371 — também na lista de Rafael —, Sol Nascente I 1/A e 2/A); passivo omitido da dação G3 (016/876).
- Ponte e deck não mudam nesta etapa: os casos afetam o 11.2 (passivos), que está em "valor do grupo, decisão item a item".
- MD5: cruzamento-compromissos-arq17.xlsx e4b6b243ffb6ef18e94c253d0f3c930b · _scripts/cruzar_compromissos_arq17.py b143072feb1bb77e3bcacf30c8e1c6c4 · 

## 2026-10-28 — decisões sobre as cobranças duplas
- Luís: os três casos que dependiam de informação do grupo (Falavigna, Colina do Retiro II, Susin) são duplicidade. Regra aplicada: sai do passivo o que os lotes pagaram, até o saldo cobrado. Falavigna, lotes de R$ 3,04 mi contra o saldo de R$ 2,0 mi: sai R$ 2,0 mi. Colina II: sai o lote, R$ 205.667. Susin: saem os R$ 1,2 mi da comissão. Mais o caso confirmado D1 (R$ 1.887.000). Total abatido R$ 5.292.667. Proposta do tópico 11.2: R$ 6,7 mi (antes R$ 1,4 mi negativo). A ponte não muda.
- Interpretação minha a confirmar: no Colina II, só o valor do lote entregue sai dos R$ 3,8 mi, não o saldo inteiro.
- Luís: lotes entregues que seguem no estoque e a dação G3 omitida estão corretos (desconto feito uma vez). Sem ajuste.
- Luís: os lotes do Morada dos Ipês II prometidos a Rovílio provam que loteamento não finalizado tem valor e que o inventário não mostra todos os loteamentos. Tarefa para depois registrada em log-decisoes.md (Ainda pendentes).
- Script build_fluxo_caixa.py passa a ler os casos de cruzamento-compromissos-arq17.xlsx (aba Avaliação). Ordem de execução: cruzar_compromissos_arq17.py antes de build_fluxo_caixa.py. Slide 10 renderizado e conferido.
- MD5: fluxo-caixa-sei-v0.2.xlsx 218ef994b596d9e1518abc27c8609340 · parte4-sei-fluxo-caixa-v0.2.pptx 8a5b998c704d4febc2c1a06219275eac · cruzamento-compromissos-arq17.xlsx 4679c10f9c673d7193d64ee6f6e1dc02 · 

## 2026-10-28 — regra geral: só cobranças futuras no 11.2
- Luís: todas as cobranças passadas são zeradas (lotes, valores e demais), porque o bem já saiu do inventário ou o pagamento já saiu da conta. Só cobranças futuras ficam na lista.
- Revisão item a item pelos textos do doc 11 (GERAL B720–B743): nova cobrança passada encontrada, a Solar Serra (R$ 150.000, lote entregue em abr/2024). Zerada. As demais são obrigações futuras ou contingências (obras a fazer, processos e multas em aberto, lotes a entregar a Celestino Rossi e Susin, parcelas a vencer). Colina do Retiro II: confirmado que sai só o valor do lote; o restante é obra futura.
- Total abatido R$ 5.442.667; total do tópico na proposta R$ 6,9 mi. A ponte não muda (R$ 37.592.641 sem explicação). Slide 10 renderizado e conferido.
- Pendência para a sessão SEI: o lote 02/07 do Medianeira, já entregue, precisa sair do estoque de 2025.
- MD5: fluxo-caixa-sei-v0.2.xlsx 452757aa9c4eb949be6dbb922aa5f378 · parte4-sei-fluxo-caixa-v0.2.pptx 0e0262e1373e6dc96ed53af1218bfb55 · 

## 2026-10-28 — deck v0.3 (18 slides)
- Pedido do Luís: um slide explicando a cobrança em duplicidade e um slide por sessão (S1, S2, S3, S4 e o restante da SEI), com grupo Siviero, Rafael e proposta.
- O v0.2 estava aberto no PowerPoint do Luís (arquivo travado); a nova versão foi salva como parte4-sei-fluxo-caixa-v0.3.pptx. O v0.2 fica intacto.
- Slide 11: duplicidade, com o exemplo do Falavigna (lotes de R$ 3,0 mi entregues; R$ 2,0 mi ainda cobrados) e a tabela das 5 cobranças passadas retiradas (R$ 5,44 mi; 1/6 = R$ 0,9 mi).
- Slides 13 a 17: números lidos das planilhas-mestre de cada sessão (só leitura; `_scripts/deck_extra_slides.py`). S1: 75,4 / 79,5 / 79,5 (C0 de referência, só 4 de 24 itens decididos). S2: 58,1 / 75,9 / 75,9. S3: 26,0 / 29,3 / 29,3. S4: 3,0 / 3,3 / 3,3. SEI bens: 152,8 / 155,3 / 155,2 (R$ mi; grupo / Rafael / proposta).
- Os 18 slides foram renderizados; conferidos os novos e o título do S4, que foi encurtado.
- MD5: parte4-sei-fluxo-caixa-v0.3.pptx 56e362b44b6337fbbe45b406a200d519 · _scripts/build_deck_parte4.py e57300615330d25cf9bf31170c723571 · _scripts/deck_extra_slides.py dbfbed50289849ee5ccec7bfd93d2289 · 
