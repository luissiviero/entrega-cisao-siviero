# Revisão independente — planilha-mestre 02/05 (28/10/2026)

Revisor de contexto novo. Fontes lidas direto: inventários 2025 (docs 02, 05, 08, 11, docx "áreas vendidas", antigos 05, 08, 11 e demais antigos/atuais por busca de matrícula), inventários 2018 no Drive (docs 13 e 17), xlsx SEI 2018 (aba "Bens em outras empresas"), PDFs de Rafael (Hotel, Agropecuária, Transportadora, Diretrizes, empreendimentos), lista `ativos-avaliados-valter.xlsx`, planilha 2_0 (todas as abas), respostas ao Dr. Pablo arq. 1 (texto). Rodei `build_mestre_02_05.py` (regenerou os dois xlsx; saída: 24 linhas, 8 + 13 linhas 2018, totais 75.410.201,00 × 79.491.957,50). Nenhum outro arquivo foi alterado.

## Achados

### Alta

**A1. "Existe no Estudo 2 (Rafael)? = Não" está errado em 5 linhas 2018 (lição 1 do plano).**
O script grava "Não" fixo para toda linha 2018 sem par (`build_mestre_02_05.py`, linha 384). Rafael avalia esses bens nos documentos das empresas onde estão em 2025:
- 02-2018-s10 (lote 12, mat. 3.062), s11 (lote 89, mat. 8.486), s12 (lote 33, mat. 1.218): `estudo-rafael/empreendimentos-pdf/ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL, BELISE LTDA.pdf`, itens 04, 08 e 09 (p. 2–3, 10, 17–19, 23–24; ex.: lote 33 "estimado entre R$ 180.000,00 e R$ 350.000,00", p. 19; o xlsx correspondente tem abas Avaliacao e Quadro_Resumo).
- 02-2018-s14 (69,8 ha, mat. 459/460/461/558) e s15 (4,5 ha, reg. 15.922): `TRANSPORTADORA IRMÃOS SIVIERO LTDA, BELISE LTDA.pdf`, p. 3, 7 (concordância expressa) e 14–15.
Correção: preencher Sim + documento e página; trazer o valor de Rafael dessas seções como referência (a decisão é da sessão dona do bem).

**A2. Item 05 (Fazenda, R$ 44.828.300,00, 59 % do Hotel, na lista de Rafael) aparece como "Concorda" sem ressalva.**
No PDF do Hotel não há "DA CONCORDÂNCIA" logo após o item 05 (só a referência genérica no item 06, p. 14). Rafael registra que os valores "deverão ser considerados parâmetros preliminares de referência, sujeitos à revisão" e recomenda avaliação independente (p. 13–14). Do jeito que está, o item entra na regra em bloco do Pacote 1 (etapa 3 do brief).
Correção: situação "Concorda com ressalva" (ou equivalente) e tratar o item 05 como exceção, fora do Pacote 1.

### Média

**M1. Falta candidato à Parte 6 e pergunta sobre o custo de regularização em duas empresas.**
Doc 11 (SEI), OBSERVAÇÕES: "áreas rurais que estão adquiridas por contrato, e que ... ainda não foram escrituradas ... custo aproximado de R$ 250.000,00" (planilha 2_0 GERAL B882, E882 = −250.000). Docs 05 e 02: R$ 400.000,00 de regularização da fazenda, que está "em nome ... de TERCEIROS, POR CONTRATO". A nota do Luís em GERAL H882 pergunta justamente "Fazenda São Crispin?", e a mestre só a liga ao item 05. Sem candidato P6, sem linha em "Passivos e deduções" e sem pergunta ao grupo. Resumo B190–B191 ("Necessário maior explicação sobre isso") também ficou de fora.
Correção: P6-08 (desconto em duplicidade, regularização da fazenda no doc 05 × R$ 250.000 do doc 11), linha de referência na aba Passivos e pergunta "quais áreas compõem os R$ 250.000 da SEI".

**M2. Observações do Luís omitidas na aba Resumo da planilha 2_0.**
A Fase 0 varre só as linhas 73 e 78–94 (`fase0_extracao.py`, linha 1963). Ficaram de fora:
- B99–B102, item 2.0 do bloco "FAZENDA SÃO CRISPIN": "Vendas: valor médio por hectare está maior em 2025 do que o apresentado em 2018" (+ "venda em 2018:" / "venda em 2025:"). Pertence ao 05-2025-05 / P6-06.
- B20 "05. Hotel Princesa dos Vales", na lista "VALUATIONS NECESSÁRIAS / A íntegra dos documentos" (B16–B18).
- B190–B191 (ver M1).
- As linhas de cálculo do arrendamento B88, E88 ("saca de soja ~R$ 120"), B92, E92 e B96–B97 entram só em parte: B87, F88, F91 e C94 foram pegas, as demais não.
Correção: ampliar as faixas da aba Resumo e mapear essas células.

**M3. A Base B do benchmark 2018 "das duas empresas" omite a Agropecuária da planilha 2_0.**
Resumo da mestre: "B. Base da planilha 2_0 (GERAL C28 / aba 5.Hotel C2)" = 48.730.897,83. A 2_0 tem a Agropecuária em GERAL C11 (= C13, trator, R$ 116.600) e na aba 2.São Crispin C2. Para as duas empresas, B = 48.847.497,83 e B − A = 2.388.230,00. A decomposição da aba "Base 2018" só fecha porque subtrai o trator.
Correção: somar C11 ou renomear a base como "só seção do Hotel na 2_0" e explicar a ausência do trator.

**M4. Matrícula 1.412 (tema 9): a mestre não registra a sobreposição nem o problema no argumento C1 de Rafael.**
- O grupo (arq. 1, p. 8–9) diz que a matrícula 1.412 tem 1.800 m² no total: 720 m² no item 01 e 1.080 m² no item 02. Mesmo assim, a mestre mostra "Área (2025)" = 1.800,00 m² no item 01 e 1.080 m² no item 02, o que soma 2.880 m².
- As faixas de Rafael para o item 01 (p. 8) tratam da "área comercial premium de 1.800 metros quadrados", isto é, da matrícula inteira (itens 01 + 02), a R$ 1.500–2.500/m² (R$ 2,7–4,5 mi). No mesmo PDF ele aceita o item 02 a R$ 7.500/m² (R$ 8,1 mi). Os R$ 17.555.000,00 do item 01, sem memória de cálculo, podem incluir o terreno do item 02. Isso pesa contra usar o valor de Rafael (C0 = maior) sem ressalva.
Correção: área "1.800 m² (720 m² avaliados no item; 1.080 m² = item 02)" e nota nas linhas 01 e 02.

**M5. NA-SOJA-RIGO marcado "Existe em 2018? Não" sem razão.**
As 4 × 20.150 sacas estão no doc 17 2018, T6R1 (venda Rigo). As safras de 2021 e 2022 (R$ 3.224.000,00 a R$ 80/saca) estão dentro do "TOTAL RECEBIDO" de R$ 12.058.000,00, abatido do bruto de 2018 (a própria Fase 0 marca DIVERGE). Ou seja, o crédito existia em 2018 e foi tratado como recebido.
Correção: "Sim, doc 17 2018 T6R1", nota ligando a P6-06. Na Empresa, indicar que a venda foi da Agropecuária (doc 17) e o docx de 2025 está no Hotel.

**M6. Falta a coluna "Existe em 2025 (grupo)?" da seção 0.7.**
Para as linhas 2018 a resposta varia: não está no doc 05, mas está no doc 11 ou no doc 08. Hoje isso só aparece no texto livre de "Localização em 2025".
Correção: incluir a coluna (Sim/Não + documento e tabela/linha).

**M7. Mapeamento de GERAL F943 sem apoio.**
F943 é a linha-título "VENDAS SÃO CRISPIN" e diz que a diferença de R$ 1.874.840 vem das "linhas 947 a 950", que são as 4 parcelas Pellegrini (107.800 + 220.000 + 575.520 + 971.520). A mestre liga F943 a 02-2018-s02 (Rigo, 106,53 ha) e 02-2018-s09 (Angeli). Nenhuma das duas é o objeto da nota.
Correção: ligar F943 a s03–s06 e/ou ao 05-2025-05 / P6-06.

### Baixa

- **B1.** RDB-1 diz valer para "Todos os itens do Hotel (05-2025-01 a 08)", mas a coluna "Registro data-base" só traz RDB-1 nos itens 01, 05 e 08 (dicionário `RDB` do script). Faltam 02, 03, 04, 06 e 07.
- **B2.** As notas H52 e H63 ("Não bate com 2025 versão anterior") não têm o dado na linha. A mestre não mostra o valor antigo: item 05 R$ 52.978.900,00 e item 06 R$ 520.000,00 (ANTIGO 05 T5R3/T5R4, a R$ 130.000/ha). Sugestão: coluna "Valor 2025 antigo".
- **B3.** "Fonte Rafael" cita só o balanço do grupo e o executado. Faltam a página da "avaliação executada" (p. 7 do item 01, com R$ 13.504.243,50; p. 15 do item 08, com R$ 54.000,00), dos "bens escolhidos" (p. 22–23) e do Diretrizes (p. 4–5). A contradição interna de Rafael (avaliação = valor do grupo; balanço executado = outro valor) está na aba "Grupo x Rafael 05" da Fase 0, mas não chega à mestre. Sugestão: coluna "Valor na avaliação executada".
- **B4.** Rafael imprime o item 08 duas vezes no balanço executado (p. 19). O total conta uma vez só, o que está certo. Vale uma nota na linha 05-2025-08.
- **B5.** Perguntas ao grupo:
  - Q2 deveria citar também "44,16 ha" × "= 44,1 hectares" (doc 17 2018 T3R7).
  - Q4/Q5 deveriam citar a coincidência das 2.606 sacas: Rigo vendida em 2024 = safra Lorenzet 2021 (Avisos da Fase 0; achado E1).
  - A 2ª parte de Q5 já tem resposta parcial no arq. 1, p. 6 (arrendamento Hotel × Agropecuária "pro forma", venda dos grãos tributada no CNPJ da Agropecuária): citar e reformular.
  - Falta pergunta sobre o pagamento em boi do arrendamento (dezembro de 2024, anterior à data-base) e sobre o destino da receita do arrendamento (arq. 1, p. 6).
- **B6.** P6-03 deveria registrar a evidência contrária: o lote 14 da SEI tem matrícula 4.194 e número de lote diferente (14 × 12). A ligação é só por área + vendedor (risco da lição 3). Na P6-05, o valor em jogo pode ser dado: R$ 3.750.000,00 (doc 08 T4R2–T4R3). Rafael também avalia esses bens (Transportadora, p. 3 e 7).
- **B7.** P6-07 e Q7 coincidem com o achado E4, que está no pacote 2, ainda não aprovado (`analises/acertos-e-areas-vendidas/log-decisoes.md`). Têm base própria nas fontes, mas não devem ser apresentados como achado aprovado.
- **B8.** O cabeçalho "Bloco (inventário 2025)" é preenchido com blocos de 2018 nas linhas 2018. As benfeitorias (s08) aparecem como "Fazenda São Crispin (parcelas)". Na coluna "Situação", as linhas 2018 usam categorias fora do vocabulário da 0.7 ("Vendido em 2018…") que se parecem com destino. O fato tem fonte no doc 17, mas convém deixar claro que o Destino (P4) continua vazio.
- **B9.** GERAL H14 / 2.São Crispin I5 (nota sobre a linha do arrendamento) está ligada só às linhas de soja Lorenzet. Ligar também ao 05-2025-05 / RDB-3.
- **B10.** A saída dos 83,47 ha da mat. 00.760 (Rigo) não tem nota própria na linha 05-2025-05. Só aparece na legenda e na linha s02. Acrescentar uma nota de destino (venda Rigo 2018, doc 17 T3R2 e T6R1).

## Verificado como correto

- **Contagem.** Doc 13 tem 4 itens e doc 17 tem 17 linhas: 21 bens 2018, todos presentes uma vez (8 em pares, 13 sem par). Os 8 itens do doc 05 2025 aparecem uma vez cada. O doc 02 2025 está zerado (R$ 0,00), assim como o doc 00.
- **Pares 2018 × 2025:**
  - 01 e 02 pela mat. 1.412 + área (1.800 / 1.080 m²);
  - 03 pela mat. 6.262;
  - 04 pelas mat. 1.941 e 7.104;
  - 05 pelas 9 matrículas (doc 17, seq. 1);
  - 06 pela mat. 11.090 + 4,0 ha + vendedor Dirceu Prestes e Lairton (seq. 7, não a parcela de 106,53 ha);
  - 07 pela mat. 7.391 (seq. 13);
  - 08 pela placa JTB 8585 (seq. 17).
- **Totais.** Itens do doc 05 = R$ 75.410.201,00. Executado de Rafael = R$ 79.491.957,50, diferença de R$ 4.081.756,50 (= 4.050.756,50 do item 01 + 31.000,00 do item 08). Doc 13 2018 = R$ 34.404.877,83. Doc 17 2018: bruto R$ 24.112.390,00, recebido R$ 12.058.000,00, líquido R$ 12.054.390,00. Base A = R$ 46.459.267,83. Base B (C28) = R$ 48.730.897,83, conferida pela fórmula da 2_0.
- **Valores de Rafael por seção.**
  - Balanço do grupo (p. 1–4): igual ao doc 05.
  - Avaliação executada (p. 7–15): igual ao grupo em todos os itens.
  - Balanço executado (p. 16–19): só os itens 01 (R$ 17.555.000,00) e 08 (R$ 85.000,00) diferem.
  - Lista xlsx, Diretrizes (p. 4–5) e bens escolhidos (p. 22–23): itens 04, 05, 06 e 08, subtotal R$ 50.914.157,50.
  - Concordâncias expressas: 02 (p. 10), 03 (p. 11), 04 (p. 12), 06 (p. 14), 07 (p. 15).
- **Coluna "Na lista de Rafael" e C0.** Lista certa nas 24 linhas. C0: item 01 = R$ 17.555.000,00 (maior, fica na empresa); item 08 = R$ 54.000,00 (menor, na lista); demais iguais. Soma C0 = R$ 79.460.957,50.
- **Localização em 2025:**
  - lote 12 em 11 T8R5 (13,0 ha; R$ 520.000,00 atual / R$ 143.000,00 antigo);
  - lote 89 em 11 T8R9 (R$ 250.200,00 nas duas versões);
  - lote 33 em 11 T9R2 (R$ 40.000,00 / R$ 75.500,00);
  - 69,8 ha em 08 T4R2 (R$ 3.500.000,00) e 4,5 ha em 08 T4R3 (R$ 250.000,00), iguais no antigo;
  - lote 52 também em ANTIGO 11 T9R3 (R$ 226.500,00).
  Trator Valtra, benfeitorias, mat. 982, 11.661 e 2.424/2.425 não aparecem em nenhum inventário 2025, atual ou antigo. A ocorrência de "982" no doc 04 é o ano de um caminhão. Sede Campestre só no doc 05.
- **Observações mapeadas.** GERAL H30/H34/H35/H38 → 01, H46 → 04, H52 → 05, H63 → 06, H230/B230 → s10, B235 → s12, F947–F950 → s03–s06, F986 → soja Rigo, F987 → soja Lorenzet. As abas das empresas repetem a GERAL, sem perda. Extras, Acertos e Acordos (linhas 36–43) e questionamentos não têm observação destas empresas.
- **Registro de data-base.** RDB-1 (arq. 1 p. 1 "termo final ... outubro/2025"; p. 9 "data da elaboração do inventário"), RDB-2 (arq. 1 p. 9, FIPE out/2025 R$ 56.854,00; PDF p. 15, R$ 54.745 e R$ 85–130 mil), RDB-3 (arrendamento 02/05/2023–30/04/2026) e RDB-4 (p. 8–9) conferem. Rovílio não aparece nos docs 02, 05, ANTIGO 05 nem no docx; aparece no doc 00.
- **Candidatos à Parte 6.** P6-01, P6-02, P6-04 (R$ 325.700,00), P6-06 e P6-07 (doc 11 T5R2, R$ 887.000,00) conferem com as fontes. P6-03 confere nos números (SEI 2018 lote 14, R$ 280.000,00; GERAL C229). SEI candidato nº 2 confere.
- **Perguntas.** Q1, Q3, Q6, Q8, Q9 e Q10 conferem (páginas e linhas citadas corretas). E1 e E3 estão de fato aprovados no log da análise de ACERTOS.
- **Regularização de R$ 400.000,00.** Não está abatida em nenhum total: doc 05 = soma dos itens, doc 00 = R$ 75.410.201,00, e Rafael também não abate.
- **Decisões.** Nenhuma decisão é apresentada como aprovada: Decisão = "Pendente", C0 rotulado como referência, `decisoes.json` vazio e contagem 0 no Resumo.
