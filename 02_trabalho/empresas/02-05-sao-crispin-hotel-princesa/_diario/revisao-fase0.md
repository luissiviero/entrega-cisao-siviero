# Revisão independente da Fase 0 — empresas 02 e 05

Revisor: agente de contexto novo, 28/10/2026. Objeto: `_scripts/fase0_extracao.py` e `fase0-extracoes-02-05.xlsx`.
Método: rodei o script (59 conferências, 2 DIVERGE, as mesmas do diário; a planilha regravada ficou igual à anterior em todas as abas) e reextraí todas as fontes por conta própria: docx com python-docx (incluindo tabelas aninhadas, cabeçalhos, caixas de texto e revisões; não há nenhuma), PDFs com pymupdf (texto e cores), xlsx com openpyxl (fórmulas, valores, comentários, preenchimentos e cores de fonte). Recalculei os números à mão. Não alterei nada além deste arquivo.

## Achados

### A1 — ALTA — Créditos de soja do docx "Áreas vendidas" não extraídos (um deles é da Agropecuária)
- **O que está errado:** o script lê só a tabela de vendas do docx (`sale_rows`, L289–321). Os parágrafos "OBSERVAÇÕES" logo abaixo dela ficam de fora, embora o papel da fonte (L109) diga "áreas vendidas, soja, arrendamento". Nenhuma aba lê esses parágrafos direto do documento. Eles só aparecem de forma indireta, pela planilha Google (GERAL B986/B987, escritos pela sessão SEI).
- **Evidência (docx `HOTEL PRINCESA … SEM OBRAS.docx`, parágrafos após T1R4):**
  - Rigo, T1R1: as parcelas de maio/2021 e maio/2022 (20.150 sacas cada) estão marcadas "(FALTA RECEBER)". O texto traz 40.300 − 2.606 (vendidas em 2024 para a escavadeira) = 37.694 sacas "(CRÉDITO)", sem valor em R$.
  - Lorenzet: 2.606 + 3.206 + 3.591 + 3.682 = 13.085 sacas não vendidas. As de 2021–2023 estão anotadas "HOTEL/OLEOPLAN" e as de **2024 (3.682 sacas) "AGROPECUÁRIA/OLEPOPLAN"**. Ou seja, a empresa 02, com inventário 2025 = R$ 0,00, tem um crédito citado em documento do grupo.
  - Consequência para o "TOTAL RECEBIDO" de R$ 12.058.000,00: ele inclui 2 × 20.150 × R$ 80 = R$ 3.224.000,00 que o docx de 2025 diz não terem sido recebidos.
  - O número 2.606 aparece duas vezes, como soja Rigo vendida em 2024 e como soja Lorenzet de 2021. Pode ser coincidência ou erro de digitação no documento.
- **Correção sugerida:** extrair esses parágrafos numa aba "Créditos de soja" (por safra, origem e empresa). Conferir 40.300 − 2.606 = 37.694 e a soma 13.085. Registrar a divergência "Agropecuária = 0 × 3.682 sacas atribuídas à Agropecuária" e o "FALTA RECEBER" contra o rótulo "TOTAL RECEBIDO". Levar ao Luís (crédito não inventariado no Hotel e na Agropecuária).

### A2 — ALTA — Fazenda (item 05 de 2025 e item 01 de 2018) marcada "VENDIDA" inteira; o texto da venda parcial se perde
- **O que está errado:** `situacao` (L274–279) marca "VENDIDA" sempre que a célula de valor contém VENDIDA. A célula diz "VENDIDA para a Família Rigo, **somente 83,47 ha** desta área total, da Matrícula nº 00.760 **407,53 ha**". Na aba "Itens 2025" o item 05 (atual e antigo) aparece como Área "491,0 ha", Situação "VENDIDA". O texto da célula de valor (`texto_valor`) e o comprador não estão em nenhuma coluna de nenhuma aba (L1196–1207 e L1222–1235). "Itens 2018" (seq 1) e "Áreas da fazenda" repetem "VENDIDA | RIGO" para a parcela de 491 ha.
- **Evidência:** doc 05 2025 T5R3 (R$ 44.828.300,00); ANTIGO 05 T5R3; doc 17 2018 T3R2 ("V E N D I D A Somente 83,47 ha desta área total Família Rigo"). Busquei "407,53" e "somente" no xlsx: só aparecem em Conferências e na Lista de Rafael.
- **Correção sugerida:** criar a situação "VENDIDA PARCIAL (83,47 ha)", uma coluna "área avaliada" (407,53 ha em 2025; 491 ha em 2018, que avaliava inclusive a parte vendida), e as colunas `comprador` e `texto_valor` em "Itens 2025" e "Itens 2018".

### M1 — MÉDIA — Benchmark 2018 digitado (46.459.267,83); a conferência compara a soma consigo mesma, e a planilha Google usa outra base
- **O que está errado:** a L651–656 compara `liquido + total(t18_13)` com `Decimal("46459267.83")`, número digitado que não está em nenhum documento (o doc 27 não traz essa soma). É a mesma soma das duas fontes, então a conferência passa por construção, o que contradiz a docstring ("nothing typed from memory").
- **Evidência / recálculo:** a planilha 2_0 usa para o mesmo perímetro (Hotel 2025 = Hotel 2018 + fazenda) a base 2018 de **R$ 48.730.897,83** (GERAL C28 / aba 5.Hotel C2 = 34.404.877,83 + 13.588.520 + 176.000 + 300.000 + 226.500 + 35.000). Diferença: R$ 2.271.630,00, conciliada assim:
  - o "líquido" do doc 17 abate o **preço de venda** (12.058.000), não o **valor de inventário** das áreas vendidas. Esse valor é 9.118.570 = Rigo 83,47 × 44.000 + 3.163.050 (= 6.835.730, que é o C944 da planilha) + Pellegrini 1.874.840 + Angeli 408.000. Ganho na venda: 2.939.430;
  - menos o trator (116.600) e os lotes 03/04/05 (551.200), que não estão no Hotel 2025;
  - 2.939.430 − 116.600 − 551.200 = 2.271.630. Fecha.
  - Além disso, o preço Pellegrini inclui R$ 100.000 de 41 cabeças de gado, e o preço Rigo inclui R$ 3.224.000 "FALTA RECEBER" (ver A1).
- **Correção sugerida:** tirar a constante. Mostrar as duas bases (líquido do doc 17 e base "itens que permaneceram") com a conciliação acima. A escolha da base é decisão de premissa (benchmark 2018) e cabe ao Luís.

### M2 — MÉDIA — Conferências de área com valores digitados; o total declarado 411,43 ha nunca é conferido
- **O que está errado:** a L792–798 compara o remanescente calculado pelas parcelas com `Decimal("411.53")`, montado a partir do próprio resultado ("407,53 + 4,0"). O doc 17 e o docx de áreas **declaram 411,43 ha**, e os docs 05/02 de 2025 usam **411,4 ha** (base do ITBI e do arrendamento a–d). Nenhum desses valores declarados é conferido. Também são digitados em vez de lidos do texto: 83,47 (L777), 190 (L783, L792), 491,0 (L694, L757), 407,53 (L736), 686,75 (L772).
- **Recálculo:**
  - 686,75 − 190,00 − 85,32 = 411,43 (confere com o declarado);
  - pelas parcelas: 686,75 − 190 − 85,22 = 411,53 = 407,53 + 4,0. A diferença de 0,10 ha está nos 85,32 declarados para a venda Pellegrini (DIVERGE já registrado);
  - 155 + 30 + 157,5 + 68,9 = 411,4 (confere com o texto de 2025).
  - Matrícula 00.760 = 48 + 100 = 148 ha = área Jane/João Paulo; 148 − 83,47 = 64,53 (confere).
- **Correção sugerida:** ler 686,75 / 190 / 85,32 / 411,43 dos parágrafos do doc 17 e do docx, e 411,4 do doc 05. Conferir: (a) a aritmética declarada (OK); (b) 411,53 contra 411,43 (diverge 0,10 ha); (c) 411,53 contra 411,4.

### M3 — MÉDIA — Faixa de mercado de Rafael omitida (p. 9) e ruído nas faixas da p. 15
- **O que está errado:** `RANGE_PATTERNS` (L363–371) não cobre "gira em torno de R$ X a R$ Y". Fica de fora a frase da p. 9 (item 01): "A média para imóveis e apartamentos novos ou bem localizados no Centro de Veranópolis gira em torno de R$ 5.600 a R$ 6.200 por m²." A aba de faixas tem 8 linhas; o PDF tem 9 frases de faixa ou referência.
- Na p. 15, a primeira linha começa com um "R$ " solto (sobra da linha do valor), e a referência FIPE "Aproximadamente R$ 54.745,00" aparece duas vezes (sozinha e dentro da frase "Varia de…").
- **Correção sugerida:** acrescentar o padrão genérico `R\$ ?[\d.,]+ a R\$ ?[\d.,]+`, remover duplicatas e limpar o prefixo.

### M4 — MÉDIA — Planilha Google: células relevantes fora das abas e faixas escolhidas
- **O que está errado:** a leitura (L985, L1007) cobre as abas 2/5/Extras e as linhas 11–15, 28–69, 941–950 e 986–987 da GERAL. Ficam de fora:
  - **GERAL 227–241 (áreas rurais da SEI):** A.3, linha 241 = Lote 52, mat. 7.391 (2018 R$ 226.500 → 2025 R$ 0), que em 2025 é o **item 07 do Hotel** (R$ 422.800). Linha 230 = lote 12 (Sangali), com comentário em fio B230 "DESCRIÇÃO ANTIGA … 13,0 ha × 11.000 = R$ 143.000"; o doc 17 de 2018 tem 20,5 ha / R$ 225.500. Linha 234 = lote 89; linha 235 = lote 33, com comentário B235. São os itens 03–06 da Agropecuária em 2018.
  - **GERAL H882:** "Quais áreas: Fazenda São Crispin? Áreas de compensação?" (custas de R$ 250.000 na SEI).
  - **Aba Resumo 78–94:** São Crispin, R$/ha 2018 × 2025 ("Valor por hectare em 2025 é o mesmo, independente da produtividade"); arrendamento estimado em 334.800 + 108.000 + 51.810; notas "* pegar contratos" e "* verificar se arrendo sobre toda área ou parcial".
  - **Resumo B73:** rotulado "Hotel Princesa dos Vales", mas aponta para GERAL!C8/E8 (caminhão-tanque da Abastecedora, 145.000 → 0). É erro da planilha e convém registrar.
  - **11.Empreendimentos B90 (lote 52) e I731; Acertos e Acordos B36–B42 (vendas São Crispin).**
- Fora isso, nas abas lidas não há comentários nem notas; as cores de preenchimento são só o zebrado. As únicas fontes coloridas (vermelho BF0000) são F943, F947–F950, F986 e F987, todas capturadas.
- **Correção sugerida:** incluir essas faixas e abas, marcando as observações.

### M5 — MÉDIA — Docx de áreas × doc 17: a conferência compara só totais e não pega mudanças de texto
- **Evidência:**
  - Angeli: docx T1R3 "10,2 hectares X R$ **50.000,00**"; doc 17 T4R5/T6R3 "X R$ **40.000,00**" (inventário de 2018 R$ 408.000; venda R$ 510.000).
  - Rigo: o docx acrescenta "= 118.100 SC SOJA" (37.500 + 4 × 20.150, confere), "37.500 sacas" à vista (3.000.000 / 80, confere) e "(FALTA RECEBER)" duas vezes.
  - A conferência da L665–670 passa porque os totais são iguais (12.058.000).
- **Correção sugerida:** fazer um diff de texto linha a linha (como no "2025 antigo x atual") e registrar as diferenças.

### M6 — MÉDIA — Notas de Rafael sem vínculo com o item; valor de 17.555.000 sem cálculo não aparece na planilha
- **O que está errado:** a aba "Rafael notas" lista só as páginas. Pelo PDF:
  - concordância expressa: p. 10 → item 02; p. 11 → 03; p. 12 → 04; p. 14 → 06 (cita "02., 03., 04., 05."); p. 15 → 07;
  - **não há concordância para 01 e 08**, os dois que ele reavalia;
  - o item 05 entra na lista de concordância da p. 14, embora as p. 13–14 recomendem avaliação independente e chamem o valor de "parâmetro preliminar".
  - No balanço executado (p. 16–17), o item 1 vale R$ 17.555.000,00, mas os componentes impressos logo abaixo são os do grupo (somam 13.504.243,50). Não há memória de cálculo. Esse fato está no diário, mas não na planilha.
- Conferi as cores do PDF: os itens em azul (#5B9BD5), que segundo a p. 21 são os escolhidos, são exatamente 04, 05, 06 e 08, nas três seções. Está correto.
- **Correção sugerida:** acrescentar uma coluna "item" e o "alcance da concordância" às notas, e uma linha "01: valor executado sem memória; componentes impressos = grupo".

### M7 — MÉDIA — Sobreposição com a SEI em 2018 capturada, mas não sinalizada
- **Evidência:** a aba "Bens em outras empresas" do inventário 2018 da SEI (xlsx L8–L13) lista como bens da SEI os itens 02–06 da Agropecuária (408.000 + 225.500 + 250.200 + 75.500 + 226.500 = **1.185.700**) e o item 04 do Hotel (Sede Campestre, **3.879.091,80**), no total **5.064.791,80**, que também estão nos docs 17 e 13 e, portanto, no benchmark 2018. O lote 52 (mat. 7.391) aparece no Hotel 2025 (item 07) e também na GERAL/SEI (A.3).
- Também não há tabela de correspondência 2018 → 2025 para os itens da Agropecuária que somem: trator 116.600, benfeitorias 300.000, lotes 03/04/05 e os itens 07/08 "NÃO DIVISÍVEL" (69,8 + 4,5 ha).
- **Correção sugerida:** criar uma conferência "possível dupla contagem SEI × 02/05 em 2018" e uma tabela de correspondência 2018 → 2025 por matrícula.

### B1 — BAIXA — "Totais 2018" mistura totais que não são do inventário
- `summary_rows` (L329–331) pega toda linha que começa com TOTAL/PARCIAL. Entram as linhas de capital social T1R8 e T2R4 (R$ 300.000 = 100%), a T8R8 vazia ("100,00% = R$") e duplicatas (T5R6/T7R1; T6R4/T7R2). O rótulo inclui o texto do valor.
- **Correção sugerida:** limitar às tabelas de bens e de vendas e tirar as duplicatas.

### B2 — BAIXA — "Resumos de capital"
- Nas linhas T3 do doc 00, o "Valor" registrado é a cota por sócio (12.568.366,83, `vals[0]`), não o total da empresa, e a coluna não diz isso.
- O parágrafo do doc 00 2025 sobre a venda das quotas de **Rovílio** (instrumento de 28/08/2025, "diluídas em partes iguais aos demais sócios") não é registrado. A conferência "Menção a Rovílio" (L1064) exclui justamente o doc 00.
- O doc 00 T2R12 tem "R$ 100.874,027,60" malformado, que o padrão MONEY leria como 100.874,027. Não afeta 02/05, mas é um risco se o resumo for reaproveitado.

### B3 — BAIXA — Aritmética interna não conferida pelo script (conferi: fecha)
- **Item 01 de 2025:** 4.320.000 + 3.845.155,20 + 574.239,60 + 1.728.000 + 1.200.000 + 1.836.848,70 = 13.504.243,50. A base dos 25% (7.347.394,80) é a soma dos componentes já depreciados.
- **Item 04 de 2025:** 4.230.000 + 664.686 + 150.000 + 250.000 + 266.171,50 = 5.560.857,50.
- **Itens 01 e 04 de 2018:** 17.475.786,03 e 3.879.091,80. Alzir: 37,275 × 44.000 + 69,225 × 22.000 = 3.163.050.
- **Pequenas inconsistências dos documentos:**
  - item 03: 750 × 3.333,34 = 2.500.005 (doc diz 2.500.000);
  - item 04: "25,00% (Vinte por cento)";
  - despesas de regularização: 100.000 + 246.840 + 5.000 + 5.000 = 356.840 contra "TOTAL de R$ 400.000" (arredondado, +43.160);
  - cota por sócio: 6 × 12.568.366,83 = 75.410.200,98, com rótulo "16,66%" mas cálculo de 1/6.
- **Correção sugerida:** acrescentar essas conferências, no mesmo padrão já usado para a fazenda.

### B4 — BAIXA — Lógica frágil no script (hoje dá o resultado certo)
- L617 `seq <= 8`: subtotal da fazenda por posição.
- L544 `zip(i25a_05, i25_05)`: pareia por posição, não por número de item.
- L394: o marcador de item só aceita 1–8, e qualquer linha com um dígito sozinho troca o item.
- L403: só capta valores que estão numa linha isolada "R$ x".
- L947–957: a conferência do Diretrizes procura o valor como substring nas p. 4–6, que incluem itens do Hotel Araçá. "85.000,00" casaria com "385.000,00" ou "185.000,00" se estivessem nessas páginas.
- L767: `area_ha` fica vazio para "20,5 ha (13,0 + 7,5)".
- L1039–1045: rótulos invertidos ("extraído" = contagem no arquivo da SEI; "declarado" = contagem do script).
- L540/L921/L927: contagens esperadas (8, 4, 0) digitadas.

### B5 — BAIXA — Detalhes de 2018 não sinalizados
- Parcela Garbin (T3R7): o cabeçalho diz 44,16 ha e a linha da matrícula "= 44,1 hectares". O vendedor fica vazio porque o texto diz só "Área Garbin".
- **Matrícula 11.090** aparece na parcela de 106,53 ha vendida aos Rigo e na de 4,0 ha que ficou (item 06 de 2025, "DENTRO DE UM TODO MAIOR"). Risco de título a registrar.
- O item 08 de 2018 cita a compra "pelo valor de R$ 50.000,00", que não é capturada.

### B6 — BAIXA — MD5 desatualizado na aba Fontes
- `achados_acertos` (`02_trabalho/analises/acertos-e-areas-vendidas/achados-para-empresas.md`): a aba registra `ca3c0b913bf4b53bf421efcd7b561d37`. O arquivo atual tem `0a0377229ad56ad7b1f0098288b0fbc5`, modificado em 28/10 às 17:06:18 por outra sessão, depois do registro.
- As outras 26 fontes conferem.
- **Correção sugerida:** recalcular o hash ao fechar a fase e registrar a versão usada.

## Verificado como correto
- **2025 atual:** 8 itens do Hotel (números, valores, matrículas, áreas) e item único "-.-" da Agropecuária (R$ 0,00, igual ao "valor inscrito"). Soma 75.410.201,00 = TOTAL GERAL = doc 00.
- **2025 antigo:** 8 itens, soma 83.640.801,00 = ANTIGO 00. Diferenças só nos itens 05 (−8.150.600; R$/ha de 130.000 para 110.000) e 06 (−80.000), e na grafia Gloria → Glorinha no item 07. Total −8.230.600.
- **2018:** Hotel com 4 itens, 34.404.877,83 = doc 27. Agropecuária com 17 linhas, sem faltas nem duplicatas e sem linhas de tabelas que não são de bens. Subtotal da fazenda 22.775.090; bruto 24.112.390; vendas 9.448.000 + 2.100.000 + 510.000 = 12.058.000; líquido 12.054.390 = doc 27. Situação de venda e compradores das 6 parcelas e do Angeli corretos (exceto a venda parcial, A2).
- **Áreas:** matrículas 2025 = 491,0 ha (10 linhas, 9 números); avaliadas 407,53 ha × 110.000 = 44.828.300; parcelas 2018 = 686,75; Rigo 83,47 + 106,53 = 190; Pellegrini 85,22 contra 85,32 declarados (DIVERGE correto); 2018 avaliava os 491 ha inteiros (17.261.200).
- **Rafael Hotel:** todos os valores das 5 seções estão nos itens certos.
  - grupo = 75.410.201;
  - "avaliação executada" repete os valores do grupo, inclusive 01 = 13.504.243,50 e 08 = 54.000;
  - executado: 01 = 17.555.000, 08 = 85.000 (repetido na p. 19), total 79.491.957,50;
  - quadro: 1/6 = 13.248.659,583, escolhidos 50.914.157,50, diferença 37.665.498 (cálculo 37.665.497,92);
  - escolhidos: 04, 05, 06, 08 (confirmado pelas cores);
  - Diretrizes p. 4–6 e `ativos-avaliados-valter.xlsx` com os mesmos 4 valores (incluindo 85.000 para o MB 709).
- **Rafael Agropecuária:** o PDF só reproduz as observações do grupo e a nota de incorporação; não tem balanço nem valor próprio (0 valores está certo).
- **Planilha Google:** as 25 observações das faixas lidas estão certas. Não há comentários nas abas 2/5/Extras. Os comentários em fio da GERAL (B223–B238) estão fora das faixas (ver M4).
- **Fontes:** 26 de 27 hashes conferem.
