# Notas de extração — Inventário SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA., dezembro de 2018

Produto: `C:\entrega-cisao-siviero\02_trabalho\sei-2018-itens.xlsx` (abas Itens, Lotes, Quantidades, Processos, Compromissos, Totais, Conferencia).

Fontes lidas (não alteradas):

- DOCX: `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = ... Dezembro de 2018.docx` (377 KB; metadados: criado em 01/12/2015 por "Luciano", editado no Google Docs; assinado "Veranópolis – RS, 31 de dezembro de 2018").
- XLSX: mesmo nome com extensão `.xlsx` (40 KB, gravado em 26/10/2026). Não é uma versão da família: é uma transcrição estruturada do docx feita anteriormente (abas Resumo, Bens, Bens em outras empresas, Processos, Credores, Aluguéis, Saldo de lotes, Investimentos, Sócios, com fórmulas de soma e totais transcritos).
- Contexto: `Material Siviero\Estudo - lotes que saíram 2018-2025.md`.

Método: Python (python-docx + openpyxl), script `build.py` no scratchpad `sei2018`. Todos os parágrafos e todas as células de todas as tabelas foram percorridos na ordem do corpo do documento; o docx não tem comentários, nem alterações controladas, nem notas de rodapé (verificado no XML: 0 `w:ins`, 0 `w:del`, 0 comentários; cabeçalho só "Página X de Y"). Números no formato brasileiro foram convertidos para float; nenhum número foi inventado ou estimado. Identificadores de lote/quadra foram mantidos como texto.

## 1. Estrutura do documento (na ordem)

| Ordem | Bloco | Tabelas | Linhas de item | Total declarado |
|---|---|---|---|---|
| 1 | Cabeçalho (CNPJ 89.472.195/0001-86, endereço) | — | — | — |
| 2 | Quadro societário real, de fato e de direito (7 sócios) | 1 | 7 | R$ 10.000,00 = 100% |
| 3 | "Constituição do novo quadro societário a contar de 01.12.2003 – NÃO REAL" (Pedro 99%, Elgides 1%) | 1 | 2 | R$ 100.000,00 |
| 4 | Veículos | 1 | 17 | R$ 953.000,00 |
| 5 | Imóveis e móveis | 1 | 2 | R$ 110.000,00 |
| 6 | Máquinas | 1 | 13 | R$ 1.350.000,00 |
| 7 | Lotes urbanos em Veranópolis e Bento Gonçalves | 1 | 12 | R$ 2.600.000,00 |
| 8 | Áreas urbanas e rurais em Veranópolis e Caxias do Sul | 1 | 10 | R$ 6.850.000,00 |
| 9 | Loteamentos executados e registrados | 4 (itens 01–46; cabeçalho repetido no meio da 3ª tabela) | 46 | R$ 64.225.000,00 |
| 10 | Loteamentos executados e não registrados | 2 | 7 | R$ 25.690.000,00 |
| 11 | Loteamentos em execução e não registrados | 2 | 6 | R$ 5.400.000,00 |
| 12 | Observação sobre o critério (lotes a R$ 100.000,00; só o valor investido nas obras) | parágrafos | — | — |
| 13 | Loteamentos contratados a executar ou em fase de execução e licenciamentos | 4 | 25 | R$ 1.250.000,00 |
| 14 | Observação sobre o critério (R$ 50.000,00 por loteamento = investimentos iniciais) | parágrafos | — | — |
| 15 | Processos judiciais (coluna "Comarca" traz comarca + nº do processo) | 3 | 19 | R$ 6.278.430,81 + 30% = R$ 8.161.960,05 |
| 16 | Credores e compromissos ("Em janeiro de 2019") | 3 | 32 | R$ 10.120.000,00 (parágrafo solto) |
| 17 | Aluguéis (receitas) | 1 | 16 | "+/- R$ 36.000,00, por mês" |
| 18 | Quantidade de lotes por loteamento | 1 (duas partes) | 35 + 5 | 447 + 245 = 692 lotes |
| 19 | Planilha de investimentos / despesas | 2 | 19 | R$ 10.376.000,00 |
| 20 | Fechamento: total geral, ações, dívidas, previsão, total líquido | 1 | 5 | R$ 108.428.000,00 → R$ 75.146.039,95 |
| 21 | Bens da SEI relacionados nos inventários de outras empresas (RÁDIO, OFICINA, AGROPECUÁRIA, HOTEL) | 5 tabelas, 3 delas sem linha de cabeçalho | 9 | sem total próprio |
| 22 | Sub-total real, soma de volta dos passivos, total geral e real | 1 | 5 | R$ 105.940.226,35 → R$ 139.222.186,40 |
| 23 | Valor real de cada sócio (só percentuais; coluna R$ em branco) | 1 | 7 | 100% |
| 24 | Observações finais (contratos de loteamento, dação, 7% impostos, 6% comissão) e assinaturas | parágrafos | — | — |

Número de página: o docx não guarda quebras renderizadas (0 `lastRenderedPageBreak`; 7 quebras explícitas + 9 `pageBreakBefore`), por isso a coluna "página/ordem" da planilha traz a **ordem no documento** (`Tnnnn-Rnnn` = índice do elemento do corpo + linha da tabela), não o número da folha. O próprio documento cita "Folhas 20, 21 e 22" (processos) e "Folhas 23, 24 e 25" (dívidas).

## 2. Como cada aba foi montada

**Itens (147 linhas)**: 138 itens dos 9 blocos de bens (17 + 2 + 13 + 12 + 10 + 46 + 7 + 6 + 25) + 9 itens de "bens relacionados nos inventários de outras empresas". A coluna "entra no TOTAL GERAL DA EMPRESA" separa os dois grupos (os 9 só entram no SUB-TOTAL REAL). Campos derivados por expressão regular sobre o texto do item:

- *loteamento*: "Loteamento X, localizado..." ou "no LOTEAMENTO X" (lotes de Bento Gonçalves e lote 05 "LOTEAMENTO BOFF").
- *cidade*: "na cidade de X – RS", "no Município de X – RS" ou, na falta, "Comarca de X".
- *qtde lotes (saldo / parte Siviero)*: "num saldo de N lotes"; "num total de (aproximadamente) N lotes" nos blocos executados; "SIVIERO: N LOTES" nos loteamentos em execução. *qtde lotes total do loteamento*: "Total de N lotes", "total = 102 lotes" ou o "num total de aproximadamente N lotes" dos loteamentos em execução. O item 09 das áreas (condomínio de chácaras) recebeu 10 na coluna de quantidade com observação "chácaras".
- *área*: "Área total de X m²" (ou a única área em m² do texto; se há várias, ficou em branco e a lista foi para a observação); hectares pela forma "= X hectares" ou a primeira ocorrência. Áreas construídas (891,00 m² do pavilhão da Sede Campestre) não entram.
- *matrícula*: todas as "Matrícula(s) nº ..." do texto; pontilhado → "(em branco)".
- *valor unitário*: "ao valor de R$ X, cada lote"; "R$ VARIÁVEL" → coluna "base" traz a abertura (ex.: "01 LOTE – R$ 550.000,00; 01 LOTE – R$ 450.000,00; 37 LOTES – R$ 230.000,00"); R$/m² e R$/hectare nos bens de outras empresas.
- *valor total (entra no total da seção)*: o último valor da célula "Valor". Casos especiais: Alto das Graças I — a célula traz "R$ 11.730.000,00 / Dividido por 2 / R$ 5.865.000,00" → bruto 11.730.000 e total 5.865.000; loteamentos em execução — a célula traz "bruto / (-) custo para conclusão / líquido", mas o total da seção (R$ 5.400.000,00) é a soma do "VALOR DA OBRA JÁ REALIZADA" de cada item (observação do próprio documento), e foi esse o valor usado como "valor total", com bruto/custo/líquido em colunas próprias.
- *qtde × unitário*: conferência automática; 60 itens conferíveis, todos CONFERE (inclusive os VARIÁVEL e os 36 lotes valorados do Salvador).
- *observação textual*: parágrafos do item após o primeiro + marcas no texto (VENDIDO, Foi roubada, Hipotecado, SALA ALUGADA, "ainda está em nome de") + texto extra da célula de valor + avisos de campos em branco.

**Lotes (7 linhas)**: o documento só identifica lote + quadra em dois lugares: os 6 lotes de Bento Gonçalves (itens 07–12 dos lotes urbanos: Ouro Verde 569/46 e 559/47 na quadra 216, 247/169 na quadra 221; Verona 188/62 na quadra 163/G; Colina do Sol 112/51 e 126/50 na quadra 188/B, estes dois VENDIDOS com valor R$ 0,00) e o Lote 09 – Quadra 1729 do Altos da Julieta, dado em pagamento no acordo Silvestrin (processo 09; é saída, não estoque — marcado na coluna "item de origem"). Não há lista lote a lote de nenhum loteamento (nem do Lago Negro, nem do Alto das Graças, nem do Altos da Julieta); os "LOTE Nº 1054 / 0024 / 1543" das áreas são números de lote urbano sem quadra e ficaram só na aba Itens. Os itens "PAGO COM LOTE" (credor 19) e "diferença do lote no Loteamento Baú" (credores 28 e 29) não identificam o lote.

**Quantidades (40 linhas)**: transcrição da tabela "Nome do Loteamento / Quantidade Lotes", com o grupo tal como o documento separa (35 linhas até o total de 447; 5 linhas até o total de 245). Os números estão grafados com espaço entre os dígitos no docx ("1 4", "1 0 5"); foram lidos removendo os espaços.

**Processos (19 linhas)**: nº, parte(s) (extraída do 1º parágrafo), objeto (1º parágrafo), loteamentos citados, comarca, nº do processo (os itens 16 e 17 não têm número judicial: "Inquérito Civil 00.771.00015/2018" e "Processo Administrativo nº 1198/2018 Auto de Infração nº 003/2018"), valor da causa, valor estimado (só o item 02: "R$ 52.818,30 + correção de 05 anos = +/- R$ 200.000,00"), prejuízo estimado (só o item 10: R$ 1.750.000,00), total do acordo (só o item 09: R$ 177.000,00), situação/acordo e texto completo.

**Compromissos (67 linhas)**: 32 credores e compromissos + 19 linhas da planilha de investimentos (pagamentos de pavimentação, compras de lotes e áreas, dívidas pagas) + 16 aluguéis recebidos (receitas; 4 deles são valores anuais). Para os credores: valor no documento = valor final da célula (ex.: "12 parcelas de +/_ R$ 10.000,00 = R$ 120.000,00" → 120.000); situação PAGO → saldo 0; "R$ 50.000,00 - PAGO / FALTA: R$ 20.000,00" → saldo 20.000; "R$ PAGO" sem valor (item 02) → valor em branco, saldo 0; juros extraídos de "a juros de X%, ao mês".

**Totais (29 conferências + transcrição literal dos 21 totais declarados)**: ver seção 3.

**Conferencia**: docx × xlsx, campo a campo (seção 4).

## 3. Totais: CONFERE / DIVERGE

| Total | Declarado | Recomputado | Status |
|---|---|---|---|
| 9 totais de seção (veículos … contratados) | — | — | todos CONFERE |
| TOTAL GERAL DA EMPRESA | 108.428.000,00 | 108.428.000,00 | CONFERE |
| Ações judiciais — "VALOR TOTAL ... PELO VALOR DECLARADO NO PROCESSO" | 6.278.430,81 | 5.040.034,45 (soma dos 19 "Valor da causa") | **DIVERGE** (−1.238.396,36) |
| idem, usando "+/- R$ 200.000,00" no item 02 | 6.278.430,81 | 5.187.216,15 | **DIVERGE** (−1.091.214,66) |
| + 20% honorários + 10% custas | 1.883.529,24 | 1.883.529,24 (30% de 6.278.430,81) | CONFERE |
| Ações com honorários e custas | 8.161.960,05 | 8.161.960,05 | CONFERE |
| Credores e compromissos | 10.120.000,00 | 10.119.900,00 (saldos em aberto) | **DIVERGE** (−100,00; o item 27 vale R$ 39.900,00 — provável arredondamento) |
| Aluguéis "+/- R$ 36.000,00 por mês" | 36.000,00 | 39.692,00 (16 itens) / 34.930,00 (só os 12 mensais) | **DIVERGE** (o documento diz "+/-"; os 4 anuais somam R$ 4.762,00/ano) |
| Planilha de investimentos | 10.376.000,00 | 10.376.000,00 | CONFERE |
| Lotes executados e registrados / não registrados / geral | 447 / 245 / 692 | 447 / 245 / 692 | CONFERE |
| TOTAL LÍQUIDO DA EMPRESA | 75.146.039,95 | 108.428.000 − 8.161.960,05 − 10.120.000 − 15.000.000 | CONFERE |
| Bens em outras empresas (sem total no doc.) | — | 30.794.186,40 | — |
| SUB-TOTAL REAL DA EMPRESA | 105.940.226,35 | 75.146.039,95 + 30.794.186,40 | CONFERE |
| TOTAL GERAL E REAL DA EMPRESA | 139.222.186,40 | sub-total real + 8.161.960,05 + 10.120.000 + 15.000.000 | CONFERE (o documento soma de volta os passivos que havia deduzido) |
| Quadros societários (R$ 10.000,00 e R$ 100.000,00) | — | — | CONFERE |
| Percentuais "valor real de cada sócio" | 100% | 6 × 16,333% + 2% = 99,998% | CONFERE (arredondamento) |

A divergência das ações judiciais é a única relevante: o documento não mostra como chegou a R$ 6.278.430,81, e nenhuma combinação simples dos valores listados (valor da causa, estimativa do item 02, prejuízo de R$ 1.750.000,00 do item 10) reproduz o número. Fica registrada como DIVERGE sem hipótese.

Conferência interna por item (qtde × valor unitário): 60 itens conferíveis, 0 DIVERGE. Alto das Graças I: 102 × 115.000 = 11.730.000 (bruto) e 51 × 115.000 = 5.865.000 (parte da Siviero) — ambos batem. Loteamentos em execução: bruto = lotes da Siviero × valor unitário em todos os 6 itens; líquido = bruto − custo para conclusão em todos.

Cruzamento entre a tabela de quantidades e os itens: os 46 itens de "executados e registrados" somam 410 lotes; a tabela dá 447 porque inclui Barbiero (20) e Carnevalli (17), que nos itens estão em "executados e não registrados" (410 + 37 = 447). Na parte "não registrados" a tabela lista Salvador com 45 (o item valora 36: 8 + 12 + 16, com "04 LOTES VENDIDOS" e "05 LOTES DESTINADOS PARA A ÁREA VERDE") e Colina do Retiro II com 105 (o item diz "aproximadamente 100", de um total de 210). Os loteamentos em execução (Caminhos do Sol 290, Santa Corona III 44, Vale das Colinas 180, Falavigna 50, Zanella 29, Caminho das Flores 81 = 674 lotes da Siviero) não entram nos 692.

## 4. Conferência docx × xlsx

A transcrição .xlsx bate com o docx em 851 de 852 campos comparados (descrição, número do item, valores, comarca, nº do processo, quantidades, grupo, periodicidade dos aluguéis, totais transcritos no Resumo):

| Aba do xlsx | Campos comparados | Iguais |
|---|---|---|
| Bens (138 itens) | 434 | 434 |
| Bens em outras empresas (9) | 28 | 28 |
| Processos (19) | 77 | 76 |
| Credores (32) | 97 | 97 |
| Aluguéis (16) | 49 | 49 |
| Saldo de lotes (40) | 121 | 121 |
| Investimentos (19) | 39 | 39 |
| Resumo (7 totais transcritos) | 7 | 7 |

Única diferença: **processo 02** — o xlsx registra R$ 200.000,00 como "valor da causa" (a estimativa "+/- R$ 200.000,00" do texto); o docx declara "Valor da causa: R$ 52.818,30 + correção de 05 anos = +/- R$ 200.000,00". Na planilha nova o valor da causa é 52.818,30 e a estimativa vai em coluna própria. É diferença de leitura, não de transcrição. Fora isso, o xlsx não tem lista de lotes individuais nem conferência qtde × unitário, e as observações do Resumo do xlsx (ações 5.187.216,15 ≠ 6.278.430,81; credores −100; aluguéis 39.692 ≠ 36.000) coincidem com as divergências encontradas aqui.

## 5. Leituras de partes ambíguas

1. **Alto das Graças I (item 05)**: célula de valor com três linhas "R$ 11.730.000,00 / Dividido por 2 / R$ 5.865.000,00". Lido como bruto do loteamento inteiro (102 lotes) e parte da Siviero (51 lotes); o total da seção (R$ 64.225.000,00) só fecha com 5.865.000.
2. **Loteamentos em execução (itens 01–06)**: célula de valor com "bruto / (-) custo / líquido" (ex.: 29.000.000 − 13.700.000 = 15.300.000), mas o total da seção é R$ 5.400.000,00 = soma dos "VALOR DA OBRA JÁ REALIZADA" (2.800.000 + 500.000 + 600.000 + 100.000 + 1.000.000 + 400.000), conforme o parágrafo "FOI CONSIDERADO ... SOMENTE O VALOR INVESTIDO NAS OBRAS REALIZADAS". A soma dos líquidos (R$ 34.250.000,00) não é usada pelo documento.
3. **Loteamentos contratados (25 itens)**: todos com "num total de aproximadamente ...... lotes" (em branco) e R$ 50.000,00 cada, que o documento explica como valor investido igualitariamente (contratos, topografia, profissionais, multas etc.). Item 12 (Família Menegol): "FEITO EM DISTRATO EM NOVEMBRO DE 2016", mas continua valorado em R$ 50.000,00.
4. **Altos da Julieta I e II (itens 33 e 34)**: "num total de aproximadamente 20 lotes" cada, com "Total de 47 lotes" / "Total de 42 lotes" (total do loteamento); valor 20 × 90.000 = 1.800.000 por etapa. A tabela de quantidades junta as duas em uma linha com 40.
5. **Lago Negro (item 23)**: 35 lotes (5 × 350.000 + 30 × 230.000 = 8.650.000). Nenhuma identificação de lote/quadra; os 7 lotes de APP aparecem só no processo 10 (prejuízo estimado R$ 1.750.000,00; "deverão ser permutados pelo mesmo número de lotes") e a troca de 1 lote no processo 15 (acordo Wasen, 2018).
6. **Bens relacionados nos inventários de outras empresas**: os números de item (06, 06, 07, 02, 03, 04, 05, 06, 04) são os números que os bens têm nos inventários das outras empresas (RÁDIO, OFICINA, AGROPECUÁRIA, HOTEL), por isso se repetem; três das cinco tabelas não têm linha de cabeçalho. O item 02 (10,2 ha em Cotiporã) está marcado "V E N D I D A – Família ANGELI – vendido em 2018 por R$ 510.000,00 à vista", mas continua valorado em R$ 408.000,00 e somado no sub-total real. O item 07 (BR-470, Oleoplan) reúne dois terrenos (1.942,03 m², mat. 17.409, e 2.850,00 m², mat. 6.136) e pavilhões de 4.170,00 m²; a coluna área m² ficou em branco e os números estão na observação.
7. **Credores**: a célula do item 02 (caminhão Dalfovo) é "R$ PAGO" sem valor; do item 19 "R$ 50.000,00 PAGO COM LOTE"; do item 04 "R$ 50.000,00 - PAGO / FALTA: R$ 20.000,00". O total R$ 10.120.000,00 é um parágrafo solto após a terceira tabela, sem rótulo; o parágrafo "Em janeiro de 2019." dá a data da posição.
8. **Processos**: item 13 (Família Farina) diz "R$ 1.013.580,55 = TOTAL, sendo: R$ 600.000,00 – Indenização + R$ 400.000,00 – Honorários" — as parcelas somam 1.000.000,00, não 1.013.580,55; mantido o total declarado. Item 14 já sentenciado a favor da empresa, mas o valor (R$ 332.150,34) permanece na lista.

## 6. Placeholders e campos em branco no documento

- Máquinas 04, 05, 11, 12: "ano....." (sem ano).
- Lotes urbanos 05 (Loteamento Boff, em nome de Cleri Fávero): medidas, área e matrícula em branco. Lotes urbanos 07–12 (Bento Gonçalves): área e matrícula em branco (há o cadastro na Prefeitura).
- Áreas 04 (Lote nº 0024, Santa Rita): vendedor e matrícula em branco (só a inscrição municipal 113 0024 001 001).
- Loteamentos 35 e 36 (Desmembramento Tedesco I e II): "Total de ..... lotes. Área total de..........m²", 00 lotes, R$ 0,00.
- Os 25 loteamentos contratados: quantidade de lotes "......".
- Sede Campestre (bens em outras empresas, HOTEL): "casa de alvenaria ... = ...............m² de área construída".
- Tabela final "VALOR REAL DE CADA SÓCIO": coluna "R$" vazia em todas as linhas (só percentuais).
- 15 loteamentos executados e registrados com saldo 00 e valor R$ 0,00 (Morada da Colina, Bella Visione, Lovison, São Roque, Tedesco I e II, Colina do Lindóia, Fracionamento Pagliari, Fracionamento Primavera, Martini, Conte, São Luiz, Valverde IV, Medianeira, Zanetti).

## 7. Outras coisas estranhas

- Na célula de valor do item 01 (Sol Nascente I) existe o texto "colina do basalto" em fonte de 1 pt (quase invisível), antes de "R$ 2.800.000,00". Provável resto de edição; registrado na observação do item.
- Nos quadros societários, os nomes IRINEU SIVIERO, PEDRO SIVIERO e ROVILIO SIVIERO estão tachados (strikethrough) no quadro real e na tabela final por sócio; PEDRO SIVIERO também está tachado no quadro "NÃO REAL". O documento não explica a marca.
- Veículo 05 (Dodge 1978): "VENDIDO PELO VALOR DE R$ 30.000,00, EM TROCA DE PALANQUES E POSTES DE MADEIRA", mas segue valorado em R$ 30.000,00 e somado no total de veículos. Já os lotes urbanos 11 e 12 (VENDIDO) e a retroescavadeira 07 (roubada) estão com R$ 0,00.
- Lote urbano 04: "Hipotecado em favor da Receita Estadual – Dívida do ICMS" (R$ 60.000,00).
- Veículos 13, 14 e 15 trazem nomes de pessoas após o ano ("– Pedro Siviero", "– Válter Siviero", "– João Lunelli"), sem dizer se é posse, uso ou propriedade.
- Investimento 17: "Pago pavimentação asfáltica Loteamento Jardim Itália – R$ 1.265.000,00" com valor na coluna de R$ 700.000,00 (o texto cita um valor e a coluna outro; mantidos os dois).
- O cabeçalho da tabela "Loteamentos executados e registrados" se repete no meio da terceira tabela (entre os itens 29 e 30), sem quebra de tabela.
- A segunda tabela do quadro societário tem o título "NÃO REAL" como Heading 1 no corpo do documento.

Gerado em 28/10/2026.
