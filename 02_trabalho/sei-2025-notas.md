# Notas da extração – Inventário 31.12.2024 da Siviero Empreendimentos Imobiliários Ltda. (documento 11)

Data da extração: 28/10/2026. Fonte (não alterada): `Material Siviero/2025/inventário-2025/11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx`. Saída: `sei-2025-itens.xlsx` (7 abas) nesta mesma pasta. Script: `scratchpad/sei2025/build.py` (python-docx + openpyxl; toda a leitura é feita em ordem do corpo do documento, parágrafo por parágrafo e célula por célula).

## 1. Metadados e forma do arquivo

- `docProps/core.xml`: criado em 29/10/2025 17:41 UTC, autor "Luciano"; sem data de modificação gravada. O arquivo na pasta do Drive tem data de 26/10/2026.
- O tamanho de 1,1 MB vem de **seis fontes TrueType embutidas** (Arial Narrow e Tahoma, ~1,8 MB descompactados), não de imagens. Há **uma única imagem**, `word/media/image1.png`, de 70 bytes, referenciada só pelo cabeçalho (`header1.xml`), que contém apenas o campo de número de página. Nenhuma imagem no corpo.
- **Sem alterações rastreadas e sem comentários**: zero elementos `w:ins`, `w:del`, `w:comment`, `w:drawing`, `w:pict`, `w:object`. Os 429 elementos `w:strike` do XML são todos `w:val="0"` (tachado desligado explicitamente): nenhum texto está riscado.
- Texto oculto: nenhum (`w:vanish` = 0).
- Único realce de cor no documento: fundo ciano na célula do **TOTAL LÍQUIDO PATRIMONIAL R$ 100.874.027,60** (tabela 51).
- O corpo tem 52 tabelas e 789 unidades de texto não vazias (parágrafos + linhas de tabela), todas na aba `Texto`.

## 2. Estrutura do documento (em ordem)

| Ordem | Bloco | Onde |
|---|---|---|
| 1 | Cabeçalho da empresa (CNPJ 89.472.195/0001-86, endereço, IE) | parágrafos 1–5 |
| 2 | QUADRO SOCIETÁRIO REAL, DE FATO E DE DIREITO – 7 sócios, capital R$ 10.000,00 | tabela 1 |
| 3 | Tabela **sem título**: Valter Siviero 99% / Elgides Siviero Binda 1%, capital R$ 100.000,00 | tabela 2 |
| 4 | CONSTITUIÇÃO DO NOVO QUADRO SOCIETÁRIO A CONTAR DE 09.03.2020 = NÃO REAL – Pedro Siviero 99% / Marcos Antônio Siviero 1%, capital R$ 300.000,00 | tabela 3 |
| 5 | Descrição dos bens – VEÍCULOS (itens 01–15) – total R$ 1.217.000,00 | tabela 4 |
| 6 | Descrição dos bens – MÁQUINAS (itens 01–14) – total R$ 2.517.000,00 | tabela 5 |
| 7 | LOTES URBANOS EM VERANÓPOLIS e BENTO GONÇALVES (itens 01–08) – total R$ 3.290.000,00 | tabela 6 |
| 8 | ÁREAS URBANAS EM VERANÓPOLIS (itens 01–04) – total R$ 26.929.394,60 | tabela 7 |
| 9 | ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL (itens 01–12, em duas tabelas) – total R$ 7.663.866,00 | tabelas 8–9 |
| 10 | LOTES / LOTEAMENTOS EXECUTADOS E REGISTRADOS (só cabeçalho) seguido de títulos de cidade e 31 tabelas de loteamentos (itens 01–31) | tabelas 10–40 |
| 11 | TOTAL GERAL DOS BENS MÓVEIS E IMÓVEIS E LOTES / LOTEAMENTOS EXECUTADOS E REGISTRADOS = R$ 152.796.260,60 | tabela 41 |
| 12 | ÁREAS DE TERRAS CONTRATADAS – LOTEAMENTOS EM FASES DE LICENCIAMENTO E/OU EXECUÇÃO (itens 01–09; valores como multas e despesas com sinal "(-)") | tabelas 42–43 |
| 13 | TOTAL GERAL DOS BENS – LOTEAMENTOS CONTRATADOS … = R$ 15.754.581,00(-) | tabela 44 |
| 14 | TOTAL GERAL DO PATRIMÔNIO DA EMPRESA = R$ 137.041.679,60 | tabela 45 |
| 15 | LOTES DE TERCEIROS PASSÍVEIS DE INDENIZAÇÃO POR CARACTERIZAÇÃO AMBIENTAL NA CIDADE DE "FARROPILHA" (sic) – 5 loteamentos, 26 lotes | tabelas 46–47 |
| 16 | VALOR DOS LOTES DE TERCEIROS … = R$ 16.240.000,00(-) | tabela 48 |
| 17 | Compromissos a pagar em obras finais – Colina do Retiro II = 3.800.000,00 (-) | tabela 49 |
| 18 | Demais compromissos a pagar … elencados nas OBSERVAÇÕES, itens "1" ao "24" = 16.127.652,00 (-) | tabela 50 |
| 19 | TOTAL LÍQUIDO PATRIMONIAL DA EMPRESA = R$ 100.874.027,60 (realçado em ciano) | tabela 51 |
| 20 | VALOR REAL DE CADA SÓCIO CONFORME A COTA DE CAPITAL – 6 sócios x 16,66% = R$ 16.812.337,93 cada | tabela 52 |
| 21 | OBSERVAÇÕES: lista numerada automaticamente "1.)" a "24.)" (a numeração não aparece no texto extraído; foi reconstruída pelo `w:numPr`), com subitens 22.1, 22.2, 23.1, 23.2 | parágrafos 707–772 |
| 22 | Nota final em caixa alta sobre o balanço contábil; data "Veranópolis – RS, 31 de dezembro de 2024"; 27 nomes para assinatura | parágrafos 773–789 |

Títulos de cidade (parágrafos soltos em caixa alta) que precedem os loteamentos: ANTÔNIO PRADO (item 01), NOVA PRATA (02), PASSO FUNDO (03), VERANÓPOLIS (04–10), "CARLOS BABOSA" (sic; 11–12), NOVA PRATA (13–15), NOVA ARAÇÁ (16), NOVA BASSANO (17), GRAMADO (18), FARROUPILHA (19–26), COTIPORÃ (27), GUAPORÉ (28), CAXIAS DO SUL (29), FLORES DA CUNHA (30), CARLOS BARBOSA (31).

Não existe seção com o título "CONTRATOS | INDENIZAÇÕES | COMPROMISSOS | EMPRÉSTIMOS". O equivalente está espalhado em: tabelas 42–43 (contratos de loteamentos), 46–48 (indenizações), 49–50 (compromissos) e nas 24 Observações (acertos, contingências, empréstimos, financiamentos). A aba `Compromissos` reúne tudo isso (39 linhas).

## 3. Como as partes ambíguas foram lidas

- **Células mescladas**: as tabelas de loteamentos têm colunas mescladas de forma irregular (ex.: tabela 17 tem 6 colunas físicas, tabela 21 e 39 têm 4 ou 5). Cada linha foi reduzida às células físicas distintas (identidade do `w:tc`), o que sempre resultou em `[item, lote, quadra, valor]`. Os asserts do script garantem isso para as 464 linhas de lote.
- **Valor na coluna Quadra**: em San Marino (3 lotes), Firminio Dalzóchio (1), Primavera (2) e Industrial Vicentino (1) a célula Quadra traz "1650 – R$ 244.000,00" e a célula Valor traz "ÁREA DE POSSÍVEL APP". Lido como: quadra = "1650", valor informado = 244.000,00, observação = "ÁREA DE POSSÍVEL APP", **não incluído no subtotal** (os subtotais declarados fecham só sem esses lotes, confirmando a leitura).
- **"ÁREA DE APP"** (Simioni 4 lotes, Lago Negro 7 lotes) e **"FAMÍLIA SIVIERO"** (Salvador/FEMAÇÃ II, 6 lotes da quadra 05): lotes listados sem valor; coluna `incluído no subtotal` = "não (sem valor)".
- **Numeração de item com quebra de linha** ("03 /" nos lotes urbanos; "Item /" no cabeçalho de máquinas): normalizada para o número de dois dígitos.
- **Lotes "O13" (Dalzóchio, quadra 0907) e "O09" (indenização Dalzóchio, quadra 0910)**: mantidos como no documento (letra O no lugar do zero, como já observado no estudo Rafael). Lote "19" da quadra 1469 (indenização Perini) está sem zero à esquerda no documento; mantido "19".
- **Tabela 2 (Valter 99% / Elgides 1%)** não tem título nem explicação no documento; está na aba `Socios` como "Quadro societário sem título (tabela 2)".
- **Área dos loteamentos contratados**: o documento alterna "Área Vendável do Loteamento em lotes" (itens 01–04) e "Área total" (05–09); a aba `Itens` traz o número e a unidade diz qual é.
- **Cidades**: para loteamentos, extraídas da própria descrição ("localizado na cidade de X – RS"); para os demais itens, lidas à mão da descrição (coluna `cidade`). Áreas, matrículas e valores unitários dos 24 imóveis não-loteamento também foram transcritos à mão da descrição (dicionário `ATTR` no script) e conferidos contra o texto.
- **Valores das Observações**: cada item numerado é seguido por um parágrafo solto "R$ x" que resume o valor; esse parágrafo foi tomado como o valor do item. Os 16 valores assim obtidos somam exatamente os R$ 16.127.652,00 da tabela 50, o que confirma a leitura.
- Formato numérico: todo "1.234,56" convertido para float; lote/quadra/matrícula mantidos como texto com zeros à esquerda.

## 4. Totais: CONFERE / DIVERGE

Todos os totais estruturais do documento fecham:

| Total declarado | Valor | Recomputado | Status |
|---|---|---|---|
| Veículos | 1.217.000,00 | 1.217.000,00 | CONFERE |
| Máquinas | 2.517.000,00 | 2.517.000,00 | CONFERE |
| Lotes urbanos | 3.290.000,00 | 3.290.000,00 | CONFERE |
| Áreas urbanas | 26.929.394,60 | 26.929.394,60 | CONFERE |
| Áreas rurais | 7.663.866,00 | 7.663.866,00 | CONFERE |
| 31 subtotais de loteamentos | — | — | todos CONFERE (soma 111.179.000,00) |
| **Total geral dos bens** | **152.796.260,60** | **152.796.260,60** | **CONFERE** |
| Loteamentos contratados (dedução) | 15.754.581,00 | 15.754.581,00 | CONFERE |
| Patrimônio da empresa | 137.041.679,60 | 137.041.679,60 | CONFERE |
| Lotes de terceiros – indenização (dedução) | 16.240.000,00 | 16.240.000,00 | CONFERE |
| Obras Colina do Retiro II (dedução) | 3.800.000,00 | — | declarado, sem composição |
| Demais compromissos (dedução) = soma das Observações | 16.127.652,00 | 16.127.652,00 | CONFERE |
| **Total líquido patrimonial** | **100.874.027,60** | **100.874.027,60** | **CONFERE** |
| Valor por sócio (÷ 6) | 16.812.337,93 | 16.812.337,93 | CONFERE |
| Acerto Rovílio (Obs. 1) | 3.440.082,00 | 3.440.082,00 | CONFERE |
| Quadros societários (tabelas 1, 2, 3) | 10.000 / 100.000 / 300.000 | iguais | CONFERE |

**DIVERGE** (todos na aba `Totais`, com a diferença):

1. **Multa do loteamento contratado 03 – Vista da Colina**: declarada R$ 2.556.900,00; 20% do custo de R$ 12.100.000,00 = R$ 2.420.000,00 (diferença –136.900,00). O documento diz "multa estimada em 20% do custo", mas o valor não corresponde.
2. **Multa do loteamento contratado 04 – Loteamento das Famílias**: declarada R$ 2.280.646,00; 20% de R$ 12.100.000,00 = R$ 2.420.000,00 (diferença +139.354,00). Mesma observação. (As multas dos itens 02, 05, 06, 07, 08 e 09 batem exatamente com 20% do custo, mais R$ 50.000,00 de dano moral nos itens 06–09.)
3. **Item 03 áreas urbanas**: 4.986,00 m² x R$ 992,78 = 4.950.001,08; documento usa 4.950.000,00 (arredondamento de R$ 1,08).
4. **Item 09 áreas rurais**: 15,1 ha x R$ 2.650,00 = 40.015,00; documento usa 40.000,00 (arredondamento de R$ 15,00).
5. **Item 06 áreas rurais (Ansolin, matrícula 30.390)**: as cinco parcelas adquiridas descritas somam 189.135,22 m², mas o documento diz que a área da empresa é 182.435,22 m² (= 202.435,22 da matrícula – 20.000 de Antônio Ansolin). Diferença de 6.700,00 m², igual à parcela de Ivo José Frainer. Os valores (R$ 2.263.666,00 = soma dos preços de aquisição) fecham; a divergência é só de área.
6. **Observação 21 (Paraná Equipamentos)**: 12 x 26.666,67 = 320.000,04; documento usa 320.000,00 (arredondamento).
7. **Tabela 52**: 6 x 16.812.337,93 = 100.874.027,58, contra 100.874.027,60 do total (arredondamento de 0,02). E 6 x 16,66% = 99,96%, não 100%.

Sem subtotal declarado (apenas recomputados): os 5 loteamentos da indenização (400.000; 986.000; 640.000; 664.000; 13.550.000).

## 5. Placeholders e lacunas no documento

- Lotes urbanos 07 e 08 (Bento Gonçalves): "Área de .........m²" e "Matrícula nº ..........." em branco.
- Máquinas 05, 06, 11 e 12: "ano....." em branco.
- Loteamento 04 Fracionamento Simioni: subtotal grafado "R$ 0.000.000,00"; os 4 lotes são "ÁREA DE APP".
- Loteamento contratado 01 Caminhos do Sol (Buriti, 612 lotes, custo 21.163.000,00): valor grafado "R$ 00.000.000,00"; é o único contratado sem multa estimada, embora o texto diga que custos excedentes correm por conta da SEI.
- Loteamentos contratados 08 (Dr. Letti) e 09 (Sol Nascente III): "ESTIMATIVA DE ??? LOTES" e participação "= ???".
- Observação 1: os 3 lotes do Condomínio Morada dos Ipês II a pagar a Rovílio (R$ 550.000,00 cada) estão com "Lote nº .... – Quadra nº ....".
- Observações 22, 23 e 24: todos os valores em branco ("R$ ............000,00 (....... MIL REAIS)").
- Observação 1: "R$ 25.500,00 (Vinte e cinco mil reais)": número e extenso divergem; o número 25.500 é o que fecha a soma de 3.440.082,00.

## 6. Itens sem valor

- 4 lotes Simioni (APP), 7 lotes Lago Negro quadra 11 B (APP), 6 lotes Salvador/FEMAÇÃ II quadra 05 ("FAMÍLIA SIVIERO"): 17 lotes listados sem valor. Mais 7 lotes "ÁREA DE POSSÍVEL APP" com valor informado mas excluído do subtotal (San Marino 012/1650 244.000, 013/1650 233.000, 004/1676 253.000; Dalzóchio 002/1292 131.000; Primavera 008/1193 170.000, 009/1193 165.000; Vicentino 013/1303 1.125.000; soma 2.321.000,00). Total: 438 lotes listados, 414 com valor no subtotal.
- Os 9 loteamentos contratados (tabelas 42–43) não têm valor de ativo; só multas/despesas negativas (8 deles) ou placeholder (1).
- Observações 4, 8, 10, 11, 12 (processo Jardim Itália; impostos 5,93%/7%; passivo de imóveis vendidos; 6% de comissão) e 22–24: sem valor.
- Lote urbano 04 hipotecado à Receita Estadual (dívida de ICMS): valor da dívida não informado.

## 7. Lotes duplicados

Chave de comparação: loteamento (normalizado) + quadra + lote, sobre as 472 linhas da aba `Lotes` (438 dos loteamentos, 26 da indenização, 8 do acerto Rovílio). **Nenhum lote aparece duas vezes.** Em particular:

- Firminio Dalzóchio: item 21 (O13/0907; 006, 007/1040; 003/1046; 001, 002, 004, 005/1182; 002/1292) x indenização 01 (O09/0910; 003, 004/1292) – quadras coincidem (1292) mas lotes não.
- Jardim Itália: item 25 (quadras 1736–1750, incl. 002/003/004 da 1741 e 015/016 da 1747) x indenização 02 (007, 009/1741; 001, 002, 009/1747) – sem coincidência.
- Perini: item 19 (002, 003, 004/1688) x indenização 03 (011, 012, 018, 19/1469) – quadras diferentes.
- Primavera: item 22 (008, 009/1193; 001/1461) x indenização 04 (005, 006, 010, 011/1199) – quadras diferentes.
- Industrial Vicentino: item 24 (007, 008, 009, 010, 013/1303) x indenização 05 (003, 004/1301; 004, 009, 010, 011/1302; 001, 012, 014, 016/1303) – mesma quadra 1303, lotes distintos.
- Morada do Ipê I (item 03: O14/2) x lotes pagos a Rovílio no "Morada dos Ipês I" (13/02, 10/02, 07/04) – distintos (note a grafia diferente do condomínio).
- Nossa Senhora da Saúde: item 28 (006/260; 008/261; 001, 003/269) x Rovílio (12/237; 15/252) – distintos.
- Monte Vêneto: item 11 (Carlos Barbosa, 031/D) e item 27 (Monte Vêneto I–II, Cotiporã, 60 lotes) são loteamentos homônimos em cidades diferentes; não há sobreposição.

Uma verificação mais frouxa (cidade + quadra + lote, ignorando o nome do loteamento) só acusa coincidências de numeração entre loteamentos diferentes de Veranópolis (ex.: Q01 L004 existe em Solar Norte, Salvador, Colina do Retiro II e Pedra Lisa), o que é normal e não indica duplicidade.

## 8. Outras coisas estranhas

- **Data-base x conteúdo**: o documento é datado de 31/12/2024, mas registra fatos de 2025: venda das quotas de Rovílio em 28/08/2025; máquina 14 (D6N XL) "adquirida em setembro de 2025"; saldo do Banco Caterpillar em 31/07/2025; saldo da Paraná Equipamentos em 31/10/2025; arquivo criado em 29/10/2025.
- **Trator D6N XL**: máquinas item 14 diz "ano 2014"; Observação 21 diz "ano 2021" para a mesma máquina (R$ 450.000,00, Paraná Equipamentos).
- **Quadro societário**: a tabela 1 tem 7 sócios (Rovílio com 2%), mas a distribuição (tabela 52) é em 6 partes de 16,66% (Irineu, Raul, Valter, Elcio, Adélia, Pedro), coerente com a venda das quotas de Rovílio. Os percentuais "reais" (18,18% / 7,10%) não são usados na distribuição; cada sócio recebe 1/6, inclusive Pedro Siviero (7,10% na tabela 1). Mesma premissa já apontada como F3 na conferência anterior.
- **Item 02 das áreas rurais (Família Lago, matrícula 33.516)** é descrito como "terras URBANAS … IMÓVEL URBANO" dentro do bloco de áreas rurais (equivalente à observação O6 da conferência anterior).
- **Item 10 das áreas rurais (2,0 ha Cecconello, R$ 600.000,00)** é o mesmo imóvel do loteamento contratado 05 (Família Cecconello), que por sua vez traz multa de R$ 1.300.000,00(-).
- **Imóvel Oleoplan (áreas urbanas 04)**: aqui o valor é R$ 13.631.378,60, igual à soma das parcelas; o estudo Rafael trazia R$ 10.631.378,60 (observação O2 da conferência anterior). O inventário 11 resolve a diferença de R$ 3.000.000,00 a favor da soma das parcelas.
- **Loteamentos contratados 02 e 03** têm a mesma "área vendável" (159.307,14 m²) apesar de estimativas de lotes diferentes (356 x 242); **03 e 04** têm a mesma estimativa (242 lotes) e o mesmo custo (R$ 12.100.000,00) com áreas diferentes. Possível copia-e-cola.
- **Salvador – Alberto Salvador – FEMAÇÃ II**: descrição diz "num total de 26 lotes", mas a tabela lista 32 (26 com valor + 6 "FAMÍLIA SIVIERO").
- Descrições com "aproximadamente": Barbiero (4), Altos da Julieta (10), Monte Vêneto I–II (60), Família Falavigna (30); as tabelas listam exatamente esses números.
- **Observação 14**: a SEI assume R$ 150.000,00 (1 lote) por placas solares instaladas no **Hotel Araçá Ltda.** (outra empresa do grupo).
- **Observação 15**: comissão de ~R$ 1.200.000,00 a Alexandre Susin pela intermediação dos contratos Buriti, mais 6 lotes a entregar (não identificados).
- **Observação 13**: 3 lotes (R$ 600.000,00) a Celestino Rossi, não identificados.
- **Observação 9** estima R$ 250.000,00 para regularizar áreas rurais compradas por contrato (itens 11 e 12, Piccolo e Scalco, não registradas em nome da empresa).
- Erros de digitação: "CARLOS BABOSA", "FARROPILHA", "MERCDES BENZ", placas "IYJ 8I14" e "NPQ OI67" (letra I/O no lugar de dígito), "Condomínio Morada do Ipê I" x "Morada dos Ipês I".
- A nota final em caixa alta afirma que o balanço contábil "NÃO CORRESPONDE À REALIDADE FÁTICA DA EMPRESA".

## 9. Contagens das abas

| Aba | Linhas | Conteúdo |
|---|---|---|
| Itens | 93 | 15 veículos + 14 máquinas + 8 lotes urbanos + 4 áreas urbanas + 12 áreas rurais + 31 loteamentos + 9 loteamentos contratados |
| Lotes | 472 | 438 lotes dos 31 loteamentos + 26 lotes de terceiros (indenização) + 8 lotes do acerto Rovílio |
| Compromissos | 39 | 9 contratos + 5 indenizações + 1 obra (tabela 49) + 24 observações |
| Passivos | 42 | 8 multas + 3 despesas + 1 placeholder (contratados); indenização; obras; demais compromissos; 24 observações; 3 ônus sem valor |
| Socios | 21 | tabelas 1, 2, 3 e 52, incluindo linhas de total |
| Totais | 78 | totais declarados, subtotais, verificações aritméticas por item |
| Texto | 789 | texto integral em ordem, com origem (parágrafo / tabela N linha M) e bloco |


**Nota de 28/10/2026 (fechamento da Parte 3):** sobre o imóvel Oleoplan, a decisão aprovada é adotar R$ 10.631.378,60 (Rafael, C0; pacote 2, item 1). A frase acima, de que o inventário "resolve a diferença a favor da soma das parcelas", descreve o documento do grupo e não a posição adotada.
