# Inventário 31.12.2024 – comparação "versão antiga" x "versão atual" (2025)

Data da análise: 28/10/2026.

## 1. Método

- Pastas comparadas:
  - ANTIGA: `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\` (arquivos "ANTIGO 00." a "ANTIGO 11."; não existe "ANTIGO 02").
  - ATUAL: `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\` (arquivos "00." a "11." + "HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx" + "ACERTOS E ACORDOS.pdf").
- Extração com Python + python-docx: todos os parágrafos e todas as células de todas as tabelas, na ordem do documento (incluindo tabelas aninhadas). Células mescladas foram lidas uma única vez.
- Comparação linha a linha com `difflib` (antigo x atual); nas linhas longas foi feita comparação palavra a palavra para localizar o trecho alterado. Em seguida cada diferença foi interpretada (tabela, item, valor antigo, valor novo).
- Alterações rastreadas (w:ins / w:del) e comentários (w:comment): verificados no XML de todos os 23 arquivos .docx. Resultado: **nenhum arquivo contém alterações rastreadas nem comentários**; todas as diferenças abaixo são diferenças de texto "aceito".
- Valores reproduzidos exatamente como constam nos documentos (formato brasileiro). Onde o documento traz valor em branco, pontilhado ou "0.000.000,00", isto é dito explicitamente. Nenhum número foi calculado por mim, salvo onde indicado como "conferência aritmética".
- "Linha do Valter": linha "VALTER SIVIERO / BENTA ELISA VELHO SIVIERO" na tabela de rateio final de cada empresa (cota de 1/6 = 16,66%).
- Scripts e dumps de texto em `C:\Users\User\AppData\Local\Temp\claude\C--Users-User--claude-spares-spare-kOyywN\eb6693c8-36bc-49ba-865b-2dcc272284c1\scratchpad\diff2025\`. Nenhum arquivo de origem foi modificado.

Contagem de linhas alteradas (saída do difflib, somando linhas removidas + adicionadas):

| Doc | Empresa | Linhas alteradas | Situação |
|---|---|---|---|
| 00 | Planilha do valor do capital | 32 | alterado |
| 01 | Abastecedora de Combustíveis Siviero | 21 | alterado |
| 02 | Agropecuária Fazenda São Crispin | – | só existe na versão atual |
| 03 | Associação Cultural Princesa dos Vales | 0 | sem diferenças |
| 04 | Ceval Terraplenagem | 0 | sem diferenças |
| 05 | Hotel Princesa dos Vales | 22 | alterado |
| 06 | Hotel Araçá | 18 | alterado |
| 07 | Rádio Comunidade Princesa dos Vales | 0 | sem diferenças |
| 08 | Transportadora Irmãos Siviero | 20 | alterado |
| 09 | Turismo Siviero | 20 | alterado |
| 10 | Siviero Transportes | 20 | alterado |
| 11 | Siviero Empreendimentos Imobiliários (SEI) | 564 | fortemente alterado |

---

## 2. Documento 00 – PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 2, linha 01 Abastecedora | "01. ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA." / R$ 9.661.414,00 | "01. ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. **x2**" / R$ 9.661.414,00 | Valor igual; acrescentada a marca "x2" ao nome (também nas linhas 08 e 11 – são as três empresas em que Rovílio tinha 2%). |
| Tabela 2, linha 05 Hotel Princesa dos Vales | R$ 83.640.801,00 | R$ 75.410.201,00 | Redução de R$ 8.230.600,00 (reavaliação da Fazenda São Crispin de R$ 130.000/ha para R$ 110.000/ha – ver doc 05). |
| Tabela 2, linha 06 Hotel Araçá | R$ 61.790.460,00 | R$ 58.076.823,20 | Redução de R$ 3.713.636,80 (obra civil de 25% para 15% – ver doc 06). |
| Tabela 2, linha 08 Transportadora | "08. TRANSPORTADORA IRMÃOS SIVIERO LTDA." | "... LTDA. **x2**" | Valor igual (R$ 16.374.986,00). |
| Tabela 2, linha 11 SEI | nome sem valor (célula vazia) | "11. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. x2" / "R$ 100.874,027,60" | Valor preenchido. Atenção: grafado com vírgula no lugar do ponto ("100.874,027,60"); na tabela 3 e no doc 11 consta R$ 100.874.027,60. |
| Tabela 2, linha de total | "PARCIAL / TOTAL" R$ 174.480.661,00 | "TOTAL" R$ 263.410.451,80 | Antes era total parcial (sem SEI). |
| Parágrafo novo | – | "CAPITAL PATRIMONIAL PARA CADA SÓCIO DO GRUPO EMPRESARIAL SIVIERO NO VALOR DE R$ 43.897.245,30" | Conferência aritmética: (263.410.451,80 − 26.980,00) / 6 = 43.897.245,30. |
| Parágrafo novo | – | "CAPITAL PATRIMONIAL PARA O SÓCIO ANTÔNIO TELMO DE TONI (RÁDIO) NO VALOR DE R$ 26.980,00" | 1% da Rádio. |
| Tabela 3 (nova) | não existia | Tabela empresa x valor com a "COTA DO CAPITAL PATRIMONIAL ... PARA CADA SÓCIO: 16,66%" por empresa: 01 R$ 1.610.235,67; 02 R$ 0,00; 03 R$ 0,00; 04 R$ 52.500,00; 05 R$ 12.568.366,83; 06 R$ 9.679.470,53; 07 R$ 445.170,00 (16,50%) + Telmo de Toni 1% R$ 26.980,00; 08 R$ 2.729.164,33; 09 R$ 0,00; 10 R$ 0,00; 11 R$ 16.812.337,93; TOTAL R$ 263.410.451,80 | Na linha 11 o nome está grafado "SIVIERO EMPRENTOS IMOBILIÁRIOS LTDA." (erro de digitação). |
| Parágrafo novo | – | Texto "Conforme INSTRUMENTO PARTICULAR DE COMPRA E VENDA DE QUOTAS DE CAPITAL ... datado de 28 de agosto de 2025, o Sócio ROVÍLIO SIVIERO ... efetuaram a VENDA das quotas que lhes pertenciam às empresas, diluídas em partes iguais aos demais sócios remanescentes." | Registra a saída de Rovílio. |
| Linha de assinaturas | "Rovílio Siviero Antônio Telmo de Toni Pedro Elias Siviero Marcos Antônio Siviero" | "Antônio Telmo de Toni Pedro Elias Siviero Marcos Antônio Siviero" | Rovílio removido das assinaturas. |

Total declarado: antigo R$ 174.480.661,00 (parcial, sem SEI) → atual R$ 263.410.451,80. Linha "Valter": não há linha por sócio neste documento; a cota por sócio passou a constar como R$ 43.897.245,30.

---

## 3. Documento 01 – ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA.

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 3, item 01 (posto, Matrícula 3.752), célula de valor | "R$ 9.661.414,00" | "R$" (número apagado, célula contém apenas "R$   ") | Descrição do item idêntica. O valor do item sumiu da célula, mas o TOTAL GERAL DA EMPRESA continua R$ 9.661.414,00. Provável apagamento acidental. |
| Tabela 4 (rateio), linhas Irineu, Raul, Valter, Elcio, Adélia, Pedro | 16,33% = R$ 1.578.030,96 (cada) | 16,66% = R$ 1.610.235,67 (cada) | Redistribuição após saída de Rovílio. |
| Tabela 4, linha Rovílio | "ROVILIO SIVIERO / 225.676.150-91 / 2,00% = R$ 193.228,24" | linha removida | Tabela passou de 9 para 8 linhas. |
| Tabela 4, TOTAL | 100,00% = R$ 9.661.414,00 | 100,00% = R$ 9.661.414,00 | Igual. |
| Parágrafo novo | – | Texto do INSTRUMENTO PARTICULAR DE COMPRA E VENDA DE QUOTAS de 28/08/2025 (Rovílio) | |
| Assinaturas | "Rovílio Siviero" (última linha) | removido | |

Total declarado: R$ 9.661.414,00 em ambas. Linha Valter: 16,33% = R$ 1.578.030,96 → 16,66% = R$ 1.610.235,67.

---

## 4. Documento 02 – AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA.

Sem contrapartida na pasta antiga (não existe "ANTIGO 02"). Não há o que comparar. Na versão atual: capital R$ 300.000,00 (6 sócios x R$ 50.000,00 = 16,66%); rateio final R$ 0,00 para todos (linha Valter: R$ 0,00 = 16,66%).

---

## 5. Documento 03 – ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES

Sem diferenças (texto e tabelas idênticos). Total geral da Associação: R$ 0,00. Linha Valter: 16,66% = R$ 0,00 em ambas.

---

## 6. Documento 04 – CEVAL TERRAPLENAGEM E ESCAVAÇÕES LTDA.

Sem diferenças. Total geral da empresa: R$ 315.000,00. Linha Valter: 16,66% = R$ 52.500,00 em ambas.

---

## 7. Documento 05 – HOTEL PRINCESA DOS VALES LTDA.

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 5, item 05 – Fazenda São Crispin (491 ha; vendida 83,47 ha à Família Rigo; saldo 407,53 ha) – preço por hectare | R$ 130.000,00/ha em todas as parcelas | R$ 110.000,00/ha em todas as parcelas | |
| idem – Área Dona Rita produtiva (100 ha) | R$ 13.000.000,00 | R$ 11.000.000,00 | |
| idem – Área Dona Rita não produtiva (186 ha) | R$ 24.180.000,00 | R$ 20.460.000,00 | |
| idem – Área Zancheta produtiva (45,6 ha) | R$ 5.928.000,00 | R$ 5.016.000,00 | |
| idem – Área Zancheta não produtiva (11,4 ha) | R$ 1.482.000,00 | R$ 1.254.000,00 | |
| idem – Área Jane/João Paulo (64,53 ha) | R$ 8.388.900,00 | R$ 7.098.300,00 | |
| idem – valor do item | R$ 52.978.900,00 | R$ 44.828.300,00 | Redução de R$ 8.150.600,00. |
| Tabela 5, item 06 – 4,0 ha Dirceu Prestes/Lairton (Matrícula 11.090) | 4,0 ha x R$ 130.000,00 = R$ 520.000,00 | 4,0 ha x R$ 110.000,00 = R$ 440.000,00 | Redução de R$ 80.000,00. |
| Tabela 5, item 07 – Lote Rural 52, Cotiporã (Matrícula 7.391, 15,1 ha) | "adquirida da Senhora GLORIA E GERALDO BASSANI" / R$ 422.800,00 | "adquirida da Senhora GLORINHA E GERALDO BASSANI" / R$ 422.800,00 | Só o nome mudou; valor igual. Observação: este mesmo imóvel (Matrícula 7.391) constava também no doc 11 antigo (tabela 9, item 10, R$ 226.500,00) e foi removido do doc 11 atual – eliminada a duplicidade. |
| Tabela 6, TOTAL GERAL DA EMPRESA | R$ 83.640.801,00 | R$ 75.410.201,00 | Redução de R$ 8.230.600,00. |
| Tabela 7 (rateio), cada um dos 6 sócios | 16,66% = R$ 13.940.133,50 | 16,66% = R$ 12.568.366,83 | |
| Tabela 7, TOTAL | 100,00% = R$ 83.640.801,00 | 100,00% = R$ 75.410.201,00 | |

Total declarado: R$ 83.640.801,00 → R$ 75.410.201,00. Linha Valter: R$ 13.940.133,50 → R$ 12.568.366,83.

---

## 8. Documento 06 – HOTEL ARAÇÁ LTDA.

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 3, item 01 (Hotel, Quadra 37, 18 terrenos, 6.900 m²) – "Valor do Edifício Construído – obra civil = 04 andares" | "25,00% (Vinte e cinco por cento) da obra feita: R$ 37.136.368,00 ... = R$ 9.284.092,00" | "15,00% (Quinze por cento) da obra feita: R$ 37.136.368,00 ... = R$ 5.570.455,20" | Percentual de "vantagem da coisa feita" reduzido. |
| Tabela 3, item 01 – valor do item | R$ 53.320.460,00 | R$ 49.606.823,20 | Redução de R$ 3.713.636,80. |
| Tabela 4, TOTAL GERAL DA EMPRESA | R$ 61.790.460,00 | R$ 58.076.823,20 | |
| Tabela 5 (rateio), cada um dos 6 sócios | 16,66% = R$ 10.298.410,00 | 16,66% = R$ 9.679.470,53 | |
| Tabela 5, TOTAL | R$ 61.790.460,00 | R$ 58.076.823,20 | |

Total declarado: R$ 61.790.460,00 → R$ 58.076.823,20. Linha Valter: R$ 10.298.410,00 → R$ 9.679.470,53.

---

## 9. Documento 07 – RÁDIO COMUNIDADE PRINCESA DOS VALES FM LTDA.

Sem diferenças. Total geral: R$ 2.698.000,00. Linha Valter: 16,50% = R$ 445.170,00 em ambas.

---

## 10. Documento 08 – TRANSPORTADORA IRMÃOS SIVIERO LTDA.

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 5 (rateio), 6 sócios | 16,33% = R$ 2.674.581,05 (cada) | 16,66% = R$ 2.729.164,33 (cada) | |
| Tabela 5, linha Rovílio | "ROVILIO SIVIERO / 2,00% = R$ 327.499,70" | removida | 9 → 8 linhas. |
| Tabela 5, TOTAL | R$ 16.374.986,00 | R$ 16.374.986,00 | Igual. |
| Parágrafo novo | – | Instrumento de compra e venda de quotas de 28/08/2025 (Rovílio) | |
| Assinaturas | "Élcio Siviero Pedro Siviero Rovílio Siviero" | "Élcio Siviero Pedro Siviero" | |

Total declarado: R$ 16.374.986,00 em ambas. Linha Valter: R$ 2.674.581,05 → R$ 2.729.164,33.

---

## 11. Documento 09 – TURISMO SIVIERO LTDA.

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 4 (rateio), 6 sócios | 16,50% = R$ 0,00 | 16,66% = R$ 0,00 | |
| Tabela 4, linha Rovílio | "ROVILIO SIVIERO / 1,00% = R$ 0,00" | removida | |
| Parágrafo novo | – | Instrumento de compra e venda de quotas de 28/08/2025 | |
| Assinaturas | "Élcio Siviero Pedro Siviero Rovílio Siviero Geni Siviero Mattiuz" | "Élcio Siviero Pedro Siviero Geni Siviero Mattiuz" | |

Total declarado: R$ 0,00 em ambas. Linha Valter: 16,50% → 16,66% (R$ 0,00).

---

## 12. Documento 10 – SIVIERO TRANSPORTES LTDA.

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 4 (rateio), 6 sócios | 16,50% = R$ 0,00 | 16,66% = R$ 0,00 | |
| Tabela 4, linha Rovílio | "ROVILIO SIVIERO / 1,00% = R$ 0,00" | removida | |
| Parágrafo novo | – | Instrumento de compra e venda de quotas de 28/08/2025 | |
| Assinaturas | "Élcio Siviero Pedro Siviero Rovílio Siviero" | "Élcio Siviero Pedro Siviero" | |

Total declarado: R$ 0,00 em ambas. Linha Valter: 16,50% → 16,66% (R$ 0,00).

---

## 13. Documento 11 – SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. (SEI)

Versão antiga: 761 linhas extraídas, 47 tabelas. Versão atual: 841 linhas, 52 tabelas. 564 linhas alteradas. A versão antiga tinha praticamente todos os totais em branco ("R$ 0.000.000,00") e os 91 lotes do Colina do Retiro II sem valor; a versão atual preenche todos os totais e acrescenta seções de passivo.

### 13.1 Cabeçalho e quadro societário (tabelas 1–3)

| Campo / item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Título antes da tabela 2 | "CONSTITUIÇÃO DO NOVO QUADRO SOCIETÁRIO A CONTAR DE 01.12.2003 = NÃO REAL" | título removido (a tabela 2 continua igual) | Só o título sumiu. |
| Tabelas 1, 2, 3 | – | – | Sem diferenças (capital R$ 10.000,00 / R$ 100.000,00 / R$ 300.000,00). |

### 13.2 Tabela 4 – VEÍCULOS

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| 05 – caminhão SCANIA placa IIQ 0751, 1963 | R$ 28.000,00 | R$ 25.000,00 | |
| 10 – caminhão SCÂNIA placa IHK 2435, 1967 | "R$ 0.000,00" (em branco) | R$ 25.000,00 | |
| 11 | "Um semirreboque ARTESANAL = PLACA IIH 6326, cor branca, ano 1986." R$ 20.000,00 | "Uma camionete TOYOTA HILUX, cor preta, ano 2018 = PLACA IYJ 8I14." R$ 188.000,00 | Bem substituído. |
| 12 – caminhão SCÂNIA placa IIX 8288, 1963 | R$ 30.000,00 | R$ 25.000,00 | |
| 13 – FORD 1929 "FUBICA", placa IKR 3B22 | R$ 40.000,00 | R$ 60.000,00 | |
| 15 | "Um veículo JEEP COMPASS = PLACA GHX 9I84, cor cinza, ano 2022." R$ 150.000,00 | "Um veículo TOYOTA RAV4 = PLACA IZQ 9H36, cor prata, ano 2019." R$ 155.000,00 | Bem substituído. |
| TOTAL GERAL DOS BENS – VEÍCULOS | "R$ 000.000,00" (em branco) | R$ 1.217.000,00 | Itens 01–04, 06–09 e 14 sem alteração. |

### 13.3 Tabela 5 – MÁQUINAS

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| 02 – Escavadeira 320DL 2014 | R$ 400.000,00 | R$ 350.000,00 | |
| 05 – Trator de Esteira D6, ano "....." | R$ 100.000,00 | R$ 80.000,00 | |
| 06 – Trator de Esteira D6, ano "....." | R$ 100.000,00 | R$ 50.000,00 | |
| 14 (novo) | – | "Um Trator de Esteira, ano 2014, modelo D6N XL, CATERPILLAR, adquirido em setembro de 2025." R$ 450.000,00 | Nas Observações o mesmo trator aparece como "ano 2021" (Paraná Equipamentos) – inconsistência de ano entre tabela e observação. |
| TOTAL GERAL DOS BENS – MÁQUINAS | "R$ 0.000.000,00" | R$ 2.517.000,00 | Tabela ganhou uma 4ª coluna vazia. Demais itens iguais. |

### 13.4 Tabela 6 – LOTES URBANOS EM VERANÓPOLIS E BENTO GONÇALVES

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| 01 – Av. Pinheiro Machado, "Casa da Nina" (Matrícula 1.597) | R$ 150.000,00 | R$ 280.000,00 | |
| 02 – Av. Dr. José Montaury, "Casa da Nona" (Matrícula 1.335) | R$ 220.000,00 | R$ 250.000,00 | |
| 05 – Av. Osvaldo Aranha esq. Rua Eduardo Duarte, 345 m² (Matrícula 1.507) | R$ 850.000,00 | R$ 550.000,00 | |
| 06 (antigo) – Rua Bento Gonçalves, fundos da Casa da Nona, adquirido de Michelle Tedesco, 375 m² (Matrícula 2325) | R$ 220.000,00 | **item removido** | Imóvel não consta mais. |
| 07 antigo → 06 atual – Rua Gen. Flores da Cunha, Família Dal Ponte (Matrícula 3.788) | R$ 1.010.000,00 | R$ 1.010.000,00 | Só renumerado. |
| 08 antigo → 07 atual – Bento Gonçalves, Loteamento Ouro Verde II, lote 569/46 | R$ 159.000,00 | R$ 130.000,00 | |
| 09 antigo → 08 atual – Bento Gonçalves, Loteamento Verona, lote 188/62 | R$ 100.000,00 | R$ 160.000,00 | |
| TOTAL | "R$ 0.000.000,00" | R$ 3.290.000,00 | Itens 03 e 04 iguais. |

### 13.5 Tabela 7 – ÁREAS URBANAS EM VERANÓPOLIS

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| 01 – Lote 1054, 1,6 ha, Moacir Zanini (Matrícula 17.038) | R$ 900.000,00 | R$ 1.200.000,00 | |
| 02 (antigo) – Lote 1543, 1,9 ha, BR 470 esq. Linha República, Família Bissani (Matrícula 21.819), "Adquirida por R$ 400.000,00" | R$ 900.000,00 | **item removido** | |
| 03 antigo → 02 atual – Matrícula 3.751, 7.200 m², antena da Rádio | R$ 7.148.016,00 | R$ 7.148.016,00 | Renumerado. |
| 04 antigo → 03 atual – Matrícula 3.754, 4.986 m², ex-Serraria | R$ 4.950.000,00 | R$ 4.950.000,00 | Renumerado. |
| 05 antigo → 04 atual – pavilhões Oleoplan | Dois terrenos: Matrícula 17.409 (1.942,03 m²) = R$ 3.884.060,00 e Matrícula 6.136 (2.850,00 m²) = R$ 5.700.000,00; pavilhões 4.170 m² R$ 3.113.322,00; 30% = R$ 933.996,60 | "Um imóvel urbano, com a área de 4.792,03m² ... Matrícula nº 29.954 – (oriunda da unificação das Matrículas nº 17.409 e 6.136) ... R$ 2.000,00/m² = R$ 9.584.060,00"; pavilhões e 30% iguais | Matrículas unificadas; valor do item igual: R$ 13.631.378,60. |
| TOTAL | "R$ 0.000.000,00" | R$ 26.929.394,60 | |

### 13.6 Tabelas 8 e 9 – ÁREAS RURAIS (Veranópolis, Cotiporã, Fagundes Varela, Caxias do Sul)

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| T8 item 02 – Família Lago, Linha Tomaz Flores (perto do CTG Rincão da Roça Reúna) | "área de terras rurais, denominada de PARTE DO LOTE RURAL LETRA 'M', com área total de 3,6 hectares, conforme contrato e área medida conforme croqui, com 33.742,78m² ... Matrícula nº 14.319 = 3,6 hectares." R$ 900.000,00 | "área de terras URBANAS, denominada de UM IMÓVEL URBANO, com a denominação administrativa de lote nº 0292, da quadra nº 009, zona 12, distrito 01, com área total de 4,0 hectares, conforme Registro e área medida, com 40.027,41m², ... Estrada Velha da Vacaria ... Matrícula nº.: 33.516." R$ 1.500.000,00 | Imóvel regularizado/registrado como urbano, nova matrícula, área maior e valor maior. |
| T8 item 03 – Lote Rural 14, Santa Rita, 7,5 ha | R$ 280.000,00 | R$ 220.000,00 | |
| T8 item 04 – Lote Rural 12, Família Sangali, 13,0 ha | "13,0 hectares X 11.000,00" R$ 143.000,00 | "13,0 hectares X 40.000,00" R$ 520.000,00 | |
| T8 item 07 – Lotes 26 e 27, chácaras, Barra do Jabuticaba (10 chácaras) | "R$ 10.000,00 cada chácara" R$ 100.000,00 | "R$ 2.000,00 cada chácara" R$ 20.000,00 | |
| T9 item 09 – Lote 33, Cotiporã, Armelindo Calza (Matrícula 1.218), 15,1 ha | "15,1 hectares X R$ 5.000,00" R$ 75.500,00 | "área de terras rurais de MATO NATIVO ... 15,1 hectares X R$ 2.650,00" R$ 40.000,00 | |
| T9 item 10 (antigo) – Lote Rural 52, Cotiporã, Gloria e Geraldo Bassani (Matrícula 7.391), 15,1 ha x R$ 15.000,00 | R$ 226.500,00 | **item removido** | O mesmo imóvel consta no doc 05 (Hotel Princesa dos Vales, item 07, R$ 422.800,00). Eliminada a duplicidade entre empresas. |
| T9 11 antigo → 10 atual – 2,0 ha Família Cecconello, Caxias do Sul | R$ 600.000,00 | R$ 600.000,00 | Renumerado. |
| T9 12 antigo → 11 atual – Lote 47, Família Piccolo, 7,5 ha | R$ 45.000,00 | R$ 20.000,00; acrescentado "Adquirida pelo valor de R$ 45.000,00. Parte da área foi destinada para compensação ambiental do Loteamento Colina do Retiro II." | |
| T9 13 antigo → 12 atual – Lote 47, Família Scalco, 7,5 ha | R$ 50.000,00 | R$ 30.000,00; mesmo acréscimo sobre compensação ambiental | |
| TOTAL áreas rurais | "R$ 0.000.000,00" | R$ 7.663.866,00 | Itens 01, 05, 06, 08 iguais. |

### 13.7 Tabelas 11–40 – LOTES EM LOTEAMENTOS EXECUTADOS E REGISTRADOS

Subtotais por loteamento (item / loteamento / antigo / atual):

| Item | Loteamento | Antigo | Atual | Comentário |
|---|---|---|---|---|
| 01 | Sol Nascente I – Antônio Prado (11 lotes) | R$ 4.031.000,00 | R$ 4.031.000,00 | igual |
| 02 | Boito – Nova Prata (1 lote) | R$ 136.000,00 | R$ 136.000,00 | igual; inserido o título de cidade "NOVA PRATA" antes da tabela (antes estava sob "ANTÔNIO PRADO") |
| 03 | Condomínio Morada do Ipê I – Passo Fundo | "saldo de 02 lotes": lote O01 quadra 2 R$ 538.000,00 + lote O14 quadra 2 R$ 629.000,00 = R$ 1.167.000,00 | "saldo de 01 lote": lote O14 quadra 2 R$ 629.000,00 = R$ 629.000,00 | **Lote O01/quadra 2 (R$ 538.000,00) saiu** (presumivelmente usado no acerto com Rovílio – ver 13.12). |
| 04 | Fracionamento Simioni – Veranópolis (4 lotes) | R$ 0.000.000,00 | R$ 0.000.000,00 | igual (sem valor em ambas) |
| 05 | Solar Norte I e II (6 lotes) | R$ 1.171.000,00 | R$ 1.171.000,00 | igual |
| 06 | Aditivo Medianeira (15 lotes) | R$ 2.647.000,00 | R$ 2.647.000,00 | igual |
| 07 | Colina do Retiro I (13 lotes) | R$ 2.585.000,00 | R$ 2.585.000,00 | igual |
| 08 | (tabela 17, 2ª parte) | R$ 8.374.000,00 | R$ 8.374.000,00 | igual |
| 09 | **Colina do Retiro II – Veranópolis** | "num total de 91 lotes"; 91 lotes listados **sem nenhum valor**; subtotal "R$ 0.000.000,00" | "num total de 102 lotes"; 102 lotes **todos valorados** (de R$ 162.000,00 a R$ 310.000,00); subtotal **R$ 20.181.000,00** | Maior alteração de valor do documento. Lotes removidos (quadra/lote): 05/015, 07/001, 08/012, 08/015. Lotes acrescentados: 06/001, 07/008, 07/011, 08/009, 08/010, 09/006, 09/007, 11/003, 13/003, 13/004, 14/011, 14/012, 17/001, 17/002, 17/003. |
| 10 | Industrial Pedra Lisa (31 lotes) | R$ 10.985.000,00 | R$ 10.985.000,00 | igual |
| 11 | Monte Vêneto – Carlos Barbosa (1 lote) | R$ 241.000,00 | R$ 241.000,00 | igual |
| 12 | Baldasso – Carlos Barbosa (5 lotes) | R$ 1.419.000,00 | R$ 1.419.000,00 | igual |
| 13 | Colina do Basalto – Nova Prata (14) | R$ 2.180.000,00 | R$ 2.180.000,00 | igual |
| 14 | Carnevalli (3) | R$ 617.000,00 | R$ 617.000,00 | igual |
| 15 | Barbiero (≈4) | R$ 822.000,00 | R$ 822.000,00 | igual |
| 16 | Araçá – Nova Araçá (1) | R$ 142.000,00 | R$ 142.000,00 | igual |
| 17 | Dalla Costa – Nova Bassano (4) | R$ 650.000,00 | R$ 650.000,00 | igual |
| 18 | **Lago Negro – Gramado** | "saldo de 11 lotes"; lotes 001–005 da quadra 01 **sem valor**; lote 014 da quadra 05 sem valor; lotes 015–021 da quadra 11 B = "ÁREA DE APP"; subtotal "R$ 0.000.000,00" | "saldo de 12 lotes"; lotes quadra 01: 001 R$ 770.000,00; 002 R$ 805.000,00; 003 R$ 879.000,00; 004 R$ 720.000,00; 005 R$ 847.000,00; **lote 014/quadra 05 removido**; 7 lotes APP mantidos sem valor; subtotal **R$ 4.021.000,00** | |
| 19 | Perini – Farroupilha (3) | R$ 4.560.000,00 | R$ 4.560.000,00 | igual |
| 20 | **San Marino – Farroupilha** (11 lotes) | lotes 012/1650 R$ 244.000,00; 013/1650 R$ 233.000,00; 004/1676 R$ 253.000,00 somados; subtotal R$ 2.810.000,00 | os mesmos três lotes marcados "ÁREA DE POSSÍVEL APP" (valor movido para a coluna de descrição) e excluídos da soma; subtotal **R$ 2.080.000,00** | Redução de R$ 730.000,00. |
| 21 | **Firminio Dalzóchio – Farroupilha** (9 lotes) | lote 002/1292 R$ 131.000,00 somado; subtotal R$ 1.921.000,00 | lote 002/1292 "ÁREA DE POSSÍVEL APP", excluído; subtotal **R$ 1.790.000,00** | Redução de R$ 131.000,00. |
| 22 | **Primavera – Farroupilha** (3 lotes) | lotes 008/1193 R$ 170.000,00 e 009/1193 R$ 165.000,00 somados; subtotal R$ 483.000,00 | ambos "ÁREA DE POSSÍVEL APP", excluídos; subtotal **R$ 148.000,00** | Redução de R$ 335.000,00. |
| 23 | Vila Vicenza (1) | R$ 157.000,00 | R$ 157.000,00 | igual |
| 24 | **Industrial Vicentino – Farroupilha** (5 lotes) | lote 013/1303 R$ 1.125.000,00 somado; subtotal R$ 6.545.000,00 | lote 013/1303 "ÁREA DE POSSÍVEL APP", excluído; subtotal **R$ 5.420.000,00** | Redução de R$ 1.125.000,00. |
| 25 | Jardim Itália I-II-III (27) | R$ 5.823.000,00 | R$ 5.823.000,00 | igual |
| 26 | Altos da Julieta I-II (≈10) | R$ 2.086.000,00 | R$ 2.086.000,00 | igual |
| 27 | Monte Vêneto I-II – Cotiporã (≈60) | R$ 5.336.000,00 | R$ 5.336.000,00 | igual |
| 28 | Nossa Senhora da Saúde – Guaporé (4) | R$ 533.000,00 | R$ 533.000,00 | igual |
| 29 | Alto das Graças I – Caxias do Sul (14) | R$ 3.699.000,00 | R$ 3.699.000,00 | igual |
| 30 | Família Falavigna – Flores da Cunha (≈30) | R$ 18.216.000,00 | R$ 18.216.000,00 | igual |
| 31 | **Reserva Paradiso – Carlos Barbosa** (novo) | não existia | "num saldo de 01 lote": lote 623, quadra 104, R$ 500.000,00; subtotal R$ 500.000,00 | Nova tabela 40 com título "CARLOS BARBOSA". |

### 13.8 Totais gerais

| Campo | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 40 antiga "TOTAL GERAL DOS BENS – LOTES / LOTEAMENTOS EXECUTADOS E REGISTRADOS" | R$ 000.000.000,00 | linha não existe mais | |
| Tabela 41 atual "TOTAL GERAL DOS BENS MÓVEIS E IMÓVEIS E LOTES / LOTEAMENTOS EXECUTADOS E REGISTRADOS" | – | **R$ 152.796.260,60** | Conferência aritmética: veículos 1.217.000,00 + máquinas 2.517.000,00 + lotes urbanos 3.290.000,00 + áreas urbanas 26.929.394,60 + áreas rurais 7.663.866,00 + loteamentos (soma dos 31 subtotais) 111.179.000,00 = 152.796.260,60 (fecha). |
| "TOTAL GERAL DOS BENS – LOTEAMENTOS CONTRATADOS A EXECUTAR OU EM FASE DE EXECUÇÃO E LICENCIAMENTOS" | R$ 00.000.000,00 | **R$ 15.754.581,00(-)** | Passou a ser um valor negativo (multas + despesas – ver 13.9). |
| "TOTAL GERAL DA EMPRESA" (antiga) / "TOTAL GERAL DO PATRIMÔNIO DA EMPRESA" (atual) | R$ 000.000.000,00 | **R$ 137.041.679,60** | = 152.796.260,60 − 15.754.581,00. |
| "TOTAL LÍQUIDO DA EMPRESA" (antiga) / "TOTAL LÍQUIDO PATRIMONIAL DA EMPRESA" (atual, tabela 51) | R$ 000.000.000,00 | **R$ 100.874.027,60** | = 137.041.679,60 − 16.240.000,00 − 3.800.000,00 − 16.127.652,00. |

### 13.9 Loteamentos contratados a executar / em licenciamento (tabelas 41–43 antigas → 42–44 atuais)

Na versão antiga eram 15 itens, todos com valor "R$ 00.000.000,00". Na atual são 9 itens, com coluna "VALOR MULTA" (negativa) e "DESPESAS".

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| 01 Caminhos do Sol – Família Boff, Caxias (612 lotes, custo R$ 21.163.000,00; Buriti 50% / SEI 10% = 61 / proprietários 40%) | R$ 00.000.000,00 | mesmos dados + "Conforme contrato firmado com a empresa BURITI, este Loteamento está em fase de execução inicial ... caso os custos excederem ao contratado, a empresa Siviero ... deverá arcar com os mesmos." Valor: R$ 00.000.000,00 | Sem multa atribuída. |
| 02 Vale das Colinas – Piccoli + Gasparin + Zanrosso, Caxias | "ESTIMATIVA DE 424 LOTES"; custo R$ 13.710.000,00; SEI 10% = 42 | "ESTIMATIVA DE 356 LOTES"; custo R$ 17.800.000,00; "MULTA ESTIMADA EM 20,00% DO CUSTO"; "VALOR VENAL DOS IMÓVEIS: R$ 39.743.378,00"; SEI 10% = 35; "AINDA NÃO ESTÁ SENDO EXECUTADO"; ônus de multa e desfazimento de escrituras. **VALOR MULTA R$ 3.560.000,00(-) / DESPESAS R$ 651.035,00(-)** | |
| 03 Vista da Colina – Zanandréa + Fioreze, Caxias (242 lotes) | custo R$ 8.356.000,00 | custo R$ 12.100.000,00; valor venal R$ 20.960.000,00; não executado. **MULTA R$ 2.556.900,00(-) / DESPESAS R$ 108.000,00(-)** | |
| 04 antigo – Girardi + Matioda + Wiltgen, Caxias (220 lotes, custo R$ 10.846.973,25) | R$ 00.000.000,00 | **item removido** | |
| 05 antigo → 04 atual – Pedrotti + Andreis + Lorandi, Caxias | 337 lotes; custo R$ 11.162.416,41; SEI 10% = 33 | 242 lotes; custo R$ 12.100.000,00; valor venal R$ 11.403.230,00; SEI 10% = 24; não executado. **MULTA R$ 2.280.646,00(-) / DESPESAS R$ 308.000,00(-)** | |
| 06 antigo – Santa Corona III – Família Paulety, Caxias (89 lotes; SEI 61,54% = 55; R$ 3.302.795,64 em lotes; "DALFOVO – tem ..........LOTES") | R$ 00.000.000,00 | **item removido** do doc 11 | Passou a constar apenas na lista "SEM OBRAS" do documento HOTEL PRINCESA ... (ver seção 14). |
| 07 antigo – Alto das Graças II – Família Bellini, Caxias (34 lotes) | R$ 00.000.000,00 | **item removido** | Idem (lista "SEM OBRAS"). |
| 08 antigo → 05 atual – Família Cecconello, Caxias | "ESTIMATIVA DE ??? LOTES"; "SIVIERO ... = 50,00% = ???"; 2,0 ha adquiridos por R$ 600.000,00 em 06/02/2014 | "ESTIMATIVA DE 130 LOTES"; custo R$ 6.500.000,00; "HECTA EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = 50,00% = 65 / PROPRIETÁRIOS = 50,00% = 65"; **MULTA R$ 1.300.000,00(-)** | O parceiro passou de "Siviero" para "HECTA". |
| 09 antigo – Bella Visione – Família Truccollo, Nova Prata (75 lotes) | R$ 00.000.000,00 | **item removido** | Lista "SEM OBRAS". |
| 10 antigo – Champagnat – Irmãos Maristas, Veranópolis (961.181,45 m², 4 etapas) | R$ 00.000.000,00 | **item removido** | Lista "SEM OBRAS". |
| 11 antigo → 08 atual – Industrial Dr. Letti, Farroupilha | sem custo; "??? LOTES" | custo R$ 8.000.000,00; multa 20% + dano moral R$ 50.000,00; risco de APP; **MULTA R$ 1.650.000,00(-)** | |
| 12 antigo → 06 atual – Bonetto, Farroupilha (61 lotes) | sem custo | custo R$ 3.660.000,00; multa 20% + dano moral R$ 50.000,00; risco de APP; **MULTA R$ 782.000,00(-)** | |
| 13 antigo → 07 atual – Pasqual, Farroupilha (29 lotes) | sem custo | custo R$ 7.250.000,00; multa 20% + dano moral R$ 50.000,00; risco de APP; **MULTA R$ 1.500.000,00(-)** | |
| 14 antigo → 09 atual – Sol Nascente III, Antônio Prado | "ESTIMATIVA DE 84 LOTES"; SEI 55% = 46; proprietários 45% = 38 | "ESTIMATIVA DE ??? LOTES"; 55% = ???; custo R$ 5.040.000,00; multa 20% + dano moral; **MULTA R$ 1.058.000,00(-)** | Quantidade de lotes deixou de ser informada. |
| 15 antigo – Família Zanella II – Condomínio Morada dos Ipês II, Passo Fundo (41 lotes) | R$ 00.000.000,00 | **item removido** | Lista "SEM OBRAS"; 3 lotes deste condomínio aparecem no acerto com Rovílio (ver 13.12). |
| TOTAL da seção | R$ 00.000.000,00 | **R$ 15.754.581,00(-)** | Conferência aritmética: soma das multas e despesas = 15.754.581,00 (fecha). |

### 13.10 Seções novas de passivo (tabelas 46–51 – só na versão atual)

| Campo | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Tabela 46/47 "LOTES DE TERCEIROS PASSÍVEIS DE INDENIZAÇÃO POR CARACTERIZAÇÃO AMBIENTAL NA CIDADE DE FARROPILHA" | não existia | 01 Firminio Dalzóchio – 3 lotes: O09/0910 R$ 130.000,00; 003/1292 R$ 135.000,00; 004/1292 R$ 135.000,00. 02 Jardim Itália I-II-III – 5 lotes: 007/1741 e 009/1741 R$ 193.000,00 cada; 001/1747, 002/1747, 009/1747 R$ 200.000,00 cada. 03 Perini – 4 lotes: 011, 012, 018, 19 da quadra 1469, R$ 160.000,00 cada. 04 Primavera – 4 lotes: 005, 006, 010, 011 da quadra 1199, R$ 166.000,00 cada. 05 Industrial Vicentino – 10 lotes: 003/1301, 004/1301, 004/1302, 009/1302, 010/1302, 011/1302, 001/1303, 012/1303, 014/1303, 016/1303, R$ 1.355.000,00 cada. | Lotes já vendidos a terceiros que podem ter de ser indenizados. |
| Tabela 48 – valor dos lotes de terceiros | – | **R$ 16.240.000,00(-)** | Conferência: 400.000 + 986.000 + 640.000 + 664.000 + 13.550.000 = 16.240.000 (fecha). |
| Tabela 49 – "compromissos à pagar em obras finais" – Colina do Retiro II | constava só como parágrafo nas Observações: custo "aproximadamente R$ 3.500.000,00 (Três milhões de reais)" | tabela própria: pavimentação + elevatória CORSAN + limpeza + arborização + "instalação da rede de iluminação pública" = "aproximadamente R$ 3.800.000,00"; valor **3.800.000,00 (-)**; garantia hipotecária de lotes da empresa e da Família Frainer | Valor subiu R$ 300.000,00 e ganhou o item iluminação pública. |
| Tabela 50 – "Demais compromissos a pagar em obras, empréstimos, prestação de serviços, honorários, acertos, multas ambientais, processos administrativos, etc., elencados nas OBSERVAÇÕES, itens 1 ao 24" | – | **R$ 16.127.652,00 (-)** | Conferência: soma dos valores destacados nas Observações (13.12) = 16.127.652,00 (fecha). |
| Tabela 51 – TOTAL LÍQUIDO PATRIMONIAL DA EMPRESA | (tabela 45 antiga: R$ 000.000.000,00) | **R$ 100.874.027,60** | |

### 13.11 Quadro "VALOR REAL DE CADA SÓCIO CONFORME A COTA DE CAPITAL" (tabela 46 antiga → 52 atual)

| Sócio | Versão antiga | Versão atual |
|---|---|---|
| Irineu / Anair, Raul, **Valter / Benta**, Elcio, Adélia / José Valério, Pedro | "16,66% = R$" (em branco) | **16,66% = R$ 16.812.337,93** (cada) |
| TOTAL | "100,00%= R$" (em branco) | 100,00% = R$ 100.874.027,60 |
| Tabela 47 antiga "ROVÍLIO SIVIERO / 225.676.150-91 / 2,00% = R$" + parágrafo "VER ACORDO" | existia (em branco) | **removida** |

Conferência: 16.812.337,93 x 6 = 100.874.027,58 (diferença de R$ 0,02 por arredondamento).

### 13.12 OBSERVAÇÕES – CONTRATOS | INDENIZAÇÕES | COMPROMISSOS | EMPRÉSTIMOS

Na versão atual cada item passou a ter uma linha de valor isolada abaixo do texto (os valores somam os R$ 16.127.652,00 da tabela 50).

| Item | Versão antiga | Versão atual | Comentário |
|---|---|---|---|
| Venda das quotas de Rovílio | – | parágrafo novo: INSTRUMENTO PARTICULAR DE COMPRA E VENDA DE QUOTAS de 28/08/2025, quotas diluídas em partes iguais aos remanescentes | |
| Acerto com Rovílio | "As cotas de capital do sócio Rovílio Siviero, que correspondem a 2,00% ... foram devidamente acertadas ... através de imóveis – lotes – e parte em dinheiro, da seguinte forma: 04 (Quatro) lotes localizados no Condomínio Morada dos Ipês, no Município de Passo Fundo + 02 (Dois) lotes localizados no Loteamento Nossa Senhora da Saúde, no Município de Guaporé + o valor de R$ 110.000,00 ... em espécie." (sem valores dos lotes) | "... 2,00% nas empresas – Siviero Empreendimentos Imobiliários Ltda., Abastecedora de Combustíveis Siviero Ltda., Transportadora Irmãos Siviero Ltda. e de 1,00% nas empresas – Turismo Siviero Ltda. e Siviero Transportes Ltda. ... no total de R$ 3.440.082,00": Morada dos Ipês I – lote 13/quadra 02 R$ 500.000,00 PAGO; lote 10/quadra 02 R$ 587.510,00 PAGO; lote 07/quadra 04 R$ 432.072,00 PAGO. N. S. da Saúde (Guaporé) – lote 12/quadra 237 R$ 75.000,00 PAGO; lote 15/quadra 252 R$ 60.000,00 PAGO. R$ 110.000,00 em espécie. "R$ 25.500,00 (Vinte e cinco mil reais), pagos em espécie". Morada dos Ipês II – 3 lotes (número/quadra "....") R$ 550.000,00 cada, À PAGAR = **R$ 1.650.000,00** | De 4 para 3 lotes no Morada dos Ipês I; acrescidos 3 lotes À PAGAR no Morada dos Ipês II. Conferência: soma = 3.440.082,00 (fecha). Inconsistência: "R$ 25.500,00" por extenso "Vinte e cinco mil reais". O saldo a pagar (R$ 1.650.000,00) entra como compromisso. |
| Parágrafos gerais removidos | "As áreas de terras conforme contratos relacionados nas Folhas 22, 23 e 24 ... NÃO estão contabilizadas como bens imóveis ..."; "Daquelas áreas com obras ... pendem de registro imobiliário ..."; "Considerado o valor investido em cada Loteamento, de forma igualitária ..."; "Não estão considerados, neste valor, ... os adiantamentos de valores às Famílias proprietárias ..."; "Ressalta-se que cada CONTRATO de LOTEAMENTO deverá ser analisado de forma individual."; "Analisar a forma de pagamento da DAÇÃO EM PAGAMENTO ..."; "Ver as antecipações feitas em espécie ..." | todos removidos | Substituídos pela quantificação de multas (13.9). |
| Loteamento Perini – infração ambiental | "Também não, as infrações ambientais ... no valor aproximado de R$ 5.000.000,00" (apenas citado, não contabilizado) | "Existe em tramitação uma infração ambiental ... R$ 5.000.000,00" + linha de valor **R$ 5.000.000,00** | Passou a ser contabilizado como compromisso. |
| Loteamento Dalzóchio – fonte de água (MP x Município x Siviero x Família Dalzóchio, desde 2006) | "no valor aproximado de R$ 500.000,00 (Quinhentos mil reais)" | "no valor aproximado de R$ 2.200.000,00 (Dois milhões e duzentos mil reais)" + linha **R$ 2.200.000,00** | Aumento de R$ 1.700.000,00. |
| Jardim Itália – canalização / APP | texto | texto igual, sem valor | sem alteração |
| San Marino – ETE | texto sem valores | acrescido: "5.1) Valor do terreno: R$ 230.000,00"; "5.2) Valor do custo aproximado para a construção da ETE: R$ 300.000,00"; "TOTAL: R$ 530.000,00" + linha **R$ 530.000,00** | |
| Lago Negro – APP em 7 lotes | "a empresa deverá ainda fazer a indenização de mais 02 lotes para a Família Candiago, objeto de Dação em Pagamento." + parágrafo sobre compensação de 03 lotes em área institucional do Município de Gramado | "a empresa fez a indenização de 02 lotes para a Família Candiago e Terceiro, objeto de Dação em Pagamento." + "Carece o cumprimento do PRAD ... aproximadamente R$ 80.000,00" + linha **R$ 80.000,00**; parágrafo da compensação com o Município de Gramado **removido** | Indenização dada como feita; compensação administrativa não é mais mencionada. |
| Família Acordi – multa MP Farroupilha | texto (multa de ≈R$ 1.000.000,00 reduzida para ≈R$ 80.000,00 já pagos; TAC/PRAD) | texto igual + "Carece o cumprimento do PRAD, sua execução final ... aproximadamente R$ 100.000,00" + linha **R$ 100.000,00** | |
| Escrituras públicas dos imóveis dos sócios | "tributação ... 5,93% (PIS-COFINS-IRPJ-CSLL)" | "5,93% ... que se houver a aplicação do adicional de 10,00%, totaliza-se um percentual de 7,00%" | |
| Áreas rurais adquiridas por contrato e não escrituradas | texto sem valor | texto + "Carece num custo aproximado de R$ 250.000,00" + linha **R$ 250.000,00** | |
| Custos 7% impostos e 6% comissão sobre vendas | texto | texto igual (reposicionado) | sem alteração de conteúdo |
| CELESTINO ROSSI – arquiteto | acerto em lotes, sem valor | "O valor dos honorários equivale a 03 (três) lotes urbanos, no valor global de R$ 600.000,00" + linha **R$ 600.000,00** | |
| SOLAR SERRA / MÁRIO BORDIGNON – placas solares Hotel Araçá | R$ 150.000,00 = 01 lote urbano | igual + linha **R$ 150.000,00** | |
| ALEXANDRE SUSIN – comissão contratos Buriti | "restando ainda a entrega de ......... lotes ... no valor de aproximadamente R$ ......" | "restando ainda a entrega de 06 lotes ... aproximadamente R$ 1.200.000,00" + linha **R$ 1.200.000,00** | Preenchido. |
| LUIS CARLOS MERCALI – empréstimo jun/2020, 1% a.m. | R$ 200.000,00 | igual + linha **R$ 200.000,00** | |
| INES TREVISAN PICCOLI – empréstimo dez/2021, 1% a.m. | R$ 350.000,00 | igual + linha **R$ 350.000,00** | |
| SAMIR TREVISAN PICCOLI – empréstimo dez/2021, 1% a.m. | R$ 1.150.000,00 | igual + linha **R$ 1.150.000,00** | |
| LOTEAMENTO RESERVA FALAVIGNA – Flores da Cunha (pavimentação + parque + galeria pluvial) | "custo estimado em aproximadamente R$ 950.000,00" | "aproximadamente R$ 2.000.000,00 ... Para a construção da galeria ... R$ 1.200.000,00 e para ... PARQUE ... R$ 800.000,00" + linha **R$ 2.000.000,00** | Aumento de R$ 1.050.000,00. |
| LOTEAMENTO COLINA DO RETIRO II – Veranópolis | parágrafo: "aproximadamente R$ 3.500.000,00 (Três milhões de reais)" | parágrafo removido das Observações; virou a tabela 49 com R$ 3.800.000,00 (-) | Ver 13.10. |
| BANCO CATERPILLAR FINANCIAL – escavadeira 320 (2024) | 33 parcelas de R$ 10.534,91; saldo R$ 347.652,03; última em 18/04/2028 | igual + linha **R$ 347.652,00** | |
| PARANÁ EQUIPAMENTOS S/A (novo) | – | "Um Trator de Esteira, ano 2021, modelo D6N XL, CATERPILLAR ... valor total de R$ 450.000,00 ... entrada já paga R$ 130.000,00, restando ... em 31 de outubro de 2025 ... 12 parcelas de R$ 26.666,67 ... total R$ 320.000,00, última parcela 10 de outubro de 2026." + linha **R$ 320.000,00** | Novo financiamento (o trator entrou na tabela 5 por R$ 450.000,00, lá com "ano 2014"). |
| Parágrafos finais ("O valor real acima apontado não é líquido...", redução por regularização "R$ ............000,00", custos notariais/registrais em branco, balanço contábil) | presentes | presentes, apenas reordenados; valores continuam pontilhados | sem alteração de conteúdo |
| Assinaturas | "Élcio Siviero Pedro Siviero Rovílio Siviero" | "Élcio Siviero Pedro Siviero" | Rovílio removido. |

Total declarado (doc 11): antigo – todos os totais em branco ("R$ 000.000.000,00"); atual – bens R$ 152.796.260,60; patrimônio após multas de loteamentos R$ 137.041.679,60; **líquido R$ 100.874.027,60**. Linha Valter: "16,66% = R$" (em branco) → **16,66% = R$ 16.812.337,93**.

---

## 14. Documentos só na versão atual (sem contrapartida antiga)

- **"HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx"** – não há versão antiga; nada a comparar. Conteúdo, em resumo: Fazenda São Crispin 686,75 ha (−190,00 ha vendidos à Família Rigo em 2018 por R$ 9.448.000,00 = 118.100 sacas de soja; −85,32 ha vendidos à Família Pellegrini por R$ 2.100.000,00 + 41 cabeças de gado; saldo 411,43 ha); Lote 24 e 24/E Cotiporã (Celito Pierozan, Matrículas 2.424 e 2.425); "TOTAL RECEBIDO DAS ÁREAS VENDIDAS R$ 12.058.000,00"; sacas de soja a receber (37.694 sacas Rigo + 13.085 sacas Lorenzet, valores em branco); contrato de arrendamento com a Família Lorenzetti (2023–2026). Segunda parte: lista dos loteamentos contratados pela SEI "SEM OBRAS": Alto das Graças II, Santa Corona III (Caxias); Condomínio Morada dos Ipês II, Congregação das Irmãs de São Carlos (Passo Fundo); Bella Visione (Nova Prata); Bonetto, Pascoal, Letti, Colombo (Farroupilha); Champagnat, Simioni (Veranópolis); Sol Nascente (Antônio Prado). Observação: Alto das Graças II, Santa Corona III, Morada dos Ipês II, Bella Visione e Champagnat são exatamente os itens retirados da tabela de loteamentos contratados do doc 11.
- **"ACERTOS E ACORDOS.pdf"** (2,4 MB) – não há versão antiga; arquivo PDF, não comparado (fora do escopo docx).

---

## 15. Resumo

**Documentos sem alteração:** 03 (Associação Cultural), 04 (Ceval), 07 (Rádio).

**Documentos alterados:** 00, 01, 05, 06, 08, 09, 10, 11.

**Padrão comum a 00, 01, 08, 09, 10 e 11 – saída de Rovílio Siviero:** linha de Rovílio (2% ou 1%) removida das tabelas de rateio; os 6 sócios remanescentes passam a 16,66% (antes 16,33% ou 16,50%); parágrafo novo citando o Instrumento Particular de Compra e Venda de Quotas de 28/08/2025; Rovílio retirado das assinaturas. Totais das empresas não mudaram por isso (01: R$ 9.661.414,00; 08: R$ 16.374.986,00).

**Reavaliações que mudaram totais:**
- 05 Hotel Princesa dos Vales: Fazenda São Crispin de R$ 130.000/ha para R$ 110.000/ha → total R$ 83.640.801,00 → R$ 75.410.201,00 (−R$ 8.230.600,00); cota/sócio R$ 13.940.133,50 → R$ 12.568.366,83.
- 06 Hotel Araçá: obra civil de 25% para 15% da coisa feita → total R$ 61.790.460,00 → R$ 58.076.823,20 (−R$ 3.713.636,80); cota/sócio R$ 10.298.410,00 → R$ 9.679.470,53.
- 00 Planilha: total parcial R$ 174.480.661,00 → total R$ 263.410.451,80 (agora inclui SEI R$ 100.874.027,60); cota por sócio R$ 43.897.245,30; tabela nova com a cota por empresa.

**Alterações mais relevantes para a SEI (doc 11):**
1. Totais finalmente preenchidos: bens R$ 152.796.260,60; patrimônio R$ 137.041.679,60; **líquido R$ 100.874.027,60**; cota de cada sócio (inclusive Valter/Benta) R$ 16.812.337,93.
2. Colina do Retiro II: 91 lotes sem valor → 102 lotes valorados, subtotal R$ 20.181.000,00 (4 lotes retirados, 15 acrescentados).
3. Nova seção de passivo "lotes de terceiros passíveis de indenização por caracterização ambiental em Farroupilha": 26 lotes, R$ 16.240.000,00 (-), dos quais 10 lotes do Industrial Vicentino a R$ 1.355.000,00 cada.
4. Loteamentos contratados: de 15 itens sem valor para 9 itens com multa estimada de 20% do custo + despesas = R$ 15.754.581,00 (-); 6 loteamentos retirados do doc 11 (Girardi/Matioda/Wiltgen, Santa Corona III, Alto das Graças II, Bella Visione, Champagnat, Morada dos Ipês II – os cinco últimos agora na lista "SEM OBRAS" do documento do Hotel); Cecconello passou de parceiro "Siviero" para "HECTA".
5. "Demais compromissos" (observações itens 1–24) quantificados em R$ 16.127.652,00 (-); maiores: Perini R$ 5.000.000,00; Dalzóchio R$ 500.000,00 → R$ 2.200.000,00; Reserva Falavigna R$ 950.000,00 → R$ 2.000.000,00; saldo Rovílio R$ 1.650.000,00; Alexandre Susin R$ 1.200.000,00; Samir Piccoli R$ 1.150.000,00; Celestino Rossi R$ 600.000,00.
6. Colina do Retiro II – obras finais: R$ 3.500.000,00 → R$ 3.800.000,00 (-), agora em tabela própria.
7. Acerto com Rovílio detalhado e totalizado em R$ 3.440.082,00 (3 lotes Morada dos Ipês I pagos, 2 lotes Guaporé pagos, R$ 135.500,00 em espécie, 3 lotes Morada dos Ipês II a pagar R$ 1.650.000,00); lote O01/quadra 2 do Morada do Ipê I saiu do inventário (R$ 538.000,00).
8. Lotes marcados "ÁREA DE POSSÍVEL APP" e retirados das somas: San Marino (3 lotes, −R$ 730.000,00), Dalzóchio (1, −R$ 131.000,00), Primavera (2, −R$ 335.000,00), Industrial Vicentino (1, −R$ 1.125.000,00). Lago Negro: 5 lotes valorados (R$ 4.021.000,00), lote 014/quadra 05 removido.
9. Imóveis retirados: lote Rua Bento Gonçalves/Michelle Tedesco (Matrícula 2325, R$ 220.000,00); Lote 1543 BR 470/Família Bissani (Matrícula 21.819, R$ 900.000,00); Lote Rural 52 Cotiporã/Bassani (Matrícula 7.391, R$ 226.500,00 – permanece no doc 05). Imóvel novo: lote 623/quadra 104 Reserva Paradiso, Carlos Barbosa, R$ 500.000,00.
10. Reavaliações de imóveis: Família Lago (Matrícula 14.319, 3,6 ha rural, R$ 900.000,00 → Matrícula 33.516, 4,0 ha urbano, R$ 1.500.000,00); Lote Zanini R$ 900.000,00 → R$ 1.200.000,00; Família Sangali R$ 143.000,00 → R$ 520.000,00; Av. Osvaldo Aranha/Eduardo Duarte R$ 850.000,00 → R$ 550.000,00; chácaras Barra do Jabuticaba R$ 100.000,00 → R$ 20.000,00; Matrículas 17.409 + 6.136 unificadas na 29.954 (valor igual).
11. Bens móveis: veículos total R$ 1.217.000,00 (Hilux 2018 R$ 188.000,00 e RAV4 2019 R$ 155.000,00 no lugar de semirreboque artesanal e Jeep Compass 2022); máquinas total R$ 2.517.000,00 (novo trator D6N XL R$ 450.000,00, financiado – saldo R$ 320.000,00 Paraná Equipamentos).

**Pontos a conferir (inconsistências encontradas nos documentos atuais):**
- Doc 00, tabela 2, linha 11: "R$ 100.874,027,60" (vírgula no lugar do ponto); tabela 3: "SIVIERO EMPRENTOS".
- Doc 01, tabela 3, item 01: valor do item apagado (célula só com "R$"); total segue R$ 9.661.414,00.
- Doc 11: trator D6N XL "ano 2014" na tabela 5 e "ano 2021" nas observações; "R$ 25.500,00 (Vinte e cinco mil reais)" no acerto com Rovílio; Sol Nascente III perdeu a estimativa de lotes (84 → "???").
