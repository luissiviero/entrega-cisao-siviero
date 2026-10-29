# Revisão independente – Pacote 4 (rascunho): imóveis da SEI fora de loteamento, 2018 × 2025

Data: 28/10/2026. Revisor independente (contexto novo). Arquivo revisado: `02_trabalho/pacotes/pacote-04-rascunho.md`.

Fontes conferidas nos documentos originais (texto extraído com python-docx, parágrafos e células de tabelas, em ordem):
- 2018: `01_fontes/inventario-2018/18. 1. SIVIERO EMPREENDIMENTOS ... Dezembro de 2018.docx` (citado abaixo como "doc 2018"; T = tabela, R = linha) e `02_trabalho/sei-2018-itens.xlsx`, aba Itens. Para contexto: docs 21, 25 e 27 de 2018.
- 2025: `01_fontes/inventario-2025/11. INVENTÁRIO - 31.12.2024 - SIVIERO ...docx` ("doc 11"), `02_trabalho/sei-2025-itens.xlsx`, aba Itens; docs 00 a 10 e "HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx" ("doc Hotel-áreas vendidas").
- Versão antiga: `01_fontes/inventario-2025-antigo/ANTIGO 05` e `ANTIGO 11`, e o relatório `02_trabalho/diff-inventario-2025-antigo-x-atual.md`.
- Para ver de onde vinham os pares antigos: `02_trabalho/planilha-mestre-sei.xlsx`, aba Mestre.

Nenhum arquivo de origem foi alterado. Scripts de apoio ficaram fora do projeto, na pasta temporária `revisor4`.

---

## Gravidade ALTA

### A1. A SEDE CAMPESTRE de 2018 (matrículas 1.941 e 7.104) não aparece no rascunho

- **Doc 2018, T35 R0** (bloco "BENS DA EMPRESA SIVIERO ... E RELACIONADOS NOS INVENTÁRIOS DE OUTRAS EMPRESAS"): "Um terreno rural, com a área de 9,4 hectares, ... Rua Luiz Siviero, ... Bairro Femaçã, ... Matrículas nº 1.941 e 7.104, conhecido como SEDE CAMPESTRE". Valor: R$ 3.879.091,80, assim composto: terra 9,4 ha × R$ 300.000 = R$ 2.820.000; pavilhão R$ 664.686; casa R$ 150.000; 30% da obra R$ 244.405,80. Rótulo: "Relacionado no Inventário da empresa: HOTEL". Na xlsx 2018 é o nº seq 147.
- **2025, doc 05 (HOTEL PRINCESA DOS VALES LTDA.), tabela 5, item 04**: mesmo texto, mesmas matrículas 1.941 e 7.104, 9,4 ha. Valor R$ 5.560.857,50, assim composto: terra 9,4 ha × R$ 450.000 = R$ 4.230.000; pavilhão R$ 664.686; quiosque R$ 150.000; casa R$ 250.000; 25% da obra R$ 266.171,50. A soma confere.
- As matrículas 1.941 e 7.104 não aparecem no doc 11 nem nos docs 01–04 e 06–10. Aparecem só no doc 05, tanto na versão atual quanto na antiga.
- **Correção:** acrescentar a Sede Campestre ao pacote. O tratamento é igual ao do caso 8: não é saída; o imóvel está no inventário 2025 do Hotel Princesa dos Vales, doc 05, item 04. Na Parte 6, conferir que está contado uma única vez. Sem esta linha, o rascunho não cobre os 9 bens de "outras empresas" de 2018: cobre 8.

### A2. O par errado "área rural 11 (Piccolo) × 2018 Áreas #08 (Ansolin)" continua na planilha-mestre, e o rascunho não o desfaz

- **planilha-mestre-sei.xlsx, aba Mestre, linha SEI-2025-ARUR-11 (Piccolo)**: "Existe em 2018? = Sim"; "Item 2018 = Áreas urbanas e rurais em Ve #08"; descrição 2018 "PARTE DO LOTE RURAL nº 73 ..."; valor 2018 R$ 900.000; similaridade 0,92. Pela mesma lógica, a linha da área rural 06 (Ansolin, 30.390) está com "Existe em 2018? = Não".
- O rascunho propõe, no caso 4, o par área rural 06 × Áreas #08 e, no caso 2, classifica a área rural 11 como adquirida depois de 2018. Mas, diferentemente do caso 1 (Scalco × lote 24), não manda desfazer o par 11 × #08 que está gravado. Se o pacote for aprovado como está, o item 2018 Áreas #08 fica ligado a dois itens de 2025.
- **Correção:** no caso 2, acrescentar "desfazer o par área rural 11 × 2018 Áreas #08 (Ansolin), hoje gravado na planilha-mestre", com o mesmo texto do caso 1.

---

## Gravidade MÉDIA

### M1. Caso 4 (Ansolin): área errada, vendedor simplificado, data não comprovada e um item de 2018 esquecido

- **Área.** O rascunho diz "matrícula 30.390, 20,2 ha". O doc 11, tabela 8, item 06, diz: "Matrícula nº 30.390 ... com área de 202.435,22m²", "área pertencente ao Senhor Antônio Ansolin ... 2,0 hectares = 20.000,00m²" e "Área da empresa ... é de 18,2 hectares = 182.435,22m²". Os 20,2 ha são a área total da matrícula; a área da SEI é **18,2 ha**.
- **Inconsistência no próprio doc 11.** As parcelas somam 60.000 + 70.519,78 + 49.787,44 + 6.700,00 + 2.128,00 = **189.135,22 m² (18,9 ha)**, e não 182.435,22 m². A diferença, 6.700 m², é exatamente a parcela de Ivo José Frainer (0,67 ha). O preço dela, R$ 388.000, está somado no valor de R$ 2.263.666, mas a área parece ficar fora da área declarada. Esse ponto deve entrar nas perguntas da Parte 4.
- **"Mesmo vendedor".** Em 2025 há três vendedores: Antônio Ansolin, Inês Ansolin Pessutto e Ivo José Frainer. A prova forte do par é outra: a primeira parcela de 2025, "área de 6,0 hectares = 60.000,00m², pelo valor de R$ 525.000,00, do Senhor Antônio Ansolin", é idêntica ao item 2018 Áreas #08 (doc 2018, T7 R8: "Adquirida a parte ideal de 6,0 hectares, pelo valor de R$ 525.000,00").
- **"Compras posteriores a 2018".** O doc 11 não dá data a nenhuma parcela. O documento "ACERTOS E ACORDOS" cita pagamentos a Antônio Ansolin "referente a compra de 7,0 ha", também sem data. A afirmação não tem suporte documental: deve ser escrita como hipótese ou virar pergunta.
- **Item de 2018 esquecido.** Em 2018, a área restante já estava contratada. Doc 2018, T16 R2 (loteamentos contratados, item 02): "Loteamento Pedra Lisa II – Ansolin ... a empresa ... tem adquirido 6,0 hectares de terras. Área total de 200.246,25m². Área de 2,00 ha não fará parte do Loteamento, pois pertence ao Senhor Antônio Ansolin.", com valor R$ 50.000. Esse loteamento contratado não aparece em 2025: nenhuma ocorrência de "Pedra Lisa II" nem de "Ansolin" fora da tabela 8 do doc 11. O pacote, ou o pacote de loteamentos, precisa dizer que esse item de 2018 se transformou na área rural 06.
- **Matrícula.** Ela mudou de 21.771 (2018) para 30.390 (2025). O rascunho pede a cadeia de matrículas no caso 3, mas não aqui.
- **Correção:** trocar "20,2 ha" por "18,2 ha da SEI (matrícula com 20,2 ha, dos quais 2,0 ha de Antônio Ansolin)". Apoiar o par na parcela idêntica de 6,0 ha por R$ 525.000. Tratar o acréscimo como "a datar". Ligar o item de 2018 "Loteamento Pedra Lisa II – Ansolin" (R$ 50.000). Pedir a cadeia 21.771 → 30.390 e a explicação da parcela Frainer.

### M2. Falta em A o par Sangali, matrícula 3.062 (2018 "outras empresas" #03 × 2025 área rural 04)

- **Doc 2018, T33 R1:** "PARTE DO LOTE RURAL Nº 12, com área total de 13,0 hectares + 7,5 hectares ... FAMÍLIA SANGALI ... Matrícula nº.: 3.062. Matrícula nº 3.062 = 20,5 hectares X 11.000,00", CCIR 2.481.950-6, valor R$ 225.500,00, "Relacionado no Inventário da empresa: AGROPECUÁRIA".
- **Doc 11, tabela 8, item 04:** "PARTE DO LOTE RURAL Nº 12, com área total de 13,0 hectares ... FAMÍLIA SANGALI ... Matrícula nº.: 3.062", mesmo CCIR 2.481.950-6, valor R$ 520.000,00.
- O par só aparece de forma implícita no caso 7. A tabela A não tem a linha; por isso a área rural 04 de 2025 não está pareada formalmente, e a contagem "todo item 2025 pareado ou adquirido depois" não fecha.
- **Correção:** acrescentar em A a linha "Outras empresas: Agropecuária (lote 12, Sangali) | Áreas rurais #04 | 3.062 (área 20,5 → 13,0 ha; ver caso 7)".

### M3. Caso 7: falta a hipótese de dupla contagem em 2018

- Os 7,5 ha a mais em 3.062 têm o mesmo tamanho da área do lote 14, também da família Sangali, na mesma Linha Júlio de Castilhos. Doc 2018, T7 R5: "PARTE DO LOTE RURAL Nº 14, da linha Dr. Júlio de Castilhos, com área total de 7,5 hectares ... FAMÍLIA SANGALI – VALDIR SANGALI ... Matrícula nº.: 4.194". Esse item estava no bloco próprio da SEI. O item 3.062 estava no bloco "outras empresas / AGROPECUÁRIA". Em 2025 os dois aparecem separados: 4.194 com 7,5 ha e 3.062 com 13,0 ha.
- O fato citado pelo rascunho, "a área rural 03 tem outra matrícula", está correto. Mas não afasta a hipótese de que o inventário de 2018 tenha contado os mesmos 7,5 ha duas vezes. Se for isso, não houve saída.
- **Correção:** no caso 7, acrescentar a hipótese "possível dupla contagem em 2018 com o lote 14 (4.194)" e pedir a matrícula 3.062 atualizada. Para referência, ao preço unitário de 2018 os 7,5 ha valem 7,5 × R$ 11.000 = R$ 82.500.

### M4. Caso 1: a classificação do lote 24 "fora da análise 2018→2025" não segue a premissa

- **Doc 2018, T33 R0:** "Vendido em 2018, a área total para o Senhor ANGELI, pelo valor de R$ 510.000,00 (Quinhentos e dez mil reais), de forma, À VISTA." Na coluna valor: "V E N D I D A / Família ANGELI / ... R$ 408.000,00".
- **"Ainda somado": confirmado.** SUB-TOTAL REAL de R$ 105.940.226,35 menos TOTAL LÍQUIDO de R$ 75.146.039,95 dá R$ 30.794.186,40. Esse valor é exatamente a soma dos 9 bens de "outras empresas", incluindo os R$ 408.000 do lote 24: 7.148.016 + 4.950.000 + 13.631.378,60 + 408.000 + 225.500 + 250.200 + 75.500 + 226.500 + 3.879.091,80. Os R$ 408.000 entram no SUB-TOTAL REAL, não no TOTAL LÍQUIDO; o rascunho deve dizer isso.
- **Fonte de 2025 que o rascunho não cita.** O doc Hotel-áreas vendidas, tabela 1, linha 2, repete o lote 24, matrículas 2.424 e 2.425, "10,2 hectares X R$ 50.000,00", e "Vendido em 2018 ... ANGELI ... R$ 510.000,00". O valor entra no "TOTAL RECEBIDO DAS ÁREAS VENDIDAS ... R$ 12.058.000,00", num documento com o título do Hotel Princesa dos Vales.
- **Classificação.** A premissa diz que todo bem de 2018 que saiu precisa de destino documentado. O lote 24 está no inventário de 2018 e no subtotal, e o destino está documentado. Ele não fica "fora da análise": é uma saída com destino documentado, a venda de 2018 por R$ 510.000 à vista. Deve entrar na Parte 4 já resolvido, com uma nota sobre onde entrou o dinheiro, porque a venda aparece num documento do Hotel.
- **Citação.** O trecho "vendida em 2018 por R$ 510.000" está entre aspas, mas não é o texto do documento. Usar a citação literal acima.

### M5. Itens retirados entre as versões de 2025 e a data-base: falta o registro (premissa P2)

- **Caso 5 (Bissani, matrícula 21.819).** Estava no ANTIGO 11, tabela 7, item 02, com R$ 900.000, e foi retirado da versão atual. O rascunho manda para a Parte 4, o que está certo. Falta dizer que é preciso apurar se a saída foi antes ou depois de 31.12.2024. Se foi depois, o bem pertence ao inventário de 31.12.2024 e a retirada precisa entrar em `registro-data-base.md`.
- **Lote "fundos da Casa da Nona" (Michelle Tedesco, matrícula 2325, 375 m², R$ 220.000).** Estava no ANTIGO 11, tabela 6, item 06, e foi retirado da versão atual. Não consta de 2018: nenhuma ocorrência de "Tedesco" como lote urbano, de "2325" ou de "fundos da Casa da Nona" no doc 2018. É um bem comprado depois de 2018 e retirado do inventário de 31.12.2024 sem explicação. O rascunho não o menciona. Mesmo sem par em 2018, a retirada precisa da mesma verificação de data-base.
- **Mudanças só na versão atual.** A unificação 17.409 + 6.136 → 29.954 e a nova matrícula 33.516 (no lugar de 14.319) só aparecem na versão atual. A versão antiga do mesmo inventário de 31.12.2024 ainda mostrava as matrículas anteriores. Os pares ficam mantidos, mas convém registrar a data desses atos no cartório, para não importar para 31.12.2024 um fato posterior.

---

## Gravidade BAIXA

- **B1. Lotes urbanos 2018 #11 e #12 (Colina do Sol 112/51 e 126/50) ficaram de fora.** Doc 2018, T6 R11 e R12, ambos "VENDIDO." com valor R$ 0,00. Não estão em 2025. Precisam aparecer no pacote, com a nota "já vendidos em 2018, valor zero no inventário". Sem isso, a cobertura dos 12 lotes urbanos de 2018 fica em 10.
- **B2. "Imóveis e móveis" #02 (móveis do escritório, R$ 10.000; doc 2018, T4 R2) não é mencionado.** Não é imóvel. Basta uma linha: "fora do escopo (bem móvel); o próprio inventário de 2018 descreve o destino".
- **B3. Caso 8: dados que faltam.** Em 2018, o lote 52 (7.391) estava rotulado "Relacionado no Inventário da empresa: AGROPECUÁRIA". Em 2025 não está no doc 02 (Agropecuária): está no doc 05 (Hotel Princesa dos Vales), tabela 5, item 07, "15,1 hectares X R$ 28.000,00", R$ 422.800,00, vendedora "GLORINHA E GERALDO BASSANI". Convém registrar essa troca de empresa e o valor. O resto do caso confere.
- **B4. O título de A ("Pares confirmados pela matrícula") não vale para três linhas.** As linhas Ouro Verde 569/46, Verona 188/62 e Cecconello não têm matrícula: os pares saem do cadastro municipal ou da descrição. Sugestão: "pela matrícula, cadastro municipal ou descrição idêntica".
- **B5. Lote urbano 06 (Dal Ponte, 3.788, R$ 1.010.000): "adquirido depois de 2018" é dedução pela ausência.** Não há "3.788" nem "Dal Ponte" em nenhum documento de 2018, e o doc 11 não dá data de compra. A classificação é aceitável, mas pelo valor do item convém pedir a data e o preço de compra.
- **B6. Caso 9, lote 2018 #05 ("LOTEAMENTO BOFF", Rua Fioravante Boff, R$ 70.000).** O doc 2018 também traz, em T10 R2, "Loteamento Irmãos Boff ... saldo de 01 lote, ao valor de R$ 60.000,00". Vale perguntar se é o mesmo lote, contado duas vezes.
- **B7. Caso 10.** Para a Parte 4 é útil registrar que o apartamento 303 foi "adquirido por R$ 85.000,00" (doc 2018, T4 R1) e que estava alugado em 2018. Doc 2018, T26 R15 (Aluguéis): "Apartamento ... Edifício ESPERENÇA - GELATTI ... adquirido do Sottili", R$ 300 por mês.
- **B8. Fora das listas de bens: "Apartamento 401 – Edifício Luiz Siviero".** Ele aparece só na tabela de aluguéis de 2018 (T26 R8, R$ 650 por mês) e em nenhum inventário de bens, nem de 2018 nem de 2025. Pode ser de terceiro, administrado pela SEI. É só um registro, para confirmar de quem é o imóvel.

---

## Aritmética da última linha

- O rascunho soma os itens 5, 6, 9 e 10: 900.000 (Bissani) + 70.000 (lote 0024) + 70.000 (Boff) + 100.000 (559/47) + 100.000 (247/169) + 100.000 (apartamento 303) = **R$ 1.340.000**. **Confere.** Os valores de 2018 conferem com o doc 2018 (T6 e T7) e com a xlsx.
- Os subtotais de 2018 usados para conferir também batem: lotes urbanos R$ 2.600.000 e áreas R$ 6.850.000 (doc 2018, T6 R13 e T7 R11).
- Com as correções acima a soma não muda: a Sede Campestre e o lote 52 não são saídas; o lote 24 e os lotes do Colina do Sol já estavam vendidos em 2018. A linha poderia dar um valor aos 7,5 ha do item 7, R$ 82.500 ao preço de 2018, ou dizer que eles caem se M3 se confirmar.

## Caso 8 e Sede Campestre: resposta pedida

- **Matrícula 7.391** (lote 52, Bassani, 15,1 ha): está **só no doc 05, HOTEL PRINCESA DOS VALES LTDA.**, tabela 5, item 07, R$ 422.800,00, nas versões atual e antiga. Não aparece em nenhum dos outros docs 01–10. Estava no ANTIGO 11 (tabela 9, item 10, R$ 226.500) e saiu do doc 11 atual, como diz o diff.
- **Sede Campestre** (matrículas 1.941 e 7.104, 9,4 ha): está **só no doc 05, HOTEL PRINCESA DOS VALES LTDA.**, tabela 5, item 04, R$ 5.560.857,50, nas versões atual e antiga. Não está no doc 06 (Hotel Araçá), no doc 11 nem nos demais.
- Consulta complementar: o doc 07 (Rádio) cita a matrícula 3.751 só numa observação, "de propriedade da empresa Siviero Empreendimentos Imobiliários Ltda.", sem lançá-la como bem. Não há dupla contagem. O doc 08 tem a matrícula 3.753 (oficina mecânica), que é outro imóvel, diferente da 3.754.

---

## Confirmado correto

Seção A: os 18 pares conferem com os originais em matrícula (ou cadastro), área, localização e vendedor.
- Lotes urbanos: 2018 #01 × 2025 #01 (1.597, 330 m², Casa da Nina); #02 × #02 (1.335, 362,50 m², Casa da Nona); #03 × #03 (9, 570 m², Bar do André); #04 × #04 (1.053, 260 m², Mercali, hipoteca de ICMS); #06 × #05 (1.507, 345 m², esquina com a Eduardo Duarte); #07 × #07 (cadastro 13082200; 569/46; quadra 216 → 216/D; Ouro Verde → Ouro Verde II); #10 × #08 (cadastro 12482700; 188/62; quadra 163/G; Verona).
- Áreas: #01 × rural 01 (14.443, 14.444 e 14.576; 3,4 ha; Ermelinda Mossi Remor; CCIR 6.206.801-6); #02 × urbana 01 (17.038; 16.913,967 m²; Moacir Zanini); #05 × rural 03 (4.194; 7,5 ha; Valdir Sangali); #06 × rural 05 (14.945; 1,8 ha; Fogali); #09 × rural 07 (matrícula 36; 10 chácaras); #10 × rural 10 (Cecconello, 2,0 ha, "ADQUIRIDA em 2014", texto idêntico).
- Outras empresas: Rádio 3.751 × urbana 02 (7.200 m², R$ 7.148.016); Oficina 3.754 × urbana 03 (4.986 m², R$ 4.950.000); Oficina 17.409 + 6.136 × urbana 04 (29.954, "oriunda da unificação"; 1.942,03 + 2.850,00 = 4.792,03 m²; R$ 13.631.378,60); lote 89 8.486 × rural 08 (27,8 ha, Sérgio Baldissera); lote 33 1.218 × rural 09 (15,1 ha, Armelindo Calza).

Seção B:
- **Caso 1.** Os fatos conferem: contrato de 01/11/2022 e área "ainda não registrada em nome da empresa"; matrículas 2.424 e 2.425; 10,2 ha; venda a Angeli por R$ 510.000; R$ 408.000 somados no SUB-TOTAL REAL. Desfazer o par Scalco × lote 24 é correto, porque a planilha-mestre tinha esse par com similaridade de 0,86. A ressalva sobre a classificação do lote 24 está em M4.
- **Caso 2.** Correto: Piccolo "por contrato em 06 de outubro de 2022", Scalco "01 de novembro de 2022", ambos "ainda não registrada em nome da empresa". Não há "Piccolo", "Scalco" nem "Ernesto Alves" em 2018. Falta só a correção A2.
- **Caso 3.** Correto: mesmo vendedor (Família Lago); mesmo lugar (Linha Tomaz Flores, Sapopema, CTG Rincão da Roça Reiuna/Reúna); ANTIGO 11, tabela 8, item 02, ainda com 14.319 e 3,6 ha; atual com 33.516 e 4,0 ha, "terras URBANAS". Aceitar o par e pedir a cadeia de matrículas está bem fundamentado.
- **Caso 5.** Os dados conferem (lote 1543, BR-470, Bissani, 21.819, 1,9 ha, R$ 900.000). Não está em nenhum documento de 2025 atual (docs 00 a 11 e doc Hotel-áreas vendidas). Estava no ANTIGO 11.
- **Caso 6.** Os dados conferem (lote 0024, Santa Rita, 2,6 ha = 26.100 m², R$ 70.000, matrícula e vendedor em branco, inscrição 113 0024 001 001). Não aparece em 2025, nem na versão antiga.
- **Caso 7.** Os números conferem (13,0 + 7,5 = 20,5 ha em 2018; 13,0 ha em 2025; a área rural 03 tem a matrícula 4.194). A ressalva está em M3.
- **Caso 8.** Confirmado: o imóvel está no doc 05 (ver acima).
- **Caso 9.** Os dados conferem (#05 Boff/Cleri Fávero R$ 70.000; #08 559/47, cadastro 13082100, R$ 100.000; #09 247/169, cadastro 13092200, R$ 100.000). Nenhum aparece em 2025.
- **Caso 10.** Os dados conferem (apartamento 303, Edifício Boa Esperança, 17.399, 89,73 m², R$ 100.000). Não aparece em 2025.

Outros pontos confirmados:
- Bens de 2025 sem par: lote urbano 06 (Dal Ponte, 3.788) e áreas rurais 11 e 12 não constam de 2018. Os 2025 têm 24 itens no escopo (8 lotes urbanos, 4 áreas urbanas, 12 áreas rurais). Tirando esses 3, sobram 21, todos pareados: 18 na tabela A e 3 nos casos de B (área rural 02 no caso 3, área rural 06 no caso 4 e área rural 04, só implícita, no caso 7; ver M2).
- A soma de R$ 1.340.000 confere.

## O que não pôde ser conferido

- Datas de compra das parcelas Ansolin, do lote Dal Ponte e da abertura das matrículas 29.954 e 33.516: nenhum documento disponível traz essas datas.
- O motivo da retirada do lote Bissani e do lote Tedesco entre as versões de 2025.
- O inventário de 2018 das outras empresas (Hotel, Agropecuária, Rádio, Oficina). Não está nas fontes, e por isso não foi possível ver como esses bens apareciam lá.
