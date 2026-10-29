# Revisão independente — `rovilio-quotas.md`

Revisor de contexto novo, 28/10/2026. Documento revisado: `02_trabalho/empresas/01-08-abastecedora-transportadora/rovilio-quotas.md` (12.103 bytes, 18:04). Script `_scripts/rovilio_numbers.py` rodado com `python -I`: todos os números impressos conferem com o documento.

Fontes conferidas: contrato (arq. 10), imagens p. 2 a 9 a 150 dpi (p. 1, 10 e 11 pela transcrição); arq. 1, p. 3–4, 17, 21, 23 e 29 (camada de texto); inventários 2025 atuais 00, 01, 08, 09, 10 e 11 e antigos 00, 01, 08, 09 e 11 (extrações em `_extracao/` e leitura direta dos .docx); inventários 2018 12, 15, 17, 18, 8, 7 e 27 (Drive); planilha Google 2_0, GERAL 873–894 e Resumo 60–67 (fórmulas e valores).

Nenhum arquivo existente foi alterado.

---

## Achados

### 1. ALTA — A seção 8.1 conclui mais do que o contrato diz e omite o texto que pesa contra essa conclusão

**Trecho:** "Pelo contrato, o grupo não pagou ao Rovílio 35 % acima dos próprios inventários; pagou perto do valor deles. O próprio contrato diz ainda que o pagamento é 'superior ao de fato devido'."

**Evidência:**
- Cl. 3ª (p. 4): o preço de **R$ 3.440.082,00** é o que as compradoras pagam "pelas cotas de capital … e ainda outros bens de direito por créditos hereditários e/ou creditórios". A comissão é forma de pagamento desse preço, não um valor à parte.
- Parágrafo primeiro (p. 6): os itens **"1" a "10"**, inclusive os itens 7 e 8 (comissão), "quitam … a todas as quotas de capital … contabilizadas e inventariadas do VENDEDOR, em valor superior ao de fato devido". Pela letra do contrato, a comissão integra o pagamento das quotas. O contrato é ambíguo: chama os itens 7 e 8 de comissão, mas os conta no preço das quotas.
- A frase "superior ao de fato devido" diz que o pagamento **excede** o valor das quotas. Ela sustenta a premissa do Dr. Pablo (preço acima do valor das quotas), com a ressalva do grupo de que o excesso foi concessão. No texto, ela aparece como reforço de "pagou perto do valor deles", que é o sentido oposto.
- O próprio grupo não usou o argumento da comissão. Arq. 1, p. 29, item 17: o grupo aceita "a diferença de cerca de novecentos mil reais" e a justifica por "acertos passados e futuros, de trabalhos realizados, de comissões, de parcerias", pela entrega futura dos imóveis e por "concessões recíprocas".
- A versão antiga do inventário da SEI (achado 4) descreve o acerto sem comissão nenhuma.

**Correção sugerida:** reescrever 8.1 como duas leituras, sem conclusão.
- (a) Leitura que retira a comissão: 96,84 % de 2 % pelo líquido (o grupo poderia usá-la, embora não a tenha usado).
- (b) Leitura pela letra da cl. 3ª e do parágrafo primeiro: os itens 1 a 10 pagam as quotas; o preço é 135,53 % de 2 % dos inventários; o grupo admite a diferença (p. 29) e o contrato admite pagamento "superior ao de fato devido".
- Retirar "pagou perto do valor deles". Mover a frase do "superior ao de fato devido" para a leitura (b).

### 2. ALTA — A cl. 11ª, item 5 tem a autoria errada (§ 5) e o texto exagerado (§ 8.2)

**Evidência:**
- A cl. 11ª abre com "Ainda, as partes, resolvem:" (p. 7). O item 5 (p. 8) é declaração conjunta das partes.
  - O § 5 escreve "o vendedor declara que as demonstrações…". Isso está errado.
  - O § 8.2 escreve "O grupo assinou que…". Essa versão contradiz o § 5.
- O item 5 diz que as demonstrações "são completas e refletem corretamente a **posição financeira e o resultado das operações** e foram elaborados de conformidade com os princípios contábeis". Não diz que refletem o **valor**. O título do 8.2, "refletem o valor", amplia a fonte.
- Omissões que pesam nesse ponto:
  - Cl. 11ª, item 9 (p. 8): o vendedor "NÃO será responsável" por prejuízo decorrente "da inexatidão ou inveracidade de qualquer declaração feita ou garantia dada".
  - Os próprios inventários 2025 (SEI, observações finais; a versão antiga também) dizem "O valor real acima apontado não é líquido" e "O BALANÇO PATRIMONIAL E CONTÁBIL … NÃO CORRESPONDE À REALIDADE FÁTICA DA EMPRESA".

**Correção sugerida:**
- § 5: "as partes resolvem/declaram (cl. 11ª, caput, p. 7; item 5, p. 8)".
- § 8.2: título neutro, por exemplo "O contrato remete aos inventários de 31.12.2024". Citar o item 5 com as palavras exatas. Acrescentar o item 9 e a ressalva dos próprios inventários.

### 3. MÉDIA — O § 9 aponta para a SEI o lote errado do Ipê I

**Trecho:** "Item 2: o lote 10/q02 do Ipê I … Verificar se ele sumiu do inventário 2025, que lista Ipê I com 'saldo de 01 lote', o 014/q2 de R$ 629.000."

**Evidência:**
- O inventário antigo da SEI (`ANTIGO 11.`, T13) lista o Ipê I com "saldo de 02 lotes": **001/q2, R$ 538.000,00**, e **014/q2, R$ 629.000,00**. O atual lista só o 014/q2.
- O lote 10/q02 não aparece em nenhuma das duas versões. O lote que saiu entre as versões é o **001/q2 (R$ 538.000,00)**. O diff já registra isso em `02_trabalho/diff-inventario-2025-antigo-x-atual.md`, linhas 250 e 395.

**Correção sugerida:**
- Registrar que o 10/q02 não consta do inventário de 31.12.2024 em nenhuma versão, embora o contrato diga que só foi entregue em fev./2025.
- Registrar que o 001/q2 (R$ 538.000,00) saiu do inventário sem explicação.
- Acrescentar as duas perguntas ao § 10.

### 4. MÉDIA — Falta a versão antiga do acerto na SEI, que difere da atual

**Evidência (`ANTIGO 11.`, observações):**
- O texto antigo diz: "As cotas de capital do sócio Rovílio Siviero … foram devidamente acertadas, anteriormente ao presente ato … da seguinte forma: 04 (Quatro) lotes localizados no Condomínio Morada dos Ipês … + 02 (Dois) lotes … Guaporé + o valor de R$ 110.000,00". Segue a linha "ROVÍLIO SIVIERO … 2,00% = R$" (em branco) e "VER ACORDO".
- A versão antiga não tem total, não tem comissão, não tem os R$ 25.500,00 e não tem os 3 lotes do Ipê II "À PAGAR".
- Os R$ 1.650.000,00 a pagar e o total de R$ 3.440.082,00 só aparecem na versão atual. A dedução de R$ 1.650.000,00 também só aparece nela.
- O `ANTIGO 00.` deixa o valor da SEI em branco, com total "PARCIAL" de R$ 174.480.661,00.

**Correção sugerida:** acrescentar ao § 6 um item "Versão antiga da SEI" com esse texto. Ligar o item ao 8.3 e ao § 9 (SEI): a dedução de R$ 1,65 mi e o total do acerto são posteriores à versão antiga.

### 5. MÉDIA — Os inventários usados nas contas do § 7 são posteriores ao contrato, e o documento não diz isso

**Evidência:**
- A versão atual do inventário da SEI cita o próprio contrato de 28/08/2025. Também traz o saldo da Paraná Equipamentos "restando ainda em 31 de outubro de 2025".
- O grupo diz que os inventários se referem a out./2025 (registro DB-05; arq. 1, p. 4: "termo definido para a composição – OUTUBRO/2025").
- Na versão antiga, o líquido da SEI estava em branco (achado 4).
- A planilha Google rotula a coluna E (onde está E874) como "outubro 2025".
- Logo, o preço não pode ter sido fixado como 2 % de R$ 100.874.027,60. Tanto os 35,53 % do Dr. Pablo quanto os 96,84 % do documento comparam o preço com uma base feita depois dele, base que já desconta parte do próprio preço (8.3).

**Correção sugerida:** registrar isso no § 7 e no 8.2. Sem isso, a leitura "o contrato usa esses inventários como base do preço" sugere que os números do § 7 eram os da negociação.

### 6. MÉDIA — Falta a comparação com o benchmark de 2018, que tem uma coincidência numérica relevante

**Evidência (inventários 2018 e resumo 27.1):**

| Empresa (2018) | Valor (R$) | % Rovílio | Parcela (R$) |
|---|---|---|---|
| SEI, total líquido | 75.146.039,95 | 2 % | 1.502.920,80 |
| Abastecedora | 12.361.851,89 | 2 % | 247.237,04 |
| Transportadora | 1.564.000,00 | 2 % | 31.280,00 |
| Siviero Transportes | 928.000,00 | 1 % | 9.280,00 |
| Turismo | 0,00 | 1 % | 0,00 |
| **Total** | | | **1.790.717,84** |

- Os itens já entregues ou pagos (1, 2, 5, 6, 7, 9 e 10) somam **R$ 1.790.082,00**: diferença de R$ 635,84 (0,04 %). A soma inclui o item 7 (comissão).
- A planilha do Luís põe exatamente −1.790.082 em GERAL D874, na coluna "2018: atualizado".
- O preço cheio fica 92,11 % acima dessa base de 2018.

**Correção sugerida:** registrar no § 7 como observação numérica, sem concluir (pode ser coincidência). Ela vale para os dois lados:
- sugere que a parte já paga foi calibrada pelo valor de 2018;
- sugere que o "à pagar" de R$ 1,65 mi foi acrescentado depois;
- enfraquece a separação do item 7 como comissão.

O benchmark 2018 é premissa do projeto.

### 7. MÉDIA — Faltam cláusulas do contrato relevantes para o valor e para o cliente (precedente de saída de sócio)

**Cláusulas que o documento não menciona:**
- **Cl. 11ª, item 1, 2ª frase (p. 7):** "quaisquer bens e/ou valores embora não relacionados no Inventário … estão incluídos no preço total referido na cláusula terceira". Pesa na leitura do excesso de preço. O cliente pode ler que os inventários não são completos; o grupo pode ler que o preço paga mais que o inventário. O documento cita só "porteira fechada".
- **Cl. 7ª (p. 7):** não recairá sobre o vendedor e seu cônjuge "quaisquer débitos … de épocas passadas e/ou … presente e futura … responsabilidades por passivos".
- **Cl. 11ª, item 10 (p. 8):** contas a pagar, empréstimos, avais e garantias são "de responsabilidade unicamente das COMPRADORAS". O sócio saiu sem passivo.
- **Cl. 11ª, item 9 (p. 8):** isenção do vendedor por declarações inexatas (ver achado 2).
- **Cl. 11ª, item 2.1 (p. 7) e cl. 5ª (p. 6):** impostos de transmissão, escrituras, registros e tributos dos imóveis correm por conta do vendedor. O valor líquido recebido é menor que o nominal.
- **Cl. 3ª, § 2º (p. 6):** "sem direito de preferência dos demais sócios", "não concorrência com pessoas alheias e terceiras" e "garantia de confidencialidade". O contrato foi entregue ao Dr. Pablo apesar da cláusula de confidencialidade.
- **Cl. 3ª:** quotas "livres e desembaraçadas".
- **Cl. 11ª, item 2:** a frase "as COMPRADORAS não dispõem de condições de fazê-lo através de moeda corrente" foi citada só parcialmente. É precedente de pagamento de saída em lotes.

**Correção sugerida:** acrescentar ao § 5 um bloco "Outras cláusulas", com página.

### 8. MÉDIA — A identificação da terra de São Brás e o "crédito hereditário" estão afirmados com mais certeza do que a fonte permite (§ 4 e 8.4)

**Evidência:**
- O contrato fala em "créditos hereditários **e/ou creditórios**" e cita São Brás como exemplo ("dentre outros bens móveis ou imóveis"). Não diz que a parte do preço é herança, nem de quem.
- Inventário 08, item 08 (4,5 ha, R$ 250.000): foi "adquirida do Senhor Tercílio Dal Más … CONTRATO DE COMPRA E VENDA, datado de 25 setembro de 2008". Está "ao lado das terras na Capela de São Brás". Não é bem herdado.
- Só o item 07 (69,8 ha, R$ 3.500.000, Capela de São Brás) é candidato a herança.
- A SEI tem outra área em Cotiporã na mesma Linha Júlio de Castilhos: lote 33, 15,1 ha, R$ 40.000 (2025, T9 item 09). Em 2018 estava na Agropecuária.

**Correção sugerida:**
- § 4: "corresponde provavelmente ao item 07 (69,8 ha, R$ 3.500.000); o item 08 (4,5 ha, R$ 250.000) é contíguo e foi comprado em 2008".
- 8.4: "pode pagar direitos do Rovílio como herdeiro ou credor".

### 9. MÉDIA — Falta registrar a assimetria do Ipê II na SEI

**Evidência:**
- A SEI deduz R$ 1.650.000,00 (3 lotes do Ipê II a R$ 550.000 cada) como compromisso.
- O Ipê II não está valorado como ativo na SEI 2025. Na versão antiga era o item 15 de "loteamentos contratados", com "R$ 00.000.000,00", e foi retirado da versão atual (diff, linha 310).
- Arq. 1, p. 29, item 18.3: no Ipê II, "tendo em vista os acertos e comissões no acordo primitivo da contratação do empreendimento, o percentual resultante foi de 36%", e "a empresa estima prejuízo".
- A comissão do Rovílio (itens 7 e 8) é exatamente pela "intermediação … do Condomínio Morada do Ipê I e II".

**Correção sugerida:** acrescentar ao § 9 (SEI), só como registro:
- os lotes do Ipê II entram como passivo a R$ 550.000 cada, sem ativo correspondente no inventário;
- o grupo liga os "acertos e comissões" ao percentual reduzido do Ipê II.

### 10. MÉDIA — As datas de baixa de 09 e 10 não aparecem, e essas empresas figuram como compradoras

**Evidência:**
- Inventários 2025, doc 09: "EMPRESA BAIXADA EM 04 DE SETEMBRO DE 2023". Doc 10: "EMPRESA BAIXADA EM 08 DE ABRIL DE 2024".
- As duas são compradoras no contrato de 28/08/2025, com CNPJ e IE, representadas por Élcio.

**Correção sugerida:** no § 3 e no § 9 (S4), registrar as datas e o fato de que empresas já baixadas figuram como compradoras, sem concluir.

### 11. BAIXA — A conferência de 2018 está incompleta e falta a diferença entre o quadro formal e o rateio

**Evidência:**
- Além de 12 e 15, os inventários 2018 também trazem Rovílio com 2 % na SEI (18.1, T1 e T37) e com 1 % na Turismo (8.1) e na Siviero Transportes (7.1). Ele assina todos.
- Em 2018, o valor de cada sócio está em branco para todos ("2,00% = R$").
- O "QUADRO SOCIETÁRIO REAL" (T1) de 2018 e de 2025 mostra 18,18 % × 5, Pedro 7,10 % e Rovílio 2,00 %. O rateio de valor e o contrato usam 16,33 % (2018: 16,333 %).

**Correção sugerida:** completar a frase do § 3. Anotar a diferença entre o quadro formal e o rateio, que o grupo explica por "acordo verbal" (arq. 1, p. 23).

### 12. BAIXA — A nota da planilha também erra o sentido, não só o número

**Evidência:**
- GERAL H874: "A avaliação que eles fizeram está 36.53% abaixo do acerto com Rovílio".
- Além do 36,53 × 35,53, a base de R$ 2.538.208,55 está **26,22 %** abaixo do preço. O preço é que está 35,53 % acima da base.

**Correção sugerida:** registrar as duas coisas no § 7.

### 13. BAIXA — O item 9 como "reembolso" é leitura, não texto

**Evidência:** o contrato qualifica o item 9 como "antecipação" que "correspondem ao bem imóvel de propriedade de VENDEDOR (apartamento) dado como forma de pagamento para a construção do muro". "Reembolso" (8.1) é interpretação plausível, mas não é o termo do contrato.

**Correção sugerida:** "o item 9 corresponde a um bem do vendedor usado numa obra do condomínio (o contrato o chama de antecipação)".

### 14. BAIXA — Coerência da base nas contas sem comissão e linhas sem percentual no § 7

**Evidência:**
- Ao retirar a comissão do preço, a base continua descontando os R$ 550.000,00 de comissão (item 8) no líquido da SEI.
  - Com a base somando de volta esses R$ 550 mil: 96,42 %.
  - Somando de volta o R$ 1,65 mi inteiro: 95,60 %.
- As linhas 3 e 4 da tabela mostram "—". Os valores seriam:
  - 133,79 %, com a SEI antes do "à pagar";
  - 96,18 %, com a SEI pelos bens brutos. Esse é um argumento que o outro lado pode usar: o preço fica abaixo de 2 % dos bens brutos.

**Correção sugerida:** preencher as linhas 3 e 4 e anotar a sensibilidade da base.

### 15. BAIXA — Detalhes da tabela do § 5

- **Item 6:** a data de 25/10/2013 é só do 6.1, que também diz "repassado o valor para o VENDEDOR". A venda original do lote 15/q252, com o valor entregue a Ivo Siviero, não tem data.
- **Itens 7 e 10:** o contrato diz "já entregue/pago anteriormente ao presente ato". Sem data, mas antes de 28/08/2025. Vale escrever isso no lugar de só "sem data".

### 16. BAIXA — Inconsistências internas do inventário da SEI que o § 6 poderia registrar

**Evidência:**
- O item 10 aparece como "R$ 25.500,00 (Vinte e cinco mil reais), pagos em espécie". O valor por extenso não confere, e o contrato diz que são impostos pagos pela SEI, não dinheiro entregue.
- O texto diz que as quotas "foram devidamente acertadas, anteriormente ao presente ato", mas lista R$ 1.650.000,00 "À PAGAR".

### 17. BAIXA — Fonte do argumento do Dr. Pablo e testemunha

**Evidência:**
- O cálculo dos 35,53 % está no arq. 1, p. 21 (item 17, "Dúvida: Distrato com Rovílio"). Ele traz também a diferença implícita de R$ 45.093.672,40 (172.004.100 − 126.910.427,60). O documento lista a p. 21 nas fontes, mas não a cita no texto.
- A testemunha Luciano Bertocco tem o mesmo nome do autor de documentos de 2018 do grupo (arquivo 2018 "20. 1. … Elaborados Luciano Bertocco"). Vale um registro neutro.

---

## Conferido sem erro

**Partes e assinaturas**
- Vendedor, CPF 225.676.150-91, cônjuge "Ciente/Anuente" Maristela Pretto Siviero (p. 1 e 9).
- Representantes das cinco compradoras: Pedro (01 e SEI), Luciane (08), Élcio (09 e 10).
- Data e local: Veranópolis, 28/08/2025. Firmas reconhecidas de 17/09 a 24/09/2025. Testemunhas Suzane e Luciano Bertocco.

**Cláusulas e citações**
- Cl. 1ª (p. 2), objeto e citação.
- Cl. 3ª (p. 4), citação integral conferida na imagem.
- § 2º da cl. 3ª (p. 6), "diluídas de forma igualitária".
- Cl. 4ª (p. 6).
- Parágrafo primeiro (p. 6).
- Cl. 10ª, multa de 10 %.
- Cl. 11ª, itens 1 e 2 ("porteira fechada" e dação).

**Quotas e percentuais**
- Percentuais das p. 3–4 e 6: 2 % em 01, 08 e SEI; 1 % em 09 e 10. Os demais sócios têm 16,33 % e 16,50 %.
- Somas de 99,98 % e 100,00 %.

**Itens de pagamento**
- Os 10 itens: valores, lotes, quadras, loteamentos, cidades, naturezas (antecipação e comissão) e datas (24/03/2023, fev./2025, 01/11/2012, 25/10/2013, maio/2022).
- Soma de R$ 3.440.082,00 igual ao preço.
- Pagos antes de 31.12.2024: R$ 745.000,00 (itens 1, 5, 6 e 9).
- Item 2: R$ 587.510,00 em fev./2025.
- A entregar: R$ 1.650.000,00 (itens 3, 4 e 8).
- Comissão: R$ 982.072,00.

**Contas (§ 7)**
- 2 % de 126.910.427,60 = 2.538.208,55.
- 135,53 % e diferença de +901.873,45.
- 2.571.208,55; 3.576.653,21; 2.458.010,00 (96,84 %); 2.348.010,00 (92,51 %).
- Valores implícitos de 172.004.100 e 122.900.500.

**Inventário 2025 da SEI**
- Bens: 152.796.260,60.
- Cadeia de cálculo: 152.796.260,60 − 15.754.581 = 137.041.679,60; − 16.240.000 − 3.800.000 − 16.127.652 = 100.874.027,60.
- "Demais compromissos" de R$ 16.127.652,00 = soma exata das 16 linhas de valor das observações. A primeira é o "à pagar" de R$ 1.650.000,00.
- Lotes 13/q02, 10/q02 e 07/q04 "PAGO"; R$ 110.000 e R$ 25.500 "pagos em espécie"; 3 lotes do Ipê II "À PAGAR".

**Inventários 01, 08 e 00**
- Totais de 01 (9.661.414,00) e 08 (16.374.986,00).
- Na versão antiga, parcelas de Rovílio de 193.228,24 e 327.499,70. Na atual, rateio em 6 × 16,66 %.
- Marca "x2" no doc 00 ao lado de 01, 08 e 11.

**São Brás**
- Itens 07 e 08 da 08: 69,8 ha + 4,5 ha = R$ 3.750.000.
- Em 2018, Agropecuária, "NÃO DIVISÍVEL".

**Planilha Google**
- GERAL E874 = −550000*3. Resumo C63 = "Diferença de 36.53%". Resumo B67 e GERAL H874 com os valores do Dr. Pablo.

**Arq. 1**
- Citações das p. 23 e 29, nas partes transcritas.

**Hashes**
- MD5 (10 primeiros dígitos) das fontes do § 1.

**Script**
- `rovilio_numbers.py` roda e reproduz todos os números. Ele confere a soma dos itens com o preço por assert e soma as 16 linhas pelos parágrafos das observações.
