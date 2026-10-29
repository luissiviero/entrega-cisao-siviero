# Revisão independente — pacote 3 (destino 2018 → 2025 e baixas; empresas 03, 04, 07, 09, 10)

Revisor de contexto novo, 28/10/2026. Objeto: `pacotes/pacote-03.md`. Conferido contra as fontes primárias, não contra o pacote: inventários 2018 (Drive, arq. 7.1, 8.1, 9.1, 11.1, 14.1, 18.1, 22.1, 23.1, 25.1, 27.1), inventários de 2015 (`Material antigo/2018_desatualizado`, incl. "A-PLANILHA… 2015"), inventários 2025 atuais e antigos (03, 04, 07, 09, 10, 11, 01, 06, 08, 00), contas pagas (planilha `Contas_Pagas_2023-07_a_2024-10.xlsx`, aba Consolidado, e PDFs de jul/2023 p. 2 e 5, jan/2024 p. 2 e 4, fev/2024 p. 4, jul/2024 p. 3, set/2024 p. 3, out/2024 p. 3, lidos na imagem), arq. 1 (texto), arq. 13 (OCR, só para localizar), arq. 17 (PDF p. 3 e planilha), PDFs de Rafael 03/09/10, `planilha-mestre` (abas Mestre e "Créditos e dívidas (notas)") e `pacote-02.md`. Numeração de parágrafos: parágrafos não vazios do .docx, como em `_scripts/fase0_extracao.py`. Scripts e imagens: `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\639aa2be-df0b-40be-bb0a-3182601981fe\scratchpad\revisor-p3\`.

**Contagem: 1 alta, 8 médias, 13 baixas.**

---

## Alta

### A1. Restituição de ICMS (item 9 e pergunta 13): é tratada como pendência, mas as fontes dão o destino (recebida antes de 2018)
- Pacote, item 9: "Restituição de ICMS (2018, valor em branco…): sem notícia depois. Se houve recebimento, é ativo; pendência." Pergunta 13: "desfecho da ação, valor recebido e empresa".
- Inventário 2018 da Turismo (arq. 8.1), par. 17, no passado: "Conforme uma redução da alíquota do ICMS buscada via esfera judicial, houve uma restituição no valor de aproximadamente R$ ……, que retornaram aos cofres da empresa."
- Doc 25 de 2018, par. 243 (seção da Siviero Transportes): "todo o crédito buscado junto a Justiça, do tributo do ICMS, … no valor de aproximadamente R$ 450.000,00 foram usados administrativamente para o funcionamento da empresa … Este valor retornou ao caixa da empresa de forma líquida". O par. 244 trata o valor como já recebido ("Se este valor não tivesse retornado…"). A versão de 2015 (A-PLANILHA 2015, par. 169) traz o mesmo texto com "aproximadamente R$ 400.000,00".
- Conclusão: pelo que o próprio grupo declara, o crédito de ICMS foi recebido e gasto antes de 2018. Não é um ativo de 2018 que precise de destino (P4). O "valor em branco" de 2018 é o valor da restituição já recebida pela Turismo, não um crédito em aberto.
- Correção: reclassificar no item 9 como "destino declarado pelo grupo: restituição recebida antes de 2018 e usada na operação (Turismo 2018 par. 17; doc 25 par. 243–244)". Tirar o ICMS da pergunta 13, que fica só com a quitação dos parcelamentos. Se ainda quiser, perguntar apenas o valor recebido pela Turismo. Atualizar a nota 09-2018-N17 da planilha-mestre (hoje "Pendente").

---

## Média

### M1. Acordo do IXO 6000: o pacote o chama de "crédito", mas a fonte não diz quem paga; "credor" e faixa de saldo sem base
- Fonte (inventário 2025 da ST, par. 22): "Existe pendente um PROCESSO JUDICIAL referente a venda do Ônibus IXO 6000, acordado pelo valor de R$ 115.000,00, sendo pagos de forma parcelada: R$ 20.000,00 (Despesas judiciais e advocatícias) e o saldo em parcelas de R$ 5.000,00, cada, a contar de novembro de 2024." O texto não diz se a ST recebe ou paga. Compare com o inventário 2025 da Transportadora, par. 15, que diz expressamente "PROCESSO JUDICIAL em tramitação **para cobrança**". A própria planilha-mestre classifica a nota 10-2025-N22 como "dívida, crédito/processo". O parágrafo vem logo depois de "O valor real acima apontado não é líquido, à medida que:" (par. 19–21). É um indício fraco de que seja um ônus, porque a ordem das observações varia entre os inventários.
- O pacote afirma a natureza como fato: resumo ("crédito do IXO 6000 em 31.12.2024: pendência"), item 1 ("Se o crédito existia…"), pergunta 2 ("credor após a baixa… valor recebido") e S4-DB2 ("crédito em aberto"). A leitura de credor é plausível: venda a prazo em dez/2015 e custas de "AVALIACAO E VENDA IMOVEL PENHORA" pagas em nome da ST em jul/2024, que normalmente cabem ao exequente. Mesmo assim, é inferência e não está marcada.
- Mais três imprecisões no mesmo ponto:
  - "nenhum inventário de 2025 registra o crédito": o inventário 2025 da ST registra o processo e o acordo no par. 22, só que fora da tabela e sem valor no total.
  - A faixa "saldo entre R$ 85.000 e R$ 115.000" inclui os R$ 20.000 que a fonte rotula "Despesas judiciais e advocatícias". Se forem honorários e custas, o líquido para o grupo seria no máximo R$ 95.000.
  - "o credor em 31.12.2024 não está identificado": falta a leitura jurídica óbvia. Extinta a sociedade, quem sucede nos direitos e obrigações remanescentes são os sócios. Pelo contrato, são os do quadro "NÃO REAL": Adélia Frizon Siviero com 98 % e Giovani Siviero com 2 % (inventário 2025 da ST, tab. 2). Não são os seis do quadro "real". Isso importa para o 1/6.
- Correção: marcar "crédito" como inferência (provável: ST credora), citar a leitura alternativa (ST ou grupo devedor) e reescrever a pergunta 2: "No acordo do IXO 6000, quem paga e quem recebe? Se o grupo recebe, quem é o titular depois da baixa (sócios formais Adélia/Giovani ou outro) e quanto foi recebido até 31.12.2024; os R$ 20.000 são de quem?". Ajustar o S4-DB2 no mesmo sentido.

### M2. Ceval faturou cerca de R$ 446 mil em 2023–2024 e tinha folha de pagamento: caixa, recebíveis e passivos de 31.12.2024 não aparecem
- Contas pagas, aba Consolidado, linhas pagas pela SEI:
  - ago/2023: "CEVAL-ISSQN S/NOTA R$ 27.200,00-LUIZ ARGENTA" e PIS/COFINS/IRPJ/CSLL sobre R$ 27.200.
  - ago/2024: "IMPOSTO S/SERVICOS-N.F.20-CERAN-R$ 126.590,00" (ISSQN e INSS).
  - set/2024 (PDF p. 3, conferido): "CEVAL - IMPOSTO R$ 126.590,0O X 0,65% - NOTA", "… X 3,00%", e na p. 4 IRPJ e CSLL sobre o mesmo valor.
  - set/2024: "N.F.22-CERAN-R$ 286.940,00". out/2024: "N.F.21-RENAN BARONIO-R$ 3.[152]" (ISS de R$ 63,04).
  - out/2024 (PDF p. 3, conferido): "CEVAL-IMPOSTO R$ 290.092,00X0,65%-NOTAS 21/22" (286.940 + 3.152 = 290.092) e "CEVAL - IMPOSTO R$ 2.362,50 X 4,80% - NOTA 23".
  - Soma das bases: R$ 446.244,50.
  - Folha: "CEVAL-FOLHA PAGAMENTO MES AGOSTO/2024-JOAO C…" (FGTS, 21/08/2024) e "CEVAL - SOBRE SALARIO FOLHA PAGAMENTO MES SET" (18/10/2024, R$ 199,70). Também alvará 2024, e-CNPJ, PPRA/PCMSO 2024 e 2025 e contabilidade mensal, tudo pago pela SEI.
- Ligação de NF 20/22 com a Ceval: é inferência minha, porque as bases da Ceval no PIS/COFINS coincidem (R$ 126.590; R$ 290.092).
- O inventário 2025 da Ceval lista só três máquinas. Não há caixa, contas a receber da CERAN, passivo trabalhista nem saldo com a SEI, que pagou os tributos e as despesas. O pacote não trata o assunto, embora a pergunta 4 da revisão peça exatamente isto. Também reforça o item 4: a Ceval estava operando em 2024.
- Correção: novo item (Ceval: atividade em 2023–2024; destino das receitas; saldo intragrupo com a SEI em 31.12.2024) e uma pergunta, ou candidato intragrupo na Parte 6. Sem valor adotado (P5).

### M3. Dívidas da Siviero Transportes em 2018 sem destino no pacote (Millenium, Santander, GUIBOR)
- Doc 25 de 2018, par. 225–228: "Ainda restam: 04 parcelas para serem pagas para a empresa MILLENIUN, de R$ 9.200,00, cada. 18 parcelas para serem pagas para a empresa GUIBOR, de R$ 8.000,00, cada, do ônibus atual. 09 parcelas para serem pagas para o Banco SANTANDER, de R$ 7.000,00, cada." São R$ 36.800 + R$ 144.000 + R$ 63.000 = R$ 243.800 de passivo da ST em 2018.
- Doc 22 (reunião), par. 49: "Tem valor para acertar com a Millenium."
- O pacote só usa a GUIBOR, como financiamento do IUY 8068 (item 2). Millenium e Santander não aparecem, e nenhum inventário de 2025 os cita.
- Correção: incluir no item 3 (ou num item próprio) como passivos de 2018 a confirmar (quitação, quem pagou) e perguntar junto com a pergunta 4/5.

### M4. Doc 18 (SEI 2018, tab. 28): a citação omite "em troca de aluguéis", e o "não batem" esconde outra leitura
- Fonte, item 02: "Pago combustíveis da Siviero Transportes Ltda., **em troca de aluguéis no Posto até 2011**. | R$ 180.000,00". O pacote resume como "R$ 180.000 de combustíveis" e soma R$ 942.000 como valor da SEI (resumo nº 3; S4-5). O doc 23 (2011), par. 9, lista "Aluguel Posto … R$ 180.000,00" entre as dívidas da TS/ST que compõem os R$ 1.733.700. É o mesmo valor, e a fonte diz que foi compensado com aluguéis. Não é crédito adicional.
- "Números que não batem com os R$ 269.300": não são a mesma grandeza. Os R$ 269.300 são o que a SEI pagou "a partir de OUTUBRO de 2011" de despesas correntes (texto repetido desde o inventário de 2015). Os R$ 762.000 são dívidas pagas de 2011 a 2015 e podem incluir a quitação dos R$ 1.733.700. Essa leitura alternativa não está no pacote.
- Doc 22, par. 61, chama o aporte de "injeção financeira de aproximadamente R$ 2.000.000,00 para pagamento de dívidas". Reforça que o grupo tratou os pagamentos como aporte, não como empréstimo, e sustenta o "efeito no 1/6 tende a zero". O pacote não cita.
- Escopo: os créditos da Rádio (R$ 25.000), da SEI (R$ 11.000) e do Hotel Araçá (R$ 161.500) são de 2011 (doc 23). Nenhum inventário de 2018 dessas empresas os registra (conferido: Rádio, Hotel Araçá e SEI de 2018 não têm crédito contra a ST fora da "Planilha de investimentos"). Não são itens 2018 → 2025 no sentido do P4.
- Correção: citar o item 02 por inteiro, tirar os R$ 180.000 da soma (ou marcar "compensado com aluguéis, segundo a fonte"), apresentar 269.300 × 762.000 como grandezas diferentes, citar o doc 22 par. 61 e dizer que os créditos de 2011 já não estavam nos inventários de 2018.

### M5. Empréstimo Caixa "Rádio" (item 3 e S4-11): "pagou à ST" não decorre da fonte; leitura alternativa omitida
- Fonte (jul/2023, p. 5, conferido): na seção da Abastecedora, Fornecedor "SIVIERO TRANSPORTES LTDA", 10/07/23, R$ 6.500,00, "P. 28/28 - EMPRESTIMO CAIXA FEDERAL - RADIO C", e "ABA -TARIFA TED-PAGTO EMPRESTIMO RADIO-SIVIER".
- Nesses relatórios, a coluna "Fornecedor" não é necessariamente quem recebe. Em jul/2023, p. 2 (conferido), há "SIVIERO TRANSPORTES LTDA | 373,00 | MERCATO CONTABILIDADE - SERVICOS JUNHO/23-TUR": o prestador é a Mercato, e a ST aparece como empresa à qual a despesa é atribuída. O próprio pacote lê assim a linha da penhora ("em nome da 'SIVIERO TRANSPORTES LTDA'", item 1), mas lê a do empréstimo como "pagou à ST". As duas leituras se contradizem.
- Leitura alternativa: o empréstimo era da própria ST (ou da Rádio com a ST), e a Abastecedora pagou a última parcela por conta dela. A ST já tinha empréstimo na Caixa em 2011 (doc 23, par. 13: "Caixa Federal – 13 x R$ 4.590,00 … R$ 62.000,00"). "RADIO C…" e "RADIO-SIVIER…" estão cortados na fonte.
- Correção: escrever "a Abastecedora pagou, atribuído à ST no relatório, a parcela 28/28 de um empréstimo da Caixa com referência à 'RADIO C…' (texto cortado)", marcar devedor e beneficiário como inferência e deixar claro que o empréstimo bancário terminou em jul/2023. Fica só a eventual conta entre Abastecedora, ST e Rádio, de valor desconhecido: os R$ 182.000 supõem 28 parcelas iguais, todas pagas pela Abastecedora. Ajustar a redação do S4-11.

### M6. Arquivo 17 (p. 3) documenta a troca do lote do Vicentino pelos 3 ônibus, e não é usado
- Arq. 17 (PDF p. 3, texto conferido): "26.02.2014 | 04 | 1302 | Alcides Basso Imóveis. | Vendido, com permuta de 03 veículos – ônibus + R$ 552.000,00 em 60 parcelas, pagas pela Siviero = R$ 1.427.000,00."
- Isso documenta a origem citada no item 1 e no S4-7: o lote da SEI saiu em 26/02/2014, em troca dos 3 ônibus que ficaram com a ST. Também diz que as 60 parcelas de R$ 9.200 (R$ 552.000) foram "pagas pela Siviero". Se "Siviero" for a SEI (é como os documentos do grupo a chamam), há mais um fluxo SEI → ST. Isso é inferência, porque "Siviero" pode ser a própria ST. As 4 parcelas Millenium em aberto em 2018 (M3) são o fim dessas 60.
- O pacote diz que leu o arquivo 17 (item 2), mas não cita esta linha.
- Correção: citar no item 1 e no S4-7 (data, comprador, composição) e pôr a pergunta sobre quem pagou as 60 parcelas junto com a pergunta 5.

### M7. Perguntas ao grupo: falta a do item 10, há duplicatas com o pacote 2 e a 2 e a 13 precisam de ajuste
- O item 10 pergunta "o que eram os R$ 10.000 de 2018" e "se a Associação chegou a obter outorga de rádio comunitária". Nenhuma das 13 perguntas cobre isso: a 11 pede só certidões e distratos.
- Duplicatas com as 7 do pacote 2 (`pacote-02.md`, linhas 92–98), que o pacote 3 propõe juntar numa lista única de 20:
  - nº 9 (Celito) = pacote 2 nº 4;
  - nº 7 (ILD 3959) = pacote 2 nº 7, que ainda pergunta "qual é a terceira caçamba";
  - nº 10 (CRLV da Toyota) se sobrepõe à pacote 2 nº 3.
  - Juntadas, ficariam 17 perguntas distintas, não 20.
- A pergunta 2 pressupõe que o grupo é credor (M1). A 13 inclui o ICMS, que já tem destino (A1).
- Correção: incluir a pergunta do item 10, eliminar as duplicatas na lista única, reescrever a 2 e a 13 e acrescentar as de M2, M3 e M8.

### M8. Falta a pergunta que fecharia vários pontos de uma vez: os balanços contábeis
- Os inventários de 2025 dizem que o balanço existe: "O BALANÇO PATRIMONIAL E CONTÁBIL ESTÁ PRONTO E CONFERIDO, PELA ESCRITA CONTÁBIL NESSA DATA, QUE NÃO CORRESPONDE À REALIDADE FÁTICA DA EMPRESA." Está na ST, par. 14; na Turismo, par. 20; na Ceval, par. 17; na Rádio, par. 19.
- O balanço de 31.12.2024 da Ceval e da Rádio, e os balanços de encerramento e os distratos da ST e da Turismo, responderiam às perguntas 1, 2, 5, 6, 8 e 9 e ao M2/M3: recebíveis, empréstimos, contas entre empresas e destino dos ativos e passivos remanescentes na extinção.
- Correção: acrescentar uma pergunta: "balanço/balancete de 31.12.2024 da Ceval e da Rádio; balanço de encerramento e distrato da ST e da Turismo, com a destinação de ativos e passivos remanescentes". Isso pode substituir parte da pergunta 11.

---

## Baixa

### B1. Penhora de jul/2024: outras leituras não ditas
O pacote marca a ligação com o IXO 6000 como "possivelmente", o que está correto. Fica sem menção que a ST também tinha outros recebíveis em 2018 (24 × R$ 10.000 "empresa de Santa Catarina"; R$ 100.000 "do Rodrigo") e um processo em que, segundo o inventário de 2015 (arq. 8, par. 18), "a empresa e demais passaram de credores/autores para RÉUS" (Operação Deus tá vendo). Conferido na imagem (jul/2024, p. 3): "SIVIERO TRANSPORTES LTDA | 24/07/24 | 328,49 | CUSTAS AVALIACAO E VENDA IMOVEL PENHORA ALIJH". Sugestão: citar as três possibilidades na pergunta 3.

### B2. Identificações feitas por inferência sem marca
- Doc 25 par. 249 não cita a placa ("foi vendido um ônibus, adquirido da empresa Millenium, pelo valor de R$ 585.000,00"). A placa IXO 6000 vem do inventário de 2018 (mesmo valor e mesma data).
- "Doc 25 par. 227: 18 parcelas de R$ 8.000 à GUIBOR 'do ônibus atual'" ligado ao IUY 8068 (item 2): o par. 227 não diz qual ônibus. O financiamento do Santander (par. 228) e as "parcelas do micro-ônibus" de 2014 (par. 239) também podem ser do Volare.
- Toyota de R$ 40.000 em 2011 (item 7): o doc 23, par. 40, diz só "Toyota".
- Sugestão: marcar como inferência.

### B3. S4-3: a proposta faz sentido, mas o rótulo deve preservar a pergunta
Conferido na imagem: "IPVA ANO 2024 - IKI 4827 - RETRO CASE ANO 200…" (18/01/24, R$ 99,65, jan/2024 p. 4) e "IPVA ANO 2024 - IKI 4325 - RETRO CASE - ANO 2…" (26/02/24, R$ 104,36, fev/2024 p. 4), ambos na seção da Transportadora. Duas placas com licenciamento pago é indício forte de dois cadastros. Não prova que as duas máquinas existiam em 31.12.2024: um cadastro sem baixa continua gerando cobrança. Também não mostra o titular, porque quem pagou foi a Transportadora, não a Ceval nem a SEI. Sugestão: "indício forte de duas máquinas (dois cadastros com licenciamento em 2024); manter a pergunta 2 do pacote 2 (chassi/NF da IKI 4827)" em vez de "verificado".

### B4. S4-6 / item 5: "origem confirmada" é forte demais; faltam dois elementos que ajudam
- O doc 25 (par. 201–202) confirma que a sala foi comprada com um lote da SEI. Que a "DÍVIDA: 01 LOTE…" de 2018 seja esse lote é leitura do pacote, corretamente marcada como "Leitura" no item 5, mas chamada de "origem confirmada" na seção da Parte 6. Sugestão: "origem provável".
- Leitura alternativa não dita: o lote ainda não teria sido entregue a Leonir e Cleci Dal Pai, vendedores da sala (arq. 13, R.4/11.382). Nesse caso o credor seria um terceiro.
- O valor (R$ 110.000) difere do valor da permuta (R$ 70.000) e não há explicação no pacote.
- Elementos a acrescentar:
  - Doc 22, par. 50: "Lote para sala da Rádio" (pendência listada na reunião).
  - Arq. 17: o Colina do Retiro I tinha 23 lotes na SEI em 2018 (doc 18, tab. 11, item 46) e 13 em 2025 (doc 11, tab. 17, item 07). O arq. 17 lista exatamente as 10 saídas, nenhuma para a Rádio ou para Dal Pai. Logo, nenhum lote saiu da SEI para quitar essa dívida entre 2018 e 2025.
- Citação do par. 202 cortada sem reticências: a fonte continua "…pela própria empresa Rádio Comunidade Princesa dos Vales FM Ltda.".

### B5. Celito: um indício útil não usado
Nas contas pagas de jul/2023 a out/2024 aparecem juros pagos pela SEI a credores pessoas físicas: Samir e Inês Piccoli, Luis Carlos Mercali, André Luiz Siviero. Não há nenhum pagamento a Celito, embora a NP tivesse "juros de forma mensal" (doc 25 par. 144) a 1,50 % ao mês (doc 18, tab. 23, item 07). É indício de que a dívida não era mais servida em 2023–2024 (inferência). Vale citar no item 6 e na pergunta. Também em 2011 Celito era credor da TS/ST em R$ 25.000 (doc 23 par. 5). É anterior a 2018, mas convém registrar.

### B6. Turismo: dois indícios não citados
- Indício da baixa: "MERCATO CONTABILIDADE-SERV SET/23-TURISMO SIV", R$ 796,00, pago em 16/10/2023 (Consolidado, out/2023 p. 2). É o dobro do honorário habitual de R$ 373 e o último da Turismo. Mesmo padrão do "MARCO/24-DIST…" (R$ 796) da ST que o pacote usa.
- Indício da quitação do parcelamento: doc 25, par. 211: "Não pode ser dado baixa da mesma nos órgãos federativos, sem antes quitar todas suas dívidas de parcelamento de impostos." Se a baixa de 04/09/2023 se confirmar, o parcelamento terá sido quitado (inferência).

### B7. Linhas do DAER da ST não tratadas
Doc 25, par. 238: "Com relação das duas linhas do DAER…"; par. 237: transporte de estudantes da ACAUVE. Permissões de linha podem ter valor e não aparecem em 2018 nem em 2025. O texto parece herdado de 2015, então a provável conclusão é "sem valor/encerradas". Sugestão: uma linha no item 8 ou uma subpergunta na 11.

### B8. Associação: contradição na fonte e natureza jurídica
O inventário 2025 diz no par. 19 "O BALANÇO … ESTÁ PRONTO E CONFERIDO…" e no par. 20 "O BALANÇO NUNCA FOI FEITO". O pacote cita só o par. 20. Além disso, uma associação não distribui patrimônio aos associados na dissolução (vai a entidade sem fins lucrativos). Os R$ 10.000 de 2018 não teriam efeito no 1/6 em nenhuma hipótese. Sugestão: registrar as duas coisas e baixar a prioridade do item 10.

### B9. "Venda documentada" no item 1 = declaração do grupo
A única base é o inventário de 2018: não há contrato nem CRV. O item 9 já escreve "documentada pelo próprio grupo". Sugestão: usar a mesma qualificação no item 1 e lembrar que as vendas são anteriores a 2018.

### B10. Cobertura das contas pagas
Não há relatório de set/2023. A aba 2023-09 de `Contas_Pagas_e_Lotes_Vendidos…xlsx` diz: "Sem relatório de contas pagas para este mês na pasta Relatórios Mensais". As frases "nenhuma menção … contas pagas de jul/2023 a out/2024" (item 2) devem ressalvar essa lacuna.

### B11. Resumo nº 3: "R$ 2.003.000 (ST, 2011)"
Os R$ 269.300 que compõem o total são pagamentos "a partir de OUTUBRO de 2011". A data do total é posterior e incerta, porque o texto se repete em 2015 e em 2018. Sugestão: "(ST; dívida de set/2011 + pagamentos posteriores da SEI)".

### B12. Citações truncadas sem reticências
- Item 1, doc 25 par. 230: a fonte continua "…do Rodrigo, referente a venda do ônibus."
- Item 5: ver B4.
- Item 3: "a ST 'NÃO paga, desde 2011, qualquer conta… nem tampouco, os salários' de Luciane e Marcos" está correta com a elipse.

### B13. Arquivos dependentes a atualizar depois da aprovação
`parte6-candidatos.md` ainda não traz o S4-11, e o texto do S4-6 ("conferir de quem era o lote") e do S4-5 ("R$ 942.000 (SEI)", ver M4) continua o anterior. A planilha-mestre mantém "Pendente" na nota do ICMS (A1). É coerente com a regra de só gravar após a aprovação, mas convém listar no pacote o que muda em cada arquivo.

---

## O que conferi e está certo

- **Item 1:**
  - ST 2018, arq. 7.1, tab. 3, itens 05–07: valores, datas de venda ("dezembro de 2015/2016, à prazo") e "adquirido em troca de um lote no Loteamento Industrial Vicentino".
  - Totais: 1.515.100; 2.443.100 − 1.515.100 = 928.000; doc 27 = R$ 928.000.
  - Doc 25 par. 229–230 (24 × 10.000 + 100.000 = 340.000) e par. 249.
  - ST 2025 par. 22 (versão antiga par. 21): 115.000, 20.000, parcelas de 5.000 desde nov/2024; 19 parcelas.
  - Custas de penhora: R$ 328,49, 24/07/2024, jul/2024 p. 3, imagem.
  - O PDF de Rafael da ST não transcreve o processo e pede verificação da baixa (p. 2).
- **Item 2:** placas, anos e valores (250.000; 70.000 "carro que fazia a Linha de Cotiporã"; 588.000; total 908.000). Nenhuma menção às placas depois de 2018 em inventários de 2025, contas pagas, arq. 1 e arq. 17. Só aparecem nos inventários de 2015/2018 e nas planilhas de trabalho do Luís.
- **Item 3:**
  - ST 2018 par. 14–16 = doc 25 par. 234; doc 23 com os credores Rádio 25.000, SEI 11.000 e Hotel Araçá 23.000 + 50.000 + 88.500.
  - Parcela 28/28 de R$ 6.500 em 10/07/2023 e tarifa TED, na imagem.
  - Honorário "MARCO/24-DIST…" de R$ 796 em 15/04/2024.
  - Nenhum crédito contra a ST nos inventários de 2025 da SEI, Rádio, Hotel Araçá e Abastecedora.
- **Item 4:**
  - Ceval 2018 item 03 (chassi, R$ 93.000).
  - Multa de 04/07/2023 (R$ 156,19), conserto de 15/01/2024 (R$ 120) e "IPVA ANO 2024 - ILD 3959 - M.BENZ 2003 - CEVA" de 23/01/2024 (R$ 99,65), pagos pela SEI (jan/2024 p. 2, imagem).
  - Ausente do inventário 2025 atual e do antigo. Doc 25 par. 146 com "03 caçambas".
  - A ILD 3959 não aparece nas contas pagas depois de jan/2024 nem no arq. 17.
- **Item 5:** Rádio 2018 par. 14 (R$ 110.000; 2015 sem valor); doc 25 par. 201–202; arq. 13 R.4/11.382 (R$ 77.600, escritura de 15/03/2013); nenhuma menção em 2025 (Rádio e SEI).
- **Item 6:** Ceval 2018 par. 13; doc 18 tab. 23 item 07; doc 25 par. 144. Nenhuma menção a Celito nos inventários de 2025, inclusive no passivo da SEI de R$ 16.127.652 (tab. 50).
- **Item 7:** ST 2018 item 04 (R$ 20.000); Rádio 2025 item 07 (R$ 70.000); IPVA de 07/02/2024, R$ 104,36, pago pela Transportadora.
- **Item 8:** baixa em 08/04/2024 (par. 15); parcelamento de R$ 7.000 = 28 × 250; Operação Deus tá vendo em 2018 (par. 17) e em 2015 ("RÉUS", par. 18); quotas de Rovílio e o pedido de rescisão (par. 13, 16–17).
- **Item 9:** baixa em 04/09/2023 (par. 18); parcelamento de R$ 77.000 = 28 × 2.750 (2015: R$ 150.000); ônibus "2000" vendido em 2015 por R$ 95.000 (Turismo 2018 par. 14; doc 25 par. 215; doc 23 par. 31 "Paradiso ano 1995").
- **Item 10:** total de R$ 10.000 com linhas em branco (arq. 14.1; doc 27); capital 6 × 16.666,66 (2018) e 6 × 1.666,66 (2025); baixa em 09/07/2020 (par. 21); custos pagos pela Rádio (par. 14).
- **Certidões e baixas:** procurei por nome de arquivo e por texto (Drive Cisao_Siviero e 01_fontes). Não há certidão, distrato nem data de baixa diferente para 03, 09 e 10. Todas as fontes, inclusive os PDFs de Rafael, repetem 09/07/2020, 04/09/2023 e 08/04/2024. O arq. 1 não trata do assunto (a penhora citada lá, p. 23–24, é da SEI, caso Perini).
- **Parte 6:**
  - S4-3: os dados de IPVA conferem (ver B3).
  - S4-6: a origem do doc 25 confere (ver B4).
  - S4-11: os dados conferem (ver M5 sobre a redação).
  - Nenhum valor reavaliado (P5). Fatos de 2025–2026 (Rovílio) registrados sem mudar números (P2).
