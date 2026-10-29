# Revisão independente da planilha-mestre (empresas 03, 04, 07, 09 e 10)

Revisor independente (contexto novo), 28/10/2026. Objeto: `planilha-mestre-03-04-07-09-10.xlsx` (md5 73cdcd6243d31c84e9319a21e2d6e9e3), gerada por `_scripts/build_mestre.py` (md5 07d529f52cfc148f0bbcc6243eb5b9ee) a partir de `_extracao/fase0-extracao.json` (md5 9eecb6219ac91e13c23232cc7aa748da) e `decisoes.json` (md5 a526f3effad300a1596ba2db16aa3526). Os hashes conferem com o diário. Também revisado: `parte6-candidatos.md`.

Conferido contra as fontes primárias, lidas de novo por mim: inventários 2018 (Drive: 14.1 Associação, 11.1 Ceval, 9.1 Rádio, 8.1 Turismo, 7.1 Siviero Transportes; também docs 18, 23, 25 e 27), inventários 2025 (00 a 11), PDFs de Rafael (03, 04, 07, 09, 10, Diretrizes, Introdução e os 13 PDFs da SEI em `empreendimentos-pdf`), `ativos-avaliados-valter.xlsx`, `inventário-estudo-rafael.xlsx`, as duas planilhas Google (via extração) e o arq. 1 das respostas ao Dr. Pablo. Nenhum outro arquivo foi alterado. Scripts de conferência: `scratchpad\revisor-mestre\`.

Numeração de parágrafos: igual à da extração (só os parágrafos não vazios do .docx). Numeração de tabelas: ordem no .docx.

Resumo: **1 alta, 11 médias, 9 baixas.**

---

## Achados

### 1. ALTA: a linha "Móveis, instalações e utensílios" (07-2025-02) aparece como bem novo, mas em 2018 os móveis estavam dentro do item 02
- Evidência: o inventário 2018 da Rádio (arquivo 9.1, tab. 3, item 02, R$ 1.700.000) diz "Equipamentos instalados no ESTÚDIO, como: mesa, computadores, telefones, demais equipamentos eletrônicos e de transmissão, **móveis e utensílios**". Em 2025 esse conteúdo virou dois itens: o 02 (móveis, R$ 150.000) e o 03 (equipamentos e canal, R$ 1.900.000). O próprio `decisoes.json` (casamento 07-2018-02) já anota: "Em 2025 'móveis' sai do item e vira o item 02".
- O que está na planilha: na aba Mestre, linha 07-2025-02, aparece "Existe em 2018? = Não", "Par = sem par em 2018" e "Destino = Entrada depois de 2018 ou desmembramento". O par proposto 07-2018-02 → 07-2025-03 compara R$ 1.700.000 com R$ 1.900.000. O certo é comparar com R$ 2.050.000 (02 + 03).
- Correção: tratar o par proposto como 1 → 2, com 07-2018-02 ligado a 07-2025-02 e a 07-2025-03, ainda "PROPOSTO". Na linha 07-2025-02, marcar "Existe em 2018? = Sim, dentro do item 02 de 2018 (desmembramento, proposto)". Mostrar também a comparação conjunta (R$ 1.700.000 × R$ 2.050.000).

### 2. MÉDIA: o destino do terreno da antena (mat. 3.751) como "Transferido para a SEI" é contrariado pelas fontes
- Evidência: nas fontes, o terreno já era da SEI em 2018 e só aparecia **listado** no inventário da Rádio.
  - Doc 18 (SEI 2018), tab. 31, cujo título é "BENS DA EMPRESA SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. E RELACIONADOS NOS INVENTÁRIOS DE OUTRAS EMPRESAS", item 06: "Relacionado no Inventário da empresa: RÁDIO", R$ 7.148.016.
  - Doc 25 (2018), par. 204: "A antena … está instalada num imóvel que pertence à empresa Siviero Empreendimentos".
  - Inventário 2025 da Rádio, par. 15: "de propriedade da empresa Siviero Empreendimentos".
  - Arq. 1, p. 12: "A Siviero Empreendimentos … cede área para instalação da antena … de forma verbal".
- O que está na planilha: na aba Mestre, linha 07-2018-06, aparecem "Destino = Transferido para a SEI…" e "Situação = Em 2025 fora destas empresas", que o Leia-me define como "bem passou a outra empresa do grupo". A coluna Ref. 2018 cita só a Rádio e omite o doc 18, tab. 31.
- Efeito: o total 2018 da Rádio (R$ 9.316.016, também no doc 27) inclui R$ 7.148.016 de um bem que não era da Rádio. A planilha sugere uma saída de patrimônio da Rádio para a SEI entre 2018 e 2025 que as fontes não mostram.
- Correção:
  - Destino: "Sempre da SEI (doc 18, tab. 31); em 2018 só listado no inventário da Rádio; em 2025 no doc 11, tab. 7, item 02".
  - Situação própria, por exemplo "Bem de outra empresa listado em 2018".
  - Incluir o doc 18, tab. 31, na referência.
  - No Resumo, mostrar o total 2018 da Rádio com e sem o terreno.

### 3. MÉDIA: os 3 ônibus da Siviero Transportes já vendidos em 2018 aparecem como bens existentes, e o Resumo usa o total bruto como "Total declarado 2018"
- Evidência: o inventário 2018 da Siviero Transportes (arquivo 7.1), tab. 3, mostra:
  - item 05 (IXO 6000): "Vendido em dezembro de 2015, à prazo";
  - item 06 (IMY 5000): "Vendido em dezembro de 2016, à prazo";
  - item 07 (IOM 9093): "Vendido em dezembro de 2015, à prazo";
  - linha de total: "R$ 2.443.100,00 / (-) R$ 1.515.100,00 / R$ 928.000,00".

  O doc 27 (2018) também dá R$ 928.000 para a Siviero Transportes.
- O que está na planilha:
  - Mestre, linhas 10-2018-05/06/07: "Existe em 2018? = Sim", com os valores 585.000, 430.100 e 500.000 na coluna Valor 2018.
  - Resumo, linha da Siviero Transportes: "Total declarado 2018" = 2.443.100 e "Confere 2018? = sim". Isso esconde que o total declarado final é R$ 928.000 e que R$ 1.515.100 já não eram bens em 2018, e sim crédito a receber.
- Correção:
  - Na coluna "Existe em 2018?", usar "Não (vendido antes; crédito a receber)" ou uma situação própria.
  - No Resumo, mostrar bruto, dedução e líquido (2.443.100 − 1.515.100 = 928.000) e conferir com o doc 27.

### 4. MÉDIA: o destino dos ônibus vendidos a prazo não cita as evidências que existem, e a Siviero Transportes tem crédito em 31.12.2024 com total zero
- Evidência:
  - Inventário 2025 da Siviero Transportes, par. 22, que já está na aba de notas como 10-2025-N22: "PROCESSO JUDICIAL referente a venda do Ônibus IXO 6000, acordado pelo valor de R$ 115.000,00 … R$ 20.000,00 (Despesas judiciais e advocatícias) e o saldo em parcelas de R$ 5.000,00 … a contar de novembro de 2024". Logo, em 31.12.2024 havia um crédito pendente.
  - Doc 25, par. 229–230: "a empresa Siviero Transportes ainda tem a receber: 24 parcelas, no valor de R$ 10.000,00 … e mais R$ 100.000,00, do Rodrigo".
  - Doc 25, par. 249: ônibus "vendido … pelo valor de R$ 585.000,00, à prazo" (dez/2015).
  - Doc 22, par. 49: "Tem valor para acertar com a Millenium".
- O que está na planilha:
  - Mestre, linhas 10-2018-05/06/07: o destino é só "recebimento a verificar — pendente (P4)". A linha do IXO 6000 não liga a 10-2025-N22 e a coluna "Parte 6" dela está vazia.
  - Resumo, linha da Siviero Transportes: "Confere 2025? = sim" com R$ 0, sem dizer que existe esse crédito numa empresa baixada em 08/04/2024. O PDF de Rafael (10) também omite o parágrafo, como a Fase 0 já registrou.
- Correção:
  - No destino do IXO 6000, citar 10-2025-N22: acordo de R$ 115.000 contra o valor de R$ 585.000 de 2018, ou seja, perda e crédito residual a apurar.
  - Nos itens 05–07, citar o doc 25, par. 229–230 e 249.
  - Acrescentar no Resumo uma nota sobre o crédito fora da tabela de bens.

### 5. MÉDIA: o alerta "valor citado na descrição difere da coluna de valor" sumiu do Mestre
- Evidência: o inventário 2018 da Siviero Transportes mostra divergência entre o texto e a coluna de valor:
  - item 05 (IXO 6000): "valor comercial de R$ 470.000,00" no texto, R$ 585.000 na coluna;
  - item 07 (IOM 9093): R$ 450.000 no texto, R$ 500.000 na coluna.

  A extração (campo `alerta`) marca os dois. O doc 25, par. 249, fala em venda por R$ 585.000.
- O que está na planilha: as linhas 10-2018-05 e 10-2018-07 mostram só o valor da coluna, sem aviso. O alerta também existe para a sala (R$ 100.000 de aquisição contra R$ 120.000) e para o terreno (R$ 992,78/m²). Esses dois são esperados, mas também não aparecem.
- Correção: levar o campo `alerta` e o valor citado no texto para uma coluna "Observação" ou "Alerta" do Mestre.

### 6. MÉDIA: a Toyota (IHW 0285) aparece como "Transferido" sem documento da transferência; é bem C0 e candidato intragrupo
- Evidência:
  - A Toyota estava no inventário 2018 da Siviero Transportes (item 04, R$ 20.000; doc 23 de 2011 também a lista no "Capital em Veículos" da ST/TS, R$ 40.000).
  - Em 2025 está no inventário da Rádio (item 07, R$ 70.000).
  - A Siviero Transportes foi baixada em 08/04/2024 (inventário 2025, par. 15).
  - Nenhuma fonte registra venda, doação ou valor da transferência.
  - É o único bem da lista C0 de Rafael (Diretrizes p. 7; PDF da Rádio p. 11).
- O que está na planilha: Mestre, linha 07-2025-07, com "Destino = Transferido de 10 Siviero Transportes para 07 Rádio", "Registro de data-base" vazio e "Parte 6" vazia.
- Correção:
  - Destino: "Mesma placa (identidade confirmada); forma, data e contrapartida da transferência ST → Rádio sem documento — pendente".
  - Registrar quem era o titular em 31.12.2024 (registro do Detran ou CRLV).
  - Incluir a Toyota em `parte6-candidatos.md` como intragrupo: bem saiu da ST sem contrapartida registrada e a ST foi baixada.

### 7. MÉDIA: o "Critério de Rafael" do caminhão LK 1313 (04-2025-03) está errado
- Evidência: PDF da Ceval, p. 4, traz faixa de mercado de R$ 90.000 a R$ 150.000 e FIPE "em torno de R$ 30.600". No balanço executado (p. 5), Rafael mantém R$ 45.000, igual ao valor do grupo.
- O que está na planilha: Mestre, linha 04-2025-03, com "Critério de Rafael = faixa de mercado" e "Faixa (C1) = R$ 90.000 a R$ 150.000". Quem ler pode achar que Rafael aplicou a faixa; o valor dele está fora da própria faixa.
- Correção: usar o critério "manteve o valor do grupo apesar da faixa (FIPE R$ 30.600 citada)" e mostrar a FIPE junto da faixa. Isso pesa na aplicação de C0/C1 (regra P10).

### 8. MÉDIA: a coluna "Registro de data-base" está vazia em todas as 18 linhas, sem marca de pendência
- Evidência: Mestre, coluna AA, vazia.
- Por que importa: é justamente onde deveria aparecer:
  - quem era o titular registral em 31.12.2024 da Toyota (achado 6) e do terreno 3.751 (SEI, achado 2);
  - a sala e o box (matrículas 11.382 e 11.324, que o grupo disse estarem em nome da Rádio, arq. 1 p. 12);
  - as empresas baixadas (datas de baixa).
- Correção: preencher com a fonte, ou escrever "Pendente — sem documento" e explicar a coluna no Leia-me.

### 9. MÉDIA: faltam na aba "Créditos e dívidas (notas)" itens citados nos textos dos inventários
- (a) Inventário 2025 da Rádio, par. 15: "HAVENDO UM CUSTO PARA O DESLOCAMENTO DA ANTENA (DESMONTAGEM E MONTAGEM + CASA DE TRANSMISSÃO + INSTALAÇÃO ELÉTRICA, ETC) E O CUSTO PARA A AQUISIÇÃO DE UM TERRENO". É um ônus da Rádio, base do S4-1; hoje só entra a redução de 8 % (par. 17).
- (b) Associação 2018, par. 14, e 2025, par. 14: "Os custos financeiros necessários foram subsidiados pela empresa Rádio". É intragrupo (crédito da Rádio sem valor).
- (c) Siviero Transportes 2018, par. 15–16: a empresa "NÃO paga, desde 2011, qualquer conta … nem tampouco, os salários dos herdeiros e funcionários". É intragrupo; o doc 25, par. 248, diz que a Rádio teve "o ônus do pagamento dos salários", cerca de R$ 100.000.
- (d) Turismo 2025, par. 14, e Siviero Transportes 2025, par. 13: venda das quotas de Rovílio (28/08/2025). Pelo arq. 1, p. 23, Turismo e Siviero Transportes, embora baixadas, aparecem como compradoras, "quitando tudo, inclusive … créditos". Os par. 19 (Turismo) e 16 (ST) pedem "rescisão do Contrato Particular das Quotas Societárias". O fato é posterior à data-base, mas pode afetar o crédito do IXO 6000 (achado 4).
- Correção: incluir esses parágrafos com natureza "intragrupo/ônus/venda de quotas", destino "Pendente", e ligar (a) ao S4-1.

### 10. MÉDIA: o total de R$ 10.000 da Associação em 2018 não tem linha nem pendência de destino
- Evidência: inventário 2018 da Associação (arquivo 14.1), tab. 2: 5 linhas pontilhadas e "TOTAL GERAL DA ASSOCIAÇÃO … R$ 10.000,00". O doc 27 (2018) também dá R$ 10.000. Em 2025 o total é R$ 0.
- O que está na planilha: o Mestre não tem linha para isso. O Resumo mostra "Confere 2018? = NÃO" e o Leia-me só cita o fato. Pela regra do método, todo valor existente em 2018 e ausente em 2025 precisa de destino, ou vira pendência.
- Correção: criar uma linha "03-2018-total (sem itens)" com R$ 10.000 e destino "Sem destino documentado — pendente (P4)". A possível relação com o capital da Associação (R$ 10.000 em 2025) fica a verificar.

### 11. MÉDIA: as conferências do Resumo para 03, 09 e 10 dizem "Confere Rafael? sim", mas comparam com o "apresentado pelo grupo"
- Evidência: o script usa `apresentado_grupo` quando não há balanço de Valter. Os PDFs 03, 09 e 10 de Rafael não têm balanço próprio (só "Balanço apresentado pelo grupo" + texto de baixa). O PDF 03 omite a data de baixa e o PDF 10 omite o processo do IXO 6000.
- O Resumo também não confere com o doc 27 (2018) nem com o doc 00 (2025). Essa conferência teria mostrado a divergência da Siviero Transportes (achado 3).
- Correção:
  - Nessas três empresas, usar "Confere Rafael? = sem balanço próprio de Rafael (só repete o grupo)".
  - Acrescentar colunas "Doc 27 (2018)" e "Doc 00 (2025)" com a conferência.

### 12. MÉDIA: faltam candidatos evidentes na Parte 6
- (a) Intragrupo SEI → Siviero Transportes. O doc 25, par. 223, diz que foi "agregado ao capital da empresa … a venda de um lote industrial no Município de Farroupilha … valor de venda de mercado de R$ 1.300.000,00 … negociado … por R$ 875.000,00". O par. 224 trata da compra dos 3 ônibus da Millenium (R$ 1.427.000). O inventário 2018 da Siviero Transportes, itens 05–07, diz que cada ônibus foi "adquirido em troca de um lote no Loteamento Industrial Vicentino". Um lote da SEI saiu para a ST; a ST vendeu os ônibus a prazo e foi baixada. É preciso verificar onde ficaram o crédito e os recebíveis (achado 4).
- (b) Toyota ST → Rádio (achado 6).
- (c) Possível dupla contagem dentro de Rafael (sala + box). PDF da Rádio, p. 3: a faixa de R$ 380–520 mil é para a sala com "94,51 m² de área total, o que inclui as áreas comuns do prédio e a vaga de garagem", e a "Inclusão do Box de Garagem … agregando entre R$ 35.000 e R$ 50.000 ao valor final do conjunto". No balanço (p. 9), Rafael põe sala R$ 420.000 **e** box R$ 40.000. É preciso conferir se o box entra duas vezes.
- Correção: acrescentar esses itens como S4-7 a S4-9 (suspeitas, com as referências acima).

### 13. BAIXA: a faixa C1 do Fox não contém o valor de Rafael
- Evidência: PDF da Rádio, p. 6: FIPE R$ 32.200 a R$ 35.500, e na mesma página anúncios de "R$ 31.000 e R$ 39.000". No balanço (p. 9) o valor é R$ 32.000.
- O que está na planilha: Mestre, linha 07-2025-06, mostra só a primeira faixa. O valor de Rafael fica abaixo do mínimo exibido.
- Correção: mostrar as duas faixas.

### 14. BAIXA: faltam vínculos na coluna "Parte 6"
- A redução de 8 % (07-2025-N17) e a antena (07-2025-04) não estão ligadas ao S4-1, que trata exatamente da redução de 8 %.
- A sala 311 (07-2025-01) não está ligada ao S4-6 (lote "permutado" da sala, doc 25, par. 202).
- Correção: ligar por ID, não só por regex de chave.

### 15. BAIXA: precisão de textos em `parte6-candidatos.md`
- S4-1: segundo o arq. 1, p. 13, a redução de 8 % cobre também "imposto de transmissão e emolumentos, afora aquisição das cotas do sócio minoritário Antonio Telmo De Toni", e não só mudar a antena e comprar terreno.
- S4-2: incluir o que o diário já registrou: série CAT0320DAA6F00902, valor da NF R$ 470.000 em 12/11/2009 e anotação "Maquina do Celito" (arq. 6, p. 1). Esta última continua dependendo da decisão sobre o E7.
- S4-5:
  - Explicitar a divergência: a ST diz que a SEI pagou R$ 269.300 a partir de out/2011; a SEI diz ter pago R$ 762.000 + R$ 180.000.
  - Doc 23 (2011): a lista de credores da TS/ST inclui "Radio Comunidade Ltda. R$ 25.000", "Siviero Empreendimentos R$ 11.000" e "Hotel Araçá" (R$ 23.000 + 50.000 + 88.500).
  - A frase "sem registrar o passivo nem o crédito" é ambígua: a ST 2025 registra um crédito (IXO 6000).
- S4-6: cruzar com a lista C0. Rafael escolheu 3 lotes do "07. Colina do Retiro": 002-05, 002-09 e 003-09 (Diretrizes p. 9; `ativos-avaliados-valter.xlsx`, linhas 24–26). É preciso ver se o lote da dívida da Rádio ou o lote da sala é um deles.
- As demais referências conferem: doc 18 tab. 5 itens 02/06, tab. 23 item 07, tab. 28, tab. 31; doc 11 tab. 5 itens 03/07, tab. 7 item 02; doc 25 par. 143–145, 202 e 234; arq. 1 p. 13; 8 % × 2.698.000 = 215.840.

### 16. BAIXA: o Resumo conta linhas pelo prefixo do ID
- A Siviero Transportes aparece com "Linhas: 6", mas tinha 7 bens em 2018: a Toyota está contada na Rádio (ID 07-2025-07). A soma 2018 está certa porque usa a coluna "Empresa (2018)".
- Correção: contar por "Empresa (2018)" e "Empresa (2025)", ou explicar no Leia-me.

### 17. BAIXA: formatos
- Faixa C1 escrita com separador inglês ("R$ 380,000 a R$ 520,000").
- Na aba de notas, a coluna "Valores" usa "150,000.00" e " | 8,0%" (começa com separador).
- A numeração de "Parágrafo" (só parágrafos não vazios) não está explicada no Leia-me.
- Correção: usar formato brasileiro e explicar a numeração.

### 18. BAIXA: a "Ref. 2018" usa o número da empresa de 2025
- Exemplo: "Inventário 2018 da empresa 07, item 01". No Drive, os arquivos de 2018 têm outra numeração: 9.1 Rádio, 11.1 Ceval, 7.1 Siviero Transportes, 14.1 Associação, 8.1 Turismo.
- Correção: citar o arquivo (por exemplo "arq. 9.1 (Rádio 2018), tab. 3, item 01").

### 19. BAIXA: o Resumo da Turismo diz "sem total no documento"
- O inventário 2018 da Turismo tem o total em branco, mas o doc 27 dá "TURISMO SIVIERO LTDA. R$ 0,00".
- Correção: citar o doc 27.

### 20. BAIXA: o rótulo "Não abordado (só em 2018)" é ambíguo
- Não deixa claro se quem não abordou foi o grupo ou Rafael. Para os bens da Siviero Transportes e da Ceval, o fato é "não existe no grupo em 2025".
- Correção: usar "Só em 2018 (ausente em 2025)".

### 21. BAIXA: a dívida do ônibus IUY 8068, "(está financiado)", não está nas notas
- Evidência: inventário 2018 da Siviero Transportes, item 01. O doc 25, par. 227, cita "18 parcelas … GUIBOR, de R$ 8.000,00 … do ônibus atual" e "09 parcelas … SANTANDER, de R$ 7.000,00".
- Correção: registrar como dívida de 2018 com destino pendente.

---

## O que conferi e está correto

- **Cobertura, 1 linha por bem.** Bens com valor em 2018: 04 = 4, 07 = 6, 10 = 7 (total 17); 03 e 09 não têm itens. Bens com valor em 2025: 07 = 7, 04 = 3 (total 10); 03, 09 e 10 só têm linhas "-.-" de R$ 0. O Mestre tem 18 linhas: 17 com "Existe em 2018? = Sim" e 10 com bem de 2025, sem duplicata nem omissão. O inventário 2025 "antigo" tem os mesmos itens e valores.
- **Pares por chave (todos corretos):**
  - sala, mat. 11.382;
  - box, mat. 11.324;
  - Fox, placa IRO 2562;
  - Toyota, placa IHW 0285, de ST 2018 item 04 para Rádio 2025 item 07;
  - 320D, NF 4946;
  - Case 580L, placa IKI 4325 e chassi JHF0048344;
  - LK 1313, placa IHM 0033 e chassi 34500312579515;
  - terreno, mat. 3.751, para doc 11, tab. 7, item 02, R$ 7.148.016.

  Não há par por semelhança de texto além dos 2 propostos (canal/equipamentos e antena), que estão marcados "PROPOSTO, aguarda aprovação".
- **Bens sem destino.** ILD 3959 (Ceval), IUY 8068, IAZ 0281 e ISS 7019 não aparecem em nenhum inventário 2025 (01 a 11). Também não há placa, chassi nem modelo "2423" ou "Volare". "Sem destino documentado — pendente (P4)" está correto.
- **Valores por linha.** Valor 2018, Valor grupo 2025 e Valor Rafael conferem com o .docx e o PDF em todas as 18 linhas. As diferenças estão certas: sala +210.000 (100 %), Fox +4.000 (14,3 %), 320D +40.000 (20 %), Case +30.000 (42,9 %); as demais são 0. A coluna Situação está certa, linha a linha.
- **Páginas das referências de Rafael.**
  - Balanço da Rádio: itens 1–6 na p. 9 (o item 1 começa na p. 8, com valor na p. 9); item 7 na p. 10.
  - Balanço da Ceval: p. 5.
  - Faixas: sala p. 3, Fox p. 6, Toyota p. 7, 320D p. 2, Case p. 3, LK p. 4.
  - Totais de Rafael: R$ 2.912.000 (p. 10) e R$ 385.000 (p. 5).
- **Lista C0.** Só a Toyota está marcada "Sim". Nenhum outro bem de 03/04/07/09/10 aparece nas Diretrizes (p. 7) nem em `ativos-avaliados-valter.xlsx` (linha 10).
- **Rafael e o terreno 3.751.** Rafael não avalia esse terreno em lugar nenhum. Procurei no PDF da Rádio, nos 13 PDFs da SEI, nas Diretrizes e na Introdução. "Existe no Estudo 2 (Rafael)? = Não" está correto.
- **Decisões.** Todas as 18 linhas estão com "Pendente."; "Valor adotado" e "Motivo" estão vazios. `decisoes.json` só tem 2 casamentos "proposto" e nenhuma decisão.
- **Aba de notas.** Os 13 valores e textos conferem com as fontes:
  - Celito R$ 150.000;
  - Colina do Retiro I R$ 110.000;
  - Receita Federal R$ 77.000 / R$ 2.750 (Turismo) e R$ 7.000 / R$ 250 (ST);
  - ônibus 2000 R$ 95.000;
  - R$ 1.733.700 + R$ 269.300 = R$ 2.003.000;
  - IXO 6000 R$ 115.000 / R$ 20.000 / R$ 5.000;
  - datas de baixa 09/07/2020, 04/09/2023 e 08/04/2024;
  - redução de 8 %.
- **Resumo.**
  - Soma 2025 = total declarado nas 5 empresas: 2.698.000, 315.000, 0, 0, 0.
  - Soma de Rafael = total de Rafael: 2.912.000 e 385.000.
  - Soma 2018: Rádio 9.316.016, Ceval 391.000 e Siviero Transportes 2.443.100 bruto (ver achado 3); Associação "NÃO" é real (R$ 10.000 sem itens).
- **Observações do Luís.** As 4 notas "Em 2018 estava na relação 9.1 RÁDIO…" estão ligadas ao terreno 3.751, com situação "Pendente". Não há outras observações do Luís nas linhas destas empresas.
- **Colunas exigidas pelo plano.** Todas existem. A única ressalva é a coluna "Registro de data-base", vazia (achado 8).
