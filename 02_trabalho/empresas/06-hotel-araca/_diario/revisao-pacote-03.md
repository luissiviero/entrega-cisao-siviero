# Revisão independente — Pacote 3 (destino 2018 → 2025), empresa 06 Hotel Araçá

Data: 28/10/2026. Revisor com contexto novo. Objeto: `pacotes/pacote-03-destinos.md` (gerado por `_scripts/pacote_03.py`, mestre md5 `9d5aa8ca61`, busca md5 `6af8bba685`, os dois conferidos). Não alterei nenhum arquivo além deste relatório. Scripts e extrações temporárias do revisor: `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\5fd9298d-a99a-4db0-a826-38f7fb4a324d\scratchpad\revisor-pacote3\`.

Contagem: **2 altas, 4 médias, 5 baixas.**

---

## Altas

### A1 — O arquivo 18 documenta um lote de Capão da Canoa dado a Válter em 2019, e o pacote diz que não há destino nenhum

- **Evidência.** Arquivo 18 ("Planilha de lotes vendidos e/ou outras formas de pagamento na aquisição de bens para acerto futuro – Válter Materiais de Construção Ltda. e Válter Siviero"), PDF p. 3, bloco de 10.06.2019, conferido também como imagem: linha "Lote Capão da Canoa", com a descrição "Lote na esquina próximo ao Hotel Araçá Ltda., com valor diferenciado aos demais lotes dados aos sócios." No docx está no bloco 12 e na tabela 3, linha 4.
- **A busca encontrou o trecho e o pacote não o usou.** O trecho está em `busca-destinos-06.xlsx`, aba Achados, termo "Capão da Canoa", arquivo 18, nos locais "bloco 12", "tabela (bloco 25) linha 4" e "p. 3". Nada no diário, na mestre, no G7, no G8 ou na Parte 6 o menciona.
- **Efeito.**
  1. "Resultado: nenhuma escritura, venda, permuta ou dação dos cinco lotes depois de 2018" é falso como está escrito. Há pelo menos um lote de Capão entregue a Válter em 2019, para acerto futuro.
  2. Para D1 (lote de Válter, 360.0019.004.000, 450 m²), a resposta "Destino documentado? não" provavelmente está errada. O destino documentado mais provável é a consignação a Válter (arquivo 18, 10.06.2019), com valor "diferenciado" ainda não informado.
  3. O documento não dá a inscrição. Se o lote "na esquina" for o item 02 de 2025 (360.0019.003.000, "Casa da Praça", vizinho do D1), esse lote já estaria consignado a Válter em 2019 e foi escolhido de novo na lista de Rafael (lado de Válter/Benta) em 2025. Isso é risco de contagem dupla e candidato à Parte 6.
  4. A frase "demais lotes dados aos sócios" é indício documental de que os lotes com nome de sócio de 2018 (D2, D4, D5 e o de Irineu) foram "dados aos sócios". Isso aponta para consignação a sócio, que é a forma de destino que o P4 pede para documentar.
- **Correção.**
  - D1 passa a "pendente, com destino parcial: lote de Capão consignado a Válter no arquivo 18 (10.06.2019); inscrição não identificada".
  - Citar o arquivo 18 em D2, D4 e D5 como indício ("lotes dados aos sócios").
  - Nova pergunta ou extensão do G7: qual lote é o "na esquina próximo ao Hotel", qual foi o "valor diferenciado" e quais foram os "demais lotes dados aos sócios".
  - Novo candidato à Parte 6: arquivo 18 × D1 / item 02.
  - Corrigir a frase "Resultado".

### A2 — D3: a leitura "troca a favor do sócio" contraria três trechos das próprias fontes

- **Evidência.**
  1. Inventário 2018 (doc 16), item 23: o valor está em aberto como "R$ …(Avaliar Lote e Casa) … ELCIO SIVIERO", e o total diz "R$ 49.684.243,20 (+ o valor da casa e lote – Élcio Siviero)". Em 2018, a casa e o lote de Élcio são tratados como bem do Hotel a somar ao total, ao contrário dos outros lotes com nome de sócio.
  2. Doc 25 de 2018, seção Hotel Araçá: "Dentre os sócios está estabelecido, que as casas pertencem para: – CASA DA ESQUINA: Élcio Siviero (que pertence à empresa Siviero)".
  3. Divisão de lotes de 2015 (`A-PLANILHA … 2015` e `6. DIVISÃO DE LOTES 2015 – VÁLTER`): "Lote 10 – Quadra 03 – Aditivo Madianeira, que conforme contrato foi dado para o Paulo Busatto. | R$ 70.000,00". O valor aparece lançado contra Élcio. Em 2018 (doc 25) o valor fica em branco ("R$……") e é acrescentado "(Trocado com o lote de Capão da Canoa – CASA)".
- **Outras leituras possíveis, as três compatíveis com o texto.**
  - (a) A leitura do pacote: a troca é crédito a favor de Élcio e a casa sai da empresa.
  - (b) Élcio entregou a casa da esquina à "empresa Siviero" em troca do lote 10/q03 da SEI, que ele passou a Busatto. Por isso o débito de R$ 70.000,00 some em 2018 e a casa "pertence à empresa Siviero". Nessa leitura, **a casa é bem de empresa (SEI ou Hotel) e falta no inventário 2025**. É exatamente o que Rafael aponta na p. 24 ("Não constaram no quadro patrimonial do Hotel Araçá… duas residências… na esquina da Rua da Divisa com a Rua Três"), com estimativa de R$ 2,4 a 3,5 milhões pelas duas casas.
  - (c) Élcio recebeu a casa de Busatto em troca do lote. Nesse caso ela é dele, mas o doc 25 diz o contrário.
- **Efeito.** Escrever na coluna "Destino 2018→2025" que o destino parcial é uma "troca a favor do sócio" registra um destino que as fontes não sustentam. Pela leitura (b) e pelo C0, o efeito vai no sentido oposto: se a casa é da empresa e fica na empresa, entra com o maior valor. O Luís aprovaria o item 1 do pacote sem ver esses três trechos. A mestre já tem os trechos 1 e 2 parcialmente nas "Notas de identificação", mas o filtro de `pacote_03.py` (lista `pistas`) os descarta: as notas que começam com "valor '(Avaliar…", "o total de 2018 ressalva…" e "possível correspondência com as duas residências…" não passam.
- **Correção.**
  - D3 passa a "pendente; leituras em conflito (a/b/c)", sem "a favor do sócio".
  - Citar os três trechos e Rafael p. 24.
  - Ligar D3 ao G8 e aos candidatos 2 e 5 da Parte 6.
  - Perguntar "qual empresa é a 'empresa Siviero'" que é dona da casa da esquina.
  - Ajustar o filtro `pistas` e tirar o texto de destino fixo no código (o `if "TROCOU"` com frase pronta).

---

## Médias

### M1 — D2 (Raul): o pacote omite a atribuição documentada e o vínculo com G8, Rafael e a Parte 6

- **Evidência.** Doc 25 de 2018, seção Hotel: "CASA AO LADO: Raul Siviero", na mesma lista em que "as casas pertencem para", e "manutenção das 02 casas de veraneio dos Sócios, com o pagamento das despesas… durante o ano todo". Rafael, p. 24: as duas residências da Rua da Divisa esquina Rua Três não constam no Hotel. O G8 e o candidato 5 já existem (Pacote 2), mas a linha D2 do pacote só cita a lista de IPTU e diz que não há destino documentado.
- **Correção.** D2 passa a "atribuída a Raul por acordo entre sócios (doc 25), sem ato de transferência; manutenção paga pelo Hotel; pendente". Citar Rafael p. 24, G8 e candidato 5. Juntar G7 e G8 numa única pergunta sobre os cinco lotes e as duas casas, para o grupo não receber duas perguntas sobre os mesmos bens.

### M2 — A cobertura declarada da busca está exagerada e há material não lido

- **"Todos os … .pdf" não é verdade na prática.** Os PDFs digitalizados sem camada de texto e fora do corpus de OCR não geram achado nenhum:
  - `ACERTOS E ACORDOS.pdf` (6 p.; analisado em outra sessão, sem achado para a 06);
  - `2019-03-11_Considerações_finais_Luciano.pdf` (5 p.);
  - `Documentos/retiradas_socios_2018.pdf`;
  - `Contrato Siviero X Buriti Holding Sul Participações.pdf` (20 p.);
  - `Holdings/VMC - 3ª alteração contratual.pdf` (VMC = Válter Materiais de Construção);
  - 15 relatórios mensais.
- **Imagens e arquivos sem extensão não foram lidos:**
  - `2018/2019-01-14_ata_quarta _reunião` (JPEG sem extensão, idêntico ao `2019-01-14 — Reunião Família Siviero.jpg`);
  - `Documentos/Relatório hotel araçá 2025.jpg`, `relatórios empresas 2025.jpg`, `Cobrança Siviero` (JPEG), `quantidade soja 2025` (JPEG);
  - `Material a entregar Pablo` (texto);
  - `.gdoc`/`.gsheet` e o `.pptx`.
- **Os .doc não precisavam ficar de fora.** O conector do Drive os lê (`application/msword`). Li dois pelo conector:
  - "1." (28 KB) não é o contrato de compromisso mútuo: é o instrumento de cessão de quotas do Hotel Princesa dos Vales, de 15/03/2004.
  - "B-" (246 KB) traz a cessão de quotas do Hotel Araçá de 2003.
  - Nenhum dos dois cita os lotes de Capão. "2." e "C-" têm o mesmo tamanho de "B-" e de "1.", respectivamente.
- **O que eu mesmo li e como ficou.** Li como imagem a ata de 14/01/2019, o relatório do hotel de 2025, a p. 1 das Considerações de 2019 (igual ao doc 26 docx) e o `retiradas_socios_2018.pdf` (é a mesma página da ata de 14/01/2019). Nenhum deles traz destino de D1–D5. Fiz busca de texto completo em todo o Drive (não só na pasta Cisao_Siviero) por 360.0057, 360.0074, 360.0019.004 e Busatto: só aparecem cópias dos inventários, do doc 6 e de "Relação de capitais.docx" (2019, igual ao doc 6). O resultado da busca se sustenta, salvo A1, mas o texto do pacote precisa dizer o que não foi lido.
- **Correção.** Reescrever o parágrafo "Busca feita" com:
  - a lista de não lidos e o motivo de cada um;
  - os itens lidos por imagem ou pelo conector;
  - a correção da descrição dos .doc.

### M3 — Os termos de busca não pegam a forma como os destinos são escritos

- **Evidência.** Faltam os termos "esquina", "casa da esquina", "casa ao lado", "consignado", "Coletto", "Gramado"/"Lago Negro", "dólar"/"cofre" e lote/quadra derivados das inscrições (lote 04/q19; lotes 01–02/q57; lotes 30–31/q74). As inscrições só são procuradas com ponto opcional, sem hífen e sem forma compacta. O achado de A1 só apareceu pelo termo genérico "Capão da Canoa" (394 achados em 32 arquivos), que ninguém revisou fora dos inventários.
- **Correção.** Acrescentar esses termos e revisar um por um os achados de "Capão da Canoa" e "Hotel Araçá" fora dos inventários 2015, 2018 e 2025.

### M4 — D7 (dólares): classificação inconsistente, linha já aplicada e o outro lado omitido

- **Evidência e problemas.**
  1. A mestre tem Situação "Não abordado" (Parte 5), mas Destino 2018→2025 "Pendente (Parte 4)", e a coluna "Ref. 2025" guarda um documento de 2018. O doc 22 faz parte do conjunto do inventário 2018 e registra os dólares como do Hotel (posição de 25/05/2017). Há leitura defensável de que é bem "existente em 2018" e que o P4 exige destino. É preciso escolher uma das duas e registrar o motivo.
  2. O cabeçalho diz "Nada está aplicado", mas o próprio D7 diz que a linha está "já incluída" na mestre, antes da aprovação (portões de revisão).
  3. O anexo registra U$ 1.000,00 a Válter ("Viagem Cruzeiro", 23.02.2012) e U$ 700,00 a Luciane ("Viagem Peru", 09.09.2013). São saídas a sócio ou família e, portanto, candidatos à Parte 6. Estão nas notas da mestre, mas não no pacote nem no candidato 6 proposto.
  4. Possível outro lado do "Pago Coletto–Gramado – 02 lotes": a lista de IPTU de 2018 de Válter (doc 6) traz dois lotes em Gramado, Loteamento Lago Negro, em nome de Válter Materiais de Construção Ltda. (L 18 Q 06 e L 08 Q 02), e um "SEI/Válter, a definir". Lago Negro é loteamento da SEI (arquivo 17; doc 25). Não está provado que sejam os mesmos, mas é a pista do outro lado que o candidato 6 precisa.
  5. A redação do G12 não é mostrada, e o Luís aprova uma pergunta que não leu.
- **Correção.** Decidir P4 ou Parte 5 e alinhar a mestre. Dizer que a linha já foi gravada ou desfazê-la até a aprovação. Incluir as retiradas de Válter e Luciane e a pista dos lotes de Gramado no candidato 6. Mostrar o texto do G12.

---

## Baixas

- **B1 — D8 com data errada.** D8 é chamado de "Proposta de divisão de 2019", mas a frase vem da convocação de 26/05/2017 (`2018_desatualizado/G-SÓCIOS… Reunião no dia 26 de maio de 2017.docx`, no disco, com o mesmo texto e sem o "R$ 170.000,00"). Em 2019 só se preencheu o faturamento. O mesmo arquivo de 2017 traz a planilha de dólares idêntica e pode ser citado como fonte em disco para D7. Vale registrar que "EXCETO os lotes adquiridos adjacentes" interessa ao C0 e à lista de Rafael (lotes adjacentes 02, 11, 12, 18 e 19), mas continua sendo só registro.
- **B2 — D4: o doc 5 não é do sócio.** O doc 5 é a lista de José Valério, não do "sócio": o titular é Adélia, e José Valério aparece com ela no quadro de 2025. Explicar o vínculo.
- **B3 — Frases que confundem.** A frase "As inscrições aparecem só em…" mistura arquivos de duas inscrições e omite que 360.0074.030 e .031 e as matrículas 77.833 e 74.857 aparecem só no doc 16. Já "Nenhum deles está no inventário de 2015", ao lado de um rascunho cujo nome diz "Dezembro de 2015", confunde. Dizer que a data do rascunho é incerta.
- **B4 — `pacote_03.py`.** O destino do D3 está escrito à mão no código, não vem da mestre. O filtro `pistas` corta notas relevantes, e é a causa técnica de A2 e M1. A variável `files_scanned` é calculada e não usada.
- **B5 — Inscrições de Élcio e Irineu.** As listas de IPTU de Élcio (360.0059.005.000) e de Irineu (360.0058.001.000 e "Circo") são de imóveis em Capão, perto da quadra 0057. Pela regra não formam par, mas a "Leitura" poderia dizer que são bens dos sócios, fora do Hotel, para não parecer que foram verificados como destino.

---

## O que conferi e está certo

- **Lista dos bens de 2018 sem par está completa: são os itens 20, 22, 23, 24 e 25.**
  - Os itens 01–19 de 2018 têm par por inscrição ou matrícula. O item 21 (360.0038.015.000) faz par com o item 17 de 2025.
  - O item 01 tem os mesmos componentes nos dois anos: 18 terrenos em 12 matrículas, móveis de 36 e de 108 aptos, obra civil, equipamentos (com o poço) e vantagem da coisa feita. A queda de 30 % para 15 % é de valor, não de existência, e já foi tratada.
  - O inventário 2018 não lista veículos nem móveis avulsos. A Fiat Toro não existe em 2018, então não é caso de P4.
- **Numeração do rascunho:** Élcio é o item 22 e Raul o item 23. Os de Adélia e Pedro não aparecem. O inventário de 2015 tem 19 itens.
- **Citações literais conferidas na fonte:**
  - doc 6 (Válter, 360.0019.004.000, R$ 500,00);
  - docs 1, 4 e 5 (Capão sem inscrição, R$ 500,00);
  - docs 2 e 3 (outras inscrições);
  - nota "TROCOU ESTE LOTE/CASA…" do doc 16;
  - doc 25, Busatto ("Madianeira", R$ 70.000,00);
  - matrícula 77.833 sequencial à 77.832.
- **Doc 22:** li o arquivo pelo conector do Drive (fileId 1HgldDKaDDFwpNAISVxnAsHpVHIpbliJK). Os trechos de `doc22-2018-trechos.txt` conferem com o texto do conector. A soma 110 + 175 + 1.800 + 10.000 + 25.613 = 37.698,00 confere. Também conferem U$ 9.000,00 "Coletto–Gramado" em 15.05.2009, a data de 25.05.2017 e a frase do D8.
- **D6:** 12.000,00 m² em 2018 e 10.500,00 m² em 2025, confere.
- **Arquivo 17:** não tem lote em Capão nem o lote 10/q03 do Aditivo Medianeira. Ligado ao Hotel, só a linha 60 (placas solares).
- **Inventários 2025 das outras empresas:** "Capão da Canoa" aparece só no 06 e no 11 (Solar Serra). Nenhuma outra empresa tem D1–D5.
- **md5** da mestre e da busca conferem com o cabeçalho. O único erro de leitura (doc 22, "No such file") foi registrado e compensado com a leitura pelo conector.
