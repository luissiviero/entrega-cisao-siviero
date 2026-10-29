# Revisão independente do pacote 1 (empresas 02 e 05)

Revisor com contexto novo (P11), 28/10/2026. Objeto: `pacotes/pacote-01.md`. Rodei `_scripts/pacote_01.py` de novo e o pacote saiu idêntico ao arquivo revisado (diff vazio; saída "R1 5 17023657.5; R2 5; R3 3; exceções 3; destino 8").

Conferi tudo nas fontes, não na planilha:
- doc 05 2025 e doc 02 2025 (docx);
- docx "áreas vendidas";
- docs 13 e 17 de 2018, lidos no Drive;
- PDF de Rafael do Hotel, todas as 24 páginas;
- PDFs de Rafael da Agropecuária, do Diretrizes, das "ÁREAS RURAIS…" (com o xlsx de apoio) e da Transportadora;
- `ativos-avaliados-valter.xlsx`;
- log da SEI (pacotes 3 e 4) e `pacotes/pacote-04.md` da SEI;
- log e achados aprovados do ACERTOS;
- arq. 1 (p. 5, 6, 8 e 9) e arq. 9 (p. 1), lidos como imagem;
- arq. 11 (matrícula 1.412), lido como imagem.

## Achados

### A1 — alta: o item 06 (4 ha da Fazenda São Crispin, mat. 11.090) não deveria estar na regra em bloco sem ressalva

- **Evidência.**
  - **Avaliação independente.** No PDF de Rafael, p. 13–14 (entre os itens 05 e 06), Rafael pede "avaliação técnica específica e independente da Fazenda São Crispim". Até lá, os valores ficam como "parâmetros preliminares de referência, sujeitos à revisão". O item 06 se chama "FAZENDA SÃO CRISPIN… DENTRO DE UM TODO MAIOR" e usa o mesmo R$ 110.000/ha do item 05.
  - **Concordância por remissão.** A concordância de Rafael com o item 06 (p. 14) diz "de acordo como citado nos itens anteriores (02., 03., 04., 05.)". Ela remete ao item 05, que tem justamente essa ressalva e nenhuma concordância própria.
  - **Regularização de R$ 400.000.** O doc 05 2025 (OBSERVAÇÕES), repetido por Rafael na p. 6, diz que o valor "não é líquido": sofrerá redução de R$ 400.000,00 pela regularização da FAZENDA SÃO CRISPÍN. O ITBI é calculado sobre 411,4 ha, que só fecham com o item 06 incluído (407,53 + 4,0; achado E2 aprovado). As escrituras "em nome… de TERCEIROS, POR CONTRATO" incluem o item 06, "adquirida por CONTRATO DE COMPRA E VENDA, em dezembro de 2015". A própria planilha-mestre (aba "Passivos e deduções") atribui os R$ 400.000 aos "itens 05 e 06".
  - **Conclusão.** A frase do pacote "Nenhuma das ressalvas muda o valor" não vale para o item 06. O valor dele depende de duas questões abertas no pacote 2: a avaliação independente e a dedução de R$ 400.000. O item 06 está na lista de Rafael, então o C0 pede o menor valor. Se a dedução for aceita para a fazenda, o menor valor do item 06 pode ficar abaixo de R$ 440.000.
- **Correção sugerida.** Levar o item 06 para o pacote 2, junto com o item 05, e decidir os dois juntos (mesma fazenda, mesmo R$/ha, mesma dedução). A R1 passa a ter 4 bens e R$ 16.583.657,50. Se o item ficar na R1, a ressalva precisa dizer que o valor "segue a decisão do item 05 quanto à avaliação independente e aos R$ 400.000" e citar Q1 (matrícula 11.090).

### A2 — média: as ressalvas do item 04 (Sede Campestre) estão incompletas

- **Evidência.**
  - Arq. 1, p. 9: "No que tange ao imóvel conhecido como SEDE CAMPESTRE… pendem de averbação as suas benfeitorias". Pavilhão, quiosque e casa, R$ 1.330.857,50 com os 25 %, não estão averbados.
  - No doc 05 e em Rafael (p. 3 e p. 12), a casa de alvenaria tem área em branco ("...............m²") e vale R$ 250.000,00.
  - O item está na lista de Rafael, ou seja, vem para o Luís. Custo e risco de regularização recaem sobre quem recebe.
  - A concordância de Rafael é condicionada à "regularidade registral" (p. 11).
- **Correção sugerida.** Acrescentar à ressalva do item 04: "benfeitorias não averbadas (arq. 1 p. 9); área da casa não informada". Não muda o valor pela P5; é ponto para a ficha de fechamento ou pergunta ao grupo.

### A3 — média: as listas de "Fora deste pacote" não mostram onde estão 8 dos 21 bens de 2018 nem os R$ 400.000

- **Evidência.**
  - Conferi nas fontes: o doc 13 de 2018 tem 4 bens e o doc 17 de 2018 tem 17 (11 numerados mais 6 parcelas e linhas sem número). São 21 ao todo.
  - O pacote mostra 5 na R2 e 8 na etapa 6. Os outros 8 têm par num item do Hotel em 2025:
    - Hotel 2018, itens 01 a 04 → itens 01 a 04;
    - Agro 2018, seq. 1 → item 05;
    - Agro 2018, seq. 7 → item 06;
    - Agro 2018, item 06 (lote 52) → item 07;
    - Agro 2018, item 10 (Mercedinha) → item 08.

    Os pares estão certos (matrícula, área e placa conferidas), mas o pacote não os mostra. A R1 não tem coluna 2018, e o leitor não consegue confirmar que cada bem aparece uma só vez.
  - A regularização de R$ 400.000 (docs 02 e 05) também não aparece em nenhuma lista.
  - O doc 02 2025 (item 01 "-.-", R$ 0,00) não é citado.
- **Correção sugerida.**
  - Coluna "Par em 2018" na R1.
  - Em "Fora deste pacote", a linha: "8 bens de 2018 com par no Hotel 2025: 5 na R1 e 3 no pacote 2 (itens 01, 05 e 08)".
  - Uma linha para os R$ 400.000, que vão para o pacote 2 com os itens 05 e 06.
  - Uma frase: "Agropecuária 2025: inventário sem bens (R$ 0,00)".

### A4 — média: a R2 afirma um destino ("transferido para outra empresa do grupo") sem documento

- **Evidência.**
  - P4: o destino precisa ser documentado.
  - A planilha-mestre marca essas 5 linhas como "Pendente (destino, etapa 6)".
  - `parte6-candidatos.md` P6-05: "Nenhuma fonte traz documento de transferência da Agropecuária para a Transportadora". As perguntas Q9 e Q13 ao grupo pedem justamente isso.
  - A sessão S3 liga a terra de 69,8 ha aos créditos da venda de quotas do Rovílio.
  - Os lotes 12, 89 e 33 já estavam na lista "bens em outras empresas" da SEI em 2018. Não se sabe se houve transferência ou só mudança de inventário.
- **Correção sugerida.** Trocar por "localizado em 2025 no inventário da SEI (doc 11) / da Transportadora (doc 08); documento de transferência pendente (Q9, Q13); destino confirmado na etapa 6".

### A5 — baixa: falta à R2 a pendência da SEI sobre os 7,5 ha do lote 12

- **Evidência.** A frase do pacote está certa nos dois pontos que conferi.
  - Pacote 4 da SEI, seção A: a SEI pareou "Outras empresas (Agropecuária)" lote 12 (mat. 3.062), lote 89 (8.486) e lote 33 (1.218) com as áreas rurais #04, #08 e #09.
  - Pacote 3 da SEI: adotou os valores de Rafael (R$ 650.000, R$ 1.500.000 e R$ 250.000). O C0 dá o maior porque os lotes não estão na lista de Rafael.

  Falta dizer que o caso 7 do pacote 4 da SEI (os 7,5 ha a mais do lote 12) segue "Pendente" na SEI. É a mesma questão de P6-03 e Q9, e sem o aviso as duas sessões podem tratá-la em dobro.
- **Correção sugerida.** Acrescentar ao lote 12: "a SEI deixou os 7,5 ha pendentes (pacote 4 da SEI, caso 7)".

### A6 — baixa: as ressalvas dos itens 02 e 03 não dizem tudo sobre a faixa de mercado de Rafael

- **Evidência.**
  - **Item 02.** A citação está correta, mas falta a página (p. 8). Falta também dizer que a faixa vale para a "área comercial premium de 1.800 metros quadrados", isto é, a matrícula 1.412 inteira (arq. 11: lote 313 com 1.800 m²), que inclui os 1.080 m² do item 02. Por isso a decisão do item 01 no pacote 2 e a do item 02 usam o mesmo argumento.
  - **Item 03.** Na p. 9, Rafael cita "lotes padrão de menor porte (como 450 m²)" no centro, "sem frente direta para a avenida principal", a cerca de R$ 1.660/m². O valor do grupo é R$ 3.333,34/m². É o mesmo tipo de risco do item 02, e o pacote não o registra.
  - Nos dois itens o valor fica, porque os dois lados dão o mesmo valor (P5) e os bens ficam na empresa (C0).
- **Correção sugerida.** No item 02: "(p. 8; a faixa cobre a matrícula inteira, inclusive este terreno; ver item 01 no pacote 2)". No item 03: a referência da p. 9.

### A7 — baixa: a ressalva comum a todos os bens da R1 está incompleta

- **Evidência.**
  - **Concordância condicionada.** A concordância de Rafael fica "sujeita à validação pelo escritório de advocacia… adequação dos valores" (p. 11). A seleção de bens também é condicionada à "conferência dos valores atribuídos" (p. 21).
  - **Data-base.** O RDB-1 (o grupo diz que os valores são da data de elaboração do inventário, arq. 1 p. 9) vale para os 5 bens e não é citado.
  - **Termo.** Chamar a concordância de "Argumento técnico (C1)" desvia do plano: C1 é a faixa de mercado de Rafael.
- **Correção sugerida.** Uma linha abaixo da tabela: "Concordância de Rafael condicionada à validação jurídica (p. 11 e p. 21); data dos valores: ver RDB-1". Trocar "Argumento técnico (C1)" por "Fundamento".

### A8 — baixa (informação, sem mudança de valor): item 07, lote 52

Nada no texto de Rafael para este item tira o lote 52 da regra. Ele fica na empresa, então o C0 pede o maior valor, e os dois lados dão R$ 422.800 (R$ 28.000/ha). Para referência, Rafael avalia terras vizinhas por valores mais altos por hectare:
- lotes 63/65/66/N de São Brás, Cotiporã, Linha Júlio de Castilhos: cerca de R$ 50.000/ha (R$ 3.500.000 por 69,8 ha);
- lote 12, Linha Júlio de Castilhos (Veranópolis): R$ 50.000/ha (R$ 650.000 por 13 ha).

Sem um segundo valor para o próprio item, a P5 não permite mudar. Fica só como argumento possível (C1), se o Luís quiser um "caso extremo".

### A9 — baixa: linguagem para leitor não técnico

1. Códigos internos sem explicação:
   - "T8R4", "11 T8R5", "doc 05 2025 T5R2";
   - "seq. 10", "02 2018 item — (seq. 8)";
   - "Fase 0 (…)", "candidato P6-0x", "C1", "Q4 e Q5", "E1", "S3".

   Sugestão: "inventário da SEI, áreas rurais, item 04"; "candidato a desconto em duplicidade (Parte 6), nº 3"; e uma legenda curta no fim.
2. Ressalva do item 03: "Fase 0: 2025 item 03: 750 m² × R$/m² = valor — O documento arredonda: 750 × 3.333,34 = 2.500.005,00" é ilegível. Sugestão: "O documento usa R$ 3.333,34/m²; 750 m² × 3.333,34 = R$ 2.500.005,00, arredondado para R$ 2.500.000,00 (diferença de R$ 5)".
3. Ressalva do lote 12 na R2 com texto truncado no meio:
   - "mesma área e vendedor em 11 T8R4: 03 | Uma área de terras… localizada no Município de…";
   - "[comentário: DESCRIÇÃO ANTIGA: … adquiri]; Mudou valor".

   Sugestão: resumir em uma frase e citar a nota do Luís pelo sentido ("descrição mudou; valor mudou").
4. A abertura diz "organiza o resto em duas regras de escopo", mas o resto do Hotel (itens 01, 05 e 08) vai para o pacote 2. A R2 e a R3 tratam de outros bens. Sugestão: "…e trata em duas regras os bens de 2018 que estão em outra empresa e os créditos em soja".
5. "Proposta: adoto…" soa como decisão já tomada. Sugestão: "Proposta: adotar…".
6. "Pendente com você: a base do benchmark 2018 (A ou B…)" não diz o que são A e B nem os valores. Faltam A = R$ 46.459.267,83 e B = R$ 48.847.497,83. Falta também o aviso que o próprio log manda dar: B mudou de R$ 48.730.897,83 para R$ 48.847.497,83 depois da aprovação da Fase 0.
7. Na R2, a coluna "Rafael" mostra R$ 3.500.000 para São Brás sem dizer que em 2018 o valor era "NÃO DIVISÍVEL" (sem valor). Convém uma coluna "Valor 2018" ou uma nota.

## O que conferi e está correto

- **Totais.**
  - Grupo, doc 05 2025: R$ 75.410.201,00.
  - Rafael, balanço executado (p. 19 e quadro-resumo p. 20): R$ 79.491.957,50.
  - R1: R$ 17.023.657,50 (8.100.000 + 2.500.000 + 5.560.857,50 + 440.000 + 422.800).
  - Bens escolhidos por Rafael (04, 05, 06, 08): R$ 50.914.157,50.
- **Os 5 bens da R1 têm o mesmo valor em todas as seções de Rafael.**
  - Balanço do grupo (p. 2–4).
  - Avaliação executada (p. 10, 11, 12, 14, 15).
  - Balanço executado (p. 17–19).
  - Bens escolhidos (p. 22–23, itens 04 e 06).
  - Diretrizes (p. 4–5, itens 04 e 06).
  - `ativos-avaliados-valter.xlsx` (04 e 06).
  - Nenhum desses 5 aparece com outro valor no PDF das áreas rurais, no da Transportadora, no da Agropecuária (sem balanço) ou na Introdução.
- **Concordâncias nas páginas citadas.** Item 02, p. 10 (com o texto continuando na p. 11); item 03, p. 11; item 04, p. 12; item 06, p. 14; item 07, p. 15.
- **Lista de Rafael.** 04 e 06 estão; 02, 03 e 07 não estão. O C0 é neutro nos 5, porque os valores são iguais.
- **Ressalvas existentes.** Estão corretas:
  - item 02: matrícula 1.412 com 720 + 1.080 = 1.800 m² (arq. 1 p. 8–9; arq. 11);
  - item 04: "25,00% (Vinte por cento)", com 25 % de R$ 1.064.686 = R$ 266.171,50;
  - item 06: versão antiga R$ 520.000; matrícula 11.090 com 106,53 ha vendidos aos Rigo e 4,0 ha que ficaram (doc 17 2018);
  - item 07: lote 52 em 2018 na Agropecuária (R$ 226.500).
- **R2.**
  - Matrículas, áreas e valores do grupo conferem: doc 17 2018 e doc 11 / doc 08 2025.
  - Valores de Rafael:
    - lote 12: R$ 650.000 (faixa R$ 550–850 mil, p. 10);
    - lote 89: R$ 1.500.000 (p. 17–18);
    - lote 33: R$ 250.000 (p. 19);
    - São Brás: R$ 3.500.000 e R$ 250.000 (Transportadora p. 7–8 e 14–15, com concordância expressa).
  - A frase sobre os pacotes 3 e 4 da SEI está correta (ver A5 para o que falta).
- **R3.**
  - Quantidades conferidas no docx "áreas vendidas": 40.300 − 2.606 = 37.694; 2.606 + 3.206 + 3.591 = 9.403; 3.682 ("OLEPOPLAN" está assim no original).
  - Contrato 9: o Hotel (CNPJ 89.471.270/0001-94) é o ARRENDADOR (p. 1, imagem).
  - Arq. 1 p. 6: "pro forma" e venda dos grãos "pelo CNPJ da Agropecuária". Confere.
  - E1 aprovado (ACERTOS, pacote 1, 28/10/2026). Rafael não trata soja em nenhum PDF lido.
  - Em 2018, o "TOTAL RECEBIDO" de R$ 12.058.000 incluía as 4 × 20.150 sacas de 2019–2022. Confere.
- **Pacote 2 e etapa 6.**
  - Itens 05, 01 e 08 com valores e "na lista" corretos, em ordem decrescente de valor.
  - As 8 linhas da etapa 6 conferem com o doc 17 de 2018:
    - seq. 2 a 6 marcadas "VENDIDA" (Rigo e Pellegrini);
    - seq. 9 vendida a Angeli;
    - seq. 8 (benfeitorias, R$ 300.000) e seq. 16 (trator, R$ 116.600) sem par.
  - Soma do 2018: R$ 22.775.090 (fazenda) e R$ 24.112.390 (total). Confere.
- **Nenhuma decisão aparece como aprovada.** O `log-decisoes.md` da sessão não registra decisão de valor.
