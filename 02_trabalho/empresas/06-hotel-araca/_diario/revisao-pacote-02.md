# Revisão independente — Pacote 2 (exceções), empresa 06 Hotel Araçá

Data: 28/10/2026. Revisor em contexto novo. Objeto: `pacotes/pacote-02-excecoes.md` e `_scripts/pacote_02.py`. Li direto das fontes: PDF de Rafael (26 p., texto completo; as imagens são só o logotipo), inventário 2025 (docx), inventário 2018 (docx 16) e doc 25 de 2018 (Drive), `ativos-avaliados-valter.xlsx`, arquivo 1 das respostas, p. 10–11, a planilha-mestre (abas Mestre, Item01_componentes, Totais, Observacoes), os registros da empresa, o `00_PLANO.md` e o log da SEI (Oleoplan). Não alterei nenhum arquivo além deste relatório. Meus scripts estão em `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\5fd9298d-a99a-4db0-a826-38f7fb4a324d\scratchpad\revisor-pacote2\`.

Contagem: **2 altos, 6 médios, 7 baixos.**

---

## Altos

### A1 — E1: o pacote diz que Rafael não informa o que entra nos 66 milhões, mas a p. 7 informa
- **Evidência:** PDF p. 7, título do método: "1. Avaliação Física (Terreno + Área Construída)". O cálculo: "Avaliado em R$ 66.000.000, calculando o terreno a R$ 5.000/m² e a depreciação técnica da estrutura construída de 10.500 m² a R$ 3.000/m²". 6.900 × 5.000 = 34.500.000 e 10.500 × 3.000 = 31.500.000, que somam exatamente 66.000.000. Os móveis e equipamentos (R$ 10.504.000,00 pelo grupo) ficam de fora da conta. Por outro lado, no balanço executado (p. 16–17), Rafael põe os 66 milhões ao lado da descrição completa do grupo, que inclui os móveis e equipamentos. O pacote só diz: "Rafael não diz se os móveis e equipamentos estão dentro dos 66 milhões".
- **Problema também na tabela de parcelas:** a linha "Móveis e equipamentos" mostra "—" na diferença. Mas a tabela só fecha em +16.393.176,80 se os móveis entrarem como −10.504.000,00: 27.600.000 − 702.823,20 − 10.504.000 = 16.393.176,80. Ou seja, a própria conta do pacote supõe que Rafael deu zero aos móveis, sem dizer isso.
- **Por que o Luís precisa ver:** (i) **a favor de Rafael (C1):** se só imóvel for comparado com imóvel, o grupo dá R$ 39.102.823,20 (terreno + obra + vantagem) contra os R$ 66.000.000,00 de Rafael. (ii) **Risco:** o outro lado pode ler a p. 17 e dizer que os 66 milhões já incluem tudo. Nós podemos ler a p. 7 e dizer que os móveis ficaram fora. Se a decisão não deixar isso escrito, abre espaço para dupla contagem ou omissão nas Partes 4 e 6.
- **Correção:** trocar a frase por: "O cálculo de Rafael na p. 7 é só terreno + construção (34,5 + 31,5 milhões). Na p. 17 ele atribui o valor ao item inteiro, com a descrição dos móveis." Na tabela, mostrar os móveis como "R$ 0 no cálculo da p. 7 (−10.504.000,00)". Na decisão, se for Rafael, escrever "item 01 inteiro, inclusive móveis e equipamentos, R$ 66.000.000,00".

### A2 — E4: o precedente Oleoplan foi descrito ao contrário, e o efeito no 1/6 do Luís ficou escondido
- **Evidência:** log da SEI, Pacote 2, linha 1: no Oleoplan, "As parcelas de Rafael somam o valor do grupo: o outro lado pode alegar erro de soma". Mesmo assim foi adotado o **valor declarado por Rafael**, que era o valor favorecido pelo C0, e o risco só foi registrado. No E4, o pacote faz o contrário: descarta o valor declarado por Rafael, que é o favorecido pelo C0 (total maior para a empresa), e diz que faz isso "como na SEI (imóvel Oleoplan)". Além disso, o texto diz "o erro passa para o 1/6 de Rafael". O estudo de Rafael é o da Valter Participações, e a lista dele é a "dos bens que virão para nós" (P10). Esses R$ 60.000,00 são do 1/6 do próprio Luís: não adotar a diferença reduz em R$ 60.000,00 o crédito em relação ao estudo de Rafael (R$ 10.825.000,00 → R$ 10.765.000,00, aba Totais).
- **Correção:** tirar "como na SEI". Explicar a diferença: no Oleoplan existia um bem, e o valor favorecido pelo C0 foi mantido; aqui não existe bem, e o P5 impede adotar valor sem item. Dar ao Luís as duas opções: (a) não adotar, como recomendado (P5), −R$ 60.000,00 no seu 1/6 em relação ao estudo de Rafael; (b) manter o total declarado por Rafael, seguindo o C0 e o desfecho do Oleoplan, com o risco registrado. Trocar "1/6 de Rafael" por "seu 1/6 no estudo de Rafael". Usar "diferença não identificada" em vez de "erro de soma", como ficou decidido na M4 da revisão do Pacote 1 (ver B7).

---

## Médios

### M1 — E1: o prédio só fica "quase igual" porque a vantagem de 15 %, calculada também sobre os móveis, foi posta no prédio
- **Evidência:** inventário 2025, item 01: os 15 % incidem sobre "R$ 37.136.368,00", que é a obra depreciada (26.632.368) mais móveis (4.104.000) mais equipamentos (6.400.000). A aba Item01_componentes da mestre compara só a obra civil: grupo R$ 26.632.368,00 (R$ 2.536,42/m² líquido) contra Rafael R$ 31.500.000,00 (R$ 3.000/m²), uma diferença de **+R$ 4.867.632,00 (+18 %)**. O pacote soma a vantagem inteira ao prédio, chega a −702.823,20 e conclui que "Rafael … chega a um prédio quase igual ao do grupo". O grupo chama a vantagem de "Valor do Edifício Construído – obra civil", então a leitura do pacote é defensável, mas é só uma das duas leituras possíveis.
- **Correção:** mostrar as duas leituras: só a obra (+4.867.632,00) e obra mais vantagem (−702.823,20). Avisar que a base da vantagem inclui os móveis. Tirar "quase igual", ou dizer de que leitura se trata.

### M2 — E1: faltam ressalvas que o outro lado vai usar
- **Evidência:**
  1. P. 7: a faixa "depende diretamente do objetivo final": quando o hotel continua operando, "o Método da Renda dita o preço". P. 8: o método da renda "é o método mais realista para fundos hoteleiros e investidores". O hotel fica na empresa e segue operando. Pela lógica do próprio Rafael, o método que vale é o da renda, que dá R$ 41,0 a 50,6 milhões, faixa que inclui o valor do grupo. O pacote cita a faixa da renda, mas não a frase de Rafael que dá peso a ela.
  2. A faixa de R$ 55 a 85 milhões, primeiro argumento C1, não tem conta no PDF. Ela "integra" a avaliação física (66 a 102,45 milhões) e a renda (41 a 50,6 milhões), mas a média simples das duas dá 53,5 a 76,5 milhões. O piso de 55 milhões não sai de nenhum número mostrado.
  3. O teto físico supõe "terrenos unificados frente ao mar" (p. 7), e o argumento do valor de venda fala em "incorporação imobiliária". O grupo responde que não há plano de unificação definido, nem projeto, nem habite-se (arquivo 1, p. 11). Os lotes 13 e 14 da quadra 37 nem se sabe de quem são (G1).
- **Correção:** acrescentar os três pontos em "Contra, ou ressalvas".

### M3 — Faltam argumentos C1 a favor de Rafael que estão nas fontes
- **E1:** a versão antiga do inventário 2025 (também do grupo) dava ao item 01 **R$ 53.320.460,00**, com vantagem de 25 %. A atual baixou para 15 %, R$ 49.606.823,20, e essa foi a única mudança do inventário entre as duas versões (diário, Fase 0). Em 2018 a vantagem era de 30 %. Nos mesmos anos, os lotes vizinhos subiram de 64 % a 100 % (Pares_2018_2025), enquanto o item 01 subiu 9,6 %. Ou seja, o próprio grupo já deu ao hotel um valor maior que o atual.
- **E3:** o grupo manteve R$ 1.000.000,00 de 2018 a 2025 (inventário 2018, item 19; Pares_2018_2025). No mesmo período, revalorizou todos os lotes de Capão de 64 % a 100 %.
- **Correção:** acrescentar os dois pontos em "A favor de Rafael (C1)", deixando claro que são fatos das versões do grupo, não reavaliação.

### M4 — E2: falta o trecho do doc 25 de 2018 que identifica as duas casas
- **Evidência:** doc 25 de 2018, seção "Hotel Araçá Ltda.": foram "agregados ao capital do Hotel, a compra de 23 lotes urbanos …, uma casa e manutenção das 02 casas de veraneio dos Sócios, com o pagamento das despesas de manutenção das mesmas durante o ano todo". Logo depois: "CASA DA ESQUINA: Élcio Siviero (que pertence à empresa Siviero)" e "CASA AO LADO: Raul Siviero". Isso bate com o que Rafael descreve (duas casas em lotes em sequência, na esquina) e com os itens 22 e 23 de 2018. É a pista mais forte de identidade e de titularidade, e indica despesa do Hotel em favor de sócios. O pacote só cita a troca do lote 10/q03 do Aditivo Medianeira.
- **Destino incoerente:** se as casas são os itens 22 e 23 de 2018 ("Só 2018", que vão para a Parte 4), não podem ir também para a Parte 5 como "não abordados".
- **Correção:** citar o trecho no E2 e na G8 ("pertence à empresa Siviero": qual empresa, e quem é o titular hoje). Pôr a despesa de manutenção na lista de candidatos da Parte 4/6. Definir um destino só: Parte 4 se forem os itens 22 e 23.

### M5 — A Fiat Toro ficou de fora do pacote e sem destino
- **Evidência:** PDF p. 25, na mesma "Observação importante" das casas: "E uma Fiat Toro automóvel de serviço do Hotel Araçá Ltda.", que deve "entrar no escopo negocial ou ter a sua destinação … comprovada" (p. 25–26). A mestre tem a linha "Só Rafael". A R3 do Pacote 1 só mandou ao pacote de exceções as "residências citadas por Rafael", e a G8 já pergunta pela Toro. Não há valor de nenhum dos lados, mas o bem precisa de destino escrito.
- **Correção:** incluir a Toro no E2 (ou numa linha E5): sem valor; G8; se for da empresa em 31/12/2024, volta com valor; se não for, vai para a Parte 5.

### M6 — Os registros de data-base 1 e 2 continuam "a decidir no pacote de exceções"
- **Evidência:** `registro-data-base.md`, linhas 1 e 2, coluna "Tratamento adotado": "a decidir no pacote de exceções". O pacote cita os dois registros como argumento, mas não propõe o tratamento. Se o Luís aprovar o pacote, os registros continuam abertos (P2).
- **Correção:** propor o tratamento de cada um. Exemplo: "nº 1: mantido sem ajuste; os valores de Rafael são adotados como estão (P5); fuga registrada". "Nº 2: mantido sem ajuste; o CUB de dez/2022 subavalia a obra na data-base, o que reforça o C1". Pedir a aprovação junto.

---

## Baixos

- **B1:** "os lotes 03, 21 e 22 do mesmo quarteirão" (E1) e "lotes 03 a 10" (E4) são números de **item**. Os lotes são 16, 15 e 17 da q37 (itens 03, 21 e 22) e 16, 23–29 (itens 03–10). E os lotes 21 e 22 da q37 fazem parte do próprio item 01 (mat. 34.769 e 34.770), o que deixa a frase errada se lida ao pé da letra. Correção: "itens 03, 21 e 22 (lotes 16, 15 e 17 da q37)".
- **B2:** "lotes de 360 m² da quadra 38, atrás do hotel": a fonte não diz "atrás" (os "Fundos do Hotel" são os itens 03–10). Acrescentar que o grupo dá os mesmos R$ 1.000/m² aos lotes de fundos (itens 03, 04 e 10) e R$ 972,22/m² aos itens 18 e 19. O argumento continua de pé: terreno unificado de frente para o mar pelo preço de lote de fundos.
- **B3:** E1, "Duas leituras de Rafael": falta o contraponto. A concordância expressa (p. 12 e p. 16) cobre só os "itens 02 à 19", não o 01 nem o 20, o que enfraquece a tese de que Rafael aceitou o valor do grupo. O mesmo vale para o E3.
- **B4:** E3: a faixa está na p. 13, mas a adoção do piso está na p. 19. Falta ressalva: Rafael aplica o preço por m² de lotes padrão de cerca de 300 m² (p. 13–14) a um lote de 725,39 m² sem fator de área. E a p. 15 liga o piso de R$ 2,4 milhões a "casa … mais antiga, necessitando de reformas", o que sustenta o piso como valor conservador (C1).
- **B5:** o arquivo 1, p. 11, diz que o poço com outorga "não está sendo utilizado", e ele está dentro dos equipamentos de R$ 6,4 milhões. O arquivo diz também que o PPCI "está aprovado e em vigor". O pacote diz só "não há projeto, habite-se nem profissional contratado", o que está certo, mas fica incompleto.
- **B6:** o script fixa no texto valores que não lê da mestre nem do PDF: as opções "R$ 66.000.000,00" e "R$ 49.606.823,20", "R$ 1.000/m²", "R$ 5.000/m²", 6.900 m², "R$ 2.400.000 a 3.500.000 (p. 13)", 2.000.000, 1.601.357,90, 12.000/10.500, 160/158,18 e 750.000. Hoje todos conferem, mas não se atualizam se a mestre mudar.
- **B7:** E4 chama a diferença de "erro de soma". A revisão do Pacote 1 (M4) fixou "diferença não identificada", e o próprio pacote admite outra leitura (um lote não listado). Manter "não identificada".

---

## O que conferi e achei certo

- **Valores do grupo:** item 01 R$ 49.606.823,20, componentes 6.900.000 / 1.512.000 / 2.592.000 / 26.632.368 / 6.400.000 / 5.570.455,20; item 20 R$ 1.000.000,00; total R$ 58.076.823,20 (docx 2025 e reprodução na p. 1–5 de Rafael).
- **Valores de Rafael:** 66.000.000 (p. 17), 2.400.000 (p. 19), total declarado 76.230.000 (p. 20), 1/6 de 12.705.000, crédito de 10.825.000 e lista de 1.880.000 (p. 21). Soma dos 22 itens = 75.870.000; diferença = 360.000.
- **Faixas e páginas:** 55–85 milhões (p. 7); físico 66 a 102,45 milhões (6.900 × 8.000 + 10.500 × 4.500 = 102.450.000); renda 41 a 50,6 milhões (p. 8); premissas 140 quartos, R$ 520, 45 %, 30 % (p. 9); R$ 1.050–1.500/m² (p. 11); Condado 2,4–3,5 milhões (p. 13), terreno 1,81–2,75 milhões e casa 600–750 mil (p. 14); residências 2,4–3,5 milhões pelo par (p. 24); repetição do valor do grupo na p. 6 (item 01) e na p. 13 (item 20).
- **Contas:** renda pelas premissas: 140 × 520 × 365 × 0,45 × 0,30 = 3.587.220,00, que dá 34.164.000,00 a 10,5 % e 42.202.588,24 a 8,5 %. Diferenças +16.393.176,80 e +1.400.000,00; no 1/6, 2.732.196,13 e 233.333,33. Soma final 7.470.000 + 66.000.000 + 2.400.000 = 75.870.000,00; +17.793.176,80 contra o grupo; 2.965.529,47 no 1/6. Nesga: 420.000 / 2.100.000. Lote da q38: 1.083,33/m². 1/6: 12.705.000 contra 12.645.000; diferença de 60.000.
- **Direção do C0:** a lista de Rafael para o Hotel tem só os itens 02, 11, 12, 18 e 19 (xlsx, linhas 5–9; p. 21–22). Os itens 01 e 20 ficam na empresa, então o C0 pede o maior (Rafael), e as recomendações de E1 e E3 seguem o C0.
- **Exceções:** na mestre só os itens 01 e 20 estão como "Discorda". Os outros valores de Rafael no PDF são faixas gerais, já tratadas no Pacote 1. Falta só a Toro (M5).
- **Outros pontos:**
  - Rafael não desconta INSS nem regularização; só cita genericamente "obrigações fiscais, eventuais ônus e gravames", na p. 26.
  - As fontes de Rafael não trazem nenhuma venda de terreno a R$ 5.000/m².
  - A reportagem do Jornal do Comércio é de dez/2025.
  - CUB de dez/2022 (registro 2).
  - A área de 12.000 m² (2018) passou a 10.500 m² (2025).
  - Contrato particular de maio/2014 por R$ 750.000,00.
  - O grupo afirma que todos os imóveis estão em nome da empresa (arquivo 1, p. 11).
  - Itens 22 e 23 de 2018 ("CASA" e "CASA/ESQUINA") e a troca do lote 10/q03 no doc 25.
  - md5 da mestre `bcaf0b225e`.
  - Para o hotel, o tratamento como exceção, com opções (a) e (b) e sem terceira via (P5), está de acordo com o brief.
