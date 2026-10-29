# Revisão independente de fechamento da Parte 3 (SEI)

28/10/2026. Revisor com contexto novo (P11). Objeto: todas as decisões aprovadas da Parte 3 (pacotes 1 a 5, correções e decisões após as revisões), `decisoes-sei.json`, `log-decisoes.md` e `planilha-mestre-sei.xlsx`.

Método:
- li o plano (P1 a P13 e §7), o log, o JSON, o script e as quatro revisões anteriores;
- regenerei a planilha-mestre numa cópia temporária;
- conferi os números contra as fontes:
  - `sei-2025-itens.xlsx` (Itens, Lotes, Passivos, Totais, Texto, Socios);
  - texto dos 26 PDFs de Rafael, extraído com PyMuPDF;
  - `ativos-avaliados-valter.xlsx`;
  - `geofarroupilha-app-lotes.xlsx`;
  - arquivo 17 (xlsx, lotes vendidos de jan./2018 a out./2025);
  - `diff-inventario-2025-antigo-x-atual.md`;
  - `analises/acertos-e-areas-vendidas/achados-para-sei.md` e `cruzamentos.xlsx`.

Não alterei nenhum arquivo do projeto além deste. Os scripts de conferência e a cópia regenerada ficaram em `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\a26375ee-fe81-4a97-9538-f335e283012f\scratchpad\revisao-fechamento\`.

**Resumo.**
- As contas fecham:
  - a planilha se reproduz célula a célula;
  - os totais do log batem;
  - o C0 foi aplicado no sentido certo em todos os itens divergentes;
  - não há dupla contagem dentro da soma adotada.
- O problema grave é de **cobertura pela data-base**. O inventário "31.12.2024" mostra o estoque do fim de 2025:
  - **faltam 27 lotes** que eram da SEI na data-base (achado 1);
  - **entraram 2 bens** que só chegaram em setembro de 2025, entre eles um bem da lista de Rafael (achado 2).
- A soma adotada e a lista de pendências da ficha ainda não refletem isso.

---

## Achados

### 1. ALTA. O doc 11 omite 27 lotes que eram da SEI em 31.12.2024; a Parte 3 não os registra nem como pendência

- **O que está errado.** O arquivo 17, que é do próprio grupo, tem 36 linhas com data posterior a 31.12.2024. Cruzei cada uma com o estoque do doc 11 (aba Lotes), com a tabela 47 (lotes de terceiros) e com a obs. 1. Resultado: **27 lotes vendidos, permutados ou dados em pagamento depois da data-base não aparecem em lugar nenhum do doc 11**. Pela P2 eles eram ativo da SEI em 31.12.2024. Saíram sem registro: o grupo apresentou o estoque do fim de 2025 como se fosse o da data-base.
- **Evidência:** arq. 17, aba "Lotes vendidos (PDF)" (data, quadra, lote, comprador):

| Loteamento | Lotes (quadra/lote) | Data | Destino no arq. 17 |
|---|---|---|---|
| Salvador – Femaçã II | 01/08 | 10/09/2025 | permuta Dal Ponte "R$ 1.010.000" |
| Salvador – Femaçã II | 03/05 | 01/07/2025 | acerto com Bertocco |
| Colina do Retiro I | 14/07 | 19/02/2025 | pagamento de serviços |
| Colina do Retiro I | 14/08 e 14/09 | 10/09/2025 | permuta Dal Ponte |
| Colina do Retiro II | 05/15 | 01/08/2025 | à vista |
| Solar Norte | 01/03 e 02/01 | jun. e jul./2025 | — |
| Baldasso | 06/200, 06/350 e 48/137 | fev. a jun./2025 | — |
| San Marino | 1673/06 | 05/01/2025 | — |
| Firminio Dalzóchio | 910/03 | 01/04/2025 | — |
| Vila Vicenza | 910/06 | 21/07/2025 | "comprador querendo rescindir" |
| Viero | 1469/10 | 21/02/2025 | — |
| Morada dos Ipês | 03/002 | 11/02/2025 | — |
| Alto das Graças | 6366/18 | 04/08/2025 | — |
| Lago Negro | 05/14 | 28/08/2025 | permuta Candiago (registro 5) |
| Barbiero | 05/03 e 05/08 | 23/05/2025 | S6 |
| Falavigna | 174/020, 021, 023; 175/012, 013; 876/016 | 2025 | S6 |
| Falavigna | 322/028 | 02/06/2026 | — |

  - O próprio grupo é incoerente. O Monte Vêneto 04/006 (vendido em 28/03/2025) e o Falavigna 876/008 (vendido em 09/04/2026) **continuam** no estoque do doc 11 (aba Lotes, linhas 344 e 424). Os demais foram tirados.
  - Também saíram entre as duas versões do doc 11, ambas "31.12.2024":
    - o Morada do Ipê I 001/q2, de R$ 538.000 (diff, §13.7, item 03);
    - o Lago Negro 014/05 (diff, item 18).
  - Os lotes do Ipê I 02/010 e 02/013 (R$ 1.087.510) estão na obs. 1 como "PAGO" ao Rovílio. Pelo arq. 17, foram entregues em 28/08/2025.
  - Dos 27 lotes, só 8 estão no S6 aprovado, e o S6 não foi lançado no `registro-data-base.md`. O Lago Negro 014/05 está nas linhas 5 e 9 do registro. Os outros 18 não estão registrados em lugar nenhum.
- **Valor em jogo (ordem de grandeza, não é avaliação).** Pela mediana dos lotes vizinhos no próprio doc 11, os 27 lotes somam cerca de **R$ 9,2 milhões (1/6 ≈ R$ 1,53 milhão)**. Somando o Ipê I 001/q2 e os 2 lotes do Rovílio, são mais R$ 1,6 milhão. A omissão vai contra o Luís, porque reduz o patrimônio que fica.
- **Correção sugerida:**
  - abrir no registro de data-base uma linha por lote (ou por grupo de lotes), com a data do arq. 17;
  - incluir os 27 lotes na ficha de fechamento como pendência explícita da Parte 3;
  - perguntar ao grupo a data do contrato de cada venda (escritura × compromisso) e o valor de cada lote em 31.12.2024;
  - nenhum dos dois lados avalia esses lotes, então a P5 exige o valor do grupo, se ele o der, ou uma decisão de "caso extremo";
  - corrigir a frase do log (situação depois do pacote 5) "itens de 2025 ainda sem valor adotado: os 7 loteamentos contratados".

### 2. ALTA. O lote urbano 06 (Dal Ponte, R$ 1.010.000, na lista de Rafael) e o Reserva Paradiso (R$ 500.000) entraram depois da data-base e foram adotados sem registro

- **Evidência:**
  - Arq. 17, linhas 31, 147 e 148: Salvador 01/08 e Colina do Retiro I 14/08 e 14/09 foram "Vendido à vista, com permuta de outro imóvel, valor de R$ 1.010.000,00", com a Família Dal Ponte, em 10/09/2025. O lote urbano 06 do doc 11 é "adquirido da Família Dal Ponte", matrícula 3.788, e vale exatamente R$ 1.010.000.
    - Ele não existia em 2018 (aba Mestre, linha 36: "Existe em 2018? Não").
    - A sessão do fluxo de caixa já o trata como "o lote da Dal Ponte de 2025" (`analises/parte4-fluxo-caixa/_diario/revisao-v0.md`, B4).
    - A pergunta 8 ao grupo ("datas de compra") já tem resposta nas fontes.
  - Arq. 17, linha 64: o Reserva Paradiso 104/623 foi "Adquirido, como forma de pagamento" da venda do Vicentino 1301/003, em 19/09/2025. Foi revendido em 07/01/2026 (`cruzamentos.xlsx`, aba Datas). O lote é novo na versão atual do doc 11 (diff, item 31).
  - Na planilha-mestre:
    - o LURB-06 está na R2, marcado "Sim" na lista de Rafael (Mestre, linha 36);
    - o LOT-31 está na R1 (linha 85);
    - os dois estão sem registro de data-base.
- **Valor em jogo.** São R$ 1.510.000 na soma adotada (1/6 = R$ 251.666,67).
  - Mais importante: **Rafael escolheu para nós um bem que a SEI não tinha na data-base**, e o grupo pode usar isso contra a lista inteira.
  - Há risco de contar duas vezes. Se o lado do ativo do S1 for lançado (Vicentino 1301/003, nível B pela R6: R$ 677.500) sem tirar o Reserva Paradiso, a mesma troca entra duas vezes.
- **Correção sugerida:**
  - registrar os dois bens no registro de data-base;
  - marcar o LURB-06 e o LOT-31 como provisórios;
  - decidir o tratamento junto com o achado 1, para que a troca entre só uma vez: os 3 lotes permutados pelo Dal Ponte, e o Vicentino 1301/003 contra o Reserva Paradiso;
  - levar ao Luís o efeito sobre a lista de Rafael (LURB-06).

### 3. MÉDIA. Os arquivos do projeto contradizem a decisão do Oleoplan, e o valor adotado não é uma avaliação

- **Evidência:**
  - `pendencias-conferencia-anterior.md`, item 3: "**Esclarecida** [...] O R$ 10.631.378,60 é erro de transcrição de Rafael. Tratamento: adotar R$ 13.631.378,60 (valor do grupo)".
  - `sei-2025-notas.md`, linha 133: "O inventário 11 resolve a diferença [...] a favor da soma das parcelas".
  - A decisão aprovada é a contrária: R$ 10.631.378,60 (log, pacote 2, item 1, e "Decisões após a revisão").
  - No PDF Diretrizes (p. 8-9), Rafael copia o texto do grupo, inclusive o CUB de agosto/2017, e imprime um total R$ 3.000.000 menor que a soma das parcelas. Não há faixa de mercado nem avaliação própria. A P5 manda adotar "uma das duas avaliações".
- **Valor em jogo:** R$ 3.000.000 (1/6 = R$ 500.000). Se esses arquivos circularem, entregam ao grupo, por escrito e do nosso lado, o argumento do erro de transcrição.
- **Correção sugerida:**
  - atualizar a pendência 3 e a nota para a decisão do Luís, com o risco anotado;
  - na ficha de fechamento, listar o Oleoplan como decisão sobre premissa frágil, porque o valor de Rafael não é avaliação.

### 4. MÉDIA. Áreas rurais 11 e 12: o valor de Rafael ignora dois ônus que o próprio texto registra

- **Evidência:** no doc 11 e na transcrição de Rafael (PDF Áreas rurais, p. 20-21 e 24), os dois imóveis estão "ainda não registrados em nome da empresa" (cessão de direitos de posse) e "Parte da área foi destinada para compensação ambiental do Loteamento Colina do Retiro II".
  - A faixa de Rafael, de R$ 250.000 a 1.200.000, é genérica para 7,5 ha na Linha Ernesto Alves, com anúncios e VTN de 2026, e não trata nenhum dos dois ônus.
  - O motivo aprovado (pacote 3) cita só o VTN.
  - Em outros bens, uma restrição ambiental documentada levou a 50 % ou 0 % (R6, C2). Aqui levou ao valor cheio de Rafael.
- **Valor em jogo:** + R$ 710.000 sobre o valor do grupo (1/6 = R$ 118.333,33).
- **Correção sugerida:**
  - registrar os dois ônus no motivo e na ficha;
  - pedir ao grupo a área destinada à compensação e a situação da posse;
  - se a área de compensação for relevante, o Luís decide se a lógica do C2 vale também aqui.

### 5. MÉDIA. Decisões aprovadas ainda não refletidas na planilha-mestre

- **Dívida do trator (R$ 320.000).** A decisão do pacote 2, item 7, exclui o bem **e** a dívida. O JSON só grava o bem (`MAQ-14`, valor adotado 0). A aba Passivos, linha 37 (obs. 21, Paraná Equipamentos), continua com R$ 320.000 e sem valor adotado.
  - Quem calcular o patrimônio líquido pela planilha tira R$ 450.000 em vez de R$ 130.000: R$ 53.333,33 a mais contra o Luís no 1/6.
  - Correção: acrescentar o passivo em `passivos` do JSON, com valor adotado 0.
- **S4, aprovado:** "Tirar o Q6371 L11 da lista de bens a receber até esclarecer". Na aba Lotes, linha 408 (Alto das Graças I, H-6371/011, R$ 141.000), o lote continua "Sim" na lista de Rafael.
- **S6, aprovado:** ainda não está em `registro-data-base.md` (ver achado 1).

### 6. MÉDIA. A lista de pendências do log e do Resumo está incompleta

- **Evidência:** o log (situação depois do pacote 5) e a aba Resumo dão "7 itens sem valor adotado", só os loteamentos contratados. Ficam de fora:
  - os 6 lotes "FAMÍLIA SIVIERO" do Salvador, quadra 05 (Lotes, linhas 77-82): valor adotado vazio, pergunta 1; os vizinhos valem cerca de R$ 2,29 milhões;
  - os 8 lotes da obs. 1, do Rovílio (Lotes, linhas 466-473): sem decisão nenhuma;
  - o lado do ativo do S1 (registro 7): Vicentino 1301/003, nível B, R$ 677.500; Jardim Itália 1747/001 e Primavera 1199/010 e 011, nível A, R$ 0;
  - o Lago Negro 014/05 e os itens dos achados 1 e 2.
- **Correção sugerida:**
  - o Resumo deve contar também os lotes pendentes;
  - a ficha de fechamento deve listar todos os itens acima;
  - os 8 lotes do Rovílio devem receber a decisão explícita "Parte 4".

### 7. MÉDIA. Registro 4 (Rovílio): "Não altera valores" está errado

- **Evidência:**
  - Doc 11, obs. 1: o contrato é de 28/08/2025, posterior à data-base. O grupo deduz R$ 1.650.000 como passivo (3 lotes do Ipê II "A PAGAR"; Passivos, linha 17) e tira do estoque 5 lotes "PAGO" (R$ 1.654.582).
  - Arq. 17, linhas 299-300: 2 desses lotes foram entregues em 28/08/2025.
  - Na data-base, o Rovílio ainda era sócio (Socios, quadro real, 2 %). O pagamento das quotas dele saiu de bens da SEI depois da data-base.
- **Valor em jogo:** até cerca de R$ 3,3 milhões no patrimônio de 31.12.2024 (1/6 ≈ R$ 550 mil), contra o Luís.
- **Correção sugerida:** reescrever a justificativa do registro 4 ("altera valores: passivo e lotes da obs. 1") e manter o tratamento como pendente, para a Parte 4.

### 8. MÉDIA. O ganho tirado de Rafael nos bens que ficam depende de avaliações de 2026, sem laudo

- **Evidência:**
  - Rafael usa VTN de 2026, anúncios e Instagram (PDF Áreas rurais, p. 8-21; Lotes urbanos, p. 3-9; Lotes loteamentos, p. 79).
  - O registro 3 ("Estudo 2 de Rafael, set./2026") está "*pendente*".
  - Itens adotados por Rafael que ficam na empresa:
    - 7 áreas rurais: + R$ 3.254.800;
    - LURB-02 e LURB-07: + R$ 220.000;
    - Lago Negro 002 a 005: + R$ 1.149.000.
  - Total: **+ R$ 4.623.800 (1/6 = R$ 770.633,33)**. É o primeiro alvo do grupo, pela P2 e pela falta de laudo.
- **Correção sugerida:** decidir o registro 3 antes de fechar a Parte 3 e listar esses itens na ficha como dependentes da premissa "valor de mercado de Rafael aceito como valor de 31.12.2024".

### 9. MÉDIA. Lago Negro: a base da "APP confirmada" e os lotes que eram de compradores

- **Evidência:**
  - Doc 11, obs. 6: "Está em trâmite, basicamente com decisão já proferido pelo Ministério Público, com acato da Juíza". A decisão não está nas fontes.
  - Arq. 17, linha 193: o lote 14/05 foi permutado "com o lote 18, da quadra 11B – APP" em 28/08/2025. O lote 18/11B só voltou à SEI depois da data-base. O lote 19/11B foi indenizado por R$ 580.000 em 14/10/2025 (registro 9).
  - O doc 11 lista os 7 lotes da quadra 11B como estoque da SEI.
- **Por que importa:**
  - Na data-base, pelo menos os lotes 18 e 19 eram de terceiros.
  - Pela extensão da R6 ao passivo (nível A: indenização de 100 %), que nós mesmos aprovamos, o grupo pode pedir provisão para eles.
  - A cobrança de R$ 7.700.000 de Rafael também conta lotes que não eram da SEI.
- **Correção sugerida:**
  - na ficha, marcar o C2 do Lago Negro como dependente da cópia da decisão;
  - anotar que a extensão da R6 ao passivo pode trazer de volta uma provisão para os lotes 18 e 19 (Parte 4).

### 10. BAIXA. Trator e financiamento da obs. 20: outros efeitos de data-base

- A obs. 21 diz "entrada já paga no valor de R$ 130.000". Excluir o bem e a dívida elimina também os R$ 130.000 que saíram da SEI, e o efeito líquido de menos R$ 130.000 (1/6 = R$ 21.666,67) cai sobre o Luís. Se a entrada saiu de recursos que já estavam na data-base e que não foram inventariados, o efeito deveria ser zero. Registrar no item 6 do registro.
- A obs. 20 (Banco Caterpillar) informa o saldo em 31/07/2025, 33 parcelas de R$ 10.534,91. Em 31/12/2024 faltariam mais 7 parcelas, cerca de R$ 73.744 a mais. Isso favorece o Luís, mas foge da P2, e o outro lado pode levantar. Fica para a Parte 4.

### 11. BAIXA. Aditivo Medianeira 07/002 vendido antes da data-base e ainda no estoque

- Arq. 17: vendido em 24/04/2024 à Solar Serra, "como forma de pagamento placas solares Hotel Araçá". No doc 11 está no estoque por R$ 165.000 (R1; Lotes, linha 37). A obs. 14 tem um passivo de R$ 150.000 com a Solar Serra ("acerto em lote").
- Ou o ativo e o passivo valem os dois (lote ainda não entregue), ou os dois saem. Há também um pagamento de outra empresa do grupo com lote da SEI (Parte 6, forma 9).
- Correção: perguntar ao grupo e anotar na Parte 6.

### 12. BAIXA. Argumento a favor não registrado: indenização do Vicentino 1303/001

- O lote mede 1.141,5 m². Os outros lotes de terceiros do Vicentino medem de 2.555 a 4.010 m² (`geofarroupilha-app-lotes.xlsx`). Todos estão listados por R$ 1.355.000.
  - Dá cerca de R$ 1.187/m², contra cerca de R$ 400/m² nos lotes próprios 1303/007 a 010.
  - Rafael dá ao Vicentino a faixa de R$ 980.000 a 1.125.000 (p. 110).
- Pela R6 (nível A), a indenização adotada é de R$ 1.355.000. Pela P5 não se reavalia, mas vale registrar o argumento para uma contestação.

### 13. BAIXA. Leitura e documentação

- Aba Lotes, coluna "Valor Rafael": está vazia nos San Marino 012, 013 e 004, no Vicentino 013 e nos Primavera 008 e 009. Para esses lotes, o motivo diz "grupo = Rafael" ("Valor em só uma das fontes"). Preencher a partir das tabelas de Rafael.
- Aba Legenda: "Decisão (P5) / Motivo — A preencher" e "Destino — A preencher na Parte 4" estão desatualizadas.
- `00_PLANO.md`, §7: ainda diz "pacotes 1 a 4", "R$ 153.970.060,60" e "Simioni (reaberto)".
- Pendências que a ficha deve lembrar, já registradas em outros lugares:
  - escavadeira 320D, MAQ-03, R$ 200.000: na SEI e na Ceval, com 1/3 anotado como de terceiro (S12, ainda não aprovado);
  - área rural 10, dentro do Cecconello (Parte 6, item 3), quando a Parte 5 decidir os contratados;
  - área rural 06: 6.700 m² da parcela Frainer fora da área declarada (pergunta 4);
  - Alto das Graças I: parceria Villa Luna de 50 % em 2018, mantida sem pergunta por decisão do Luís.

---

## Decisões sobre premissa frágil (para a ficha de fechamento)

| Decisão | Premissa frágil | Valor em jogo (1/6) |
|---|---|---|
| Oleoplan por Rafael | O valor é erro de soma, não avaliação (achado 3). | R$ 500.000 |
| 7 áreas rurais, LURB-02/07 e Lago Negro 002 a 005 por Rafael | Avaliação de 2026, sem laudo; registro 3 pendente (achado 8). | R$ 770.633,33 |
| Áreas rurais 11 e 12 | Posse não registrada e compensação ambiental (achado 4). | R$ 118.333,33 |
| Lago Negro, 7 lotes a R$ 0 (C2) | Decisão "em trâmite", sem cópia; lotes 18 e 19 de terceiros na data-base (achado 9). | a apurar |
| R6, leitura intermediária | Texto da Lei 4.814/2023 não obtido; mapa de 2026 (registro 10). | Leitura estrita: + R$ 4.513.000 no patrimônio |
| Simioni, nível B | Sem cópia do TAC nem localização da nascente. | R$ 43.333,33 contra R$ 0 |
| LURB-06 e Reserva Paradiso | Bens posteriores à data-base (achado 2). | R$ 251.666,67 |
| Estoque de lotes | O doc 11 mostra o estoque do fim de 2025 (achado 1). | ≈ R$ 1,5 milhão |
| Trator | Entrada de R$ 130.000 (achado 10). | R$ 21.666,67 |
| Pendentes sem valor | Lotes "FAMÍLIA SIVIERO", 7 contratados (Parte 5), lado do ativo do S1, lotes da obs. 1. | — |

---

## Conferido e correto

1. **Teste 1, regeneração.** `build_mestre_sei.py` com `MESTRE_OUT` na pasta temporária reproduz a planilha oficial célula a célula, sem nenhuma diferença nas 6 abas (Mestre 154×41, Lotes 473×23, Loteamentos 83×13, Passivos 43×9, Resumo, Legenda). O md5 difere só pelos metadados do arquivo. Resumo:
   - soma adotada: R$ 155.157.560,60;
   - 7 itens sem valor adotado (CONTR-01 a 07);
   - passivo de terceiros adotado: R$ 8.570.000;
   - grupo: R$ 152.796.260,60, igual ao total da tabela 41 do doc 11 (aba Totais: CONFERE).
2. **Soma adotada = soma dos itens.** Os 93 itens de 2025 somam R$ 155.157.560,60. Os lotes não entram na soma, o que evita dupla contagem entre item e lotes. Nos 31 loteamentos, a soma dos lotes é igual ao valor do item, tanto pelo grupo quanto pelo adotado (Lago Negro R$ 5.170.000; San Marino R$ 2.445.000; Vicentino R$ 5.982.500; Simioni R$ 260.000).
3. **Somas por regra.**
   - R1: 395 lotes em 28 loteamentos, R$ 103.459.000. Hoje são 24 itens com a R1 (R$ 94.021.000) mais 4 itens do pacote 5 pelo subtotal do grupo (R$ 9.438.000).
   - R2: 12 itens, R$ 8.560.666.
   - R3: 30 itens, R$ 16.515.016 (veículos R$ 1.150.000; máquinas R$ 2.067.000; áreas urbanas 01 a 03 R$ 13.298.016).
   - Total do pacote 1: R$ 128.534.682.
4. **Totais do log, refeitos passo a passo.**
   - Depois do pacote 2: R$ 148.635.060,60.
   - Depois do pacote 3: R$ 153.970.060,60 (+ R$ 5.335.000 das 7 áreas, antes sem valor).
   - Depois do pacote 5: R$ 155.157.560,60.
   - R6: + R$ 6.970.500 = ativo + R$ 1.187.500 e passivo − R$ 5.783.000 (R$ 14.353.000 → R$ 8.570.000); 1/6 = R$ 1.161.750.
   - S1: R$ 1.887.000; 1/6 = R$ 314.500.
   - Custo frente ao C0: R$ 1.653.500; 1/6 = R$ 275.583,33.
   - Passivo por loteamento: Vicentino R$ 6.775.000; Jardim Itália R$ 586.000; Perini R$ 560.000; Dalzóchio R$ 400.000; Primavera R$ 249.000.
5. **C0, sentido certo em todos os itens divergentes:**
   - 7 áreas rurais, LURB-02 e LURB-07: fora da lista, recebem o maior valor;
   - AURB-04: na lista, recebe o menor;
   - Lago Negro 001/01: na lista, valor do grupo (R$ 770.000 contra R$ 1.100.000 de Rafael);
   - Lago Negro 002 a 005: ficam na empresa, R$ 1.100.000.
   - O "Imóvel urbano 02" da lista de Rafael é a área da Família Lago (ARUR-02, matrícula 33.516), não a Casa da Nona (`ativos-avaliados-valter.xlsx`, linha 21). O LURB-02 está, portanto, corretamente fora da lista.
   - Marcas da lista: 91 lotes "Sim", 5 "só nas tabelas" e 7 bens fora de loteamento.
6. **Valores de Rafael, conferidos no texto dos PDFs:**
   - Áreas rurais: total R$ 10.918.666, igual à soma adotada do bloco; faixas citadas no pacote 3 (p. 8, 10, 19-21); VTN de R$ 130.615 e R$ 120.000.
   - Lotes urbanos: total R$ 3.510.000, igual à soma adotada do bloco; Casa da Nona de R$ 220.000 a 450.000 (p. 3); Ouro Verde II de R$ 135.000 a 190.000 (p. 8).
   - Lago Negro: R$ 1.100.000 por lote e R$ 5.500.000 (p. 81-83).
   - Oleoplan: R$ 10.631.378,60 (Diretrizes, p. 9).
   - Contratados (Quadro_Resumo): valores do log, Dr. Letti e Sol Nascente III em R$ 0.
7. **Níveis da R6** contra `geofarroupilha-app-lotes.xlsx`, leitura intermediária: os 33 lotes e os 4 do Simioni estão no nível certo, com os percentuais do motivo. O Vicentino 1301/003 (S1) daria nível B (100 % na faixa de 30 m), como diz o registro 7.
8. **Sem dupla contagem dentro da soma adotada:**
   - nenhum lote de terceiros (tabela 47) e nenhum lote da obs. 1 coincide com lote do estoque próprio;
   - os 4 lotes do S1 estão só no passivo, com valor 0;
   - os 11 lotes próprios da R6 entram uma única vez, pelos itens LOT-04, 20, 21, 22 e 24.
   - O risco existe para o futuro: Reserva Paradiso × Vicentino 1301/003 (achado 2).
9. **Pacote 4.** Os pares do JSON estão na Mestre: ARUR-02 ↔ 2018 #03; ARUR-06 ↔ #08; ARUR-04 ↔ OUTRAS s143; ARUR-11 e 12 sem par. O par errado ARUR-11 × #08 foi desfeito. Os destinos dos 12 itens de 2018 aparecem na coluna "Destino 2018→2025".
10. **Trator.** O texto do doc 11 (máquinas, item 14: "adquirido em setembro de 2025") sustenta a exclusão do bem.

## Não conferido

- Imagens dos PDFs: a cor da seleção de Rafael não foi revista.
- Valores dos 27 lotes do achado 1: só ordem de grandeza, pelos vizinhos.
- Datas de contrato das vendas de 2025: o arq. 17 traz uma data por linha, sem dizer se é a do compromisso ou a da escritura.
- Planilhas Google (Parte 7).
