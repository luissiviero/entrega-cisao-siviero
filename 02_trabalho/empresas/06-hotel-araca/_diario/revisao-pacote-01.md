# Revisão independente — Pacote 1 (regras em bloco), empresa 06 Hotel Araçá

Data: 28/10/2026. Revisor com contexto novo. Li direto das fontes: inventário 2025 (docx), inventário 2018 (doc 16, Drive), PDF de Rafael (26 p., texto extraído página a página com pymupdf), `ativos-avaliados-valter.xlsx` (linhas "06 - Hotel Araçá"), arquivo 1 de respostas (p. 10–11), `00_PLANO.md`, a planilha-mestre (abas Mestre, Pares_2018_2025, Totais, Legenda), `perguntas-ao-grupo.md`, `parte6-candidatos.md` e `registro-data-base.md`. **Não rodei `pacote_01.py`**, porque ele roda `build_mestre_06.py` e regrava a planilha-mestre e o pacote (ver B5). Os scripts de conferência estão no scratchpad (`revisor-pacote1/`).

Resultado: **1 alto, 6 médios, 9 baixos.** Nenhum valor, soma, contagem, página ou par 2018×2025 errado. Os problemas estão nas ressalvas e no enquadramento de R2 e R3.

---

## Achados

### A1 (alta). A ressalva "Abaixo da faixa" diz que os itens 18 e 19 ficam na empresa. Eles estão na lista de Rafael.
- **Evidência.** O pacote diz: "Nos itens 18, 19 (R$ 350.000,00); 03 … 10 (R$ 360.000,00), o valor fica abaixo do piso … Esses bens ficam na empresa, mas sem reavaliação própria (P5) não há valor maior disponível." Mas os itens 18 e 19 estão em `ativos-avaliados-valter.xlsx` (linhas 8 e 9) e no quadro de bens escolhidos do PDF (p. 21–22, "02., 11., 12., 18. e o 19"). O próprio pacote os marca como "lista de Rafael" na lista do R1.
- **Efeito.** O C0 aparece invertido para 2 dos 5 bens que vêm para nós. Para 18 e 19, estar abaixo do piso de Rafael *favorece* o C0 (menor valor para o que vem). É exatamente o argumento técnico C1 que o P10 manda registrar, e o pacote o apresenta como desvantagem.
- **Correção.** Dividir a ressalva em duas:
  - (a) itens 03–10, que ficam na empresa: valor abaixo do piso, sem valor maior entre os dois lados;
  - (b) itens 18 e 19, da lista de Rafael: o valor abaixo do piso da faixa do próprio Rafael (p. 10) sustenta o valor adotado. Registrar isso como argumento C1 em cada um.

### M1 (média). A comparação com o piso de R$ 380.000 não vale para os itens 05–09 e é incerta para os itens 04–10.
- **Evidência.**
  - A faixa da p. 10 é para "lotes de terreno padrão (geralmente medindo 12m x 30m ou 360 m²) localizados nas quadras 37 e 38". Os itens 05–09 têm 307,80 m².
  - Na faixa por m² que Rafael dá na p. 11 (R$ 1.050 a 1.500/m²), esses lotes saem a R$ 360.000 / 307,80 = R$ 1.169,59/m². Ficam **dentro** da faixa, não abaixo.
  - Pela régua por m², só ficam abaixo do piso os itens 03, 04 e 10 (R$ 1.000/m²) e 18 e 19 (R$ 972/m²).
  - Os itens 04–10 têm inscrição da quadra 0071 (pergunta 5 em aberto). Se a quadra registral não for 37, a faixa "quadras 37 e 38" nem se aplica.
  - A p. 11 diz ainda que "lotes de meio de quadra ficam mais próximos do piso".
- **Correção.**
  - Dizer na ressalva que a faixa é para lotes de 360 m², e que por m² os itens 05–09 estão dentro dela.
  - Condicionar 04–10 à pergunta 5.
  - Quantificar o que se deixa de argumentar. Para os itens que ficam e estão abaixo por m² (03, 04, 10), o piso por m² daria 3 × R$ 18.000 = R$ 54.000, e o piso por lote daria até 8 × R$ 20.000 = R$ 160.000. Isso dá no máximo cerca de R$ 26.700 em 1/6.
  - Assim o Luís sabe que aprovar o R1 fecha essa porta, salvo se ele quiser tratar o caso como extremo (P5).

### M2 (média). O R2 remete a titularidade do item 17 às perguntas P7 e P10, mas nenhuma das duas trata do item 17.
- **Evidência.** O R2 diz: "Titularidade e origem seguem como perguntas (P7, P10)". Em `perguntas-ao-grupo.md`:
  - a P7 cobre as inscrições 360.0019.004.000 e 360.0057.001/002.000 e as matrículas 77.833 e 74.857;
  - a P10 cobre os lotes 15 e 17 da **quadra 37** (itens 21 e 22).

  O item 17 (inscrição 360.0038.015.000, o lote "IRINEU SIVIERO" de 2018, item 21 de 2018, sem matrícula) não está em nenhuma pergunta. Ele só aparece como candidato 3 de `parte6-candidatos.md`.
- **Efeito.** O item com a dúvida de titularidade mais forte do pacote (R$ 390.000, lote com nome de sócio em 2018) entra no patrimônio da empresa sem pergunta ao grupo que o acompanhe.
- **Correção.** Criar uma pergunta própria ("O lote 15/q38, inscrição 360.0038.015.000, anotado 'IRINEU SIVIERO' em 2018, é da empresa desde quando e por qual título? Matrícula 33.003 em nome de quem?") e citá-la no R2. Citar também:
  - o candidato 3 da Parte 6;
  - a afirmação do grupo de que "todos os imóveis [estão] registrados em nome da empresa" (arq. 1, p. 11);
  - o doc 3 de 2018: a lista de IPTU de Irineu não traz esta inscrição.

### M3 (média). O R3 diz "sem valor de nenhum lado", mas Rafael dá faixa de valor para as duas residências da Rua da Divisa e pede que entrem na negociação.
- **Evidência.**
  - Na p. 24 Rafael estima "entre R$ 2.400.000 e R$ 3.500.000" para as duas casas.
  - Na p. 25–26 ele diz que esses bens "devem entrar no escopo negocial ou ter a sua destinação … comprovada documentalmente".
  - A mestre traz essa faixa na coluna "Faixa de mercado de Rafael". O pacote não a mostra.
  - Os itens 22 e 23 de 2018 (casas de Raul e Élcio, inscrições 360.0057.002 e 360.0057.001, lotes em sequência, um de "ESQUINA") batem com a descrição de Rafael ("dois lotes … em sequência, na esquina"). A mestre já anota isso como possível correspondência, mas o pacote não.
- **Efeito.** É a maior soma em jogo no R3: de R$ 400 mil a R$ 583 mil em 1/6. O Luís pode aprovar o R3 achando que são linhas sem valor.
- **Correção.** Na linha do R3, mostrar a faixa de Rafael (p. 24) e a possível correspondência com os itens 22 e 23 de 2018. Deixar explícito que aprovar o R3 não descarta o pedido de Rafael.

### M4 (média). A ressalva "Erro de soma de Rafael" afirma mais do que se sabe.
- **Evidência.**
  - O pacote diz: "Os R$ 360.000,00 a mais no total dele não pertencem a nenhum destes itens".
  - A conta confere: a soma dos itens de Rafael é 66.000.000 + 6.710.000 + 2.400.000 + 760.000 = R$ 75.870.000, contra o total declarado de R$ 76.230.000 (p. 20). A diferença é R$ 360.000.
  - Mas a própria aba Totais da mestre chama isso de "Diferença de soma de Rafael (não identificada)" e anota que ela é "igual ao valor de cada um dos itens 03 … 10".
  - Uma hipótese plausível é Rafael ter contado em dobro um lote do R1. O pacote a descarta sem base.
- **Correção.** Reescrever assim: "Diferença não identificada de R$ 360.000,00 no total de Rafael; coincide com o valor de cada lote 03–10 (R1). Não muda o valor de nenhum item deste pacote; afeta o 1/6 e o crédito declarados por Rafael (p. 21), tratados no pacote de exceções."

### M5 (média). A concordância de Rafael é condicional, e o R1 a apresenta como plena.
- **Evidência.**
  - Na p. 12 a concordância vem "sem prejuízo da necessária conferência documental e jurídica quanto à titularidade, existência, regularidade registral".
  - Na p. 12–13 ela fica "sujeita à validação pelo escritório de advocacia".
  - Na p. 24 a seleção de bens fica "condicionada à verificação da titularidade".
  - O grupo afirma a titularidade, mas recusa mandar as matrículas (arq. 1, p. 11).
- **Efeito.** Aprovar o R1 pode ser lido como confirmação de existência e titularidade. Isso pesa nos bens que serão transferidos para nós (02, 11, 12, 18, 19), em especial os itens 18 e 19, com matrícula única 73.482 (pergunta 4).
- **Correção.** Na coluna "Proposta" do R1 e do R2, dizer que se aprova **o valor**, não a existência registral nem a titularidade. Citar a condição de Rafael (p. 12–13 e 24) e destacar que, para os 5 bens da lista, as matrículas são pré-requisito da transferência.

### M6 (média). Data-base dos itens 21 e 22: falta perguntar quando foram adquiridos.
- **Evidência.**
  - Os itens 21 e 22 não estão no inventário 2018.
  - A matrícula 86.669 é maior que a 77.832, que já existia em 2018, o que indica registro posterior.
  - A 124.906 é de Osório.
  - A P10 pergunta "Como foram adquiridos", não quando. A mestre diz "data e forma de aquisição a documentar", mas o pacote não traz isso.
  - Se algum foi adquirido depois de 31/12/2024, ele sai do inventário 06 e vai para o `registro-data-base.md` (P2).
- **Correção.** Acrescentar à P10 a pergunta "em que data (escritura/registro)?". No R2, anotar que a inclusão de 21 e 22 depende de a aquisição ser até 31/12/2024.

### B1 (baixa). A numeração das perguntas colide com a das premissas do plano.
- No mesmo bloco aparecem "sem reavaliação própria (P5)", que é premissa do plano, e "(pergunta P5)", que é a pergunta da quadra 0071.
- Também colidem "Parte 4 (P4 do plano)" e "(pergunta P4)", e no R2 "P10" é pergunta, enquanto P10 no plano é o C0.
- Usar outra sigla para as perguntas ao grupo (ex.: "G5").

### B2 (baixa). O R2 cita só a p. 20 para os itens 21 e 22.
- Rafael também lista os dois a R$ 380.000 cada na seção "avaliação executada" (p. 15, bloco "21 e 22").
- A concordância da p. 16 vem logo depois desse bloco, mas diz "itens anteriores de 02., à 19.". Vale anotar essa ambiguidade, que é de onde vem o "sem declaração expressa".
- Citar "p. 15 e 20".

### B3 (baixa). A amostra não pega um item de R2 com ressalva real.
- Os 5 valores conferem (ver abaixo), mas a única linha de R2 na amostra é o item 21, o de menor risco.
- Trocar ou acrescentar o item 17 (lote Irineu) ou o 22 (Osório).

### B4 (baixa). O script classifica por texto livre.
- `pacote_01.py` separa R1 e R2 procurando "IRINEU", "VÁLTER" e "concordância expressa" na coluna de notas. Uma mudança de redação na nota move itens entre regras sem nenhum aviso.
- O valor de R$ 360.000,00 do erro de soma e as páginas (p. 10, p. 12–13, 16, 20) estão escritos fixos no texto.
- Usar colunas estruturadas (por exemplo "Concordância": expressa/por valor; "Nome de sócio em 2018": sim/não) e ler o 360.000 da aba Totais.

### B5 (baixa). Rodar o script regrava a planilha-mestre.
- `pacote_01.py` roda `build_mestre_06.py` antes de ler, o que regrava a mestre já aprovada pelo Luís.
- Para auditoria, convém um modo só leitura, ou registrar o md5 da mestre usada no pacote.

### B6 (baixa). O cabeçalho do balanço executado de Rafael tem um erro de cópia.
- A p. 16 diz "GRUPO SIVIERO — SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA.", mas o conteúdo é do Hotel Araçá.
- Como o R1 se apoia nos valores dessa seção, vale uma nota de rodapé para que ninguém conteste a identidade da seção.

### B7 (baixa). Os itens 01 e 20 têm um segundo valor de Rafael.
- Na seção "avaliação patrimonial executada", Rafael repete o valor do grupo: R$ 49.606.823,20 (p. 6) e R$ 1.000.000,00 (p. 13). A mestre só cita o balanço executado (p. 17 e 19).
- Isso fica fora deste pacote, mas deve constar no pacote de exceções (lição "o valor pode estar em qualquer seção").

### B8 (baixa). Os itens 03, 21 e 22 ficam no próprio quarteirão do hotel.
- São os lotes 16, 15 e 17 da quadra 37. O item 01 cobre os lotes 01–12 e 18–22, mais a nesga.
- Junto com os lotes 13 e 14 (pergunta 1), eles formam o vão do quarteirão que o inventário diz ter de ser unificado.
- Vale uma ressalva: se parte do hotel (piscinas, estacionamento) estiver sobre eles, o valor como lote avulso pode conflitar com o item 01.

### B9 (baixa). Faltam pistas no R3 para a Parte 4.
- O lote "VÁLTER SIVIERO" (360.0019.004.000, lote 04/q19) é vizinho do item 02 (lote 03/q19, escolhido por Rafael) e aparece na lista de IPTU de Válter de 2018 (doc 6).
- A matrícula 77.833 (lote "ADÉLIA", 2018) é sequencial à 77.832 (item 02).
- São pistas para a Parte 4 que valem uma linha no R3, porque tocam o próprio lado do Luís.

---

## O que conferi e está certo

- **Valores do R1 e do R2 nas duas fontes.** Os 20 itens têm o mesmo valor no inventário 2025 e no balanço executado de Rafael:
  - 02: R$ 400.000;
  - 03–10: R$ 360.000;
  - 11–17: R$ 390.000;
  - 18 e 19: R$ 350.000;
  - 21 e 22: R$ 380.000.
- **Páginas da mestre.** As páginas estão certas:
  - p. 17: itens 01–05;
  - p. 18: itens 06–13 (o valor do 13 está na p. 18);
  - p. 19: itens 14–20;
  - p. 20: itens 21 e 22.
- **Outras seções do PDF.**
  - Seção "avaliação executada" (p. 9–15): os itens 02, 11, 12, 18, 19, 21 e 22 têm o mesmo valor.
  - Quadro de escolhidos (p. 21–22): 02, 11, 12, 18 e 19 com o mesmo valor, somando R$ 1.880.000.
  - Fora isso, não há valor específico de Rafael para nenhum item do R1 ou do R2, só faixas genéricas (p. 10–11).
- **Contagens e somas.**
  - R1: 17 itens, R$ 6.320.000,00.
  - R2: 3 itens, R$ 1.150.000,00.
  - Divergentes: 01 e 20, R$ 50.606.823,20.
  - Total: R$ 58.076.823,20, igual ao TOTAL GERAL do inventário. Os componentes do item 01 somam R$ 49.606.823,20.
  - R3: 9 linhas (5 só 2018, 2 só Rafael, 2 não abordados). Mestre: 31 linhas = 22 + 9.
  - Os itens fora das regras são exatamente os dois que divergem.
- **Pares 2018 × 2025.**
  - 02–16, 18 e 19 batem por inscrição, matrícula e código.
  - 17 ↔ item 21 de 2018, por inscrição e código (2018 sem matrícula).
  - 20 ↔ item 19 de 2018.
  - 21 e 22 não têm par.
  - Os 25 itens de 2018 estão todos cobertos: os itens 20, 22, 23, 24 e 25 estão no R3.
- **Lista de Rafael.** É 02, 11, 12, 18 e 19, com inscrição, matrícula e valor iguais aos da mestre.
- **Amostra.** As 5 linhas (02, 05, 13, 18, 21) conferem nas três fontes: descrição, chave, item e valor de 2018, valores de 2025, página e regra.
- **Ressalvas certas.**
  - "C0 não se aplica".
  - Quadra 0071 nos itens 04–10.
  - Matrícula 73.482 nos itens 18 e 19.
  - Comarca de Osório no item 22.
  - A conta de R$ 360.000.
- **Referências do R3.** As perguntas P1, P7, P8 e P9 correspondem às linhas do R3.
- **Concordância expressa.** A concordância das p. 12–13 e 16 é de fato para os itens 02 a 19, e não cobre 21 e 22.
