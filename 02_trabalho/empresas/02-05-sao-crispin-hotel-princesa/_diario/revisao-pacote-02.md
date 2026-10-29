# Revisão independente — Pacote 2 (Hotel Princesa dos Vales, 05)

28/10/2026. Revisor com contexto novo. Objeto: `pacotes/pacote-02.md` (gerado por `_scripts/pacote_02.py`).

**Como conferi.** Rodei `_scripts/pacote_02.py` de novo. A saída é idêntica ao pacote em revisão (diff vazio) e o total impresso é 79.060.957,5. Depois conferi direto nas fontes:
- doc 05 e doc 02 de 2025 (tabelas e observações);
- doc 00;
- PDF do Hotel de Rafael, p. 1–24, texto completo;
- PDF da Agropecuária de Rafael;
- PDF Diretrizes e PDF INTRODUÇÃO;
- busca de "17.555" em todos os PDFs de Rafael;
- `ativos-avaliados-valter.xlsx` e `inventário-estudo-rafael.xlsx`;
- arq. 1, p. 5–9;
- matrícula 1.412 (arq. 11), lida como imagem fora do projeto;
- planilha 2_0 (GERAL e Resumo);
- logs de decisões da SEI, das empresas 01-08 e da 06.

Não alterei nenhum outro arquivo. Rodar o script regravou o pacote, que ficou idêntico, e a planilha-mestre.

---

## Achados

### A1 — ALTA — Item 01: a faixa de mercado de Rafael (C1) não sustenta os R$ 17.555.000. Ela contradiz esse valor.

**O que o pacote diz:** as faixas de Rafael "somam de R$ 11.900.000,00 a R$ 17.800.000,00 [...] O valor de Rafael está dentro dessa soma; o do grupo também."

**O erro:** a soma junta duas coisas que não são só do item 01:
- a faixa do terreno, de R$ 2,7 a 4,5 milhões, é para os **1.800 m²** da matrícula inteira (p. 8). É R$ 1.500 a 2.500 por m² × 1.800;
- a faixa do prédio, de R$ 9,2 a 13,3 milhões, é para os 2.040 m² construídos (p. 9).

O item 01 avalia só **720 m²** de terreno. Os outros 1.080 m² são o item 02, como o próprio grupo esclarece (arq. 1, p. 8–9) e a matrícula confirma: lote único nº 313, 1.800 m², 30 m × 60 m.

Feita a conta certa:

| Comparação | Faixa de Rafael | Grupo | Rafael | Resultado |
|---|---|---|---|---|
| Só o item 01 (720 m² + prédio) | 720 × R$ 1.500 a 2.500 = R$ 1,08 a 1,8 milhão; mais o prédio: cerca de **R$ 10,3 a 15,1 milhões** | R$ 13.504.243,50 | R$ 17.555.000,00 | Grupo dentro da faixa. Rafael **cerca de R$ 2,45 milhões acima do teto** |
| Itens 01 + 02 juntos (os 1.800 m²) | R$ 11,9 a 17,8 milhões | R$ 21.604.243,50 | R$ 25.655.000,00 | Os dois bem acima do teto |

A faixa do prédio é o "valor estimado de mercado para o Hotel Princesa dos Vales" (p. 9). Rafael diz que ela combina estrutura e "potencial de geração de receita". Pode, portanto, já incluir a fração do terreno, e somar a faixa do terreno seria contar o terreno duas vezes. Nesse caso o teto do item 01 cai para cerca de R$ 13,3 milhões.

**Por que importa:** nos precedentes aprovados, o C1 confirmou o C0:
- SEI, pacote 3: "C1 confirma em todas";
- Hotel Araçá, E1: "grupo abaixo do piso da faixa".

Aqui o C1 vai **contra** o valor que o C0 escolhe. A faixa de Rafael também pesa contra o item 02, já aprovado em R$ 8.100.000 (R$ 7.500/m² contra R$ 1.500–2.500/m²). O pacote cita isso como risco (3), mas não tira a consequência.

**Correção sugerida:**
- Trocar o parágrafo do C1 por: "O C0 indica o valor de Rafael. O C1 aponta o contrário: a faixa de mercado do próprio Rafael para o item 01 vai até cerca de R$ 15,1 milhões. O valor do grupo está dentro dela; o de Rafael está acima."
- Apresentar ao Luís as duas opções como escolha real: C0, R$ 17.555.000; ou C1, R$ 13.504.243,50.
- Mostrar o efeito de cada uma no 1/6.

### A2 — ALTA — Redução de R$ 400.000: incompatível com o tratamento já aprovado para o mesmo tipo de redução nas empresas 01 e 08 (P5)

**O que o pacote diz:** "É o próprio número do grupo (não é reavaliação)".

**Evidência contrária:**
1. **Empresas 01 e 08.** O Luís aprovou, no pacote 3, a regra C1 (`empresas/01-08-abastecedora-transportadora/log-decisoes.md`): "reduções de 0,70 % e 1,5 % [...] ônus de dissolução [...] Nada é descontado do valor da 01 e da 08 (P5 + diretriz)". A justificativa no pacote 3 foi: "P5: nenhum dos lados aplica". O grupo usa o mesmo texto ("O valor real acima apontado não é líquido [...] sofrerá redução pela regularização dos imóveis") em todos os inventários (planilha 2_0 GERAL B9, B18, B26; doc 05).
2. **Hotel Araçá.** A sessão tratou a regularização do hotel como pergunta (G14, "risco de desconto em dobro"), não como dedução.
3. **SEI.** Ninguém decidiu abater o rol de custas (obs. 9, R$ 250.000). A análise de cobranças trata esses textos como "cobrança anunciada" a verificar.
4. **Aqui também nenhum dos lados aplica a redução.** O TOTAL GERAL do doc 05 é R$ 75.410.201,00 (T6R3). A tabela "VALOR REAL DE CADA SÓCIO" do próprio doc 05 calcula 16,66 % sobre esse total bruto (T7). O doc 00 usa os mesmos R$ 75.410.201,00 como PL do Hotel (T2R6). O total de Rafael, R$ 79.491.957,50, também é bruto (p. 19–20).

Aplicar a redução no Hotel, onde ela favorece o Luís (a fazenda vem para ele), e não aplicá-la na 01 e na 08, onde ela o prejudicaria, é uma assimetria que o outro lado vai apontar. Pelo critério das empresas 01 e 08, abater é um ajuste fora das duas avaliações. Exige, no mínimo, a justificativa escrita de "caso extremo" (P5).

**Correção sugerida:**
- Dizer isso ao Luís com todas as letras.
- Oferecer duas opções:
  - (a) não abater: coerente com 01, 08 e P5; fato registrado sem valor;
  - (b) abater: como "caso extremo" com justificativa escrita, por ser o único caso em que o grupo dá um valor em reais, e não um percentual ou um texto em branco.
- Mostrar o efeito de cada opção no 1/6.

### A3 — ALTA — "Quem recebe paga a regularização" contraria a fonte, e o efeito para o Luís depende de como a redução é lançada

**O que o pacote diz:** "A fazenda vem para nós: quem recebe paga a regularização, então o valor líquido para nós é menor (C0)."

**A fonte diz o contrário.** Doc 05 e doc 02, OBSERVAÇÕES f): "Perfazem todas as despesas, um TOTAL de R$ 400.000,00 [...] que deverão ser satisfeitas pela atual credora e proprietária das áreas de terras HOTEL PRINCESA DOS VALES LTDA." Pelo texto, quem paga é a empresa.

O custo também é de uma operação específica: passar as escrituras do Hotel e de terceiros para a **Agropecuária** e unificar as matrículas. O grupo diz que "ainda não se deu início a qualquer providência" (arq. 1, p. 6, item 2). Na data-base, portanto, é uma estimativa de custo futuro, sem nada desembolsado (P2). Se a fazenda vier para o Luís, essa transferência para a Agropecuária deixa de existir e é substituída por outra, com ITBI próprio.

**Efeito, pelas duas leituras possíveis:**
- **Redução do valor da fazenda**, como propõe o pacote: o Luís recebe a fazenda R$ 400.000 mais barata e o total cai R$ 400.000, o que reduz a cota dele em R$ 66.666,67. Ganho líquido para o Luís: **R$ 333.333,33**.
- **Passivo da empresa**, como diz o texto f): a fazenda mantém o valor e o total cai R$ 400.000. Efeito para o Luís: **menos R$ 66.666,67**.

Há também risco de benefício duplo: o valor da fazenda é reduzido e, depois, a empresa paga o ITBI e as escrituras da transferência ao Luís. O grupo vai contestar.

**Correção sugerida:**
- Retirar o argumento "quem recebe paga" ou trocá-lo pelo texto da fonte.
- Mostrar as duas leituras com o efeito no 1/6.
- Condicionar qualquer abatimento a que o instrumento da cisão defina quem paga a regularização e o ITBI da transferência ao Luís.

### M1 — MÉDIA — Item 05: "não registra concordância" é incompleto

Logo após o item 05 não há bloco de concordância (p. 13–14 trazem a avaliação técnica). Mas Rafael remete ao item 05 em duas concordâncias seguintes:
- item 06, p. 14: "DA CONCORDÂNCIA [...] de acordo como citado nos itens anteriores (02., 03., 04.,05.)";
- item 07, p. 15: "(02., 03., 04., 05., 06.)".

Além disso, o quadro-resumo de Rafael usa R$ 44.828.300 no conjunto escolhido de R$ 50.914.157,50 (p. 20 e 23). A planilha-mestre já registra isso ("citado na concordância de outros itens (p. 14, p. 15)"). O outro lado pode dizer que Rafael concordou com o item 05 por remissão, o que enfraquece o argumento dos "parâmetros preliminares".

**Correção:** trocar por "não há concordância expressa logo após o item 05, mas as concordâncias dos itens 06 e 07 remetem a ele (p. 14 e 15)".

### M2 — MÉDIA — Item 05: "provisório" não é uma das saídas da P5, e faltam as consequências

A P5 e a coluna "Decisão (P5)" da planilha-mestre preveem três saídas: "Adoto grupo / Adoto Rafael / Caso extremo". "Provisório" é uma categoria nova e precisa de decisão explícita do Luís. O pacote não responde:
- **E se a avaliação independente nunca vier?** Então o "provisório" é definitivo (R$ 44.828.300). Convém dizer isso.
- **Quem encomenda e quem paga, e em que prazo?** O grupo já disse que prefere não ter despesa nesta fase (arq. 1, p. 6, item 1).
- **E se o resultado vier maior?** Rafael pede que a avaliação considere também a "exploração mercantil por arrendamento" e a "capacidade produtiva" (p. 13–14), o que pode **subir** o valor. O laudo seria um terceiro valor, e pela P5 só entra como caso extremo. Usá-lo só quando for menor (C0) é uma assimetria visível. O Luís deve decidir antes se aceita o laudo em qualquer sentido.
- **O efeito no total.** O total do Hotel fica sem fechar até o laudo.

**Correção:**
- Registrar a decisão como "Adoto o valor do grupo (igual ao de Rafael)", com a ressalva do pedido de avaliação.
- Explicar ao Luís as quatro consequências acima.

### M3 — MÉDIA — Item 05: o argumento técnico usa números fracos, rotulados como C1, e omite o mais forte

1. **Os valores do ITBI não são C1.** O C1 é a faixa de mercado de Rafael (legenda do pacote; log da SEI). Os R$ 40.000 / 18.000–20.000 / 8.000 por hectare são valores fiscais mínimos informados pelo grupo. O próprio grupo diz que "são estimativas obtidas do conhecimento na prática" (arq. 1, p. 6, item 4).
2. **Campo nativo.** A fonte diz "R$ 18.000,00 [...] a R$ 20.000,00", e o pacote só cita R$ 18.000.
3. **Dr. Pablo.** "O Dr. Pablo apontou o mesmo (arq. 1, p. 5)" é impreciso. Ele não cita o ITBI. Ele aponta que áreas produtivas e não produtivas "passaram a ter o mesmo valor por hectare" e pede: "Esclarecer o motivo para o aumento do preço para R$110.000,00/hectare, uma vez que ainda estamos recebendo o pagamento de áreas vendidas por R$ 49.726,30/hectare."
4. **Argumentos omitidos:**
   - **Preço de venda real, o mais forte:** a venda aos Rigo em 2018, R$ 9.448.000 por 190 ha, dá R$ 49.726,32/ha. É uma transação efetiva, não um valor fiscal.
   - **O próprio inventário do grupo de 2018** diferenciava: R$ 44.000/ha para área produtiva e R$ 22.000/ha para não produtiva (planilha 2_0, Resumo C81–C85).
   - **A versão antiga do inventário de 2025** usava R$ 130.000/ha, R$ 52.978.900 no item 05 (planilha-mestre, coluna da versão antiga). O valor por hectare do grupo já mudou.

**Correção:** separar "C1 de Rafael", que aqui é só o texto das p. 13–14, sem faixa, de "outros indícios". Citar o Dr. Pablo pelo que ele escreveu. Incluir os R$ 49.726,32/ha da venda de 2018 e a diferenciação de 2018.

### M4 — MÉDIA — Item 01: "sem memória de cálculo" está certo, mas faltam dois fatos para o risco de "erro de transcrição"

1. **O valor de Rafael passa pelo total dele.** O total de Rafael, R$ 79.491.957,50 (p. 19 e 20), e a cota de 1/6, R$ 13.248.659,58 (p. 20), só fecham com R$ 17.555.000. O valor aparece numa página só (p. 16), mas passa pelo total e pela apuração da cota dele. Isso enfraquece a tese de erro de transcrição.
2. **Possível origem, não documentada.** R$ 13.504.243,50 × 1,30 = R$ 17.555.516,55. A razão entre os dois valores é 1,29996, e R$ 17.555.000 é esse resultado arredondado para baixo, no milhar. Pode ser um acréscimo de 30 % sobre o valor do grupo. Também está perto do valor de 2018 (R$ 17.475.786,03, planilha-mestre). É hipótese, sem documento.

Confirmei que não há cálculo de Rafael para esse valor em nenhum dos seus documentos:
- PDF do Hotel, p. 1–24, inclusive o quadro-resumo das p. 19–20 e a seleção das p. 21–23;
- Diretrizes: o item 01 não está na lista do Hotel, p. 4–6;
- INTRODUÇÃO: 16 páginas, sem menção ao Hotel nem ao valor;
- os dois xlsx da lista;
- busca de "17.555" em todos os PDFs de Rafael: só aparece na p. 16 do Hotel.

**Correção:** acrescentar os dois pontos ao risco (1).

### M5 — MÉDIA — Totais: comparação entre valores brutos e líquidos, sem o efeito no 1/6

- **A conta confere:** pacote 1, R$ 16.583.657,50; mais os itens 05, 06, 01 e 08 (R$ 44.828.300 + 440.000 + 17.555.000 + 54.000), igual a R$ 79.460.957,50; menos R$ 400.000, igual a **R$ 79.060.957,50**.
- **A comparação mistura bruto e líquido.** O grupo (R$ 75.410.201,00) e Rafael (R$ 79.491.957,50) são totais **brutos**, e o recomendado é **líquido**. Mostrar também o bruto recomendado (R$ 79.460.957,50).
- **O resumo atribui a dedução ao grupo.** A linha da redução mostra "− R$ 400.000,00" na coluna **Grupo**, como se o grupo a abatesse. Ele não abate (A2). Pôr "anunciada, não abatida".
- **Falta o efeito no 1/6 de cada decisão:**
  - item 01: + R$ 675.126,08 na cota do Luís com o valor de Rafael;
  - R$ 400.000: + R$ 333.333,33 ou − R$ 66.666,67, conforme a leitura (A3);
  - item 08: + R$ 25.833,33 a favor do Luís com R$ 54.000 em vez de R$ 85.000.

### B1 — BAIXA — Precedente Oleoplan: justo no essencial, mas incompleto

O padrão é o mesmo: as parcelas impressas por Rafael somam o valor do grupo, e o valor final dele é outro. O pacote, porém, não diz três coisas:
- **Direção oposta:** o Oleoplan estava na lista de Rafael, e o valor adotado foi o **menor**.
- **Mais suporte:** o valor de Rafael para o Oleoplan também está no PDF Diretrizes (p. 9), em dois documentos. Os R$ 17.555.000 estão em uma só página.
- **Situação diferente no C1:** no Hotel o C1 contradiz o valor de Rafael (A1).

**Correção:** acrescentar essas três diferenças.

A citação do log na legenda está cortada no meio de uma tabela markdown (pipes e "…"). Trocar por uma frase.

### B2 — BAIXA — Item 01: data-base, matrícula e notas do Luís não aparecem

- **Data-base.** As faixas de Rafael são de 2026 (RDB-4). O pacote não menciona.
- **Matrícula.** A matrícula 1.412 (certidão de 15/05/1996, Av. 4/1.412, de 1981) averba um prédio de 1.973,28 m². O grupo avalia 2.040 m² mais 510 m² de subsolo e afirma que "não há pendências de averbações" (arq. 1, p. 9). A certidão é antiga demais para concluir. Vale uma pergunta ao grupo (certidão atualizada).
- **Notas do Luís** na planilha 2_0 sobre os componentes do item 01, não citadas:
  - H30, H34 e H35: CUB de 2022 e de 2017;
  - H38: "Que obra é essa?".
- **Respostas do grupo** sobre o CUB e a reforma: arq. 1, p. 9, item 4.

### B3 — BAIXA — Ressalvas de área da fazenda

- **Origem dos 411,53 ha.** "407,53 + 4,0 = 411,53 ha pelas matrículas" é impreciso: as matrículas listadas somam 491,0 ha. Os 411,53 ha saem da área avaliada nos itens 05 e 06.
- **Matrícula 00.760 duplicada.** O grupo já explicou a duplicidade: "áreas de 48,0ha e 100,00ha foram adquiridas em momentos distintos [...] remanesce apenas a área de 64,53ha" (arq. 1, p. 9). O pacote deve citar a explicação, ainda a confirmar (Q3).
- **Contas conferidas:** a diferença de 0,1 ha (cerca de R$ 11.000) e o ITR de 259 ha estão certos.

### B4 — BAIXA — Redução: ressalvas a completar

- **"Hipótese fraca."** Deixar claro que o fraco é a hipótese de duplicidade com os R$ 250.000 da SEI, não a redução.
- **Contar uma vez só.** O mesmo texto está nos docs 02 e 05 e nos dois PDFs de Rafael: Hotel, p. 4–6; Agropecuária, p. 1–3. A planilha-mestre conta uma vez, o que está certo, e o pacote deve dizer isso.
- **Ônus de dissolução.** O texto reserva "os ônus de eventual dissolução da empresa em relação a parte de algum(ns) sócio(s)". São ônus não quantificados, dirigidos ao sócio que sai. Nas empresas 01 e 08 isso ficou como aviso para a negociação. Repetir o aviso aqui.
- **Sede Campestre (item 04), também na lista de Rafael.** As benfeitorias "pendem de averbação" (arq. 1, p. 9), sem custo estimado. O valor está fora dos R$ 400.000, que são calculados só sobre os 411,4 ha da fazenda, e o mesmo vale para o item 07. A dedução pertence só aos itens 05 e 06.

### B5 — BAIXA — Clareza para o leitor

- **Recomendação ambígua.** "Recomendação: adotar R$ 17.555.000,00 [...] Alternativa mais defensável: R$ 13.504.243,50" recomenda uma opção e chama a outra de mais defensável. Depois da A1, apresentar as duas opções lado a lado, com prós e contras e efeito no 1/6.
- **Item 08:**
  - as duas referências FIPE são posteriores à data-base (RDB-2);
  - o FIPE do grupo é de um "veículo similar MB 709 2p";
  - o de Rafael é "somente o chassi". Os R$ 54.000 dão valor zero à carroceria boiadeiro. Isso favorece o C0, e o grupo não pode contestar o próprio número. Basta dizer em uma linha.
- **Siglas.** Pôr no início uma frase simples do que é C0, C1 e P5, além da legenda no fim.

---

## Conferido e correto

- **Valores do grupo**, doc 05, T4R2, T5R3, T5R4 e T6R2: item 01, R$ 13.504.243,50; item 05, R$ 44.828.300,00; item 06, R$ 440.000,00; item 08, R$ 54.000,00. Total do grupo, R$ 75.410.201,00 (T6R3), sem abater os R$ 400.000.
- **Valores de Rafael:**
  - R$ 17.555.000,00 no balanço executado, p. 16. Os componentes impressos nas p. 16–17 são os do grupo. Na avaliação, p. 7, ele repete R$ 13.504.243,50;
  - item 08: R$ 85.000,00 na p. 19, impresso duas vezes e contado uma no total; R$ 54.000 na p. 15;
  - itens 05 e 06 iguais aos do grupo, p. 12, 14, 18–19 e 23;
  - total R$ 79.491.957,50 (p. 19–20), que confere com a soma.
- **Citações:** todas literais e nas páginas indicadas. São as das p. 14 ("parâmetros preliminares…"), 13 ("valores inferiores…"), 8 (faixa do terreno e "entre R$ 1.500 e R$ 2.500 por m²."), 9 (faixa do prédio) e 15 (caminhão, FIPE de R$ 54.745,00 e "unidades reformadas"). A recomendação de avaliação independente está nas p. 13–14. O texto dos R$ 400.000 é idêntico nos docs 05 e 02.
- **Lista de Rafael:** itens 04, 05, 06 e 08 no PDF do Hotel (p. 20 e 22–23), no Diretrizes (p. 4–6) e nos dois xlsx. O item 01 não está na lista. A direção do C0 está certa nos dois itens com divergência:
  - item 01, fica na empresa: maior valor, o de Rafael;
  - item 08, vem para o Luís: menor valor, o do grupo.
- **FIPE:** R$ 54.745,00 (Rafael, p. 15) e R$ 56.854,00 (grupo, FIPE de out/2025, arq. 1, p. 9).
- **Contas da regularização:** 100.000 + 246.840 + 5.000 + 5.000 = R$ 356.840. 411,4 ha × R$ 30.000 = R$ 12.342.000, e 2 % disso dá R$ 246.840.
- **Áreas da fazenda:**
  - 100 + 186 + 45,6 + 11,4 + 64,53 = 407,53 ha;
  - matrículas = 491,0 ha;
  - R$ 110.000/ha para toda a área;
  - 411,43 ha declarados no doc 17 de 2018;
  - ITR de 259 ha nas contas pagas pela SEI (09/2024).
- **Risco de dupla contagem no item 01 (risco 2):** está bem formulado como condicional. A matrícula tem um lote só, de 1.800 m². O grupo divide 720 m² no item 01 e 1.080 m² no item 02 (arq. 1, p. 8–9), sem dupla contagem do lado do grupo. A faixa do terreno de Rafael é para os 1.800 m².
- **Precedente Oleoplan:** a citação do log da SEI está correta, e o valor foi mantido depois da revisão ("Mantido o valor de Rafael, R$ 10.631.378,60").
- **Total:** R$ 79.060.957,50 = valores da R1 (R$ 16.583.657,50) + valores recomendados (R$ 62.877.300,00) − R$ 400.000,00.
- **Script:** reexecutado, com saída idêntica. Os asserts de C0 dos itens 01 e 08 e de cobertura de todos os itens passam.
