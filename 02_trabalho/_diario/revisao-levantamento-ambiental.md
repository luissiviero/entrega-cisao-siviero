# Revisão independente: levantamento ambiental da SEI

- Data: 2026-10-28
- Objeto: `02_trabalho/levantamento-ambiental-sei.xlsx` (7 abas) e `_scripts/levantamento_ambiental.py`
- Revisor: agente de contexto novo. Li as fontes diretamente: doc 11 atual e antigo (python-docx), inventários de 2018 (docx e xlsx) e de 2015 (docx do Drive), arquivo 1 (PDF, 32 p.), arquivo 17 (xlsx, 2 abas), PDFs e xlsx de Rafael (executados, terceiros, Bonetto, demais compromissos, áreas rurais), ACERTOS E ACORDOS (gdoc), contas pagas (aba Consolidado), decisoes-sei.json e registro-data-base.md.
- Data-base: 31.12.2024.

## Achados

### 1. ALTA: quatro dos 26 "lotes de terceiros" ainda eram da SEI em 31.12.2024, e o levantamento não tira a consequência disso
- **Problema:** o arquivo 17 e o arquivo 1 mostram que quatro lotes só saíram da SEI depois da data-base:
  - Jardim Itália 001/1747, vendido em 10/03/2025;
  - Primavera 010/1199 e 011/1199, dados à SAFRA em fev. e mar./2025;
  - Vicentino 003/1301, vendido em 19/09/2025.

  Em 31.12.2024 eram estoque próprio. Não podiam existir como "lote de terceiro passível de indenização". Somam R$ 1.887.000 (200.000 + 166.000 + 166.000 + 1.355.000), lançados no passivo de R$ 16.240.000.
- **O que falta nas abas:** as sobreposições 3 e 4 tratam esses lotes como "contradição" ou "dado em pagamento e cobrado de novo". Não dizem o principal. O doc 11 descreve a situação de out./2025, e não a de 31.12.2024. Na data-base:
  - os 4 lotes deveriam estar no ativo, ao menos como "possível APP", como os 7 lotes próprios;
  - não havia terceiro a indenizar;
  - no caso SAFRA, havia um passivo próprio na data-base: a ação rescisória de 2023 (multa de 20% mais 5 lotes pedidos). O doc 11 não o traz, porque o acordo de fev./2025 já o tinha "quitado" com esses dois lotes.

  O próprio grupo assume a data de 31/10/2025.
- **Evidência:**
  - arquivo 17, aba "Lotes vendidos (PDF)": linhas 6 (p.1, "Lote dado como parte da indenização pelo distrato", 2025-03-27), 24 (p.3, 2025-09-19) e 162 (p.8, 2025-03-10);
  - arquivo 1, p.26: "Em fevereiro de 2025, foi celebrado acordo ... dação em pagamento de dois lotes urbanizados de nº 10 e 11 da quadra 1199";
  - arquivo 1, p.28, item 12: "termo final adotado para o encerramento da participação – 31 de outubro de 2025".
- **Correção:** criar uma sobreposição ou ajuste próprio: "lote próprio na data-base lançado como passivo de terceiro" (R$ 1.887.000). Lançar no `registro-data-base.md` uma linha para cada um dos 4 lotes e outra para o passivo SAFRA na data-base. Registrar na aba "Situação" que o doc 11 retrata out./2025.

### 2. ALTA: o casamento com o arquivo 17 perde os fatos dos lotes 015, 018, 019 e 020/11B do Lago Negro; a aba "Apos data-base" está incompleta
- **Problema:** o script só casa pelas colunas Lote/Quadra (`arq17_lot`). As permutas do Lago Negro estão registradas no lote *recebido*, e o lote de APP aparece só no texto da coluna Observações. A indenização em dinheiro está só na aba "Totais por loteamento". Por isso:
  - a coluna "Arquivo 17" da aba Lotes está vazia para 015, 018, 019 e 020/11B;
  - "Após a data-base?" está em branco para 018 e 019;
  - a aba "Apos data-base" lista só 3 lotes. Faltam o 018/11B (permuta Candiago, 28/08/2025) e o 019/11B (R$ 580.000 a Zucoloto, 14/10/2025). Falta também o Primavera 010/1199 (dação de fev./2025, citada só no arquivo 1).
- **Evidência:**
  - arquivo 17, "Lotes vendidos (PDF)", linha 190 (p.9): lote 05/05, Adriana Wasen, 2019-04-24, "Permuta, com o lote 20, da quadra 11B – APP";
  - linha 194 (p.9): lote 14/05, Paulo Roberto Candiago, **2025-08-28**, "Permuta, com o lote 18, da quadra 11B";
  - linha 216 (p.9): lote 10/09, Michael Schneider Flack, 2019-08-20, "Permuta, com o lote 15, da quadra 11B";
  - "Totais por loteamento", linha 34: "indenizou, em espécie, o Lote 19 ... R$ 580.000,00, no dia 14 de outubro de 2025".
- **Correção:**
  - no script, procurar também no texto de Observações o padrão "lote NN, da quadra 11B" e atribuir o fato ao lote de APP;
  - incluir à mão o 019/11B (nota de Totais) e o 010/1199 (arquivo 1);
  - reemitir a aba "Apos data-base", que deve ter pelo menos 6 lotes: 001/1747, 010/1199, 011/1199, 003/1301, 018/11B e 019/11B.

  Outros casos de grafia testados não esconderam nada (ver a lista no fim).

### 3. MÉDIA: a posição de Rafael no San Marino vem copiada sem alerta e contradiz o total da sobreposição 9
- **Problema:** a aba Lotes põe, como posição de Rafael para os 3 lotes de possível APP do San Marino, a citação de p.94: "diferença nominal em seu favor no montante de R$ 7.700.000,00". O valor é erro de Rafael. É a mesma frase do Lago Negro, em que 7.700.000 = 7 lotes × R$ 1.100.000. Os três lotes do San Marino somam R$ 730.000 (244.000 + 233.000 + 253.000).
- **Total da sobreposição 9:** diz que Rafael cobra "R$ 2.321.000 no total" pelos 7 lotes de possível APP. Esse número só fecha se o San Marino for 730.000. Com o que Rafael escreveu, a soma é 7.700.000 + 131.000 + 335.000 + 1.125.000 = R$ 9.291.000. A correção foi feita em silêncio.
- **Como Rafael lista os lotes:** a sobreposição 9 diz que ele "lista os 7 lotes de possível APP entre os escolhidos". É impreciso. Os lotes de APP aparecem logo abaixo da lista de escolhidos, depois de "IMPORTANTE!", e **não entram no total escolhido**:
  - San Marino: escolhidos R$ 517.000 = 001/1673 + 003/1676 (p.93);
  - Dalzóchio: R$ 745.000 = O13/0907 + 003/1046 (p.99);
  - Vicentino: R$ 1.355.000 = só o 007/1303 (p.115).

  Rafael os cobra à parte, pelo valor cheio (100%, e não 1/6).
- **Lotes de APP confirmada:** pelo mesmo raciocínio, Rafael dá valor implícito de R$ 1.100.000 a cada um dos 7 lotes do Lago Negro (p.81 e p.84) e R$ 130.000 a cada lote do Simioni (R$ 520.000, p.20-21). Na coluna "Valor de Rafael" esses lotes aparecem como "sem valor". Isso deve ficar explícito.
- **Evidência:** Rafael, executados, p.94 ("R$ 7.700.000,00 mil reais. (Verificar documentação relativa as APPS)") e p.84 (o mesmo texto do Lago Negro); p.93, p.99, p.100, p.106, p.115 e p.116.
- **Correção:**
  - marcar a citação de p.94 como "erro aparente de Rafael (os 3 lotes somam R$ 730.000)";
  - explicar na sobreposição 9 de onde vêm os R$ 2.321.000;
  - trocar "lista entre os escolhidos" por "lista logo abaixo dos escolhidos, fora do total escolhido, e cobra o valor cheio à parte";
  - na coluna "Valor de Rafael" dos 11 lotes de APP confirmada, registrar o valor implícito (1.100.000 e 130.000).

### 4. MÉDIA: Lago Negro. A obrigação já existia em 31.12.2024, e "sem duplicidade" esconde uma omissão; contradições de contagem não apontadas
- **Problema:** a sobreposição 6 conclui "Sem duplicidade em 2025" e diz que os R$ 580.000 e a permuta Candiago "são posteriores à data-base e não estão no passivo (registro de data-base)". Três ressalvas:
  - **A obrigação é anterior à data-base.** A ação civil pública de 2016 pedia "indenizar ou substituir os adquirentes", e a liminar tornou 3 lotes indisponíveis. Em 31.12.2024 estavam pendentes o lote 18 (Candiago) e o lote 19 (Zucoloto). Só o pagamento e a escritura são de 2025. A obs. 6 diz que a empresa "fez a indenização de 02 lotes", ou seja, retrata out./2025. Por isso não há provisão.
  - **O `registro-data-base.md` só tem a linha 5 (Candiago).** Não há linha para os R$ 580.000. A remissão "(registro de data-base)" não sustenta a afirmação.
  - **As fontes contam de modo diferente, e isso não foi apontado.** Obs. 6: indenização de 02 lotes "para a Família Candiago e Terceiro". Arquivo 17: "supriu 04 lotes de sua propriedade e obrigou-se a indenizar 03 lotes vendidos para terceiros pela Família Candiago". Arquivo 1: três adquirentes (Wasen, lote 20; Flach, lote 15; Zucoloto, lote 19) mais o lote 18 do termo de divisão com Paulo Candiago, isto é, 4 casos. O Dr. Pablo, por sua vez, escreve "Lago Negro (-8 lotes APP)".
- **Evidência:** arquivo 1, p.24-25; doc 11 P138; arquivo 17, Totais, linha 34; arquivo 1, p.18 (tabela do Dr. Pablo); `registro-data-base.md`, linhas 1 a 6.
- **Correção:** reescrever a "Situação" da sobreposição 6: "obrigação existente na data-base, sem provisão no doc 11 (lote 19: R$ 580.000; lote 18: um lote, o 14/05)". A decisão já aprovada manda isso para as Partes 4 e 6. Acrescentar a linha dos R$ 580.000 ao registro de data-base e listar a divergência de contagem (2 x 3 x 4 x 8).

### 5. MÉDIA: falta a sobreposição possível Perini (obs. 2, R$ 5.000.000) x Dalzóchio (obs. 3, R$ 2.200.000)
- **Problema:** o levantamento anota que o número 048/1.08.0001924-5 aparece em 2015 como "ACP – Loteamento Perini" e em 2018 como a ACP do Dalzóchio (fonte da água). Não leva isso para a aba Sobreposicoes. Os fatos:
  - o TAC do Perini é de 04/12/2012;
  - o inventário de 2018 não lista nenhum processo do Perini;
  - o doc 11 antigo citava a infração do Perini sem a contabilizar ("Também não, as infrações ambientais ...");
  - o doc 11 atual soma as duas observações.

  Não se pode excluir que o mesmo processo de 2008 tenha dado origem ao TAC de 2012 e esteja nas duas observações (R$ 7.200.000 juntos).
- **Evidência:** inv. 2015, T13.L9 (processo 08, "Loteamento Perini", nº 048/1.08.0001924-5); doc 18, T21.L2 (o mesmo número, "Ivanor Carlos Dalzóchio ... Fonte da Água"); arquivo 1, p.23 ("TAC celebrado em 04.12.2012"); doc 11 antigo, P118.
- **Correção:**
  - criar uma sobreposição candidata: "obs. 2 e obs. 3 podem ser o mesmo processo; pedir certidão de objeto e pé do 048/1.08.0001924-5";
  - registrar na aba Processos, linha Perini, que o caso aparece em 2015, some em 2018 e volta em 2025;
  - acrescentar o início "2008/2012" ao lado de "dez./2024".

### 6. MÉDIA: Lago Negro. A compensação de 3 lotes com o Município (doc 11 antigo) e a recusa do MP em ago./2025 ficaram de fora
- **O que diz a versão antiga:** "poderá ser compensado 03 (três) lotes em área de Uso Institucional do Município, dentro da mesma área do Loteamento Lago Negro" (doc 11 antigo, P129). O parágrafo foi removido na versão atual.
- **O que diz o arquivo 1:** o perito sugeriu essa permuta. O Município aceitou, mas "na audiência de agosto de 2025 ... o Ministério Publico foi contra esse acordo" (p.25).
- **Por que importa:** era uma recuperação possível de 3 lotes, perdida depois da data-base. Também explica por que o texto mudou.
- **Correção:** incluir a informação nas abas Itens (Lago Negro) e "Apos data-base", e na linha da diferença entre as versões do doc 11.

### 7. MÉDIA: Bonetto. Falta a posição de Rafael, que soma dois cenários que se excluem
- **Problema:** o item Bonetto traz só o doc 11 e o arquivo 1. Rafael:
  - avalia a parte da SEI como ativo: 30,5 lotes × R$ 102.000 = R$ 3.111.000;
  - e ao mesmo tempo deduz a multa de R$ 782.000, ficando com R$ 388.167 líquidos.

  Se o loteamento é inviável por APP, como diz o doc 11, a multa existe e os lotes não. Se é viável, há lotes e não há multa. O grupo, por sua vez, só deduz a multa e não dá valor aos lotes.
- **Evidência:** Rafael, LOTEAMENTO BONETTO, p.3 a 6 ("R$ 3.111.000,00, resultante de 30,5 lotes × R$ 102.000,00"; "R$ 518.500,00 – R$ 130.333,33 = R$ 388.166,67"); doc 11, T42.L8.
- **Correção:** acrescentar a posição de Rafael e uma sobreposição "ativo e multa do mesmo empreendimento em cenários excludentes". A mesma lógica vale para Pasqual e Dr. Letti, se Rafael tiver feito o mesmo (não conferi esses PDFs).

### 8. MÉDIA: as contas pagas têm itens ambientais que o levantamento não traz, e a aba "Fontes verificadas" diz que não há
- **Problema:** a aba diz "Só despesas pequenas: vistoria de nascentes do Lago Negro, laudo SAFRA, multa FEPAM do Pedra Lisa, mudas doadas à SEMMA". Faltaram:
  - **Multa ambiental da Prefeitura de Cotiporã:** "P.01/36-MULTA AMB COLONIA SAO BRAS-REMOCAO VE…", NF de R$ 25.243,37 em 36 parcelas de R$ 701,19 (Consolidado, linhas 1777, 2044 ... 3576; P.08/36 em out./2024). Em 31.12.2024 restavam cerca de 26 parcelas, uns R$ 18.200. Não está no doc 11.
  - **ART "PROJETO DE CONSTRUÇÃO LOTES SIMIONI - F…":** bióloga Diana Fiori, 08/12/2023 (linha 1083). Se for o Fracionamento Simioni, contradiz o "ÁREA DE APP" sem valor do grupo e reforça o TAC ("avisar ... restrições ... para a construção de benfeitorias"). Conferir.
  - **Reposição florestal em processos administrativos:** "570 MUDAS PROCESSO ADM LOTEAMENTO BOFF", "585 MUDAS ... LOT PALAVRO", "345 MUDAS ... LOT PICCOLI" (linhas 915-917) e "1200 MUDAS ARAUCARIA PROC ADM LOT BOFF" (linha 1400).
  - **PRAD do Alto das Graças:** linhas 123 e 1921.
  - **"CUSTAS RECURSO REF MULTA R$ 40.000,00-LOT COL…":** linha 602, natureza não identificada.
- **Correção:**
  - incluir a multa de Cotiporã como item ambiental (pequeno, mas é passivo na data-base e está fora do doc 11);
  - registrar a ART do Simioni na aba de documentos do Simioni;
  - corrigir o texto da aba "Fontes verificadas".

### 9. MÉDIA: a data da Lei 4.814/2023, de fonte externa, é usada como premissa de conclusões
- **Problema:** a aba Processos marca o "Marco Regulatório" com nota de fonte externa e "A conferir". Mesmo assim, a sobreposição 4 afirma "Lotes vendidos pela própria SEI depois da Lei 4.814/2023" e tira dela a "Contradição". Nenhuma fonte do grupo dá número ou data da lei. O arquivo 1 diz só "aprovação da Lei do Marco Regulatório" (p.28).
- **Correção:** até conferir o texto oficial, reescrever a sobreposição 4 sem a lei: "lotes vendidos em out./2024 a set./2025 e listados como inedificáveis". O argumento se mantém para as vendas depois da data-base (achado 1).

### 10. MÉDIA: o critério de APP é inconsistente dentro das mesmas quadras, e isso não foi apontado
- **Problema:** o grupo lista vizinhos imediatos como inedificáveis e dá valor cheio a lotes próprios nas mesmas quadras:
  - **Vicentino, quadra 1303:** lotes próprios 007, 008, 009 e 010 a R$ 1.355.000 cada (T33); o 013 como possível APP; os lotes de terceiros 001, 012, 014 e 016 da mesma quadra como inedificáveis a R$ 1.355.000 (T47.L33-36).
  - **Jardim Itália:** quadra 1741, lotes próprios 002, 003 e 004 cheios e terceiros 007 e 009 indenizáveis; quadra 1747, próprios 015 e 016 cheios e terceiros 001, 002 e 009 indenizáveis (T34 x T47).
  - **Dalzóchio, quadra 1292:** o mesmo padrão, já notado em sei-2025-notas.md.
- **Correção:** acrescentar como divergência a levar ao pacote 5 e ao pedido de mapa GEOFarroupilha. Sem delimitação por lote, o mesmo critério teria de valer para os lotes próprios das mesmas quadras.

### 11. BAIXA: Acordi. O pagamento dos R$ 130.000 não tem prova, e os valores não fecham
- **Problema:** a aba Processos diz "Multa paga com desconto (R$ 61.920,30 + R$ 130.000 de acerto)". O ACERTOS E ACORDOS não diz que os R$ 130.000 foram pagos. O doc 11 (P145) fala em multa municipal de "aproximadamente R$ 70.000,00 (já quitada)" e não cita os R$ 130.000.
- **Os totais não fecham:** ACERTOS dá 88.457,52 + 61.920,30 + 130.000,00 = R$ 280.377,82. O doc 11 dá cerca de R$ 80.000 + R$ 70.000. As multas de origem também divergem: ACERTOS fala em cerca de R$ 1.200.000, o doc 11 em cerca de R$ 1.000.000, e 2018 em R$ 1.499.492,79 + R$ 1.069.049,00.
- **Atribuição incerta:** no gdoc, a linha "Multa na Prefeitura ... Acerto com o Município" (P55) vem depois dos marcadores "30.)" e "31.)". Atribuí-la ao item 29 é interpretação.
- **TAC:** o doc 11 diz que ainda falta "firmar um TAC junto ao Ministério Público" (P143). Chamar o valor pago de "TAC R$ 88.457,52" mistura as duas coisas.
- **Correção:** escrever "pagamento dos R$ 130.000 não comprovado" e registrar a divergência de valores.

### 12. BAIXA: Jardim Itália. "Processo judicial" (doc 11) x "inquérito civil" (arquivo 1)
- A obs. 4 do doc 11 fala em "Processo Judicial impetrado pelo Ministério Público". O arquivo 1 (p.24) trata de "IC 00771.001.451/2020", que é inquérito civil, arquivado em jul./2026.
- A aba Processos adota "Inquérito civil" sem registrar a contradição. Anotá-la.

### 13. BAIXA: Vicentino 003/1301. "Indenização maior que o preço da venda" é extrapolação
- A nota do arquivo 17 é ambígua: "Vendido à vista, com parte permuta de outro imóvel, no valor de R$ 500.000,00". Não se sabe se R$ 500.000 é o preço total ou o valor do imóvel permutado.
- A contrapartida em imóvel é o lote 623/104 do Reserva Paradiso, recebido pela SEI na mesma data (arquivo 17, linha 64, p.4).
- Corrigir para "preço informado de R$ 500.000 (redação ambígua), com parte paga com o lote 623/104 do Reserva Paradiso, recebido na mesma data".

### 14. BAIXA: sobreposição 8. "Rafael não menciona a compensação" não é exato
- Rafael transcreve a frase "Parte da área foi destinada para compensação ambiental do Loteamento Colina do Retiro II" (Áreas rurais, p.3, p.20, p.21 e p.24).
- Ele só não a considera na avaliação (faixa de R$ 250.000 a R$ 1.200.000, p.21).
- Trocar por "transcreve a compensação mas não a considera no valor".

### 15. BAIXA: área rural 09 (lote 33 de Cotiporã, "MATO NATIVO") ficou fora do levantamento
- O doc 11 (T9.L2) descreve 15,1 ha de "MATO NATIVO", a R$ 40.000.
- Rafael avalia em R$ 250.000 e cita as restrições do Código Florestal e da Lei da Mata Atlântica (Áreas rurais, p.18 a 20).
- É restrição ambiental explícita no doc 11. Incluir na aba Itens, mesmo sem processo.

### 16. BAIXA: Pasqual e Dr. Letti. O risco de APP só aparece na nova versão do doc 11, e as contas pagas contradizem "sem trâmite"
- **Versões do doc 11:** na versão antiga, Bonetto, Pasqual e Dr. Letti não tinham a frase do risco de APP nem multa. A frase entrou junto com as multas de R$ 782.000, R$ 1.500.000 e R$ 1.650.000 (diff, linhas 306-308).
- **Contas pagas:** registram gastos em 2023 e 2024:
  - "SERVICOS TOPOGRAFIA LOT. PASCOAL", R$ 7.100 (linha 1650);
  - "TESTES SOLO LAUDO GEOLOG" do Loteamento Pascoal (linha 2834);
  - "FRETE 02 RESERVATORIOS P/ LOT. DR. LETTI" (linha 428).

  Isso contrasta com "sem nenhuma trâmite de regularização" (arquivo 1, 18.5) e com "sequer projeto" (doc 11, Letti).
- **Correção:** anotar nos itens "Contratado com risco de APP".

### 17. BAIXA: posição de Rafael sobre as observações 2 a 7 (R$ 7.910.000) não registrada
- Rafael aceita o bloco de R$ 16.127.652 das observações à razão de 1/6 (R$ 2.687.942), sujeito a conferência documental (DEMAIS COMPROMISSOS, p.2-3).
- Nos R$ 16.240.000 dos lotes de terceiros faz o contrário: trata o valor como patrimônio e cobra crédito de 1/6 (p.4-5). Nesse PDF ainda há divergência interna: R$ 2.706.666,67 x R$ 2.706.333,33.
- Registrar na aba Itens e apontar a inversão de sinal.

### 18. BAIXA: afirmações por inferência que deveriam vir marcadas como tal
- **Simioni, "Não impede a venda" (Processos):** é inferência a partir do TAC, que obriga a avisar os compradores. Sem cópia do TAC.
- **Simioni, "O item 18.6 trata de outro empreendimento":** plausível, porque o 18.6 está na lista de áreas contratadas sem contrato. Mas o arquivo 1 não diz isso.
- **"Acordi fora do valor" (Fontes, 2018):** o P69 de 2018 tira as infrações do Acordi do valor dos loteamentos. Mas o total de ações deduzido (R$ 6.278.430,81 + 30% = R$ 8.161.960,05) lista os processos 16 e 17. A soma das 19 causas é R$ 5.187.216,15, diferente do total declarado.
- **Perini, matrículas 34892, 24305 e 24306 indicadas à penhora:** não se verificou se são lotes do estoque, por exemplo os Perini 002, 003 e 004/1688 de R$ 1.520.000 cada. Se forem, há restrição sobre bens a dividir.
- Marcar todas como "inferência" ou "a verificar".

## Conferido e correto

- **Cobertura dos 44 lotes:** 4 do Simioni, 7 APP do Lago Negro, 7 de possível APP e 26 de terceiros. Lote, quadra, valor, bloco e linha da tabela conferem com o doc 11 (T14, T27, T29-T31, T33, T47). Os subtotais do doc 11 fecham sem os lotes de possível APP: San Marino 2.080.000, Dalzóchio 1.790.000, Primavera 148.000, Vicentino 5.420.000 e Lago Negro 4.021.000.
- **Aba Conferencias:**
  - terceiros: 400.000 + 986.000 + 640.000 + 664.000 + 13.550.000 = 16.240.000 (doc 11, T48; Rafael, terceiros, p.2);
  - possível APP: 2.321.000;
  - "Multas ambientais" do Dr. Pablo: 5.000.000 + 2.200.000 + 530.000 + 80.000 + 100.000 = 7.910.000 (arquivo 1, p.21);
  - contagem de 44 lotes.
- **Observações do doc 11:** valores das obs. 2 a 7 e dos contratados 06 a 08 conferem (782.000, 1.500.000 e 1.650.000; multa de 20% do custo mais R$ 50.000). O Sol Nascente III, sem menção a APP, ficou fora com razão. Não achei outra observação ambiental no doc 11.
- **Processos de 2018:** os 5 ambientais estão todos (08, 10, 15, 16, 17), com número e valor da causa iguais ao docx e ao xlsx (935,50; 8.635,00; 285.000,00; 1.499.492,79; 1.069.049,00). Os de 2015 também (08 "Perini" e 14 Lago Negro). Os demais processos de 2015 e 2018 não são ambientais.
- **Arquivo 1:**
  - Perini: nº 5009118-66.2024, R$ 4.663.325,88, TAC 04/12/2012, embargos 5001467.46.2025, perícia sem perito;
  - Lago Negro: ACP 5000454-62.2016; Wasen lote 20 → 05/05 (acordo ago./2018; arquivo 17 com 2019-04-24); Flach lote 15 → 10/09; Zucoloto R$ 580.000; gravame em jun./2026;
  - SAFRA: nº 5004974-83.2023, acordo em fev./2025;
  - Jardim Itália: IC arquivado em jul./2026;
  - Marco Regulatório (p.28), GEOFarroupilha (p.26), escrow do Dr. Pablo (p.20), Colina do Retiro II com 99 e 96 lotes (p.27).

  Páginas e citações conferem.
- **Casamentos com o arquivo 17:** estão certos 011/1469 (Farina, 2019), 018 e 19/1469 (Efetiva, 10/10/2024), 001/1747 (Erick, 2025), 002/1747 (Ambitec, 2021), 011/1199 (SAFRA, 2025), 003/1301 (Fattori, 2025) e 004/1302 (Basso, 2014). Testei grafias:
  - "O09" com quadra "0910"/"910": o único lote dessa quadra no arquivo 17 é o 03/910, Belusso, 2025, que não está no escopo;
  - "11 B" x "11B", "19" x "019";
  - quadra 1469 do Loteamento Viero (lotes 01, 02, 05, 09 e 10, nenhum no escopo).

  Fora o Lago Negro (achado 2) e o 010/1199, não há venda ou permuta escondida dos lotes do escopo.
- **Rafael:**
  - Simioni R$ 520.000 (p.21), Lago Negro R$ 7.700.000 (p.84), Dalzóchio R$ 131.000 (p.100), Primavera R$ 335.000 (p.106) e Vicentino R$ 1.125.000 (p.116): as citações batem com o PDF;
  - terceiros: transcrição p.1-2, pedido de relatório técnico p.2-4, 1/6 como crédito p.4-5.
- **Sobreposições 1, 2, 3, 4 e 5:** as somas estão corretas (2.731.000; 5.640.000; 332.000; 1.875.000; 360.000).
- **Fatos das sobreposições 3, 4 e 5:** conferem com o arquivo 17 e o arquivo 1 (Ambitec, Farina, SAFRA).
- **Sobreposições 7 e 8:** valores do ACERTOS (88.457,52; 61.920,30; 130.000,00) e decisões ARUR-11/12 (R$ 380.000 cada) conferem com o gdoc e com decisoes-sei.json.
- **Diferenças entre o doc 11 antigo e o atual:** Dalzóchio de 500.000 para 2.200.000; Perini não contabilizado e depois contabilizado; os 7 lotes de possível APP passaram a ser excluídos. Batem com o docx antigo.
- **Colina do Retiro II:** 105 para 99 lotes por APP (arquivo 17, Totais, linha 8; arquivo 1, p.27) e compensação ambiental nas áreas rurais 11 e 12 (doc 11, T9.L4-5).
