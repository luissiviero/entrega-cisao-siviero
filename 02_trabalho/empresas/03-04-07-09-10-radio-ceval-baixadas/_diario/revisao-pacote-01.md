# Revisão independente — pacote 1 (regras em bloco, empresas 03, 04, 07, 09 e 10)

Revisor independente, 28/10/2026. Objeto: `pacotes/pacote-01.md` (md5 conferido no diário: 5359eeb2…) e `pacotes/pacote-01-tabelas.md`, gerado de `planilha-mestre-03-04-07-09-10.xlsx` (md5 b50578cd1d0d5bf73c1ef0881d1c4dbb, igual ao cabeçalho das tabelas). Conferido contra as fontes primárias: inventários 2025 (docs 00, 03, 04, 07, 09, 10 e, para duplicidade, 01, 02, 05, 06, 08 e 11), PDFs de Rafael (Rádio, Ceval, Associação, Turismo, Siviero Transportes, Diretrizes), `ativos-avaliados-valter.xlsx`, inventários 2018 (9.1, 11.1, 7.1, 8.1, 14.1), docs 18 e 25 de 2018, arq. 1 (p. 1–2 e 12–13), arq. 13 e arq. 14 (lidos nas imagens), `planilha-mestre-sei.xlsx` (SEI-2025-AURB-02) e `registro-data-base.md` desta pasta. Scripts e extratos em `scratchpad\revisor-p1\`.

Contagem: **1 alta, 8 médias, 8 baixas.**

---

## Achados

### 1. ALTA — R2 diz "sem valor de ativo" na Siviero Transportes, mas há um crédito em aberto em 31.12.2024 que o pacote não nomeia

- **Evidência.** Inventário 2025 doc 10, par. 22: "Existe pendente um PROCESSO JUDICIAL referente a venda do Ônibus IXO 6000, acordado pelo valor de R$ 115.000,00 … R$ 20.000,00 (Despesas judiciais e advocatícias) e o saldo em parcelas de R$ 5.000,00, cada, a contar de novembro de 2024." Na data-base estavam pagas no máximo 2 parcelas. Sobravam algo entre R$ 85.000 e R$ 115.000, conforme a natureza dos R$ 20.000. Na planilha-mestre isso está em 10-2025-N22 e 10-2018-05 e no S4-7. No `registro-data-base.md` está em S4-DB2, como "pendente".
- **Por que é alta.** Um crédito a receber é um ativo. A frase da R2, "Sem valor de ativo em 2025", aprovada como está, fixa R$ 0,00 para a ST com um ativo conhecido na data-base. Além disso, o acordo começa em nov/2024, depois da baixa da ST (08/04/2024; doc 10 par. 15). Em 31.12.2024 o credor já não pode ser a ST. Então o crédito não está em nenhum inventário de 2025 e pode sumir de todos os totais. O PDF de Rafael da ST (p. 1–2) **não transcreve** o par. 22: ele nunca considerou o crédito.
- Outros créditos e processos das baixadas que a R2 também não nomeia (todos de 2018, a verificar se existiam em 31.12.2024):
  - ST: processo "Operação Deus tá vendo" (10-2018-N17) e recebíveis dos ônibus vendidos a prazo (10-2018-06/07; doc 25 par. 229–230).
  - Turismo: restituição de ICMS (09-2018-N17).
- **Correção.** Reescrever a R2 como "nenhum **bem** na tabela do inventário de 2025 (grupo = Rafael = doc 00 = R$ 0,00)". Acrescentar uma linha: "Ficam fora da R2 e não são zerados: crédito do acordo do IXO 6000 (10-2025-N22, S4-DB2, S4-7; saldo em 31.12.2024 e titular a apurar na etapa 6), …". Na ficha de fechamento, perguntar ao grupo quem recebe as parcelas.

### 2. MÉDIA — A "concordância expressa" de Rafael tem condições, e as ressalvas 1 e 6 as omitem

- **Evidência.** PDF Rádio p. 4–5. A concordância vale "sem prejuízo da necessária conferência documental e jurídica quanto à titularidade, existência, regularidade registral" e "fica sujeita à validação pelo escritório de advocacia contratado, especialmente quanto à regularidade formal dos documentos, à titularidade dos bens e à adequação dos valores adotados". Os itens 03, 04 e 05 remetem a essa cláusula (p. 5).
- **Por que importa.** Quatro dos seis bens da R1, R$ 2.390.000 dos R$ 2.505.000, entram só por essa concordância. Rafael não avaliou esses itens. Aprovar a R1 sem a condição pode ser lido como renúncia às conferências que o próprio Rafael reservou: box, outorga e, no caso da Toyota, a condição equivalente da p. 12.
- **Correção.** Em cada ressalva com "concordância expressa", escrever "concordância expressa condicionada à conferência de titularidade, existência e regularidade (PDF p. 4–5)". Acrescentar à R1: "a aprovação fixa o valor, não a titularidade nem a existência".

### 3. MÉDIA — Ressalva 1 (canal, R$ 1.900.000): citação inexata, e falta dizer que a estimativa é antiga

- **Evidência.**
  - O inventário 2025 doc 07, par. 14, diz "estimado por algumas informações verbais, **no passado**, do Senhor Élcio Siviero". O pacote cita "por informações verbais" e corta o "no passado".
  - O mesmo parágrafo já está, palavra por palavra, no inventário 2018 da Rádio (arq. 9.1, par. 15).
  - O grupo responde (arq. 1 p. 12–13) que o valor "seguiu o valor de mercado na data da sua elaboração", com a "expertise … [do] saudoso sócio Elcio Siviero".
  - O item 02 de 2018, de R$ 1.700.000 com os móveis, virou em 2025 R$ 2.050.000 (itens 02 + 03) sobre a mesma base.
- **Por que importa.** É o maior item da R1 (76 % do total). O valor comum não vem de avaliação na data-base: é uma estimativa verbal antiga de um sócio já falecido, que Rafael só aceitou com a condição do achado 2. Isso não muda o valor (P5), mas o Luís precisa ler isso na ressalva.
- **Correção.** Citar o texto literal com "no passado". Registrar que a observação é a mesma de 2018 e que a resposta do grupo está no arq. 1 p. 12–13. Ver também o achado 10.

### 4. MÉDIA — Ressalva 3 (box): não diz que a titularidade em 31.12.2024 está pendente; a certidão é de 2016

- **Evidência.**
  - A planilha-mestre, linha 07-2025-05, coluna "Registro de data-base", diz: "Titularidade em 31.12.2024 a conferir no arq. 13 (matrícula)". A ressalva 3 não fala disso.
  - Arq. 13 p. 3–5: a matrícula 11.324 tem R.6 de "arrematação" pela Rádio (praça em 11/07/2013, registro em 18/01/2016) e a certidão é de **18/01/2016**. A certidão da sala (11.382, p. 2) é de 25/04/2013.
  - Não há certidão atualizada, apesar de o grupo falar em "matrículas atualizadas" (arq. 1 p. 12).
- **Correção.** Acrescentar à ressalva 3: "titular em 31.12.2024 não comprovado; certidão de 2016; pedir certidão atual". Acrescentar também, para o S4-9: a matrícula mostra que o box é unidade autônoma ("box-garagem nº 04 sito no subsolo", área total 17,28 m², arq. 13 p. 3). Isso pesa contra a ideia de que os 94,51 m² da sala incluem a vaga.

### 5. MÉDIA — Ressalva 4 (Toyota): falta a titularidade na data-base e a condição de Rafael; e o tratamento não é simétrico ao do caminhão

- **Evidência.**
  - A planilha-mestre, linha 07-2025-07, diz: "Verificar titular em 31.12.2024 (ST baixada em 08/04/2024; sem documento da transferência)".
  - PDF Rádio p. 12: a seleção de Rafael "fica condicionada à verificação da titularidade, existência, regularidade documental e jurídica dos bens".
  - A faixa de Rafael (p. 7) vai de R$ 45.000 a R$ 120.000, com FIPE de R$ 15.000 a R$ 30.000.
- **Por que importa.** É o único bem destas empresas que vem para "nós" (confirmado em `ativos-avaliados-valter.xlsx`, linha 15, e em Diretrizes p. 7). A ressalva só diz que "a transferência … não tem documento". Ela não diz que, em 31.12.2024, o veículo pode estar em nome de uma empresa extinta.
- **Assimetria.** A ressalva 5 discute subir o caminhão até a faixa de Rafael como "caso extremo", que favorece o C0 do lado do maior valor. A ressalva 4 não registra o lado favorável do C0 na Toyota (o piso da faixa, R$ 45.000, ou a FIPE). O C1 pede registrar o argumento técnico nos dois casos.
- **Correção.** Acrescentar à ressalva 4 a titularidade em 31.12.2024 (DETRAN/CRLV) e a condição da p. 12. Registrar a faixa de R$ 45.000 a R$ 120.000 (p. 7) e a FIPE de R$ 15.000 a R$ 30.000 como C1, sem propor mudança. Ou tirar das duas ressalvas a discussão de "caso extremo" e levá-la ao Luís em pergunta própria (achado 6).

### 6. MÉDIA — Ressalva 5 (caminhão LK 1313): uma decisão do Luís aparece como mecânica e já vem com recomendação

- **Evidência.**
  - Inventário 2025 doc 04, item 03: "avaliado **na data da compra** no valor de R$ 45.000,00". O inventário 2018 (arq. 11.1, item 04) traz o mesmo texto com **R$ 38.000**. Logo, "data da compra" é frase de modelo, e nenhum dos dois números é avaliação em 31.12.2024.
  - Rafael (PDF Ceval p. 4) dá mercado de R$ 90.000 a R$ 150.000, de 2 a 3,3 vezes o valor adotado, e mesmo assim mantém R$ 45.000 no balanço (p. 5).
  - O bem fica na empresa: pelo C0, o lado que interessa é o maior valor.
- **Por que importa.** A R1 deve cobrir só o que é mecânico. Aqui o "acordo" é formal: Rafael copiou o valor do grupo e, no próprio texto, diz que o mercado está muito acima. Decidir se é "caso extremo" cabe ao Luís (regra de não reavaliar, salvo caso extremo justificado). O pacote já traz "que não recomendo".
- **Correção.** Tirar o LK 1313 da R1 e fazer dele uma pergunta neutra de sim/não. Isso pode ir no pacote 2 ou ficar à parte. A pergunta deve trazer: valor de compra × faixa de Rafael, FIPE de R$ 30.600 abaixo do valor adotado, e efeito no 1/6 de no máximo R$ 17.500. Se ficar na R1, apagar a recomendação e escrever "decisão do Luís: manter R$ 45.000 (padrão) ou tratar como caso extremo".

### 7. MÉDIA — R2: "Rafael só transcreve o grupo" é incompleto; Rafael faz uma reserva expressa

- **Evidência.** Nos PDFs da Associação, da Turismo e da ST, sempre na p. 2, Rafael acrescenta a seção "DA BAIXA E DA REGULARIDADE DA EXTINÇÃO DA EMPRESA". Nela, a baixa fica "pendente de validação formal, não se presumindo … a inexistência de eventual patrimônio, direitos ou obrigações remanescentes". O PDF da ST omite o par. 22 do grupo (IXO 6000).
- **Correção.** Trocar por: "Rafael repete o quadro do grupo (R$ 0,00), mas ressalva que a baixa precisa ser validada e que não se presume a falta de patrimônio, direitos ou obrigações (p. 2 de cada PDF); o PDF da ST não traz o acordo do IXO 6000". Isso reforça o achado 1.

### 8. MÉDIA — "8 linhas só de 2018 (R$ 2.526.100,00 em 2018)" mistura bens com ônibus vendidos antes de 2018

- **Evidência.** As 8 linhas se dividem em dois grupos:
  - 5 bens, que somam R$ 1.011.000: 04-2018-03 (R$ 93.000), 03-2018-TOTAL (R$ 10.000) e 10-2018-01, 02 e 03 (R$ 250.000, R$ 70.000 e R$ 588.000).
  - 3 ônibus vendidos a prazo em 2015 e 2016 (10-2018-05, 06 e 07, R$ 1.515.100). O próprio inventário 2018 da ST tira esses ônibus do total: "R$ 2.443.100,00 / (−) R$ 1.515.100,00 / R$ 928.000,00" (arq. 7.1, tab. 3).
- **Por que importa.** R$ 2.526.100 não é valor de 2018 de nada. Lido como benchmark, superestima em R$ 1.515.100 o que "sumiu" entre 2018 e 2025.
- **Correção.** Escrever: "5 linhas de bens de 2018 (R$ 1.011.000,00) e 3 ônibus vendidos a prazo antes de 2018 (R$ 1.515.100,00, fora do total de 2018), todos para a etapa 6".

### 9. MÉDIA — "Efeito" e "Ficam fora" deixam de fora os créditos e dívidas (aba "Créditos e dívidas", 23 notas)

- **Evidência.** O pacote diz "Todas as 19 linhas da planilha estão enquadradas" e "Restam as 4 exceções … e os pontos de nível de empresa". Ficam sem menção, entre outras:
  - a nota promissória de Celito na Ceval, R$ 150.000 (04-2018-N13, S4-4);
  - a dívida da Rádio de 1 lote do Colina do Retiro, R$ 110.000 (07-2018-N14, S4-6);
  - a Associação subsidiada pela Rádio (03-2025-N14);
  - a dívida da ST com a SEI, R$ 2.003.000 (10-2018-N14, S4-5);
  - o IXO 6000 (achado 1).
- **Correção.** Acrescentar a "Ficam fora das regras" uma linha: "23 notas de créditos, dívidas e intragrupo (aba 'Créditos e dívidas'), nenhuma coberta por R1–R3; vão para a etapa 6 e para a Parte 6 (S4-4 a S4-7)".

### 10. BAIXA — Ressalva 1: citações e fatos da outorga

- O protocolo de renovação de 22/05/2026 está no arq. 1 **p. 12** e no arq. 14 p. 2–3, não na p. 13. A p. 13 só traz a classe.
- "Classe A3 com pedido para A2": a A3 aparece na licença (arq. 14 p. 1). O pedido de A2 **não tem documento** no arq. 14, que só traz a licença, o e-mail, o recibo e o requerimento de renovação. Há apenas a afirmação do grupo, "conforme documentos anexos" (arq. 1 p. 13). Dizer isso no texto.
- A situação em 31.12.2024 já se lê no arq. 14 p. 1: licença emitida em 13/01/2022, "Validade da radiofrequência: 08/04/2027", FM 96,1, classe A3. A outorga estava vigente na data-base. A renovação (período de 08/04/2027 a 08/04/2037) é fato posterior e não muda número. A ressalva pode dizer isso já, em vez de "será levantada à parte", e atualizar S4-DB3.
- Ligar a ressalva a S4-DB3 do `registro-data-base.md`. O registro foi criado depois do pacote.

### 11. BAIXA — Ressalva 3: a frase de Rafael foi atribuída à sala

PDF Rádio p. 3: "O **anúncio** menciona 94,51 m² de área total, o que inclui as áreas comuns do prédio e a vaga de garagem". O pacote escreve "Rafael diz que a área da sala já 'inclui a vaga de garagem'". Reescrever como "Rafael, comentando um anúncio de 94,51 m² (mesma área da sala), diz que a área inclui a vaga…". Ver também a matrícula do box no achado 4.

### 12. BAIXA — R1 adota valores brutos, e isso não está escrito

- A redução de 8 % incide sobre o "valor acima apontado" da empresa (doc 07, par. 17).
- O quadro de Rafael (p. 10) calcula o 1/6 sobre os 100 % da Rádio, sem os 8 % e sem descontar Telmo de Toni.
- O pacote remete os dois pontos ao pacote 2, o que está certo, mas o leitor pode tomar os R$ 2.505.000 por valor líquido. Acrescentar: "valores brutos; a redução de 8 % e a base do 1/6 com Telmo de Toni são decididas no pacote 2 e podem alterar o valor da Rádio".

### 13. BAIXA — S4-DB4 ("outubro/2025") não aparece

O arq. 1 (p. 1–2) diz que os inventários se referem a um "termo final adotado em prévio consenso como sendo outubro/2025". O pacote apresenta tudo como 31.12.2024. Isso está registrado de forma preventiva em S4-DB4. Vale uma linha nas ressalvas da R1 ("registrado em S4-DB4; nenhum número afetado").

### 14. BAIXA — R3: "sempre da SEI" se apoia em declarações

As bases são o doc 18 tab. 31, o doc 25 par. 204, o doc 07 2025 par. 15 e o arq. 1 p. 12. Não há matrícula 3.751 nas fontes; a revisão da mestre já pedia o titular registral. Trocar "sempre da SEI" por "declarado como da SEI em 2018 e em 2025". O valor e o enquadramento não mudam.

### 15. BAIXA — Duas regras com o mesmo nome "R3"

"Regra R3 da SEI" e "R3" deste pacote se confundem. Escrever "regra R3 do log de decisões da SEI (bem avaliado só pelo grupo)".

### 16. BAIXA — Página citada para a Toyota na amostra

Rafael, item 7: o texto começa na p. 9 e o valor está na p. 10. Escrever "p. 9–10". Não muda nada.

### 17. BAIXA (fora do pacote, na planilha-mestre) — a Toyota é contada em duas empresas no resumo

Na aba "Resumo por empresa", a Toyota entra em 07 Rádio (8 linhas, "Concorda: 5") e em 10 ST (7 linhas, "Concorda: 1"). A soma dá 20 linhas para 19 reais. O pacote não usa esses números, mas convém corrigir na próxima regeneração.

---

## O que conferi e está certo

- **Valores da R1.** Os 6 bens têm o mesmo valor no inventário 2025 e no balanço executado de Rafael:
  - Rádio: móveis R$ 150.000; equipamentos e canal R$ 1.900.000; antena R$ 300.000; box R$ 40.000; Toyota R$ 70.000 (doc 07, tab. 3, itens 02–05 e 07; PDF Rádio p. 9–10).
  - Ceval: LK 1313 R$ 45.000 (doc 04, item 03; PDF Ceval p. 5).
- **Enquadramento.** Nenhum bem que concorda ficou de fora, e nenhum que discorda entrou. Os 4 que discordam são: sala (R$ 210.000 × R$ 420.000), Fox (R$ 28.000 × R$ 32.000), 320D (R$ 200.000 × R$ 240.000) e Case (R$ 70.000 × R$ 100.000).
- **Páginas de Rafael.**
  - Concordância: p. 4 para o item 02 e p. 5 para os itens 03, 04 e 05.
  - Faixas: Toyota de R$ 45.000 a R$ 120.000 (p. 7); LK 1313 de R$ 90.000 a R$ 150.000 e FIPE de cerca de R$ 30.600 (p. 4).
  - Box: "agregando entre R$ 35.000 e R$ 50.000 ao valor final do conjunto" (p. 3).
  - Toyota escolhida: p. 10–11.
- **Lista C0.** A Toyota é o único bem destas empresas na lista de Rafael (`ativos-avaliados-valter.xlsx`, linha 15; Diretrizes p. 7). A Ceval não tem bens selecionáveis (PDF Ceval p. 6).
- **Duplicidade.** Nenhum bem da R1 aparece nos outros inventários de 2025 (docs 01, 02, 05, 06, 08 e 11), buscando por placa, matrícula, "Bandeirante", "1313" e "antena".
- **Totais.** Todos conferem:
  - R1: R$ 2.505.000 (Rádio R$ 2.460.000; Ceval R$ 45.000).
  - 10 bens de 2025: R$ 3.013.000 (R$ 2.698.000 + R$ 315.000, iguais ao doc 00).
  - Exceções: R$ 508.000 pelo grupo e R$ 792.000 por Rafael. Rafael também fecha: R$ 2.505.000 + R$ 792.000 = R$ 3.297.000 = R$ 2.912.000 + R$ 385.000.
  - A soma das 8 linhas é R$ 2.526.100 (a conta está certa; o problema é o rótulo, achado 8).
  - Efeito no 1/6: (R$ 150.000 − R$ 45.000) / 6 = R$ 17.500.
  - 19/19 linhas enquadradas.
- **Ressalvas que conferem com as fontes.**
  - Antena no terreno da SEI (doc 07 par. 15; doc 25 par. 204).
  - Os 8 % incluem mudar a antena e comprar as quotas de Telmo de Toni (arq. 1 p. 13).
  - Móveis desmembrados do item 02 de 2018 (par aprovado em `decisoes.json`).
  - Protocolo de 22/05/2026 e classe A3 (arq. 14 p. 1–3).
  - Transferência da Toyota sem documento (ST 2018 item 04 → Rádio 2025 item 07).
- **R2.** As três empresas valem R$ 0,00 no grupo, no doc 00 e no quadro que Rafael repete. As baixas são anteriores à data-base: 09/07/2020, 04/09/2023 e 08/04/2024. A venda de quotas de 28/08/2025 é posterior e está registrada em S4-DB1, sem mudar número.
- **R3.** O terreno 3.751 está na SEI 2025 (doc 11, tab. 7, item 02). Ele foi adotado na `planilha-mestre-sei.xlsx` (SEI-2025-AURB-02) por R$ 7.148.016, "R3 … Aprovada em 2026-10-28", e não está na lista de Rafael. Em 2018 entrou só no total da Rádio (R$ 9.316.016), já que a SEI o lista "relacionado no inventário da Rádio", fora do total (doc 18, tab. 31). O valor na SEI e zero aqui não deixa o terreno sem valor nem com valor dobrado em 2025. O ajuste do benchmark 2018 da Rádio (R$ 2.168.000 sem o terreno) já está na mestre.
- **Decisões do Luís não tomadas por conta própria.** O pacote não aplica os 8 %, a base do 1/6 com Telmo de Toni nem a anotação da NF 4946 (E7): manda tudo ao pacote 2. A única decisão tratada como mecânica é a do LK 1313 (achado 6).
