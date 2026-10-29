# Revisão independente final — todas as decisões aprovadas (empresas 03, 04, 07, 09 e 10)

Revisor de contexto novo, 28/10/2026. Esta revisão vem antes das fichas de fechamento (P11). Não reabre mérito de decisão do Luís. Aponta erro, incoerência, fato novo ou risco.

**Objeto:** `log-decisoes.md`, `decisoes.json` (md5 1762dffd…), `planilha-mestre-03-04-07-09-10.xlsx` (md5 22f8e2e0…; abas Leia-me, Mestre, Créditos e dívidas, Resumo), `pacotes/pacote-01/02/03.md` e tabelas, `parte6-candidatos.md`, `perguntas-ao-grupo.md`, `registro-data-base.md`, `_scripts/build_mestre.py` e as revisões anteriores em `_diario/revisao-*.md`. Os md5 de todos esses arquivos batem com o último registro do `diario.md`.

**Fontes conferidas diretamente:**
- inventários 2025 (docs 00, 03, 04, 07, 09, 10 e 11, relidos do .docx; varri também 01, 02, 05, 06 e 08);
- inventários 2018 do Drive (7.1, 9.1, 11.1 e doc 27);
- PDFs de Rafael da Rádio (p. 3, 6, 7, 9–12) e da Ceval (p. 2–6);
- `ativos-avaliados-valter.xlsx`;
- arq. 17 (xlsx);
- `Contas_Pagas_2023-07_a_2024-10.xlsx` (aba Consolidado);
- da SEI, só leitura: `planilha-mestre-sei.xlsx`, `sei-2018-itens.xlsx`, `sei-2025-itens.xlsx`, `log-decisoes.md`, `parte6-candidatos-duplicidade.md`;
- da S3: `rovilio-quotas.md` e `pacotes/levantamento-relatorios-mensais.md`.

Scripts em `…\scratchpad\revisor-final\`.

**Contagem: 0 altas, 8 médias, 10 baixas.** Nenhum valor adotado errado, nenhuma decisão aplicada fora do aprovado e nenhuma direção do C0 invertida.

---

## Achados

### Média

#### M1. Planilha, linha 10-2018-05 (IXO 6000): a coluna "Registro de data-base" ainda diz "crédito em aberto"
- **Evidência:** a célula diz "S4-DB2: acordo do IXO 6000 com parcelas desde nov/2024 (crédito em aberto em 31.12.2024)". O texto vem de um mapeamento fixo por ID no `build_mestre.py`.
- **O que foi aprovado (pacote 3, item 1):** a natureza do acordo é **inferência** ("provável crédito"; a fonte não diz quem paga a quem), com leitura alternativa (grupo devedor). O `registro-data-base.md` (S4-DB2) e a nota 10-2025-N22 da própria planilha já dizem isso. O `pacote-01.md` (R2) também diz "em aberto em 31.12.2024", mas é texto anterior à correção.
- **Risco:** a ficha da ST copiar "crédito em aberto" como fato.
- **Correção:** no script, trocar o texto por "S4-DB2: acordo do IXO 6000 (parcelas desde nov/2024; natureza, titular e saldo a confirmar — pergunta 10)" e regenerar.

#### M2. Planilha, "Critério de Rafael" = "concordância expressa com o valor do grupo" em 4 linhas da R1
- **Linhas:** 07-2025-02, 03, 04 e 05.
- **O que foi aprovado (log, R1):** "concordância de Rafael é condicionada, PDF Rádio p. 4–5 e 12". A correção nº 2 da revisão do pacote 1 foi feita no pacote, mas não chegou à coluna da planilha. O `pacote-01-tabelas.md` também repete "expressa".
- **Correção:** rótulo "concordância condicionada à conferência de titularidade, existência e regularidade (p. 4–5)".

#### M3. Planilha, aba Resumo: faltam o total adotado e a quota
- **Evidência:** a aba tem soma do grupo, soma de Rafael e totais de 2018, mas nenhuma coluna com a soma dos valores adotados nem com a base e o valor da quota.
- **Onde os números estão hoje:** R$ 2.912.000 / 16,50 % = R$ 480.480,00 e R$ 385.000 / 1/6 = R$ 64.166,67 só aparecem no log e no diário.
- **Conferi por script, a partir da coluna "Valor adotado":** Rádio R$ 2.912.000 e Ceval R$ 385.000. O resultado é o mesmo.
- **Risco:** a ficha transcrever os números à mão, sem a planilha como fonte.
- **Correção:** acrescentar no Resumo as colunas "Soma adotada", "Base da quota (decisão)" (Rádio 16,50 %, pacote 2 item 7; demais 1/6) e "Quota".

#### M4. S4-11 e S4-12 (aprovados no pacote 3) não aparecem na planilha
- **Evidência:**
  - Nenhuma linha ou nota da planilha cita S4-11 (empréstimo Caixa "Rádio") ou S4-12 (atividade da Ceval em 2023–2024, R$ 446.244,50).
  - Os demais candidatos (S4-1 a S4-10) estão ligados por ID.
  - O Resumo da Ceval não sinaliza que o inventário omite caixa, recebíveis e passivos, nem a possível omissão do 2423 K.
- **Risco:** a ficha da Ceval, montada a partir da planilha, sair sem o principal ponto aberto da empresa.
- **Correção:** criar notas próprias ("04-2024-ATIV", S4-12, perguntas 8 e 17; "07/10-2023-CAIXA", S4-11, pergunta 15) ou uma coluna "Pontos abertos (Parte 6 / perguntas)" no Resumo.

#### M5. Coerência com a SEI: a planilha da SEI casa as escavadeiras de forma deslocada, e o S4-2 não tira a consequência
- **Evidência:**
  - `sei-2018-itens.xlsx`: a SEI tinha em 2018 a 320DL 2014 (item 01, R$ 400.000) e a **320D 2009** (item 02, R$ 250.000).
  - `sei-2025-itens.xlsx`: a SEI tem em 2025 a 320 2024 (01, R$ 887.000, financiada pela Caterpillar, obs. 20), a 320DL 2014 (02, R$ 350.000) e a "320D 2009 – SEI" (03, R$ 200.000).
  - A `planilha-mestre-sei.xlsx` casou por similaridade de texto 2025-01 ↔ 2018-01 e 2025-02 ↔ 2018-02 (320DL 2014 com a 320D 2009). Com isso, marca a MAQ-03 "320D 2009 – SEI" como "Existe em 2018? Não".
  - O par natural é 2018-02 → 2025-03 e 2018-01 → 2025-02; a 2024 é nova.
- **Consequência para estas empresas:** em 2018 e em 2025 coexistem duas 320D 2009, a da SEI (sem NF nas fontes) e a da Ceval (NF 4946, série CAT0320DAA6F00902). Isso é **indício de duas máquinas**, do mesmo tipo do S4-3. Hoje o S4-2 só diz "em 2018 a SEI já tinha uma 320D 2009". Não muda número aqui (a 320D da Ceval fica em R$ 240.000).
- **Risco:** a ficha (ou a SEI) afirmar que a "320D – SEI" só aparece em 2025.
- **Correção:**
  - No S4-2: "indício de duas máquinas: SEI 2018 item 02 (R$ 250.000) e Ceval 2018 item 01 (R$ 200.000) coexistiam; manter a pergunta 1".
  - Avisar a sessão SEI do casamento deslocado, na união da P13. O arquivo da SEI é só leitura para esta sessão.

#### M6. Coerência com a SEI: o lote 004/1302 do Vicentino, trocado pelos ônibus da ST, é hoje passivo indenizável da SEI
- **Evidência:**
  - Arq. 17 (xlsx, p. 3): "2014-02-26 | 04 | 1302 | Alcides Basso Imóveis | Vendido, com permuta de 03 veículos – ônibus + R$ 552.000,00 em 60 parcelas…". Os 3 ônibus são os da ST 2018, "adquirido em troca de um lote no Loteamento Industrial Vicentino" (arq. 7.1, itens 05–07).
  - `planilha-mestre-sei.xlsx`, aba Lotes, linha 458: "Industrial Vicentino | 1302 | 004", em "Lotes de terceiros passíveis de indenização" (doc 11, tab. 47, linha 29). Valor do grupo R$ 1.355.000; **passivo adotado R$ 677.500** (R6 nível B, pacote 5 da SEI).
- **Lacuna:** o S4-7 ("Onde ficou o resultado?") cita o lote como "outro lado: SEI" e não registra que o mesmo lote gera passivo na SEI.
- **Risco/argumento:** a SEI assume uma indenização de R$ 677.500 (R$ 112.916,67 no 1/6) por um lote cujo produto foi para a ST. Os ônibus foram vendidos a prazo, os recebíveis de 2018 (R$ 340.000) não têm destino, o acordo do IXO 6000 tem titular incerto, e a ST foi baixada.
- **Correção:** acrescentar ao S4-7 a referência ao passivo da SEI (doc 11 tab. 47; pacote 5 da SEI, R6) e cruzar com a lista da SEI na união (P13).

#### M7. Coerência com a S3: a parcela 28/28 do empréstimo Caixa "Rádio" tem leituras opostas nas duas sessões
- **S3** (`01-08-abastecedora-transportadora/pacotes/levantamento-relatorios-mensais.md`, linha 61): "Última parcela do empréstimo Caixa da Rádio, paga à Siviero Transportes … Intragrupo, quitado antes da data-base, sem efeito".
- **S4** (pacote 3 aprovado, S4-11): "pagou à ST" não decorre da fonte. A coluna "Fornecedor" é a empresa à qual a despesa é atribuída. Devedor e conta entre Abastecedora, ST e Rádio ficam a confirmar (pergunta 15).
- A pergunta 15 envolve a empresa 01, que é da S3.
- **Risco:** fichas de sessões diferentes registrarem conclusões contraditórias sobre o mesmo lançamento, e a pergunta ir duplicada ou contraditória ao grupo.
- **Correção:** alinhar a leitura na união (P13) e decidir com a S3 quem faz a pergunta 15.

#### M8. `log-decisoes.md`: duas frases desatualizadas contradizem decisões posteriores do próprio log
- **Fase 0:** "Ainda sem resposta do Luís: uso do achado E7 … Até a resposta, o E7 não é usado." O pacote 2, item 4b, resolveu o ponto: a hipótese do E7 foi retirada.
- **Pacote 2:** "Perguntas ao grupo propostas no pacote 2 (7): 'Ainda não' — ficam só no pacote." Desde o pacote 3, elas estão na lista única aprovada (`perguntas-ao-grupo.md`, nº 1–7; nada enviado).
- **Correção:** acrescentar "(resolvido no pacote 2, item 4b)" e "(incorporadas à lista única do pacote 3; envio continua decisão do Luís)".

### Baixa

#### B1. Planilha, linha 03-2018-TOTAL: "Motivo da decisão" vazio
O `decisoes.json` tem o motivo, mas o bloco que cria a linha (`build_mestre.py`, ~linha 474) preenche só a "Decisão (P5)". Usar `decision_cols(rid)` também nessa linha.

#### B2. Leia-me da planilha desatualizado
- Pares: ainda diz "marcado 'PROPOSTO' até aprovação" e "é proposto como desmembrado". Os dois pares estão aprovados.
- Situação: ainda diz "terreno 3.751, **sempre da SEI**", contra o aprovado "declarado como da SEI" (correção 14 do pacote 1). O texto está fixo no `build_mestre.py`, linhas 670 e 678.

#### B3. Tabelas dos pacotes 1 e 2 (históricas, geradas antes das correções e decisões)
- `pacote-01-tabelas.md` mantém "concordância expressa" e "Sempre da SEI".
- `pacote-02-tabelas.md` calcula os efeitos da Rádio sobre 1/6. Com a base decidida de 16,50 %, ficam:
  - sala: +R$ 34.650,00 (não R$ 35.000,00);
  - Fox: +R$ 660,00 (não R$ 666,67);
  - 8 %: −R$ 38.438,40 (não −R$ 38.826,67).
- A ficha deve usar 16,50 % para a Rádio. Não é preciso reescrever os pacotes aprovados; basta não copiar deles.

#### B4. Planilha, linha 07-2018-06: observação do Luís marcada como "Pendente"
A observação ("Em 2018 estava na relação 9.1 RÁDIO…", em 4 células das planilhas Google) já está tratada pela R3 e pelo S4-1. O script marca "Pendente" sempre que existe observação. Marcar "Respondida (R3; S4-1)".

#### B5. Alertas da fonte em formato numérico inglês
Exemplos: "R$ 100,000.00" (sala), "R$ 992.78" (terreno), "R$ 470,000.00" e "R$ 450,000.00" (IXO e IOM). É resto do achado 17 da revisão da mestre. Usar o formato brasileiro.

#### B6. Destino 10-2018-06: recebíveis atribuídos ao IMY 5000 sem marca de inferência
O destino liga ao IMY 5000 os "Recebíveis de 2018 (doc 25 par. 229–230: 24 × R$ 10.000 + R$ 100.000)". O doc 25 não identifica os ônibus ("empresa de Santa Catarina"; "do Rodrigo"). Escrever "recebíveis de 2018 dos ônibus vendidos (ônibus não identificado na fonte)" nas três linhas.

#### B7. `registro-data-base.md`: falta a FIPE de out/2025 usada no motivo do LK 1313
O motivo da decisão 04-2025-03 cita a FIPE de out/2025 (R$ 41.201, arq. 1 p. 7). A P2 pede registrar fato posterior mesmo sem mudança de número. Sugestão: S4-DB5, "argumento posterior; nenhum número muda".

#### B8. Rafael, Ceval p. 6: 1/6 escrito como R$ 64.166,00
O valor exato é R$ 64.166,67 (truncamento de Rafael). Registrar na ficha para não parecer divergência.

#### B9. Lançamento das contas pagas não tratado
Consolidado, linha 3399 (set/2024, pago pela SEI): "IMPOSTO SIMPLES S/NOTA N 6631 DA RADIO - RECE…", R$ 3.000. Não está no pacote 3 nem na busca de destinos. Pode ser de rádio de terceiros: Rádio Planalto e Rádio Veranense aparecem como fornecedores. Também pode ser intragrupo SEI × Rádio, do mesmo tipo do S4-12. Valor pequeno: classificar ou cobrir com a pergunta 17 (balanço da Rádio).

#### B10. Na união com a SEI (P13)
O candidato nº 1 da SEI pede "Conferir se também constam no inventário 2025 da Rádio". O S4-1 responde: em 2025, o 3.751 está só na SEI. A resposta precisa ser levada na união.

---

## Pontos abertos e riscos para a ficha de fechamento, por empresa

Efeitos aproximados, para a ficha dimensionar cada ponto. Não são propostas de ajuste (P5). Rádio na base decidida de 16,50 %; demais em 1/6.

### 07 Rádio — total adotado R$ 2.912.000,00; quota de 16,50 % = R$ 480.480,00

1. **Base da quota (decisão do Luís, P1).** Registrar as três leituras que existem nas fontes:
   - 16,50 % (doc 00 e tab. 4 do inventário): R$ 480.480,00 — **decidida**;
   - 1/6 (Rafael p. 10): R$ 485.333,33, com crédito de Rafael de R$ 415.333,33 depois da Toyota (com 16,50 %: R$ 410.480,00);
   - 18 % do "quadro societário real … originariamente" (tab. 1 do inventário 2025: 5 × 18 %, Pedro 9 %, Telmo 1 %): R$ 524.160,00.

   Só registro. A decisão não é reaberta.
2. **Redução de 8 % (não aplicada).** Se o grupo insistir: −R$ 232.960, ou −R$ 38.438,40 na quota.
   - Argumento do outro lado: Diretrizes de Rafael p. 15–16 ("valor econômico líquido").
   - Resposta registrada: custo futuro sem cronograma, intragrupo (S4-1), e a parcela de Telmo já está fora da base (S4-10).
3. **Canal + equipamentos (R$ 1.900.000, 65 % da Rádio).** É estimativa verbal antiga ("no passado", Élcio). Rafael não avalia; só concorda com condição.
   - A outorga estava vigente em 31.12.2024 (até 08/04/2027). A renovação de 2026 é posterior (S4-DB3).
   - É o maior ponto contestável: cada 10 % equivale a R$ 31.350 na quota.
4. **Sala (R$ 420.000, Rafael, contra R$ 210.000 do grupo).**
   - O grupo pode atacar a fonte de Rafael (um anúncio) e alegar dupla contagem do box (S4-9, até R$ 40.000, R$ 6.600 na quota). A matrícula sustenta: 66,02 m² úteis, sem vaga.
   - As certidões são de 2013 (sala) e 2016 (box); o titular em 31.12.2024 não está comprovado (pergunta 3).
5. **Toyota IHW 0285 (C0, lista de Rafael, R$ 70.000).**
   - Transferência ST → Rádio sem documento; IPVA de 2024 pago pela Transportadora; ST baixada (S4-8, pergunta 3).
   - Se a titularidade falhar: a Rádio vale −R$ 70.000 (−R$ 11.550 na quota), e o bem escolhido pode não ser entregável.
   - FIPE de Rafael: R$ 15–30 mil (C1, registrado; sem mudança, P5).
6. **Dívida de 2018 de "01 lote no Colina do Retiro I" (R$ 110.000; S4-6, pergunta 16).**
   - Se o credor for terceiro e a dívida for da Rádio: −R$ 18.150 na quota.
   - Se a dívida for com a SEI: efeito líquido perto de zero.
   - Falta ver se o lote está entre os 3 escolhidos por Rafael no "07. Colina do Retiro" (002/05, 002/09, 003/09). Isso continua aberto desde a revisão da mestre.
7. **Inventário sem caixa, recebíveis nem passivos** de uma empresa em operação (pergunta 17).
   - Empréstimo Caixa (S4-11; leitura divergente da S3, M7).
   - Lançamento "S/NOTA N 6631 DA RADIO" (B9).
8. **Terreno 3.751.** Valor só na SEI (R$ 7.148.016, R3 da SEI). Conferido: está no Mestre da SEI e não aparece em Rafael nem nos outros inventários de 2025. A antena (R$ 300.000) fica sobre terreno da SEI, cedido de forma verbal.
9. **Titularidade formal das quotas.** O quadro "não real" é Eliana 99 % + Telmo 1 %; a cessão aos irmãos está no doc 20 de 2003. Compra das quotas de Telmo: pergunta 6.

### 04 Ceval — total adotado R$ 385.000,00; quota de 1/6 = R$ 64.166,67

Rafael: "não há interesse"; só crédito (p. 6, R$ 64.166,00 — B8).

1. **320D (R$ 240.000, piso da faixa de Rafael).**
   - A fonte de Rafael é um link de "320-dl", contestável.
   - Alienação fiduciária à Caterpillar Financial na NF: quitação a comprovar (pergunta 5).
   - Anotação "Celito": decidida (4a/4b).
   - S4-2: agora com indício de duas máquinas (M5). Se fosse a mesma, sairia a linha da SEI (−R$ 33.333,33 no 1/6 da SEI).
2. **Case 580L (R$ 100.000).** S4-3: indício forte de duas máquinas. Se fosse a mesma: −R$ 10.000 na SEI. Pergunta 2.
3. **LK 1313 (R$ 45.000; caso extremo não adotado).** A faixa de Rafael (R$ 90–150 mil) fica registrada como argumento não usado (até +R$ 17.500).
4. **Caminhão 2423 K ILD 3959.** Em uso em jan/2024 e ausente do inventário: possível omissão (Parte 5). Referência de 2018: R$ 93.000, ou R$ 15.500 no 1/6. Pergunta 7.
5. **Atividade de 2023–2024 (S4-12).**
   - Bases de R$ 446.244,50; possível recebível da CERAN (NF 22, R$ 286.940, set/2024).
   - Tributos e folha pagos pela SEI: saldo intragrupo.
   - Despesas de 2025 (PPRA/PCMSO) também pagas pela SEI.
   - O efeito no 1/6 é desconhecido e pode ir para os dois lados. Perguntas 8 e 17.
6. **NP de Celito (R$ 150.000; S4-4).** Se ainda existir em 31.12.2024 e for contada uma vez, na Ceval ou na SEI: −R$ 25.000 no 1/6. Há indício de quitação. Pergunta 4.
7. **Titularidade formal.** O quadro "não real" é Marcos Antônio Siviero 66,67 % + Pedro 33,33 %.

### 03 Associação Cultural — R$ 0,00 em bens

1. Baixa em 09/07/2020 sem certidão (pergunta 17). A fonte se contradiz: par. 19 "balanço pronto" × par. 20 "nunca foi feito".
2. R$ 10.000 de 2018 sem item (pergunta 20). Sem efeito possível no 1/6: associação não distribui patrimônio.
3. Custos pagos pela Rádio: intragrupo sem valor.

### 09 Turismo — R$ 0,00 em bens

1. Baixa em 04/09/2023 sem certidão; indício: honorário em dobro em out/2023 (pergunta 17).
2. Parcelamento com a Receita de 2018 (R$ 77.000; R$ 12.833,33 no 1/6): quitação provável, por inferência (pergunta 19).
3. ICMS recebido antes de 2018 (declarado). Ônibus "2000" vendido em 2015.
4. Rovílio (S4-DB1; pergunta nº 7 da S3): a Turismo aparece como compradora já baixada.
5. Quadros: o "real" é 5 × 18,20 % + Pedro 8 % + Rovílio 1 %; o formal é Geni 98 % + André 2 %. Só importa se surgir ativo ou passivo remanescente.

### 10 Siviero Transportes — R$ 0,00 em bens

1. **Acordo do IXO 6000** (R$ 115.000; no máximo R$ 95.000 líquidos se o grupo for credor, R$ 15.833,33 no 1/6).
   - A natureza é inferência.
   - Depois da baixa, o titular formal seria Adélia 98 % / Giovani 2 %: risco de o valor sair do rateio de 1/6.
   - Perguntas 10–11; S4-DB2.
2. **Recebíveis de 2018** (R$ 340.000; R$ 56.666,67 no 1/6): destino desconhecido (pergunta 9).
3. **3 ônibus em uso em 2018** (R$ 908.000; R$ 151.333,33 no 1/6, a valores de 2018): sem destino documentado (P4, pergunta 12). É o maior ponto aberto das baixadas.
4. **Dívidas com terceiros de 2018** (R$ 243.800) e parcelamento (R$ 7.000): perguntas 13 e 19.
5. **"Operação Deus tá vendo"**: possível passivo contingente, de valor desconhecido (pergunta 18). Penhora de jul/2024 (pergunta 11).
6. **Cadeia lote 004/1302 → ônibus → recebíveis → baixa** (S4-7), com o lote hoje como passivo indenizável da SEI (R$ 677.500 adotado; M6).
7. **Aportes da SEI** = "injeção financeira" (S4-5, efeito perto de zero). Toyota (S4-8). Linhas do DAER, provavelmente encerradas. Rovílio (S4-DB1).

---

## O que conferi e está certo

- **Decisões × `decisoes.json` × planilha:**
  - As 10 decisões de bens de 2025, as 8 linhas de 2018 sem bem em 2025, a R3 (terreno), as regras de empresa (R2, 8 %, 4a, 4b, 16,50 %), os 2 pares sem chave e os 32 destinos estão no JSON e na planilha, com decisão, valor e destino iguais aos aprovados no log.
  - Nada aplicado sem aprovação; nada aprovado sem aplicação. As exceções de texto estão em M1, M2, B1 e B6.
- **Valores adotados × fontes:**
  - Grupo: doc 04 (200.000 / 70.000 / 45.000; total 315.000) e doc 07 (210.000 / 150.000 / 1.900.000 / 300.000 / 40.000 / 28.000 / 70.000; total 2.698.000). O doc 00 confere (2.698.000; 315.000; 03, 09 e 10 = 0; Telmo 26.980).
  - Rafael: Rádio p. 9–10 (420.000 / 150.000 / 1.900.000 / 300.000 / 40.000 / 32.000 / 70.000; total 2.912.000; 1/6 = 485.333,33). Ceval p. 5 (240.000 / 100.000 / 45.000; total 385.000).
  - Faixas de Rafael: p. 3 (380–520 mil; 6.000–7.500/m²; box 35–50 mil); p. 6 (Fox 32.200–35.500; anúncios 31.000–39.000); p. 7 (Toyota 45–120 mil; FIPE 15–30 mil); Ceval p. 2–4 (240–295 mil; 85–130 mil; 90–150 mil; FIPE 30.600).
- **C0:** só a Toyota está na lista de Rafael entre estas empresas (`ativos-avaliados-valter.xlsx`, linha 10). A direção foi conferida por script nas 10 linhas: maior valor nos 9 bens que ficam, menor (igual) na Toyota.
- **Somas e quotas:**
  - Rádio: soma adotada R$ 2.912.000 e 16,50 % = R$ 480.480,00.
  - Ceval: R$ 385.000 e 1/6 = R$ 64.166,67.
  - 03, 09 e 10: R$ 0 em bens nas duas avaliações e no doc 00.
  - Contas dos pacotes conferidas: 1.515.100; 908.000; 1.011.000; 340.000; 243.800; 2.003.000; 446.244,50; 232.960; 38.438,40; 396.120–495.150; 6.361,71/m²; 182.000; 95.000.
- **2018:** Rádio 9.316.016 (com o terreno; 2.168.000 sem ele), Ceval 391.000, ST 2.443.100 − 1.515.100 = 928.000 e Associação 10.000 conferem com o inventário (arq. 9.1, 11.1, 7.1) e com o doc 27. O terreno entrou uma vez só em 2018 (na Rádio; a SEI o lista fora do total).
- **Terreno 3.751:** valor só na SEI (R$ 7.148.016, R3 da SEI, "não abordado por Rafael"). Não está no estudo da Rádio de Rafael.
- **Bens destas empresas em outras planilhas:** nenhum (sala, box, Fox, Toyota, LK, 2423 K) aparece no Mestre da SEI ou em outro inventário de 2025. As duplicidades possíveis são só a 320D (S4-2) e a Case (S4-3).
- **Referências cruzadas:**
  - Todas as perguntas citadas em `decisoes.json`, `parte6-candidatos.md` e `registro-data-base.md` (1–20) apontam para o assunto certo.
  - S4-1 → candidato nº 1 da SEI e S4-2 → candidato nº 4 da SEI conferem.
  - S4-DB1 → pergunta nº 7 da S3 (`rovilio-quotas.md`, seção 10) confere.
  - Os parágrafos citados dos inventários de 2025 conferem: ST 13–16 e 22; Turismo 14, 18–20; Associação 12–14, 19–21; Rádio 14–19.
- **Revisões anteriores:**
  - Os achados dos pacotes 1, 2 e 3 e da mestre foram incorporados.
  - Restam: M2 (concordância expressa na planilha), M1 (crédito "em aberto" na planilha), B2 ("sempre da SEI" no Leia-me) e B5 (formato inglês nos alertas).
  - O cruzamento S4-6 × lotes escolhidos por Rafael segue aberto, como pendência legítima.
