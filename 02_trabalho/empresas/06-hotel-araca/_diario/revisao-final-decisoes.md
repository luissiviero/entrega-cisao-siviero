# Revisão independente final — decisões aprovadas da empresa 06 Hotel Araçá (antes da ficha de fechamento)

Data: 28/10/2026. Revisor com contexto novo (P11). Objeto: `log-decisoes.md`, `decisoes.json`, `planilha-mestre-06.xlsx` (md5 `4e16d10033`, gerada por `_scripts/build_mestre_06.py`; `decisoes.json` md5 `f797da2909`), os pacotes `fase0-resumo-06.md`, `pacote-01-regras.md`, `pacote-02-excecoes.md` e `pacote-03-destinos.md`, os registros (`registro-data-base.md`, `perguntas-ao-grupo.md`, `parte6-candidatos.md`), os relatórios dos revisores anteriores e o diário.

Conferi contra as fontes, lidas direto por script: inventário 2025 (docx 06), inventário 2018 (doc 16 no Drive), PDF de Rafael (26 p.), `ativos-avaliados-valter.xlsx` e o arquivo 1 das respostas ao Dr. Pablo (p. 11). Só alterei este relatório. Scripts e extrações temporárias: `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\5fd9298d-a99a-4db0-a826-38f7fb4a324d\scratchpad\revisor-final\`.

**Contagem: 0 altas, 5 médias, 7 baixas.**

Nenhuma decisão de valor está mal aplicada, e nenhum número está errado. Os 22 valores adotados, as decisões e os motivos da mestre são iguais aos de `decisoes.json`, que por sua vez segue o log. A soma (R$ 75.870.000,00) e o efeito no 1/6 (+R$ 2.965.529,47) conferem com as fontes. Os achados tratam de decisões aprovadas que não chegaram à mestre, de referências desatualizadas e de pendências que ficaram só no texto dos pacotes.

---

## Médias

### M1 — Duas decisões aprovadas (E4 e D6) estão em `decisoes.json`, mas não foram aplicadas na planilha-mestre

**Evidência.**
- `decisoes.json` tem as seções `excecoes_sem_item.E4` (os R$ 360.000,00 não adotados, opção a) e `pendencias_componentes.D6` (1.500 m² de área construída sem destino, G3).
- `build_mestre_06.py` só lê `dec["decisoes"]` (linha 325), `dec["destinos_2018"]` (linha 575) e `dec["regras"]` (linha 672). As duas seções acima não são lidas em ponto nenhum do script.
- Na mestre:
  - item 01: a coluna "Destino 2018→2025" mostra "—" e o motivo (E1) não cita D6 nem G3. A nota diz só "a esclarecer";
  - aba Totais, linha 5: diz "Diferença de soma de Rafael (não identificada)", sem dizer que o Luís decidiu não adotá-la (E4).
- O header do log diz que as decisões "são aplicadas na planilha-mestre … a cada regeneração". Para E4 e D6 isso não acontece.

**Efeito.** Nenhum valor muda: E4 vale R$ 0,00 e D6 não tem valor. Mas a mestre é o registro aplicado. Quem montar a ficha só a partir dela perde uma decisão aprovada (E4: o Luís abriu mão de R$ 60.000,00 no 1/6 contra o C0) e uma pendência de existência dentro do item de R$ 66 milhões (D6).

**Correção.** O script passa a ler as duas seções:
- E4: na aba Totais, linha 5, nota "não adotada (E4, Pacote 2)";
- D6: no item 01, coluna Destino, o texto "Componente obra civil: 12.000 → 10.500 m² sem destino documentado (D6, G3)", e o mesmo na nota da obra civil em Item01_componentes.

Depois, regenerar a mestre e registrar o md5 no diário.

### M2 — O motivo aprovado de E2 manda as casas para a pergunta G8, que o Pacote 3 (também aprovado) tornou exclusiva da Fiat Toro

**Evidência.**
- `decisoes.json`, `regras.E2.texto`: "… destino na Parte 4 com os itens 22 e 23 de 2018; pergunta G8". O mesmo texto está na mestre, na linha "Duas residências … Rua da Divisa", coluna Motivo.
- O log (Pacote 2) diz "pergunta G8".
- O Pacote 3, aprovado depois, diz "G7 (refeita, junta os cinco lotes e as duas casas; a G8 fica só para a Fiat Toro)". `perguntas-ao-grupo.md` já segue isso: a G8 trata só da Toro, e a G7 pergunta pelas casas.

**Efeito.** É uma contradição entre dois registros aprovados, mas só de referência. Não muda valor nem decisão, por isso a classifico como média e não alta. Quem seguir a mestre acha na G8 só a Toro e pode concluir que as casas ficaram sem pergunta.

**Correção.**
- Em `regras.E2.texto`, trocar "pergunta G8" por "pergunta G7 (refeita no Pacote 3)".
- Acrescentar ao log, no Pacote 3, a nota "E2 passa a apontar para a G7".

A mudança só executa o item 4 aprovado no Pacote 3. Não é decisão nova, mas convém dizer isso ao Luís na ficha.

### M3 — Item 02 (lista de Rafael): o risco de já ter sido "dado aos sócios" em 2019 não tem pergunta que o resolva, e a decisão R1 não traz ressalva

**Evidência.**
- Arquivo 18, p. 3, bloco de 10.06.2019: "Lote Capão da Canoa … Lote na esquina próximo ao Hotel Araçá Ltda., com valor diferenciado aos demais lotes dados aos sócios", sem inscrição. O trecho é citado no Pacote 3, em D1 e na mestre.
- O candidato 7 da Parte 6 diz: "Se for um dos lotes do Hotel que ainda estão no inventário 2025 (por exemplo o item 02 …), o mesmo bem estaria entregue e escolhido duas vezes."
- A G7, porém, pergunta: "O 'Lote Capão da Canoa' do arquivo 18 (2019) é qual **deles**?". "Deles" são as cinco inscrições de 2018 sem par em 2025. A pergunta não deixa a resposta ser "o item 02" ou outro lote que ainda está no inventário 2025. Também não pergunta o "valor diferenciado" nem quais foram os "demais lotes dados aos sócios", como o revisor do Pacote 3 tinha sugerido (A1).
- A mestre aplica R1 ao item 02 ("bem identificado em 2018 … concordância expressa de Rafael"). A nota de identificação desse item não cita o arquivo 18 nem o candidato 7. O risco só aparece na linha do D1.

**Efeito.** O item 02 vale R$ 400.000,00 e está na lista de Rafael (vem para nós). Se já foi entregue a Válter em 2019 para acerto futuro, três coisas mudam:
1. não é da empresa em 31/12/2024;
2. a escolha de Rafael repete um bem já recebido;
3. a conta de Válter na Parte 4 muda.

Como está, nenhuma pergunta ao grupo resolve isso.

**Correção.**
- Ampliar a G7 com três perguntas: "O lote do arquivo 18 é um desses cinco ou um lote que ainda consta do inventário 2025 (por exemplo 360.0019.003.000, item 02)? Qual foi o 'valor diferenciado'? Quais foram os 'demais lotes dados aos sócios' e quando?"
- Na mestre, item 02, acrescentar a nota: "ver candidato 7 da Parte 6 (arquivo 18, 2019)".

### M4 — Regularização do hotel (R$ 2.000.000,00 de redução, INSS de R$ 1.601.357,90, "ônus de dissolução"): o Pacote 2 manda o tema para a Parte 4 e a Parte 6, mas ele não está em registro nenhum

**Evidência.**
- Inventário 2025, observações: "redução de aproximadamente 4,00% = R$ 2.000.000,00"; INSS "aproximadamente R$ 1.601.357,90, afora ainda os custos de: A.P.P.C.I., A.R.T., engenheiro…"; "não podendo serem considerados como líquidos os valores".
- Arquivo 1, p. 11: não há projeto, habite-se nem profissional contratado.
- Pacote 2, E1: "o tema vai para os passivos (Parte 4) e a Parte 6". Mas `parte6-candidatos.md` (candidatos 1 a 7) não o traz, e `perguntas-ao-grupo.md` não pergunta nada sobre isso. O tema só aparece na aba Observacoes e nas notas do item 01.

**Efeito.**
- A favor do Luís: nenhum dos lados deduz os custos hoje.
- Contra o Luís: o outro lado vai pedir a dedução no valor de R$ 66 milhões, que é custo de reposição de prédio regular.
- O inventário não deixa claro se o INSS está dentro dos R$ 2 milhões de redução ou se soma a eles. Se a Parte 4 lançar os dois, pode haver desconto em dobro (forma 5 da Parte 6: passivo de obra mais bem avaliado líquido).
- Ordem de grandeza: até R$ 3.601.357,90 na empresa, ou R$ 600.226,32 no 1/6.

**Correção.**
- Novo candidato 8 em `parte6-candidatos.md`: "regularização do item 01: redução de 4 % (R$ 2.000.000,00) e INSS (R$ 1.601.357,90); verificar se se sobrepõem e se algum passivo da Parte 4 os repete".
- Nova pergunta G13: "O INSS de R$ 1.601.357,90 está dentro da redução de R$ 2.000.000,00 ou soma a ela? Há passivo dessa regularização lançado no balanço contábil?"

### M5 — Material não lido (P9): a lista está só no texto do Pacote 3, sem registro de pendência nem aviso ao Luís

**Evidência.**
- O Pacote 3, em "Busca feita", diz que não foram lidos: PDFs escaneados fora do OCR (ACERTOS E ACORDOS, considerações de 2019, retiradas de 2018, relatórios mensais), 3 imagens e 4 .doc antigos.
- O revisor do Pacote 3 leu parte disso: a ata de 14/01/2019, o relatório do hotel de 2025, as retiradas de 2018 e dois dos .doc.
- Continuam sem leitura: os 15 relatórios mensais (jul/2023–out/2024), o contrato Buriti, a 3ª alteração da VMC, os .doc "2." e "C-", as imagens restantes e os `.gdoc`/`.gsheet`.
- O P9 diz: "Nenhum material fica de fora por não conseguir abri-lo. Primeiro tentar conectores … se ainda falhar, avisar para o Luís resolver."

**Efeito.** Os relatórios mensais são justamente onde podem aparecer a Fiat Toro, a manutenção das casas paga pelo Hotel (candidato 5) e o pagamento das placas solares (candidato 1). Os destinos D1 a D5 ficam "sem documento" sem que essas fontes tenham sido varridas.

**Correção.** Antes da ficha:
1. Ler como imagem, ou pelo conector, os relatórios mensais, com os termos Toro, Capão, Araçá, casa e manutenção.
2. Ler os dois .doc restantes pelo conector.
3. O que não se conseguir ler vai para a ficha como "pendência de acesso (P9)", com o nome de cada arquivo.

---

## Baixas

### B1 — D8 foi aprovado no Pacote 3, mas não está no log, no json nem em registro nenhum

- **Evidência.**
  - O log do Pacote 3 lista D1 a D7.
  - O D8 ("pode ser agregado a casa, EXCETO os lotes adquiridos adjacentes", convocação de 26/05/2017) existe só em `pacote-03-destinos.md`.
  - O Pacote 3 não o incluía em "O que peço para aprovar", e por isso o log é coerente com a lista de aprovação. Ainda assim, o enunciado desta revisão fala em "D1–D8".
- **Efeito.** A frase interessa à lista de Rafael, que escolheu lotes adjacentes ao hotel, e pode ser usada pelo outro lado. Ela some se ninguém reler o pacote.
- **Correção.** Uma linha no log, "D8: só registro", e uma menção na ficha.

### B2 — Fiat Toro (E5): a condição está invertida e "volta com valor" não tem fonte

- **Evidência.** Pacote 2, E5: "Se for da empresa em 31/12/2024, volta com valor; se não, vai para a Parte 5". O json diz "Parte 5 se não for da empresa".
- **Efeito.**
  - Pela definição do plano, a Parte 5 é a do bem não abordado. Uma Toro que é da empresa e não está no inventário é exatamente o caso da Parte 5. Uma Toro que não é da empresa só precisa de destino, se tiver saído depois de 2018.
  - Nenhum dos dois lados dá valor à Toro (PDF p. 25). Pelo P5, "volta com valor" não tem de onde sair sem que o grupo a avalie.
- **Correção.** Reescrever: "Se for da empresa, é bem não abordado (Parte 5) e precisa de valor do grupo. Se não for, registrar o titular." Na ficha, mostrar a Toro como bem sem valor possível hoje.

### B3 — Dólares do cofre: "Situação" ficou como "Não abordado", e o D7 classificou como destino de bem de 2018 (P4)

- **Evidência.**
  - Mestre, linha dos dólares: Situação "Não abordado" (a Legenda diz "nem o grupo nem Rafael avaliam"). "Existe em 2018?": "Sim, fora do inventário". Destino: "Pendente (Parte 4)".
  - O revisor do Pacote 3 (M4) pediu que se escolhesse uma das classificações. O D7 escolheu P4, mas a Situação não acompanhou.
- **Correção.** Situação "Só 2018 (fora do inventário)" ou uma nota que explique a dupla classificação. Na ficha, mostrar os dólares no bloco Parte 4.

### B4 — As mesmas casas aparecem em até três linhas

- **Evidência.** A linha "Duas residências … Rua da Divisa" (só Rafael) e as linhas D2 e D3 (lotes 360.0057.002 e .001, com casa) são, pela hipótese registrada (doc 25: "CASA DA ESQUINA"/"CASA AO LADO"), os mesmos bens.
- **Efeito.** Não há valor em dobro, porque nenhuma das três linhas tem valor. Mas as "10 linhas sem valor nesta etapa" do diário não são 10 bens diferentes.
- **Correção.** Na ficha, agrupar como "casas da Rua da Divisa = lotes D2/D3 (provável)" e contar os bens, não as linhas.

### B5 — R2: o motivo cita "G10, G11" para os três itens, mas cada pergunta cobre só parte deles

- **Evidência.**
  - A G11 trata só do item 17.
  - A G10 trata dos itens 21 e 22, mas só de quando e como foram adquiridos: não pergunta em nome de quem estão registrados.
  - O item 22 está matriculado em Osório (mat. 124.906), e nenhuma pergunta explica por quê.
- **Correção.**
  - Motivo por item: 17 → G11; 21 e 22 → G10.
  - Acrescentar à G10: "Estão registrados em nome da empresa? Por que a matrícula 124.906 está em Osório?"

### B6 — Lotes 15, 16 e 17 da quadra 37 (itens 21, 03 e 22): a "sobreposição com o item 01" está superestimada nos pacotes e não tem pergunta

- **Evidência.**
  - Os 6.900 m² do item 01 vêm de 12 matrículas que cobrem os lotes 01–12, 18–22 e a nesga: 4 × 450 + 13 × 360 + 420 = 6.900 m². Os lotes 15, 16 e 17 têm matrículas próprias (86.669, 64.782 e 124.906). Não há área contada duas vezes.
  - O Pacote 1 diz que o "valor de lote avulso pode se sobrepor ao item 01". O único risco real é a obra (piscinas, estacionamento) estar sobre esses lotes, o que pesa na regularização e na unificação. O arquivo 1, p. 11, diz que "a quadra de terrenos que abriga a edificação do hotel" será unificada.
  - Nenhuma pergunta trata disso: a G1 cobre só os lotes 13 e 14.
- **Correção.**
  - Na ficha, não chamar isso de dupla contagem de valor.
  - Acrescentar à G1: "Alguma parte do hotel (piscinas, salão, estacionamento) está sobre os lotes 13 a 17 da quadra 37?"
  - A dupla contagem de área que existe de fato é a da nesga (G2): até R$ 2.100.000,00 no valor de Rafael adotado, ou R$ 350.000,00 no 1/6.

### B7 — O erro da planilha Google do Luís (lote 28/q38 com valor 0 em 2018) está só no resumo da Fase 0

- **Evidência.**
  - `fase0-resumo-06.md`, item 5: o erro está na GERAL (linhas 70, 80 e 99), nas abas Hotel Araçá das duas planilhas e num bloco da aba 7.Rádio.
  - A "Diferença" aparece como R$ 1.500.000,00; o acréscimo real é R$ 1.150.000,00.
  - A mestre registra "Observação do Luís: nenhuma". O erro não chegou a nenhum registro nem à Parte 7.
- **Correção.** Uma linha na ficha: "a sua planilha 2_0 tem um erro de R$ 200.000,00 no 2018 do Hotel (lote 28/q38); corrigir antes de usá-la".

---

## O que a ficha de fechamento precisa dizer para não induzir o Luís a erro

1. **Números.**
   - Grupo: R$ 58.076.823,20, ou R$ 9.679.470,53 no 1/6.
   - Valores adotados: R$ 75.870.000,00, ou R$ 12.645.000,00 no 1/6 (+R$ 2.965.529,47).
   - Lista de Rafael: R$ 1.880.000,00, igual nos dois lados. Crédito pelos valores adotados: R$ 10.765.000,00. Rafael escreve R$ 12.705.000,00 e R$ 10.825.000,00 porque inclui os R$ 360.000,00 não identificados, que o Luís decidiu não adotar (E4; abriu mão de R$ 60.000,00 contra o C0).
2. **92 % da diferença (R$ 16.393.176,80 de R$ 17.793.176,80) depende só de E1**, e a base de Rafael é fraca:
   - os 66 milhões são o piso do custo de reposição (terreno + prédio), sem móveis;
   - o método que o próprio Rafael chama de "mais realista" (renda) dá R$ 41,0 a 50,6 milhões, faixa que contém o valor do grupo;
   - a faixa de R$ 55 a 85 milhões não tem conta no PDF;
   - as fontes são de 2025.

   Se E1 cair, o 1/6 volta para R$ 9.912.803,87 (só com E3).
3. **Valores brutos.** Nada está líquido de regularização, INSS ou "ônus de dissolução" (até R$ 3.601.357,90; R$ 600.226,32 no 1/6). A Parte 4 (passivos e destinos) não foi feita; o método depende da explicação do Luís.
4. **Aprovou-se o valor, não a titularidade.**
   - Dependem de resposta do grupo os itens 17 (G11), 20 (G6), 21 e 22 (G10): R$ 3.550.000,00, ou R$ 591.666,67 no 1/6.
   - O item 02 (lista de Rafael) tem o risco do arquivo 18 (candidato 7; M3).
   - O grupo afirma que "todos os imóveis" estão em nome da empresa (arquivo 1, p. 11), mas não apresentou matrículas.
5. **Fora do valor, sem como somar hoje (P5).**
   - Possíveis acréscimos à empresa: as duas casas (Rafael: R$ 2,4 a 3,5 milhões pelo par, R$ 400.000,00 a 583.333,33 no 1/6), provavelmente os lotes D2/D3; os lotes D1, D4 e D5; os dólares do cofre (U$ 37.698,00 em 2017); a Fiat Toro; as placas solares pagas pela SEI; os lotes 13 e 14 da q37.
   - Possível redução: área do terreno (nesga, G2) e da obra (D6, G3).
6. **Pendências abertas.**
   - Perguntas G1 a G12 aprovadas e **não enviadas** (mais as ampliações propostas acima).
   - Candidatos 1 a 7 da Parte 6, abertos, e o candidato 8 se for aceito.
   - D1 a D7 pendentes, D3 com três leituras sem escolha; D8 só registro.
   - Registros de data-base 1 a 3, mantidos sem ajuste (E6).
   - Material não lido (M5).
   - Os registros desta empresa ainda não foram unidos aos da SEI (P13).
7. **Não dizer** que os lotes 15 a 17 da q37 contam em dobro (B6), nem contar as linhas sem valor como bens distintos (B4).

---

## O que conferi e achei certo

- **Decisões.** Os 22 itens de 2025 têm na mestre o mesmo valor adotado, a mesma decisão e o mesmo motivo de `decisoes.json`, conferidos por script. Os destinos 2018-20, 22, 23, 24 e 25 da mestre são iguais aos do json. As regras sem valor (R3, E2, E5, D7) estão aplicadas às linhas certas, e o filtro por descrição funciona: a linha dos dólares recebe D7, não R3. Não há decisão aplicada sem aprovação. Aprovadas e não aplicadas, só E4 e D6 (M1).
- **Valores contra as fontes.**
  - Inventário 2025: 22 itens, total R$ 58.076.823,20. O item 01 confere parcela a parcela: terreno 6.900.000; móveis 1.512.000 e 2.592.000; obra 26.632.368; equipamentos 6.400.000; vantagem 15 % sobre 37.136.368 = 5.570.455,20.
  - Rafael, balanço executado, p. 17–20: os itens somam R$ 75.870.000,00; o total declarado é R$ 76.230.000,00. Os 66 milhões (p. 7) = 6.900 m² × 5.000 + 10.500 m² × 3.000. A repetição do valor do grupo na p. 6 (49.606.823,20) e na p. 13 (1.000.000,00) confere, e as faixas das p. 7, 8, 10, 11, 13 e 14 também.
  - Lista de Rafael: itens 02, 11, 12, 18 e 19, R$ 1.880.000,00, igual à p. 21–22.
  - Inventário 2018: os 19 itens com valor somam R$ 49.684.243,20; os itens 20 a 25 têm nome de sócio no lugar do valor, e o 23 traz a nota da troca.
- **Somas.**
  - R1 = 6.320.000 (17 itens).
  - R2 = 1.150.000 (3 itens).
  - E1 + E3 = 68.400.000.
  - Soma dos adotados: 75.870.000,00. Diferença: +17.793.176,80, ou +2.965.529,47 no 1/6.
  - E1: +2.732.196,13 no 1/6. E3: +233.333,33.
  - Rafael corrigido: 12.645.000 e 10.765.000.
  - Renda pelas premissas de Rafael: 3.587.220 por ano, R$ 34.164.000,00 e R$ 42.202.588,24.
- **C0.** A direção está certa em todos os itens. Os dois que divergem (01 e 20) ficam na empresa e recebem o maior valor. Os cinco da lista têm valor igual dos dois lados, e os itens 03–10 não têm valor maior disponível (P5).
- **Pares 2018 × 2025.** Batem por inscrição, código e matrícula. O item 17 de 2025 tem a mesma inscrição do lote "IRINEU" de 2018 (item 21). Os 5 lotes de 2018 sem par estão completos.
- **Dupla contagem de valor entre itens adotados.** Não há. O item 01 é formado por matrículas distintas das dos itens 03, 21 e 22, e a área de 6.900 m² fecha. A única dupla contagem possível fica dentro do item 01 (nesga, G2), registrada.
- **Registros.**
  - `registro-data-base.md` (3 registros, E6 aplicado) confere com o Pacote 2.
  - O texto das perguntas G7, G8 e G12 é igual ao do Pacote 3.
  - Os candidatos 6 e 7 da Parte 6 foram acrescentados como aprovados.
- **Premissas e brief.** Não achei decisão que contrarie o plano ou o brief.
  - E1 foi tratado como exceção, com as duas opções, como pede o brief.
  - E4 se afasta do C0 com registro e por escolha expressa do Luís.
  - Nenhum ajuste de data-base foi feito em silêncio.
  - O destino 2018 → 2025 ficou como pendência, não como duplicidade (P4).
  - Os arquivos escritos estão só na pasta da empresa (P13).
