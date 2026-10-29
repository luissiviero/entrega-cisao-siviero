# Revisão independente da Fase 0 — empresa 06 Hotel Araçá

Data: 28/10/2026. Revisor em contexto novo. Li direto das fontes originais (docx 2018, 2025 atual e antigo; PDF de Rafael, 26 páginas, texto e imagem das p. 17, 20 e 21; Diretrizes p. 6–7; `ativos-avaliados-valter.xlsx`; as duas planilhas Google, com valores e fórmulas; arquivo 1 p. 10–11; arquivo 12 renderizado como imagem; doc 27 de 2018 e doc 00 de 2025). Não usei os `.txt` de `_extracoes` como fonte. Scripts do revisor: `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\5fd9298d-a99a-4db0-a826-38f7fb4a324d\scratchpad\revisor-fase0\`.

**Resultado geral:** não achei número extraído errado nem bem perdido na extração (0 achados altos). Os achados são riscos no script, uma aba de saída com valores truncados e omissões no resumo do pacote (`pacotes/fase0-resumo-06.md`). Vários pontos já estão no diário, mas não chegaram ao resumo que vai ao Luís.

Contagem: **alta 0 · média 10 · baixa 8**.

---

## Achados

### M1 — média — Aba `Rafael_textos` com valores truncados e páginas faltando
- **Evidência:** `fase0-06.xlsx`, aba `Rafael_textos`. A regex `[^.]*R\$ ?[\d\.,]+[^.]*\.` corta a frase no ponto de milhar. Exemplos: p. 7 "calculando o terreno a R$ 5." e a linha seguinte "500 m² a R$ 3.000/m²"; p. 10 "varia entre R$ 380." / "000 e R$ 550.000"; p. 13 "varia entre R$ 2." (é R$ 2.400.000); p. 14 "Entre R$ 1.810.000 e R$ 2." (é R$ 2.750.000) e "Entre R$ 600.000 e R$ 750."; p. 15 "R$ 380."; p. 21 "de R$ 10." (é R$ 10.825.000,00); p. 22 "R$ 1.880."; p. 24 "entre R$ 2.400.000 e R$ 3."; p. 25 várias. A aba também não varre as p. 5–6 e 16–19.
- **Por que importa:** é o único registro, no xlsx, das faixas de mercado de Rafael (hotel 55–85 mi; físico 66–102,45 mi; renda 41–50,6 mi; lotes q37/38 380–550 mil, esquina 550–600 mil; item 20 2,4–3,5 mi, terra 1,81–2,75 mi, casa 600–750 mil; casas da Rua da Divisa 2,4–3,5 mi). Quem ler a aba vê números errados.
- **Corrigir:** separar trechos por linha ou por uma regex que não quebre em `\d\.\d`. Criar uma tabela estruturada das faixas: item, mínimo, máximo, base (m², R$/m²) e página. Os valores corretos estão acima e no diário.

### M2 — média — `check()` dá CONFERE quando os dois lados são `None`, e DIVERGE não para o script
- **Evidência:** `fase0_extracao_06.py`, `check()`: `abs((got or 0) - (expected or 0)) < 0.005`. Se a regex do total falhar (`total = None`) e a soma der 0, ou se `first()` devolver `None` nos dois lados, o resultado sai "CONFERE". A docstring diz que "The script stops (assert) when a parsed total does not match the document", mas `check()` só grava o resultado. Os DIVERGE (Rafael 360 mil; GERAL C70 e C80) não interrompem nada.
- **Corrigir:** `assert got is not None and expected is not None` dentro de `check()`. No fim, comparar a lista de DIVERGE com a lista esperada e documentada, e falhar se surgir uma nova. Ajustar a docstring.

### M3 — média — Pareamento por ordem no balanço executado sem conferência independente
- **Evidência:** `lot_values()` pareia inscrições e valores por índice e confere só as contagens (21). Na transcrição, um deslocamento seria pego pela conferência item a item contra o inventário. No executado não: 20 dos 21 lotes têm valores repetidos (360/390/350/380 mil), e um deslocamento que movesse os R$ 2.400.000,00 do item 20 para o 19 ou o 21 passaria sem alarme. Hoje o pareamento está **certo**: conferi na imagem e no texto das p. 17–20, e o item 13 (valor na p. 18 antes da inscrição na p. 19) está bem pareado. Também são frágeis: (a) a exclusão do "valor total global de R$ 750.000,00" depende de uma janela de 40 caracteres; (b) `49606823.20`, `66000000.00` e as páginas 2 e 17 do item 01 estão fixos no código, sem leitura.
- **Corrigir:** ancorar cada valor ao rótulo do item que o precede ("20.", "21. e 22.", "02" etc.) ou à inscrição anterior mais próxima, e conferir que o valor diferente do grupo cai no item 20. Ler o valor e a página do item 01 do texto em vez de fixá-los.

### M4 — média — Erro do lote 28/q38 nas planilhas Google vai além da GERAL C70, e o resumo só cita a C70
- **Evidência:** planilha 2_0, aba `6.Hotel Araçá`, linha 31 (item 19, lote 28/q38): C31 = 0 e D31 = 0 (digitados; as linhas vizinhas usam `=G..`). H31 = 350.000. Pelo inventário de 2018, esse lote é o item 18, inscrição 360.0038.028.000, R$ 200.000,00. O erro passa para: C2 = 49.484.243,20 (certo: 49.684.243,20); C12 = 4.210.000 (certo: 4.410.000); D2 "2018 atualizado" = 56.576.823,20 (com D31 = 350.000 seria 56.926.823,20); H2 "Diferença" = 1.500.000 (certo: 1.150.000 = item 17 de 390 mil, que era de Irineu em 2018, mais os itens 21 e 22 de 380 mil cada). O mesmo erro aparece na GERAL, linhas 70/80/99 (C, D e G), na aba `6.Hotel Araçá` da planilha "empreendimentos e Rafael" (C/D/I) e num **bloco duplicado do Hotel Araçá dentro da aba `7.Rádio`** dessa planilha (linhas 1–33, antes do bloco da Rádio na linha 35). Fora isso, F3 da aba `6.Hotel Araçá` = `SUM(F4:F18)` = 63.630.460 soma o item 01, o subtotal e cinco lotes: fórmula errada da fonte.
- **Corrigir:** no resumo, item 3.4, citar também D2/D70 e H2/G70 e dar os valores certos (1.150.000 de acréscimo real entre 2018 e 2025). Anotar o bloco duplicado em `7.Rádio` e o F3.

### M5 — média — Itens 04–10 de 2025: texto diz "Quadra nº 37", inscrição diz quadra 0071
- **Evidência:** inventário 2025 atual e antigo, tabela 3, itens 04–10: "lote urbano nº 23 a 29 – Quadra nº 37", com inscrições 360.0071.023.000 a 360.0071.029.000. Nos outros itens, o segundo grupo da inscrição é a quadra (0019, 0037, 0038). Em 2018 esses itens não têm lote nem quadra no texto, só a inscrição 0071. Os itens 08/09 ("lotes 27/28, q37") e 18/19 ("lotes 27/28, q38") repetem os números de lote. Rafael copia "Quadra nº 37" (p. 3, 17–18) e aplica a faixa "quadras 37 e 38".
- **Por que importa:** o pareamento por lote e quadra na planilha-mestre pode trocar itens. A quadra real dos itens 04–10 pode ser outra (fundos do hotel).
- **Corrigir:** na planilha-mestre, parear só por inscrição e matrícula. Registrar a inconsistência como pergunta: confirmar a quadra pela matrícula (26.373–26.379).

### M6 — média — Observações do inventário 2025 (redução de R$ 2 milhões, INSS, "valor não é líquido") não foram extraídas
- **Evidência:** inventário 2025 atual, parágrafos depois da tabela 5: INSS da obra "aproximadamente R$ 1.601.357,90", além de APPCI, ART, engenheiro etc.; "O valor patrimonial lançado no ÍTEM 01 DO INVENTÁRIO, sofrerá redução pela regularização do imóvel, o que se estima numa redução de aproximadamente 4,00% = R$ 2.000.000,00"; "não podendo serem considerados como líquidos os valores". O total de 58.076.823,20 não desconta esses valores. O script lê só as tabelas. O resumo não cita os R$ 2 milhões, e o diário cita só o INSS (via arquivo 1).
- **Corrigir:** extrair os parágrafos de observação das três versões para uma aba própria. Registrar no resumo que o próprio grupo estima redução de cerca de R$ 2.000.000,00 no item 01 e INSS de cerca de R$ 1.601.357,90, sem lançá-los no total. Matéria de passivo e contingência.

### M7 — média — Mudanças de parâmetro do item 01 entre 2018 e 2025 não estão no resumo nem na aba de componentes
- **Evidência:** item 01 no docx de 2018 e no de 2025. Área construída **12.000,00 m² (2018) → 10.500,00 m² (2025)**. CUB ago/2017 R$ 2.154,09 → CUB **dez/2022** R$ 3.170,52, dois anos antes da data-base. Terreno R$ 550,72/m² → R$ 1.000,00/m². Móveis R$ 50.000 e 35.000 por apartamento → R$ 42.000 e 30.000. Vantagem 30 % (2018) → 25 % (2025 antigo) → 15 % (2025 atual). O resumo só cita 25 → 15 %. A aba `Item01_componentes` guarda os valores, mas não área, preço unitário nem percentual de depreciação. Na fonte de 2018 há ainda um erro de digitação: "R$ 31.903,264,00" (= 31.903.264,00, conferido pela soma).
- **Corrigir:** acrescentar à aba as colunas área, R$/unidade, base (CUB e mês) e % de depreciação. No resumo, registrar a queda de área construída e o CUB defasado, porque o 2018 é o benchmark.

### M8 — média — Nota de troca do lote de Élcio (2018) ausente do resumo, e ela pesa na hipótese das "duas residências"
- **Evidência:** inventário 2018, tabela 6, item 23 (inscrição 360.0057.001.000, "CASA/ESQUINA"): "TROCOU ESTE LOTE/CASA POR UM LOTE NO LOTEAMENTO ADITIIVO MEDIANEIRA, EM VERANÓPOLIS", com valor "(Avaliar Lote e Casa) … ELCIO SIVIERO". O total de 2018 traz "(+ o valor da casa e lote – Élcio Siviero)". O item 22 (Raul, 360.0057.002.000) é "CASA". O resumo, item 3.3, levanta a hipótese de que as casas de Rafael (p. 24) seriam os lotes de Raul e Élcio, mas não diz que a de Élcio foi trocada por lote da SEI. O diário registra.
- **Corrigir:** incluir a nota no resumo e na lista de perguntas: quem é hoje o titular das inscrições 360.0057.001/002; se a troca passou pela SEI (Aditivo Medianeira); se a casa de Élcio ficou fora do total de 2018.

### M9 — média — Lotes 13 e 14 da quadra 37 não aparecem em nenhum item, e o resumo não diz
- **Evidência:** item 01 (2025) cobre os lotes 01–12, 18–22 e a nesga. Os itens 03, 21 e 22 são os lotes 16, 15 e 17. Os lotes 13 e 14 da mesma quadra do hotel não aparecem em nenhuma das três versões nem no PDF de Rafael. O diário registra; o resumo não.
- **Corrigir:** levar ao resumo como pergunta de existência e titularidade, porque são lotes encravados no quarteirão a unificar.

### M10 — média — Coluna "Estudo Rafael" da planilha Google traz só os lotes escolhidos
- **Evidência:** "Comparativo … empreendimentos e Rafael", aba `6.Hotel Araçá`: H2 = `SUM(H3:H34)` = 1.880.000. Só H14, H23, H24, H30 e H31 estão preenchidas (itens 02, 11, 12, 18, 19). Não há os 66.000.000 do item 01, os 2.400.000 do item 20 nem o total de 76.230.000. O rótulo leva a ler 1,88 milhão como "valor de Rafael". O diário registra; o resumo não.
- **Corrigir:** dizer no resumo, item 3.4, e na aba `Planilhas_Google` que a coluna é a seleção, não a avaliação.

### B1 — baixa — Defeitos de forma no PDF de Rafael
- **Evidência:** a p. 16 tem o título "GRUPO SIVIERO — SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA." num documento do Hotel Araçá. A p. 21 traz "1/6 do patrimônio total 3,66 bens imóveis". A concordância da p. 16 cobre "02., à 19."; os itens 21 e 22 ficam fora do texto, embora estejam valorados em R$ 380.000,00 (p. 15 e 20). O item 20 é avaliado com casa de 160 m², contra 158,18 m² no inventário (p. 13–14).
- **Corrigir:** anotar no pacote de exceções.

### B2 — baixa — Rafael repete os valores do grupo na própria seção de avaliação
- **Evidência:** na seção "AVALIAÇÃO PATRIMONIAL EXECUTADA" (p. 5–6), o item 01 aparece com R$ 49.606.823,20 (p. 6); o item 20, com R$ 1.000.000,00 (p. 13). Os 66 milhões (p. 17) e os 2,4 milhões (p. 19) só aparecem no "balanço executado". O resumo cita "p. 7 e 17" e "13–15 e 19".
- **Corrigir:** dizer que o valor adotado por Rafael é o do balanço executado e que as p. 6 e 13 repetem o valor do grupo, para evitar dúvida sobre qual é o "valor de Rafael".

### B3 — baixa — O resumo não dá os números de Rafael corrigidos
- **Evidência:** com a soma correta (75.870.000,00), 1/6 = 12.645.000,00 e o crédito = 12.645.000,00 − 1.880.000,00 = 10.765.000,00. Rafael escreve 12.705.000,00 e 10.825.000,00 (p. 21).
- **Corrigir:** acrescentar ao resumo, item 2.

### B4 — baixa — Forma da aba `Planilhas_Google`
- **Evidência:** a coluna "linha" fica vazia nas linhas das abas `6.Hotel Araçá` (não dá para rastrear a célula). As linhas de subtotal "Lotes" (12 na aba, 80 na GERAL) não entram, porque a coluna A está vazia; a C80 é conferida à parte. A coluna E ("dezembro 2024", vazia ou 0) foi ignorada. A coluna F não tem cabeçalho na fonte; o rótulo "col. F (antigo)" é uma inferência (confere: F2 = 61.790.460 = inventário antigo). A GERAL foi lida nas linhas 70–102; a 103 é texto de observações e a 104 está vazia: nada se perdeu.
- **Corrigir:** gravar o número da linha e incluir os subtotais.

### B5 — baixa — Forma da aba `Itens`
- **Evidência:** o item 02 perde "antiga matrícula 37.549". As inscrições dos itens 21 e 22 têm um grupo ".000" a mais (360.0037.015.000.000), e os códigos 39394 e 39396 vêm sem o zero à esquerda. O item 22 está registrado na Comarca de **Osório**, matrícula 124.906; todos os outros estão em Capão da Canoa. Os itens 18 e 19 têm a mesma matrícula, 73.482, nas três versões.
- **Corrigir:** normalizar inscrição e código antes de qualquer cruzamento; manter a matrícula antiga numa coluna.

### B6 — baixa — Resumo 2018 (doc 27) e doc 00 de 2025 só recebem hash, sem conferência
- **Evidência:** conferi à mão. O doc 27 (2018) traz "HOTEL ARAÇA LTDA. R$ 49.684.243,20", igual ao inventário. O doc 00 (2025) traz R$ 58.076.823,20 e R$ 9.679.470,53 por sócio, iguais ao inventário.
- **Corrigir:** acrescentar as duas conferências ao script.

### B7 — baixa — Poço: o grupo diz que não é usado, e o fato não está no resumo
- **Evidência:** arquivo 1, p. 11, item 6: "não está sendo utilizado para abastecimento e consumo". Arquivo 12 (imagem): Portaria DRH nº 1398/2012, de 06/09/2012, poço tipo ponteira na Rua 2, nº 330; usos "sistema de prevenção de incêndio, jardinagem e piscina de uso coletivo"; 20 m³/dia; "caráter precário". O poço entra, sem valor próprio, nos "equipamentos diversos" (R$ 8.000.000 − 20 % = R$ 6.400.000). O diário transcreve isso corretamente.
- **Corrigir:** uma linha no resumo.

### B8 — baixa — Divergências menores entre as versões
- **Evidência:** o endereço é "Rua Dois, nº 333" em 2018, "Rua Miguel Main Felo, nº 333" em 2025 e "rua 2, nº 330" na outorga. O "novo quadro societário não real" é Élcio 98 % / Joel 2 % em 2018 e Élcio 95 % / Joel 5 % em 2025. A tabela de valor por sócio de 2025 diz "16,66 %", mas calcula 1/6 exato (9.679.470,53 × 6 = 58.076.823,18).
- **Corrigir:** só registrar; não mudam números.

---

## O que conferi e achei correto

- **Itens, nas três versões, contra o docx:** número, valor, matrícula, inscrição, código, área, lote e quadra de todos os itens; sócio nos itens 20–25 de 2018 (Válter, Irineu, Raul, Élcio, Adélia F., Pedro). Contagens: 25 em 2018 (19 com valor), 22 no 2025 antigo, 22 no 2025 atual. Os docx não têm controle de alterações, texto riscado nem texto oculto.
- **Totais:** 2018 = 49.684.243,20 (item 01 45.274.243,20 + lotes 4.410.000); 2025 antigo = 61.790.460,00; 2025 atual = 58.076.823,20 (item 01 49.606.823,20 + lotes 8.470.000). Os três batem com o TOTAL GERAL. A tabela por sócio de 2025 (9.679.470,53) é 1/6 do total.
- **Componentes do item 01, refeitos à mão:**
  - 2018: 3,8 mi + 1,8 mi + 3,024 mi + 20.679.264 + 6,4 mi + 9.570.979,20 (30 % de 31.903.264) = 45.274.243,20.
  - 2025 atual: 6,9 mi + 1.512.000 + 2.592.000 + 26.632.368 (10.500 × 3.170,52 = 33.290.460, menos 20 %) + 6,4 mi + 5.570.455,20 (15 % de 37.136.368) = 49.606.823,20.
  - 2025 antigo: com 25 % = 9.284.092,00 → 53.320.460,00.
  - A área de 6.900 m² = 4 × 450 + 13 × 360 + 420.
  - As duas versões de 2025 só diferem no percentual de vantagem e nos totais.
- **PDF de Rafael, todas as 26 páginas:**
  - A transcrição (p. 1–5) é igual ao inventário 2025, item a item.
  - No balanço executado (p. 16–20), o item 01 vale 66.000.000,00 (p. 17) e o item 20 vale 2.400.000,00 (p. 19); os outros 20 lotes têm o valor do grupo. A soma dá 75.870.000,00, contra 76.230.000,0 declarados (p. 20, visto na imagem): diferença de 360.000,00.
  - 1/6 de 76.230.000 = 12.705.000; o crédito de 10.825.000 confere com o total errado.
  - A conta dos 66 milhões fecha: 6.900 × 5.000 + 10.500 × 3.000. O máximo físico também: 6.900 × 8.000 + 10.500 × 4.500 = 102.450.000.
  - Bens citados por Rafael fora do inventário: as duas residências da Rua da Divisa com a Rua Três e a Fiat Toro (p. 24–25). Nenhum outro.
  - O pareamento item ↔ valor do script está certo hoje (ver M3).
- **Lista de Rafael:** 5 linhas da empresa 06 (itens 12, 11, 02, 18, 19; 390 + 390 + 400 + 350 + 350 = 1.880.000). Lote, quadra, matrícula, inscrição e código batem com o PDF (p. 22) e com as Diretrizes (p. 6–7). A coluna "Página" da lista se refere às Diretrizes.
- **Planilhas Google:**
  - Os valores transcritos na aba `Planilhas_Google` batem com as células da fonte (2_0 `6.Hotel Araçá` linhas 2–34; GERAL 70–102; "empreendimentos e Rafael" `6.Hotel Araçá` linhas 2–34) e com o mapeamento das colunas.
  - Os valores de out/2025 e os componentes do item 01 de 2018 e 2025 conferem com os inventários.
  - A única divergência contra o inventário é a do lote 28/q38 em 2018 (ver M4).
  - O pareamento implícito 2018 ↔ 2025 da planilha está certo nos demais itens: item 17 de 2025 = Irineu em 2018, com 0; itens 21 e 22 sem par em 2018.
- **Respostas do grupo:** a p. 10 do arquivo 1 reproduz a composição do item 01 e os valores dos itens 02–22 iguais ao inventário. A p. 11 e o arquivo 12 estão transcritos corretamente no diário.
- **Resumo do pacote:** totais, diferenças (16.393.176,80 + 1.400.000 + 360.000 = 18.153.176,80), móveis e equipamentos (10.504.000) e o par Irineu → item 17 conferem.
