# Revisão independente — planilha-mestre da empresa 06 (Hotel Araçá)

Revisor de contexto novo, 28/10/2026. Objeto: `_scripts/build_mestre_06.py` (com os parsers de `fase0_extracao_06.py`), `planilha-mestre-06.xlsx` (abas Mestre, Item01_componentes, Totais, Pares_2018_2025, Observacoes) e `decisoes.json`. Li as fontes direto: os inventários 2025 atual e antigo, o inventário 2018 (Drive, doc 16), o PDF de Rafael (26 p.), `ativos-avaliados-valter.xlsx`, o arquivo 1 das respostas (p. 10–11), o arquivo 17 (lotes vendidos), o inventário 11 da SEI, as duas planilhas Google, `observacoes-planilhas-google.xlsx` e, para conferir os bens de 2018 sem par, os docs 1 a 6, 20 e 25 da pasta `inventário-2018`. Não alterei nenhum arquivo além deste relatório. Meus scripts estão em `…\scratchpad\revisor-mestre\`.

Referências de célula da aba Mestre: linha 2 = item 01 de 2025; linhas 3–23 = itens 02–22 de 2025; linhas 24–28 = itens 20, 22, 23, 24 e 25 de 2018 sem par; linhas 29–30 = bens citados só por Rafael. Colunas: F = Ref. 2018, L = Ref. Rafael, N = Valor Rafael, Q = Situação, R = Lista de Rafael, S = Direção C0, W = Destino, X = Registro de data-base, Y = Observação do Luís, AA = Notas.

**Resultado geral:** os números da aba Mestre conferem com as fontes, item a item. Os pares 2018 × 2025 estão certos. Os totais fecham. Achei 1 bem fora da planilha (alto), 8 riscos ou omissões relevantes (médios) e 13 problemas de forma (baixos). Nenhum par errado e nenhum valor de grupo, de Rafael ou de 2018 transcrito errado.

---

## Achados

### A1 — alta — Placas solares do hotel (instaladas em 2024, pagas pela SEI) não estão no inventário nem na planilha-mestre
- **Evidência:**
  - Inventário 11 (SEI) 2025, observação 14: "SOLAR SERRA PLACAS SOLARES – MÁRIO BORDIGNON – A empresa deverá pagar e fazer um acerto pelos serviços prestados e materiais utilizados na instalação e colocação das placas solares no HOTEL ARAÇÁ LTDA. …", com "R$ 150.000,00 … representados por 01 LOTE URBANO". Está também na GERAL da planilha 2_0, linha 887, e na planilha "empreendimentos e Rafael", linha 735.
  - Arquivo 17 (lotes vendidos), linha 60: lote 02/q07 do Aditivo Medianeira vendido à Solar Serra Ltda. em 24/04/2024, "como forma de pagamento placas solares Hotel Araçá Ltda.".
  - Inventário 06: nenhum texto fala em placas ou energia solar. Os "Móveis e Equipamentos Diversos" valem R$ 8.000.000,00 antes da depreciação em 2018 e também em 2025. As placas, instaladas antes da data-base, não entraram nesse valor.
  - A planilha-mestre não tem linha para elas. O diário as trata só como candidato à Parte 6.
- **Por que é alto:** é um bem físico do hotel, anterior a 31.12.2024, que fica fora da lista de bens, e não só uma questão de duplicidade entre empresas. O valor é desconhecido: no mínimo um lote de R$ 150.000,00, mais o "valor pago".
- **Corrigir:** criar uma linha no bloco "Hotel" (ou "Citados em outras fontes"), com Situação "Não abordado" (nem o grupo nem Rafael avaliam). Na referência, citar o inventário 11, obs. 14, e o arquivo 17, linha 60. Marcar Parte 6 na nota (despesa do Hotel paga pela SEI). Levar ao grupo a pergunta sobre o valor total da instalação e quem a registrou.

### M1 — média — A diferença de R$ 360.000 de Rafael leva o rótulo "sem bem", mas é exatamente o valor de um lote
- **Evidência:** aba Totais, A5 "Diferença de soma de Rafael (sem bem)" = 360.000. No balanço executado (p. 17–20), 360.000 é o valor de cada um dos itens 03 a 10 ("Fundos do Hotel"). É também 360 m² × R$ 1.000/m², o preço do grupo. Conta: 76.230.000 − 66.000.000 − 2.400.000 = 7.830.000 para os 20 lotes restantes, contra os 7.470.000 listados, ou seja, um lote de 360 mil a mais. Os itens 03–10 não estão repetidos no texto das p. 17–20 (conferi a sequência 02, 03, …, 19, 20, 21, 22). Não há planilha de Rafael para esta empresa em `Material Rafael/Estudo 2/versão-excel` que mostre a conta. A transcrição do balanço do grupo (p. 5) fecha certo em 58.076.823,20. Fato relacionado: os lotes 13 e 14 da quadra 37 não aparecem em nenhum item (ver M5).
- **Por que importa:** "sem bem" é uma conclusão, não um fato. Se Rafael contou um lote de 360 mil duas vezes, o erro é de soma. Se contou um lote que não está listado, há um bem por identificar. Nos dois casos o 1/6 (12.705.000) e o crédito (10.825.000) que ele declara na p. 21 ficam R$ 60.000 acima do que a soma dos itens dá (12.645.000 e 10.765.000).
- **Corrigir:** mudar o rótulo para "Diferença de soma de Rafael (não identificada)". Na nota, dizer que ela é igual ao valor de um lote dos itens 03–10. Levar à lista de perguntas a Rafael ou ao grupo.

### M2 — média — Item01_componentes: o "0,00" atribuído a Rafael nos móveis, nos equipamentos e na vantagem é uma inferência, mas aparece como valor
- **Evidência:** aba Item01_componentes, E3–E5 e E7 = 0, F3–F7 = −1.512.000 / −2.592.000 / −6.400.000 / −5.570.455,20. PDF p. 7: "Valor Mínimo Físico: Avaliado em R$ 66.000.000, calculando o terreno a R$ 5.000/m² e a depreciação técnica da estrutura construída de 10.500 m² a R$ 3.000/m² pelo CUB-RS modificado". A conta 6.900 × 5.000 + 10.500 × 3.000 = 66.000.000 fecha. O "Valor Máximo Físico" segue a mesma estrutura: 6.900 × 8.000 + 10.500 × 4.500 = 102.450.000. A decomposição em terreno e construção, portanto, se sustenta no texto.
- **Outra leitura possível:**
  1. No balanço executado (p. 16–17), Rafael copia a descrição inteira do grupo, com móveis, equipamentos e "15 % da obra feita", e põe R$ 66.000.000 como valor do item todo. Ele substitui o item e não diz que os móveis valem zero. Ele simplesmente não os avalia em separado.
  2. Os 66 milhões estão dentro da faixa de "valor médio de mercado" de 55 a 85 milhões. Essa faixa, segundo o próprio Rafael, "integra a avaliação física … com a capacidade de geração de receita hoteleira", ou seja, a de um hotel "totalmente equipado". Também está acima da faixa por renda (41,0–50,6 milhões), que por definição inclui o hotel equipado.
  3. Os R$ 3.000/m² já estão depreciados ("depreciação técnica"). O comparável no grupo é o líquido de R$ 2.536,42/m² (3.170,52 − 20 %), não o bruto.
- **Por que importa:** se alguém decidir por componente, por exemplo aplicando o maior valor de cada componente pelo C0, soma 34,5 + 31,5 milhões de Rafael com 1,512 + 2,592 + 6,4 + 5,57 milhões do grupo, cerca de 82,07 milhões. Esse valor não é nem o de Rafael (66 mi) nem o do grupo (49,6 mi). E as diferenças negativas sugerem que Rafael "zerou" os móveis, coisa que ele não diz.
- **Corrigir:** nas linhas 1.2, 1.3, 1.5 e 1.6, trocar 0 por vazio com o texto "não avaliado em separado (incluído ou omitido nos 66 mi: a definir)". Deixar a diferença só no total. Escrever na aba que a decisão é no nível do item, não do componente. Citar as duas leituras.

### M3 — média — Queda de 1.500 m² na área construída do hotel (2018 → 2025) sem alerta na planilha-mestre
- **Evidência:** inventário 2018, item 01: "12.000,00m² de área construída X R$ 2.154,09". Inventário 2025 (atual e antigo): "10.500,00m² de área construída X R$ 3.170,52". Na aba Item01_componentes, G5/H5 só repetem os dois números em texto. A aba Mestre (AA2) não diz nada. A Fase 0 registrou o fato (revisão M7, resumo item 2), mas ele não chegou à mestre.
- **Por que importa:** o 2018 é a referência de existência e a regra pede conferir no nível de cada componente do item 01. Uma área construída menor é uma diferença de existência dentro do componente 1.4 (12.000 × 2.154,09 = 25,8 mi brutos contra 10.500 × 3.170,52 = 33,3 mi). Rafael usa os mesmos 10.500 m².
- **Corrigir:** nota explícita na linha 1.4 da aba Item01_componentes e na AA2: "área construída 12.000 m² em 2018 e 10.500 m² em 2025; a esclarecer com o grupo (demolição? erro de 2018? área equivalente?)". Também o CUB de dez/2022 (ver M7).

### M4 — média — Área do terreno: a "nesga" de 420 m² é "parte do lote nº 12", que também entra entre os 13 lotes de 360 m²
- **Evidência:** inventário 2025 e PDF p. 6/16: "(Lote nº 12 – Matrícula nº 33.101)" está entre os 13 lotes de 360 m², e "01 lote de 420,00m² - lote da Nesga – sobra com muro. (Parte do Lote nº 12 – Matrícula nº 33.101)". 4 × 450 + 13 × 360 + 420 = 6.900 m². Os "18 terrenos" são 17 lotes mais a nesga. O inventário de 2018 não dá a lista por lote.
- **Por que importa:** se o lote 12 tiver só 360 m² e a nesga for parte dele, a área de 6.900 m² estaria contada a mais em 420 m². Isso dá R$ 420.000 no preço do grupo e R$ 2.100.000 no de Rafael (R$ 5.000/m²). Não está provado que haja erro: o lote 12 pode ter 780 m² na matrícula 33.101.
- **Corrigir:** nota na linha 1.1 (Item01_componentes) e pergunta ao grupo: área da matrícula 33.101.

### M5 — média — Os lotes 13 e 14 da quadra 37 não têm linha nem pergunta na planilha-mestre
- **Evidência:** item 01 (2025) cobre os lotes 01–12 e 18–22 da quadra 37, mais a nesga. Os itens 21, 03 e 22 são os lotes 15, 16 e 17. Os lotes 13 e 14 não aparecem em nenhuma das três versões, nem no PDF de Rafael, nem nas respostas do grupo. A Fase 0 registrou o fato (revisão M9, resumo item 4), mas a planilha-mestre não tem linha "a verificar" nem nota em AA2.
- **Por que importa:** são lotes encravados no quarteirão que o grupo diz que será unificado (inventário 2025, obs. 2; arq. 1, p. 11, resposta 1). Se forem da empresa, são bens perdidos. Se forem de terceiros ou de sócios, pesam na unificação e no valor do item 01. Podem ainda explicar o lote de 360 mil a mais de Rafael (M1), mas isso é só hipótese.
- **Corrigir:** criar a linha "Lotes 13 e 14 da quadra 37 — titularidade a verificar", com Situação "Não abordado", ou ao menos nota em AA2, e levar a pergunta ao grupo.

### M6 — média — Bens de 2018 sem par: a nota "sem documento da troca" é falsa, e há fontes de 2018 não usadas que indicam bem do sócio
- **Evidência:**
  - (a) Mestre AA26 (item 23 de 2018, Élcio) diz "sem documento da troca". O doc 25 de 2018 (`25. 1. PLANILHA = DIVISÃO DE LOTES=ANO 2018 …`), na seção de Élcio Siviero, tabela 3, linha 7, diz: "Lote 10 – Quadra 03 – Loteamento Aditivo Madianeira, que conforme contrato foi dado para o Paulo Busatto, pelo valor de R$ 70.000,00 (Trocado com o lote de Capão da Canoa – CASA)". É um documento do próprio grupo sobre a troca, lançado como retirada de Élcio. Há candidato à Parte 6: lote da SEI usado para o sócio.
  - (b) No mesmo item, o inventário 2018 traz no campo do valor "R$……(Avaliar Lote e Casa) … ELCIO SIVIERO", e o TOTAL GERAL vem com a ressalva "(+ o valor da casa e lote – Élcio Siviero)". A planilha-mestre não registra nenhuma das duas. A aba Totais, A11 "Total 2018 (19 itens com valor)" = 49.684.243,20, também não dá a ressalva do próprio documento de que o total está incompleto.
  - (c) O doc 6 de 2018 (`6. 1. VALTER SIVIERO - Dezembro de 2018.docx`, tabela 2, linha 3) lista "Valter Siviero || 360.0019.004.000 || Capão da Canoa || R$ 500,00", na relação de imóveis do sócio, com o valor de 2018. É a inscrição do item 20 de 2018 (Mestre, linha 24). Os docs 1 (Raul), 4 (Pedro) e 5 (José Valério) também listam cada um um imóvel em Capão da Canoa (sem inscrição). O doc 3 (Irineu) traz 360.0058.001.000 e não traz a 360.0038.015.000, que é o item 21 de 2018, hoje o item 17 da empresa.
  - Na Mestre, AA24–AA28 só dizem "bem da empresa ou do sócio a esclarecer".
- **Por que importa:** o Destino dos cinco bens de 2018 sem par fica "Pendente". Isso está certo, porque ainda não há decisão. Mas a nota sobre a troca está errada e há indícios documentais de que pelo menos o item 20 (Válter) estava em nome do sócio em 2018. Sem isso, a Parte 4 começa sem as fontes.
- **Corrigir:** AA26: trocar "sem documento da troca" por "doc 25 de 2018 (seção Élcio, tab. 3): lote 10/q03 Aditivo Medianeira dado a Paulo Busatto, R$ 70.000, 'trocado com o lote de Capão da Canoa – CASA' (candidato Parte 6)". Acrescentar "(Avaliar Lote e Casa)" e a ressalva do total. AA24: citar o doc 6 de 2018. Na linha A11 de Totais, acrescentar "o total de 2018 exclui, por ressalva do documento, a casa e o lote de Élcio".

### M7 — média — Coluna "Registro de data-base" vazia, mas há dois desvios da data-base
- **Evidência:**
  - Rafael, p. 7, primeira fonte citada para o valor do hotel: `jornaldocomercio.com/especiais/construcao-civil/2025/12/1228669-imoveis-de-capao-xangri-la-e-torres-tem-alta-valorizacao`, uma reportagem de dezembro de 2025 sobre alta de preços. Os 66 milhões (item 01) e os 2,4 milhões (item 20) vêm de anúncios e reportagens atuais, não de 31.12.2024.
  - No lado do grupo, o item 01 usa o "CUB/DEZEMBRO DE 2022", dois anos antes da data-base. Isso consta da Fase 0, mas não da coluna X.
  - O grupo afirma (arq. 1, p. 11, resposta 1) que "A valoração segue preço médio de mercado na data do inventário".
- **Por que importa:** a premissa P2 manda registrar no 0.8 todo valor que fuja de 31.12.2024. Nos dois itens divergentes, Rafael pode estar medindo em outra data. Isso pesa na decisão do item que vale 16,4 milhões de diferença.
- **Corrigir:** abrir no registro de data-base desta empresa uma linha para Rafael (item 01 e item 20: fontes de 2025) e outra para o grupo (CUB dez/2022), e pôr o link nas células X2 e X21.

### M8 — média — A Lista de Rafael (C0) é ligada pelo número do item, e a conferência só compara valores que se repetem
- **Evidência:** `build_mestre_06.py`, l. 97–102: `lista[str(r[3]).zfill(2)] = (r[10], r[13])`; l. 176–178: `na_lista = it["item"] in lista` e `assert abs(lista[...][0] - g) < TOL`. Os valores 390.000 (itens 11–17) e 350.000 (18–19) se repetem. Se a lista trouxesse "13" no lugar de "12", ou se a numeração da lista seguisse outro ano, o assert passaria e a Direção C0 sairia trocada sem aviso. Hoje está **certo**: conferi as inscrições na coluna de observação da lista (360.0038.010, 360.0038.009, 360.0019.003, 360.0038.027, 360.0038.028) com os itens 12, 11, 02, 18 e 19. A lição de método 3 pede chave por inscrição ou matrícula, não por número.
- **Corrigir:** ligar a lista pela inscrição (extraída da coluna "observação", r[14]) ou pela matrícula (r[8]), e conferir que o número do item bate, com assert.

### B1 — baixa — O par do item 01 é forçado no código, e a conferência é fraca
- **Evidência:** l. 123 `pairs["01"] = i18[0]`, sem condição. l. 564–565 só conferem que "34.768" aparece uma vez na descrição de 2018. A lista das 12 matrículas está fixa em texto em `chave()` (l. 52). As 12 matrículas são idênticas nas três versões (conferi), então o par está certo, mas o script não acusaria se uma delas mudasse.
- **Corrigir:** extrair as 12 matrículas das duas descrições e comparar os conjuntos com assert. Gerar a Chave a partir delas.

### B2 — baixa — A aba Totais não mostra o efeito do erro de soma no 1/6 e no crédito de Rafael
- **Evidência:** p. 21: 1/6 = 12.705.000,00 e crédito = 10.825.000,00. Pela soma dos itens dá 12.645.000,00 e 10.765.000,00. Está no resumo da Fase 0, mas não nesta aba.
- **Corrigir:** acrescentar as duas linhas.

### B3 — baixa — "bem novo em 2025" nos itens 21 e 22 é uma suposição
- **Evidência:** Mestre W22 e W23, "— (bem novo em 2025: origem a documentar)". O que se sabe é que os dois não estão no inventário de 2018. O item 22 está registrado na Comarca de **Osório**, matrícula 124.906, o que sugere matrícula antiga, anterior ao registro de Capão da Canoa. Pode ser um bem antigo omitido em 2018, ou vindo de sócio (os docs 1, 4 e 5 de 2018 listam lotes em Capão sem inscrição). Na Pares_2018_2025, o critério "sem par" está certo.
- **Corrigir:** "sem par no inventário 2018; data e forma de aquisição a documentar".

### B4 — baixa — As referências dão a tabela, não a página, e o código de imóvel não aparece
- **Evidência:** colunas F e I: "item NN (tabela N)". O plano (0.7) pede "item, página do doc". O código de imóvel (por exemplo, 039416 no item 21 de 2018 e no item 17 de 2025; 038999 do item 20 de 2018, vizinho do 038998 da Casa da Praça) é uma terceira chave que confirma os pares e não está na Chave nem na Pares_2018_2025.
- **Corrigir:** acrescentar o código à Chave e uma coluna "código" à Pares_2018_2025. Página do docx, se viável; senão, dizer que "tabela" é o número da tabela do docx.

### B5 — baixa — Valores de Situação e de "Existe no Estudo 2" fora do padrão do plano
- **Evidência:** Q24–Q28 = "Só 2018", que não está na lista do plano (Concorda / Discorda / Só grupo / Só Rafael / Não abordado). K25–K26 = "Não" para os itens 22 e 23 de 2018, enquanto a nota diz que eles podem ser as residências de Rafael (p. 24).
- **Corrigir:** registrar "Só 2018" como extensão na legenda, ou usar "Não abordado". Nos itens 22 e 23 de 2018, pôr "a verificar (possíveis residências da p. 24)".

### B6 — baixa — "Observação do Luís" é texto fixo no script e cita só a planilha 2_0
- **Evidência:** l. 237, 290 e 338: o texto fixo "sem observação na planilha 2_0 para esta empresa". O script não lê `observacoes-planilhas-google.xlsx`. Conferi: não há observação para as linhas 70–103 da GERAL da 2_0 nem para as abas `6.Hotel Araçá` das duas planilhas. O conteúdo de hoje está certo, mas não se atualiza.
- **Corrigir:** ler o arquivo de observações filtrando planilha, aba e linha, e citar as duas planilhas.

### B7 — baixa — As faixas de mercado de Rafael não aparecem na aba Mestre
- **Evidência:** p. 10–11 (lotes das quadras 37/38: 380–550 mil; esquina 550–600 mil; meio de quadra perto de 380 mil), p. 13–15 (item 20: 2,4–3,5 milhões), p. 7–8 (hotel: 55–85; físico 66–102,45; renda 41,0–50,6 milhões). A Mestre só traz o valor adotado. Nos itens 03–10 (360 mil), Rafael adota um valor abaixo do piso da própria faixa, e a Situação "Concorda" está certa pela concordância expressa (p. 12 e 16). Os itens 04–10 têm inscrição da quadra 0071, e a faixa de Rafael é para as quadras 37/38.
- **Corrigir:** uma coluna ou nota "faixa de Rafael (p.)" nas linhas 2–23.

### B8 — baixa — A concordância expressa de Rafael cobre os itens 02 a 19, não os itens 21 e 22
- **Evidência:** p. 16, "de acordo como citado nos itens anteriores de 02., à 19.", e p. 12. Os itens 21 e 22 aparecem a 380.000 nas p. 15 e 20, sem texto de concordância. A Situação "Concorda" (Q22–Q23) sai só da igualdade de valores.
- **Corrigir:** nota "concorda por valor (p. 20); sem declaração expressa".

### B9 — baixa — O rótulo "15 % da obra feita" esconde que a base da vantagem inclui móveis e equipamentos
- **Evidência:** inventário 2025: a base é R$ 37.136.368,00 = 26.632.368 (obra) + 6.400.000 + 2.592.000 + 1.512.000 (móveis e equipamentos). Em 2018, R$ 31.903.264 segue a mesma composição. Item01_componentes, G7: "15,00 % da obra feita".
- **Corrigir:** acrescentar "base = obra civil depreciada + móveis + equipamentos (37.136.368,00)". É relevante para M2: se os móveis ficam fora dos 66 mi, a vantagem também perde parte da base.

### B10 — baixa — O método da renda de Rafael não se reproduz com as premissas dele
- **Evidência:** p. 8–9: 140 quartos (o hotel tem 144), diária de R$ 520, ocupação de 45 %, margem de 30 %. 140 × 520 × 365 × 0,45 × 0,30 = R$ 3,59 mi/ano; ÷ 10,5 % = 34,2 mi (ele escreve 41,0 mi); ÷ 8,5 % = 42,2 mi (ele escreve 50,6 mi). Não é reavaliação nossa, só a aritmética dele. Útil para a decisão do item 01.
- **Corrigir:** registrar na nota do Total de Item01_componentes ou no pacote de exceções.

### B11 — baixa — A nota do item 20 omite a afirmação do grupo
- **Evidência:** AA21, "titularidade registral a conferir". O arq. 1, p. 11, resposta 1, diz "reafirmando estarem todos os imóveis registrados em nome da empresa". O inventário cita a matrícula 69.488 e o "Contrato Particular" de maio/2014. Rafael avalia uma casa de 160 m² (p. 13–14), contra 158,18 m² no inventário.
- **Corrigir:** citar a afirmação do grupo e manter o pedido da matrícula.

### B12 — baixa — Descrições e notas fixas no código, presas ao número do item
- **Evidência:** l. 62–64 (descrição do item 01), l. 203–206 (nota do item "20"), l. 242–243 ("21", "22"), l. 258–260 (itens "22" e "23" de 2018), rótulos de Totais com contagens fixas ("22 itens", "19 itens com valor", "5 itens", l. 462–499). Se a numeração mudar, o texto vai para o bem errado ou a contagem fica errada sem acusar.
- **Corrigir:** prender as notas à inscrição e calcular as contagens.

### B13 — baixa — Pagamentos feitos pelo Hotel Araçá a favor de sócios (doc 25 de 2018), fora do escopo da mestre mas sem registro
- **Evidência:** doc 25 de 2018, tabela 1 (Adélia): apartamento no Edifício Contemporâneo (2011) "pago pela empresa Hotel Araçá Ltda. e Siviero Empreendimentos", R$ 680.000,00, e mais pagamentos do Hotel (Multiaquecimento, Castellar, SPA, Hidrolar, Vuelma). Tabela 11 (Raul): outro apartamento no mesmo edifício, R$ 400.000,00, pago pelas mesmas empresas. Não são bens do hotel em 2025, mas são saídas de patrimônio da empresa a favor de sócios.
- **Corrigir:** anotar como candidato à Parte 4 ou 6 no diário desta empresa. Não entra na mestre.

---

## O que conferi e achei correto

1. **Aba Mestre, linhas 2–23 (22 itens de 2025):** valor do grupo (inventário atual), valor 2025 antigo (difere só no item 01: 53.320.460,00), valor de Rafael do balanço executado (66.000.000 no item 01, p. 17; 2.400.000 no item 20, p. 19; os outros 20 iguais ao grupo), páginas de Rafael (p. 17: 01–05; p. 18: 06–13, com o valor do 13 na p. 18 e a inscrição na p. 19; p. 19: 14–20; p. 20: 21–22), Diferença R$ e % (16.393.176,80 = 33,0 %; 1.400.000 = 140 %), Situação (20 concordam, 2 discordam), valor de 2018 de cada par e referências de item e tabela de 2018 e 2025. Tudo confere com o texto das fontes.
2. **Lista de Rafael:** 5 linhas "06 - Hotel Araçá" em `ativos-avaliados-valter.xlsx` (itens 12, 11, 02, 18, 19; total 1.880.000 = p. 21–22). R3, R12, R13, R19 e R20 = "Sim"; o resto = "Não". Conferi também pela inscrição.
3. **Direção C0:** só os itens 01 e 20 divergem. Nenhum dos dois está na lista, logo "maior (fica na empresa)" está certo nos dois. Nenhuma decisão foi aplicada: `decisoes.json` tem `decisoes: {}` e `destinos_2018: {}`, Decisão = "Pendente", Valor adotado vazio, e Totais, A13 = 0.
4. **Pares 2018 × 2025:** refiz os pares pelas inscrições e matrículas de 2018 e de 2025, conferindo também pelo código de imóvel. Os 19 lotes pareados estão certos: 2018 02–16 → 2025 02–16; 17 → 18; 18 → 19; 19 → 20; 21 (Irineu, sem matrícula, código 039416) → 17 (mat. 33.003, código 039416). O item 01 tem as mesmas 12 matrículas nas três versões. Os itens de 2018 sem par (20, 22, 23, 24, 25) não aparecem em 2025 nem por inscrição nem por matrícula (77.833 e 74.857 não constam de 2025). Os itens de 2025 sem par (21 e 22: mat. 86.669 e 124.906, códigos 39394 e 39396) não constam de 2018. Nenhum par se apoia em semelhança de texto.
5. **Item01_componentes, valores do grupo:** 2018 (3.800.000; 1.800.000; 3.024.000; 20.679.264; 6.400.000; 9.570.979,20 = 45.274.243,20), 2025 antigo (… 9.284.092 = 53.320.460,00) e atual (… 5.570.455,20 = 49.606.823,20). Conferidos nos docx. A conta de Rafael 34.500.000 + 31.500.000 = 66.000.000 confere com a p. 7.
6. **Totais:** 58.076.823,20; 75.870.000; 76.230.000; 360.000; 7.470.000 (20 itens); 50.606.823,20 × 68.400.000; 17.793.176,80 + 360.000 = 18.153.176,80; 2018 = 49.684.243,20; lista = 1.880.000. Tudo fecha. A transcrição do grupo por Rafael (p. 5) fecha em 58.076.823,20, ou seja, o erro de 360 mil está só no balanço executado.
7. **Linhas "Só Rafael":** residências da Rua da Divisa (p. 24–25, faixa de 2,4 a 3,5 milhões pelo par) e Fiat Toro (p. 25, sem valor) estão bem referenciadas e fora do total de Rafael. Procurei "Toro", "Divisa" e as inscrições dos bens de 2018 sem par em todos os inventários 2025 e 2018 das outras empresas e dos sócios: nenhuma outra empresa lista a Toro ou as casas (sem risco aparente de dupla contagem).
8. **Notas de identificação:** "texto diz quadra 37; inscrição é da quadra 0071" (itens 04–10), "matrícula 73.482 repetida nos lotes 27 e 28/q38", "registro na comarca de Osório" (item 22), "antiga matrícula 37.549" (item 02) e "em 2018 era o lote 'IRINEU SIVIERO'" (item 17) estão corretas.
9. **Aba Observacoes:** os parágrafos de 2025 atual (11) e de 2018 (2) conferem com os docx. As observações do inventário antigo são iguais às do atual.
10. **Colunas do plano 0.7:** todas presentes (Empresa, Bloco, Chave, existência nos três estudos com referência, valores, diferença R$ e %, Situação, Decisão, Motivo, Destino, Registro de data-base, Observação do Luís, Situação da observação), mais Descrição, Valor 2018, Valor 2025 antigo, Lista C0, Direção C0, Valor adotado e Notas.
11. **Script:** os valores de Rafael por lote são ligados pela inscrição (não pela ordem), com 1 matrícula, 1 inscrição e 1 valor por bloco (assert). O valor do item 01 de Rafael é conferido pela decomposição da p. 7 (assert). A soma dos itens de cada inventário é conferida com o TOTAL GERAL (assert). As duas conciliações de Totais têm assert.
