# Revisão independente — Fase 0, empresas 01 (Abastecedora) e 08 (Transportadora)

- **Data:** 28/10/2026
- **Revisor:** agente revisor independente (contexto novo). Não li `_diario/` (exceto para gravar este arquivo) nem `pacotes/`.
- **Objeto:** `_extracao/fase0-itens-01-08.xlsx` (e `_scripts/extract_fase0.py`), `_extracao/*.txt`, `arq10-rovilio-transcricao.md`, `contratos-abastecedora.md`, `hashes.txt`.
- **Método:** releitura direta das fontes (docx via XML e python-docx, PDFs do Rafael via PyMuPDF, xlsx com fórmulas e com valores em cache, imagens 150/300 dpi dos contratos), recálculo das somas e composições, busca por matrícula, placa e chassi nos inventários 2018 e em todos os inventários 2025 (01 a 11).

Resumo: **2 achados de gravidade alta, 9 média, 7 baixa.** Os itens e valores da aba "Itens inventários" estão todos certos. Os erros estão na leitura do Rafael da Abastecedora, nas faixas de preço do Rafael, nas conferências que faltam e no que não foi extraído da planilha Google.

---

## Achados

### A-01 · ALTA · Rafael 01: o valor que Rafael adota para o item 01 (R$ 12.361.850,00) ficou sem número de item e sem conferência
- **Descrição:** Na aba "Rafael itens", a linha `Rafael 01 · Balanço executado por Valter · p. 3 · 12.361.850` está com o **item vazio**. O único valor de Rafael ligado ao item 01 é 9.661.414 (balanço apresentado pelo grupo, p. 1). Quem filtrar por item 01 conclui que Rafael usa R$ 9.661.414, mas ele adota **R$ 12.361.850,00** em três lugares: no texto da p. 2 ("adota-se ... o valor previamente estabelecido de R$ 12.361.850,00"), no balanço executado da p. 3 e no quadro-resumo da p. 4.
- **Contradições internas do PDF, a registrar:**
  1. Na p. 3, o valor de R$ 12.361.850,00 aparece ao lado da **descrição de 2025**, cujas parcelas somam R$ 9.661.414,00 (7.245.000 + 439.275 + 769.505 + 650.000 + 557.634).
  2. R$ 12.361.850 é o **total da empresa em 2018** (doc 12: R$ 12.361.851,89, truncado em R$ 1,89). Esse total **inclui o caminhão-tanque ICW 8342 (R$ 145.000)**, que não existe em 2025. O posto sozinho, em 2018, valia R$ 12.216.851,89 (doc 12, T3.R2).
  3. Na p. 3 consta "BEM DISPONÍVEIL — imóvel 3.752", mas a p. 5 diz que "não foram identificados bens ou direitos passíveis de seleção".
- **Evidência:** `ABASTECEDORA ... BELISE LTDA.pdf` p. 2–4. Doc 12 de 2018, T3.R2–R4. Aba "Rafael itens", linha 81 do dump.
- **Correção sugerida:** preencher item = 01 nessa linha e criar a coluna "valor adotado por Rafael" (01 item 01 = 12.361.850, com página 2/3/4). Criar a conferência "Rafael 01 adotado × doc 12 (2018)", que dá diferença de −1,89, com nota de que o valor de 2018 inclui o caminhão ICW 8342.

### A-02 · ALTA · Rafael 01: quadro-resumo (p. 4–5) não extraído; o 1/6 de Rafael está errado
- **Descrição:** Os valores do quadro estão escritos "R$ 12.361,850" e "R$ 2.060,166", fora do padrão `,dd`, e o script não os captura. Ficaram fora da planilha: patrimônio R$ 12.361.850, participação de 1/6 R$ 2.060.166, "0,16 bens" e "Crédito Patrimonial a Valter Participações Ltda R$ 2.060,166". Além disso, **1/6 de 12.361.850 = R$ 2.060.308,33, e não R$ 2.060.166**. O número de Rafael corresponde a 12.361.000 / 6 = 2.060.166,67. A diferença é de R$ 142,33 no crédito atribuído.
- **Evidência:** PDF Rafael 01 p. 4 e p. 5. No PDF 08, a mesma conta está certa: 16.924.986 / 6 = 2.820.831 (p. 16–17).
- **Correção sugerida:** extrair os quatro valores do quadro (p. 4–5), normalizando a escrita, e incluir a conferência "Rafael 01: 1/6 × total", que hoje dá diferença de −142,33, como contradição interna do PDF.

### M-01 · MÉDIA · "Rafael itens" não diz qual valor Rafael adota por item da 08; a seção "Avaliação" repete os valores do grupo
- **Descrição:** A aba lista três seções por item sem marcar o valor adotado. Na seção "Avaliação executada por Valter" (p. 4–12), os itens 03 e 10–13 trazem os **valores do grupo** (500.000; 140.000; 30.000; 140.000; 30.000). Rafael diverge deles no texto. Os valores que ele adota estão no "Balanço executado" (p. 13–15): item 03 = **1.000.000**, que vem da "Declaração expressa" de Benta (p. 5–6, "aproximadamente R$ 1.000.000,00", que o script não captura por estar no meio da linha); item 11 = **50.000**; item 12 = **160.000**; item 13 = **40.000**. Item 10 = 140.000, sem mudança. Os demais itens têm o mesmo valor do grupo, com a "concordância" das p. 4–8. O item 03 e os itens 10–13 **não têm parágrafo de concordância**.
- **Correção sugerida:** criar a coluna "valor adotado" com a página (08: p. 13–15; total p. 16 = 16.924.986 = 16.374.986 + 500.000 + 20.000 + 20.000 + 10.000) e registrar a declaração da p. 6.

### M-02 · MÉDIA · "Rafael faixas": duas faixas no item errado e várias faixas não capturadas
- **Descrição:** O script toma o "item corrente" no fim da página, mas a faixa está no texto da página inteira.
  - p. 9, R$ 145.000–170.000 (preço de anúncio do **Scania 2002, item 10**) está marcada como item 11.
  - p. 10, R$ 65.000–75.000 (Graneleira/LS, **carreta Randon 2009, item 11**) está marcada como item 12.
  - O regex só pega "entre R$ X e R$ Y". Ficaram fora as faixas escritas "R$ X a R$ Y":
    - item 10: FIPE R$ 106.000 a 145.000 (p. 9);
    - item 11: caçamba R$ 50.000 a 80.000, baú/sider R$ 80.000 a 90.000, tanque "a partir de R$ 60.000" (p. 10);
    - item 12: truck/bitruck com prancha R$ 200.000 a 310.000 (p. 11) e conjunto com carreta prancha R$ 240.000 a 280.000 (p. 12);
    - item 13: R$ 40.000–50.000, 50.000–60.000, 65.000–70.000 e 83.000–85.000 (p. 12).
- **Correção sugerida:** atribuir o item pela posição do texto (bloco entre o "N." e o "N+1."), aceitar o padrão "R$ X a R$ Y" e corrigir as duas linhas.

### M-03 · MÉDIA · Faltam conferências de composição nos itens 07 e 08 da 08 de 2025 e na nota de R$ 1.135.000
- **Descrição:** Os inventários de 2025 (atual e antigo) e o Rafael (p. 3, 7, 14–15) trazem contas que não fecham, e a extração não aponta:
  - item 07: 69,8 ha × R$ 50.000 = **R$ 3.490.000**, mas o valor declarado é R$ 3.500.000 (+10.000);
  - item 08: 4,5 ha × R$ 50.000 = **R$ 225.000**, mas o valor declarado é R$ 250.000 (+25.000);
  - nota final do doc 08: "FOI ACRESCENTADO ... R$ 1.135.000,00 ... ITENS 03 – 04 – 05", mas esses itens somam **R$ 1.020.000** (500.000 + 300.000 + 220.000), uma diferença de 115.000.
- **Evidência:** `08. INVENTÁRIO ...docx` T4.R2, T4.R3 e observações finais.
- **Correção sugerida:** acrescentar essas três conferências na aba "Conferências" como divergências do documento.

### M-04 · MÉDIA · A extração não traz os pares 2018 × 2025 por matrícula, placa ou chassi
- **Descrição:** Nenhuma aba relaciona os bens de 2018 com os de 2025. Fiz o pareamento por chave (resultado em "Conferido sem erro", item 6). Pontos que a extração deveria registrar:
  - Bens de 2018 **sem par** em 01/08:
    - caminhão-tanque **ICW 8342** (doc 12 item 02), sem par em nenhum inventário de 2025;
    - sala do consultório, **mat. 18.952** (doc 15 item 05), sem par em nenhum inventário de 2025;
    - móveis do escritório (doc 15 item 07, sem chave);
    - apto 102, **mat. 7.395** (doc 15 item 09, doado, já deduzido do total);
    - terreno **mat. 3.754** (doc 10 item 06) e pavilhões **mat. 17.409 + 6.136** (doc 10 item 07). Esses dois **estão em 2025 na SEI** (doc 11, item 03 = mat. 3.754 com R$ 4.950.000; item 04 = mat. 29.954, oriunda da unificação de 17.409 e 6.136, com R$ 13.631.378,60), e não em 01/08.
  - Os itens 02 a 05 da 08 em 2025 (pavilhão, equipamentos, rampa, borracharia) e os itens 02 a 05 do doc 10 **não têm chave**. O par entre eles é só por descrição e posição.
- **Correção sugerida:** criar a aba "Pares 2018×2025" com a chave usada e o tipo de par (chave / descrição / sem par / migrou para outra empresa).

### M-05 · MÉDIA · Planilha Google exportada como fórmulas, sem valores calculados e sem cabeçalho
- **Descrição:** O script abre o xlsx com `data_only=False`. Por isso os totais aparecem como `=SUM(C4:C18)` e não como número, e a linha 1 (cabeçalhos) não é exportada. Sem cabeçalho, a coluna D ("2018: atualizado") parece um valor de 2018, mas é `=G`/`=E`, ou seja, o **valor de 2025**. Os valores em cache que deveriam constar e ser conferidos:
  - 1.Abastecedora/GERAL linha 5: C = 12.361.851,89 (= doc 12); D = 9.806.414 (= 9.661.414 + 145.000); E/F/G = 9.661.414 (= doc 00); diferença −145.000.
  - 8.Transportadora C2 / GERAL C116: "2018 apresentado" = **12.285.998,84**. Esse valor **não bate com nenhum documento de 2018**: é a soma dos itens 01–05 do doc 10 (Oficina, 10.721.998,84) com os itens do doc 15 (1.564.000), sem os itens 06–07 da Oficina, que foram para a SEI. D = 16.553.986; E/F/G = 16.374.986 (= doc 00); diferença −179.000.
  - Nas abas 1 e 8, os valores de dez/2024 estão na coluna **F, que não tem cabeçalho**. O cabeçalho "dezembro 2024" está na coluna E, que está vazia.
  - GERAL C124 = "Não divisível" e 8.Transportadora C10 = "?" para o mesmo item 07; C125/C11 = "?" para o item 08.
- **Correção sugerida:** reabrir o arquivo também com `data_only=True`, exportar o cabeçalho e criar a conferência "planilha × inventário" com nota sobre a base composta de 2018 da 08.

### M-06 · MÉDIA · Textos do autor da planilha sobre 01/08 fora das linhas extraídas
- **Descrição:** Nas linhas pedidas, o único texto próprio do autor é "Por que reduziu R$2.555.437,89" (1.Abastecedora I7 / GERAL H7), e ele foi extraído. A conta confere: 12.216.851,89 − 9.661.414. Os blocos de "Observações" são cópia do texto dos docx. Fora dessas linhas, porém, há notas do autor que tratam diretamente de 01/08:
  - **GERAL H874** e **Resumo B57–C63/B67**: "Rovílio tinha direito à 2% ... Abastecedora R$ 193.228,28 / Transportadora R$ 327.499,72 ... total R$ 2.538.208,55 ... A avaliação que eles fizeram está 36.53% abaixo do acerto com Rovílio".
    - A conta do autor não fecha: 3.440.082 / 2.538.208,55 − 1 = **35,53%**, e 2.538.208,55 fica **26,22% abaixo** de 3.440.082.
    - A nota ignora o 1% de Rovílio na Turismo e na Siviero Transportes.
  - **Resumo B73**: o rótulo diz "Hotel Princesa dos Vales", mas a célula aponta para `=GERAL!C8/E8` (R$ 145.000 → 0), que é o **caminhão-tanque ICW 8342 da Abastecedora**. O rótulo está trocado.
- **Correção sugerida:** registrar essas notas numa aba de observações, com a ressalva de que estão fora do intervalo pedido.

### M-07 · MÉDIA · Planilha Google: linha de OBSERVAÇÕES da aba 8 omitida e textos cortados em 200 caracteres
- **Descrição:** `range(2, 19)` para "8.Transportadora" para na linha 18, e a **linha 19 (OBSERVAÇÕES, 4.503 caracteres)** fica de fora. Nas demais abas, as observações (1.Abastecedora B9 com 1.955 caracteres; GERAL B9 e B133) foram cortadas em 200 caracteres. O conteúdo é igual ao dos docx de 2025, e GERAL B133 difere de 8.Transportadora B19 só em linhas em branco. Mesmo assim, a extração não mostra isso.
- **Correção sugerida:** incluir a linha 19 e gravar o texto completo, ou registrar que ele é idêntico ao docx.

### M-08 · MÉDIA · Contradições internas do PDF Rafael 08 não registradas
- **Descrição:**
  1. A p. 17 diz que existe "crédito patrimonial ... pelos **bens imóveis selecionados** de R$ 2.820.831,00" e, logo depois, "não há interesse ... **não foram identificados bens ou direitos passíveis de seleção**".
  2. A p. 16 chama de "bens imóveis" um total que inclui equipamentos, caminhões e carretas.
  3. As p. 13 (08) e p. 3 (01) têm o cabeçalho "GRUPO SIVIERO — SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA.", que parece copiado de outro documento.
  4. As concordâncias das p. 5–8 omitem o item 03 e os itens 10–13, o que é coerente com a divergência de valor.
- **Correção sugerida:** registrar como notas na aba Rafael.

### M-09 · MÉDIA · Conferências faltantes com fontes que existem na pasta
- **Descrição:** A extração não confere:
  1. os totais antigos de 01/08 com o **`ANTIGO 00. PLANILHA ...docx`**, que existe em `01_fontes/inventario-2025-antigo/`. Conferi: 01 = 9.661.414 e 08 = 16.374.986, ambos certos;
  2. item a item, o "Balanço apresentado" do Rafael com o docx de 2025. Conferi: os 13 itens da 08 e o item 01 da 01 estão iguais;
  3. o preço do contrato de Rovílio com a soma dos 10 itens de pagamento. Conferi: 3.440.082,00, certo.
- **Correção sugerida:** incluir essas três conferências na aba "Conferências".

### B-01 · BAIXA · Conferência tautológica
- A linha "2025 atual 01 item 01: parcelas + 30 % = valor do item" compara o valor recomposto com ele mesmo, porque o valor da célula está vazio. A verificação independente é só "soma = total". Marcar como "não aplicável" ou explicar na nota.

### B-02 · BAIXA · Marcas "x2" do doc 00 (2025) não registradas
- `00. PLANILHA ...docx` T2.R2, T2.R9 e T2.R12 trazem o texto literal "x2" depois de 01, 08 e 11, que são exatamente as três empresas em que Rovílio tinha 2%. Também em T2.R12 está escrito "R$ 100.874,027,60" (na T3 o valor é 100.874.027,60). Registrar como anotação do documento.

### B-03 · BAIXA · hashes.txt incompleto
- O doc 17 de 2018 (Agropecuária) é fonte da extração e aparece na aba Fontes (`dc0476f1f7`, que conferi), mas **não está em `hashes.txt`**. Também faltam o `ANTIGO 00` e o `11. INVENTÁRIO` (SEI), que foi lido para gerar `2025-11-sei.txt`.

### B-04 · BAIXA · Pequenas inconsistências nos docx não anotadas
- Antigo 01/08: a cota de Rovílio está em R$ 193.228,24 e R$ 327.499,70, mas 2% dá 193.228,28 e 327.499,72. Os percentuais somam 16,33% × 6 + 2% = 99,98%.
- Atual: 16,66% × 6 = 99,96%.
- Contrato de Rovílio, p. 3–4: 16,33% × 6 + 2% = 99,98% com "TOTAL 100,00%".

### B-05 · BAIXA · arq10: lista de números incompleta
- A seção "Percentuais" omite os **10,00% de honorários** (Décima Primeira, item 9, p. 8). Falta também a soma de conferência dos 10 itens (= 3.440.082,00).

### B-06 · BAIXA · arq4: dados da p. 14 omitidos
- CPFs das testemunhas, manuscritos: Marcos A. Siviero "817.112.550-6[8?]" e Fabiana B[e]idrich "004.195.130-19".
- Etiqueta notarial: Serviço Notarial de Veranópolis, tabeliã Caroline Zechin Sbrissa, selo 0736.01.2100001.26848, escrevente Vilson Lucas Nalin, emolumentos R$ 7,80 + selo R$ 1,40.
- Cl. 22ª, §2º: ao fim da locação, a Locadora escolhe livremente o transportador.

### B-07 · BAIXA · Doc 17 de 2018 com duas linhas de "TOTAL GERAL" na aba "Totais declarados"
- As linhas T5.R6 e T7.R1 têm o mesmo valor (24.112.390). Isso não afeta nenhuma conferência, mas convém deduplicar ou indicar a tabela de origem.

---

## Conferido sem erro

1. **Itens inventários (49 linhas):** conferi cada item e valor contra o docx de origem (python-docx e XML bruto).
   - 2025 atual e antigo, 01 e 08: 1 + 13 itens, todos certos.
   - O item 01 da 01 atual tem de fato a célula de valor vazia ("R$"), confirmado no XML. Os docx não têm alterações controladas, caixas de texto, campos ou comentários, e as notas de rodapé estão vazias.
   - Antigo × atual: itens idênticos. Mudam só o quadro de cotas (16,33% + Rovílio 2% → 16,66%) e a nota sobre a venda das quotas.
   - 2018: doc 12 (2 itens), doc 15 (9 itens), doc 10 (7 itens; a lista de equipamentos com "...." foi corretamente excluída) e doc 17 (itens 07/08 "NÃO DIVISÍVEL", sem valor) estão todos certos. Nenhum item de outra tabela foi capturado por engano.
2. **Totais e composições:**
   - 01 2025: 7.245.000 + 439.275 + 769.505 + 650.000 + 557.634 = 9.661.414; 30% × 1.858.780 = 557.634.
   - 08 2025: soma = 16.374.986.
   - Doc 12: 12.216.851,89 + 145.000 = 12.361.851,89; 30% × 2.762.962,99 = 828.888,90.
   - Doc 15: 1.784.000 − 220.000 = 1.564.000.
   - Doc 10: 29.303.377,44; item 02: 969.086,80 + 290.726,04; item 07: 3.884.060 + 5.700.000 + 3.113.322 + 933.996,60 = 13.631.378,60; item 06: 4.986 × 992,78 ≈ 4.950.000.
   - Doc 15 itens 05/06: 2 CUB × 1.407,45 × área ≈ 149.000 / 685.000.
   - 08 item 01: 8.700 × 992,78 = 8.637.186; item 02: 1.298 × 1.100 = 1.427.800.
   - Doc 27 (2018): 12.361.851,89 / 1.564.000 / 29.303.377,44 / 12.054.390. Doc 00 (2025): 9.661.414 / 16.374.986; a soma do doc 00 fecha em 263.410.451,80.
   - As 24 conferências da aba estão numericamente corretas.
3. **Rafael 08:**
   - p. 1–3 = inventário 2025, 13 itens, total 16.374.986.
   - p. 13–16: itens 01–13 = 8.637.186 / 1.427.800 / 1.000.000 / 300.000 / 220.000 / 840.000 / 3.500.000 / 250.000 / 360.000 / 140.000 / 50.000 / 160.000 / 40.000, total 16.924.986 (soma confere).
   - 1/6 = 2.820.831 (confere). "2,16 bens" = 13/6.
   - Nenhum bem de 01/08 aparece em `ativos-avaliados-valter.xlsx` (abas Ativos/Resumo/Composição; busquei matrículas, placas e nomes). Os PDFs "Diretrizes" e "Introdução" só citam o grupo de forma genérica e o 1/6. A citação a "Cotiporã" nas Diretrizes refere-se a áreas da SEI, não às matrículas 459/460/461/558 e 15.922.
4. **Planilha Google:** todos os valores de 2018 e 2025 das linhas extraídas batem com os inventários, item a item (com a ressalva do M-05 sobre a base composta de 2018 da 08). As linhas 10 e 134 da GERAL e a linha 10 da 1.Abastecedora estão vazias, então o recorte do script não perdeu dado nelas. Os 6 comentários da planilha estão em GERAL B223–B238 (SEI), fora do escopo.
5. **Transcrições conferidas nas imagens:**
   - arq10: p. 1, 3, 4, 5, 6. Percentuais 16,33 / 2,00 / 16,50 / 1,00; preço R$ 3.440.082,00; itens 1–10 (valores, lotes, quadras, condomínios, datas 24/03/2023, fev/2025, 01/11/2012, 25/10/2013, mai/2022); tabelas da p. 6. Tudo certo.
   - arq2: p. 1, 2, 4. Partes, CNPJs, volumes (2.283,60 m³/ano; 10.847,10 m³ = 2.283,60 × 4,75), vigência 01/01/2021–30/09/2025, garantidor siviero@..., tabela da multa. Tudo certo.
   - arq3: p. 1. Interveniente Abastecedora Siviero, R$ 750.000 e as quatro parcelas (250.000 + 250.000 + 170.762,47 + 79.237,53 = 750.000), "(duzentos e cinquenta reais)" sic. Tudo certo.
   - arq4: p. 1, 2, 7, 9, 12, 14. Partes, 144 meses a partir de 01/01/2022, aluguel R$ 15.000, área de 3.450 m² mat. 3.752, área de 2.500 m² mat. 3.754, mat. 3.753 só na cl. 9ª §2º, multas de 10%/12%, cl. 22ª (21/03/2022; 22/03/2021), data 19/11/2021. Tudo certo.
6. **Pares 2018 × 2025 por chave:**
   - mat. 3.752: doc 12 item 01 ↔ 01 item 01;
   - mat. 3.753: doc 10 item 01 ↔ 08 item 01;
   - IKR 2054, IQB 0612, IJZ 5399, IJZ 8207: doc 15 itens 01–04 ↔ 08 itens 10–13;
   - mat. 7.393: doc 15 item 06 ↔ 08 item 06;
   - mat. 7.400: doc 15 item 08 ↔ 08 item 09;
   - mat. 459/460/461/558: doc 17 item 07 ↔ 08 item 07;
   - Registro 15.922: doc 17 item 08 ↔ 08 item 08.
   - Sem chave (só descrição): doc 10 itens 02–05 ↔ 08 itens 02–05.
   - Nenhum bem de 01/08 em 2025 ficou sem par em 2018. Os bens de 2018 sem par estão listados no M-04.
7. **Hashes (10 primeiros dígitos do MD5):** conferi 08 atual `d3882994f6`, Rafael 01 `84a05e9b14`, arq10 `23d9b1732d`, planilha Google `11a81ef0aa`, doc 15 `19a23a2931`, arq4 `361f8b3a52` (MD5 completo igual ao do `contratos-abastecedora.md`) e doc 17 `dc0476f1f7` (igual ao da aba Fontes). Todos batem.
