# Revisão independente da Fase 0, rodada 2 — empresas 02 e 05

Revisor: agente de contexto novo, 28/10/2026. Objeto: `_scripts/fase0_extracao.py` (v2) e `fase0-extracoes-02-05.xlsx`.

**Método.** Rodei o script uma vez: 37 fontes, 116 conferências, 10 DIVERGE, 5 avisos. A planilha foi regravada e não alterei mais nada. Li as fontes por conta própria:
- docx com python-docx (parágrafos e tabelas): docs 02, 05 e 00 de 2025, docx "áreas vendidas", docs 13 e 17 de 2018 e doc 11 de 2025;
- PDF de Rafael (Hotel) com pymupdf, p. 5–17;
- planilha 2_0 com openpyxl (valores e fórmulas), `observacoes-planilhas-google.xlsx` e xlsx da SEI 2018.

Também busquei, com limite de dígito, cada matrícula de 2018 em **todos** os docx de 2025 (atuais e antigos, parágrafos e tabelas). Recalculei os MD5 das 37 fontes: todos conferem.

## 1. Achados da rodada 1 → situação

| Achado | Situação | Evidência |
|---|---|---|
| A1 Créditos de soja | **Corrigido** | Há a aba "Créditos de soja" com 11 linhas: Rigo 40.300 − 2.606 = 37.694 e Lorenzet 13.085, ambos lidos dos parágrafos P12–P29 do docx. Conferências OK. A divergência de 3.682 sacas da "AGROPECUÁRIA/OLEPOPLAN" (doc 02 = R$ 0) aparece como DIVERGE. A conferência FALTA RECEBER dá 2 × 20.150 × 80 = 3.224.000 (DIVERGE). Há aviso para o número 2.606 repetido. Ressalvas em N3 e N4. |
| A2 Venda parcial | **Corrigido** (resta um detalhe) | Item 05 de 2025 (atual e antigo) e seq. 1 de 2018: situação "VENDIDA PARCIAL (83,47 ha)", comprador RIGO e colunas "Texto da célula de valor" e "Área avaliada" (407,53 ha) preenchidas. Falta a área avaliada em 2018 (491 ha), que fica vazia. |
| M1 Benchmark digitado | **Corrigido** | A constante saiu. A aba "Base 2018" traz A = 46.459.267,83 (calculada) e B = 48.730.897,83 (GERAL C28, = 5.Hotel C2, conferido). A conciliação de 2.271.630 confere. A 2ª conferência não acrescenta informação (N7). |
| M2 Áreas digitadas | **Corrigido** | 686,75 / 190 / 85,32 / 411,43 são lidos dos parágrafos do doc 17 e do docx; 411,4 vem do doc 05; 491 e 83,47 vêm da descrição; 407,53 vem da célula de valor. Conferências (a) OK, (b) DIVERGE 0,10 e (c) DIVERGE 0,13. Na (c), as colunas "extraído" e "declarado" estão invertidas (N12). |
| M3 Faixas de Rafael | **Corrigido** | São 8 linhas. Entra a frase da p. 9 "gira em torno de R$ 5.600 a R$ 6.200". O FIPE ficou só dentro da frase da p. 15 e o "R$" solto foi removido. Conferi no PDF (p. 8, 9 e 15): não falta nenhuma faixa. |
| M4 Planilha Google | **Corrigido** | Entraram GERAL 227–241, 882 e 986–987, Resumo 73 e 78–94, 11.Empreendimentos 90 e 731, e Acertos e Acordos 36–43. Os comentários B230, B233, B235, B237 e B238 foram capturados. O aviso do Resumo B73 está certo (C73 = GERAL!C8, caminhão-tanque). Há ruído na marcação "É observação?" (N10). |
| M5 Diff do docx × doc 17 | **Corrigido** | A aba "Vendas docx x doc 17" tem 5 diferenças: Rigo +118.100 SC, +37.500 sacas e 2 × "(FALTA RECEBER)"; Angeli 40.000 → 50.000. Conferi no texto: são exatamente essas. |
| M6 Notas de Rafael | **Corrigido em parte** | A aba "Rafael notas" tem as colunas item e "itens citados". As páginas conferem com o PDF: p. 10 → 02, p. 11 → 03, p. 12 → 04, p. 14 → 06 (cita 02–05), p. 15 → 07, e p. 13 → 05 (avaliação independente). "Grupo x Rafael" marca 01 e 08 sem concordância e traz a nota de 01. Falta: a coluna "Concordância expressa?" diz "não" para 05, sem indicar que as p. 14–15 citam o 05. |
| M7 Sobreposição SEI 2018 | **Corrigido** (resta uma nota imprecisa) | Há a aba "SEI 2018 outras empresas" (6 linhas, R$ 5.064.791,80), a coluna "Também listado pela SEI em 2018?" e a aba "2018 → 2025 por matrícula". A frase "contado no benchmark 2018 das duas empresas" não vale para todas as linhas (N9). |
| B1 Totais 2018 | **Corrigido** | Ficaram 5 linhas, sem capital social e sem duplicatas (as refs foram unidas: T5R6, T7R1 e T6R4, T7R2). |
| B2 Resumos de capital | **Corrigido em parte** | Há a coluna "Tipo" (cota × total) e a aba "Menções" com o parágrafo de Rovílio. O valor malformado "100.874,027,60" não foi tratado (sem impacto em 02/05). Rótulo novo errado (N11). |
| B3 Aritmética interna | **Corrigido** (uma conferência é trivial) | Os componentes de 01, 03 e 04 (2025) e de 01 e 04 (2018), a regularização e as 6 cotas foram conferidos, e há aviso "25% (Vinte por cento)". Recalculei à mão: tudo fecha. A conferência "2018 item 04: base da obra" passa por construção (N5). |
| B4 Lógica frágil | **Corrigido em parte** | Corrigido: `seq <= 8` (agora `bloco_fazenda`), o pareamento antigo × atual (agora por número de item), a busca no Diretrizes (com limite de dígito e corte no Araçá), `area_ha` "20,5 ha (13,0 + 7,5)" e os rótulos da conferência de observações. Continua igual: o marcador de item só aceita 1–8 (`0?([1-8])\.?`), os valores só são captados em linha isolada "R$ x", e o diff de vendas volta a parear por posição (`zip(v18, va)`). |
| B5 Detalhes 2018 | **Corrigido em parte** | Há avisos para Garbin (44,16 × 44,1) e para a mat. 11.090. Falta capturar a compra do item 08 de 2018 "pelo valor de R$ 50.000,00" (T5R3), e "Adquirido de" continua vazio na seq. 6 (Garbin). |
| B6 MD5 | **Corrigido** | Os 37 hashes conferem com os arquivos atuais. `achados_acertos` = 0a0377229ad56ad7b1f0098288b0fbc5. |

**Nada do que estava certo ficou errado.** Conferi de novo:
- **Totais:** 2025 atual 75.410.201,00 e antigo 83.640.801,00 (diferenças só nos itens 05, −8.150.600, e 06, −80.000; texto do 07 mudou). 2018 Hotel 34.404.877,83. 2018 Agropecuária: subtotal da fazenda 22.775.090, bruto 24.112.390, recebido 12.058.000, líquido 12.054.390.
- **Rafael, por seção:**
  - grupo = 75.410.201;
  - avaliação executada: 01 = 13.504.243,50, 08 = 54.000;
  - executado: 01 = 17.555.000, 08 = 85.000 (2 vezes), total 79.491.957,50;
  - quadro: 1/6 = 13.248.659,583, escolhidos 50.914.157,50, diferença 37.665.498;
  - escolhidos: 04, 05, 06 e 08.
- **Lista e Diretrizes:** iguais aos escolhidos.
- **Componentes do item 01 no balanço executado (p. 16–17):** os 10 valores do grupo aparecem uma vez cada.

## 2. Achados novos

### N1 — MÉDIA — Lote 12 (seq. 10 de 2018): a área cai de 20,5 ha para 13,0 ha na correspondência 2018 → 2025, sem nenhuma sinalização
- **Evidência:** o doc 17, T4R6, diz "13,0 hectares + 7,5 hectares … Matrícula nº 3.062 = 20,5 hectares X 11.000,00" (R$ 225.500). Em 2025, a mat. 3.062 aparece só no doc 11, T8R5, com "13,0 hectares X 40.000,00" (R$ 520.000). Os 7,5 ha batem em área e vendedor com o doc 11, T8R4: lote 14, mat. **4.194**, "FAMÍLIA SANGALI – VALDIR SANGALI", 7,5 ha, R$ 220.000. A planilha 2_0 trata os dois lotes na SEI já em 2018: GERAL C229 (lote 14) = R$ 280.000 e C230 (lote 12) = R$ 143.000 = 13 × 11.000, com o comentário B230 "DESCRIÇÃO ANTIGA". **Possível dupla contagem dos 7,5 ha em 2018:** dentro dos 20,5 ha da Agropecuária e nos R$ 280.000 da SEI.
- **Por que passa:** a busca é só por número de matrícula, e a coluna "Texto em 2025" corta em 160 caracteres, antes da área e do valor.
- **Correção sugerida:** acrescentar as colunas "área 2025" e "valor 2025" e uma conferência "área 2018 × área 2025 da mesma matrícula". Registrar a hipótese lote 12 + 7,5 ha × mat. 4.194 como aviso, para decisão em P4.

### N2 — MÉDIA — Mat. 11.090: a parcela de 106,53 ha **vendida** aos Rigo aparece "em 2025" no item 06 do Hotel
- **Evidência:** a aba "2018 → 2025 por matrícula", seq. 2 (Alzir, 106,53 ha, VENDIDA, RIGO), mostra "05 T5R4" com o texto do item 06 (4,0 ha, Dirceu/Lairton). A correspondência é só de número: a mesma matrícula cobre as duas parcelas (B5). Quem lê a aba sozinha entende que o bem vendido continua no Hotel. O aviso "Matrícula em mais de uma parcela" está em outra aba.
- **Correção sugerida:** quando a situação de 2018 for VENDIDA e a matrícula também tiver uma parcela não vendida, escrever na própria linha "matrícula compartilhada; a parcela vendida não está em 2025". Ou casar também por área e vendedor.

### N3 — MÉDIA — O texto da venda de 2.606 sacas (Rigo) é cortado em "PARA"; perde-se o destino, a escavadeira hidráulica, que está na SEI
- **Evidência:** no docx, P12 termina em "…VENDIDO EM 2024 A QUANTIA DE 2.606 SACAS PARA" e P13 continua "AQUISIÇÃO DA ESCAVADEIRA HIDRÁULICA". Na aba "Créditos de soja", as duas linhas de 2021 terminam em "…SACAS PARA". O código `line = f"{prev} {p}" if prev.endswith("PARA")` é código morto: só seria usado se o parágrafo seguinte (P13) casasse com algum padrão, e não casa.
- **Relevância:** o doc 11 de 2025, T5R2, traz "Uma Escavadeira Hidráulica, ano 2024, modelo 320, CATERPILLAR", R$ 887.000. A soja do Hotel (crédito Rigo) pode ter pago um bem que hoje está no inventário da SEI. É um fluxo entre empresas, a registrar para P4 e para a sessão SEI.
- **Correção sugerida:** juntar ao parágrafo anterior o parágrafo de continuação, que não casa com nenhum padrão. Criar o aviso "2.606 sacas → escavadeira 2024 (SEI, doc 11 T5R2?)", com o vínculo marcado como hipótese.

### N4 — MÉDIA — "FALTA RECEBER" (R$ 3.224.000) tratado como certo; o próprio docx sugere soja recebida e ainda não vendida
- **Evidência:** a tabela T1R1 marca as parcelas de 2021 e 2022 "(FALTA RECEBER)". Mas o título das OBSERVAÇÕES (P10) diz "O SOJA DE 02 (DUAS) SAFRAS DA FAMÍLIA RIGO AINDA NÃO FORAM **VENDIDOS**", e P12 registra que 2.606 sacas da safra 2021 foram **vendidas** em 2024. Isso só é possível se a soja estava à disposição do grupo. Logo:
  - o saldo atual é de 37.694 sacas, não 40.300;
  - o R$ 80/saca é a cotação de outubro/2018 (T1R1), não a de 31/12/2024.
- A conferência "parcelas Rigo marcadas FALTA RECEBER incluídas no TOTAL RECEBIDO" e a nota da aba "Base 2018" levam ao Luís uma leitura (crédito contra os Rigo) que o documento não sustenta sozinho.
- **Correção sugerida:** reescrever a nota assim: "marcadas FALTA RECEBER na tabela; as OBSERVAÇÕES dizem 'não vendidos' e registram venda parcial em 2024 — pode ser soja em depósito, não crédito contra os Rigo; saldo declarado 37.694 sacas; R$ 80 = cotação 10/2018". Levar a dúvida ao Luís em vez de afirmar.

### N5 — BAIXA — "2018 item 04: base da obra = pavilhão + casa + quiosque" passa por construção
- **Evidência:** em `component_checks`, quando o documento não imprime a base (2018), o código faz `base = parts["pavilhão"] + parts["casa"] + parts.get("quiosque", 0)` e depois compara essa mesma soma com `base` (L818–828). A conferência informativa é a seguinte (30% × 814.686 = 244.405,80), que está OK.
- **Correção sugerida:** em 2018, não emitir a conferência da base, ou registrá-la como "base inferida".

### N6 — BAIXA — "Redução de R$ 400.000 fora do TOTAL GERAL" repete a 1ª conferência e não testa a redução
- **Evidência:** compara `sum(itens 05)` com `total(t25_05)`, igual a "2025 atual 05: soma dos itens = TOTAL GERAL". A redução lida do texto (`reduction`) só aparece na nota. Também não se confere que a "redução de R$ 400.000" (P77) é igual ao "TOTAL de R$ 400.000" das despesas (P56).
- **Correção sugerida:** conferir `reduction == total` (despesas) e registrar a ausência de abatimento como aviso, não como conferência OK.

### N7 — BAIXA — A conferência "Base planilha − base líquida = ganho − trator − lotes" é uma identidade
- **Evidência:** com `kept == C28` (conferência anterior), a igualdade se reduz a "soma das 17 linhas do doc 17 = TOTAL GERAL", que já é conferida. Não acrescenta prova. A composição de `kept` foi montada para reproduzir a C28; a 1ª conferência mostra que é **uma** composição que fecha, não que a planilha usou essa. A premissa "83,47 ha a R$ 44.000 (parte produtiva)" vem da C944 (6.835.730) e não está escrita no rótulo.
- **Correção sugerida:** rotular a conciliação como "decomposição que reproduz C28", escrever a premissa dos R$ 44.000/ha e deixar só a 1ª conferência.

### N8 — BAIXA — "Regularização: texto idêntico nos docs 02 e 05" esconde a única diferença de redação
- **Evidência:** antes de comparar, o código troca "para a empresa" por "em nome da empresa" no doc 05. Os documentos diferem exatamente nisso: doc 05 P36 "transmitir … **para a** empresa AGROPECUÁRIA"; doc 02 P26 "**em nome da** empresa AGROPECUÁRIA". O resultado mostra OK, e a nota diz "se divergir, ver a diferença".
- **Correção sugerida:** escrever na nota a diferença que foi normalizada. As âncoras `find()` também devem falhar alto se faltarem: hoje, -1 nos dois daria strings iguais e um OK trivial.

### N9 — BAIXA — Nota da conferência SEI 2018: "contado no benchmark 2018 das duas empresas" é impreciso
- **Evidência:**
  - Angeli (408.000) foi vendido em 2018. Na base A, o preço de venda (510.000) o anula. Na base B, ele fica de fora, assim como os lotes 03/04/05 (551.200).
  - Na base B, a sobreposição com a SEI é só lote 52 + Sede Campestre = 226.500 + 3.879.091,80 = **4.105.591,80**, não 5.064.791,80.
  - O título da aba da SEI é "Bens da Siviero Empreendimentos relacionados nos inventários": é uma lista, sem prova de que entra no total da SEI.
- **Correção sugerida:** mostrar a sobreposição por base (A e B) e trocar "contado" por "listado".

### N10 — BAIXA — "Planilhas Google": a marcação "É observação?" infla a contagem (67) e a conferência dá DIVERGE informativo
- **Evidência:**
  - Na aba Resumo, a partir da linha 78, qualquer texto fora da coluna A é marcado como observação, inclusive rótulos de tabela: B78 "FAZENDA SÃO CRISPIN", C88/D88 "anual"/"média anual", C81/D81 "R$ 44.000,00/ha".
  - As linhas GERAL 227–229, 231–233 e 236–240 são da SEI, sem relação com 02/05. Os comentários B233, B237 e B238 e as notas H233, H237, H238 e H240 entram na contagem.
  - O arquivo da SEI não tem nenhuma linha para as abas 2.São Crispin, 5.Hotel e Extras, então o "5" está certo. Mas a conferência aparece como DIVERGE numa aba cujo cabeçalho diz "DIVERGE = achado a tratar".
- **Correção sugerida:** restringir a marcação a células de observação de fato e mover a contagem para a aba Avisos, como informação.

### N11 — BAIXA — "Resumos de capital": as linhas TOTAL do grupo estão rotuladas como "total da empresa"
- **Evidência:** doc 00 2025 T2R13 e T3R13 (263.410.451,80), antigo T2R13 "PARCIAL TOTAL" (174.480.661,00) e doc 27 T1R14 (225.163.796,31) têm Tipo = "total da empresa". São totais do grupo.
- **Correção sugerida:** criar o tipo "total do grupo".

### N12 — BAIXA — Números e formatação digitados em rótulos e notas; colunas invertidas numa conferência
- **Números digitados em vez de lidos** (não afetam os valores, mas contrariam a docstring "every number comes from the documents"):
  - nota de "Rafael executado item 01": "17.555.000,00";
  - nota da conferência FALTA RECEBER: "R$ 12.058.000,00";
  - nota da soja da Agropecuária: "R$ 0,00";
  - rótulos "411,4 ha × R$ 30.000", "100.000 + ITBI …", "Redução de R$ 400.000", "(48 + 100 ha)";
  - nota da aba Base 2018: "R$ 100.000 de gado";
  - na aba "Grupo x Rafael", a nota fixa "componentes impressos = os do grupo" não depende do resultado da conferência.
- **Formatação americana num entregável em português:** "R$ 400,000.00", "R$ 5,064,791.80", "25.00%", "100.0 ha × 110000.00".
- **Colunas invertidas:** em "2025: área declarada no texto × itens 05 + 06", o valor declarado (411,4) está em "Extraído" e o calculado (411,53) em "Declarado".
- **Correção sugerida:** montar rótulos e notas a partir das variáveis, formatar no padrão brasileiro e trocar a ordem dos argumentos.

### N13 — BAIXA — "2018 → 2025 por matrícula" busca só nos inventários atuais: o lote 52 no "ANTIGO 11" não aparece
- **Evidência:** o ANTIGO 11, T9R3, lista a mat. 7.391 (lote 52, R$ 15.000/ha) na SEI. No atual, a matrícula está no Hotel, item 07 (R$ 28.000/ha, 422.800). A mudança da SEI para o Hotel entre as versões de 2025 confirma a GERAL A.3 (C241 = 226.500 → E241 = 0), mas não fica registrada.
- **Correção sugerida:** incluir os inventários antigos na busca, com a coluna "versão", ou registrar um aviso.

## 3. Resumo

Todos os 15 achados da rodada 1 foram tratados: 10 estão corrigidos e 5 corrigidos em parte (M6, B2, B4, B5 e, por um detalhe, A2). Não há regressão nos totais, nos valores por item nem nos valores de Rafael por seção.

Achados novos:
- **4 médios**, que são de interpretação ou de correspondência e não de leitura:
  - N1: lote 12, 20,5 → 13,0 ha, com possível dupla contagem dos 7,5 ha;
  - N2: parcela vendida da mat. 11.090 associada ao item 06;
  - N3: as 2.606 sacas e a escavadeira da SEI;
  - N4: "FALTA RECEBER" × "não vendidos".
- **9 baixos**, entre eles duas conferências que passam por construção (N5, N7) e uma redundante (N6).
