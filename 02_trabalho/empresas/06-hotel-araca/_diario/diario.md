# Diário de trabalho — empresa 06 Hotel Araçá (uso interno de Claude e do revisor)

Somente acréscimos. Cada entrada: data, passo, entradas (MD5 10 dígitos), saídas, contagens e conferências, pendências. Para retomar: ler o brief `02_trabalho/empresas/_orquestracao/prompts/S2-06-hotel-araca.md`, `00_PLANO.md`, o fim deste diário e `log-decisoes.md` desta pasta; números sempre re-derivados das fontes por `_scripts/fase0_extracao_06.py` (e scripts seguintes), nunca da memória.

## 2026-10-28 — Fase 0 (extração)

### Leituras de método
- Brief S2-06 inteiro; `00_PLANO.md` (P1–P13, seção 7); `02_trabalho/log-decisoes.md`; `02_trabalho/_diario/diario.md` (até a entrada "sessão SEI 2 — levantamento ambiental refeito"); `_orquestracao/analise-coordenacao.md`; memória do projeto.
- Pasta de notas da família de sessões: vazia no início.

### Entradas (md5 10 dígitos, calculado pelo script)
- inv2025 `01_fontes/inventario-2025/06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx` `4371e018c7`
- inv2025_antigo `01_fontes/inventario-2025-antigo/ANTIGO 06. …docx` `4d33595ff0`
- inv2018 (Drive) `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\16. 1. HOTEL ARAÇÁ LTDA. = I N V E N T Á R I O = … Dezembro de 2018.docx` `9dcf78ee97`
- resumo2018 (Drive) `…\inventário-2018\27. 1. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - RESUMO - Dezembro de 2018.docx` `af5ec6c22e` (só hash; ainda não lido)
- doc00_2025 `01_fontes/inventario-2025/00. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - 31.12.2024.docx` `a1670b10fe`
- rafael_pdf `01_fontes/estudo-rafael/HOTEL ARAÇÁ LTDA - BELISE LTDA.pdf` `f844ffc907` (26 p., todas com texto)
- diretrizes_pdf `01_fontes/estudo-rafael/DIRETRIZES PARA SELEÇÃO DE ATIVOS … BELISE LTDA.pdf` `aff2b08f2a`
- lista_rafael `01_fontes/outros/ativos-avaliados-valter.xlsx` `935282c1aa`
- gsheet_2_0 `01_fontes/planilhas-google/2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx` `11a81ef0aa`
- gsheet_rafael `01_fontes/planilhas-google/Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx` `c499486a26`
- observacoes `02_trabalho/observacoes-planilhas-google.xlsx` `18bf9de495`
- conferencia_md `01_fontes/outros/conferencia-discrepancias.md` `bdd7c87e70`; conferencia_xlsx `01_fontes/outros/inventario-2025-avaliacoes-corrigido.xlsx` `a565450170`
- pendencias_md `02_trabalho/pendencias-conferencia-anterior.md` `1e9eca2c47`; diff_md `02_trabalho/diff-inventario-2025-antigo-x-atual.md` `7afde86365`
- resp_arq1 `01_fontes/respostas-docs-2026-08-13/1 - grupo siviero - encaminhamento esclarecimentos ao Dr. Pablo - 13.08.2026.pdf` `03968eb843` (32 p., com texto)
- resp_arq12 `…/12 - OUTORGA_POCO_HOTEL_ARACA.pdf` `a8b7aa7ef5` (1 p., escaneada, lida como imagem)
- resp_arq17 `…/17 - … LOTES VENDIDOS DE JANEIRO DE 2018 A OUTUBRO DE 2025 - REV.13.08.2026.xlsx` `7ef0dc35dc`

### Saídas
- `_scripts/dump_docx.py`, `_scripts/dump_pdf.py` (texto bruto, só para leitura) → `_extracoes/inv2025.txt`, `inv2025-antigo.txt`, `inv2018.txt`, `rafael-pdf.txt`, `diretrizes-pdf.txt`, `resp-arq1.txt`, `resp-arq12.txt` (vazio: escaneado), `doc00.txt`.
- `_scripts/fase0_extracao_06.py` → `_extracoes/fase0-06.xlsx` md5 `b19887005d` (abas Fontes, Conferencias, Itens, Item01_componentes, Rafael_itens, Rafael_textos, Lista_Rafael, Planilhas_Google).

### Contagens e conferências (saída do script)
- 2018: 25 itens; 19 com valor; soma 49.684.243,20 = TOTAL do documento (que acrescenta "+ o valor da casa e lote – Élcio Siviero"). Itens 20–25 com nome de sócio no lugar do valor.
- 2025 atual: 22 itens, soma 58.076.823,20 = TOTAL. 2025 antigo: 22 itens, soma 61.790.460,00 = TOTAL. Só o item 01 muda (53.320.460,00 → 49.606.823,20; vantagem 25 % → 15 % sobre 37.136.368,00).
- Componentes do item 01 somam o valor do item nas três versões (2018 45.274.243,20; antigo 53.320.460,00; atual 49.606.823,20).
- Rafael, transcrição do balanço do grupo (p. 1–5): 22 itens = inventário 2025, item a item (pareados por inscrição e matrícula).
- Rafael, balanço executado (p. 16–20): item 01 = 66.000.000,00 (p. 17); item 20 = 2.400.000,00 (p. 19); os outros 20 lotes iguais ao grupo. Soma 75.870.000,00 × total declarado 76.230.000,00 (p. 20, impresso "R$76.230.000,0") → DIVERGE 360.000,00. 1/6 de 76.230.000 = 12.705.000 (p. 21) confere; crédito 12.705.000 − 1.880.000 = 10.825.000 confere: o total errado propaga para o 1/6 e o crédito.
- Lista de Rafael: 5 itens (02, 11, 12, 18, 19), soma 1.880.000 = "lotes escolhidos" p. 21–22; Diretrizes p. 6–7.
- GERAL C70 49.484.243,20 × 2018 49.684.243,20 → DIVERGE 200.000 (linha 99, lote 28/q38, 2018 = 0; era 200.000 no item 18 de 2018). GERAL C80 lotes 4.210.000 × 4.410.000, mesma causa. E70 out/2025 confere.

### Achados de leitura (a confirmar nas fases seguintes)
- Rafael p. 7: R$ 66.000.000 é o "Valor Mínimo Físico" (terreno R$ 5.000/m² e 10.500 m² a R$ 3.000/m²); faixa "valor médio de mercado" 55–85 milhões; máximo físico 102.450.000; renda 41,0–50,6 milhões. Pela conta 6.900 × 5.000 = 34.500.000 e 10.500 × 3.000 = 31.500.000, soma 66.000.000: aparentemente só terreno + construção, sem móveis/equipamentos (10.504.000 no grupo) nem vantagem (5.570.455,20). Rafael repete no balanço executado o texto do grupo com as parcelas do grupo e põe 66 milhões ao lado.
- Rafael p. 10–11: lotes das quadras 37 e 38 entre 380.000 e 550.000 (meio de quadra perto do piso 380.000; esquina 550–600 mil). Mesmo assim adota 360.000 nos itens 03–10 (abaixo do piso dele) e "concorda" com os itens 02–19 (p. 12–13 e 16).
- Rafael p. 13–15: item 20 faixa 2.400.000–3.500.000 (terra 1,81–2,75 milhões; casa 600–750 mil); adota o piso.
- Rafael p. 24–25: duas residências de ~200 m² na esquina da Rua da Divisa com a Rua Três, faixa 2,4–3,5 milhões pelo par; **e uma Fiat Toro, automóvel de serviço do Hotel** (novo: não estava no brief). Não somadas ao total (a soma dos itens fecha 75.870.000 sem elas).
- Hipótese (não confirmada): as duas residências são os itens 22 (Raul, lote 360.0057.002, casa) e 23 (Élcio, 360.0057.001, casa/esquina) de 2018 — dois lotes em sequência, um de esquina. O 23 tem anotado "trocou este lote/casa por um lote no Loteamento Aditivo Medianeira". Endereço não consta do inventário 2018.
- Pares 2018 → 2025 por inscrição imobiliária: 2018 02–18 → 2025 02–16, 18, 19; 2018 19 → 2025 20; 2018 21 (IRINEU) → 2025 17 (lote 15/q38, mat. 33.003, "Adquirido do Senhor Hugo Ferri"). Sem par em 2025: 2018 20 (VÁLTER, 360.0019.004.000), 22 (RAUL), 23 (ELCIO), 24 (ADÉLIA, mat. 77.833), 25 (PEDRO, mat. 74.857). Sem par em 2018: 2025 21 (mat. 86.669) e 22 (mat. 124.906, Osório).
- Item 01: área 6.900 = 4×450 + 13×360 + 420 (nesga). Lotes 13 e 14 da quadra 37 não aparecem em nenhum item (conferência anterior, linha 92).
- Matrícula 73.482 nos lotes 27 e 28/q38 nas três versões e na lista de Rafael ("Documento indica a mesma matrícula do item 19"). Inscrições e códigos distintos.
- Item 20: "Contrato Particular de Compra e Venda" (maio/2014, R$ 750.000 quitado), mas com matrícula 69.488 citada. Grupo afirma (arq. 1, p. 11) que "todos os imóveis" estão registrados em nome da empresa.
- Arq. 1 p. 10–11 (Anexo E): composição do item 01 confere com o inventário; sem projetos de regularização, sem habite-se, sem profissional contratado; INSS ≈ 1.601.357,90 "estimativa baseada em tabela de valores da receita federal"; PPCI aprovado e em vigor; justificativa dos 15 %; poço com outorga mas "não está sendo utilizado para abastecimento e consumo".
- Arq. 12: Portaria DRH nº 1398/2012, 06/09/2012, poço tipo ponteira, Rua 2 nº 330, usos "sistema de prevenção de incêndio, jardinagem e piscina de uso coletivo", 20 m³/dia, caráter precário. Coerente com a resposta do grupo. O poço entra no texto dos "equipamentos diversos" (6.400.000) nas três versões, sem valor separado.
- Data-base: arq. 1 p. 1–2 fala em "termo final … outubro/2025". Os textos do inventário 06 são datados de 31/12/2024 e não citam fatos posteriores; a coluna "outubro 2025" das planilhas Google = inventário atual. Nenhum número do Hotel Araçá muda por isso até aqui.
- Rovílio: nenhuma menção nos inventários 06 nem no PDF de Rafael desta empresa.
- Cruzamento com a SEI: placas solares do Hotel Araçá pagas pela SEI (doc 11, obs. 14: Solar Serra – Mário Bordignon, compromisso de 1 lote em Veranópolis, R$ 150.000) e arq. 17 linha 60: lote 02/q07 do Aditivo Medianeira vendido a Solar Serra em 24/04/2024 "como forma de pagamento placas solares Hotel Araçá". Candidato à Parte 6 (pagamento em duplicidade? despesa do Hotel paga pela SEI sem registro no Hotel).
- Troca do item 23 de 2018 (Élcio) por lote no Aditivo Medianeira: não achei no arq. 17 por nome do sócio. Pendência.
- Planilha "empreendimentos e Rafael", aba 6.Hotel Araçá, coluna "Estudo Rafael": só os 5 lotes escolhidos (1.880.000); não traz os 66 milhões nem os 2,4 milhões.
- Observações do Luís (observacoes-planilhas-google.xlsx): nenhuma nas linhas 70–104 da GERAL nem em aba do Hotel Araçá.

### Pendências
- Revisor independente da Fase 0 (P11) antes de apresentar ao Luís.
- Ler o resumo 2018 (doc 27) e o doc 00 2025 no que toca ao Hotel Araçá (linha do doc 00: 58.076.823,20 e 9.679.470,53 por sócio, conferida em texto bruto).

## 2026-10-28 — revisor independente da Fase 0 e correções
- Revisor (contexto novo, fontes + script + xlsx, sem conclusões) → `_diario/revisao-fase0.md`. 0 altos, 10 médios, 8 baixos. Nenhum número extraído errado nem bem perdido; pareamento do balanço executado conferido por ele nas imagens das p. 17–20.
- Correções no script `fase0_extracao_06.py`:
  - M2: `check()` exige os dois lados (assert); lista `EXPECTED_DIVERGE` com as 9 divergências documentadas; qualquer divergência nova ou desaparecida para o script.
  - M3: `balance_section()` ancora cada valor ao bloco do lote ("Um lote urbano" … próximo), 1 matrícula + 1 inscrição + 1 valor por bloco; exclui "valor total global de R$ 750.000,00" pelo texto; valor e página do item 01 lidos do texto. Nova conferência por lote "Rafael executado = grupo" (divergem só itens 01 e 20) e item 20 achado pela inscrição 645.0027.015.000.
  - M1: aba `Rafael_faixas` estruturada (página, bem, tipo, mínimo, máximo, trecho), p. 5–25; `Rafael_textos` refeita por frase (corte só em ponto seguido de espaço e maiúscula), p. 5–26.
  - M6: aba `Observacoes_inventario` (parágrafos das três versões); conferências INSS 1.601.357,90 e redução 2.000.000,00.
  - M7: `Item01_componentes` com área, valor unitário, base CUB, depreciação e % de vantagem.
  - B4: linha da fonte e subtotais "Lotes" na aba `Planilhas_Google`; rótulos de coluna F e "Estudo Rafael" explicados. B5: colunas `matricula_antiga`, `comarca`, `inscricao_norm`, código com 6 dígitos. B6: conferências doc 27 (2018) e doc 00 (2025).
  - M4: novas conferências GERAL D70 (56.576.823,20 × 56.926.823,20), G70 e aba H2 (1.500.000 × 1.150.000), aba C2, F2.
- Saída: `_extracoes/fase0-06.xlsx` md5 `c1963c6bf8`; 70 conferências, 9 divergências = as documentadas.
- Resumo `pacotes/fase0-resumo-06.md` reescrito com M4, M5, M7–M10, B1–B3, B7 e os números de Rafael corrigidos (1/6 12.645.000; crédito 10.765.000).
- Não fiz: o bloco repetido do Hotel Araçá na aba 7.Rádio da planilha "empreendimentos e Rafael" e a fórmula F3 (=SUM(F4:F18)) só estão registrados, não extraídos.
- Próximo: apresentar a Fase 0 ao Luís; esperar OK para a planilha-mestre.

## 2026-10-28 — Fase 0 aprovada; planilha-mestre gerada
- Luís aprovou a Fase 0 no chat ("Aprovo, siga"). Registrado em `log-decisoes.md` e `decisoes.json` (aprovacoes; sem decisões de valor).
- Novo script `_scripts/build_mestre_06.py` (importa os parsers de `fase0_extracao_06.py`; lê decisões de `decisoes.json`) → `planilha-mestre-06.xlsx`, abas Mestre, Item01_componentes, Totais, Pares_2018_2025, Observacoes.
- Saída do script: 29 linhas (22 de 2025; 5 só 2018: itens 20, 22, 23, 24, 25 de 2018; 2 só Rafael: residências da Rua da Divisa e Fiat Toro). Concorda 20 itens = 7.470.000,00; discorda 2 (01 e 20): grupo 50.606.823,20 × Rafael 68.400.000,00; erro de soma de Rafael 360.000,00. Pares: item 01 + 19 lotes por inscrição normalizada (matrícula conferida quando os dois anos têm); sem par 2025: 21, 22.
- Item 01, Rafael derivado da p. 7: terreno 6.900 × 5.000 = 34.500.000; construção 10.500 × 3.000 = 31.500.000; soma = 66.000.000 (assert). Diferença por componente: terreno +27.600.000; obra civil +4.867.632; móveis/equipamentos −10.504.000; vantagem −5.570.455,20.
- Revisor independente da planilha-mestre disparado → `_diario/revisao-mestre.md`. Não apresentar ao Luís antes.

## 2026-10-28 — aviso da coordenação: OCR revisado dos arquivos 2–16
- Fonte (só leitura, da sessão SEI): `02_trabalho/ocr-respostas-docs/ocr.jsonl` md5 `9b5b52b7c1`; revisão em `02_trabalho/_diario/revisao-ocr-respostas-docs.md`. Uso: localizar trechos; números conferidos na imagem.
- Busca por araçá/capão/condado/divisa/fiat toro/hugo ferri/bigode: só o arquivo 12 (outorga, já lido como imagem). Os demais acertos são falsos ("desembaraçado"). Nenhuma menção ao Hotel Araçá nos arquivos 2–11 e 13–16, inclusive no contrato de Rovílio (arquivo 10).

## 2026-10-28 — revisor da planilha-mestre e correções
- Revisor → `_diario/revisao-mestre.md`: 1 alto, 8 médios, 13 baixos. Nenhum valor, par ou conciliação errado.
- Conferido por mim antes de corrigir: doc 25 de 2018 (md5 `07fe4f03f4`), seção Élcio: lote 10/q03 Aditivo Madianeira dado a Paulo Busatto, R$ 70.000, "trocado com o lote de Capão da Canoa – CASA". Doc 6 de 2018 (md5 `bc34729945`) é a relação de imóveis para IPTU de Válter e família e lista 360.0019.004.000 (Capão, R$ 500,00, valor do IPTU). Doc 1 (Raul, `362444641c`), doc 4 (Pedro, `a4d6c9c61e`), doc 5 (José Valério, `a1ebb4d7aa`): um imóvel em Capão sem inscrição cada. Doc 3 (Irineu, `ff9ed49fd3`): 360.0058.001.000 e "Circo", não 360.0038.015.000. Doc 25: apartamentos do Ed. Contemporâneo pagos pelo Hotel e pela SEI (Adélia 680.000; Raul 400.000).
- Correções em `build_mestre_06.py`:
  - A1: linha "Placas solares" (Não abordado), lida do inventário 11 e do arquivo 17 (linha 60).
  - M1: rótulo "Diferença de soma de Rafael (não identificada)", com nota de que 360.000 = valor de cada item 03–10.
  - M2: componentes 1.2, 1.3, 1.5 e 1.6 de Rafael ficam vazios ("não avaliado em separado"); decisão no nível do item.
  - M3, M4, M5: notas sobre área construída, nesga e lotes 13/14; linha própria para os lotes 13/14.
  - M6: notas do doc 25 e do doc 6, ressalva do total de 2018.
  - M7: `registro-data-base.md` criado (3 registros) e coluna ligada nos itens 01 e 20 e nas residências.
  - M8: lista de Rafael ligada pela inscrição, com assert de item, matrícula e valor.
  - B1: as 12 matrículas comparadas nas três versões. B2: 1/6 e crédito pela soma. B3, B5: textos. B4: código de imóvel na chave e nos pares (assert de código igual nos pares). B6: observações lidas do arquivo. B7: coluna da faixa de Rafael. B8: concordância expressa × por valor. B9: base da vantagem. B10: aritmética da renda de Rafael (lucro 3.587.220/ano; 34,16 mi e 42,20 mi). B11: afirmação do grupo no item 20. B12: notas por inscrição e contagens calculadas.
  - B13: em `parte6-candidatos.md`.
- Criados `parte6-candidatos.md` (4 linhas) e `perguntas-ao-grupo.md` (10 perguntas propostas, não enviadas).
- Saída: 31 linhas (20 concorda, 2 discorda, 5 só 2018, 2 só Rafael, 2 não abordados). Conciliações sem mudança: 7.470.000 + 50.606.823,20 = 58.076.823,20; 17.793.176,80 + 360.000 = 18.153.176,80.

## 2026-10-28 — planilha-mestre aprovada; Pacote 1 gerado
- Luís aprovou a planilha-mestre ("Aprovo, siga"). Registrado em `log-decisoes.md` e `decisoes.json` (aprovacoes).
- `build_mestre_06.py`: a faixa de lotes de Rafael (quadras 37 e 38, p. 10) só se aplica a itens com quadra 37 ou 38 no texto (o item 02, q19, ficou sem faixa).
- Novo `_scripts/pacote_01.py` (reconstrói a mestre e lê dela) → `pacotes/pacote-01-regras.md`.
  - R1: 17 itens (02–16, 18, 19) = 6.320.000,00.
  - R2: 3 itens (17, 21, 22) = 1.150.000,00.
  - R3: 9 linhas sem valor.
  - Fora das regras: 01 e 20 = 50.606.823,20.
  - Fecha 6.320.000 + 1.150.000 + 50.606.823,20 = 58.076.823,20 (assert).
- Revisor independente do Pacote 1 disparado → `_diario/revisao-pacote-01.md`. Não apresentar antes.

## 2026-10-28 — revisor do Pacote 1 e correções
- Revisor → `_diario/revisao-pacote-01.md`: 1 alto, 6 médios, 9 baixos. Valores, somas, contagens, páginas, pares e amostra conferem.
- A1: a ressalva "abaixo da faixa" foi separada. Itens 03–10 ficam na empresa (C0 pediria o maior; não há maior, P5). Itens 18–19 são da lista de Rafael: abaixo do piso, o que favorece o C0 (argumento C1).
- M1: notas por m² na mestre (faixa de Rafael R$ 1.050–1.500/m², p. 11). Itens 05–09 a 1.169,59/m², dentro da faixa. Efeito máximo informativo: 8 × 20.000 = 160.000 (26.666,67 em 1/6).
- M2: pergunta G11 (item 17, lote Irineu).
- M3: residências da Rua da Divisa vão ao pacote de exceções como "só Rafael".
- M4: texto sobre os 360 mil = "não identificada".
- M5: frase "aprova-se o valor, não a titularidade".
- M6: G10 inclui a data de aquisição.
- B1: perguntas renumeradas G1–G11. B2: p. 15 e 20. B3: amostra 02, 05, 13, 18, 17, 22.
- B4: mestre ganhou as colunas "Concordância de Rafael" e "Nome de sócio em 2018", e o pacote classifica por elas. O 360.000 é lido da aba Totais.
- B5: `pacote_01.py` é só leitura e grava o md5 da mestre usada.
- B6: nota do cabeçalho p. 16. B8: lotes do quarteirão. B9: pistas no R3 (vizinho, matrícula sequencial).
- B7 fica para o pacote de exceções: Rafael repete o valor do grupo nos itens 01 (p. 6) e 20 (p. 13) na "avaliação patrimonial executada".
- Mestre: nota "abaixo do piso" agora para todos os itens com faixa, com texto distinto. Concorda: "que mesmo assim o adota". Discorda: "valor do grupo abaixo do piso da faixa de Rafael" (itens 01 e 20).
- Sem nova rodada de revisor: as mudanças são de texto e enquadramento; números sem mudança (R1 6.320.000; R2 1.150.000; fora 50.606.823,20).

## 2026-10-28 — revisor do Pacote 1 e correções
- Revisor → `_diario/revisao-pacote-01.md`: 1 alto, 6 médios, 9 baixos. Valores, somas, contagens, páginas, pares e amostra conferem.
- A1: ressalva "abaixo da faixa" separada. Itens 03–10 ficam na empresa: o C0 pediria o maior, mas não há maior (P5). Itens 18–19 são da lista de Rafael: estar abaixo do piso favorece o C0 (argumento C1).
- M1: notas por m² na mestre (faixa de Rafael R$ 1.050–1.500/m², p. 11). Itens 05–09 a 1.169,59/m², dentro da faixa. Efeito máximo, só informativo: 8 × 20.000 = 160.000 (26.666,67 em 1/6).
- M2: pergunta G11 (item 17, lote Irineu).
- M3: residências da Rua da Divisa vão ao pacote de exceções como "só Rafael".
- M4: diferença de 360 mil descrita como "não identificada".
- M5: frase "aprova-se o valor, não a titularidade".
- M6: G10 inclui a data de aquisição.
- B1: perguntas renumeradas G1–G11. B2: p. 15 e 20. B3: amostra 02, 05, 13, 18, 17, 22.
- B4: mestre ganhou as colunas "Concordância de Rafael" e "Nome de sócio em 2018"; o pacote classifica por elas e lê o 360.000 da aba Totais.
- B5: `pacote_01.py` é só leitura e grava o md5 da mestre usada.
- B6: nota do cabeçalho da p. 16. B8: lotes do quarteirão. B9: pistas no R3 (lote vizinho, matrícula sequencial).
- B7 fica para o pacote de exceções: Rafael repete o valor do grupo nos itens 01 (p. 6) e 20 (p. 13) na "avaliação patrimonial executada".
- Mestre: a nota "abaixo do piso" agora vale para todos os itens com faixa. Nos itens que concordam diz "que mesmo assim o adota"; nos itens 01 e 20, "valor do grupo abaixo do piso da faixa de Rafael". Citação da p. 25–26 corrigida para o texto literal.
- Saídas: mestre md5 `2bfad8d982`; pacote `2d2d13e327`; `pacote_01.py` `c12b6233fa`; `build_mestre_06.py` `7726bf9b0f`; perguntas `caad15d240`.
- Sem nova rodada de revisor: as mudanças são de texto e enquadramento. Números sem mudança: R1 6.320.000,00; R2 1.150.000,00; fora das regras 50.606.823,20.

## 2026-10-28 — Pacote 1 aprovado e aplicado; Pacote 2 gerado
- Luís: "Aprovo R1–R3". `decisoes.json`: 20 decisões (2025-02 a 2025-22, exceto 01 e 20), "Adoto grupo (= Rafael)", com regra R1/R2. Seção `regras.R3` (sem valor). `build_mestre_06.py` aplica R3 às linhas Só 2018 / Só Rafael / Não abordado. `log-decisoes.md` atualizado.
- Mestre regenerada: Totais "Valores adotados" = 7.470.000,00 (20 itens); Decisão: 20 adotadas, 9 sem valor, 2 pendentes (01 e 20).
- Conferido: o PDF de Rafael não tem "INSS", "habite", "regulariz", "descont" ou "dedu" (tema 2 do brief). Rafael não desconta a regularização; o grupo também não deduz do total.
- Novo `_scripts/pacote_02.py` → `pacotes/pacote-02-excecoes.md`. E1 hotel: Rafael 66.000.000 (C0 maior), com ressalvas. E2 residências: sem valor. E3 item 20: Rafael 2.400.000. E4 diferença de soma de 360.000: não adotar.
- Asserts: valor do grupo repetido por Rafael na p. 6 (49.606.823,20) e na p. 13 (1.000.000,00). Soma depois das recomendações = 75.870.000,00 = soma dos itens de Rafael.
- Revisor do Pacote 2 disparado → `_diario/revisao-pacote-02.md`.

## 2026-10-28 — revisor do Pacote 2 e correções
- Revisor → `_diario/revisao-pacote-02.md`: 2 altos, 6 médios, 7 baixos. Números conferem; E1 e E3 seguem o C0.
- Conferido por mim: doc 25 de 2018, seção Hotel Araçá: "manutenção das 02 casas de veraneio dos Sócios … durante o ano todo"; "CASA DA ESQUINA: Élcio Siviero (que pertence à empresa Siviero)"; "CASA AO LADO: Raul Siviero".
- Correções em `pacote_02.py` (reescrito):
  - A1: E1 diz o que entra nos 66 milhões (p. 7: terreno + construção) e que a decisão é pelo item inteiro, com móveis.
  - A2: E4 em duas opções, (a) não adotar e (b) adotar 360.000 (+60.000 no 1/6 do Luís). A diferença em relação ao precedente Oleoplan fica explícita. "Erro de soma" passou a "diferença não identificada".
  - M1: tabela de parcelas com obra civil (+4.867.632; +18 %), vantagem e móveis separados.
  - M2: frases de Rafael sobre a renda (p. 7–8, citadas literalmente), faixa 55–85 sem conta e teto com terreno unificado.
  - M3: versões do grupo (antigo 53.320.460,00, vantagem 25 %; 2018 30 %); lotes +60 % a +100 % × hotel +9,6 %; item 20 parado em 1.000.000.
  - M4: E2 com o trecho do doc 25; destino só Parte 4; G8 refeita; candidato 5 da Parte 6 (manutenção das casas).
  - M5: E5 Fiat Toro. M6: E6 propõe o tratamento dos registros de data-base 1 e 2.
  - B1–B5, B7 no texto. B6: mais valores lidos da mestre, do PDF e dos inventários; ficam literais 420 m², 3.000/m², 750.000, 160/158,18.
- Variação dos lotes por script: +60 % (Casa da Praça, 250.000 → 400.000) a +100 %; o revisor tinha escrito 64 %.
- Saídas: pacote `a0d8504a1e`; script `fcbbb635bc`; perguntas `81a41b0fe8`; parte6 `0207e2fb9e`.
- Sem nova rodada de revisor: os números não mudaram (E1 +16.393.176,80; E3 +1.400.000; E4 360.000; soma 75.870.000).

## 2026-10-28 — Pacote 2 aprovado e aplicado; busca de destinos; Pacote 3 gerado
- Luís: "Aprovo E1–E6". `decisoes.json`:
  - 2025-01: Adoto Rafael, 66.000.000.
  - 2025-20: Adoto Rafael, 2.400.000.
  - `excecoes_sem_item.E4`: 360.000 não adotados.
  - Regras E2 e E5, com filtro por descrição.
- `registro-data-base.md`: registros 1 e 2 "mantido sem ajuste (E6)". `log-decisoes.md` atualizado.
- Mestre: soma dos valores adotados 75.870.000,00 (22 itens).
- `build_mestre_06.py`:
  - Regras com filtro sobrepõem a R3.
  - A R3 agora vale só para as 9 descrições aprovadas (`regras.R3.descricoes`, extraídas do Pacote 1). Assim a linha nova dos dólares ficou "Pendente", e não "sem valor" por arrasto.
- Busca `_scripts/busca_destinos_06.py` → `_extracoes/busca-destinos-06.xlsx`. Cobertura: 301 arquivos, 191 distintos por md5, no Drive Cisao_Siviero e em 01_fontes, mais o OCR. 1 erro: doc 22 de 2018 não baixado. Os .doc de 2003 (compromisso mútuo) não foram varridos.
- Doc 22 lido pelo conector do Drive (fileId 1HgldDKaDDFwpNAISVxnAsHpVHIpbliJK) → `_extracoes/doc22-2018-trechos.txt` (trechos literais). Hotel: "pode ser agregado a casa, EXCETO os lotes adquiridos adjacentes". Anexo "resumo dos dólares": U$ 37.698,00 no cofre em 25/05/2017; U$ 9.000,00 "Pago Coletto–Gramado – 02 lotes" em 15.05.2009. Soma conferida: 110 + 175 + 1.800 + 10.000 + 25.613 = 37.698.
- Achado: a versão do inventário do Hotel na pasta "Material antigo/2018_desatualizado" (arquivo "Dezembro de 2015", md5 `d9be609088`) lista como itens comuns do Hotel, sem nome e sem valor, as inscrições 360.0019.004 (20), 360.0038.015 (21), 360.0057.001 (22) e 360.0057.002 (23). O inventário da pasta "inventário-2015" (`2b482542d6`) tem 19 itens = 2018 01–19 (item 02 pela matrícula antiga 37.549).
- Doc 25 de 2018, seção Élcio, tabela "DESCRIÇÃO DO CAPITAL PAGO PELAS EMPRESAS": a troca Busatto aparece como capital pago a Élcio. O mesmo texto já constava da planilha de divisão de 2015 (sem a nota "trocado com o lote de Capão").
- Mestre: linha nova "Dólares do Hotel no cofre" (Não abordado, Pendente); nota do rascunho de 2018 nas linhas dos lotes de sócio e no item 17.
- `pacote_03.py` → `pacotes/pacote-03-destinos.md`: D1–D5 pendentes (D3 com destino parcial), D6 área construída, D7 dólares, D8 casa de 2019.
- Revisor do Pacote 3 disparado → `_diario/revisao-pacote-03.md`.

## 2026-10-28 — revisor do Pacote 3 e correções
- Revisor → `_diario/revisao-pacote-03.md`: 2 altos, 4 médios, 5 baixos. A lista dos 5 lotes está completa e as citações conferem; o doc 22 foi conferido por ele pelo conector.
- Conferido por mim:
  - Arquivo 18 (`8da060b1a3`), p. 3, lida como imagem: bloco 10.06.2019, linha "Lote Capão da Canoa | — | Lote na esquina próximo ao Hotel Araçá Ltda., com valor diferenciado aos demais lotes dados aos sócios." Planilha "de lotes vendidos e/ou outras formas de pagamento na aquisição de bens para acerto futuro" de Válter Materiais e Válter Siviero.
  - Convocação de 26/05/2017 no disco (`Material antigo/2018_desatualizado/G-SÓCIOS…`, `b159422801`): tem a mesma planilha de dólares (U$ 37.698,00) e a frase da casa.
- Correções:
  - A1: D1 cita o arquivo 18; a frase do resultado diz "possível entrega a sócio"; candidato 7 da Parte 6 (risco de cruzar com o item 02).
  - A2: D3 em três leituras, sem escolher.
  - M1: D2 com "CASA AO LADO: Raul". G7 junta os lotes e as casas; G8 fica só para a Toro.
  - M2: parágrafo de cobertura reescrito (o que foi lido e o que não foi).
  - M3: termos esquina, casa ao lado/da esquina, consignado, Coletto, dólar/cofre e lote/quadra acrescentados; busca refeita: 1.278 achados, nenhum destino novo; a única menção nova é histórica (U$ 50.000 para a entrada na compra do Hotel, doc 25 e divisões de 2015).
  - M4: dólares tratados como bem de 2018 fora do inventário (P4), fonte em disco, retiradas de Válter e Luciane, pista dos lotes do Lago Negro (doc 6), texto da G12 no pacote.
  - B1: D8 com data de 2017. B2: doc 5 é de José Valério. B3: frases. B4: as evidências vêm de todas as notas da mestre. B5: inscrições próprias de Élcio e Irineu explicadas.
- Mestre:
  - Linha dos dólares com "Existe em 2018? Sim, fora do inventário" e fonte em disco.
  - Notas do arquivo 18 (D1) e do doc 25 (casas, D2 e D3).
- Saídas: mestre `a19d843afb`; pacote 3 `md5 na tabela acima`.
- Sem nova rodada de revisor: só texto e evidências; nenhum valor envolvido.

## 2026-10-28 — Pacote 3 aprovado e aplicado
- Luís: "Aprovo o Pacote 3".
- `decisoes.json`:
  - `destinos_2018` com 2018-20, 22, 23, 24 e 25.
  - `pendencias_componentes.D6`.
  - `regras.D7`, com filtro "cofre em Veranópolis".
- `perguntas-ao-grupo.md`: G7 refeita, G8 só para a Toro, G12 nova.
- `parte6-candidatos.md`: candidatos 6 e 7.
- `log-decisoes.md` atualizado.
- Mestre regenerada: 32 linhas; 22 itens de 2025 com valor (75.870.000,00); 10 linhas sem valor nesta etapa (5 lotes de 2018, 2 só Rafael, 3 não abordados).
- Próximo: revisão independente de todas as decisões aprovadas (P11), depois a ficha de fechamento.

## 2026-10-28 — revisão final das decisões e ficha de fechamento
- Revisor final → `_diario/revisao-final-decisoes.md`: 0 altos, 5 médios, 7 baixos. As 22 decisões, a soma de 75.870.000,00 e o C0 conferem; não há decisão aplicada sem aprovação.
- Corrigido (só registro):
  - M1: Totais da mestre mostra o E4 e o D6, lidos de `decisoes.json`.
  - M2: E2 passa a citar a G7 (`correcoes` no json).
  - B1: o D8 entra no log.
- Levado à ficha como proposta, sem aplicar: G13 (item 02 × arquivo 18, M3), G14 e candidato 8 (regularização, M4), G1 e G10 ampliadas (B5, B6), redação do E5 (B2), motivo do R2 por item (B5).
- Também na ficha: M5 (material não lido), B3 (dólares tratados no bloco da Parte 4), B4 (casas = 2 bens em 3 linhas), B6 (sem área em dobro nos lotes 15–17), B7 (erro da planilha 2_0).
- Novo `_scripts/ficha_fechamento_06.py` → `ficha-fechamento-06.md`. Asserts: 22 decisões; soma adotada = soma dos itens de Rafael; E1 + E3 = diferença.
- A ficha segue a lista de conteúdo do revisor final; sem rodada nova de revisor sobre ela, porque só reúne números já revisados.
- Pendente: aprovação da ficha pelo Luís e das 6 propostas; depois, a união com os arquivos da SEI (P13), que cabe à sessão SEI.

## 2026-10-28 — ficha aprovada; empresa 06 fechada
- Luís: "Aprovo tudo" (ficha de fechamento e as 6 propostas).
- `decisoes.json` (`5560115d85`):
  - E5 com nova redação.
  - Motivos de 2025-17, 21 e 22 por item.
  - Entradas em `correcoes` e `aprovacoes`.
- `perguntas-ao-grupo.md` (`b5dbe781f8`): G1 e G10 ampliadas; G13 e G14 novas; 14 perguntas, nenhuma enviada.
- `parte6-candidatos.md` (`b5584fed47`): candidato 8 (regularização).
- `ficha_fechamento_06.py` marca a ficha como aprovada quando `aprovacoes` tem a etapa "Ficha de fechamento".
- Mestre `932fd241d3`; ficha `2d86bab84a`; log `e23e128e47`. Valores sem mudança: 75.870.000,00 (1/6 12.645.000,00).
- Fim do brief S2-06. A união dos registros desta pasta com os da raiz de `02_trabalho` cabe à sessão SEI (P13). Nota de passagem atualizada na pasta de notas das sessões.
