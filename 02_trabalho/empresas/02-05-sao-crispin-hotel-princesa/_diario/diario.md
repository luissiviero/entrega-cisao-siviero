# Diário de trabalho — empresas 02 (Agropecuária Fazenda São Crispin) e 05 (Hotel Princesa dos Vales)

Somente acréscimos. Uso interno de Claude e do revisor (P12). Para retomar: ler o brief `02_trabalho/empresas/_orquestracao/prompts/S1-02-05-sao-crispin-hotel-princesa.md`, `00_PLANO.md`, o fim deste diário e `log-decisoes.md` desta pasta; números sempre re-derivados por `_scripts/fase0_extracao.py`, nunca da memória.

## 2026-10-28 — Fase 0: fontes, extração e conferências

### Leitura prévia
- Brief S1 lido inteiro; `00_PLANO.md` (P1–P13, seção 7); `02_trabalho/log-decisoes.md` e `_diario/diario.md` da SEI; `analise-coordenacao.md` (seções 1–5); memória do projeto.
- `02_trabalho/analises/acertos-e-areas-vendidas/achados-para-empresas.md` (md5 `3fca77e8c9`): situação "aguardando aprovação do Luís". **Não usado.** Observação: o achado E2 de lá usa 85,32 ha para a Pellegrini; as parcelas do doc 17 de 2018 somam 85,22 ha (ver conferências).

### Entradas (md5, 10 dígitos)
- `01_fontes/inventario-2025/02. …SÃO CRISPIN….docx` `63f1037122`
- `01_fontes/inventario-2025/05. …HOTEL PRINCESA….docx` `39bb7a345a`
- `01_fontes/inventario-2025/HOTEL PRINCESA … ÁREA DE TERRAS VENDIDAS….docx` `6cdf8807c9`
- `01_fontes/inventario-2025/00. PLANILHA DO VALOR DO CAPITAL….docx` `a1670b10fe`
- `01_fontes/inventario-2025-antigo/ANTIGO 05. …(2).docx` `d50daf8944`
- `01_fontes/inventario-2025-antigo/ANTIGO 00. ….docx` `8ed9d82a32`
- Drive `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\13. 1. HOTEL PRINCESA DOS VALES LTDA. = I N V E N T Á R I O = … Dezembro de 2018.docx` `ca1a85e9f7`
- Drive `…\inventário-2018\17. 1. AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA. = I N V E N T Á R I O = … Dezembro de 2018.docx` `dc0476f1f7`
- `01_fontes/inventario-2018/27. …RESUMO….docx` `af5ec6c22e`
- `01_fontes/inventario-2018/18. …SIVIERO EMPREENDIMENTOS….xlsx` `8c63e21506`
- `01_fontes/estudo-rafael/HOTEL PRINCESA DOS VALES LTDA- BELISE LTDA.pdf` `067bfc3721` (24 p., com texto)
- `01_fontes/estudo-rafael/AGROPECUARIA E FAZENDA SÃO CRISPIN LTDA - BELISE LTDA.pdf` `b3a7709f8d` (4 p., com texto)
- `01_fontes/estudo-rafael/DIRETRIZES ….pdf` `aff2b08f2a`
- `01_fontes/outros/ativos-avaliados-valter.xlsx` `935282c1aa`
- `01_fontes/planilhas-google/2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx` `11a81ef0aa`
- `02_trabalho/observacoes-planilhas-google.xlsx` `18bf9de495`
- Lidas depois (temas e pacotes), hash agora: respostas ao Dr. Pablo arq. 1 `03968eb843` (32 p., com texto); arq. 7 `1af6d0b269`, arq. 8 `99104ed82f`, arq. 9 `7b519bed33` (arrendamentos, escaneados, 6+5+6 p. sem texto); arq. 11 `d27c0a1be3` (matrícula 1.412, 3 p. sem texto); `conferencia-discrepancias.md` `bdd7c87e70`; `inventario-2025-avaliacoes-corrigido.xlsx` `a565450170`; `diff-inventario-2025-antigo-x-atual.md` `7afde86365`; `pendencias-conferencia-anterior.md` `1e9eca2c47`; `parte6-candidatos-duplicidade.md` (SEI) `32ee3e62a4`.

### Saídas
- `_scripts/dump_docx.py` md5 `5ee2f7614f` (texto dos docx, para leitura; saídas no scratchpad da sessão).
- `_scripts/fase0_extracao.py` md5 `f5f8dc961e`.
- `fase0-extracoes-02-05.xlsx` md5 `d795a680f9` — abas: Conferências, Fontes, Itens 2025, 2025 antigo x atual, Itens 2018, Totais 2018, Vendas 2018, Áreas da fazenda, Rafael 05 valores, Rafael 05 faixas, Rafael notas, Grupo x Rafael 05, Lista de Rafael, Planilhas Google, SEI 2018 outras empresas, Resumos de capital.

### Contagens
- 2025 atual: 05 = 8 itens; 02 = 1 linha "-.-" (R$ 0,00). 2025 antigo 05 = 8 itens.
- 2018: Hotel (doc 13) = 4 itens; Agropecuária (doc 17) = 17 linhas (8 da fazenda, numeradas "01" ou sem número; itens 02 a 10); 3 vendas.
- Rafael Hotel: 36 valores lidos em 5 seções (grupo p. 1–4; avaliação executada p. 7–15; balanço executado p. 16–19; quadro-resumo p. 19–20; escolhidos p. 22–23); 8 frases de faixa de mercado (item 01 p. 8–9; item 08 p. 15). Rafael Agropecuária: nenhum valor de item ou total.
- Lista de Rafael: 4 itens da empresa 05 (04, 05, 06, 08); nenhum da 02. Diretrizes p. 4–6.
- Planilhas Google: 579 células lidas (abas 2.São Crispin, 5.Hotel Princesa, Extras; GERAL 11–15, 28–69, 941–950, 986–987); 25 células de observação (colunas OBSERVAÇÕES / F e H da GERAL).

### Conferências (59; 57 OK, 2 divergem)
- OK: totais 2025 atual (75.410.201,00) e antigo (83.640.801,00) do Hotel; doc 00 atual e antigo; Agropecuária 2025 = 0; diferença antigo → atual = −8.230.600,00 (itens 05 e 06).
- OK: 2018 Hotel 34.404.877,83 (= doc 27); 2018 Agropecuária: subtotal fazenda 22.775.090,00; bruto 24.112.390,00; recebido 12.058.000,00 (= 3 vendas); líquido 12.054.390,00 (= doc 27).
- OK: matrículas do item 05 somam 491,0 ha em 10 linhas, 9 números (00.760 duas vezes); 5 sub-áreas × R$ 110.000 = 44.828.300,00 em 407,53 ha. 2018: os 491 ha avaliados inteiros (17.261.200,00) apesar do "vendida somente 83,47 ha". Parcelas da fazenda 2018 somam 686,75 ha; Rigo = 83,47 + 106,53 (mat. 11.090) = 190,00.
- **DIVERGE 1:** parcelas vendidas à Pellegrini no doc 17 somam 85,22 ha (4,9 + 10,0 + 26,16 + 44,16), o quadro diz 85,32 ha. Com 85,22, o remanescente 686,75 − 190 − 85,22 = 411,53 ha = 407,53 (item 05) + 4,0 (item 06) de 2025. Achado, não erro de leitura.
- **DIVERGE 2:** as 25 observações destas abas não estão em `observacoes-planilhas-google.xlsx` (sessão SEI). Extraídas aqui, aba "Planilhas Google". Avisar a sessão SEI só pelo relatório (não escrevo lá).
- OK: Rafael reproduz o balanço do grupo (75.410.201,00); balanço executado soma 79.491.957,50; diferença +4.081.756,50 = item 01 (+4.050.756,50: 17.555.000,00 × 13.504.243,50) + item 08 (+31.000,00: 85.000 × 54.000). Na seção "avaliação executada" (p. 7–15) Rafael repete 13.504.243,50 no item 01 e 54.000,00 no item 08; os valores diferentes só aparecem no balanço executado (p. 16 e 19). O item 08 aparece duas vezes no balanço executado (p. 19), contado uma vez no total.
- OK: quadro-resumo: 1/6 = 13.248.659,583; escolhidos 50.914.157,50; diferença a complementar 37.665.498,00 (cálculo dá 37.665.497,92).
- OK: nenhum texto sobre Rovílio ou "outubro/2025" nos docs 02, 05, antigo 05 e docx de áreas.

### Notas para as próximas etapas (fatos lidos, sem decisão)
- Rafael item 01: 17.555.000,00 sem memória de cálculo; faixas citadas: terreno 1.800 m² R$ 2,7–4,5 mi (R$ 1.500–2.500/m²), prédio 2.040 m² R$ 9,2–13,3 mi. O grupo avalia o terreno a R$ 6.000/m² (720 m², item 01) e R$ 7.500/m² (1.080 m², item 02).
- Item 08: faixa de Rafael R$ 85.000–130.000 (FIPE ~R$ 54.745); grupo cita FIPE out/2025 R$ 56.854 (arq. 1 p. 9) — referência após a data-base.
- Arq. 1 p. 6: Agropecuária criada "para fins operacionais fiscais-tributários", arrendamento Hotel × Agropecuária "pro forma"; receitas do arrendamento Lorenzet "saldadas conforme fluxo financeiro necessário" para a SEI. P. 9: 00.760 = duas compras (48 e 100 ha), restam 64,53 ha; avaliações "na data da elaboração do inventário" (data-base). P. 1: "termo final … outubro/2025".
- Inventário 2018 da SEI, aba "Bens em outras empresas": Agropecuária itens 02 a 06 (Angeli 10,2 ha; lote 12; lote 89; lote 33; lote 52) e Hotel item 04 (Sede Campestre) listados como bens da SEI.
- Sede Campestre 2025: texto diz "25,00% (Vinte por cento)"; cálculo usa 25 %.
- Planilha 2_0 GERAL F943: a planilha soma R$ 13.932.840,00 de vendas (inclui as 4 parcelas Pellegrini além do preço de R$ 2.100.000) contra R$ 12.058.000,00 do documento.

### Pendências
- Revisor independente da Fase 0 (P11): a disparar agora; saída `_diario/revisao-fase0.md`.
- Ler como imagem: arq. 7, 8, 9 (arrendamentos) e arq. 11 (matrícula 1.412) — na etapa dos temas.
- Achados de ACERTOS: aguardando aprovação do Luís.

## 2026-10-28 — achados de ACERTOS aprovados (E1–E3); OCR disponível
- Mensagem da sessão de ACERTOS: pacote 1 aprovado. Conferido no log dela (`analises/acertos-e-areas-vendidas/log-decisoes.md` md5 `997d118fa0`): "aprovado pelo Luís ('Aprovo o pacote 1'), sem ajustes", inclui E1, E2, E3. `achados-para-empresas.md` agora md5 `ca3c0b913b`: E1–E3 aprovados; E4–E7 no pacote 2, não aprovados (não usar).
- Ressalva ao E2 (aprovado): o texto soma 106,53 + 85,32 = 191,85 ha e compara com 686,75 − 491 − 4 = 191,75 ha. As parcelas Pellegrini do doc 17 somam 85,22 ha (conferência DIVERGE 1); com 85,22, 106,53 + 85,22 = 191,75 e tudo fecha sem resto. A conclusão do E2 (áreas fecham) se mantém; os 0,10 ha estão no total declarado da venda Pellegrini (85,32), não nas áreas de 2025. Levar ao Luís junto com a Fase 0.
- Aviso da coordenação: Tesseract 5.4 em `C:\Program Files\Tesseract-OCR\tesseract.exe`, só `eng`/`osd`. Usar para localizar trechos nos arq. 7, 8, 9, 11; conferir números lendo a página como imagem.

## 2026-10-28 — revisor 1 da Fase 0 e correções (script v2)
- Relatório: `_diario/revisao-fase0.md` — 2 altas, 7 médias, 6 baixas. Todos os itens dos docs, totais e valores de Rafael confirmados pelo revisor.
- Correções aplicadas em `_scripts/fase0_extracao.py` (v2, md5 `83dcfd574f`); saída `fase0-extracoes-02-05.xlsx` md5 `41c878b74d`:
  - A1: nova aba "Créditos de soja" (parágrafos OBSERVAÇÕES do docx de áreas): Rigo 40.300 − 2.606 = 37.694; Lorenzet 2.606 + 3.206 + 3.591 + 3.682 = 13.085, sendo 3.682 (2024) anotadas "AGROPECUÁRIA". Conferências: R$ 3.224.000 (2 × 20.150 × R$ 80) marcados "FALTA RECEBER" dentro do "TOTAL RECEBIDO" de R$ 12.058.000; 3.682 sacas da 02 × inventário 02 = R$ 0. Aviso: 2.606 repetido (Rigo vendida 2024 e Lorenzet 2021).
  - A2: situação "VENDIDA PARCIAL (83,47 ha)"; colunas "Área avaliada", "Comprador" e "Texto da célula de valor" em Itens 2025/2018.
  - M1: constante 46.459.267,83 removida. Aba "Base 2018": base A (líquido doc 17 + doc 13) 46.459.267,83; base B (planilha GERAL C28) 48.730.897,83, recomposta pelas fontes; B − A = 2.271.630 = ganho na venda 2.939.430 (preço 12.058.000 − inventário 9.118.570) − trator 116.600 − lotes 03/04/05 551.200. Conferências de identidade, sem número digitado. Escolha da base = Luís.
  - M2: áreas declaradas lidas dos parágrafos (doc 17 e docx: 686,75 / 190 / 85,32 / 411,43; doc 05: 411,4 e a)–d)). Novas DIVERGE: 411,53 × 411,43 (vem dos 85,32) e 411,4 (texto 2025) × 411,53 (itens 05 + 06).
  - M3: padrão "R$ x a R$ y" (pega "gira em torno de R$ 5.600 a R$ 6.200", p. 9); duplicatas e "R$ " solto removidos.
  - M4: planilha Google ampliada: GERAL 227–241 e 882; Resumo 73 e 78–94; 11.Empreendimentos 90 e 731; Acertos e Acordos 36–43; comentários em fio lidos. 777 células, 67 observações; só 5 delas estão no arquivo da sessão SEI. Aviso: Resumo B73 "Hotel Princesa dos Vales" aponta para o caminhão-tanque da Abastecedora (GERAL linha 8).
  - M5: aba "Vendas docx x doc 17" (5 diferenças: Rigo "= 118.100 SC SOJA", "37.500 sacas", 2 × "FALTA RECEBER"; Angeli R$ 40.000/ha em 2018 × R$ 50.000/ha no docx).
  - M6: "Rafael notas" por item (concordância expressa: 02 p. 10, 03 p. 11, 04 p. 12, 06 p. 14, 07 p. 15; nenhuma para 01 e 08; 05 citado na concordância da p. 14 mas com recomendação de avaliação independente p. 13–14). Conferência: os componentes impressos no item 01 do balanço executado são os do grupo; 17.555.000 sem memória de cálculo.
  - M7: coluna "Também listado pela SEI em 2018" (6 linhas: Agro 02–06 e Hotel 04) e aba "2018 → 2025 por matrícula" (busca nos docs 2025 01, 03, 04, 06–11 e 05). Sem ocorrência em 2025: parcelas vendidas (982, 11.661, Ademar, Garbin), Angeli 2.424/2.425, benfeitorias R$ 300.000 e trator Valtra. Lotes 12, 89, 33 → doc 11 (SEI); 459/460/461/558 e 15.922 → doc 08 (Transportadora); 7.391 → Hotel item 07; mat. 11.090 → Hotel item 06.
  - B1–B6: Totais 2018 só com totais do inventário (sem capital social, sem duplicatas); "Resumos de capital" separa total × cota por sócio (6 × 12.568.366,83 = 75.410.200,98); aba "Menções" (doc 00: venda das quotas de Rovílio); conferências internas dos itens 01, 02, 03, 04, 06, 07 (2025) e 01, 04 (2018); despesas de regularização 356.840 × "TOTAL R$ 400.000" (arredondado) e redução de R$ 400.000 fora do total; pareamento antigo × atual por número de item; subtotal da fazenda por posição do subtotal; Diretrizes com limite de dígitos; avisos: Garbin 44,16 × 44,1 ha; mat. 11.090 em duas parcelas; "25,00% (Vinte por cento)".
- Fontes: 37 (inclui os docs 2025 01, 03, 04, 06–11 para a busca por matrícula e o log da análise de ACERTOS). `achados-para-empresas.md` agora md5 `0a0377229a` (a sessão de ACERTOS anotou no E2 a soma de 85,22 ha).
- Resultado: 116 conferências, 10 DIVERGE, todas achados das fontes (ver aba Conferências, coluna Nota).
- Próximo: re-revisão curta das correções (P11) antes de apresentar ao Luís; saída `_diario/revisao-fase0-r2.md`.

## 2026-10-28 — revisor 2 da Fase 0 e correções (script v3)
- Relatório: `_diario/revisao-fase0-r2.md`. Rodada 1: 10 corrigidos, 5 em parte; nenhuma regressão (totais, itens, Rafael por seção, 37 MD5). Novos: 4 médios (N1–N4), 9 baixos (N5–N13).
- Correções (script md5 `1c0eaf5bb6`; planilha md5 `91b19d1691`; 46 fontes, 114 conferências, 10 DIVERGE, 10 avisos):
  - N1: aba "2018 → 2025 por matrícula" com Área 2025, Valor 2025 e Nota; aviso: lote 12 (mat. 3.062) 20,5 ha em 2018 × 13,0 ha em 2025; os 7,5 ha batem em área e vendedor (Sangali) com o doc 11 T8R4 (lote 14). Possível dupla contagem em 2018 → candidato à Parte 6 (cruza com a SEI).
  - N2: nota na linha da parcela vendida de 106,53 ha (mat. 11.090 compartilhada; a parcela vendida não está em 2025).
  - N3: parágrafo "…2.606 SACAS PARA" unido ao seguinte; aviso: soja → "escavadeira hidráulica"; candidata doc 11 T5R2 (escavadeira 320, 2024, R$ 887.000,00). Hipótese de fluxo Hotel → SEI (P4/Parte 6).
  - N4: nota do "FALTA RECEBER" reescrita como leitura em aberto (tabela "FALTA RECEBER" × observações "não vendidos" com venda parcial em 2024; R$ 80 = cotação de out/2018).
  - N5: conferência trivial da base 2018 do item 04 removida. N6: redução de R$ 400.000 conferida contra o TOTAL das despesas (OK); ausência de abatimento virou aviso. N7: conciliação da Base 2018 rotulada "decomposição que reproduz B − A", premissa R$ 44.000/ha escrita; conferência-identidade removida. N8: diferença de redação 02 × 05 mostrada ("para a" × "em nome da"); âncoras falham alto. N9: nota da sobreposição SEI 2018 por base (A: R$ 5.064.791,80 listados; B: lote 52 + Sede = R$ 4.105.591,80). N10: marcação de observação restrita (36 observações; 2 no arquivo da SEI) e contagem movida para Avisos. N11: tipo "total do grupo". N12: números de rótulos/notas lidos de variáveis, formato brasileiro, colunas da conferência 411,4 corrigidas. N13: busca também nos inventários 2025 antigos (lote 52 no ANTIGO 11 T9R3 a R$ 226.500).
  - M6: coluna Notas do item 05 cita "citado na concordância de outros itens (p. 14, p. 15)".
  - Não tratados (baixos, sem efeito nos números): marcador de item de Rafael só 1–8 (o PDF só tem 8 itens); valores de Rafael só em linha isolada "R$ x" (conferido pelos dois revisores); compra do item 08 de 2018 por R$ 50.000 fica só no texto; vendedor vazio na parcela Garbin; "100.874,027,60" do doc 00 (não afeta 02/05).
- As correções v3 não passaram por uma 3ª revisão isolada; o revisor do pacote 1 (planilha-mestre) relerá esta extração.
- Próximo: apresentar a Fase 0 ao Luís e aguardar OK para a planilha-mestre.

## 2026-10-28 — Fase 0 aprovada; planilha-mestre gerada
- Luís: "Seguir com a planilha" (OK da Fase 0 e da etapa planilha-mestre). Base 2018 sem escolha: a planilha mostra A e B. Registrado em `log-decisoes.md`.
- Novo script `_scripts/build_mestre_02_05.py` (md5 `f1c32e710a`): roda `fase0_extracao.main()` e lê `fase0-extracoes-02-05.xlsx` (md5 `bd1878b09c`); decisões de `decisoes.json` (vazio). Saída `planilha-mestre-02-05.xlsx` (md5 `b89ee5a3cb`): abas Mestre, Resumo, Passivos e deduções, Legenda.
- 24 linhas: 8 bens 2025 do Hotel (6 Concorda, 2 Discorda: 01 e 08); 13 itens 2018 sem par (6 vendidos em 2018; 3 na SEI 2025 — lotes 12, 89, 33; 2 na Transportadora — Cotiporã; 2 sem destino — benfeitorias R$ 300.000 e trator Valtra R$ 116.600); 3 não abordados (soja Rigo 37.694; Lorenzet 2021–2023 9.403; Lorenzet 2024 3.682).
- Pares 2018 × 2025 (matrícula + área/vendedor/placa): 01↔Hotel 01 (mat. 1.412 + 1.800 m²); 02↔Hotel 02 (1.412 + 1.080 m²); 03↔Hotel 03; 04↔Hotel 04; 05↔Agro seq. 1 (491 ha); 06↔Agro seq. 7 (11.090 + 4,0 ha); 07↔Agro seq. 13 (7.391); 08↔Agro seq. 17 (placa JTB 8585).
- Conferências do script: total grupo 75.410.201,00 = doc 05; total Rafael 79.491.957,50 = balanço executado; 21 linhas de 2018 todas contadas (8 pares + 13 sem par); lista de Rafael = 04, 05, 06, 08. Soma pelo C0 (referência) R$ 79.460.957,50.
- Observações da planilha 2_0 mapeadas célula a célula (OBS_MAP; duplicatas das abas das empresas e cópias do texto do inventário excluídas explicitamente; o script para se uma célula não estiver mapeada).
- Criados: `registro-data-base.md` (RDB-1 a 4), `parte6-candidatos.md` (P6-01 a 07), `perguntas-ao-grupo.md` (Q1–Q10, rascunho não aprovado), `log-decisoes.md`, `decisoes.json`.
- Próximo: revisor independente da planilha-mestre (P11), saída `_diario/revisao-mestre.md`. Depois, apresentar a planilha ao Luís; o pacote 1 (regras em bloco) só com novo OK.

## 2026-10-28 — OCR revisado dos arquivos 2 a 16 disponível (aviso da coordenação)
- Pasta da sessão SEI (só leitura): `02_trabalho/ocr-respostas-docs/` — um .md por arquivo, `ocr.jsonl` (100 p.) md5 `9b5b52b7c1`, `LEIA-ME.md` md5 `3cb20e4a81`. Revisão: `02_trabalho/_diario/revisao-ocr-respostas-docs.md` md5 `ae2b42eb74` (22 páginas, ~450 números).
- Uso nesta sessão: localizar trechos nos arq. 7, 8 e 9 (arrendamentos Lorenzet) e menções cruzadas. Números do OCR têm 3–4 % de erro: conferir cada número na imagem da página e citar arquivo e página. Arq. 11 (matrícula 1.412): ler a imagem, não o OCR.
- Sem mudança na etapa atual (revisão da planilha-mestre em andamento).

## 2026-10-28 — revisor da planilha-mestre e correções (mestre v2)
- Relatório: `_diario/revisao-mestre.md` — 2 altas, 7 médias, 10 baixas. Confirmados: 21 bens 2018 e 8 de 2025 uma vez cada, pares certos, totais, valores de Rafael por seção, lista, C0, localizações nos docs 08 e 11, nenhuma decisão apresentada como aprovada.
- A1 (lição 1, 3ª ocorrência do tipo no projeto): Rafael avalia 5 bens de 2018 nos documentos das empresas onde estão em 2025. Agora lidos por matrícula: lotes 12, 89, 33 no xlsx "ÁREAS RURAIS…" (Balanco_Grupo + Quadro_Resumo: R$ 650.000, 1.500.000, 250.000; grupo 520.000, 250.200, 40.000); Cotiporã 69,8 ha e 4,5 ha no PDF da Transportadora (executado p. 14–15: R$ 3.500.000 e 250.000 = grupo; concordância p. 7). Decisão desses valores é das sessões SEI e S3. Para ler a Transportadora, o marcador de item de `rafael_sections` passou a aceitar 1–19 (Hotel sem mudança: 115 conferências, 10 DIVERGE).
- A2: item 05 agora "Concorda com ressalva" (valor igual, sem concordância expressa, Rafael chama de preliminar e pede avaliação independente, p. 13–14). Fica fora da regra em bloco do pacote 1.
- M1: P6-08 (R$ 250.000 de regularização de áreas rurais "por contrato" no doc 11 × R$ 400.000 da fazenda nos docs 02/05) + linha de referência na aba "Passivos e deduções" (texto lido do doc 11) + Q11.
- M2: aba Resumo da 2_0: células de nota listadas uma a uma (RESUMO_OBS: B20, B80, E80, B87, B88, E88, F88, F91, B92, E92, C94, B96, E96, B100, B191); B20 vai para "observações da empresa" no Resumo da mestre. 45 observações; 2 no arquivo da SEI.
- M3: Base B = GERAL C28 (Hotel) + C11 (Agropecuária, trator) = R$ 48.847.497,83; B − A = R$ 2.388.230,00 = ganho na venda 2.939.430 − lotes 03–05 551.200. Conferências: kept = C28; C11 = trator.
- M4: item 01 com área "1.800 m² na matrícula; 720 m² avaliados neste item" e nota sobre a faixa de Rafael (p. 8) cobrir a matrícula inteira; item 02 com nota da mesma matrícula.
- M5: soja Rigo "Existe em 2018: Sim (doc 17 T6R1)", empresa "05 (docx 2025) / venda pela 02 em 2018", ligada a P6-06. M6: coluna "Existe em 2025 (grupo)?". M7: F943 → parcelas Pellegrini s03–s06 e item 05.
- Baixas: RDB-1 em todos os bens do Hotel; colunas "Valor 2025 versão antiga" e "Valor na 'avaliação executada'"; "Fonte Rafael" com todas as seções e o Diretrizes; nota do item 08 impresso duas vezes; Q2, Q4, Q5, Q7 ajustadas, Q12 nova; P6-03 com evidência contrária (mat. 4.194, lote 14 × 12); P6-05 com R$ 3.750.000 e Rafael; P6-07/Q7 marcados como coincidentes com E4 não aprovado; cabeçalho "Bloco"; benfeitorias em bloco próprio; vendas 2018 como "Saída em 2018 por venda … destino (P4) a confirmar"; H14 também no item 05; nota dos 83,47 ha na linha 05.
- Saídas (md5): build_mestre_02_05.py `67aa14048a`; fase0_extracao.py `fd6456d95d`; planilha-mestre-02-05.xlsx `073a74e542`; fase0-extracoes-02-05.xlsx `b61ceb53c5` (o md5 dos xlsx muda a cada geração por causa da data interna); parte6-candidatos.md `eba6071d83`; perguntas-ao-grupo.md `8abb565a9c`.
- Próximo: re-revisão curta das correções (P11), saída `_diario/revisao-mestre-r2.md`.

## 2026-10-28 — aviso da sessão S3 sobre as terras de São Brás (Cotiporã)
- S3 (`01-08-abastecedora-transportadora/rovilio-quotas.md`, md5 `c6ace6758d`, seções 4 e 8.4; segundo a S3, aprovado pelo Luís): o contrato de venda das quotas do Rovílio (arq. 10, cl. 3ª, p. 4) inclui no preço "créditos hereditários e/ou creditórios" sobre a "área de terras rurais localizada na Comunidade de São Brás, no Município de Cotiporã", "em nome das COMPRADORAS". A S3 liga isso ao item 07 da Transportadora 2025 (69,8 ha, R$ 3.500.000), que em 2018 era o item 07 da Agropecuária ("NÃO DIVISÍVEL"); o item 08 (4,5 ha) foi comprado de Tercílio Dal Más em 2008. Nenhuma fonte traz a transferência Agropecuária → Transportadora.
- Conferido no doc 17 2018: itens 07 (mat. 459, 460, 461, 558; 69,8 ha) e 08 (registro 15.922; 4,5 ha; comprado de Tercílio Dal Más em 25/09/2008 por R$ 50.000) com "R$ NÃO DIVISÍVEL". Não li o arq. 10 (fato posterior à data-base, analisado pela S3).
- A incorporar em P6-05 e nas linhas 02-2018-s14/s15 depois da re-revisão da mestre em andamento (para não alterar arquivos durante a revisão): referência cruzada à S3, falta de documento de transferência (pergunta ao grupo).

## 2026-10-28 — re-revisão da planilha-mestre (r2) e ajustes finais
- Relatório: `_diario/revisao-mestre-r2.md`: 18 de 19 achados corrigidos, 1 em parte (M1); nenhuma regressão (totais, pares, C0 R$ 79.460.957,50, lista 04/05/06/08); valores de Rafael dos 5 bens em outras empresas conferidos nos PDFs (item por matrícula, seção e página certas). Novos: N1 (média) e N2–N8 (baixas).
- Conferido por mim nas fontes: doc 11 T9R4–T9R5 (lote 47 Piccolo e Scalco: "por contrato", "área ainda não registrada em nome da empresa", "Parte da área foi destinada para compensação ambiental do Loteamento Colina do Retiro II"); PDF "ÁREAS RURAIS…" p. 17 "Matrícula 8.846" no texto do lote 89; PDF do Hotel p. 13 "tendo como resultado valores inferiores àqueles apresentados ou atribuídos ao imóvel em outros levantamentos".
- Correções: N1 — P6-08 marcada "hipótese fraca" com a evidência contrária; Q11 reescrita (lote 47?); aba Passivos com E882 e T9R4–T9R5. N2 — citação completa ("em nome da empresa HOTEL PRINCESA DOS VALES LTDA. e de TERCEIROS, POR CONTRATO"). N3 — s10 e P6-03 com o item 03 de Rafael para os 7,5 ha (R$ 220.000 → 375.000, p. 1 e 22). N4 — documento de Rafael = PDF (xlsx como apoio), páginas da avaliação, aviso automático "p. 17 grafa a matrícula como 8.846". N5 — linha no log-decisoes sobre a base B (R$ 48.730.897,83 → R$ 48.847.497,83), sem reescrever o registro. N6 — frase da p. 13 lida por script na nota do item 05. N7 — legenda de "Concorda com ressalva". N8 — texto do doc 11 sem corte, "Diferença %" nas linhas s10–s15, ordem Q10–Q13, nota de área sem corte no meio da palavra.
- Aviso da S3 incorporado: P6-05 com a referência à venda das quotas do Rovílio (arq. 10, cl. 3ª) e Q13 (documento da transferência Agropecuária → Transportadora).
- Hashes finais: ver linha seguinte. As correções desta rodada não passaram por uma 3ª revisão isolada; todas as decisões serão revistas antes da ficha de fechamento (P11).
- `_scripts/build_mestre_02_05.py` md5 `5b5e5623cb`
- `_scripts/dump_docx.py` md5 `5ee2f7614f`
- `_scripts/fase0_extracao.py` md5 `2d5865c52f`
- `planilha-mestre-02-05.xlsx` md5 `181d62d59a`
- `parte6-candidatos.md` md5 `f6f4b5678a`
- `perguntas-ao-grupo.md` md5 `a5cef30745`
- `log-decisoes.md` md5 `ac13dfcb54`

## 2026-10-28 — planilha-mestre aprovada; pacote 1 (rascunho)
- Luís: "Montar o pacote 1" (OK da planilha-mestre e da etapa 3).
- Novo script `_scripts/pacote_01.py`: refaz a mestre e escreve `pacotes/pacote-01.md`. Mestre ganhou notas automáticas dos achados da Fase 0 por item (DIVERGE e avisos que citam "item NN") e o risco C1 do item 02 (faixa de Rafael da p. 8 × R$ 7.500/m²).
- Conteúdo: R1 = 5 bens do Hotel (02, 03, 04, 06, 07) com valor igual e concordância expressa, R$ 17.023.657,50 → adoto grupo; R2 = 5 bens de 2018 em outra empresa (lotes 12, 89, 33 → SEI; São Brás 69,8 + 4,5 ha → S3) fora do total 02/05, decisão das sessões donas; R3 = 3 créditos de soja → Parte 5. Fora: pacote 2 (itens 05, 01, 08) e etapa 6 (8 destinos de 2018).
- Próximo: revisor independente do pacote 1 → `_diario/revisao-pacote-01.md`.

## 2026-10-28 — revisor do pacote 1 e correções (pacote v2)
- Relatório: `_diario/revisao-pacote-01.md` — 1 alta, 3 médias, 5 baixas. Confirmados nas fontes: totais, os 5 valores iguais em todas as seções de Rafael, páginas de concordância, lista, C0 neutro, R2 e R3, frase sobre os pacotes 3 e 4 da SEI.
- A1 (alta): item 06 (4 ha da fazenda) saiu da Regra 1. Na mestre, item da "FAZENDA SÃO CRISPIN" com a ressalva da fazenda vira "Concorda com ressalva", com nota (avaliação independente p. 13–14; redução de R$ 400.000; concordância da p. 14 remete ao 05). Regra 1 = 4 bens (02, 03, 04, 07), R$ 16.583.657,50. Pacote 2 = 05, 01, 06, 08 + redução de R$ 400.000.
- A2: item 04 com benfeitorias não averbadas (arq. 1 p. 9, citado por script) e casa sem área; Q14 nova.
- A3: coluna "Par em 2018" e conferência dos 21 bens de 2018 (8 pares + 5 Regra 2 + 8 etapa 6); Agropecuária 2025 R$ 0,00 citada; redução de R$ 400.000 listada com o pacote 2.
- A4: Regra 2 diz "localização", destino pendente até documento (Q9, Q13). A5: caso 7 do pacote 4 da SEI (lote 12) citado, fica com a SEI. A6: faixas de Rafael p. 8 (item 02) e p. 9 (item 03) citadas com escopo. A7: concordância condicionada (p. 11 e 21, citadas por script) e RDB-1; "Argumento técnico (C1)" removido. A8: informação do lote 52 (Rafael ~R$ 50.000/ha em terras próximas × R$ 28.000/ha). A9: linguagem — nomes curtos dos bens, sem códigos de tabela, legenda, bases A e B com valores e aviso da mudança de B.
- Saídas: `_scripts/pacote_01.py` md5 `fa43947496`; `_scripts/build_mestre_02_05.py` md5 `5c8f110d28`; `pacotes/pacote-01.md` md5 `c10a581255`. Correções não passaram por segunda revisão isolada do pacote.

## 2026-10-28 — pacote 1 aprovado; início do levantamento "dinheiro das vendas da São Crispin"
- Luís: "sigo suas recomendações" → pacote 1 aprovado (R1, R2, R3) e base B do benchmark 2018. Gravado em `decisoes.json` (md5 `33c7a8c140`) e `log-decisoes.md` (md5 `566f64b0ad`); mestre refeita (`build_mestre_02_05.py` md5 `a4c24ad45e`): valor adotado nos itens 02, 03, 04, 07; R2 e R3 com decisão; Resumo mostra a base B escolhida.
- Pedido do Luís: "determinar quanto a empresa recebeu com a venda de áreas, quanto já foi retirado e quanto já foi pago". Leitura de trabalho: (1) quanto foi contratado e efetivamente recebido em cada venda (dinheiro, soja, gado), (2) quanto saiu da empresa como retirada de sócios, (3) quanto foi pago com esse dinheiro (despesas, dívidas, compras) ou, se for o sentido, quanto os compradores já pagaram. Apresentar as três leituras se as fontes permitirem.
- Fontes novas lidas (Drive `Documentos`, md5 10 dígitos): `retiradas_socios_2018.pdf` `96c512ef14` (1 p., manuscrito "Quarta Reunião", 14/1/19: retiradas pessoais desde 2011 — Adélia 640.000 cobertura, Raul 830.000 casa/apto, Élcio 330.000 faculdade, Pedro 600.000, Valter "580,00" [ilegível se mil], Irineu "não sabe valor"; Irineu/Raul "têm dinheiro a juros na empresa 1,3% a 1,5%"); `quantidade soja 2025` (jpg) `9f8fcf84db` = foto das OBSERVAÇÕES de soja do docx de áreas vendidas (mesmos números); `Cobrança Siviero` (jpg) `cd96a19e7d` = lista manuscrita de cobranças (não cita a fazenda); `relatórios empresas 2025.jpg` `672065c28a` = caixa jan–set/2025 por empresa (Hotel Princesa: receitas 1.390.404,48; despesas 1.401.648,13; resultado −11.243,65), sem a Agropecuária.
- Varredura por script de todas as fontes com texto (01_fontes, Drive Material Siviero e Documentos, OCR das respostas, análise de ACERTOS) por RIGO / PELLEGRINI / ANGELI / SOJA / SACAS / CRISPIN / AGROPECU / RETIRAD / LORENZ: em andamento (scratchpad `search_sales.py` → `hits.tsv`).

## 2026-10-28 — levantamento "vendas da São Crispin: recebido, retirado, pago" (rascunho)
- Varredura das fontes (scratchpad `search_sales.py` → `hits.tsv`, 1.368 ocorrências). Relevantes: doc 17 2018 (condições das vendas); docx 2025 de áreas vendidas (FALTA RECEBER; soja); `01_fontes/outros/Contas_Pagas_2023-07_a_2024-10.xlsx` md5 `a29fc2f2eb` (transcrição dos relatórios mensais escaneados do Drive; conferido na imagem do relatório de 05/2024, p. 3: históricos truncados no próprio PDF); doc 25 2018 (retiradas de sócios para moradia/estudo, sem relação com as vendas); arq. 1 p. 5 (frase do Dr. Pablo "ainda estamos recebendo o pagamento de áreas vendidas por R$ 49.726,30/hectare").
- Relatórios mensais do Drive (2024-04 a 2024-10) md5: 695db3fe74, 71a486a8cb, b855395e79, 613eac4150, 00dc78c5b1, 0341e6ed88, 098a235ffa (sem texto; imagem).
- Novo script `_scripts/vendas_sao_crispin.py` → `levantamento-vendas-sao-crispin.md`. Conferências internas: parcelas de cada venda = preço da venda (Pellegrini: R$ 1.100.000 à vista + R$ 1.000.000 jun/2019 = R$ 2.100.000; os R$ 100.000 de 41 cabeças de gado são parte do que foi vendido, não forma de pagamento); soma = TOTAL RECEBIDO R$ 12.058.000.
- Números: pagável em 2018 R$ 4.610.000; depois R$ 7.448.000; FALTA RECEBER (Rigo 2021 e 2022 a R$ 80) R$ 3.224.000; vendas de soja em 2024 reconstituídas pelos impostos pagos pela SEI R$ 756.588 (1.300 sacas a R$ 123 "COLONIA SA"; 2.606 sacas a R$ 123; "SOJA LORENZET" R$ 190.028,75; R$ 86.121,25); escavadeira CAT 320 da SEI (R$ 887.000) — pagos no período R$ 573.209,46; imóvel Rua Flores da Cunha da SEI (R$ 1.010.000) — pagos R$ 305.603,76, dos quais R$ 78.750,74 com "CHEQUES SOJA TITTON".
- Divergências: soja Lorenzet "não vendida" (docx) × imposto sobre "SOJA LORENZET" vendida em 2024; soja Rigo "falta receber" × soja do lote 2021 vendida pela empresa (2.606) com saldo 17.544 (+ 20.150 de 2022 = 37.694).
- Próximo: revisor independente → `_diario/revisao-vendas-sao-crispin.md`.

## 2026-10-28 — revisor do levantamento das vendas e correções (v2)
- Relatório: `_diario/revisao-vendas-sao-crispin.md` — 2 altas, 8 médias, 1 baixa. Quadro de vendas e números de imposto confirmados nas imagens dos relatórios.
- Correções no `_scripts/vendas_sao_crispin.py` (v2): A1 vendas de 1.300 e 2.606 sacas tributadas como pessoa física (FUNRURAL 1,5 %, sem IRPJ/CSLL) × Agropecuária (2,05 % + IRPJ + CSLL) — separadas, com a pergunta "em nome de quem saíram as notas"; A2 só R$ 320.538 (2.606 sacas Rigo) ligados à venda das terras; Lorenzet = aluguel; 1.300 sacas = origem não identificada; A3 imóvel Flores da Cunha somado só pelas parcelas de Maria Tereza (R$ 303.043,76; a comissão de R$ 2.560 de Ricardo Dal Ponte saiu); A4 escavadeira: soja cobre ~1/3; financiamento 6 de 48 parcelas; soja Titton "a confirmar"; A5 sem afirmar que a soja de 2021 estava com a empresa; coincidência 2.606 = Lorenzet 2021; A6 recebido em duas partes (dinheiro R$ 5.610.000; soja em sacas), sem teto em R$; A7 seção "retirado" com as retiradas das contas pagas da SEI por sócio (R$ 1.511.500 no período, classificação pelo texto), frase do grupo sobre o caixa comum da SEI (arq. 1 p. 6), manuscrito de 14/01/2019 (juros a sócios), datas corrigidas (documentos posteriores às vendas); frase do Dr. Pablo movida para "recebido" como indício; A8 inventário 2018 sem dinheiro/recebíveis entre os 17 itens; A9 total dividido (terra São Crispin R$ 11.448.000; gado R$ 100.000; Angeli/Cotiporã R$ 510.000); 118.100 sacas, 37.500 equivalentes ao à vista; A10 resumo inicial em tabela, siglas trocadas por frases, impostos em anexo; A11 preço R$ 123 marcado como dedução; ITR 259 ha.
- Saídas: script md5 e levantamento md5 na linha seguinte. Correções não passaram por segunda revisão isolada.
- `_scripts/vendas_sao_crispin.py` md5 `8d66a1b5f7`
- `levantamento-vendas-sao-crispin.md` md5 `29804a0253`

## 2026-10-28 — resposta do Luís sobre os Rigo
- Luís: "Não recebemos" (o lado dele não recebe parcelas dos Rigo diretamente). Registrado no `log-decisoes.md`; levantamento atualizado (seções 2 e 4): se parcelas ainda são recebidas, são recebidas pelo grupo (empresa ou outros sócios). Pergunta sobre o sentido de "quanto já foi pago" continua sem resposta; o levantamento mostra as duas leituras.
- `_scripts/vendas_sao_crispin.py` md5 `48ee543397`
- `levantamento-vendas-sao-crispin.md` md5 `0e5f8f5d85`
- `log-decisoes.md` md5 `dda3c1b6d3`

## 2026-10-28 — pacote 2 (exceções, rascunho)
- Luís: "Seguir com o pacote 2".
- Novo script `_scripts/pacote_02.py` → `pacotes/pacote-02.md`. Recomendações: item 05 R$ 44.828.300 e item 06 R$ 440.000 como provisórios (iguais nas duas avaliações; Rafael os chama de preliminares e pede avaliação independente; pedir essa avaliação — favorece C0); redução de R$ 400.000 abatida como dedução dos itens 05/06 (número do próprio grupo; C0); item 01 Rafael R$ 17.555.000 (C0: fica; C1: faixas de Rafael p. 8–9 somam R$ 11,9–17,8 mi; riscos: sem memória de cálculo, faixa do terreno inclui o item 02, frase de R$ 1.500–2.500/m²; precedente Oleoplan da SEI); item 08 grupo R$ 54.000 (C0: vem para nós; FIPE ~R$ 54.745 / R$ 56.854).
- Total do Hotel com pacotes 1 e 2: R$ 79.060.957,50 (grupo R$ 75.410.201,00; Rafael R$ 79.491.957,50).
- Próximo: revisor independente → `_diario/revisao-pacote-02.md`.

## 2026-10-28 — revisor do pacote 2 e correções (pacote 2 v2)
- Relatório: `_diario/revisao-pacote-02.md` — 3 altas, 5 médias, 5 baixas; contas certas. Conferido por mim: log da sessão 01-08 (pacote 3, regra C1: "Nada é descontado do valor da 01 e da 08 (P5 + diretriz)") e texto f) do doc 05 ("deverão ser satisfeitas pela atual credora e proprietária … HOTEL PRINCESA DOS VALES LTDA").
- Correções em `_scripts/pacote_02.py` (v2): A1 C1 do item 01 recalculado só para os 720 m² (faixa de Rafael R$ 10,28–15,1 mi; Rafael R$ 2,455 mi acima do teto; teto R$ 13,3 mi se o prédio já incluir o terreno) — apresentado como escolha real C0 × C1; A2 redução de R$ 400.000 com opções (a) não abater (recomendada, coerente com 01/08) e (b) abater como caso extremo; A3 "quem recebe paga" removido, texto da fonte citado, efeito das duas leituras (+ R$ 333.333,33 ou − R$ 66.666,67); M1 remissões das concordâncias p. 14–15; M2 sem "provisório": adotar o valor comum e pedir avaliação aceitando qualquer direção, com consequências; M3 indícios separados do C1 (venda Rigo R$ 49.726,30/ha; 2018 R$ 44.000/22.000/ha; antigo R$ 130.000/ha; pergunta literal do Dr. Pablo); M4 valor de Rafael entra no total e na cota; razão 1,3000; M5 efeito na cota de cada decisão; totais brutos; B1 precedente Oleoplan com as diferenças; B2 RDB-4, pergunta da matrícula 1.412 (número da área averbada não citado: não conferido por mim); B3 explicação do grupo para a 00.760; B4 ônus de dissolução como aviso; B5 siglas explicadas no início.
- Total do Hotel com as recomendações e o pacote 1: R$ 79.460.957,50 (bruto). Correções não passaram por segunda revisão isolada.
