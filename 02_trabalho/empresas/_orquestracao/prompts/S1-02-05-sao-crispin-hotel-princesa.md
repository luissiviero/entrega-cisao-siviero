Você estuda duas empresas do grupo Siviero para a cisão: **02 AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA.** e **05 HOTEL PRINCESA DOS VALES LTDA.** Eu, Luís, sou sócio de 1/6 e estou saindo. Trabalhe em português, na pasta `C:\entrega-cisao-siviero`. A SEI e outras empresas seguem em sessões paralelas.

## 1. Objetivo

Classificar cada bem das duas empresas em concorda × discorda entre a avaliação do grupo (inventário 2025, data-base 31.12.2024) e a de Rafael (Estudo 2), propor o valor a adotar item a item e documentar o destino de todo bem que existia em 2018 e não existe em 2025. As duas ficam juntas porque a Fazenda São Crispin está registrada em nome do Hotel: em 2018 era avaliada no inventário da Agropecuária; em 2025, no do Hotel.

Porte: Hotel, 8 itens, R$ 75.410.201,00 pelo grupo e R$ 79.491.957,50 por Rafael. Agropecuária, inventário 2025 zerado (R$ 0,00), mas R$ 12.054.390,00 líquidos em 2018.

## 2. Antes de começar, leia

- `00_PLANO.md` inteiro (premissas P1 a P13, seção 7 com as lições de método e a regra de arquivos).
- `02_trabalho/log-decisoes.md` e o fim de `02_trabalho/_diario/diario.md`: é o método aplicado na SEI, com as correções que o revisor exigiu.
- `02_trabalho/empresas/_orquestracao/analise-coordenacao.md`: levantamento da coordenação, seções 2 a 4.
- A memória do projeto traz as regras de trabalho. Siga todas.

Premissas que mais pesam aqui (o resto está no plano):
- **P1:** 1/6 em tudo.
- **P2:** data-base 31.12.2024. Arrendamentos, safras e vendas de soja depois dessa data vão para o registro de data-base, nunca ajustados em silêncio. Atenção: o grupo afirma que os inventários se referem a um "termo final adotado em prévio consenso como sendo outubro/2025" (respostas ao Dr. Pablo, arquivo 1, p. 1–2). Onde isso mudar um número, registre; não troque a data-base.
- **P3 e P4:** 2018 é benchmark de existência. Como a fazenda mudou de empresa, use os inventários 2018 da Agropecuária e do Hotel **somados** como benchmark das duas.
- **C0 (P10):** quatro bens do Hotel estão na lista de Rafael e recebem o **menor** valor: Sede Campestre (item 04), Fazenda São Crispin 407,53 ha (item 05), 4 ha da matrícula 11.090 (item 06) e Mercedinha MB 709 (item 08; R$ 54.000 no grupo e R$ 85.000 na lista). Os demais (itens 01, 02, 03 e 07) recebem o **maior**. Registre o argumento técnico (C1, faixa de mercado de Rafael) de cada escolha.

## 3. Fontes

Caminhos relativos a `C:\entrega-cisao-siviero\`.

- Inventário 2025: `01_fontes/inventario-2025/02. INVENTÁRIO - 31.12.2024 - AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA..docx` e `05. INVENTÁRIO - 31.12.2024 - HOTEL PRINCESA DOS VALES LTDA..docx`; docx `HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx` na mesma pasta.
- Versão antiga: `01_fontes/inventario-2025-antigo/ANTIGO 05. … HOTEL PRINCESA DOS VALES LTDA.(2).docx` (não há antigo da 02). Resumo das mudanças: `02_trabalho/diff-inventario-2025-antigo-x-atual.md`, seções 4 e 7 (de R$ 130.000/ha para R$ 110.000/ha: −R$ 8.150.600,00 na fazenda e −R$ 80.000,00 nos 4 ha; −R$ 8.230.600,00 no total).
- Inventário 2018, só no Drive: `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\17. 1. AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA. = …docx` e `13. 1. HOTEL PRINCESA DOS VALES LTDA. = …docx`. Leia do Drive e registre caminho e hash no seu diário. A lista "bens relacionados nos inventários de outras empresas" do inventário 2018 da SEI (`01_fontes/inventario-2018/18. …xlsx`, aba "Bens em outras empresas") também cita bens do Hotel e da Agropecuária.
- Rafael: `01_fontes/estudo-rafael/HOTEL PRINCESA DOS VALES LTDA- BELISE LTDA.pdf` e `AGROPECUARIA E FAZENDA SÃO CRISPIN LTDA - BELISE LTDA.pdf`; PDF Diretrizes na mesma pasta; lista de Rafael em `01_fontes/outros/ativos-avaliados-valter.xlsx` (linhas da empresa 05). Lição 1 do plano: o valor de Rafael pode estar no balanço executado, no quadro-resumo, nas tabelas de bens escolhidos ou no Diretrizes. Procure em todos antes de dizer que Rafael não avalia.
- Planilhas Google: abas `2.São Crispin` e `5.Hotel Princesa`, aba `Extras` e aba `GERAL` (linhas 11–15, 28–69, 941–950 com as notas do Luís sobre as vendas, 986–987) em `01_fontes/planilhas-google/2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx`; observações do Luís em `02_trabalho/observacoes-planilhas-google.xlsx`.
- Conferência anterior: `01_fontes/outros/conferencia-discrepancias.md` e `inventario-2025-avaliacoes-corrigido.xlsx` (abas "Vendas e créditos" e "Correções e pendências"; ela não cobriu estas duas empresas como ativos). `02_trabalho/pendencias-conferencia-anterior.md`, pendências 1 e 2.
- Respostas ao Dr. Pablo, `01_fontes/respostas-docs-2026-08-13/`: arquivo 1, p. 5–6 (Anexo B, São Crispin) e p. 8–9 (Anexo D, Hotel; explicação da matrícula 1.412 e da 00.760); arquivos 7, 8 e 9 (arrendamentos Lorenzet 2017–2020, 2020–2023, 2023–2026); arquivo 11 (matrícula 1.412). São PDFs escaneados: leia as páginas como imagem. Para tabelas em PDF pode usar a skill `pdf-to-excel`.
- **Análise paralela de ACERTOS E ACORDOS e das áreas vendidas:** `02_trabalho/analises/acertos-e-areas-vendidas/achados-para-empresas.md`. Use só os achados que eu tiver aprovado; se não houver registro de aprovação, pergunte. Nunca escreva naquela pasta. Se os achados ainda não existirem, faça a Fase 0 e a planilha-mestre e deixe as questões de área e de soja para depois deles.

## 4. Temas e riscos conhecidos (confirme cada um nas fontes)

1. **Área da fazenda.** 2025 (Hotel, item 05): 491,0 ha em 10 linhas de matrícula (9 números, a 00.760 repetida), menos 83,47 ha vendidos aos Rigo = 407,53 ha, a R$ 110.000/ha = R$ 44.828.300,00. 2018 (Agropecuária): 686,75 ha − 190,00 ha (Rigo) − 85,32 ha (Pellegrini) = 411,43 ha. O texto de 2025 fala em 411,4 ha. Pista do revisor da coordenação, a confirmar: no inventário 2018 da Agropecuária, os 190 ha dos Rigo são 83,47 ha da mat. 00.760 mais 106,53 ha da **mat. 11.090**, a mesma matrícula dos 4 ha do item 06 do Hotel; e as quatro parcelas dos Pellegrini (4,9 + 10,0 + 26,16 + 44,16) somam 85,22 ha, não 85,32. Com isso, 407,53 + 4,0 = 411,53 ha fecha com 686,75 − 190 − 85,22. A matrícula 11.090 usada em duas parcelas merece pergunta ao grupo.
2. **Matrícula 00.760** listada duas vezes (48 ha e 100 ha). O grupo explica que são compras distintas e que restam 64,53 ha nela (arquivo 1, p. 9). Confira contra as áreas da lista. Rafael registra a mesma dúvida e recomenda avaliação independente da fazenda (PDF do Hotel, p. 13–14).
3. **Vendas de 2018.** Rigo: 190 ha por R$ 9.448.000,00 (R$ 3.000.000,00 à vista + 4 × 20.150 sacas, 2019 a 2022). Pellegrini: 85,32 ha por R$ 2.100.000,00 (inclui 41 cabeças de gado). Angeli: 10,2 ha por R$ 510.000,00. O "líquido" de 2018 (R$ 12.054.390,00) é o bruto de R$ 24.112.390,00, que inclui os 491 ha inteiros e as parcelas vendidas, menos R$ 12.058.000,00 "recebidos das áreas vendidas". Risco de desconto em duplicidade (forma 6 da Parte 6): área reduzida pela venda e valor da venda descontado de novo.
4. **Soja a receber, sem valor em R$:** Rigo 37.694 sacas; Lorenzetti 13.085 sacas. Candidatos à Parte 5 (não abordados).
5. **Arrendamento à família Lorenzet:** 155 ha × 18 sacas, 30 ha × 30 sacas, 157,5 ha de campo a 3.000 kg de boi por 100 ha. O grupo diz que o arrendamento Hotel × Agropecuária é "pro forma" e que a receita vai para o fluxo da SEI (arquivo 1, p. 6). O Dr. Pablo aponta que áreas produtivas e não produtivas têm o mesmo valor por hectare (p. 5).
6. **Incorporação da Agropecuária ao Hotel:** Rafael registra como informação "pendente de validação formal" (PDF da Agropecuária, p. 3–4; PDF do Hotel, p. 6–7). Verificar se há documento. O PDF de Rafael da Agropecuária não tem balanço nem total: não há "valor de Rafael" para ela.
7. **Regularização de R$ 400.000,00** (ITBI, georreferenciamento, cartório) descrita igual nos inventários 02 e 05. Confira se é aplicada uma vez só.
8. **Sede Campestre (mat. 1.941 e 7.104) e lote 52 Bassani (mat. 7.391):** em 2018 a Sede Campestre era o item 04 do próprio inventário do Hotel (R$ 3.879.091,80), e a lista "bens em outras empresas" da SEI remete a ela; o lote 52 estava no inventário 2018 da Agropecuária (item 06) e no antigo 2025 da SEI, retirado da versão atual. Risco de dupla contagem; é o candidato nº 2 da SEI em `02_trabalho/parte6-candidatos-duplicidade.md` (só leitura).
9. **Matrícula 1.412:** itens 01 (1.800 m²) e 02 (1.080 m²) com a mesma matrícula; o grupo diz que o edifício ocupa 720 m² e a avaliação foi fracionada (arquivo 1, p. 8–9). Confira que nenhuma parte do terreno é avaliada duas vezes (arquivo 11).
10. **Bens da Agropecuária de 2018 que mudaram de empresa:** lotes rurais 12, 89 e 33 estão na SEI em 2025 (o lote 12, mat. 3.062, tinha 20,5 ha em 2018 e aparece com 13,0 ha na SEI); lotes 63, 65, 66, "N" e parte do 63 (Cotiporã) estão na Transportadora; trator Valtra BM100 (R$ 116.600,00) e benfeitorias da fazenda (R$ 300.000,00: cercas, galpão, implementos) sem destino conhecido; Mercedinha passou ao Hotel. Cada um precisa de destino (P4). Os que cruzam com outra empresa vão para os seus candidatos à Parte 6.
11. **Diferença total de R$ 4.081.756,50** entre Rafael e o grupo no Hotel. Pela conferência da coordenação, vem do item 01 (Rafael R$ 17.555.000,00 × grupo R$ 13.504.243,50) e da Mercedinha (+R$ 31.000,00). Confirme item a item.
12. **Todas:** venda das quotas de Rovílio Siviero (28/08/2025, depois da data-base). Nesta empresa, só registre se o texto aparece e se muda alguma linha. A análise do contrato é feita pela sessão da Abastecedora e Transportadora.

## 5. Sequência de trabalho

Não comece uma etapa sem o meu OK na anterior.

1. **Fase 0 das duas empresas.** Liste as fontes com hash. Extraia por script os itens de 2018 (as duas), 2025 atual e antigo, Rafael (todas as seções do PDF e a lista) e as observações das planilhas Google. Confira cada total extraído contra o total do documento. Mostre um resumo curto.
2. **Planilha-mestre** `planilha-mestre-02-05.xlsx`, com as colunas da seção 0.7 do plano e a coluna Empresa. Gerada por script em `_scripts/`; nunca editada à mão. Pares 2018 × 2025 por matrícula, área e vendedor, nunca só por semelhança de texto (lição 3).
3. **Pacote 1: regras em bloco** para o que concorda, com uma amostra de cerca de 5 linhas para eu conferir.
4. **Pacotes de exceções**, até 10 itens cada, maior valor primeiro. Cada item mostra valor do grupo, valor de Rafael, recomendação, motivo (C0, C1, C2 quando houver) e página da fonte.
5. **Revisor independente antes de cada pacote** (P11): agente de contexto novo que recebe as fontes e os arquivos produzidos, não as suas conclusões, e tenta achar erros. Corrija, registre no diário e diga o que mudou.
6. **Destino 2018 → 2025** de todo bem que saiu (P4). Bem sem destino documentado é pendência, não duplicidade.
7. **Ficha de fechamento por empresa** (uma página: decisões, números, pontos abertos), depois de uma nova revisão independente de todas as decisões aprovadas. Eu aprovo cada empresa separadamente.

Números sempre por script, nunca de memória (P12). Ao retomar depois de compactação, leia o seu diário e o seu log antes de agir.

## 6. Arquivos

Você escreve só em `02_trabalho/empresas/02-05-sao-crispin-hotel-princesa/`. Lá ficam, próprios desta sessão:
- `_diario/diario.md` (só acréscimos: entradas com hash, saídas, contagens, conferências, pendências) e `_diario/revisao-*.md` (relatórios dos revisores);
- `log-decisoes.md` e `decisoes.json` (decisões aprovadas, aplicadas pelo script);
- `perguntas-ao-grupo.md`, `parte6-candidatos.md`, `registro-data-base.md`;
- `planilha-mestre-02-05.xlsx`, `_scripts/`, `pacotes/`, `ficha-fechamento-02.md` e `ficha-fechamento-05.md`.

Todo o resto é só leitura: `01_fontes/`, os arquivos da raiz de `02_trabalho/` (são da sessão SEI), as pastas das outras empresas, `02_trabalho/analises/`. Bens que cruzam com a SEI ou com outra empresa vão para os **seus** candidatos à Parte 6, com a referência do outro lado. Não altere arquivos de memória existentes; pode criar novos. O material desta sessão só é unido aos arquivos da SEI depois que eu aprovar cada empresa.

## 7. Idioma

Português em tudo o que eu leio: documentos, planilhas, pacotes, fichas. Inglês só no código.
