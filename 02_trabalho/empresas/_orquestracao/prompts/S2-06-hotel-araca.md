Você estuda a empresa **06 HOTEL ARAÇÁ LTDA.** (Capão da Canoa – RS) para a cisão do grupo Siviero. Eu, Luís, sou sócio de 1/6 e estou saindo. Trabalhe em português, na pasta `C:\entrega-cisao-siviero`. A SEI e outras empresas seguem em sessões paralelas.

## 1. Objetivo

Classificar cada bem do Hotel Araçá em concorda × discorda entre a avaliação do grupo (inventário 2025, data-base 31.12.2024) e a de Rafael (Estudo 2), propor o valor a adotar item a item e documentar o destino de todo bem que existia em 2018 e não existe em 2025.

Porte: 22 itens; R$ 58.076.823,20 pelo grupo e R$ 76.230.000,00 por Rafael. São R$ 18.153.176,80 de diferença, a maior do grupo fora da SEI. O item 01 (o hotel, 18 terrenos na quadra 37) sozinho vale R$ 49.606.823,20 pelo grupo.

## 2. Antes de começar, leia

- `00_PLANO.md` inteiro (premissas P1 a P13, seção 7 com as lições de método e a regra de arquivos).
- `02_trabalho/log-decisoes.md` e o fim de `02_trabalho/_diario/diario.md`: é o método aplicado na SEI, com as correções que o revisor exigiu.
- `02_trabalho/empresas/_orquestracao/analise-coordenacao.md`: levantamento da coordenação.
- A memória do projeto traz as regras de trabalho. Siga todas.

Premissas que mais pesam aqui (o resto está no plano):
- **C0 (P10):** cinco lotes estão na lista de Rafael e recebem o **menor** valor: item 02 (lote 03/q19, "Casa da Praça", R$ 400.000), itens 11 e 12 (lotes 09 e 10/q38, R$ 390.000 cada), itens 18 e 19 (lotes 27 e 28/q38, "Estacionamento do Bigode", R$ 350.000 cada). Os outros 17 itens, incluindo o hotel, recebem o **maior**. Com o hotel nessa regra, o C0 aponta para Rafael na maior parte da diferença: a escolha exige argumento técnico escrito (C1, faixa de mercado de Rafael) e, se o valor de Rafael não tiver base, trate como exceção para eu decidir, não como regra.
- **P5:** sem reavaliação própria. Escolha entre os dois valores; caso extremo exige justificativa escrita.
- **P3 e P4:** 2018 é benchmark de existência, não de valor.
- **P2:** data-base 31.12.2024. Atenção: o grupo afirma que os inventários se referem a um "termo final adotado em prévio consenso como sendo outubro/2025" (respostas ao Dr. Pablo, arquivo 1, p. 1–2). Onde isso mudar um número, registre; não troque a data-base.

## 3. Fontes

Caminhos relativos a `C:\entrega-cisao-siviero\`.

- Inventário 2025: `01_fontes/inventario-2025/06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx`.
- Versão antiga: `01_fontes/inventario-2025-antigo/ANTIGO 06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx`. Resumo em `02_trabalho/diff-inventario-2025-antigo-x-atual.md`, seção 8: a "vantagem da coisa feita" da obra civil caiu de 25 % para 15 % (−R$ 3.713.636,80).
- Inventário 2018, só no Drive: `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\16. 1. HOTEL ARAÇÁ LTDA. = …docx` (25 itens; total R$ 49.684.243,20 "+ o valor da casa e lote – Élcio Siviero"). Leia do Drive e registre caminho e hash no seu diário.
- Rafael: `01_fontes/estudo-rafael/HOTEL ARAÇÁ LTDA - BELISE LTDA.pdf` (26 páginas, com avaliação item a item, balanço executado nas p. 16–20 e quadro-resumo); PDF Diretrizes na mesma pasta; lista de Rafael em `01_fontes/outros/ativos-avaliados-valter.xlsx` (linhas da empresa 06). Lição 1 do plano: procure o valor de Rafael em todas as seções antes de dizer que ele não avalia um bem.
- Planilhas Google: aba `6.Hotel Araçá` e aba `GERAL` (linhas 70–104) em `01_fontes/planilhas-google/2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx` e na planilha "…empreendimentos e Rafael…" (aba "Rafael × Inventário"); observações do Luís em `02_trabalho/observacoes-planilhas-google.xlsx`. **Não use a GERAL como fonte dos valores de 2018 desta empresa:** a linha 99 (lote 28/q38) mostra 2018 = 0, mas valia R$ 200.000,00, e a linha 97 (lote 15/q38) mostra 0 porque em 2018 era o lote com nome de sócio (tema 5). Por isso a linha 70 dá R$ 49.484.243,20 em vez dos R$ 49.684.243,20 do documento.
- Conferência anterior, que já cobriu o Hotel Araçá: `01_fontes/outros/conferencia-discrepancias.md`; `01_fontes/outros/inventario-2025-avaliacoes-corrigido.xlsx` (abas Ativos, Composição, Passivos e contingências, Correções e pendências). Reaproveite, mas confira contra as fontes. `02_trabalho/pendencias-conferencia-anterior.md`, pendência 6.
- Respostas ao Dr. Pablo, `01_fontes/respostas-docs-2026-08-13/`: arquivo 1, p. 10–11 (Anexo E); arquivo 12 (outorga do poço). PDFs escaneados: leia as páginas como imagem. Para tabelas em PDF pode usar a skill `pdf-to-excel`.

## 4. Temas e riscos conhecidos (confirme cada um nas fontes)

1. **Composição do item 01.** Segundo o Anexo E: terreno 6.900 m² a R$ 1.000/m²; móveis de 36 apartamentos novos e de 108 com depreciação; obra civil de 10.500 m² (CUB dez/2022, depreciação de 20 %); equipamentos; e mais 15 % "da obra feita" (R$ 5.570.455,20). O grupo defende a "vantagem da coisa feita" pela localização em frente ao mar e por índices urbanísticos difíceis de aprovar hoje (arquivo 1, p. 11). Compare parcela a parcela com Rafael.
2. **Regularização sem habite-se:** INSS da obra estimado em R$ 1.601.357,90, e o inventário fala numa redução de cerca de 4 % = R$ 2.000.000,00 do item 01. Nenhuma das duas está deduzida do total. Não há projeto nem profissional contratado (arquivo 1, p. 11). Verifique se Rafael desconta algo disso, para não descontar duas vezes.
3. **Matrícula 73.482 repetida** nos lotes 27 e 28/q38 (itens 17 e 18 em 2018; itens 18 e 19 em 2025). A numeração dos itens mudou entre os anos: não pareie item 17 com item 17 (lição 3). Confirme se é matrícula única para dois lotes ou erro.
4. **Outorga do poço** (arquivo 12): existe, mas o grupo diz que o poço não está sendo usado para abastecimento (arquivo 1, p. 11). Verifique se entra em algum valor (equipamentos do item 01) e se Rafael o trata.
5. **Lotes de 2018 com nome de sócio.** Itens 20 a 25 de 2018 (lotes com área de 450, 360, 308,40 e 300 m²) têm, no lugar do valor, o nome de Válter, Irineu, Raul, Élcio, Adélia e Pedro, e há "casa e lote – Élcio Siviero" fora do total. Para cada um: é bem da empresa ou do sócio? Onde está em 2025? Pistas do revisor da coordenação, a confirmar:
   - o item 17 de 2025 (lote 15/q38, mat. 33.003, R$ 390.000,00) tem a mesma inscrição imobiliária (360.0038.015.000) do item 21 de 2018, o lote "IRINEU SIVIERO": um lote em nome de sócio em 2018 aparece valorado na empresa em 2025;
   - o item 23 de 2018 (Élcio, casa, esquina) tem destino anotado: trocado por um lote no Loteamento Aditivo Medianeira (cruza com a SEI);
   - os itens 22 (Raul) e 23 (Élcio) de 2018 são lotes vizinhos com casa (tema 8).
   Os itens 21 e 22 de 2025 (lotes 15 e 17/q37, matrículas 86.669 e 124.906, Osório) também precisam de origem.
6. **Valores dos lotes de 2018 para 2025:** em geral, de R$ 180.000–220.000 para R$ 360.000–390.000 (Casa da Praça de R$ 250.000 para R$ 400.000; lotes do Bigode de R$ 200.000 para R$ 350.000). Não é tema de valor (P3), mas confira identidade lote a lote pela matrícula e pela inscrição imobiliária.
7. **Lote 15/q02 (Condado de Capão, 725,39 m², mat. 69.488, item 20):** grupo R$ 1.000.000,00, Rafael R$ 2.400.000,00, a segunda maior diferença. Comprado por contrato particular com móveis, R$ 750.000,00 quitados, com casa de 158,18 m² (Anexo E). Verifique a titularidade registral. Pelo C0 (bem que fica na empresa) aponta para Rafael: precisa de argumento técnico.
8. **Bens fora do inventário apontados por Rafael.** Em "Observação importante" (PDF, p. 24–25), Rafael diz que não constam do quadro patrimonial duas residências de cerca de 200 m² em dois lotes na esquina da Rua da Divisa com a Rua Três (bairro Araçá), com faixa de R$ 2.400.000 a R$ 3.500.000 pelo par. Confira se ele as soma no total de R$ 76.230.000,00, se correspondem aos lotes com nome de sócio de 2018 (tema 5) e a quem pertencem. Se não estiverem em nenhum inventário, são candidatas à Parte 5 (não abordados).
9. **Erro de soma de Rafael.** Pela conferência da coordenação, os itens do balanço executado de Rafael somam R$ 75.870.000,00, mas o total declarado é R$ 76.230.000,00 (p. 20): R$ 360.000,00 sem item correspondente. A parte explicável da diferença é R$ 17.793.176,80: item 01 (Rafael R$ 66.000.000,00, +R$ 16.393.176,80) e item 20 (+R$ 1.400.000,00); os demais lotes são iguais nos dois. Confirme e não procure um bem para os R$ 360.000,00 antes de confirmar que é erro de soma; se for, registre como na Oleoplan da SEI (o outro lado pode alegar erro).
10. **Todas:** venda das quotas de Rovílio Siviero (28/08/2025, depois da data-base). Nesta empresa, só registre se o texto aparece e se muda alguma linha. A análise do contrato é feita pela sessão da Abastecedora e Transportadora.

## 5. Sequência de trabalho

Não comece uma etapa sem o meu OK na anterior.

1. **Fase 0 da empresa.** Liste as fontes com hash. Extraia por script os itens de 2018, 2025 atual e antigo, Rafael (todas as seções do PDF e a lista) e as observações das planilhas Google. Confira cada total extraído contra o total do documento. Mostre um resumo curto.
2. **Planilha-mestre** `planilha-mestre-06.xlsx`, com as colunas da seção 0.7 do plano. Gerada por script em `_scripts/`; nunca editada à mão. Pares 2018 × 2025 por matrícula e inscrição imobiliária, nunca só por semelhança de texto (lição 3). Confira no nível do item e no nível de cada componente do item 01 (lição 2).
3. **Pacote 1: regras em bloco** para o que concorda, com uma amostra de cerca de 5 linhas para eu conferir.
4. **Pacotes de exceções**, até 10 itens cada, maior valor primeiro. Cada item mostra valor do grupo, valor de Rafael, recomendação, motivo (C0, C1) e página da fonte.
5. **Revisor independente antes de cada pacote** (P11): agente de contexto novo que recebe as fontes e os arquivos produzidos, não as suas conclusões, e tenta achar erros. Corrija, registre no diário e diga o que mudou.
6. **Destino 2018 → 2025** de todo bem que saiu (P4). Bem sem destino documentado é pendência, não duplicidade.
7. **Ficha de fechamento** (uma página: decisões, números, pontos abertos), depois de uma nova revisão independente de todas as decisões aprovadas.

Números sempre por script, nunca de memória (P12). Ao retomar depois de compactação, leia o seu diário e o seu log antes de agir.

## 6. Arquivos

Você escreve só em `02_trabalho/empresas/06-hotel-araca/`. Lá ficam, próprios desta sessão:
- `_diario/diario.md` (só acréscimos: entradas com hash, saídas, contagens, conferências, pendências) e `_diario/revisao-*.md` (relatórios dos revisores);
- `log-decisoes.md` e `decisoes.json` (decisões aprovadas, aplicadas pelo script);
- `perguntas-ao-grupo.md`, `parte6-candidatos.md`, `registro-data-base.md`;
- `planilha-mestre-06.xlsx`, `_scripts/`, `pacotes/`, `ficha-fechamento-06.md`.

Todo o resto é só leitura: `01_fontes/`, os arquivos da raiz de `02_trabalho/` (são da sessão SEI), as pastas das outras empresas, `02_trabalho/analises/`. Bens que cruzam com a SEI ou com outra empresa, ou com um sócio, vão para os **seus** candidatos à Parte 6, com a referência do outro lado. Não altere arquivos de memória existentes; pode criar novos. O material desta sessão só é unido aos arquivos da SEI depois que eu aprovar a empresa.

## 7. Idioma

Português em tudo o que eu leio: documentos, planilhas, pacotes, fichas. Inglês só no código.
