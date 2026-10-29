# Diário de trabalho — uso interno de Claude e do revisor

Somente acréscimos. Cada entrada: data, passo, entradas (MD5 10 dígitos), saídas, contagens e conferências, pendências. Para retomar: ler `00_PLANO.md`, o fim deste diário, `log-decisoes.md` e `decisoes-sei.json`; números sempre re-derivados das fontes por script (`_scripts/build_mestre_sei.py`), nunca da memória.

## 2026-10-28 — retroativo: Fase 0 e pacotes 1 a 3 da SEI

### Entradas (hash na data deste registro)

- `01_fontes/inventario-2025/11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx` md5 `5c9d13d0d9`
- `01_fontes/inventario-2018/18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.docx` md5 `a9097fa4ec`
- `01_fontes/inventario-2018/18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.xlsx` md5 `8c63e21506`
- `01_fontes/outros/ativos-avaliados-valter.xlsx` md5 `935282c1aa`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/00 - RESUMO BELISE LTDA.xlsx` md5 `caf72e6376`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/DEMIAS COMPROMISSOS A PAGAR, BELISE LTDA.xlsx` md5 `03648fc4aa`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/LOTEAMENTO BONETTO  BELISE LTDA.xlsx` md5 `fac7b04862`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/LOTEAMENTO COLINA DO RETIRO II BELISE LTDA.xlsx` md5 `876d5dc0b3`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/LOTES DE TERCEIROS INDENIZAVEIS, BELISE LTDA.xlsx` md5 `6e70c914da`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.xlsx` md5 `ca58713b9c`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/LOTES URBANOS EM VERANÓPOLIS RS - BELISE LTDA.xlsx` md5 `db1ad58136`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO CAMINHOS DO SOL,  BELISE LTDA.xlsx` md5 `ce58270ef1`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO DA FAMÍLIA CECCONELLO – CAXIAS DO SUL, BELISE LTDA.xlsx` md5 `75cd53bf08`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO DAS FAMILIAS PEDROTTI, ANDREIS ,LORANDI,  BELISE LTDA.xlsx` md5 `0c7b9bf4a4`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO DR LETTI E SOL NASCENTE III, BELISE LTDA.xlsx` md5 `13af411e1c`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO PASQUAL,  BELISE LTDA.xlsx` md5 `025d26bccf`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO VALE DAS COLINAS,  BELISE LTDA.xlsx` md5 `d4827994c2`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO VISTA DA COLINA,  BELISE LTDA.xlsx` md5 `bc4c784229`
- `01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL,  BELISE LTDA.xlsx` md5 `67838bd992`
- `01_fontes/planilhas-google/2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx` md5 `11a81ef0aa`
- `01_fontes/planilhas-google/ACERTOS E ACORDOS (gdoc).docx` md5 `8c0f668319`
- `01_fontes/planilhas-google/Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx` md5 `c499486a26`

### Saídas

- `02_trabalho/sei-2025-itens.xlsx` md5 `62ed740e2d`
- `02_trabalho/sei-2018-itens.xlsx` md5 `dcabdf2209`
- `02_trabalho/diff-inventario-2025-antigo-x-atual.md` md5 `7afde86365`
- `02_trabalho/observacoes-planilhas-google.xlsx` md5 `18bf9de495`
- `02_trabalho/decisoes-sei.json` md5 `32e1f28ea6`
- `02_trabalho/planilha-mestre-sei.xlsx` md5 `4e091742b7`
- `02_trabalho/_scripts/build_mestre_sei.py` md5 `dcdbf65c03`
- `02_trabalho/log-decisoes.md` md5 `0f42a57a74`
- `02_trabalho/registro-data-base.md` md5 `011fedc900`
- `02_trabalho/perguntas-ao-grupo.md` md5 `6f27d0195f`
- `02_trabalho/pendencias-conferencia-anterior.md` md5 `1e9eca2c47`

### Passos
1. Fase 0: estrutura da pasta; 104 fontes copiadas para `01_fontes` (índice `00_FONTES.md`); duplicados Drive × repositório decididos (canônico de Rafael = Drive `versão-excel`, 13 documentos).
2. Extrações por agente: doc 11 (SEI 2025) → 93 itens, 472 lotes, 39 compromissos, 42 passivos; totais conferem (bens 152.796.260,60; líquido 100.874.027,60). Doc 18 (SEI 2018) → 138 itens + 9 bens em outras empresas; total 108.428.000,00 confere.
3. Diff inventário 2025 antigo × atual (agente). Planilhas Google exportadas pelo conector; 236 observações extraídas.
4. Planilha-mestre gerada por `build_mestre_sei.py`. Casamento lote a lote com Rafael: 441 mesmo valor. Correção de casamento: Rafael escreve aviso de APP na célula da quadra; nomes com apelidos (Falavigha/Falavigna, Monte Vêneto I–II).
5. Achado: 7 imóveis listados em 2018 como "bens relacionados nos inventários de outras empresas" (Rádio, Oficina, Agropecuária, Hotel) entram no total da SEI em 2025 → verificar dupla contagem (Parte 6).
6. Pacote 1 aprovado (R1–R5). **Erro detectado depois:** R4 incluía 11 lotes de APP como "sem valor nas duas fontes", mas Rafael os valora no nível do loteamento (Simioni R$ 520.000; Lago Negro R$ 7.700.000). Corrigido antes do pacote 2 e registrado no log. Causa: conferência feita só no nível do lote. Lição: conferir item e lote.
7. Pacote 2 aprovado com ajustes; critério C0 (P10). Lista de Rafael extraída de `ativos-avaliados-valter.xlsx` (PDF Diretrizes): SEI 98 itens = 91 lotes (24.520.000) + 7 bens; confere com os quadros "Lotes escolhidos" dos documentos de empreendimentos (mesma soma). Grafias corrigidas: "Colina do Retiro" = Colina do Retiro I; "Altos da Graça I" = Alto das Graças I.
8. Pacote 3 aprovado (7 áreas rurais por Rafael, C0 + C1).

### Conferências numéricas (saída do script após o pacote 3)
- Bens 2025 pelo grupo: 152.796.260,60. Soma dos valores adotados: 153.970.060,60. Itens 2025 sem valor adotado: 1 (SEI-2025-CONTR-06, Bonetto).
- Lotes da lista de Rafael casados: 91/91; valor adotado 24.190.000 (diferença de 330.000 para a soma de Rafael = lote 001/q01 do Lago Negro pelo grupo, C0).
- Pacote 1: R1 395 lotes / 28 itens / 103.459.000; R2 12 itens / 8.560.666; R3 30 itens / 16.515.016; R4 8 itens; R5 13 itens.

### Pendências
- Revisão independente dos pacotes 1 a 3 (regra P11): a fazer.
- Pacote 4: 7 casamentos com 2018 feitos por semelhança de texto.
- Ficha de fechamento da Parte 3 da SEI.

## 2026-10-28 — pacote 3 registrado; regras P11/P12; revisões disparadas
- decisoes-sei.json: 7 áreas rurais (ARUR 03, 04, 05, 08, 09, 11, 12) por Rafael; LOT-18 confirmado (5.170.000). Soma adotada: 153.970.060,60; pendente só CONTR-06.
- Interpretação registrada da resposta do Luís ao alerta do C0: diferença de tratamento aceita; C0 mantido.
- Revisor 1 (pacotes 1–3) disparado: lê decisoes-sei.json, planilha-mestre-sei.xlsx, sei-2025-itens.xlsx, workbooks de Rafael, ativos-avaliados-valter.xlsx e PDFs gerais de Rafael. Saída esperada: _diario/revisao-pacotes-1-3.md.
- Pacote 4 (rascunho) em pacotes/pacote-04-rascunho.md: pares 2018×2025 de imóveis fora de loteamento. Erro do casamento automático encontrado ao preparar: ARUR-12 (Scalco, 2022) estava ligado ao lote 24 de 2018 (vendido em 2018). Revisor 2 disparado sobre o rascunho; saída _diario/revisao-pacote-04.md.
- Não apresentar o pacote 4 ao Luís antes do resultado do revisor 2.

## 2026-10-28 — revisor 2 (pacote 4) concluído
- Relatório: _diario/revisao-pacote-04.md. 2 altas, 5 médias, 7 baixas.
- Conferidas por mim nas fontes: A1 (Sede Campestre, mat. 1.941/7.104, em 2025 só no doc 05, item 04) e A2 (planilha-mestre liga ARUR-11 Piccolo a 2018 Áreas #08 Ansolin, sim 0.92). Ambas confirmadas.
- Versão revisada: pacotes/pacote-04.md (incorpora todas as correções). Rascunho mantido em pacotes/pacote-04-rascunho.md.
- Pendente: aplicar os pares aprovados na planilha-mestre (seção "casamentos" em decisoes-sei.json) depois da aprovação do Luís.
- Revisor 1 (pacotes 1–3) ainda em execução.

## 2026-10-28 — revisor 1 (pacotes 1–3) concluído
- Relatório: _diario/revisao-pacotes-1-3.md. 1 alta, 3 médias, 5 baixas.
- Conferido por mim: Quadro_Resumo "Patrimônio total" dos 7 documentos de áreas contratadas (Caminhos 8.601.000; Cecconello 9.165.000; Vale 5.019.600; Vista 3.384.000; Pedrotti 3.384.000; Pasqual 1.479.000; Dr. Letti/Sol Nascente III 0). Confirmado. Tabelas "Lotes escolhidos" com os 5 lotes de possível APP (p. 94, 100, 116), sem valor, com valor escrito ao lado. Confirmado.
- Causa raiz (2ª vez): o script lia só as abas Avaliacao/Balanco_Grupo; nos documentos de áreas contratadas o valor de Rafael está no Quadro_Resumo. Lição: para cada bem, procurar o valor de Rafael em todas as abas e documentos, não só na aba Avaliação.
- Script corrigido (patch_rev1): Quadro_Resumo dos contratados, valores do Diretrizes, reavaliação do Lago Negro, flag "só nas tabelas", casamento de passivos com palavra em comum. Soma adotada sem mudança: 153.970.060,60. Itens 2025 sem valor adotado: 7 (CONTR-01 a 07).
- Decisões que precisam do Luís: (a) os 6 contratados + Bonetto: adiar para a Parte 5 com regra única; (b) C2 × C0 nos 11 lotes de APP confirmada (R$ 8.220.000 segundo Rafael); (c) os 5 lotes de "possível APP" das tabelas de Rafael contam como lista de Rafael?
- Observações sem decisão agora: Oleoplan (risco de alegação de erro de soma, já no log); escavadeira 320D 2009 na SEI e no estudo de Rafael para a Ceval (R$ 240.000): dupla contagem possível, vai para a Parte 6.

## 2026-10-28 — decisões pós-revisão aplicadas; levantamento de APP disparado
- decisoes-sei.json: CONTR-01..05, 07 adiados para a Parte 5 (valor pendente); LOT-04 Simioni reaberto (pendente; lotes 006–009 pendentes); Lago Negro APP: C2 > C0 confirmado; Oleoplan reafirmado; seção "casamentos" (ARUR-11 e ARUR-12 sem par; ARUR-02 → SEI-2018-ARUR-03; ARUR-06 → SEI-2018-ARUR-08; ARUR-04 → SEI-2018-OUTRAS-s143); destinos de 12 itens de 2018.
- Script: chave única para "outras empresas" de 2018 (SEI-2018-OUTRAS-s<seq>), pares forçados, itens "outras empresas" sem par entram na lista de 2018 sem par, "Pendente." quando não há data de aprovação.
- Conferido após rebuild: pares de ARUR-02/04/06/11/12 corretos; 12 destinos preenchidos; soma adotada 153.970.060,60 (sem mudança); pendentes 8 (LOT-04 + CONTR-01..07).
- Simioni (conferido por mim): 2018 item 02, 4 lotes × R$ 250.000 = 1.000.000, com "restrições ambientais – Fonte de Água e Mato"; 2025 item 04, mesmo texto, sem valor; não está no arquivo 17; sem processo citado. Rafael: R$ 520.000, sem base de mercado.
- Criados: parte6-candidatos-duplicidade.md (5 linhas); perguntas 3–10.
- Agente de levantamento de APP disparado → levantamento-app.xlsx / .md. Depois: revisor independente sobre ele, antes de levar ao Luís.

## 2026-10-28 14:40 — levantamento de APP cancelado pelo Luís
- Agente parado antes de terminar. Ficou só `levantamento-app.xlsx`, gravado às 14:38, possivelmente incompleto e **não revisado**. Não há `levantamento-app.md`.
- Não usar o xlsx como base sem refazer ou revisar. As leituras intermediárias do agente estão no scratchpad da sessão, pasta `app\`, e se perdem com ela.
- Pendências abertas para a próxima sessão: regra única para "possível APP", com o Simioni incluído (Simioni reaberto: 2018 valia R$ 1.000.000 com o mesmo texto de restrição, sem processo citado em 2025); verificação de descontos duplos ambientais (Dalzóchio: lote próprio + lotes de terceiros + fonte de água nas observações); depois, ficha de fechamento da Parte 3 da SEI.

## 2026-10-28 — encerramento da primeira sessão
- Plano atualizado: seção 7 (estado, lições de método, regra P13 de propriedade de arquivos por sessão). README de 02_trabalho reescrito. Memória: siviero-session-structure.
- Prompts de abertura em 00_prompts/: 1-sei-continuacao.md, 2-coordenacao-demais-empresas.md, 3-apresentacao-powerpoint.md.
- Próxima sessão SEI começa pelo passo 1 do prompt 1 (refazer o levantamento de APP).
- Acrescentado prompt 4 (00_prompts/4-analise-acertos-e-areas-vendidas.md): sessão de análise do ACERTOS E ACORDOS.pdf (6 páginas escaneadas, sem texto; conferir contra a versão gdoc e a extração anterior) e do docx do Hotel (São Crispin, créditos em soja sem valor: 37.694 + 13.085 sacas; arrendamento; áreas contratadas pela SEI sem obras). Pasta própria: 02_trabalho/analises/acertos-e-areas-vendidas/. Prompts 1 e 2 e P13 atualizados para referenciá-la.

## 2026-10-28 (sessão SEI 2) — levantamento ambiental refeito
- Corpus de texto das fontes: `_scripts/corpus_fontes.py` → scratchpad da sessão (`corpus/`): 184 fontes, 60.085 unidades, 40 PDFs com páginas sem texto, 1 erro (Drive "22. 1. SÓCIOS… reunião 2017/2019.docx" não abre: arquivo não baixado).
- Novo script `_scripts/levantamento_ambiental.py` (md5 após correções na entrada seguinte) → `levantamento-ambiental-sei.xlsx`. Entradas: sei-2025-itens.xlsx (Lotes, Compromissos), arquivo 17 xlsx, xlsx de Rafael (executados, terceiros), xlsx 2018 (Processos). Cada citação é conferida no corpus; o script para se não achar.
- Saída: 44 lotes (4 Simioni + 7 Lago Negro APP + 7 possível APP + 26 terceiros), 11 itens, 10 processos, 9 sobreposições, 3 lotes com fato após a data-base; 56 citações conferidas.
- Conferências: terceiros 16.240.000 (doc 11 e Rafael) OK; possível APP 2.321.000 OK; "multas ambientais" do Dr. Pablo (obs. 2+3+5+6+7) 7.910.000 OK.
- Achados principais (a confirmar pelo revisor): TAC do MP no Simioni permite vender com aviso (arq17 p.2); 7 lotes de terceiros vendidos/dados pela SEI (Perini 11, 18, 19; Jardim Itália 01 e 02/1747; Primavera 11/1199 SAFRA; Vicentino 03/1301 por R$ 500.000 em set./2025; Vicentino 004/1302 em 2014); base dos 7 possíveis APP e dos 26 = Marco Regulatório de Farroupilha (Lei 4.814/2023, fonte externa: MP-RS lista como lei que reduziu APP); Lago Negro: R$ 580.000 pagos em 14/10/2025 e permuta Candiago em 28/08/2025 (após a data-base); Rafael cobra os 7 possíveis APP (inclusive Primavera, p.106) — a planilha-mestre marca Primavera 008/009 como fora da lista de Rafael (erro de marcação a corrigir, sem efeito em valor adotado); processo 048/1.08.0001924-5 é "Perini" em 2015 e "Dalzóchio" em 2018.
- Contas pagas jul/2023–out/2024: só despesas ambientais pequenas; nenhuma repete as contingências de 2025.
- Revisor independente disparado (fontes + xlsx + script, sem conclusões). Saída esperada: `_diario/revisao-levantamento-ambiental.md`. Não apresentar ao Luís antes.
- `levantamento-app.xlsx` (sessão 1) mantido sem uso; substituído por `levantamento-ambiental-sei.xlsx`.

## 2026-10-28 (sessão SEI 2) — revisão do levantamento ambiental e correções
- Revisor: `_diario/revisao-levantamento-ambiental.md`, 18 achados (2 altos, 8 médios, 8 baixos). Conferidos por mim nas fontes: multa de Cotiporã (contas, linhas 1777–3576: 36 × 701,19, última P.08 em out./2024); contrapartida do Vicentino 003/1301 = lote 623/104 do Reserva Paradiso (arq17, linha 64); compensação de 3 lotes do Lago Negro no doc 11 antigo (P129). A ART "LOTES S…" está cortada no relatório e fica a verificar.
- Todos incorporados no script: permutas do Lago Negro casadas pelo texto das observações; eventos extras (019/11B com R$ 580.000; 010/1199 SAFRA); valor implícito de Rafael (Lago Negro 1.100.000; Simioni 130.000); erro de Rafael no San Marino (7,7 mi, frase copiada); novas sobreposições (lote próprio na data-base como passivo de terceiro, R$ 1.887.000; Perini x Dalzóchio com o mesmo processo, R$ 7.200.000; critério de APP diferente na mesma quadra; Bonetto ativo + multa); Lago Negro reescrito (omissão de provisão); sem uso da lei municipal como premissa; textos de inferência marcados; itens novos (área rural 09, multa de Cotiporã, despesas pequenas, posição de Rafael sobre as obs. 1–24).
- Saída: levantamento-ambiental-sei.xlsx md5 `8f60e0bb2e` (44 lotes, 15 itens, 10 processos, 12 sobreposições, 10 fatos após a data-base, 71 citações conferidas; conferências OK). Script md5 `6298caec79`.
- Resumo para o Luís: `pacotes/levantamento-ambiental.md`. Registro de data-base: candidatos listados, nada lançado sem aprovação.
- OCR (tarefa intercalada, aprovada pelo Luís): Tesseract 5.4.0 instalado via winget (aprovado pelo Luís em AskUserQuestion); por.traineddata (tessdata 4.1.0) + eng/osd + configs em C:/Users/User/AppData/Local/tessdata. MD5 dos arquivos 2–16 conferem com 00_FONTES.md (100 páginas). Script `_scripts/ocr_respostas_docs.py` md5 `3858a25e0a` em execução. 1ª execução falhou (faltavam as pastas configs/tessconfigs). corpus_fontes.py (md5 `9ed179d192`) passa a usar o OCR nas páginas sem texto, que continuam em sem_texto.json; unidade marcada com "metodo": "OCR: ...".
- Recebido da sessão ACERTOS: pacote 1 aprovado pelo Luís (log-decisoes.md da pasta analises/acertos-e-areas-vendidas, linha 17). S1 (R$ 1.887.000) e S5 (R$ 360.000) coincidem com as sobreposições 1 e 6 do levantamento ambiental, achadas aqui de forma independente. A incorporar nas Partes 4–6: S1, S5 (Parte 6 + registro de data-base), S2, S3, S4 (lote Q6371 L11 na lista de Rafael), S6 (disposições de 2025 fora do doc 11), S7 (crédito Boff, Parte 5), S9. Nota no resumo pacotes/levantamento-ambiental.md.
- OCR concluído: ocr-respostas-docs/ocr.jsonl md5 `9b5b52b7c1`, 100 páginas, 15 .md; confiança < 70 em 4 p.16 (60), 4 p.17 (45), 5 p.2 (62), 6 p.1 (66), 15 p.8 (61); sem texto reconhecido: 4 p.15, 10 p.10, 13 p.6. Corpus refeito: 185 fontes, 64.214 unidades, 97 unidades OCR (marcadas "metodo"). Busca por termos ambientais no OCR: nada sobre os lotes de APP da SEI (arquivo 16 incluído). Levantamento regenerado (aba Fontes: arquivo 16 lido por OCR), md5 `95b041fefb`, mesmas contagens.
- Revisor por amostra do OCR disparado (imagens × texto, ~15 páginas, prioridade para números, manuscritos e baixa confiança). Saída esperada: _diario/revisao-ocr-respostas-docs.md.
- Revisão do OCR concluída: _diario/revisao-ocr-respostas-docs.md (22 páginas, cerca de 450 números). Contratos digitados: utilizáveis para busca; erro em números de 3 a 4%, às vezes com aparência de número válido. NFs (arq. 5 e 6), matrícula 11, selos, carimbos e imagens (arq. 4 p.15-17): mais de 50% de erro; ler a imagem. 10 p.10 e 13 p.6 estão em branco; 4 p.15 é imagem de satélite com rótulos de matrícula. Achado: a NF do arq. 6 tem anotação à mão "Maquina do Celito. Ceval 66,66 Siviero 33,33 Celito" (confirma o S12 da sessão ACERTOS, escavadeira 320D, Parte 6). Guia de uso: ocr-respostas-docs/LEIA-ME.md. Tarefa intercalada de OCR encerrada; aviso ao Luís e à sessão de organização.

## 2026-10-28 (sessão SEI 2) — levantamento aprovado; início do pacote 5
- Luís: "Aprovo, siga" (levantamento ambiental). Registrado em log-decisoes.md.
- Pacote 5: regra única de "possível APP" por nível de documento (A: decisão/perícia/gravame no lote; B: processo ambiental em curso na data-base sobre a área do loteamento; C: sem processo nem laudo, ou documento que permite a venda). Aplicação por script `_scripts/pacote05.py`.
- pacote05.py md5 `4abca9d331` → pacotes/pacote-05.md md5 `a7baf2bb78`. 10 itens por loteamento. Ativos +2.775.500 (7 possível APP: 2.255.500; Simioni 520.000); passivo de terceiros 16.240.000 → 913.000 (S1 retira 1.887.000; B 50%: Dalzóchio, Perini, Jardim Itália; C 0%: Primavera, Vicentino). Soma adotada 153.970.060,60 → 156.745.560,60. Efeito +18.102.500; 1/6 = 3.017.083,33.
- parte6-candidatos-duplicidade.md: itens 6 a 12 acrescentados; item 5 (Lago Negro) atualizado.
- Revisor do pacote 5 disparado; saída esperada _diario/revisao-pacote-05.md. Não apresentar antes.
- Revisor do pacote 5 (v1): _diario/revisao-pacote-05.md, 13 achados (2 altos). Achado 1 conferido por consulta ao GEOFarroupilha (serviços ArcGIS públicos da prefeitura, P9): Marco_Regulatorio/MapServer (camadas APP 50 m, APP 30 m, faixas não edificantes de 15 m e 5 m, Sensibilidade, Rev_MP_APPs com processo_adm) e Cadastro_Imobiliario (Lotes por inscrição setor.quadra.lote). SSL do servidor sem certificado intermediário: usado o repositório do Windows (truststore), sem desligar a verificação.
- Novo script geofarroupilha_app.py md5 `368b908d60` → geofarroupilha-app-lotes.xlsx md5 `0ee00303a8`; brutos em 01_fontes/geofarroupilha/. 95 lotes de Farroupilha localizados (quadra 1747 em 2 setores: usado o 24; quadra 910 no setor 06). % restrito = união das camadas 2 a 5 dentro do lote. 62 lotes próprios de valor cheio: nenhum com restrição; os lotes marcados pelo grupo, em geral, sim.
- pacote05.py reescrito (md5 `29b225fd8c`): nível pelo mapa (A ≥ 50 %; B 10 a 50 % ou Revisão MP com processo ≤ 2024; C < 10 %); Simioni B pelo TAC; asserts lote a lote de que os valores de Rafael = grupo. Resultado: ativos 1.187.500 (antes 2.775.500); passivo de terceiros 4.422.000 (antes 913.000); R6 +11.118.500 (1/6 1.853.083,33); S1 +1.887.000 separado; soma adotada → 155.157.560,60.
- pacote-05.md v2 md5 `f5b058866d` (todos os 13 achados tratados). Lei 4.814/2023: texto não obtido (busca web sem resultado); camadas lidas pelo nome.
- Segundo revisor disparado (mapa + regra v2); saída _diario/revisao-pacote-05-v2.md.
- Recebido da sessão S3 (Abastecedora/Transportadora): levantamento da venda das quotas do Rovílio, aprovado pelo Luís segundo a S3 (empresas/01-08-abastecedora-transportadora/rovilio-quotas.md, md5 c6ace6758d). A considerar na Parte 4 da SEI:
  - obs. 1 = R$ 1.650.000 (3 lotes do Ipê II "à pagar", contrato de 28/08/2025, posterior à data-base), dos quais R$ 550.000 são comissão; é dívida sem ativo correspondente;
  - Ipê I 001/q2 (R$ 538.000) saiu entre as versões do doc 11 sem explicação;
  - Ipê I 10/q02 (entregue em fev./2025) não consta de nenhuma versão;
  - pela cl. 4ª, todo o pagamento saiu de bens da SEI;
  - destinos 2018→2025: Ipê I 13/q02, 07/q04 e 10/q02; Guaporé 12/q237 e 15/q252; um lote do Dalla Costa.
  Item ainda não aprovado (não usar): mat. 3.754 (serraria, SEI) com 2.500 m² locados à SIM como estacionamento, com renda para a Abastecedora.
- Revisor v2: _diario/revisao-pacote-05-v2.md. Identificação e medidas confirmadas. Achado alto: a "Sensibilidade_Ambiental" é a faixa de 30 m em volta de todos os cursos d'água, e o grupo marcou exatamente os lotes que a tocam (33/33). "Revisão MP" não sustenta o nível B (sem descrição, processos de 2026). Custo frente ao C0 não declarado.
- geofarroupilha_app.py (md5 `25b8919623`) passa a calcular também a união das camadas 2 a 5 e 10 → xlsx md5 `37be4e4aa9`. pacote05.py (md5 `3c13466a91`) com três leituras (estrita, intermediaria, ampla), sem uso da Revisão MP. Resultados da R6: estrita +11.483.500; intermediária +6.970.500 (1/6 = 1.161.750,00); ampla +2.782.000 (iguais aos do revisor). S1 +1.887.000 separado. Custo frente ao C0 (intermediária): 1.653.500; 1/6 = 275.583,33.
- pacote-05.md v3 md5 `b560389ee0`: decisão central = leitura da faixa de 30 m; recomendada a intermediária. Pronto para o Luís.

## 2026-10-28 (sessão SEI 2) — pacote 5 aprovado e aplicado
- Luís: "Aprovo como recomendado" (R6 com extensão ao passivo; leitura intermediária; Simioni B R$ 260.000; lotes de possível APP ficam na empresa, lista = PDF Diretrizes).
- Cópia de segurança do decisoes-sei.json anterior no scratchpad da sessão (decisoes-sei.antes-pacote5.json).
- _scripts/aplica_pacote05.py (md5 `91f7638a3c`) grava a partir de pacote05.py (intermediaria): regras.R6; 37 lotes (11 próprios, 26 de terceiros, inclusive os 4 do S1); itens LOT-04 (260.000), LOT-20 (2.445.000), LOT-21 (1.790.000), LOT-22 (148.000), LOT-24 (5.982.500); passivos."Lotes de terceiros…" = 8.570.000; nota no C2. decisoes-sei.json md5 `ee803b0f11`.
- build_mestre_sei.py (md5 `b8c7a5658e`): a aba Passivos ganha "Valor adotado" e "Decisão" (seção "passivos" do JSON); o Resumo ganha o passivo adotado. Teste em cópia temporária: soma adotada 155.157.560,60; passivo 8.570.000; pendentes 7 (CONTR-01..07). **A planilha-mestre oficial não foi regravada: o arquivo está aberto no Excel (PermissionError).** Regravar quando o Luís fechar.
- Escritos: log-decisoes.md (pacote 5); registro-data-base.md linhas 7 a 11; perguntas-ao-grupo.md (pergunta 2 atualizada); parte6-candidatos-duplicidade.md (S1 aplicado); 00_FONTES.md (GEOFarroupilha e OCR).
- A marca "na lista de Rafael" dos Primavera 008 e 009 não precisa de correção (eles aparecem na nota "IMPORTANTE!", fora da tabela de escolhidos).
- Planilha-mestre oficial regravada depois que o Luís fechou o Excel: md5 `9484fb5a4e`. Resumo: bens adotados 155.157.560,60; passivo de terceiros adotado 8.570.000; 7 itens sem valor adotado (CONTR-01..07, Parte 5).

## 2026-10-28 (sessão SEI 2) — fechamento da Parte 3
- Início: revisor independente sobre TODAS as decisões aprovadas da Parte 3 (P11) + ficha de fechamento (pacotes/ficha-fechamento-parte3.md).
- Revisão de fechamento: _diario/revisao-fechamento-parte3.md (2 altas, 7 médias, 4 baixas). A planilha regenerada pelo revisor é idêntica à oficial.
- Achado 1 reproduzido por _scripts/saidas_pos_database.py → saidas-pos-database.xlsx (md5 `5a2aca244a`): 36 linhas do arq. 17 depois de 31.12.2024; 27 lotes ausentes do doc 11 (a 1ª versão do script deu 26 por casar Colina do Retiro I com II; corrigido com a regra de numeral romano).
- Aplicado (decisões já aprovadas): passivos."Observação 21 – Paraná Equipamentos" = 0 (pacote 2, item 7); lista_rafael_suspensos: Alto das Graças I H-6371/011 (S4). build_mestre_sei.py lê a suspensão (coluna "Na lista de Rafael" = "Suspenso da lista: ..."). decisoes-sei.json md5 `a07ea14fa8`; mestre md5 `3cca34ec4e`; soma adotada inalterada (155.157.560,60).
- registro-data-base.md: linha 4 corrigida; linhas 12 (S6), 13 (19 lotes, pendente), 14 (lote urbano 06 e Reserva Paradiso, pendente) e 15 (Ipê I 001/q2). Notas de decisão acrescentadas em pendencias-conferencia-anterior.md e sei-2025-notas.md (Oleoplan).
- Ficha de fechamento atualizada com 5 decisões para o Luís (seção 4). Pendente: aprovação do Luís.
- Pedido do Luís: verificar se as 5 decisões de fechamento já foram tomadas na sessão "Fila de tarefas SEI" (local_cd6bd22e). Lidos: analises/parte4-fluxo-caixa/log-decisoes.md e _diario/diario.md; analises/cobrancas-11.2/_diario/diario.md; busca nas transcrições ("Dal Ponte", "compensação ambiental").
  - Nenhuma das 5 foi decidida lá. Os "26 lotes" que o Luís aprovou ("os 26 lotes entram na conta") são outro conjunto: lotes vendidos em 2019–2024 além da redução do inventário. Os lotes saídos em 2025 (34/37 na conta deles) aparecem só como candidatos ao registro de data-base.
  - Compatível com a decisão 5: "regra geral do 11.2" (toda cobrança passada é zerada; no Lago Negro o grupo já lista só o que falta).
  - Decidido lá e ainda não aplicado na SEI: Solar Serra (obs. 14, R$ 150.000) zerada; o lote 02/07 do Aditivo Medianeira (R$ 165.000), já entregue, deve sair do estoque de 2025 (pendência para a sessão SEI); crédito da SEI contra o Hotel Araçá.
  - Cuidado na fusão: o 11.2.2 (S1, R$ 1.887.000) faz parte do abatimento de R$ 5.442.667 da Parte 4; o S1 já foi aplicado no pacote 5. Não aplicar duas vezes.
  - Linha 15 do log deles: lotes entregues que seguem no estoque (Medianeira 02/07, Alto das Graças 11/6371, Sol Nascente I 1/A e 2/A) "sem ajuste", decisão depois alterada para o Medianeira pela regra geral. O S4 (Alto das Graças fora da lista de Rafael) não conflita: o lote continua no estoque.
