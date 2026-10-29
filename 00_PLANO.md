# Plano de trabalho — Estudo da cisão Siviero (quota de 1/6)

Versão: 28/10/2026 (plano corrigido após revisão do Luís).
Pasta de trabalho: `C:\entrega-cisao-siviero`. Fontes originais ficam onde estão (Drive e `C:\inventario-siviero`); aqui entram cópias, planilhas de análise e o material final.

---

## 1. Premissas fixas

| # | Premissa | Consequência prática |
|---|---|---|
| P1 | Participação do sócio que sai é **1/6 em todas as instâncias**. | Ignorar os percentuais dos quadros societários (18,60 %, 18,18 %, 16,50 % etc.). Toda apuração usa 1/6. |
| P2 | **Data-base é 31.12.2024.** Rafael não a respeitou; o grupo Siviero às vezes também não. | Nada é retirado ou alterado por esse motivo sem um registro explícito (ver 0.8). Nunca ajustar em silêncio. |
| P3 | **Inventário de 2018 é o benchmark de existência do patrimônio**, não de valor. É o patrimônio tal como foi entregue ao sair da administração. | A comparação 2018 → 2025 é feita em quantidade/identidade de bens (lote, quadra, matrícula, área). Os valores de 2018 só servem de referência secundária. |
| P4 | **Todo bem que saiu entre 2018 e 2025 precisa ter um destino** (venda, permuta, dação, quitação de dívida, consignação a sócio, perda judicial…). | Um bem sem destino documentado é uma pendência a ser apontada, não uma duplicidade. |
| P5 | **Não há reavaliação própria.** Para cada item adota-se uma das duas avaliações (grupo Siviero ou Rafael), salvo casos extremos. | A planilha-mestre registra a decisão por item e o motivo. "Caso extremo" exige justificativa escrita. |
| P6 | **Prioridade absoluta: SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. (SEI).** Sem pressão de prazo. | Tudo que envolve a SEI é feito antes das demais empresas. |
| P7 | **Idioma: português** nos documentos e planilhas da entrega. Inglês só em código e README técnico. | Sobrepõe a regra global de arquivos em inglês para este projeto. |
| P8 | Material explicativo (arquivo único ou apresentação) só depois de **todo o material técnico estar finalizado e aprovado**. | Nenhum esforço de apresentação antes disso. |
| P10 | **Objetivo da cisão (28/10/2026):** reduzir o valor de mercado dos bens que virão para os sócios que saem e aumentar o valor do patrimônio que fica na empresa. Até nova aprovação, a lista dos bens que virão para nós é a de Rafael (bens selecionados pela Valter no Estudo 2). | **Critério C0:** quando grupo e Rafael divergem, bem da lista de Rafael recebe o **menor** dos dois valores; bem que fica na empresa recebe o **maior**. Vale salvo decisão diferente do Luís. Em cada item, registrar também o argumento técnico (ex.: faixa de mercado de Rafael) que sustenta o valor escolhido. |
| P11 | **Revisão independente obrigatória (28/10/2026).** Todo pacote, fase ou entregável passa por um revisor antes de chegar ao Luís, e as decisões aprovadas são revisadas outra vez antes de fechar cada parte. | O revisor é um agente com contexto novo, que recebe as fontes e os arquivos produzidos, mas não as conclusões, e tenta encontrar erros. O resultado e o que foi corrigido vão para o diário de trabalho. |
| P12 | **Diário de trabalho (28/10/2026).** Todo passo fica registrado em `02_trabalho/_diario/diario.md`, só com acréscimos: entradas com hash, saídas, contagens, conferências e pendências. | O diário serve para Claude retomar o trabalho e para o revisor. Não é material para leitura humana. Números nunca são tirados da memória: sempre vêm das fontes por script. |
| P9 | Nenhum material fica de fora por não conseguir abri-lo. Primeiro tentar conectores (Drive, Sheets, Docs); se ainda falhar, avisar para o Luís resolver. | Vale para as planilhas Google, `.gdoc`, PDFs escaneados etc. |

---

## 2. Fontes e seus papéis

| Fonte | Onde está | Papel no estudo |
|---|---|---|
| **Inventário 2018** (31.12.2018, 27 docx) | Drive `Material Siviero/2018/inventário-2018` | Benchmark de existência (P3). Item 18 = SEI, inclusive uma versão xlsx. |
| **Inventário 2025** (31.12.2024, docs 00 a 11 + ACERTOS E ACORDOS + docx Hotel Princesa) | Drive `Material Siviero/2025/inventário-2025` e cópia idêntica em `C:\inventario-siviero/2025/inventário-2025` | Avaliação do grupo Siviero (5/6). Doc 11 = SEI, atualizado em 26/10/2026. |
| **Versão antiga do inventário 2025** (11 docx "ANTIGO") | Drive `Material Siviero/Material antigo/2025 antigo` | Conferir o que mudou entre versões (0.4). |
| **Estudo 2 de Rafael** (13 PDFs gerais + 13 PDFs de empreendimentos, assinados BELISE LTDA) | Drive `Material Rafael/Estudo 2` | Avaliação de Rafael. Os PDFs de empreendimentos já estão em Excel (`versão-excel` no Drive e `belise-empreendimentos` no repositório). |
| **Respostas do grupo ao Dr. Pablo** (arquivos 1 a 18, 13.08.2026) | Drive `Material Siviero/2025/[siviero] respostas docs` | Arquivo 17 = lotes vendidos jan/2018–out/2025 (391 linhas). Arquivo 18 = lotes consignados ao Válter. Arquivo 16 = acordos Susin/AGI. Arquivo 15 = locação Oleoplan. |
| **Planilhas Google do Luís** | Drive `Material Siviero/2_0_Comparativo Patrimonial Siviero 2018 x 2025` e `Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025` (+ 2 "Cópia de…") | Resumo já feito das duas visões; contém as observações da Parte 7. Acesso via conector (0.5). |
| **Conferência de 26/10/2026** | `C:\inventario-siviero/conferencia-discrepancias.md` + `inventario-2025-avaliacoes*.xlsx` | Ponto de partida das planilhas das demais empresas. Parcialmente desatualizada (0.6). |
| **Estudo "lotes que saíram 2018-2025"** | Drive `Material Siviero/Estudo - lotes que saíram 2018-2025.md` + `Lotes vendidos 2018-2025 x GERAL.xlsx` | Base da Parte 4 e da Parte 6 (Lago Negro, Alto das Graças I, Altos da Julieta em aberto). |
| **Relatórios mensais** (jul/2023–out/2024, PDFs + 3 xlsx de contas pagas) | Drive `Material Siviero/Relatórios Mensais` | Apoio: comissões, obras, pagamentos que podem duplicar descontos (Parte 6). |
| **Documentos diversos** (Buriti, Susin, soja, retiradas 2018, DWG Sol Nascente, holdings) | Drive `Documentos`, `Holdings` | Apoio às Partes 4 e 5. |

---

## 3. Fase 0 — Revisão e consolidação

**0.1 Estrutura da pasta de entrega**

```
C:\entrega-cisao-siviero\
  00_PLANO.md                  este arquivo
  00_FONTES.md                 índice: cada fonte usada, caminho de origem, data, hash, papel
  01_fontes\                   cópias dos originais efetivamente usados (somente leitura)
  02_trabalho\                 planilhas e notas de análise (planilha-mestre, registros, estudos)
  03_entrega\                  material final, só após aprovação do técnico
```

**0.2 Duplicidades de arquivos entre as duas pastas** — decidir a versão canônica e registrar em `00_FONTES.md`.
- `00 - RESUMO BELISE LTDA.xlsx` difere entre Drive e repositório (e o repositório tem essa mudança sem commit). Decidir qual vale.
- `LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.xlsx` só existe no Drive (`versão-excel`).
- `inventario-irmaos/` no repositório tem versão anterior da planilha 2025 e cópia idêntica de `ativos-avaliados-valter.xlsx`. Confirmar que é obsoleta.
- Demais xlsx de `versão-excel` são idênticos por hash aos de `belise-empreendimentos`; `_fonte` está duplicado nos dois lugares. Regra proposta: o repositório `C:\inventario-siviero` é a origem dos arquivos gerados; o Drive recebe cópias.
- `inventário-estudo-rafael.pdf` no repositório é o PDF "DIRETRIZES PARA SELEÇÃO DE ATIVOS…" do Drive (mesmo hash). Registrar o nome real.

**0.3 Três fotografias do patrimônio da SEI** — montar a lista de bens de 2018 (doc 18 + xlsx), de 2025 (doc 11) e do Estudo 2 (PDFs de empreendimentos + "LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS"), item a item, com identificação única (loteamento + quadra + lote, ou matrícula, ou descrição + área). Essa é a chave de ligação da planilha-mestre.

**0.4 Diff entre "2025 antigo" e inventário 2025 atual** — para cada um dos 11 documentos, listar o que mudou (itens incluídos, retirados, valores alterados, texto de compromissos). O doc 11 (SEI) tem prioridade. Resultado: `02_trabalho/diff-inventario-2025-antigo-x-atual.md`.

**0.5 Acesso às planilhas Google** — localizar as duas planilhas (e as cópias) pelo conector do Drive/Sheets, exportar para xlsx em `01_fontes` e extrair todas as notas/comentários (alimenta a Parte 7). Se o conector não alcançar, avisar o Luís.

**0.6 Atualizar a conferência de 26/10** — ela registra como ausentes os inventários 01, 02, 05 e 11, que hoje existem. A planilha `inventario-2025-avaliacoes-corrigido.xlsx` não cobre Abastecedora, São Crispin, Hotel Princesa dos Vales e SEI. Esses quatro entram direto na planilha-mestre (0.7), não na planilha antiga.

**0.7 Planilha-mestre** (`02_trabalho/planilha-mestre.xlsx`) — uma linha por bem. Começa pela SEI; as demais empresas entram depois.

| Coluna | Conteúdo |
|---|---|
| Empresa, Bloco | Empresa e bloco do inventário (ex.: SEI › Lotes em estoque › Loteamento X) |
| Chave | Identificador único do bem (loteamento/quadra/lote, matrícula, ou descrição+área) |
| Existe em 2018? | Sim/Não + referência (item, página do doc 2018) |
| Existe em 2025 (grupo)? | Sim/Não + referência (item, página do doc 2025) |
| Existe no Estudo 2 (Rafael)? | Sim/Não + referência (PDF, página) |
| Valor grupo 2025 | R$ |
| Valor Rafael | R$ |
| Diferença | R$ e % |
| Situação | Concorda / Discorda / Só grupo / Só Rafael / Não abordado |
| Decisão (P5) | Adoto grupo / Adoto Rafael / Caso extremo |
| Motivo da decisão | Texto curto |
| Destino 2018→2025 | Preenchido na Parte 4: venda, permuta, dação, consignação, perda, sem destino |
| Registro de data-base | Link para a linha do registro 0.8, quando houver |
| Observação do Luís | Trazida da planilha Google (Parte 7) |
| Situação da observação | Conferido / Alterado / Pendente |

**0.8 Registro de ajustes de data-base** (`02_trabalho/registro-data-base.md`) — tabela com: item, fonte, data do fato, por que foge de 31.12.2024, quem fugiu (Rafael / grupo), tratamento adotado e justificativa. Alimentado ao longo de todas as partes (P2).

**0.9 Estudo à parte: pendências da conferência anterior** (`02_trabalho/pendencias-conferencia-anterior.md`) — área da Fazenda São Crispin (407,53 ha × 411,43 ha), matrícula 00.760 duplicada, R$ 3.000.000 de diferença interna no imóvel Oleoplan, versão oficial da planilha 2025, nota de R$ 1.135.000 da Transportadora. Não bloqueia o estudo principal; cada pendência recebe o que falta para decidir.

---

## 4. Partes do estudo e ordem de execução

Ordem: **0 → mestre (SEI) → 3 → 4 → 6 → 5 → 7 → 1 e 2 (demais empresas)**. Motivos: a SEI vem primeiro (P6); a Parte 6 depende do destino de cada lote, que é apurado na Parte 4; a Parte 7 se resolve em boa parte enquanto se classifica; as demais empresas ficam por último porque a conferência de 26/10 já fez metade do trabalho.

### Parte 3 — SEI: concorda × discorda, exceto o tópico "CONTRATOS | INDENIZAÇÕES | COMPROMISSOS | EMPRÉSTIMOS"
Por bloco do doc 11: lotes em estoque por loteamento; lotes urbanos; áreas rurais; imóvel BR-470 (Oleoplan); veículos e outros. Para cada bloco: lista de concordâncias, lista de discordâncias com valor dos dois lados e página, itens só de um lado. Cruzar com os 13 Excel do Estudo 2 e com o estoque de out/2025 da planilha Google.

### Parte 4 — Tópico "CONTRATOS | INDENIZAÇÕES | COMPROMISSOS | EMPRÉSTIMOS" e destino de cada bem que saiu
Abordagem própria, a ser explicada pelo Luís no início desta parte. Preparação independente da explicação:
- Extrair cada item do tópico no doc 11 para `02_trabalho/sei-compromissos.xlsx`: contraparte, objeto, lotes entregues (chave), valor em espécie, data, fonte/página.
- Cruzar com ACERTOS E ACORDOS.pdf, arquivos 16/17/18, planilhas de Rafael "DEMAIS COMPROMISSOS A PAGAR" e "LOTES DE TERCEIROS INDENIZÁVEIS", termos Susin/AGI.
- Preencher a coluna "Destino 2018→2025" da planilha-mestre para todo bem que existia em 2018 e não existe em 2025 (P4). Continuar o estudo "lotes que saíram" (Lago Negro, Alto das Graças I, Altos da Julieta).
- Saída: lista de bens **sem destino documentado**.

### Parte 6 — Descontos em duplicidade
Verificar, para a SEI (depois para as demais), todas as formas em que um mesmo valor pode ter sido reduzido duas vezes. Formas a investigar (lista inicial; a análise pode encontrar outras):
1. **Saída dupla do mesmo lote** — lote retirado da lista patrimonial (desconto 1) e o mesmo lote descontado de novo ao aparecer como quitação de dívida/dação/indenização (desconto 2). Exemplo dado pelo Luís.
2. **Mesmo lote em dois acordos** — o mesmo lote aparece como entregue em mais de um documento (ACERTOS E ACORDOS × doc 11 × arquivo 17 × termos Susin/AGI).
3. **Venda e dação do mesmo lote** — lote listado como vendido (arquivo 17) e também como dado em pagamento.
4. **Lote consignado a sócio** (arquivo 18) — saiu do estoque da empresa e, ao mesmo tempo, é tratado como crédito/débito do sócio ou como passivo.
5. **Passivo de obra + lote avaliado líquido** — pavimentação/infraestrutura descontada como compromisso e o lote avaliado já como "sem obras" ou com valor reduzido.
6. **Área vendida descontada duas vezes** — área total reduzida pela venda e o valor da mesma venda descontado de novo (São Crispin: 686,75 ha × 491 ha; vendas Rigo/Pellegrini).
7. **Percentual de parceria aplicado duas vezes** — na quantidade e no valor (Alto das Graças I: 50 % Villa Luna; Altos da Julieta: parte Perini).
8. **Contingência judicial + retirada do bem** — prejuízo de processo lançado como passivo e os lotes do mesmo processo retirados do estoque ou zerados (Lago Negro: 7 lotes de APP sem valor + prejuízo de R$ 1,75 milhão).
9. **Empréstimo intra-grupo** — passivo em uma empresa sem o ativo correspondente na outra (varrer os 11 inventários).
10. **Custos de venda** — comissões/custos descontados do valor dos lotes vendidos e também como despesa nos relatórios mensais.
11. **Reduções de valor sobre bens já retirados** — ajustes de valor (desvalorização, "em estudo de viabilidade", "saldo zero") aplicados a itens que já não estão na lista.
Saída: `02_trabalho/duplicidades.xlsx` com uma linha por ocorrência: forma, bem, documentos envolvidos (página), valor duplicado, efeito sobre 1/6.

### Parte 5 — Assuntos não abordados
Para cada candidato, responder: o doc 11 menciona? com valor? Rafael avaliou? Candidatos iniciais:
- Os 12 loteamentos contratados pela SEI "sem obras" citados só no docx do Hotel Princesa dos Vales (Alto das Graças II, Santa Corona III, Morada dos Ipês II, Congregação São Carlos, Bella Visione, Bonetto, Pascoal, Letti, Colombo, Champagnat, Simioni, Sol Nascente).
- Os 7 loteamentos de "áreas contratadas" avaliados por Rafael (Caminhos do Sol, Cecconello, Pedrotti/Andreis/Lorandi, Dr. Letti/Sol Nascente III, Pasqual, Vale das Colinas, Vista da Colina) e os loteamentos Bonetto e Colina do Retiro II.
- Memorando e contrato Buriti Holding × Siviero (loteamentos de Caxias do Sul).
- Projeto urbanístico Sol Nascente (DWG).
- Contrato de locação dos pavilhões à Oleoplan (arquivo 15): receita e prazo.
- Qualquer bem do inventário 2018 que não apareça em 2025 nem tenha destino (resultado da Parte 4).
Saída: `02_trabalho/nao-abordados.xlsx`.

### Parte 7 — Observações do Luís na planilha 2_0
Extrair toda nota/comentário/célula destacada para uma lista com: aba, célula, texto, item da planilha-mestre, situação (Conferido / Alterado / Pendente) e o que foi feito. Boa parte se resolve nas Partes 3 a 6; o que sobrar vira pendência explícita.

### Partes 1 e 2 — Demais empresas: concorda × discorda
Mesmo método da Parte 3, para Abastecedora, São Crispin, Associação Cultural, Ceval, Hotel Princesa dos Vales, Hotel Araçá, Rádio, Transportadora, Turismo e Siviero Transportes. Reaproveitar `inventario-2025-avaliacoes-corrigido.xlsx` (7 empresas já conferidas) e incluir as 4 que faltam. Concorda = valor igual (tolerância a definir). Discorda = valor diferente, ou item presente em só uma das avaliações.

---

## 5. Produtos finais (após aprovação técnica)

- `03_entrega/` recebe: planilha-mestre consolidada, duplicidades, não abordados, registro de data-base, pendências e o estudo da SEI.
- Material explicativo simplificado (arquivo único ou apresentação) só depois de tudo acima aprovado (P8).

## 6. Pendências de acesso

- Planilhas Google: tentar pelo conector do Drive/Sheets (0.5). Se falhar, pedir os links ou a exportação para xlsx.
- `ACERTOS E ACORDOS.gdoc` (Drive): idem.
- PDFs escaneados (ACERTOS E ACORDOS, respostas docs): leitura página a página como imagem quando a extração de texto falhar.

---

## 7. Estado em 28/10/2026 (fim da primeira sessão) e organização em sessões

### Estado da SEI
- **Fase 0 concluída** (fontes, extrações 2018 e 2025, diff de versões, planilhas Google, planilha-mestre).
- **Parte 3: pacotes 1 a 4 aprovados** e revisados por revisor independente. Detalhes em `02_trabalho/log-decisoes.md`; decisões aplicáveis em `02_trabalho/decisoes-sei.json`; planilha `02_trabalho/planilha-mestre-sei.xlsx`.
- Bens 2025 da SEI: R$ 152.796.260,60 pelo grupo; R$ 153.970.060,60 pelos valores adotados. Sem valor adotado: 7 loteamentos contratados (Parte 5) e o Fracionamento Simioni (reaberto).
- **Pendente na Parte 3:** regra única para todos os casos de "possível APP", com o Simioni incluído, e a verificação de descontos duplos ambientais. O levantamento foi cancelado: `02_trabalho/levantamento-app.xlsx` é parcial e **não revisado**, sirva só de pista. Depois: ficha de fechamento da Parte 3 e nova revisão de todas as decisões aprovadas (P11).
- **Partes 4 a 7 não iniciadas.** A Parte 4 depende da explicação do método pelo Luís.

### Lições de método (valem para todas as empresas)
1. O valor de Rafael pode estar em qualquer aba ou documento: Avaliacao, Quadro_Resumo, "balanço executado por Valter", tabelas "Lotes escolhidos", PDF Diretrizes. Procurar em todos antes de afirmar que "Rafael não avalia" um bem. Esse erro aconteceu duas vezes na regra 4 da SEI.
2. Conferir no nível do lote **e** no nível do loteamento ou item.
3. O casamento automático 2018 × 2025 por semelhança de texto erra. Conferir pela matrícula, pelo cadastro municipal ou pelo vendedor.
4. O inventário de 2018 da SEI lista "bens relacionados nos inventários de outras empresas", fora do total. Em 2025 eles aparecem na SEI ou em outra empresa: risco de dupla contagem (Parte 6).
5. Numeração: no Estudo 2 de Rafael, "Empresa 10" é a SEI; nos inventários 2025, o doc 10 é a Siviero Transportes. Usar sempre a numeração do inventário 2025.

### P13 — sessões em paralelo e propriedade dos arquivos

| Sessão | Escreve em | Só lê |
|---|---|---|
| SEI (continuação) | `02_trabalho/` (raiz, `_diario/`, `pacotes/`, `_scripts/`), `00_PLANO.md`, `00_FONTES.md` | `02_trabalho/empresas/`, `03_entrega/` |
| Coordenação das demais empresas | `02_trabalho/empresas/_orquestracao/` | todo o resto |
| Uma sessão por empresa (ou grupo de empresas) | `02_trabalho/empresas/NN-nome/`, com diário, log, decisões, perguntas ao grupo, candidatos à Parte 6 e registro de data-base **próprios** | todo o resto |
| Apresentação (PowerPoint) | `03_entrega/apresentacao/` | todo o resto |
| Análise de ACERTOS E ACORDOS e das áreas vendidas | `02_trabalho/analises/acertos-e-areas-vendidas/` | todo o resto |

- Os arquivos da raiz de `02_trabalho` (`perguntas-ao-grupo.md`, `parte6-candidatos-duplicidade.md`, `registro-data-base.md`, `log-decisoes.md`) são da sessão SEI. Os das empresas são unidos a eles só depois que o Luís aprovar cada empresa.
- Memória do projeto: só a sessão SEI edita arquivos existentes. As outras sessões podem criar arquivos novos, mas não alterar os existentes.
- Prompts de abertura das sessões: `00_prompts/`.
