Você coordena o estudo das demais empresas do grupo Siviero na cisão. São todas, exceto a SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. (SEI), que segue em outra sessão. Eu, Luís, sou sócio de 1/6 e estou saindo. Trabalhe em português, na pasta `C:\entrega-cisao-siviero`.

**Seu papel nesta sessão** é planejar e abrir as sessões de trabalho de cada empresa. Você mesmo não analisa as empresas.
1. Analisar cada empresa o bastante para escrever um prompt **específico** para ela. Não mande o mesmo prompt para todas.
2. Mostrar os prompts para mim.
3. Depois que eu aprovar, abrir uma sessão por prompt aprovado.

**Leia primeiro:**
- `00_PLANO.md`, inteiro, em especial as premissas P1 a P13 e a seção 7, com as lições de método e a regra de propriedade dos arquivos;
- `00_FONTES.md`;
- `02_trabalho/README.md`;
- `02_trabalho/log-decisoes.md`, para ver o método já aplicado na SEI: regras em bloco, pacotes de exceções, critérios C0, C1 e C2, revisor independente e as correções que ele exigiu;
- `02_trabalho/_diario/diario.md`.

A memória do projeto traz as regras de trabalho. Siga todas.

**Empresas.** Use a numeração dos inventários 2025, em `01_fontes/inventario-2025/`:

| Nº | Empresa |
|---|---|
| 01 | Abastecedora de Combustíveis |
| 02 | Agropecuária Fazenda São Crispin |
| 03 | Associação Cultural Princesa dos Vales |
| 04 | Ceval Terraplenagem |
| 05 | Hotel Princesa dos Vales |
| 06 | Hotel Araçá |
| 07 | Rádio Comunidade Princesa dos Vales FM |
| 08 | Transportadora Irmãos Siviero |
| 09 | Turismo Siviero |
| 10 | Siviero Transportes |

No Estudo 2 de Rafael, a "Empresa 10" é a SEI. Não confunda as numerações.

**Para cada empresa, levante** só o necessário para um bom prompt, sem fazer a extração completa:

- **Fontes, com caminhos:**
  - inventário 2025 atual e antigo, em `01_fontes/inventario-2025/` e `01_fontes/inventario-2025-antigo/`, e o resumo das diferenças em `02_trabalho/diff-inventario-2025-antigo-x-atual.md`;
  - inventário 2018 em `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\`. Só o da SEI foi copiado para `01_fontes`;
  - PDF de Rafael da empresa, em `01_fontes/estudo-rafael/`;
  - itens da empresa na lista de bens escolhidos por Rafael, em `01_fontes/outros/ativos-avaliados-valter.xlsx`. Essa lista define o critério C0;
  - aba da empresa nas planilhas Google exportadas, em `01_fontes/planilhas-google/`, e as observações do Luís em `02_trabalho/observacoes-planilhas-google.xlsx`;
  - conferência anterior em `01_fontes/outros/conferencia-discrepancias.md` e `01_fontes/outros/inventario-2025-avaliacoes-corrigido.xlsx`, que já cobriu 7 empresas;
  - `02_trabalho/pendencias-conferencia-anterior.md`;
  - documentos das respostas ao Dr. Pablo ligados à empresa, em `01_fontes/respostas-docs-2026-08-13/`.
- **Porte:** número de itens e total pelo grupo e por Rafael.
- **Situação:** ativa ou baixada.
- **Temas específicos já conhecidos.** Confirme cada um nas fontes:
  - **02 São Crispin:** área de 407,53 ha × 411,43 ha, matrícula 00.760 duplicada, vendas às famílias Rigo e Pellegrini, arrendamentos à família Lorenzet e soja a receber.
  - **05 Hotel Princesa:** a Sede Campestre e o lote 52 Bassani, que em 2018 estavam listados na SEI. Há risco de dupla contagem.
  - **04 Ceval:** escavadeira 320D de 2009, que também pode estar na SEI.
  - **08 Transportadora:** nota de R$ 1.135.000 contra itens que somam R$ 1.020.000.
  - **07 Rádio:** participação de 16,5% no inventário, mas vale 1/6 pela premissa P1; outorga do Ministério das Comunicações.
  - **01 Abastecedora:** contratos com a Petrobras e locação com a SIM.
  - **06 Hotel Araçá:** outorga do poço e matrícula 73.482 repetida.
  - **03, 09 e 10:** empresas baixadas.
  - **Todas:** venda das quotas de Rovílio Siviero.

**Análise em paralelo.** Uma sessão à parte analisa o arquivo ACERTOS E ACORDOS e o documento de áreas vendidas do Hotel, com a Fazenda São Crispin, os créditos em soja e o arrendamento. Os achados ficam em `02_trabalho/analises/acertos-e-areas-vendidas/achados-para-empresas.md`. Os prompts da São Crispin (02) e do Hotel Princesa (05) devem mandar usar esses achados depois que eu aprovar, sem escrever naquela pasta.

**Proponha um agrupamento e justifique.** Por exemplo, as empresas baixadas sem bens poderiam ficar numa única sessão, e empresas que compartilham bens poderiam ficar juntas.

**Cada prompt de empresa deve conter:**
1. Objetivo.
2. As premissas que mais pesam para aquela empresa, remetendo ao `00_PLANO.md` e à memória para o resto.
3. As fontes específicas, com caminhos.
4. Os temas e riscos conhecidos da empresa.
5. A sequência de trabalho: Fase 0 da empresa, planilha-mestre da empresa, pacotes de regras e de exceções com revisor antes de cada pacote, e ficha de fechamento para minha aprovação.
6. A regra de arquivos. A sessão escreve só em `02_trabalho/empresas/NN-nome/`, com diário, log de decisões, arquivo de decisões, perguntas ao grupo, candidatos à Parte 6 e registro de data-base próprios. Não altera memória existente. Bens que cruzam com a SEI ou com outra empresa vão para os candidatos à Parte 6 dela.
7. Idioma: português.

**Mostre para mim:**
1. Uma tabela-resumo das empresas, com o agrupamento proposto.
2. Cada prompt.

Espere a minha aprovação. Posso aprovar todos de uma vez ou pedir ajustes.

**Depois de aprovado,** abra uma sessão por prompt aprovado, como sessão nova e sem worktree, porque a pasta não é repositório git. Registre as sessões abertas em `02_trabalho/empresas/_orquestracao/sessoes.md`. Salve a análise e os prompts em `02_trabalho/empresas/_orquestracao/`.

**Não escreva fora de `02_trabalho/empresas/`.**
