# Revisão independente — SEI, pacotes 1 a 3

Revisor com contexto novo, 28/10/2026. Arquivos lidos sem alteração. Scripts de conferência em `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\eb6693c8-36bc-49ba-865b-2dcc272284c1\scratchpad\revisor\`.

## Método

1. Li P1–P12, `decisoes-sei.json`, `log-decisoes.md` e `build_mestre_sei.py`. O script não foi executado.
2. Recalculei, a partir das abas Mestre e Lotes, os totais, as somas por regra e as somas por loteamento.
3. Conferi os valores com o texto do doc 11 (docx), com os 13 xlsx de Rafael (Balanco_Grupo, Lotes, Avaliacao, Quadro_Resumo) e com o texto dos 26 PDFs de Rafael, extraído com pdfplumber.
4. Comparei a lista Diretrizes (`ativos-avaliados-valter.xlsx`, 98 linhas da SEI) com as marcas da planilha-mestre e com as tabelas "Lotes escolhidos pela Valter".
5. Procurei nos PDFs gerais de Rafael os veículos, as máquinas e as áreas urbanas da SEI pela placa, pelo modelo, pela matrícula e pelo endereço.

## Achados, por impacto financeiro

### 1. ALTA — A R4 está errada em 6 dos 8 loteamentos contratados: Rafael dá valor de ativo, no nível do quadro-resumo

- **O que está errado.** A R4 diz "loteamento contratado sem valor de ativo nas duas avaliações". Para 6 dos 8 loteamentos isso é falso. O grupo não dá valor de ativo. Rafael dá: no quadro-resumo de cada PDF de "Áreas contratadas", registra um "Patrimônio total dos bens imóveis desta empresa". É o mesmo tipo de erro de premissa já corrigido nos lotes de APP. O script só lê `Valor avaliado` da aba Avaliacao, e essas pastas não têm essa aba. Lê também o Balanco_Grupo, onde o valor é placeholder. Por isso classificou as 6 como "Sem valor nas duas fontes". No Caminhos do Sol o campo Item está vazio na pasta de Rafael, e o item nem casou.
- **Evidência** (aba Mestre, coluna AB "Situação" = "Sem valor nas duas fontes" e coluna AD "Valor adotado" = 0):

| Linha Mestre | Chave | Valor de ativo de Rafael | Base de Rafael | Fonte |
|---|---|---|---|---|
| 86 | CONTR-01 Caminhos do Sol | R$ 8.601.000,00 | 61 lotes (10% de 612) × R$ 141.000 | PDF Caminhos do Sol, p. 2-3 e 6 |
| 87 | CONTR-02 Vale das Colinas | R$ 5.019.600,00 | 10% de 356 lotes × R$ 141.000 (R$ 50.196.000) | PDF Vale das Colinas, p. 6-7 |
| 88 | CONTR-03 Vista da Colina | R$ 3.384.000,00 | 24 lotes × R$ 141.000 | PDF Vista da Colina, p. 5 e 8 |
| 89 | CONTR-04 Pedrotti/Andreis/Lorandi | R$ 3.384.000,00 | 24 lotes × R$ 141.000 | PDF Pedrotti, p. 4 e 6 |
| 90 | CONTR-05 Cecconello | R$ 9.165.000,00 | 65 lotes × R$ 141.000 | PDF Cecconello, p. 4 e 6 |
| 92 | CONTR-07 Pasqual | R$ 1.479.000,00 | 14,5 lotes × R$ 102.000 | PDF Pasqual, p. 5 |
| | **Soma** | **R$ 31.032.600,00** | | |

  A R4 vale de fato só para o CONTR-08 (Dr. Letti) e o CONTR-09 (Sol Nascente III): Rafael registra patrimônio R$ 0 (PDF Dr. Letti/Sol Nascente III, p. 6).
- **Valor em jogo.** Nenhum dos 6 está na lista de Rafael. Pelo C0, cada um ficaria com o maior valor, o de Rafael. O total adotado subiria até R$ 31.032.600,00, de R$ 153.970.060,60 para R$ 185.002.660,60. Isso equivale a R$ 5.172.100,00 em 1/6.
- **Ressalvas para a decisão.**
  - Rafael chama esses números de "valor patrimonial de referência" sem avaliação específica (Caminhos do Sol, p. 3-4).
  - Vista da Colina e Pedrotti usam os mesmos 242 lotes e o mesmo custo; a área vendável de Vista da Colina é igual à de Vale das Colinas. Parece cópia.
  - A área rural 10 (R$ 600.000, R2) fica "dentro do loteamento de 130.446 m²" da família Cecconello (PDF Áreas rurais, p. 20 e 24). Há risco de contagem dupla com o CONTR-05.
  - As multas já são deduzidas pelo grupo nos R$ 15.754.581,00 e não mudam.
- **Correção sugerida.** Retirar os 6 da R4 e apresentá-los ao Luís como exceções, num pacote só, com o C0. No script, ler o "Patrimônio total" da aba Quadro_Resumo nas pastas de áreas contratadas e casar o Caminhos do Sol pelo nome. Corrigir também o texto da R4 no JSON e no log.

### 2. MÉDIA — C2 contra C0 nos lotes de APP: o conflito não foi apresentado ao Luís

- **O que está errado.** Os 11 lotes de APP confirmada ficam na empresa. Pelo C0 receberiam o maior valor, o de Rafael. A decisão C2 os zerou. O log (pacote 2, item 3) registra a C2, mas não diz que o C0 levaria a outro resultado. A confirmação do Luís sobre o C0 diz "bem que fica sobe para o maior valor quando isso beneficia o nosso lado", e isso vai no sentido contrário ao da C2.
- **Evidência.**
  - Lago Negro, quadra 11B, 7 lotes: Rafael apura "diferença nominal em seu favor no montante de R$ 7.700.000,00 (Verificar documentação relativa às APPs)", ou seja, 7 × R$ 1.100.000 (PDF Lotes loteamentos, p. 84).
  - Fracionamento Simioni, 4 lotes: Rafael R$ 520.000 (xlsx Lotes loteamentos, Avaliacao, item 04, p. 17; Quadro_Resumo, p. 20).
  - Aba Lotes, linhas 15-18 e 255-261, coluna W = 0.
- **Valor em jogo.** Até R$ 8.220.000,00, R$ 1.370.000,00 em 1/6. O valor de R$ 7,7 mi é uma cobrança de Rafael, não uma avaliação de mercado; para o Simioni, o próprio JSON diz "sem referência de mercado".
- **Correção sugerida.** Levar ao Luís a pergunta explícita: a C2 prevalece sobre o C0 para bens que ficam na empresa? Registrar a resposta no JSON e no log.

### 3. MÉDIA — 5 lotes em "possível APP" estão entre os escolhidos de Rafael, mas a mestre os marca "Não"

- **O que está errado.** As tabelas "Lotes escolhidos pela Valter" têm 96 lotes da SEI. A lista Diretrizes tem 91. A diferença são 5 lotes em "possível APP" que Rafael põe entre os seus escolhidos. A mestre segue a Diretrizes e marca os 5 como "Não" (aba Lotes, coluna V).
- **Evidência.**

| Lote | Linha (aba Lotes) | Valor do grupo | Fonte de Rafael |
|---|---|---|---|
| San Marino 012/1650 | 269 | R$ 244.000 | PDF Lotes loteamentos, p. 94 |
| San Marino 013/1650 | 270 | R$ 233.000 | p. 94 |
| San Marino 004/1676 | 275 | R$ 253.000 | p. 94 |
| Firminio Dalzóchio 002/1292 | 284 | R$ 131.000 | p. 100 |
| Industrial Vicentino 013/1303 | 293 | R$ 1.125.000 | p. 116 |

  Os 5 aparecem na aba Quadro_Resumo do xlsx, linhas 156-158, 167 e 183.
- **Valor em jogo.** R$ 1.986.000,00 de valor cheio. Esses lotes estão pendentes no JSON (% do lote cheio, Parte 4). Quando o % for decidido, a marca define o sentido do C0: menor valor se o lote vier para nós, maior se ficar.
- **Correção sugerida.** O Luís decide qual lista vale: os 91 da Diretrizes ou os 96 dos quadros-resumo. Registrar no JSON antes de decidir o %. As outras 91 marcas conferem uma a uma com a Diretrizes e com o PDF (p. 9-12).

### 4. MÉDIA (risco de decisão aprovada) — Oleoplan: o valor de Rafael parece erro de soma, não avaliação

- **O que está errado.** O AURB-04 foi adotado em R$ 10.631.378,60 pelo C0. No PDF Diretrizes (p. 8-9), Rafael transcreve as parcelas do grupo sem avaliar nada: R$ 9.584.060,00 + R$ 3.113.322,00 + R$ 933.996,60 = R$ 13.631.378,60. O total que ele imprime é R$ 3 mi menor. A P5 manda escolher entre duas avaliações, e este número é uma transcrição com erro aritmético. O log já anota o risco de alegação de erro de soma; a decisão é do Luís.
- **Valor em jogo.** R$ 3.000.000,00, R$ 500.000,00 em 1/6.
- **Correção sugerida.** Manter a decisão, mas guardar o argumento técnico do C1 para o caso de contestação (P10 pede esse registro). Hoje não há nenhum: Rafael não apresenta faixa de mercado para esse imóvel.

### 5. BAIXA — Lago Negro, quadra 01: a aba Lotes mostra que Rafael concorda, mas ele não concorda

- **O que está errado.** Nas linhas 250-254 da aba Lotes, a coluna N "Valor Rafael" traz o valor que Rafael transcreve do grupo (R$ 770.000 a 879.000), e a coluna Q diz "Mesmo valor". O valor que Rafael avalia é R$ 1.100.000 por lote (PDF Lotes loteamentos, p. 81). O script pega o primeiro valor casado (`vals[0]`). Os valores adotados estão certos, porque vêm do JSON. O problema é que qualquer lote com dois valores em Rafael é classificado como "Mesmo valor" e pode receber a R1 se sair da lista de exclusões.
- **Valor em jogo.** Nenhum hoje. Seria R$ 1.149.000 se a exclusão do Lago Negro fosse removida.
- **Correção sugerida.** Usar o último valor, o da avaliação, ou marcar "Valor diferente" quando há mais de um valor.

### 6. BAIXA — Colunas de Rafael na aba Mestre não refletem o PDF Diretrizes

- **O que está errado.** O AURB-04 (linha 42) e o VEIC-14 (linha 15) aparecem como "Não abordado por Rafael", com a coluna W vazia. Rafael lhes dá R$ 10.631.378,60 e R$ 67.000,00 na Diretrizes (p. 8-9). Os lotes em possível APP de San Marino e Industrial Vicentino aparecem sem valor de Rafael, embora ele escreva o valor ao lado (p. 94 e 116).
- **Valor em jogo.** Nenhum, porque os valores adotados vêm do JSON. O risco é de leitura.
- **Correção sugerida.** Preencher essas colunas a partir da Diretrizes.

### 7. BAIXA — Máquinas (R3): a premissa se mantém, mas há risco de contagem dupla e uma referência de Rafael

- **O que está errado.**
  - No PDF da Ceval (p. 1-5), Rafael avalia uma escavadeira Caterpillar 320D de 2009 (modelo 2010, NF 4946) em R$ 240.000, com faixa de R$ 240.000 a 295.000. Avalia também uma retro Case 580L 2001 (placa IKI 4325) em R$ 100.000, com faixa de R$ 85.000 a 130.000.
  - Na SEI, o MAQ-03 é uma 320D de 2009 a R$ 200.000, e o MAQ-07 é uma Case 580L 2001 (placa IKI 4827) a R$ 60.000.
  - A retro é outro bem, pela placa. A escavadeira pode ser a mesma, contada nas duas empresas: o inventário SEI de 2018 já tinha uma 320D de 2009 (R$ 250.000), e o doc 04 da Ceval também lista uma.
  - Fora isso, nenhum veículo, máquina ou área urbana da SEI aparece em PDF de Rafael. A única exceção é a Ranger (R2). A Transportadora tem um terreno vizinho (mat. 3.753) a R$ 992,78/m², a mesma taxa dos AURB-02/03, e Rafael concorda com ela (PDF Transportadora, p. 4). Isso apoia o valor do grupo.
- **Valor em jogo.** R$ 200.000 se a escavadeira estiver duplicada.
- **Correção sugerida.** Conferir a NF 4946 e o chassi na Parte 6.

### 8. BAIXA — Aba Passivos: casamento falso pelo valor

- **O que está errado.** A linha 17 (Observação 1, Rovílio, R$ 1.650.000) aponta para o documento de Rafael "Dr. Letti e Sol Nascente III". Os dois valores só coincidem com a multa do Dr. Letti. O script casa os passivos apenas pelo valor.
- **Valor em jogo.** Nenhum.
- **Correção sugerida.** Casar também pelo item ou pelo documento.

### 9. BAIXA — Inconsistências internas de Rafael (sem efeito nos valores adotados)

- Vila Vicenza: "Patrimônio de bens considerado" de R$ 148.000 (p. 110), contra R$ 157.000 na p. 108 e na lista de lotes. A R1 usa R$ 157.000, que está correto.
- `00 - RESUMO BELISE LTDA.xlsx`, aba Observacoes, diz que o lote 07 Ouro Verde II passou de 130k para 160k. O PDF (Lotes urbanos, p. 10) e o total de R$ 3.510.000 dão R$ 150.000. A mestre usa R$ 150.000, que está correto.
- Multas do próprio grupo, já marcadas DIVERGE na aba Totais de `sei-2025-itens.xlsx` (linhas 52-53): Vista da Colina R$ 2.556.900 contra 20% × 12,1 mi = R$ 2.420.000; Pedrotti R$ 2.280.646 contra R$ 2.420.000. São insumo para a Parte 4.

## Conferido e correto

- **Resumo.** Bens 2025 pelo grupo: R$ 152.796.260,60. Soma dos adotados: R$ 153.970.060,60. Pendente: 1 item (CONTR-06). Os três números foram recalculados a partir dos 93 itens de 2025 da Mestre e conferem com a tabela 41 do doc 11. As contagens de situação também conferem.
- **Somas por regra.** R1: 395 lotes em 28 loteamentos, R$ 103.459.000,00. R2: 12 itens, R$ 8.560.666,00. R3: 30 itens, R$ 16.515.016,00. Cobertura: R$ 128.534.682,00. Nos itens e lotes de R1, R2 e R3, valor adotado = valor do grupo, sem nenhuma divergência.
- **Lotes × loteamentos.** Nos 31 loteamentos, a soma dos lotes bate com o valor do item, tanto para o grupo quanto para o adotado. Lago Negro dá R$ 5.170.000.
- **Situações intermediárias do log.** Depois do pacote 2: R$ 148.635.060,60. Depois do pacote 3: R$ 153.970.060,60. As duas reproduzidas.
- **Lotes conferidos com as fontes.** 25 lotes sorteados conferem com o texto do doc 11. 18 lotes conferem com o PDF de Rafael. As 7 áreas rurais do pacote 3 conferem com o PDF Áreas rurais (p. 22-24). LURB-02 e LURB-07 conferem com o PDF Lotes urbanos (p. 10).
- **C0 nos 11 itens e lotes divergentes.** LURB-02, LURB-07, as 7 áreas rurais, Oleoplan e Lago Negro q01 (001 a 005): a marca de lista e o sentido menor/maior estão corretos.
- **Lista Diretrizes.** As 98 linhas da SEI (91 lotes e 7 bens) estão todas marcadas "Sim", e nenhum outro bem está marcado. Valores da lista iguais aos do grupo, exceto o Lago Negro 001/01 (R$ 1.100.000 contra R$ 770.000; adotado R$ 770.000). Os 20 lotes azuis da pasta Colina do Retiro II e os 2 lotes Firminio da pasta Bonetto conferem.
- **R2.** Rafael não dá valor diferente em lugar nenhum, inclusive no Quadro_Resumo. Itens 01, 06 e 07: ele mantém o valor do grupo mesmo citando faixas de mercado.
- **R3.** Para os 14 veículos, as 13 máquinas e os AURB-01 a 03, nenhum valor de Rafael nos 26 PDFs (ver achado 7).
- **R5.** Os 13 loteamentos têm saldo zero em 2018 e não são citados no doc 11 nem nos PDFs de Rafael.
- **R4.** Vale de fato para Dr. Letti e Sol Nascente III.

## Não conferido

- Casamento 2018 → 2025, itens "gone" e destinos (Parte 4).
- Os 26 lotes de terceiros, o acerto Rovílio e os compromissos (Partes 4 e 6).
- Bonetto (Parte 5).
- Planilhas Google.
- Imagens e cores dos PDFs: o azul da seleção foi lido nas transcrições xlsx e no texto da Diretrizes.
