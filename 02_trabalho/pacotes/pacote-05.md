# Pacote 5 — regra única para "possível APP" (SEI)

28/10/2026, versão 3. A versão passou por duas revisões independentes, incorporadas aqui:
- `_diario/revisao-pacote-05.md`: levou à consulta do mapa oficial;
- `_diario/revisao-pacote-05-v2.md`: conferiu o mapa e mostrou a decisão central.

Os números vêm de `_scripts/pacote05.py`, que deve ser rodado com o nome da leitura (`estrita`, `intermediaria` ou `ampla`). O mapa vem de `_scripts/geofarroupilha_app.py` e o resultado está em `geofarroupilha-app-lotes.xlsx`; os dados brutos estão em `01_fontes/geofarroupilha/`. O script confere lote a lote que os valores de Rafael batem com os do grupo.

Nada foi gravado em `decisoes-sei.json`.

## 1. O que o mapa oficial mostra

O grupo diz que os lotes estão em faixas de APP do Marco Regulatório de Farroupilha (Lei Municipal 4.814/2023, arq. 1, p.26 e 28). O mapa da lei é público. Localizei os 95 lotes da SEI em Farroupilha e medi a parte de cada um em cada camada. O segundo revisor conferiu a identificação e as medições.

| Camada do mapa | O que é (leitura pelo nome e pela geometria; o texto da lei não foi obtido) |
|---|---|
| APP 30 m, APP 50 m | APP da lei municipal |
| Faixas não edificantes de 15 m e 5 m | Faixas onde a lei municipal proíbe construir |
| "Sensibilidade ambiental" | Faixa de 30 m em volta de **todos** os cursos d'água. É a medida da APP federal (Código Florestal) para riachos. Sem descrição no serviço |

**Os lotes que o grupo marcou são exatamente os que tocam a faixa de 30 m.** São 33 de 33, e nenhum dos 62 lotes próprios avaliados cheios a toca de forma relevante. Só 22 dos 33 tocam a APP municipal ou as faixas não edificantes.

Uma leitura possível, sem o texto da lei: a lei municipal, feita para áreas urbanas consolidadas, reduziu a faixa federal de 30 m às faixas de APP e não edificantes do mapa. A faixa de 30 m seria a referência federal, que não vale mais ali. O MP-RS, porém, lista a lei entre as que reduziram limites de APP, sinal de que a contesta. Se a lei cair, volta a faixa federal.

## 2. A regra R6 (para aprovar em bloco)

É uma **regra de decisão** sob o C2 (percentual do lote cheio), não uma reavaliação. O lote cheio é sempre o valor do grupo ou o de Rafael, como manda a P5. A mesma regra vale para o ativo e para o passivo; **a extensão ao passivo de terceiros precisa da sua aprovação expressa**.

| Nível | Lote próprio | Lote de terceiro (indenização) | Quando se aplica |
|---|---|---|---|
| **A** | 0 % | 100 % | APP reconhecida no lote por decisão, perícia ou gravame; ou 50 % ou mais do lote nas camadas restritivas |
| **B** | 50 % | 50 % | De 10 % a 50 % do lote nas camadas restritivas; ou dúvida documentada (ver as leituras abaixo e o Simioni) |
| **C** | 100 % | 0 % | Menos de 10 % do lote nas camadas restritivas |

O meio (50 %) traduz a incerteza: o lote pode ou não perder o uso. A camada "Revisão MP" do mapa **não** é usada: não tem descrição e traz processos de 2026, depois da data-base.

## 3. A decisão central: como ler a faixa de 30 m

| Leitura | Faixa de 30 m | Efeito da R6 | 1/6 | Soma adotada dos bens | Passivo de terceiros |
|---|---|---|---|---|---|
| Estrita | não conta | + R$ 11.483.500 | R$ 1.913.916,67 | R$ 155.522.560,60 | R$ 4.422.000 |
| **Intermediária (recomendada)** | lote nela fica no mínimo em B | **+ R$ 6.970.500** | **R$ 1.161.750,00** | **R$ 155.157.560,60** | **R$ 8.570.000** |
| Ampla | conta como APP | + R$ 2.782.000 | R$ 463.666,67 | R$ 154.356.560,60 | R$ 11.957.500 |

Em todas as leituras soma-se o S1, que você já aprovou: + R$ 1.887.000, ou R$ 314.500 no 1/6, aplicado uma vez só. O passivo do doc 11 é de R$ 16.240.000.

**Por que recomendo a intermediária:**
- Ela trata como dúvida, que é o nível B, exatamente o que é duvidoso: lotes que a lei municipal em vigor na data-base parece liberar, mas que a faixa federal atingiria se a lei cair.
- A estrita depende de uma leitura da lei que não pude conferir.
- A ampla aceita a marcação do grupo sem considerar a lei que o próprio grupo invoca.
- **Se você obtiver o texto da lei e ele confirmar que a faixa de 30 m não restringe, a leitura estrita passa a ser a correta.** O ganho seria de + R$ 4.513.000 no patrimônio, ou R$ 752.166,67 no 1/6.

## 4. Aplicação na leitura intermediária: 10 itens, do maior valor para o menor

Antes da regra saem do passivo de terceiros os 4 lotes do S1 (R$ 1.887.000).

| # | Lado | Loteamento | Lotes | Níveis | Grupo | Rafael | Recomendado |
|---|---|---|---|---|---|---|---|
| 1 | passivo | Industrial Vicentino | 9 | 1 A, 8 B | R$ 12.195.000 | mesmos valores (1) | **R$ 6.775.000** |
| 2 | ativo | Industrial Vicentino 013/1303 | 1 | B (32,6 % em faixa não edificante) | R$ 1.125.000 listado, fora do total | R$ 1.125.000 | **R$ 562.500** |
| 3 | passivo | Jardim Itália | 4 | 2 A, 2 B | R$ 786.000 | mesmos valores (1) | **R$ 586.000** |
| 4 | ativo | San Marino 012, 013/1650 e 004/1676 | 3 | B (só na faixa de 30 m: 56 %, 95 % e 31 %) | R$ 730.000 listado | R$ 730.000 (2) | **R$ 365.000** |
| 5 | passivo | Perini | 4 | 3 A, 1 B | R$ 640.000 | mesmos valores (1) | **R$ 560.000** |
| 6 | ativo | Fracionamento Simioni | 4 | B (TAC; ver abaixo) | sem valor ("ÁREA DE APP") | R$ 520.000 (3) | **R$ 260.000** |
| 7 | passivo | Firminio Dalzóchio | 3 | 3 A (58 % a 65 %) | R$ 400.000 | mesmos valores (1) | **R$ 400.000** |
| 8 | ativo | Primavera 008 e 009/1193 | 2 | A (82 % e 64 % em APP de 30 m) | R$ 335.000 listado | R$ 335.000 | **R$ 0** |
| 9 | passivo | Primavera 005 e 006/1199 | 2 | 005 A (90 %), 006 B (só faixa de 30 m) | R$ 332.000 | mesmos valores (1) | **R$ 249.000** |
| 10 | ativo | Firminio Dalzóchio 002/1292 | 1 | A (70 % em faixa não edificante de 15 m) | R$ 131.000 listado | R$ 131.000 | **R$ 0** |

**Notas da tabela:**
- (1) Rafael transcreve os mesmos valores, trata os R$ 16.240.000 como patrimônio, calcula um crédito de 1/6 e pede relatório técnico de APP por lote (terceiros, p.2-5).
- (2) O "R$ 7.700.000,00" que Rafael escreve no San Marino (p.94) é erro de cópia do Lago Negro.
- (3) R$ 520.000 = 4 × R$ 130.000, o piso de "lotes padrão" de Veranópolis, com a ressalva "Verificar documentação relativa às APPs" (p.17-21).

**Pontos de atenção:**
- **Lotes perto dos limites.** O Perini 018 tem 50,6 % do lote nas camadas; faltam 2,9 m² para o limite de 50 %. O Vicentino 016/1303 tem 13 % e fica em B por esses 3 pontos acima do limite de 10 %. Uma regra proporcional (valor × parte livre do lote) evitaria os saltos, mas daria a falsa impressão de que o mapa tem precisão de metro.
- **Simioni: decisão própria.** A recomendação de nível B (R$ 260.000) vale porque não há mapa da nascente nas fontes. As outras leituras são R$ 0 (APP confirmada, C2: o doc 11 diz "ÁREA DE APP", e o TAC reconhece a nascente e restringe a construção) e R$ 520.000 (valor cheio: em 2018 o grupo avaliava os 4 lotes em R$ 1.000.000, com o mesmo texto de restrição).
- **C0.** Os lotes próprios ficam na empresa pela definição aprovada da lista de Rafael (PDF Diretrizes). Peço que você confirme essa leitura: 5 dos 7 aparecem nas tabelas de "lotes escolhidos" de Rafael, fora do total. Para bens que ficam, o C0 daria o valor cheio. A R6 prevalece porque é a sua regra de APP (C2), mas isso custa ao Luís **R$ 275.583,33 no 1/6** na leitura intermediária.

## 5. Efeito na leitura intermediária

| | Hoje | Com o pacote 5 |
|---|---|---|
| Soma dos valores adotados (bens) | R$ 153.970.060,60 | **R$ 155.157.560,60** (+ R$ 1.187.500) |
| Passivo de lotes de terceiros | R$ 16.240.000 (doc 11) | **R$ 8.570.000** |
| Efeito da R6 | | **+ R$ 6.970.500; 1/6 = R$ 1.161.750,00** |
| Efeito do S1 (já aprovado) | | + R$ 1.887.000; 1/6 = R$ 314.500 |

**Lado do ativo do S1, em aberto:** na data-base, os 4 lotes eram estoque, com estes níveis pela R6:
- Jardim Itália 001: A;
- Primavera 010 e 011: A;
- Vicentino 003/1301: B, porque está 100 % na faixa de 30 m.

O doc 11 já traz o lote 623/104 do Reserva Paradiso (R$ 500.000), recebido na venda de 2025. Fica para o registro de data-base e a Parte 4.

## 6. Amostra por lote (conferir 5 linhas)

| Lote | No mapa | Nível | Cheio | Recomendado |
|---|---|---|---|---|
| San Marino 013/1650 (próprio) | 0 % em APP/faixa; 94,6 % na faixa de 30 m | B | R$ 233.000 | R$ 116.500 |
| Primavera 008/1193 (próprio) | 82,3 % em APP de 30 m | A | R$ 170.000 | R$ 0 |
| Vicentino 004/1302 (terceiro) | 0 % em APP/faixa; 24,3 % na faixa de 30 m | B | R$ 1.355.000 | R$ 677.500 |
| Vicentino 001/1303 (terceiro) | 63,8 % em APP/faixa | A | R$ 1.355.000 | R$ 1.355.000 |
| Perini 011/1469 (terceiro) | 100 % em APP/faixa | A | R$ 160.000 | R$ 160.000 |

## 7. Riscos e argumentos

- **Escrow:** o Dr. Pablo, do nosso lado, propôs uma conta garantidora para esses lotes (arq. 1, p.19-20), e o grupo recusou (p.28). A R6 dá valores definidos no lugar dela.
- **Argumento de reserva:** o passivo de R$ 1.355.000 por lote do Vicentino está acima da faixa de mercado de Rafael para o loteamento, de R$ 980.000 a R$ 1.125.000 (p.110).
- **Prioridade:** pedir o texto da Lei 4.814/2023 e o significado das camadas ao grupo ou à prefeitura. É o que decide entre a leitura intermediária e a estrita.

## 8. Descontos duplos que a regra não resolve

Estão em `parte6-candidatos-duplicidade.md`, itens 6 a 12:
- **Dalzóchio:** indenização cheia dos 3 lotes de terceiros, lote próprio sem valor e mais R$ 2.200.000 de contingência pela fonte de água no mesmo loteamento (obs. 3).
- **Perini:** R$ 560.000 de indenização mais R$ 5.000.000 de contingência (obs. 2).
- **Perini e Dalzóchio:** podem ser o mesmo processo.
- **S5:** são os lotes dados em acordo.
- **Perini 018 e 19:** vendidos pela SEI em out./2024; os dois estão em APP no mapa.

## 9. O que gravo se aprovar

- **`decisoes-sei.json`:** a regra R6 com a leitura escolhida; os 7 lotes de possível APP; o Simioni, que fecha o item reaberto LOT-04; o passivo de terceiros por lote; o S1, marcado como aplicado aqui.
- **`log-decisoes.md`:** o registro do pacote 5.
- **Planilha-mestre:** regenero e confiro a soma. Corrijo também a marca "na lista de Rafael" desses lotes.
- **`registro-data-base.md`:** os 4 lotes do S1 e o lado do ativo deles; o passivo SAFRA na data-base; o mapa na versão de 2026.
- **Pergunta 2 ao grupo:** atualizada para pedir o texto da Lei 4.814/2023 com o significado das camadas e a cópia do TAC do Simioni.
- **`00_FONTES.md`:** o GEOFarroupilha entra como fonte.
