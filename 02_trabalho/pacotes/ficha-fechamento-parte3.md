# Ficha de fechamento — Parte 3 da SEI (concorda × discorda)

28/10/2026. Números de `_scripts/ficha_parte3.py`, lidos da planilha-mestre (md5 `9484fb5a4e`). Revisão independente de todas as decisões aprovadas: `_diario/revisao-fechamento-parte3.md`.

## 1. Decisões aprovadas

| Pacote | O que decidiu | Valor |
|---|---|---|
| 1 | R1 a R5: lotes e bens com o mesmo valor nas duas fontes, bens avaliados só pelo grupo, contratados sem valor, saldos zero de 2018 | R$ 128.534.682 cobertos |
| 2 | C0 (bem que vem para nós: menor valor; bem que fica: maior) e C1. Oleoplan por Rafael (−R$ 3.000.000). Lago Negro R$ 5.170.000, com os 7 lotes de APP em R$ 0. Casa da Nona e Ouro Verde II por Rafael. Trator D6N e a dívida dele excluídos pela data-base | por item |
| 3 | 7 áreas rurais por Rafael (C0 + C1) | + R$ 3.254.800 |
| 4 | Pares 2018 × 2025 e destinos de 12 bens de 2018 | sem efeito em valor |
| Revisão 1 | 6 contratados e Bonetto adiados para a Parte 5. Simioni reaberto (fechado no pacote 5) | — |
| 5 | R6: regra única de "possível APP" pelo mapa oficial, leitura intermediária. Simioni em R$ 260.000. S1 aplicado | bens + R$ 1.187.500; passivo de terceiros de R$ 16.240.000 para R$ 8.570.000 |

## 2. Números

| Bloco | Itens | Grupo 2025 | Adotado | Diferença |
|---|---|---|---|---|
| Loteamentos executados e registrados | 31 | R$ 111.179.000,00 | R$ 113.515.500,00 | + R$ 2.336.500,00 |
| Áreas urbanas em Veranópolis | 4 | R$ 26.929.394,60 | R$ 23.929.394,60 | − R$ 3.000.000,00 |
| Áreas rurais | 12 | R$ 7.663.866,00 | R$ 10.918.666,00 | + R$ 3.254.800,00 |
| Lotes urbanos | 8 | R$ 3.290.000,00 | R$ 3.510.000,00 | + R$ 220.000,00 |
| Máquinas | 14 | R$ 2.517.000,00 | R$ 2.067.000,00 | − R$ 450.000,00 |
| Veículos | 15 | R$ 1.217.000,00 | R$ 1.217.000,00 | — |
| Áreas contratadas (Parte 5) | 9 | R$ 0 | R$ 0 (7 sem decisão) | — |
| **Bens** | **93** | **R$ 152.796.260,60** | **R$ 155.157.560,60** | **+ R$ 2.361.300,00** |

**Patrimônio líquido, provisório:** as Partes 4 e 5 ainda vão mexer nos passivos e nos contratados.

| | Doc 11 | Adotado até a Parte 3 |
|---|---|---|
| Bens | R$ 152.796.260,60 | R$ 155.157.560,60 |
| (−) Loteamentos contratados (multas) | R$ 15.754.581,00 | R$ 15.754.581,00 (Parte 5) |
| (−) Lotes de terceiros indenizáveis | R$ 16.240.000,00 | R$ 8.570.000,00 |
| (−) Obras do Colina do Retiro II | R$ 3.800.000,00 | R$ 3.800.000,00 (Parte 4) |
| (−) Demais compromissos (obs. 1–24) | R$ 16.127.652,00 | R$ 15.807.652,00 (sem a dívida do trator; Parte 4) |
| **Líquido** | **R$ 100.874.027,60** | **R$ 111.225.327,60** |
| **1/6** | R$ 16.812.337,93 | **R$ 18.537.554,60** (+ R$ 1.725.216,67) |

## 3. Revisão de fechamento: o que achou e o que já foi feito

O revisor regenerou a planilha-mestre e ela é idêntica à oficial. Os totais do log batem, o C0 foi aplicado no sentido certo e não há bem contado duas vezes na soma. Relatório completo: `_diario/revisao-fechamento-parte3.md`.

**Já aplicado (eram decisões suas ainda não gravadas, ou correções de texto):**
- Dívida do trator (R$ 320.000) excluída do passivo, junto com o trator (pacote 2, item 7).
- Lote Alto das Graças I H-6371/011 suspenso da lista de Rafael (S4 do ACERTOS).
- S6 lançado no registro de data-base (linha 12).
- Registro 4 (Rovílio) corrigido: altera valores na SEI e vai para a Parte 4.
- Notas antigas sobre o Oleoplan, que diziam o contrário da sua decisão, anotadas com a decisão aprovada.

**Precisa de você para fechar (detalhes abaixo, na seção 4):**
1. **27 lotes que eram da SEI em 31.12.2024 não estão no doc 11.** Foram vendidos ou dados em pagamento em 2025 e 2026 (arquivo 17). O doc 11 mostra o estoque do fim de 2025, não o da data-base. Conferi por script: são 27 (`saidas-pos-database.xlsx`). Só 8 estavam no S6. Pelos vizinhos, a ordem de grandeza é de R$ 9,2 milhões, 1/6 ≈ R$ 1,53 milhão. É uma estimativa do revisor, não uma avaliação.
2. **Dois bens entraram depois da data-base.**
   - O lote urbano 06 (Dal Ponte, R$ 1.010.000) veio em 10/09/2025, em troca de 3 desses 27 lotes, e **está na lista de Rafael**.
   - O Reserva Paradiso (R$ 500.000) veio em 19/09/2025 pelo Vicentino 1301/003 e foi revendido em 2026.
3. **Valores de Rafael de 2026 tomados como valores de 31.12.2024 (registro 3).** Sustentam + R$ 4.623.800 nos bens que ficam (áreas rurais, Casa da Nona, Ouro Verde II, Lago Negro 002 a 005), 1/6 ≈ R$ 770.633. É o primeiro alvo do grupo, porque são avaliações sem laudo.
4. **Áreas rurais 11 e 12.** O valor de Rafael (+ R$ 710.000) não considera a posse não registrada nem a parte destinada à compensação ambiental do Colina do Retiro II.
5. **Lago Negro.** A decisão judicial que confirma a APP não está nas fontes. Na data-base, os lotes 18 e 19 eram de compradores, e a R6 estendida ao passivo permite ao grupo pedir provisão para eles.

## 4. Recomendações para as decisões de fechamento

| # | Recomendação | Por quê |
|---|---|---|
| 1 | Fechar a Parte 3 com os 27 lotes registrados (linhas 12 e 13 do registro) e sem valor. Tratar destino e valor na Parte 4, pedindo ao grupo o valor de cada um em 31.12.2024 | Nenhuma das duas fontes avalia esses lotes. Pela P5, sem o valor do grupo, só cabe "caso extremo" com justificativa, melhor decidido com a Parte 4 |
| 2 | Manter o lote urbano 06 e o Reserva Paradiso na soma, como provisórios (linha 14), e **suspender o lote urbano 06 da lista de Rafael**, como no S4 | Tirar só os bens recebidos, sem pôr de volta os lotes que saíram, prejudica o Luís. A troca entra uma vez só quando a Parte 4 decidir. Rafael escolheu um bem que a SEI não tinha na data-base |
| 3 | Aceitar os valores de Rafael como valores da data-base, com a justificativa escrita: o próprio grupo apresenta o estoque e os valores de out./2025 (arq. 1, p.28: "termo final [...] 31 de outubro de 2025"). Listar como premissa frágil | Os dois lados avaliaram depois da data-base. A P5 manda escolher entre as duas avaliações |
| 4 | Manter o valor das áreas 11 e 12 e perguntar ao grupo a área destinada à compensação e a situação da posse | O C2 só se aplicaria com a área de compensação conhecida |
| 5 | Manter o Lago Negro em R$ 0 (C2) e pedir a cópia da decisão. Na Parte 4, tratar a provisão dos lotes 18 e 19 junto com os R$ 580.000 já pagos | O efeito da R6 no passivo desses lotes é o próprio custo já pago depois da data-base; não é um passivo novo |

**Decisões sobre premissa frágil, para a negociação:**
- Oleoplan: R$ 10.631.378,60. O valor de Rafael é a soma com erro, não uma avaliação.
- Valores de Rafael de 2026.
- R6 na leitura intermediária: depende do texto da Lei 4.814/2023.
- Simioni: depende do TAC.
- Lago Negro: depende da decisão judicial.
- Áreas rurais 11 e 12.

## 5. Pendências

**Na Parte 3 (dependem do grupo):**
- 6 lotes "FAMÍLIA SIVIERO" do Salvador, quadra 05, sem valor (pergunta 1). Os vizinhos valem de R$ 374.000 a R$ 387.000.
- Texto da Lei 4.814/2023 e cópia do TAC do Simioni (pergunta 2). Se a lei confirmar que a faixa de 30 m não restringe, a R6 passa à leitura estrita: + R$ 4.513.000.

**Vão para as próximas partes:**
- **Parte 4 (abordagem a explicar por você):**
  - observações 1 a 24 (inclui Perini, Dalzóchio, Acordi e Rovílio);
  - obras do Colina do Retiro II;
  - lado do ativo do S1;
  - passivo SAFRA na data-base;
  - indenização do Lago Negro (R$ 580.000, out./2025);
  - destinos de 2018 sem documento;
  - perguntas 3 a 10;
  - achados S2, S3, S4, S6 e S9 do ACERTOS e os da S3 (Rovílio).
- **Parte 5:** os 7 loteamentos contratados (R$ 31.032.600 segundo Rafael) e o Bonetto; o crédito Boff (S7).
- **Parte 6:** candidatos 1 a 12 de `parte6-candidatos-duplicidade.md`. Os principais são Dalzóchio e Perini (contingência + indenização), Perini e Dalzóchio como possível mesmo processo, e a escavadeira 320D.
- **Parte 7:** as observações do Luís nas planilhas Google.

**Registro de data-base:** linhas 1 a 5 e 8 a 9 com tratamento pendente.
