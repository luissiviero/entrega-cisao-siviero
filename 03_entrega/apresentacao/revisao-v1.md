# Revisão independente — esboço v1 (28/10/2026)

O revisor foi um agente com contexto novo. Recebeu o texto do deck, `fontes-dos-numeros.md` e as fontes primárias, mas não as conclusões. Conferiu cerca de 240 números e os achou certos, com exceção dos itens abaixo. O deck foi regerado com as correções, e o gerador está em `_build/build_deck.py`.

| # | Achado do revisor | Gravidade | O que foi feito |
|---|---|---|---|
| 1 | Bonetto: Rafael dá ativo de R$ 3.111.000 (PM, W91). O total dos contratados passa de R$ 31.032.600 para R$ 34.143.600. | erro | Corrigido no slide 11: 7 loteamentos, "até R$ 34 mi". |
| 2 | Oleoplan: não há argumento técnico; a diferença de R$ 3 mi parece erro de soma. Além disso, `pendencias-conferencia-anterior.md` (item 3) contradiz o log. | risco alto | O slide 7 ficou provisório; a coluna mudou para "Argumento registrado"; o texto diz "ponto fraco". A contradição entre os arquivos vai ao Luís; o arquivo é da sessão SEI. |
| 3 | "Rafael não avalia nem contesta" é falso para a escavadeira 320D (estudo da Ceval). | risco | Mudado para "o estudo de Rafael para a SEI não avalia"; a escavadeira entra nas ressalvas. |
| 4 | "Concordam": em vários itens de R2, o valor de Rafael copia o balanço e fica abaixo da faixa dele. | risco | O título passou a "onde os dois valores são iguais", com uma ressalva explícita e a marca PROVISÓRIO. |
| 5 | C0 dedutível: os únicos valores menores adotados são os 2 itens da lista de Rafael; o lote 001 do Lago Negro não tem motivo técnico. | risco alto | A composição lote a lote do Lago Negro saiu do slide. O restante depende de decisão do Luís (ver a entrega). |
| 6 | Slide 9 dava a entender que os outros 20 bens já tinham destino. | risco | Reescrito: 3 com situação anotada e 17 ainda não pesquisados. |
| 7 | Áreas rurais 04, 08 e 09 estavam em 2018 no inventário da Agropecuária. | risco | Linha incluída no slide 10 e ressalva no slide 8. |
| 8 | Os R$ 25,73 mi dos terrenos estão pelo valor do grupo. | observação | O slide passou a dizer "pelo valor do grupo". |
| 9 | Pierozan: recebimento a conferir (AF152). | observação | "O que verificar" passou a ser o recebimento (Parte 4). |
| 10 | "Rafael cobra R$ 7,7 mi a seu favor" diverge da fonte (X72: "em favor da Valter"). O revisor também não achou nas fontes que Rafael é sócio. | risco | O texto passou a "diferença de R$ 7,7 mi a favor do lado que sai". "Rafael, outro sócio que sai" ficou mantido: vem do próprio Luís, no brief da sessão. |
| 11 | A revisão independente dos pacotes 1 a 3 veio depois da aprovação. | risco | O slide 3 agora diz isso. |
| 12 | Registro de data-base: só o trator tem tratamento; as avaliações de Rafael de 2026 estão pendentes. | risco | Incluído nos slides 3 e 16. |
| 13 | Slides 6, 7 e 8 sem a marca PROVISÓRIO, apesar da nova revisão prevista e das perguntas 3, 4, 7 e 8. | risco | Os três ficaram marcados; o slide 16 cita as perguntas. |
| 14 | Faltavam os 26 lotes de terceiros e os 8 do acerto Rovílio. | observação | Incluídos no slide 5. |
| 15 | Multas dos contratados (R$ 15,75 mi negativos; Dr. Letti e Sol Nascente III). | observação | Incluídas no slide 11. |
| 16 | A frase dos 84% sugeria que o resto diverge. | observação | O texto agora é "resolvidos em bloco, por valor igual ou avaliação única". |
| 17 | 6.700 m² × 18,9 − 18,2 ha. | observação | Mantido: é o texto do log (R2). |
| 18 | Onde se decide a "possível APP" (Parte 3 × Parte 4). | observação | O texto ficou neutro: "regra única, salvo documento". |
| 19 | Demais empresas: "a levantar" sugere que não há dado. | observação | Mudado para "não conferido". Os totais ficaram de fora, porque o brief limita a fonte à conferência de 26/10. |
| 20 | P8 do plano (apresentação só ao final). | observação | O deck é um rascunho pedido pelo Luís; a capa diz que se baseia em material em andamento. |

## Conferência visual (28/10/2026)

Os 16 slides foram exportados em imagem pelo PowerPoint instalado, com o script `_build/render.ps1`. As imagens ficam em `_build/png/`, e cada slide foi inspecionado. Não havia texto cortado nem sobreposição.

Defeitos corrigidos:
- valores quebrados entre linhas ("R$" de um lado, número do outro), nos slides 7, 9, 10 e 12. Agora há espaço não separável entre "R$", o número e a unidade (mi, mil, ha);
- itens de uma linha desalinhados das bolinhas, no slide 13;
- passo 2 colado no passo 3, no slide 15;
- espaço vazio nos cartões dos slides 6 e 12.

Nenhum número foi alterado nessa etapa.
