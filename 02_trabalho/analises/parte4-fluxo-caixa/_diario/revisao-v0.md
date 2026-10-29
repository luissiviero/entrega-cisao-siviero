# Revisão independente — Parte 4 SEI, fluxo de caixa v0

Revisor: agente de contexto novo (P11). Data: 28/10/2026.
Objeto: `_scripts/build_fluxo_caixa.py`, `fluxo-caixa-sei-v0.xlsx` e `numeros-v0.json`.
Método: recalculei a partir das fontes com Python/openpyxl. Rodei as funções do próprio script e as alterei uma de cada vez para medir cada efeito. Li as 409 linhas da lista lote a lote, não só uma amostra. Os MD5 das 7 fontes conferem com a aba Leia-me.
Convenção: "H" é a linha VALOR SEM EXPLICAÇÃO da Ponte. Os efeitos estão em R$ e na ordem piso / central / teto. O sinal + quer dizer que H sobe.

## Efeito indicativo no cenário central (H base = +0,37 mi)

| Achado | Efeito em H (central) |
|---|---|
| A1 + A2 + A3 + M1 (quantidades e preços dos lotes) | −9,03 mi (calculado em conjunto) |
| A4 concessão limitada ao que está documentado | até +7,05 mi |
| A5 parte em dinheiro das vendas "BEM" (estimativa) | cerca de +2,6 mi (mais até 3,1 mi de bens que não estão em 2025) |
| A6 dívida de 2019 corrigida | +2,55 mi |
| M2 bens novos | +1,9 a +2,4 mi |
| M3 (pró-labore pago ao Hotel Araçá, Flávia Manica) | +1,45 mi |
| A7 lotes além da redução, que o próprio arquivo 17 confirma | +17,1 mi (decisão de premissa) |

Esses efeitos não se somam de forma limpa, porque alguns se sobrepõem. Eles mostram duas coisas. Primeiro, os preços dos lotes estão altos demais em pontos específicos. Segundo, os usos (C2 a C4) estão inflados a favor do grupo. Nas duas direções, os erros são maiores que o próprio H central.

---

## Severidade ALTA

### A1. Lotes vendidos antes de 2018 entram como "sem destino" dentro da ponte
- **O que está errado.** Em `lot_model` (script, linhas 445–471), os registros anteriores a 2018 ficam fora de `usable`. A redução que eles explicariam vira `extra` e entra como classe SEM, com ano "sem registro", dentro de 2018–2024. Então lotes cuja venda está documentada em 2013–2017 aparecem como dinheiro sem destino no período.
- **Evidência.** Aba "Lotes que saíram 2018-2025": Pedra Lisa (linha 15: "4 das vendas têm data anterior a 2018"), Barbiero (5 vendas de 2017), Dalla Costa (linha 25: "as 2 vendas são de 2014 e 2017, para uma redução de 1"), Prata (linha 44: saídas de 1994 e 2001), Vicentino (2), Medianeira (1), Baú (1), Monte Vêneto I–II (1). No total são 16 lotes.
- **Efeito.** SEM (linha A6 e G) cai 3,54 / 4,56 / 5,86 mi, e H cai o mesmo valor.
- **Correção.** Os registros anteriores a 2018 consomem a redução primeiro, fora do período. Só o que sobrar vira SEM. Se a premissa P3 for mantida no sentido de que toda a redução é do período, então esses lotes devem aparecer numa linha própria ("venda documentada antes de 2018"), e não como "sem informação".

### A2. Saídas de 2025 sem data foram postas no meio do período
- **O que está errado.** Registros sem data recebem o ano de 2021 (`interp`) e entram na ponte. Mas pelo menos 5 deles são de 2025:
  - Colina do Retiro II 1/7, 12/8 e 15/8 (lista, linhas 130–132). O GERAL (linhas 311–314) marca esses lotes com "Não consta no inventário 2025 atualizado". A mesma marca aparece no 15/5, vendido em 01/08/2025, e no Lago Negro 14/5 (GERAL linha 410), permutado em 28/08/2025. Logo, a marca indica saída em 2025.
  - Morada O01/2 (linha 75). A aba resumo, linha 8, diz que saiu "entre dez/2024 e out/2025".
  - Lote urbano A.1 (linha 468): "Constava em dez/2024 e não consta em out/2025". Esse registro ainda foi classificado como BEM porque o texto "Casa da Nona" bate na regra "casa".
- **Efeito.** SEM cai 0,34 / 0,65 / 1,01 mi.
- **Correção.** Tratar esses registros como 2025, fora da ponte. Na regra BEM, trocar "casa" por um padrão que não pegue nomes de lugar, ou aplicar a regra só às frases "permuta de … casa".

### A3. Perini avaliado pelos lotes grandes que restaram
- **O que está errado.** Para os 7 lotes do Perini que saíram, o modelo usa v18 = 257.000 (média de 4 lotes a 500.000 e 6 a 95.000) e v25 = 1.520.000. Mas os 3 lotes que restam em 2025 são todos da quadra 1688, de R$ 1.520.000 cada (GERAL linhas 421–423). Os que saíram são lotes pequenos das quadras 1468 e 1469. O próprio grupo avalia esses lotes pequenos em R$ 160.000 em 2025: os lotes 11, 12, 18 e 19/1469 aparecem no 11.2.2 (GERAL linhas 694–697), e 18 e 19/1469 foram vendidos à Efetiva em 10/2024.
- **Efeito.** Usando 95.000 para 2018 e 160.000 para 2025, a linha VENDA cai 0,81 / 2,77 / 6,80 mi.
- **Correção.** Usar o valor do tipo de lote que saiu, quando o próprio grupo informa esse valor. Em geral, "Valor médio out/2025" é a média dos lotes que *restaram* (viés de seleção). Conferir também Morada do Ipê I (1 lote restante, a 629.000, aplicado a 28 saídas) e Lago Negro (ver M1).

### A4. Concessão de R$ 7,5 mi (C4) contradiz os números do próprio grupo e conta passivos de 2025 duas vezes
- **Evidência.**
  - Falavigna. No inventário de 2018 (`sei-2018-itens.xlsx`, Itens seq. 111), o grupo informa custo total da obra de 2.550.000, obra já realizada de 100.000 e, portanto, 2.450.000 a concluir. Em 2025 o grupo ainda lista −2.000.000 para concluir a pavimentação, o parque e a galeria (GERAL E737). Pelos números do grupo, no período foram feitos no máximo cerca de 0,45 mi. O arquivo 17 mostra obra do Falavigna paga em 2025 e 2026, "50% pagos em espécie pela empresa Siviero" (lista, linhas 421–427).
  - Colina do Retiro II. Em 2025 o grupo lista −3.800.000 para "conclusão da pavimentação…" (GERAL E719 e E720). Nos 15 meses, o dinheiro com a palavra "retiro ii" soma **R$ 145** (aba 4).
  - O custo de R$ 50.000 por lote vem do custo total de um loteamento novo, feito a partir de terra crua (E669: 12.100.000 / 242). Não serve para concluir obras que já estavam executadas em parte.
  - A aba Contratos 11.2, coluna "Proposta v0", mantém 11.2.3 (−3,8 mi) e 11.2.4, que inclui os −2,0 mi do Falavigna. Assim o grupo é creditado duas vezes: como gasto no período e como passivo de 2025.
  - O dinheiro dos 15 meses que foi para concluir obra aparece em Monte Vêneto (289 mil) e no acerto final de pavimentação do Barbiero (57 mil). Não aparece em Falavigna nem em Colina do Retiro II.
- **Efeito.** Até +7,05 mi em H, nos três cenários, se a concessão ficar em cerca de 0,45 mi.
- **Correção.** Limitar a concessão ao que está documentado: custo a concluir em 2018 menos o saldo a pagar em 2025. Outra opção é manter os 7,5 mi e retirar 11.2.3 e o item 19 de 11.2.4 dos passivos. Os dois juntos, não. A premissa "concluídos até 2025" é contradita pelo GERAL: levar ao Luís.

### A5. Classe BEM: o teste "o bem está no patrimônio" não foi feito, e muitos registros são vendas com parte em dinheiro
- **O que está errado.** 15,3 lotes do período (4,56 mi no central, de 6,87 mi da classe BEM) são vendas "à vista" ou "à prazo" com parte em veículo ou imóvel. A regra BEM vem antes da regra VENDA, então a parte em dinheiro se perde. Exemplos:
  - Sol Nascente I 24/D (lote − 67 mil)
  - Morada 001/01 (− Jeep Commander de 230 mil)
  - Morada 006/03 (− apartamento de 130 mil)
  - Morada 012/04 (− 75 mil)
  - Vicentino 03/1302 (− carro de 233 mil; central 1,12 mi)
  - Jardim Itália 13 e 14/1744 (R$ 140.000 em espécie)
  - Rivacci (R$ 55.000 em espécie)
  - Carnevalli 06/A (R$ 24.000 em espécie)
  - Altos da Julieta 03/1733 (− Amarok de 90 mil)
  - Lago Negro 25/06 (− RAV4 de 70 mil)
- **Os bens recebidos não estão no inventário de 2025.**
  - Nenhuma placa recebida aparece no GERAL 11.1.1, que só traz Hilux IYJ 8I14, Ranger e Compass GHX 9I84: Jeep Commander RXZ 7I62 (nos 15 meses há "ANUNCIO PUBLICACAO VENDA JEEP COMANDER", 08/2024), Hilux IZR 2E40, Amarok IXL 8H85, S10 JID 4877 (usada para pagar a pavimentação do Monte Vêneto), RAV4, Jeep Compass IXP 2137 (trocado pelo lote do Boito).
  - Também não aparecem o apartamento de Passo Fundo nem o pavilhão do Rivacci (370 mil).
  - Terras pagas com lotes também não aparecem como bens novos de 2025: Vila Azul, da Inês Ansolin Pessutto (R$ 714.000, 2024; o item de Vila Azul do GERAL, linha 80, tem o mesmo 2.263.666 em 2018 e 2025); 7 ha do Antônio Ansolin (R$ 586.666, 2021); 0,67 ha do Ivo Frainer (R$ 388.000, 2023).
  - Os móveis da sede do condomínio Morada dos Ipês II (Dhejota 287.959 e Essencial 80.669; linhas 80 e 97) são obra do loteamento, não bem da SEI.
- **Efeito estimado (central).** Cerca de +2,6 mi de dinheiro. Além disso, cerca de 1,4 mi em veículos e apartamento e cerca de 1,69 mi em terras recebidas que não estão em 2025: ou foram revendidos (entrada de dinheiro) ou são bens faltantes.
- **Correção.** Dividir cada registro em VENDA (parte em dinheiro) e BEM (valor do bem). Para cada bem recebido, procurar no inventário de 2025. Se não estiver lá, tratar como entrada de dinheiro ou pedir ao grupo. Os móveis da sede vão para OBRA ou investimento.

### A6. A dívida de jan/2019 inclui itens que não são dívida da SEI e dívidas pagas com lotes
- **Evidência** (`sei-2018-itens.xlsx`, aba Compromissos; aba 5 do livro).
  - Item 26, "Pagamento de compromissos mensais: títulos, luz, água, … salários, pró-labores, juros, impostos", **750.000**. É uma estimativa de custo corrente, não dívida, e duplica C1.
  - Itens 30 (TIS, 400.000) e 32 (HOTEL, 700.000). São limites de outras empresas; a própria aba 5 aponta isso, mas eles continuam somados.
  - Irineu Siviero, itens 8, 9, 10 e 28 (**949.000**). Foram pagos com 7 lotes em 21/12/2020, com o texto "devolução de empréstimo e pagamento de juros" (Medianeira 01 e 02/04, Lago Negro 23 e 24/06, Jardim Itália 04/1736, Femaçã 02 e 03/10). Esses lotes estão em SOC, fora do dinheiro, e a dívida quitada ainda entra em C2 como uso de dinheiro. É contagem dupla.
  - Item 24, Sicredi **5.400.000** (53% do saldo). A descrição não diz de qual empresa é a conta; os itens 23, 30, 31 e 32 dizem. É preciso confirmar com extrato.
  - Lado de 2025: a dívida do trator D6N com a Paraná Equipamentos (320.000) está em E738, mas o trator foi comprado em 09/2025 e o script o exclui dos bens. É incoerente. Além disso, o saldo da escavadeira está na posição de 31/07/2025 (33 parcelas); em 31/12/2024 eram cerca de 40 × 10.534,91 ≈ 421 mil.
- **Efeito.** Saldo de 2019 de 10,12 → 7,32 mi. Dívida de 2025 na data-base de 2,37 → cerca de 2,12 mi. C2 de 7,75 → cerca de 5,20 mi. H sobe cerca de 2,55 mi.
- **Correção.** Retirar os itens 26, 30 e 32 e a dívida do Irineu (ou reclassificar os 7 lotes como dívida paga com lote). Pôr a dívida de 2025 na posição de 31.12.2024, sem o D6N. Pedir o extrato da Sicredi.

### A7. A fonte do grupo contradiz a quantidade de 2018 (P3): os lotes "além da redução" são da SEI segundo o próprio arquivo 17
- **Evidência.** Arquivo 17 .xlsx, aba "Totais por loteamento":
  - Lago Negro: "QUANTIDADE DE LOTES EM 2018: 61 … + 03 APP … − 06 Família Siviero − 07 APP − saldo 05 = 46". O estudo de 2018 dizia 35 e a aba usa 27.
  - Altos da Julieta: 45 + 6 = 51 → 41 vendidos (a aba usa 40 → 30).
  - Baldasso: 17 → 12 (aba 15 → 10).
  - Morada dos Ipês: 32 → 30 (aba 29 → 28).
  - Monte Vêneto: 82 → 22 (aba 80 → 20).
  - Salvador: 45 → 13 (aba 36 → 10).
  - Pedra Lisa: 40 → 9.
  - Falavigna: 51 → 21.
- **Efeito.** Esses lotes correspondem ao "excedente" de 46 lotes que a aba Sensibilidade deixa fora da base: 9,86 / 17,06 / 24,27 mi. É mais que todo o H central.
- **Correção.** É decisão do Luís, porque mexe numa premissa: se o arquivo 17 vale para a quantidade de 2018 quando ele mesmo a declara. No mínimo, a Ponte deve dizer que o próprio grupo confirma esses lotes como estoque de 2018 vendido no período.

### A8. Os impostos e comissões pagos pela SEI tornam o cenário "piso" incoerente
- **Evidência** (contas classificadas, SEI, 15 meses). COFINS 317.334 e PIS 68.756. A razão entre os dois é 0,2167, igual a 0,65/3,00, o que indica regime cumulativo. A receita bruta tributada nos 15 meses fica em **R$ 10,58 mi**. As comissões de corretagem somam 407.299; com os 6% do próprio grupo (GERAL linha 733), isso dá **R$ 6,79 mi** de vendas. Nos mesmos meses (jun/2023 a set/2024), a linha VENDA do modelo dá 2,85 / 5,76 / 6,66 mi.
- **O que está errado.** C1 já inclui os impostos e comissões no nível real, cerca de 1,25 mi em 15 meses, que corresponde a cerca de 10 mi de receita. Ao mesmo tempo, o piso avalia as vendas em 2,85 mi. Um mesmo cenário não pode ter custos proporcionais a 10 mi de vendas e receitas de 2,85 mi. Ressalvas: a base da COFINS inclui aluguéis (cerca de 0,5 mi), serviços, parcelas de vendas antigas e permutas e dações, que também são tributadas.
- **Correção.** Mostrar esse teste na Ponte. Retirar o piso ou escalar impostos e comissões conforme o cenário. A evidência aponta para o teto ou acima dele.

---

## Severidade MÉDIA

### M1. Lago Negro: v18 inflado
- **O problema.** A aba "Valor médio dos lotes" divide 8.650.000 por 27 lotes e chega a 320.370. Mas o estudo de 2018 (Itens seq. 77) dá preço a 35 lotes: 5 × 350.000 e 30 × 230.000. O GERAL K403 pede verificar se os 35 incluem os lotes de APP.
- **Efeito.** Com 247.143 (8,65 mi / 35), VENDA cai 1,40 / 1,08 / 0 mi.
- **Correção.** Usar 247.143, ou 230.000 para os lotes comuns.

### M2. Bens novos (C3 = 3.877.000)
- **Linha 54, lote da Família Dal Ponte, 1.010.000.** Pelo arquivo 17, o lote foi pago em 10/09/2025 com 3 lotes (Colina I 08 e 09/14, Salvador 08/01; "permuta de outro imóvel, valor de R$ 1.010.000"). Antes disso houve parcelas em 2024: nos 15 meses aparecem 160 mil do consórcio Sicredi, 70 mil em "cheques soja Titton" e 4 × 13.800. O dinheiro dentro do período foi de cerca de 0,3 mi. Superestimativa: cerca de 0,7 mi.
- **Linha 76, Sapopema, 4,0 ha, Família Lago, 1.500.000.** Provavelmente é a mesma área do item A.2 de 2018 (3,6 ha, Família Lago, Sapopema, 900.000; GERAL K88: "Mesma área do item 2?"), somada aos 3.100 m² comprados em 2023 (R$ 350.000, pagos com o lote 06/01 do Solar Norte e parcelas). Superestimativa: cerca de 1,2 a 1,4 mi.
- **Bens de 2018 que não estão em 2025 e não entram como entrada nem como distribuição:**
  - lote 1543, 1,9 ha, perto da Oleoplan: 900 mil
  - 15,1 ha em Cotiporã (A.3): 226,5 mil
  - apartamento 303 do Ed. Boa Esperança: 100 mil
  - veículos: 173 mil, incluindo duas F1000 com os nomes "Pedro Siviero" e "Válter Siviero" (seriam SOC)
- **Correção.** Corrigir as linhas 54 e 76. Listar as baixas de bens na aba 6.

### M3. C1 "manter a empresa"
- **Pró-labore pago a empresa do grupo.** R$ 186.000 de "pró-labore" foram pagos a **HOTEL ARAÇÁ LTDA**. É transferência entre empresas. Extrapolado, dá 1,04 mi.
- **Pró-labore de sócios.** Válter, Adélia e Pedro recebem 558 mil, que extrapolados dão 3,12 mi. São 3 dos 6 ramos (o aluguel e o lucro são divididos por 6). Se o ramo que sai não recebe, isso é benefício de sócio, não custo da empresa. Decisão do Luís.
- **Flávia Manica.** É uma dívida de 500 mil em 50 parcelas. Extrapolada por 84 meses dá 840 mil: superestima cerca de 0,41 mi.
- **Consórcios.** As parcelas (118,6 mil) financiaram a compra do lote Dal Ponte, que também está em C3: contagem dupla parcial.
- **Gastos pessoais e políticos** extrapolados (funeral, doações e campanhas, despesas de familiares, empréstimo a familiares): cerca de 0,53 mi.
- **Extraordinários.** Os 514 mil de 15 meses são multiplicados por 5,6, dando 2,88 mi. A regra do Luís conta dívidas e multas pagas como "manter a empresa", mas extrapolar gastos pontuais é uma escolha. Mostrar à parte.

### M4. Sensibilidade com 72 meses é incoerente
- **O problema.** Custos e aluguéis são contados por 72 meses (2019–2024), mas as vendas continuam sendo as de 84 meses, com 2018 incluído. Pelos meus cálculos, as vendas em dinheiro de 2018 somam cerca de 2,5 / 2,5 / 7,0 mi.
- **Correção.** Se o período começa em 2019, tirar também as saídas de 2018. Isso fica coerente com a dívida medida em jan/2019. Na base de 84 meses há a incoerência inversa: a dívida é medida em jan/2019, mas as vendas e os custos começam em jan/2018.

### M5. A classe SOC mistura coisas diferentes, e a nota "compensar 1/6" engana
A regra `siviero` também roda no texto da forma de saída, então pega "Grupo Empresarial Siviero" e "pagos pela empresa Siviero". Os registros mal classificados são:
- pagamento de dívida a sócios (Irineu, André; linhas 103, 104, 172, 211, 212, 278, 451, 452; ver A6);
- gratificação a funcionário (linha 315, Carraro);
- indenização da Família Fioreze (linha 391, Valdir Angest);
- comissão (linha 85, Structura/Rovílio);
- troca por 2,5 ha em **Monte Bérico**, loteamento não finalizado (linhas 386, 387 e 389), que é investimento;
- venda do Attolini com R$ 80.000 em espécie por lote (linhas 449 e 450), que é em parte VENDA;
- registros de 2025 e 2026 da Florense e da Bett (linhas 421–427), que são OBRA do Falavigna.

Não muda H, mas muda as linhas A3 a A5 e a leitura de "distribuição".

### M6. Lotes que foram para loteamentos não finalizados ou distratados ficam fora da linha I
- **Evidência.** Na classe ACORDO há 38,3 lotes ligados a distrato, atraso ou não execução de loteamentos contratados: Susin (24 lotes + 200 mil), Via Nova (9 lotes + 90 mil), Ambitec, Colombo, Zattera, Fioreze, Vale das Colinas (linhas 247 e 396). Valem 3,74 / 6,16 / 9,85 mi. Há também o Morada dos Ipês II: área doada ao município (linha 96, ACORDO), móveis da sede (linhas 80 e 97, BEM) e paisagismo (linha 82).
- **O problema.** A linha I soma só H e A2. Esses lotes são justamente o valor que o grupo já gastou em 2018–2024 com loteamentos que não deram certo, enquanto ainda desconta 15,75 mi de multas em 2025.
- **Correção.** Mostrar esses lotes à parte, para a discussão do 11.2.1.

### M7. A amostra de preços citados (aba 8) não mede o valor do lote
Em 7 dos 12 registros, o valor citado é só uma parte do pagamento ou um acréscimo em dinheiro:
- Lago Negro 25/06: 70 mil é a RAV4.
- Alto das Graças 05/6372: 270 mil é o lote somado a 139,5 mil em dinheiro.
- Alto das Graças 14/6372: 62,5 mil é parte do preço.
- Alto das Graças 09/6373: 50 mil é o dinheiro somado ao lote.
- Altos da Julieta 12/1729: 200 mil é o dinheiro somado a 24 lotes.
- Monte Vêneto 11/09: 70 mil é o dinheiro somado ao lote.
- Morada 004/06: 80,7 mil são os móveis.

O filtro `partial` não pega "+ o valor de", "+ a importância", "sendo que a importância" nem "com permuta de uma Toyota". Em mais 2 casos de terra, a SEI também pagou em dinheiro: parcelas ao Ivo Frainer pelos 0,67 ha e ao Lindolfo Lago por Sapopema. A Pedra Lisa 15/6371, com 465 mil, é o valor do contrato inteiro de pavimentação.

Sobram como confiáveis só a Pedra Lisa 01/11 (1,66 × v25) e o San Marino 03/1342 (1,04 ×). A mediana de 1,02 não tem significado. Se mostra algo, é que esses registros ficam acima de v25.

### M8. Entradas que a ponte ignora
- **Soja.** "Cheques soja Titton" pagaram imóvel da SEI. Sacas de soja pagaram a escavadeira. A SEI paga 19 mil de impostos sobre soja.
- **Serviços.** Há notas de serviço da SEI para CERAN, CEVAL e Prefeitura (aba Dúvidas, item 1). Os impostos delas, cerca de 110 mil em 15 meses, estão em C1, mas a receita não está na ponte.
- **Contas a receber.** Nenhum dos dois inventários tem contas a receber. Vendas em 26, 30, 48 e 60 parcelas deixam saldo a receber em 31.12.2024, que é um ativo ausente em 2025.

---

## Severidade BAIXA

- **B1. Lotes urbanos.** São avaliados pela média (C47/11 = 256 mil; E47/8 = 411 mil), mas cada lote tem valor próprio no GERAL (A.1 220 mil, A.2 70 mil, A.3 100 mil, A.4 100 mil). No período, 769 mil no piso deveriam ser cerca de 270 mil: −0,5 / −0,7 / −0,96 mi, com sobreposição parcial com A2.
- **B2. Femaçã.** O v18 vem do vizinho (143,3 mil), mas o GERAL, linha 645, dá 120 mil. Não afeta H, porque esses lotes não estão em classes de dinheiro.
- **B3. Índice de 2,032 para loteamentos esgotados.** É conservador: o índice com o mesmo mix dá 2,43 e a mediana dá 2,19. Favorece o grupo. Documentar.
- **B4. Totais de lotes.** O JSON traz `reducao` = 336, mas a fonte dá 335 (Reserva Paradiso, −1, foi excluído). Nos lotes urbanos, a redução líquida é 3, mas saíram 4 e entrou 1, o lote da Dal Ponte de 2025.
- **B5. Código e planilha.** `assert … or True` (linha 755) desliga o teste da trajetória. A linha 616 é um placeholder sem efeito. Na aba "Relatórios Mensais" da planilha do Luís, o mês chamado "Set. 2023" tem o valor que é set/2024 na planilha classificada. As somas conferem.
- **B6. Contagem da Colina do Retiro II.** O arquivo 17 dá 99 lotes ao Siviero ("os 105 lotes fora obtida através de um levantamento") e saldo de 96. Já o GERAL de 2025 tem 102, incluindo 13 lotes marcados "Não consta no inventário 2025 antigo". A redução de 3 é incerta.
- **B7. Boito.** O lote conta como de 2018, mas não está no inventário de 2018: foi recebido em 2023 em troca do Jeep Compass IXP 2137, que por sua vez veio do lote Carnevalli 06/A. Não altera a redução.
- **B8. Juros.** Os juros de 2023–24, cerca de 17,4 mil por mês, são extrapolados para 2018–2024. Em 2019–2020 a dívida era bem maior, então os usos dos primeiros anos ficam subestimados (a favor do sócio). Declarar.

---

## Verificações que não encontraram problema
- Aritmética da Ponte nos três cenários: A a I, H/6 e usos de 47.118.233.
- Fórmula da Sensibilidade, exceto a incoerência de M4.
- Total de 11.2 (E664 = −51.922.233,03 = E666 + E677 + E719 + E721 + E738) e Rafael (F664 = −17.778.633). F721 de fato já inclui o 11.2.5.
- Total da SEI nos 15 meses: 9.877.224,33, igual à coluna Empreendimentos de "Relatórios Mensais". Totais por classe (M = 4.998.032,98; INV = 1.233.601,36; LOTE, BENS, DIV) e média de 333.202,20 por mês, todos recalculados.
- Aluguel de 35.326,83 por mês (34.930 mensais + 4.762 anuais / 12).
- Saldo de jan/2019 de 10.119.900 (o documento diz 10.120.000) e empréstimos de 2025 de 2.367.652,03 (soma de 5 itens).
- Bens novos: 9 linhas somando 3.877.000, com o D6N corretamente fora.
- Quantidades por loteamento e totais 761 / 426 / 335 iguais à linha 47 da aba. Valores de "Valor médio dos lotes" transcritos corretamente.
- Saídas datadas de 2025 fora da ponte (POS2025). Saídas no período mais saídas posteriores somam 336.
- Transcrição do arquivo 17 conferida em amostra (Reali, Efetiva, Attolini, Dal Ponte).
- Mapa de classes dos 15 meses sem categoria órfã. MD5 das fontes conferem.
