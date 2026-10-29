# Revisão independente — cruzamento dos compromissos do doc 11 (11.2.2 a 11.2.5) com o arquivo 17

Revisor: agente de contexto novo (P11). Data: 28/10/2026.
Objeto: `_scripts/cruzar_compromissos_arq17.py` e `cruzamento-compromissos-arq17.xlsx`, abas Cruzamento e Avaliação.
Fontes conferidas:
- texto integral de GERAL, coluna B, linhas 677–743;
- arquivo 17 .xlsx, aba "Lotes vendidos (PDF)", com cada lote do 11.2.2 procurado por lote e quadra, não só por palavra-chave;
- arquivo 16 (OCR completo dos dois Termos);
- `diff-inventario-2025-antigo-x-atual.md` e `acertos-e-areas-vendidas/achados-para-sei.md` (S1 a S6).

**Critério.** Só há cobrança dupla quando o mesmo valor reduz o patrimônio duas vezes: o lote sai do estoque E a dívida que esse lote pagou continua no passivo. Se o lote saiu e o passivo é o que **resta** depois dessa entrega, não há cobrança dupla.

## Resultado por caso

### D1 Susin (11.2.4-15, R$ 1,2 mi): **não confirmado como cobrança dupla; fica como pergunta**
**O que dizem os acordos (arquivo 16)**
- **Termo A, 13/07/2020.** Susin e AGI recebem 17 lotes e R$ 200 mil. Dá quitação de "todas as intermediações e/ou prestação de serviços até a presente data" (cl. 1.1, 5, 5.1 e 6) e cita as famílias Piccoli, Boff, Andreis, Pedrotti e Fiorese, que são donas das áreas de Caminhos do Sol, Vale das Colinas, Pedrotti/Andreis e Vista da Colina. A cl. 8 exclui "novas intermediações … a partir da presente data".
- **Termo B, 30/04/2020.** São 7 lotes: o 12/1729 do Altos da Julieta, 5 lotes no Alto das Graças e o 7º (Pedrotti). Os destinatários são Alexandre Venzon, Bernardete Venzon, Valmir Susin e Alexandre Susin.

**Os contratos Buriti não estão no Termo A, e a data deles não está nas fontes.**
- Os contratos com a Buriti Holding Sul (50% Buriti, 10% SEI, 40% proprietários) são outro negócio: SEI com a Buriti, não SEI com as famílias.
- No inventário de 2018 a SEI ainda tinha 290/577 lotes no Caminhos do Sol e 180/359 no Vale das Colinas, sem Buriti. Logo, o negócio com a Buriti é de 2019–2025.
- Se for anterior a 13/07/2020, a comissão está quitada pela cláusula geral. Se for posterior, a cl. 8 a deixa fora do acordo. **Pedir a data e o contrato Buriti.**

**Contagem de lotes da aba Avaliação: errada**
- No arquivo 17 há 16 registros com Susin ou AGI:
  - 1 em 18/04/2019 (Viero 05/1469, "dívida de comissão, negociado com Rodrigo Razzera"): é anterior aos Termos e não faz parte deles;
  - 1 em 30/04/2020 (12/1729, o lote "a" do Termo B);
  - 14 em 13/07/2020, que valem 15 lotes, porque os lotes 05 e 06/6364 "equivalem a 03 lotes".
- Ou seja, foram entregues 16 lotes-equivalentes dos 24 dos acordos, não "25 lotes".
- A frase "num total de 24 lotes" do arquivo 17 é o total dos dois Termos (17 + 7), não o que foi entregue.

**Valor da aba Avaliação**
- Os R$ 3.792.104 usam a média de 2025 de 16 registros, entre eles o Viero de 2019. Esse valor não tem relação com os R$ 1,2 mi: os lotes pagaram os Termos de 2020, não a comissão Buriti.

**Os 6 lotes que faltam**
- O arquivo 17 não registra nenhuma entrega a Venzon, Bernardete ou Valmir. Isso é compatível com 6 lotes do Termo B ainda devidos, como dizem o doc 11 e o achado ACERTOS S2.
- Mas o E736 só põe valor na comissão (R$ 1,2 mi). Os 6 lotes devidos continuam dentro do estoque de 2025, o que favorece o sócio que sai.
- Proposta da aba sobre os 6 lotes do Termo B: concordo. Retirar a comissão só se o contrato Buriti for anterior a 13/07/2020.

### D2 Falavigna (11.2.4-19, R$ 2,0 mi): **possível, não provável**
- Os 5 lotes conferem: 020, 021 e 023/174 com a Florense em 20/03/2025; 012 e 013/175 com a Bett em 10/07/2025. Nos dois casos o texto diz "50% pagos em espécie pela empresa Siviero e 50% em permuta do lote".
- Todas as entregas são de 2025. A estimativa de R$ 2,0 mi é da versão atual do doc 11; na versão antiga eram R$ 950 mil. Não se sabe se os R$ 2,0 mi são o que ainda resta depois dessas entregas ou se incluem o que já foi pago.
- Só há cobrança dupla no segundo caso.
- Na data-base o efeito é o mesmo nos dois casos: os 5 lotes eram estoque da SEI em 31.12.2024 e as obras ainda não estavam pagas. O certo é devolver os lotes ao estoque e somar ao passivo o que eles pagaram.
- Valor: 5 × 607.200 = 3.036.000 (média 2025 do Falavigna) confere. Mas os lotes pagaram só 50% de cada fornecimento, então esse valor não é o valor da obra quitada.
- **Pedir a composição dos R$ 2,0 mi.**

### D3 Ambitec 02/1747 e Farina 11/1469 (11.2.2): **rejeitado como cobrança dupla**
- O 11.2.2 é, por definição, uma lista de "lotes de terceiros passíveis de indenização por caracterização ambiental": lotes que a SEI já vendeu ou deu, e cujos donos podem pedir indenização se o lote virar APP.
- Ter saído do estoque é exatamente o pressuposto do item. Não é cobrança dupla.
- Continuam valendo os dois argumentos do próprio Luís (K677 e K678): é uma contingência, não uma dívida certa; e está avaliada pelo preço de 2025, não pelo que a SEI recebeu.
- Os valores (200.000 + 160.000) conferem.

### D4 lotes da SEI vendidos em 2024–2025 (11.2.2): **confirmado só para os que saíram em 2025**
- 18 e 19/1469 (Perini, vendidos à Efetiva em 10/10/2024): já eram de terceiros na data-base. Mesma leitura que D3: contingência, não cobrança dupla.
- 01/1747 (10/03/2025), 11/1199 (dação à Safra, 27/03/2025) e 03/1301 (19/09/2025): **eram estoque da SEI em 31.12.2024**. Na data-base, voltam ao estoque e saem do 11.2.2; o risco de APP vira redução do valor do lote. Somam 200.000 + 166.000 + 1.355.000 = **1.721.000**, e não 2.041.000.
- **Faltou o Primavera 10/1199 (R$ 166.000).** O ACERTOS S1 cita uma dação à Safra em fev/2025 (arq. 1, p. 26), que não aparece no arquivo 17. Mesma situação: total de **1.887.000** se for confirmado.

### D5 Solar Serra (11.2.4-14, R$ 150 mil): **confirmado; o efeito está no crédito, não no estoque**
- O arquivo 17 (linha 60) registra o lote 02/07 do Aditivo Medianeira entregue à Solar Serra em 24/04/2024, "como forma de pagamento placas solares Hotel Araçá Ltda.".
- O lote continua no estoque de 2025, a 165.000 (GERAL linha 150), e o doc 11 ainda cobra R$ 150.000 ("01 LOTE URBANO, a escolher").
- Estoque e passivo se compensam: +165 mil e −150 mil. O ganho real para a SEI é o **crédito contra o Hotel Araçá**, porque a dívida é do hotel. Corrigindo, o patrimônio da SEI sobe cerca de 150 mil.
- A proposta da aba está correta.

### D6 Rovílio (11.2.4-1, R$ 1,65 mi): **inconsistência confirmada; cobrança dupla não provada**
- O doc 11 dá como PAGO o lote 07/04 do Morada I por R$ 432.072. O arquivo 17 diz que esse lote foi para a "Structura Corretora COMISSÃO -Rovílio Siviero (Vendido por R$ 432.072,00)" em 30/08/2021.
- É o mesmo lote e o mesmo valor, ligado a Rovílio nas duas fontes. A pergunta é se foi comissão ou pagamento de quotas.
- Lotes 010 e 013/02 (28/08/2025): conferem com o arquivo 17. A data do contrato difere: 25/08 no arquivo 17, 28/08 no doc 11.
- **Faltou:** os lotes 12/237 e 15/252 do N. S. da Saúde, dados como PAGO a Rovílio por R$ 135 mil, não aparecem no arquivo 17 e não estão no estoque de 2025 (que tem 6/260, 8/261, 1/269 e 3/269). São saídas sem registro no arquivo 17.
- Os R$ 1,65 mi "à pagar" vêm de um contrato posterior à data-base, entre sócios. Não deveriam reduzir a SEI em 31.12.2024. Isso é registro de data-base e decisão do Luís, porque mexe na relação 1/6 com os 2% de Rovílio.
- O grupo usa R$ 550 mil por lote do Morada II para pagar Rovílio, mas esse loteamento não aparece no patrimônio de 2025 nem no 11.2.1. É um ativo faltando.
- A coluna "Valor dos lotes" traz 0 com a base "valor do lote no 11.2.2", que está errada. Os valores do grupo são 500.000 + 587.510 + 432.072.

### D7 Colina do Retiro II (11.2.3, R$ 3,8 mi): **possível; manter como pergunta**
- O lote 07/A do Carnevalli foi dado à CVC em 01/08/2024 "como forma de pagamento de pavimentação no Loteamento Colina do Retiro II". Os 3,8 mi do doc 11 descrevem o que falta e não mencionam esse pagamento.
- Valor de 205.666,67 (média 2025 do Carnevalli) confere.
- Mesma leitura de D2: só há cobrança dupla se os 3,8 mi incluírem o trecho já pago.

## Descartes
- **Dalzóchio (2,2 mi, fonte de água):** correto. Os lotes encontrados são do acordo Farina (Vicentino).
- **Lago Negro (80 mil):** correto. O item cobra só o PRAD e as custas; a dação ao Candiago (14/05, 28/08/2025) é a indenização citada no próprio item.
- **Piccoli (empréstimos de Inês e Samir):** correto quanto aos lotes. Vale confirmar que esses empréstimos não têm ligação com o acerto de Arlindo Piccoli (Vale das Colinas: lote 03/1342 em 2024 e R$ 300 mil em 2018).
- **Celestino Rossi, Mercali, Caterpillar, Paraná:** nenhuma entrega de lote no arquivo 17. Correto.
- **Palavras soltas** (Antônio Prado, Rossim, Dal Más): correto.

## Erro na entrada do cruzamento
Na linha 682 do GERAL, o lote 9/0910 aparece repetido. No doc 11 (diff, linha 317), o 11.2.2 do Dalzóchio é O09/0910 a 130 mil, 003/1292 a 135 mil e **004/1292** a 135 mil. O cruzamento procurou o 9/0910 duas vezes e não procurou o 004/1292. Conferi: nenhum dos três está no arquivo 17, e o estoque de 2025 tem só o 2/1292 (APP). Não muda o resultado.

## Lotes do estoque de 2025 que já foram entregues
Estes casos inflam o estoque, o que favorece o sócio que sai. Não são cobrança dupla, mas precisam entrar no registro de data-base:
- Aditivo Medianeira 02/07, a 165 mil: entregue à Solar Serra (D5).
- Alto das Graças 11/6371, a 141 mil: dado no acerto Basso/HECTA (arq. 18, ACERTOS S4).
- Sol Nascente I 1/A e 2/A, a 351 mil cada: citados no pagamento de 2018 a Válter (arq. 18, aba "Lotes que saíram", linhas 61–64).
- Falavigna 016/876 (dação à G3 em 10/09/2025, por um termo de 2010): na data-base já era um lote devido à G3, mas não está no 11.2.4. É um passivo omitido, a favor do grupo.

## Resumo
- **Cobrança dupla confirmada:** D4, para 3 lotes de 2025 (1,72 mi), mais o Primavera 10/1199 se confirmado (1,89 mi no total).
- **Correção confirmada, mas não é cobrança dupla:** D5 (crédito de cerca de 150 mil contra o Hotel Araçá).
- **Perguntas ao grupo:** D1 (data do contrato Buriti), D2 e D7 (composição dos saldos de obra), D6 (lote 07/04 e lotes do N. S. da Saúde).
- **Rejeitado:** D3 (contingência, não cobrança dupla).
