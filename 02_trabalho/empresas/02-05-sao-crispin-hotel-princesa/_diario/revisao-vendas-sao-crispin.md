# Revisão independente: levantamento das vendas da São Crispin

Revisor com contexto novo, 28/10/2026. Objeto: `levantamento-vendas-sao-crispin.md` e `_scripts/vendas_sao_crispin.py`. Conferi tudo diretamente nas fontes: inventário 2018 doc 17; docx 2025 das áreas vendidas; `Contas_Pagas_…xlsx` (Consolidado), com as páginas dos PDFs 2024-04 p.4, 2024-05 p.1 e 2024-09 p.4 lidas na imagem; docs 22, 25 e 27 de 2018; manuscrito de 14/01/2019; arq. 1 (Dr. Pablo) p.5–7; contratos 7, 8 e 9 (OCR); inventários 2025 (02, 04, 05, 08, 11); planilhas Google; e a pasta `analises/acertos-e-areas-vendidas`. Não alterei nada além deste arquivo.

Resumo: os números do quadro de vendas estão certos. O rascunho, porém, tem um erro de soma (imóvel) e atribui a "Agropecuária/Hotel" duas vendas de soja tributadas como de pessoa física. Também soma ao "dinheiro das vendas" soja que é aluguel (Lorenzet) ou de origem desconhecida. A seção "retirado" não usa as fontes de retiradas que existem. Algumas conclusões vão além do que as fontes mostram.

## Achados

### A1 (alta): as vendas de 1.300 e 2.606 sacas foram tributadas como de pessoa física, não da Agropecuária ou do Hotel

- **Evidência.** As duas linhas de 04/2024 e 05/2024 estão no fornecedor "FUNRURAL - IMPOSTO 1,50% PRODUTOR RUR[AL]". Conferi as linhas na imagem dos PDFs (2024-04 p.4: R$ 2.398,50; 2024-05 p.1: R$ 4.808,07). Os 1,5 % são a alíquota do produtor pessoa física (1,2 + 0,1 + 0,2).
- As vendas da Agropecuária aparecem com outro fornecedor ("FUNRURAL-AGROP FAZENDA SAO CRISPIN LT") e outra alíquota: 3.895,57 ÷ 190.028,75 = 2,05 %, e 1.765,48 ÷ 86.121,25 = 2,05 %. Essa é a alíquota de pessoa jurídica. Essas vendas também têm IRPJ (1,20 %) e CSLL (1,08 %).
- Sobre R$ 159.900 e R$ 320.538 não há nenhum IRPJ ou CSLL no período (busquei todas as linhas de IRPJ e CSLL).
- O próprio grupo diz que "tributando a venda dos grãos na inscrição estadual pelo CNPJ da Agropecuária" (arq. 1, p.6). Isso não bate com essas duas linhas.
- **Consequência.** A frase "A SEI pagou os impostos das vendas de soja da Agropecuária/Hotel" não tem base para essas duas vendas. Se as 2.606 sacas da Rigo foram vendidas em nome de uma pessoa física (talvez um sócio com bloco de produtor), o dinheiro pode ter passado por essa pessoa. Isso interessa justamente à pergunta "quanto foi retirado".
- **Correção.** Separar as duas regras de tributação na tabela. Dizer que a venda das 2.606 sacas foi tributada como de pessoa física e que não se sabe em nome de quem foi a nota. Incluir a pergunta: "em nome de quem saiu a nota das 2.606 sacas e para qual conta foi o dinheiro?".

### A2 (alta): o total de R$ 756.588 entra na seção "pago com esse dinheiro", mas só R$ 320.538 vêm da venda de terras

- **Evidência.**
  - Dos R$ 756.588:
    - R$ 276.150 (190.028,75 + 86.121,25) são soja da Lorenzet, ou seja, **aluguel** da fazenda, não preço de venda de terra.
    - R$ 159.900 (1.300 sacas, "…COLONIA SA", cortado no próprio PDF) não têm origem identificada. "COLONIA SA…" pode ser outra área. Por exemplo, a Colônia de São Brás, em Cotiporã, citada no doc 22 e que está no inventário 2025 da Transportadora (08, itens 07/08). É hipótese: o texto está cortado.
  - Só as 2.606 sacas (R$ 320.538, pelo histórico "20.150 (-) 17.544") se ligam a uma parcela da venda (Rigo).
- **Consequência.** Um leitor leigo vai entender que entraram R$ 756 mil da venda das terras em 2024.
- **Correção.** Mudar o título para algo como "Soja vendida em 2024 (só parte vem da venda das terras)". Classificar cada linha como "parcela Rigo", "aluguel Lorenzet" ou "origem não identificada". Dar só R$ 320.538 como valor ligado à venda das terras.

### A3 (média): erro de soma no imóvel da Rua Flores da Cunha

- **Evidência.** No script, `pick(r"DAL PONTE", 3)` também pega "RICARDO AFONSO DAL PONTE - CORRETOR" (10/2024, R$ 2.560, comissão de lotes do Solar Norte), que não tem relação com o imóvel. As parcelas de Maria Tereza Filippin Dal Ponte somam **R$ 303.043,76**, não R$ 305.603,76.
- **Correção.** Filtrar por "MARIA TEREZA" ou por "AQUISICAO IMOVEL RUA FL" e corrigir o valor no texto.

### A4 (média): usos da soja apresentados como "documentados" sem essa força

- **Escavadeira.**
  - O vínculo com as 2.606 sacas está escrito no docx de 2025, então está ok.
  - O problema é pôr ao lado "pagos no período R$ 573.209,46 (entrada, duas parcelas e o financiamento)", o que sugere que a soja pagou isso. Ela não pagou.
  - As 2.606 sacas renderam R$ 320.538 brutos, cerca de R$ 315.730 depois do FUNRURAL. Isso cobre no máximo uns 36 % dos R$ 887.000 e uns 62 % da entrada de R$ 510.000.
  - "E o financiamento" está errado: o período tem só 6 das 48 parcelas (R$ 63.209,46) de um financiamento de R$ 377.000. Ainda restavam R$ 347.652,03 em 31/07/2025 (inventário 11).
- **Imóvel da Flores da Cunha.**
  - "CHEQUES SOJA TITTON" quer dizer cheques da Titton, que compra soja. Não há nada ligando essa soja à São Crispin, muito menos à venda das terras.
  - O próprio rascunho diz "origem não identificada" e mesmo assim põe o item em "Usos documentados do dinheiro da soja".
- **Correção.** Escavadeira: "o grupo diz que 2.606 sacas (≈ R$ 320 mil) foram usadas na compra; o resto da máquina foi pago com outros recursos da SEI". Imóvel: tirar da lista de usos e pôr em "a confirmar", com a pergunta "de quem era a soja vendida à Titton?".

### A5 (média): "a soja de 2021 estava com a empresa" é conclusão além das fontes

- **Evidência.**
  - O histórico "20.150 (-) 17.544 = 2.606" mostra só que o grupo dispôs de 2.606 sacas daquele lote em 2024. O vendedor pode mandar o comprador (Rigo) entregar direto a quem compra, sem ter a soja em depósito.
  - O docx diz "(FALTA RECEBER)" nos lotes de 2021 e 2022.
  - Só as linhas da Lorenzet têm a marca "HOTEL/OLEOPLAN", que indica depósito. As da Rigo não têm.
  - Há também uma coincidência que o rascunho omite: **2.606 sacas é também a safra Lorenzet de 2021** no mesmo docx. A análise anterior já tinha anotado isso (`analises/acertos-e-areas-vendidas/diario.md`, MÉDIA 5).
- **Correção.** Escrever: "pelo menos 2.606 sacas do lote de 2021 foram vendidas pelo grupo em 2024. Não se sabe se o restante foi entregue e está depositado, ou se ainda é devido pela Rigo". Citar a coincidência com a Lorenzet 2021 e manter a pergunta.

### A6 (média): "recebido no máximo R$ 8.834.000" não fecha com o próprio rascunho e trata soja como dinheiro

- **Evidência.**
  - O teto supõe que as parcelas de 2021 e 2022 não entraram. Mas a seção 4 mostra 2.606 sacas de 2021 já vendidas.
  - Além disso, R$ 80 por saca é preço de referência do contrato (log de decisões: "só como referência de ordem de grandeza, nunca como avaliação"). Não é dinheiro recebido.
  - As safras de 2019 e 2020 entraram em sacas e foram vendidas sabe-se lá a que preço. As OBSERVAÇÕES ("O SOJA DE 02 SAFRAS … AINDA NÃO FORAM VENDIDOS") indicam que as outras duas foram vendidas.
- **Correção.** Mostrar o recebido em duas partes:
  - Dinheiro: R$ 4.610.000 à vista em 2018 + R$ 1.000.000 da Pellegrini em jun/2019 = R$ 5.610.000, todos sem comprovante.
  - Soja da Rigo: 40.300 sacas (2019–2020), presumidas entregues e vendidas, mais pelo menos 2.606 sacas de 2021. Lote de 2022 e saldo de 2021 em aberto: 37.694 sacas.

  Dar o valor em R$ só como ordem de grandeza.

### A7 (média): a seção "retirado" ignora fontes que existem

- **Contas pagas da SEI.** Têm retiradas de sócios em todo o período. Pela minha classificação por texto (aproximada), de jul/2023 a out/2024:

  | Tipo de retirada | Valor |
  |---|---|
  | Pró-labore (Válter, Pedro e Adélia, R$ 6.500 por mês cada) | R$ 558.000 |
  | Divisão de aluguéis (R$ 3.500 por mês por sócio) | R$ 315.000 |
  | "Divisão de lucros empresas ano 2023" (12/2023) | R$ 80.000 |
  | Devolução de empréstimos a sócios (Adélia R$ 100.000; Pedro R$ 325.000, dos quais R$ 149.000 em 04/2024, mês das vendas de soja) | R$ 425.000 |
  | "Hotel Araçá – Raul Siviero + pagto contas individuais" | R$ 186.000 |

  Total ≈ R$ 1,56 milhão. Nenhuma linha cita a fazenda. O grupo, porém, diz que as receitas da fazenda são "depositadas e saldadas conforme fluxo financeiro necessário para a empresa Siviero Empreendimentos Imobiliários" (arq. 1, p.6). Ou seja, o dinheiro da fazenda se mistura no caixa da SEI, de onde saem essas retiradas.
- **Manuscrito de 14/01/2019.** Não trata só de moradia e estudo. Diz: "Irineu/Raul – têm dinheiro a juros na empresa = 1,3 % a 1,5 % … Devolver capital – empresa para pagar juros". O doc 22, da mesma reunião, fala em "Juros para os Sócios … Não tem mais como pagar estes juros", "Pró labore", "Divisão de lucros" e na dívida da SEI de cerca de R$ 10 milhões. Isso é contemporâneo à entrada dos R$ 4,61 milhões. É um destino possível (pagar sócios credores ou a dívida da SEI), mas sem prova.
- **Datas.** "Anteriores às vendas" está errado: o doc 25 é "Em 31.12.2018" e a reunião é de 14/01/2019, ambos **depois** das vendas de 2018. O fato relevante é o contrário: mesmo escritos depois das vendas, nenhum dos dois menciona o dinheiro delas. O doc 25 também traz o pró-labore anual de cada sócio de 2012 a 2018, não só moradia e estudo.
- **Citação do Dr. Pablo.** A frase "ainda estamos recebendo…" está na **pergunta** do Dr. Pablo (p.5), não numa resposta do grupo. O grupo não respondeu esse ponto (p.6). Ela não fala de sócios: é indício de recebimento (seção 2), não de retirada. A conclusão "se parte das parcelas é paga diretamente aos sócios…" não tem base nessa frase.
- **Correção.** Reescrever a seção 3 com três pontos:
  1. Nenhuma fonte liga o dinheiro das vendas a retiradas.
  2. Retiradas documentadas no único extrato disponível (SEI, 2023–24), com os totais acima e a ressalva de que saem do caixa comum.
  3. Indícios de 2018/2019 (sócios credores a juros, dívida da SEI).

  Mover a frase do Dr. Pablo para a seção 2, como "afirmação na pergunta do Dr. Pablo, não confirmada pelo grupo".

### A8 (média): falta dizer o que o inventário de 2018 mostra sobre o destino do dinheiro

- **Evidência.** No doc 17, o "TOTAL GERAL" (R$ 24.112.390) ainda inclui as áreas vendidas pelo valor de inventário, e dele se abatem os R$ 12.058.000. Nenhum item da empresa registra dinheiro em caixa, aplicação ou parcelas a receber (Pellegrini 2019, soja da Rigo). Os docs 22 e 27 tratam a Agropecuária como "sem nenhuma dívida", com receita de "arrendamentos de terras e venda de gado", e também não registram esse dinheiro.
- **Consequência.** Já em 31/12/2018 o dinheiro (e o crédito) das vendas não aparecia como bem da Agropecuária nem do Hotel. Esse é o dado central para "para onde foi".
- **Correção.** Incluir esse parágrafo na seção 2 ou 5.

### A9 (média): o total mistura venda fora da São Crispin e gado

- **Evidência.** A venda Angeli (R$ 510.000) é do lote rural 24/24-E em **Cotiporã**, não da São Crispin. Dos R$ 2.100.000 da Pellegrini, R$ 100.000 são 41 cabeças de gado: a terra é R$ 2.000.000 (85,32 ha × R$ 23.441,16).
- A leitura de que o gado foi vendido junto (e não recebido como pagamento) é a coerente: R$ 2.000.000 de terra + R$ 100.000 de gado = R$ 2.100.000 = R$ 1.100.000 à vista + R$ 1.000.000 em jun/2019. O rascunho trata isso certo. Mas o título "Vendas de terras da Fazenda São Crispin" e o "Total das três vendas" não avisam o leitor.
- **Correção.** Mostrar o total de R$ 12.058.000 dividido em três partes:
  - terra da São Crispin: R$ 11.448.000 (Rigo R$ 9.448.000 + Pellegrini R$ 2.000.000);
  - gado: R$ 100.000;
  - terra em Cotiporã (Angeli): R$ 510.000.

  Dizer também que a venda da Rigo foi fechada em soja: 118.100 sacas, das quais o à vista equivale a 37.500 (docx 2025).

### A10 (média): clareza para o Luís

- **Falta resposta direta.** Não há um resumo inicial que responda às três perguntas em três frases.
- **Jargão e siglas internas.** O texto usa "Parte 5, decisão do pacote 1", "candidato P6-07", "Q6", "base = imposto ÷ 1,5 % (FUNRURAL)", IRPJ/CSLL sem explicação e históricos cortados ("X ALÍQ", "COLONIA SA").
- **Correção.**
  - Abrir com um quadro "Recebido / Retirado / Pago": o que se sabe, o que não se sabe e de onde vem.
  - Trocar os códigos internos por frases.
  - Explicar uma vez: "o valor da venda foi calculado a partir do imposto pago, que é 1,5 % do valor vendido".
  - Pôr a tabela de impostos num anexo.

### A11 (baixa): detalhes

- A tabela de despesas da fazenda está certa. O ITR "CAMPO HOTEL + PROPRIE…" (R$ 1.856) foi bem excluído, porque parece ser a Sede Campestre. Vale anotar que o ITR 2024 da São Crispin declara **259 ha**, contra 411 ha da fazenda (dado para a sessão da empresa).
- O preço de R$ 123,00 para as 2.606 sacas é **dedução**: o histórico está cortado no PDF depois de "VENDIDOS X". A conta fecha exata (2.606 × 123 × 1,5 % = 4.808,07). Dizer que é dedução.
- A data provável das vendas é março/abril de 2024: o FUNRURAL vence no mês seguinte à venda e foi pago em 26/04 e 06/05. Isso combina com as parcelas da escavadeira de 09/04 e 09/05.
- A Pellegrini deixa de ter "sem registro posterior" só porque nenhuma fonte cita o pagamento de jun/2019. Pode valer uma pergunta específica sobre a quitação.

## O que está correto (conferido)

- **Quadro de vendas.** Rigo R$ 3.000.000 à vista + 4 × 20.150 sacas × R$ 80 = R$ 9.448.000. Pellegrini R$ 1.100.000 + R$ 1.000.000 = R$ 2.100.000, com o gado tratado corretamente como parte do preço. Angeli R$ 510.000. Total R$ 12.058.000, igual ao "TOTAL RECEBIDO" dos docs de 2018 e de 2025.
- **Prazos.** À vista em 2018: R$ 4.610.000. Depois de 2018: R$ 7.448.000. "FALTA RECEBER" em 2021 e 2022: R$ 3.224.000 à referência de R$ 80.
- **Saldo de soja da Rigo.** 37.694 sacas = 40.300 − 2.606 = 17.544 + 20.150.
- **Transcrição das contas pagas.** Conferi na imagem: FUNRURAL 04/2024 (R$ 2.398,50, "1.300 SC … X R$ 123,00") e 05/2024 (R$ 4.808,07, "20.150 (-) 17.544 = 2.606 SC SOJA VENDIDOS X"), e ITR 09/2024 (R$ 3.090). As contas batem: 1.300 × 123 × 1,5 % = 2.398,50; IRPJ 1,2 % e CSLL 1,08 % sobre R$ 190.028,75 e R$ 86.121,25 batem até o centavo.
- **Escavadeira.** Pago R$ 573.209,46 = 10.000 + 2 × 250.000 + 6 × 10.534,91. Nota de R$ 887.000. Está no inventário 2025 da SEI (item 01) e no passivo Caterpillar. Não há crédito do Hotel ou da Agropecuária contra a SEI em nenhum inventário.
- **Cheques soja Titton.** R$ 78.750,74 (70.000 + 8.750,74).
- **Contradição Lorenzet.** O docx diz "não vendida", mas há imposto sobre "SOJA LORENZET" (R$ 190.028,75). Bem apontada.
- **Lacunas.** Contas pagas só de jul/2023 a out/2024, sem relatório de entradas nem extratos do Hotel ou da Agropecuária. A pergunta sobre o destino dos R$ 12.058.000 continua necessária.
