# Impostos e custos de obra nos relatórios mensais (jul/2023 a out/2024)

> **USO INTERNO — registro para argumentação futura.**

28/10/2026. Ao aprovar o pacote 3, o Luís pediu duas coisas:
- verificar se INSS, INCC e impostos semelhantes já aparecem nos relatórios mensais;
- se aparecerem, registrar o argumento de que o grupo desconta esses custos das vendas de lotes e também do valor patrimonial (duplicidade).

**Fonte.** `01_fontes/outros/Contas_Pagas_2023-07_a_2024-10.xlsx`, aba Consolidado (md5 `a29fc2f2eb`).
- São 3.785 lançamentos: SEI 2.664, 08 919, 01 202.
- Há relatório de 15 meses; falta setembro de 2023.
- A coluna de observações vem cortada em cerca de 45 caracteres.

**Números e revisão.**
- Script: `_scripts/relatorios_mensais_tributos.py`, que gera `_extracao/relatorios-mensais-tributos.xlsx`.
- O revisor independente (`_diario/revisao-relatorios-mensais.md`) achou 22 pontos, 5 de gravidade alta. Todos foram corrigidos. A primeira versão deste documento tinha a alíquota e o total da SEI errados.

## 1. Resposta curta

| Empresa | O que os relatórios mostram | Duplicidade hoje? |
|---|---|---|
| 01 Abastecedora | Impostos sobre o próprio faturamento (aluguel): PIS, COFINS, IRPJ e CSLL, R$ 37.883,86. Faturamento de R$ 15.858/mês em 2023 e R$ 26.580/mês em 2024. Nenhum INSS de obra. Mas **foram pagos custos da regularização do posto**: ISSQN sobre a mão de obra da ampliação do posto, R$ 4.597,23 (jul/2023), e topografia da área do posto, R$ 2.832,00 (abr–mai/2024) | **Não**, porque a redução de 0,70 % não é aplicada (pacote 3). Se o grupo aplicá-la, parte do custo de regularização já foi paga e lançada como despesa (C-11) |
| 08 Transportadora | INSS só de **folha de salários**, R$ 94.032,06 (16 lançamentos). SIMPLES sobre o faturamento, R$ 138.734,16. Licença de operação e alvará, R$ 4.133,43; IBAMA, R$ 515,28; CREA, R$ 520,70. **Nenhum INSS de obra** do pavilhão, da rampa ou da borracharia | **Não**: nada foi pago e nada é descontado (pacote 3) |
| SEI | PIS, COFINS, IRPJ e CSLL sobre o faturamento mensal: **R$ 726.170,66** sobre R$ 10,59 milhões, **alíquota efetiva de 6,86 %** (60 lançamentos, 15 meses). Comissões de corretagem: R$ 398.414,36 (63 lançamentos), **3,76 % do faturamento**. A SEI pagou também impostos de outras empresas: Ceval e soja da Agropecuária, R$ 47.785,49 | **Argumento a registrar**: seção 2 |

- Nenhum lançamento cita INCC ou CUB.
- Base de setembro/2024: o PIS e a COFINS foram calculados sobre R$ 211.000 e o IRPJ e a CSLL sobre R$ 221.000. Por isso a alíquota fica entre 6,86 % e 6,87 %.

## 2. Argumento de duplicidade (SEI): registrar para depois

**Fatos**
1. A SEI recolhe PIS, COFINS, IRPJ e CSLL todo mês sobre o faturamento, em média 6,86 %. O faturamento mensal inclui vendas de lotes, parcelas de vendas anteriores e outras receitas, não só os lotes vendidos no mês. No período houve 57 lotes vendidos, com R$ 3.752.000 de valores citados.
2. O inventário 2025 da SEI, nas observações, diz:
   - "Sobre a venda – valor declarado na Escritura Pública de Compra e Venda – deve ser considerado um custo de 7,00 % […] de impostos: PIS – COFINS – IRPJ – CSLL";
   - e "um custo de 6,00 % […] de comissão" sobre o valor do negócio;
   - na transmissão dos imóveis "pertencentes aos sócios", 5,93 %, ou 7 % "se houver a aplicação do adicional de 10,00 %" do IRPJ.
3. Esses percentuais **não estão no total** da SEI. A conta é R$ 137.041.679,60 − 16.240.000 − 3.800.000 − 16.127.652 = R$ 100.874.027,60. As 16 linhas de valor dos "demais compromissos" não incluem nenhum percentual. Hoje é intenção declarada, não desconto aplicado.

**Argumentos, do mais forte ao mais fraco, para quando o grupo aplicar os percentuais**
1. **Imóveis já vendidos e ainda não escriturados.** No inventário, os 7 % aparecem logo depois do "passivo de imóveis vendidos, ou por contrato de compra e venda, ou como o compromisso de transmiti-los na forma de dação em pagamento, que ainda não foram transmitidas". Se essas vendas já entraram no faturamento mensal e foram tributadas (relatórios), cobrar 7 % de novo na escritura é tributar a mesma receita duas vezes. Ação: pedir a lista desses imóveis e a data em que cada um foi faturado.
2. **Comissões já descontadas uma a uma.** Os "demais compromissos" já deduzem R$ 1.200.000 do acerto Alexandre Susin, pago "como forma de comissão", e R$ 1.650.000 do Rovílio, que incluem R$ 550.000 de comissão (C-04). Aplicar mais 6 % genéricos sobre os mesmos lotes desconta a comissão duas vezes. Além disso, a comissão efetivamente paga no período foi de 3,76 % do faturamento, não 6 %.
3. **Na cisão, o bem não é vendido.** O bem passa ao sócio ou fica na empresa. Não há venda nem corretagem; o custo é hipotético.
4. **Valor justo.** O valor justo (CPC 46) não é reduzido por custos de transação, como a corretagem.

**Contra-argumento que o grupo vai usar.** Descontar 7 % e 6 % na avaliação e pagar na venda seria o mesmo custo visto em dois momentos, não dois custos. Avaliar pelo valor líquido de realização trata todos os sócios igualmente, e quem sai não paga de novo na venda futura. Por isso o argumento genérico ("desconta duas vezes") é fraco. Os argumentos 1 e 2 são os que se sustentam.

**Lotes vendidos em 2025.** Os relatórios terminam em out/2024, antes da data-base. Não há fonte para afirmar que o imposto das vendas de 2025 já foi pago. Se o inventário de 2025 tiver retirado lotes vendidos em 2025 e também aplicar os 7 %, o argumento vira de data-base (forma 9). Para provar, faltam os relatórios de 2025.

Registro: `parte6-candidatos.md`, **C-10** (SEI) e **C-11** (01 e 08). Não é decisão desta sessão.

## 3. Achados laterais (fora da pergunta)

Pela diretriz, passivos e ônus da 01 e da 08 não são lançados. Ativos fora do inventário são a favor.

| Empresa | Achado | Valor nos relatórios | Tratamento sugerido |
|---|---|---|---|
| 01 | ISSQN da ampliação do posto e topografia | R$ 4.597,23 + R$ 2.832,00 | Investimento no posto depois de 2017 (a favor do valor); base do C-11 |
| 01 | Títulos de capitalização Banrisul (cerca de 7 por mês) | 96 parcelas, R$ 20.386,54 | Ativo financeiro fora do inventário: registrar e pedir o extrato |
| 01 | Seguro da Ford Ranger 2010 | R$ 3.774,97 | A Ranger está no inventário da SEI: a 01 paga custo da SEI (intragrupo) |
| 01 | "Devolução cfe acordo aluguéis" à Borracharia Valverde | R$ 1.000,00 (jul–ago/2023) | O aluguel da borracharia, bem da 08, passa pela 01: confirma o C-08 |
| 01 | Última parcela do empréstimo Caixa da Rádio, paga à Siviero Transportes | R$ 6.500,00 (jul/2023) | Intragrupo, quitado antes da data-base, sem efeito |
| 08 | Empréstimo Banrisul das placas solares: 60 parcelas, a 39ª em out/2024 | 15 parcelas pagas, R$ 21.452,98 | Em 31.12.2024 restam cerca de 19 parcelas, ≈ R$ 25 mil **com juros**; o saldo devedor é menor e não foi informado. Passivo não lançado |
| 08 | Placas solares da oficina e do escritório (manutenção paga) | R$ 580,00 | Ativo fora do inventário, salvo se estiver no item 03: a favor |
| 08 | Benfeitorias no pavilhão: iluminação Enertec (notas de R$ 25.485,26, R$ 10.098,99 pagos no período), vidros, tanque e soleiras (notas de R$ 9.958,00), compressor e central telefônica (R$ 6.757,01) | cerca de R$ 42 mil em notas | Investimentos depois de 2018 no pavilhão (item 02) e nos equipamentos (item 03): a favor |
| 08 | Títulos de capitalização da conta TIS (texto cortado) | 63 parcelas, R$ 16.287,23 | Ativo fora do inventário: registrar e pedir o extrato |
| 08 | Consórcio Banrisul (parcelas de jul/2023 a abr/2024) | R$ 9.914,33 | Ativo (carta ou bem contemplado) fora do inventário: registrar |
| 08 | Garantia da "conta TIS": registro de "matrícula bem garantia" e "averbação garantia" (mai–jun/2024); "avaliação sala escritório garantia" (out/2024); seguro da sala pago pela 08 | R$ 1.374,01 | Possível garantia sobre a sala 7.393 ou outro imóvel (a ordem das datas sugere talvez o apto 401, mat. 7.400). Não é ônus lançado; só registrar. Certidões das mat. 7.393 e 7.400 |
| 08 | IPVA 2024 de veículos fora do inventário da 08 | 12 lançamentos, R$ 7.382,83 | De outras empresas: IHW 0285 (Rádio), IKI 4325 (Ceval), IHK 2435, IHW 8294, IIX 8288 e IKI 4827 (SEI), ou seja, a 08 paga custo de outras (intragrupo). **Fora de qualquer inventário:** IIH 6326 (reboque), IYI 5423 (VW Fox 2018) e OYB 0G90 (Toyota 2014), possíveis bens a favor: pedir os CRLV |
| 08 | TAC por infração ambiental na Promotoria de Veranópolis (área cortada no texto) | nota de R$ 2.502, quitada em nov/2023 | Já pago; sem ônus |
| 08 | Água de "oficina/posto/lancheria" (CORSAN) paga pela 08 | R$ 4.536,95 | Custo da 01 pago pela 08 (intragrupo; ligado ao C-08) |
| SEI | Impostos da Ceval e da Agropecuária (soja) pagos pela SEI | R$ 47.785,49 | Intragrupo, para a sessão SEI |

## 4. Proposta

- Aprovar o registro do **C-10** (SEI; argumentos 1 e 2) e do **C-11** (01: custo de regularização já pago; 08: nada pago).
- Registrar os achados laterais na aba "Créditos e contingências" da planilha-mestre, **sem valor**, com o tratamento acima:
  - ativos fora do inventário a favor: capitalização, consórcio, placas solares, veículos sem inventário, benfeitorias;
  - passivo e ônus só como registro: empréstimo das placas, garantia;
  - fluxos intragrupo como registro para a Parte 6.
- Avisar a sessão SEI do C-10 e dos pagamentos intragrupo depois da sua aprovação.
