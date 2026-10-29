# Revisão independente — levantamento dos relatórios mensais (impostos e custos de obra)

28/10/2026. Revisor independente, sem contexto anterior.

**O que foi revisado:** `pacotes/levantamento-relatorios-mensais.md`, `_scripts/relatorios_mensais_tributos.py` (rodado com `python -I`; a saída bate com o documento) e as linhas C-10 e C-11 de `parte6-candidatos.md`.

**Fontes:** `Contas_Pagas_2023-07_a_2024-10.xlsx` (Consolidado, md5 `a29fc2f2eb`), `Contas_Pagas_e_Lotes_Vendidos_…xlsx`, `Contas_Pagas_Recorrentes_x_Extraordinarias_…xlsx`, inventários 2025 (00, 01, 08 e 11) e a planilha-mestre, aba "Créditos e contingências". Todos os números abaixo foram recalculados a partir das fontes.

**Atenção:** a coluna Observações da fonte vem cortada em cerca de 45 caracteres. Algumas leituras abaixo completam a palavra final (por exemplo, "AMPLIACAO PO…" lido como "posto"). Nesses casos, isso está indicado.

---

## Achados

### Gravidade alta

**A1. A alíquota efetiva de 6,71 % e os "12 meses" estão errados. O script perdeu 3 meses e contou só metade dos impostos em outros 2.**
- *Evidência:* a regex `FATURAMENTO MES (\w+/\d{4}) = R\$ ([\d.]+,\d{2})` falha quando:
  - o valor está cortado: out/2023 "R$ 1.256.000,0"; nov/2023 "R$ 1.843.500,"; fev/2024 "R$ 604.833,3";
  - o mês vem abreviado: IRPJ e CSLL de nov/2023 trazem "NOV/2023 - R$";
  - há letra O no lugar de zero: PIS e COFINS de ago/2024 trazem "297.145,O0";
  - falta espaço: IRPJ e CSLL de set/2024 trazem "SETEMBRO/2024=R$ 221.000,00".

  Por isso, ago/2024 (4,38 %) e set/2024 (3,65 %) entram com faturamento cheio e só 2 dos 4 tributos. A observação cita o faturamento nos **15 meses**, não em 12. Pelos 60 lançamentos "FATURAMENTO MES" (15 meses × PIS, COFINS, IRPJ e CSLL), recalculado:

  | | Valor |
  |---|---|
  | Faturamento | R$ 10.577.794,97 |
  | Impostos | R$ 726.170,66 |
  | **Alíquota efetiva** | **6,87 %** (PIS + COFINS 3,65 %; IRPJ + CSLL 3,22 %) |

  Por mês: ago/2024 = 8,03 %; set/2024 = 6,56 %; out/2023 = 6,70 %; nov/2023 = 6,71 %; fev/2024 = 6,67 %.
- *Correção:* trocar "6,71 % / 12 meses / R$ 460.972,43 sobre R$ 6.873.461,64" por "6,87 % / 15 meses / R$ 726.170,66 sobre R$ 10.577.794,97". Isso vale para a seção 1, para o fato 1 da seção 2 e para o C-10. No script:
  - agrupar por mês com `FATURAMENTO MES (\w+)/(\d{4})`, mapeando NOV para NOVEMBRO;
  - ler o valor com `R\$\s*([\d.]+,[\dO]{0,2})` e trocar O por 0;
  - conferir o faturamento pelo PIS (PIS ÷ 0,0065).

  Registrar também a divergência de set/2024: PIS e COFINS sobre R$ 211.000,00; IRPJ e CSLL sobre R$ 221.000,00.

**A2. Os R$ 760.242,24 não são só "impostos sobre o faturamento das vendas" da SEI.**
- *Evidência:* dos 72 lançamentos da categoria, 60 são "FATURAMENTO MES" (R$ 726.170,66). Os outros 12 (R$ 34.071,58) são:
  - impostos da **Ceval** pagos pela SEI ("CEVAL-IMPOSTO … NOTA"): R$ 16.201,70;
  - IRPJ sobre venda de **soja** da Agropecuária: R$ 5.366,12;
  - IRPJ sobre **notas de serviço** ("IMPOSTO S/SERVICOS-N.F. 08/09/12-PREFEITURA"): R$ 12.503,76.

  A CSLL desses mesmos fatos ("CONTRIBUICAO SOCIAL", "C.S.L.L") ficou fora porque a regex só aceita "CSLL". O total mistura empresas e fatos e ainda é incoerente por dentro.
- *Correção:* usar R$ 726.170,66 (60 lançamentos) como imposto sobre o faturamento da SEI. Os outros 12 podem ser citados à parte como impostos de outras empresas pagos pela SEI. Isso é um achado intragrupo, a avisar à sessão SEI.

**A3. 01: não é verdade que há "só impostos sobre o faturamento". Há ISSQN sobre a mão de obra da ampliação do posto.**
- *Evidência:* em jul/2023, PREFEITURA MUNICIPAL DE VERANÓPOLIS, R$ 4.597,23, "ISSQN-IMPOSTO SOBRE MAO DE OBRA AMPLIACAO PO…". O texto está cortado, mas a 01 só tem o posto. É imposto de obra, da mesma família do INSS de obra; o ISS da obra costuma ser exigido para o habite-se e a averbação. O script não tem categoria para ISS e deixou o lançamento passar. Junto vem o achado M9, a topografia da área do posto.

  O que fica errado no texto:
  - na tabela da seção 1, "As reduções e o INSS de regularização da 01 … não foram pagos" vale para o INSS, mas não para todo o custo de regularização;
  - no C-11, falta mencionar esse lançamento.
- *Correção:* incluir o ISSQN de obra (e a topografia) na linha da 01 e no C-11. Exemplo: "nenhum INSS de obra; houve ISSQN sobre a mão de obra da ampliação do posto, R$ 4.597,23 em jul/2023, e topografia da área do posto, R$ 2.832,00 em abr–mai/2024". Acrescentar ao C-11: se o grupo aplicar a redução de 0,70 % (R$ 67.629,90), parte do custo de regularização do posto já foi paga e lançada como despesa. É a mesma lógica do C-10, agora com fato documentado. A conclusão "não há duplicidade hoje" continua válida, porque a redução não está no total (ver Q5).
- *Ampliação depois de 2017:* a ampliação também é investimento no posto depois de ago/2017, a data do CUB usado na avaliação. Vale registrar como achado a favor.

**A4. 01: a CSLL ficou fora. O total de impostos sobre o faturamento é R$ 37.883,86, não R$ 28.254,10.**
- *Evidência:* 15 lançamentos "CONT. SOCIAL - ABAST. COMB. SIVIERO", R$ 9.629,76: R$ 456,71 por mês em 2023 e R$ 765,50 por mês em 2024 (2,88 % do faturamento). A regex `CSLL` não reconhece "CONT. SOCIAL". Os 45 lançamentos do documento são 15 × (PIS, COFINS, IRPJ).
- *Correção:* "PIS, COFINS, IRPJ e CSLL, R$ 37.883,86 (60 lançamentos)". Na regex, acrescentar `CONT\.? SOCIAL|CONTRIBUICAO SOCIAL|C\.S\.L\.L`.

**A5. O C-10 e o 3º ponto da seção 2 afirmam um fato sem fonte.**
- *Evidência:* o texto diz que, "nos lotes vendidos entre a data-base e a avaliação (de jan. a out./2025, por exemplo), o imposto já foi pago e lançado como despesa". Os relatórios vão de jul/2023 a out/2024, ou seja, terminam **antes** da data-base. Nenhuma fonte revisada mostra imposto pago em 2025, nem diz onde ele teria sido "lançado como despesa" no cálculo da cisão. Além disso, para uma avaliação em 31.12.2024, um imposto pago em 2025 não reduz nada nessa data, a menos que o inventário misture datas. Isso seria outro argumento (forma 9, data-base), não o desta linha.
- *Correção:* tirar a afirmação ou reescrevê-la como hipótese condicional: "se o inventário de 2025 tiver retirado lotes vendidos em 2025, ou usado o caixa dessas vendas, e também descontar 7 % …". Indicar que falta a fonte, por exemplo os relatórios de 2025.

### Gravidade média

**M1. O argumento central do C-10 ("descontado duas vezes: na avaliação e de novo quando o lote for vendido") é frágil.**
- *Contra-argumento previsível do grupo:* o desconto na avaliação e o pagamento na venda são o **mesmo** custo visto em dois momentos, não dois custos. Quem sai recebe a cota em 31.12.2024. O imposto da venda futura cai sobre a empresa e sobre os sócios que ficam, não de novo sobre quem sai. Avaliar pelo valor líquido de realização (preço menos tributos e corretagem) seria só tratar todos os sócios por igual. Do jeito que está escrito, o grupo derruba o argumento com facilidade.
- *Argumentos mais sólidos, a considerar:*
  1. **Contexto do próprio inventário.** Os 7 % aparecem logo depois do parágrafo "a empresa ainda tem um passivo de imóveis vendidos … que ainda não foram transmitidas, incidindo neste caso … os impostos … sobre o valor declarado na Escritura". Os relatórios mostram que a SEI recolhe PIS, COFINS, IRPJ e CSLL todo mês sobre o faturamento. Se esses imóveis já vendidos já entraram no faturamento tributado, cobrar 7 % de novo na escritura é duplicidade de verdade, e documentada. Pedir a lista desses imóveis e quando foram faturados.
  2. **Comissões específicas já deduzidas.** Os "demais compromissos" já tiram R$ 1.200.000 de comissão de Alexandre Susin (Buriti) e os R$ 1.650.000 do Rovílio, que incluem R$ 550.000 de comissão (C-04). Aplicar 6 % genérico aos mesmos lotes desconta a comissão duas vezes.
  3. **Bem que não é vendido.** Na cisão, o bem passa ao sócio ou fica na empresa sem venda. Não há fato gerador de PIS e COFINS de venda nem corretagem; é custo hipotético.
  4. **Valor justo.** O valor justo (CPC 46, item 25) não é reduzido por custos de transação, como a corretagem.
- *Ressalva:* tirar a afirmação "pagos com o dinheiro da própria venda, como mostram os relatórios". Os relatórios mostram o pagamento, não de onde veio o dinheiro.

**M2. O texto cita a comissão "como mostram os relatórios", mas não dá o valor.**
- *Evidência:* a planilha Recorrentes × Extraordinárias, categoria "Comissões de corretagem (variável com vendas)", tem 70 lançamentos, R$ 407.598,98. Uma busca por COMISS, CORRETAG e INTERMEDIA no Consolidado da SEI dá 63 lançamentos, R$ 398.414,36. Isso é cerca de 3,9 % do faturamento de R$ 10,58 milhões, bem abaixo dos 6 %, embora a época dos pagamentos não coincida com a do faturamento.
- *Correção:* incluir o número. Ele reforça que a comissão real é menor que os 6 % do grupo.

**M3. "Os impostos de cada venda de lote" e a comparação entre 6,87 % e 7 % precisam de ressalva sobre a base.**
- *Evidência:* no período houve 57 lotes vendidos, com valor citado de R$ 3.752.000. O faturamento tributado foi de R$ 10,58 milhões. O faturamento mensal inclui outras receitas (parcelas de vendas anteriores, áreas e outras), não só os lotes vendidos no mês. Além disso, os 7 % do inventário incidem "sobre o valor declarado na Escritura", e a comissão de 6 % sobre "o valor do negócio". São bases diferentes do faturamento.
- *Correção:* trocar "de cada venda de lote" por "sobre o faturamento mensal da SEI". Dizer que as alíquotas são comparáveis em ordem de grandeza, não na base. A explicação da diferença entre 5,93 % e 7 %, dada pelo próprio inventário ("adicional de 10,00 %", do IRPJ), também deve constar.

**M4. Empréstimo das placas solares: os "≈ R$ 25 mil" são a soma das parcelas que faltam, com juros, e não o saldo devedor.**
- *Evidência:*
  - as parcelas caem de P.24, R$ 1.619,86 (jul/2023), para P.39, R$ 1.314,54 (out/2024), cerca de R$ 20 por parcela;
  - a contagem confere: 15 lançamentos (falta a P.26, de set/2023, mês sem relatório), soma de R$ 21.452,98; com P.40 e P.41 em nov e dez, restam 19 parcelas, P.42 a P.60;
  - 19 × R$ 1.314,54 = R$ 24.976,26 é o que ainda será pago, não a dívida em 31.12.2024;
  - com parcelas decrescentes (perfil SAC), o saldo de principal fica em torno de R$ 16–17 mil. É uma estimativa grosseira, porque a parcela oscila (P.31 > P.30).
- *Correção:* "parcelas restantes ≈ R$ 25 mil (com juros); saldo devedor menor, não informado". O efeito é pequeno, porque o passivo não é lançado.

**M5. Iluminação do pavilhão: o investimento é cerca de R$ 25,5 mil, não "≈ R$ 10 mil". Há outras benfeitorias e equipamentos na 08 que ficaram de fora.**
- *Evidência:*
  - no período foram pagos R$ 10.098,99 (Enertec), mas as notas somam R$ 25.485,26: NF de R$ 8.041,46 em 4 parcelas, NF de R$ 14.206,94 em 5 parcelas, mais R$ 1.270,23, R$ 1.800,00 e R$ 166,63. A P.1/4 e as P.1 a P.4/5 foram pagas antes de jul/2023;
  - ainda no pavilhão e no escritório: vidros para as janelas da oficina (NF de R$ 3.460,00, abr/2024); tanque e soleiras de portas e janelas da oficina (R$ 5.400,00, set–out/2024); compressor de ar e arco de serra (R$ 2.157,01); central telefônica do escritório (R$ 4.600,00).
- *Correção:* registrar o investimento pelo valor das notas e acrescentar esses itens. Todos são a favor do valor do pavilhão (item 02) e dos equipamentos (item 03).

**M6. Ativos financeiros da 01 e da 08 que não estão nos inventários: títulos de capitalização e consórcio.**
- *Evidência:*
  - **títulos de capitalização** do Banrisul: na 01, 96 lançamentos "PARCELA CAPITALIZACAO", R$ 20.386,54, cerca de 7 títulos por mês; na 08, 63 lançamentos, R$ 15.577,18;
  - **consórcio** da 08: 9 lançamentos "TIS - PAGTO PARCELA CONSORCIO BANRISUL" (jul/2023 a abr/2024), R$ 9.914,33. O fim em abr/2024 sugere contemplação ou quitação;
  - nenhum dos inventários cita capitalização ou consórcio.
- *Correção:* registrar como ativos a favor, sem valor (saldo resgatável e carta de crédito ou bem adquirido), e pedir os extratos. Pela diretriz, estão no lado da versão de maior valor.

**M7. Veículos com IPVA ou seguro pagos pela 01 ou pela 08 que não estão nos inventários.**
- *Evidência:*
  - **08 — IPVA 2024 pago (inventário 2025 da 08 só lista IKR 2054, IJZ 5399, IJZ 8207 e IQB 0612):** IIH 6326 (reboque artesanal), IKI 4827 e IKI 4325 (retroescavadeiras Case), IIX 8288 (Scania 1963), IHK 2435 (Scania 1967), IHW 0285 (Toyota Bandeirante), IHW 8294 (MB 1113), **IYI 5423 (VW Fox 2018, R$ 1.708,45)** e **OYB 0G90 (Toyota 2014, R$ 4.654,33)**;
  - **01:** seguro de uma Ford Ranger 2010, R$ 3.774,96 (mar–jun/2024). O inventário da 01 não tem veículo.
- *Correção:* registrar como possíveis bens fora do inventário, a favor, e pedir os CRLV. A 08 pode estar pagando IPVA de veículos de outras empresas, o que seria um achado intragrupo; por isso, sem valor por enquanto.

**M8. Receita da 01: os relatórios dão o faturamento mensal que a planilha-mestre registra como "não informado".**
- *Evidência:* o faturamento da 01 (base de PIS, COFINS, IRPJ e CSLL) foi de R$ 15.858,00 por mês de jun a dez/2023 e de **R$ 26.580,00 por mês** de jan a set/2024. A planilha-mestre, linha "Aluguel da locação à SIM", diz "valor em 31.12.2024 não informado" e usa R$ 15.000 (2022). Há também "BORRACHARIA VALVERDE — DEVOLUCAO CFE ACORDO ALUGUEIS", R$ 500 por mês (jul–ago/2023), pago pela 01. Ou seja, o aluguel da borracharia, que é bem da 08 (item 05), passa pela 01. Isso documenta o C-08.
- *Correção:* acrescentar os dois à tabela de achados, ligando o segundo ao C-08.

**M9. 01: topografia da área do posto.**
- *Evidência:* TOPO AMBIENTAL, "SERVICOS TOPOGRAFIA AREA DO POSTO E…", 2 × R$ 1.416,00 (abr–mai/2024). A própria planilha Recorrentes × Extraordinárias classifica como "Projetos, licenças e regularização". É um custo típico de averbação ou regularização, já pago.
- *Correção:* juntar ao A3.

### Gravidade baixa

**B1. A rubrica da 08 "Licença de operação da oficina e alvará, R$ 5.605,33" mistura outras coisas.**
- *Evidência:* a soma é:
  - Boff, licença de operação: R$ 3.930,00;
  - alvará: R$ 203,43;
  - anuidade do CREA do engenheiro: R$ 520,70;
  - **Panazzolo, "avaliação sala escritório garantia conta TIS": R$ 951,20.** A regex pegou "PROJETOS" no nome do fornecedor.

  A taxa trimestral do IBAMA ligada à LO da oficina (R$ 515,28) ficou de fora.
- *Correção:* "licença de operação e alvará, R$ 4.133,43; anuidade do CREA, R$ 520,70; avaliação da garantia, R$ 951,20 (achado da garantia); taxa do IBAMA, R$ 515,28". As categorias da SEI (ITBI e registro, projeto e engenharia, IPTU) foram feitas só por palavra-chave, sem revisão. Apresentá-las como aproximadas.

**B2. Citações do inventário da SEI.**
- *Evidência:* foram suprimidos, sem reticências, os trechos "(Sete por cento)" e "(Seis por cento)". Não é dito que 7 % = 5,93 % + o adicional de 10 % do IRPJ. Os itens das observações vão de "1" a "24", e só 16 têm valor.
- *Correção:* marcar as supressões com […]. Isso é forma.

**B3. A 08 paga custos de outras empresas.**
- *Evidência:*
  - água "OFICINA/POSTO/LANCHERIA" (CORSAN, R$ 4.536,95), que inclui custos do posto e da lancheria, da 01;
  - "areia — obras Lot. Falavigna", R$ 850, que é da SEI;
  - "matrículas da área da SEI em Sapopema", R$ 47;
  - IPVA de veículos de Giovani como "vale".
- *Correção:* registrar como dado intragrupo para o C-08 (renda na 01, custo na 08).

**B4. TAC ambiental da 08.**
- *Evidência:* "TERMO ACORDO INFRACAO AMBIENTAL AREA SA…", Promotoria de Veranópolis, parcelas 2 e 3 de 3, R$ 834 cada (nota de R$ 2.502), quitado em nov/2023. A área "SA…" está cortada no texto.
- *Correção:* registrar sem valor. Já foi pago, não é ônus pendente.

**B5. Indícios que reforçam a "garantia conta TIS".**
- *Evidência:* IOF da conta TIS no Banrisul em jun/2023 (R$ 404,03), que indica uso de crédito, e seguro da sala do escritório pago em parcelas pela 08. A ordem das datas (matrícula e averbação da garantia em mai–jun/2024, avaliação da sala em out/2024) sugere que a garantia registrada em maio e junho pode ser de outro bem, por exemplo o apto 401, mat. 7.400, item 09, ou que houve uma renovação.
- *Correção:* manter "sala 7.393 ou outro imóvel" e acrescentar a mat. 7.400 entre as certidões a pedir.

**B6. "SIMPLES sobre fretes e serviços" é inferência.**
- *Evidência:* a fonte diz apenas "IMPOSTO S/FATURAMENTO".
- *Correção:* "SIMPLES sobre o faturamento".

**B7. Empréstimo da Rádio pago pela 01.**
- *Evidência:* "P. 28/28 - EMPRESTIMO CAIXA FEDERAL - RADIO C…", R$ 6.500, jul/2023, pago pela 01 à Siviero Transportes. Era a última parcela, quitada antes da data-base.
- *Correção:* registro intragrupo, sem efeito.

**B8. Formato da saída do script.**
- *Evidência:* o script imprime "6.71 %", com ponto, enquanto o documento usa vírgula. O cálculo das parcelas restantes supõe uma parcela por mês.
- *Correção:* nenhuma necessária além de A1.

---

## Respostas às 5 perguntas

1. **Números e classificação:** ver A1 a A4, M9 e B1. Não há INSS de obra, CNO, CEI, INCC, CUB, DARF, GPS, parcelamento, CND ou "Receita Federal" em lançamentos da 01 ou da 08. O INSS da 08 é todo de folha (16 de 16), e não há INSS de folha classificado como de obra. O que passou foi o **ISSQN de obra do posto** e a **CSLL da 01**.
2. **Alíquota:** o valor certo é 6,87 % em 15 meses, não 6,71 % em 12 (A1). As citações do inventário estão corretas no conteúdo (B2). É verdade que os 7 % e os 6 % não estão no total: R$ 137.041.679,60 − 16.240.000,00 − 3.800.000,00 − 16.127.652,00 = R$ 100.874.027,60. As 16 linhas com valor somam exatamente R$ 16.127.652,00, e nenhuma é percentual.
3. **C-10:** afirma mais do que os fatos permitem (A5 e M1). O contra-argumento natural é o valor líquido de realização e o tratamento igual entre os sócios. Há argumentos mais fortes disponíveis (M1, itens 1 e 2).
4. **Achados laterais da 08:**
   - placas solares: as parcelas estão corretas, mas a estimativa é a soma das parcelas, não a dívida (M4);
   - garantia TIS: datas e textos corretos (B5);
   - iluminação: valor subestimado (M5).

   Faltam consórcio, capitalização, veículos, benfeitorias, receita da 01 e devolução de aluguel da borracharia (M5 a M8).
5. **"Não há duplicidade na 01 e na 08":** correta em efeito hoje. Os totais não aplicam as reduções: o total da 08 confere com a soma dos itens, R$ 16.374.986,00, sem os 1,5 %; a 01 não aplica os 0,70 %, conforme o pacote 3. Mas a justificativa está incompleta para a 01: parte do custo de regularização do posto (ISSQN de obra e topografia) **foi paga** no período (A3 e M9). Se o grupo aplicar os 0,70 %, há base documental para alegar duplicidade parcial na 01.

---

## Conferido sem erro

- 3.785 lançamentos (SEI 2.664, 08 919, 01 202); 15 meses com relatório; set/2023 sem relatório; md5 `a29fc2f2eb`.
- 08: INSS de folha, 16 lançamentos, R$ 94.032,06, todos "S/SALARIO FUNC" (um deles é de 13º).
- 08: SIMPLES, 15 lançamentos, R$ 138.734,16.
- 01: PIS + COFINS + IRPJ = R$ 28.254,10 (o valor em si; falta a CSLL, ver A4).
- Nenhum lançamento cita INCC ou CUB, em nenhuma empresa.
- INSS da SEI:
  - "retido sobre notas de serviço" (R$ 26.199,84) é INSS sobre notas de serviço à Prefeitura, à CERAN e à Buriti, não INSS de obra própria. A classificação está aceitável;
  - "INSS (outro)" (R$ 445,72) é contribuição individual para aposentadoria.
- Empréstimo das placas solares: P.24 a P.39, soma de R$ 21.452,98; última parcela R$ 1.314,54 em out/2024; 19 parcelas restantes em 31.12.2024; 19 × R$ 1.314,54 = R$ 24.976,26 (a aritmética está certa; a interpretação está em M4).
- Manutenção das placas solares em nov/2023 (R$ 360) e abr/2024 (R$ 220).
- Garantia: "DESP REGISTRO IMOVEIS MATRICULA BEM GARANTIA" (mai/2024, R$ 149,11), "AVERBACAO GARANTIA" (jun/2024, R$ 273,70) e "AVALIACAO SALA ESCRITORIO GARANTIA CONTA TIS" (out/2024, R$ 951,20). As datas e os textos estão corretos.
- Iluminação: R$ 10.098,99 pagos no período (o valor pago está correto; o investimento está em M5).
- Inventário da SEI, observações: "7,00 % … PIS – COFINS – IRPJ – CSLL" sobre o valor declarado na escritura; "6,00 % … de comissão" sobre o valor do negócio; 5,93 % e 7 % com o adicional na transmissão dos imóveis "pertencentes aos sócios". Os 16 valores somam R$ 16.127.652,00 e não incluem os percentuais.
- O exemplo citado "FATURAMENTO MES JUNHO/2023 = R$ 441.900,00" está na fonte.
- Por amostragem, as alíquotas mensais da SEI batem com o presumido: PIS 0,65 % e COFINS 3 %.
- O script roda com `python -I`, regrava só `_extracao/relatorios-mensais-tributos.xlsx`, e a saída bate com o documento.

**Não conferido** (fora do escopo, só registro da SEI): a classificação interna de ITBI e registro (R$ 58.435,29), projeto e engenharia (R$ 256.727,60) e IPTU (R$ 873.805,43).
