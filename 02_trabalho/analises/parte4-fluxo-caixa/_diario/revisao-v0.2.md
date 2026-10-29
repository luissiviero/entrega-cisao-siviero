# Revisão independente — Parte 4 SEI, verificação da v0.2 (só as mudanças)

Revisor: agente de contexto novo (P11). Data: 28/10/2026.
Objeto: `build_fluxo_caixa.py` (v0.2), comparado com `build_fluxo_caixa_v0_1.py`, mais `fluxo-caixa-sei-v0.2.xlsx`, `numeros-v0.2.json` e a nova planilha de contas pagas `contas-pagas-classificadas_snapshot-2026-10-28_1853.xlsx` (abas Extrapolação e IPCA).
Efeitos em R$, na ordem piso / central / teto.

## O que confere

**IPCA**
- Os números-índice da aba IPCA reproduzem a inflação anual oficial: 2019 4,31%; 2020 4,52%; 2021 10,06%; 2022 5,79%; 2023 4,62%; 2024 4,83%.
- Recalculei a soma dos fatores de jan/2019 a dez/2024: com base em mar/2024 (índice 6.869,14) dá **63,5962**; com base em dez/2018 (índice 5.100,61) dá **85,6469**.

**Custos**
- C1: 4.940.478,77 / 15 × 63,5962 = **20.946.388**.
- Os 4.940.478,77 são os recorrentes da SEI na planilha nova (5.165.698,54), menos combustível (75.219,77) e menos Flávia (150.000).
- O método é igual ao da aba Extrapolação (B7 e D12 por SUMIFS), com duas exceções intencionais: combustível fica em INV, por instrução direta do Luís; Flávia é contada pelas parcelas.
- C1b: 510.768,34, contados uma vez. Recalculado pela soma das 11 categorias extraordinárias classificadas como M.
- Mapeamento das categorias novas: sem categoria órfã. O pró-labore do Hotel Araçá e as despesas de familiares e de outras empresas voltaram para M, pela regra do Luís de 28/10.

**Aluguéis e dívida**
- Aluguéis: 35.326,83 × 85,6469 = **3.025.634**.
- Flávia: a parcela P.26 foi paga em jul/2023; somando os 17 meses até dez/2024 chega-se a 43 parcelas, ou 430.000.
- C2 = 7.320.900 − 2.121.396,40 + 430.000 = **5.629.503,60**.

**Deflator do teto**
- Fórmula: v25 × IPCA(mês da saída) / IPCA(dez/2024).
- Amostra conferida: Cichelero 12/02/2019 dá 0,7237, 22/03/2022 dá 0,8895 e 26/08/2024 dá 0,9811.
- Registros só com ano usam julho; registros sem data usam dez/2021; registros de 2025 ficam com fator 1 (o índice é travado nos limites da tabela).

**Alocação depois do N1**
- Os registros anteriores ao período só usam a redução que sobra. Totais: 271,96 no período + 34,04 em 2025 + 30 antes do período = 336.
- Conferi a parte anterior ao período em 11 loteamentos (Carnevalli 6, Barbiero 7, Jardim Itália 5, Pedra Lisa 3, entre outros).
- Lago Negro: não sobra redução para os registros anteriores ao período. Está certo pela regra.

**Outras mudanças**
- N2: Morada do Ipê I voltou a v25 = 629.000. A regra THIN foi retirada.
- Linha 65 (móveis): foi para "destino declarado". G2 passou a 1.246.500.

**Ponte**
- Conferida nos 4 cenários: piso, central, teto e teto_nominal. B − C = D; D + E = F; F + G1 + G2 + G3 = H; I = H + A2 + A3c; o total A é igual à soma das 8 classes. Diferença 0 em todos.
- Valores de H: piso −1,47 mi, central +16,79 mi, teto +28,46 mi (H/6 de 4,74 mi), teto_nominal +38,21 mi.
- Sensibilidade (início 2019 ou 2018, concessão documentada ou máxima): recalculada.

## O que ainda está errado

**V1 (média). A regra de troca lote por lote não pega a Monte Vêneto.**
O padrão `permuta de outro imovel . lote` exige um caractere entre "imovel" e "lote". Mas `norm()` apaga o travessão e junta os espaços, e o texto vira "permuta de outro imovel lote 03". Com isso, as linhas 332, 333, 339 e 340 continuam em VENDA.
- 332/333: o lote recebido (Femaçã II 03/04) está no estoque de 2025, a 189.000 (GERAL linha 187). Deveriam ser **BEM**.
- 339/340: o lote recebido (Medianeira 04/04) não aparece em 2025. Deveriam ser BEMX. Para H, BEMX e VENDA dão o mesmo resultado.

Efeito: H está alto em cerca de 0,14 / 0,17 / 0,17 mi. Correção: usar `permuta de outro imovel\W*lote`.

**V2 (baixa). Linha 99 (Solar Norte 06/01, os 3.100 m² de Sapopema) está como BEMX, mas a área está em 2025.**
A própria aba 6 trata a linha 76 como "mesma área do item A.2 de 2018 mais 3.100 m²", ou seja, uma área que está no inventário de 2025. Deveria ser BEM. Efeito: H está alto em cerca de 0,13 / 0,18 / 0,18 mi.

**V3 (baixa). As linhas 82 (paisagismo) e 449/450 (Attolini) mudaram de classe inteiras, sem dividir o valor.**
Cada uma ainda tem uma parte em dinheiro: na 82, a venda é a prazo em 16 vezes; na 449/450, há R$ 80.000 em espécie. Essa parte saiu de VENDA. Efeito: H está baixo em cerca de 0,3 mi no central. Correção: dividir o valor do registro entre as classes.

**V4 (baixa). Concessão do Falavigna e custo extrapolado se sobrepõem.**
Nos 15 meses, R$ 84.700 de obra do Falavigna estão em "Obras e manutenção em loteamentos finalizados" (M). Extrapolados, viram cerca de 0,36 mi, e isso se soma à concessão de 0,45 mi pela mesma obra. Efeito: H está baixo em até cerca de 0,36 mi. Na mesma categoria há um frete para o Lot. Dr. Letti (4.617), que é um loteamento contratado e não finalizado; isso vem da classificação da planilha do Luís.

**V5 (baixa). O teto agora é uma mudança de método e precisa de aprovação do Luís.**
O deflator IPCA no teto (item 4 da lista de mudanças) não consta como decisão do Luís na aba Extrapolação, e mexe na premissa P5 (só valores do grupo). Pela mudança, H no teto cai de 38,21 para 28,46 mi. O central continua interpolando até o v25 nominal, sem deflator, então o central e o teto não estão na mesma base.

**V6 (baixa). Rótulos que ficaram com o período antigo.**
- Ponte, linha A: diz "2018–2024".
- Aba 2: "antes de 2018 (vendas documentadas na contagem de 2018)" e "Forma de saída (2018–2024)". O período agora começa em 2019, então deveria ser "antes de 2019" e "2019–2024".
- Sensibilidade: o texto "saídos em 2018–2024 além da redução" se refere aos 26 lotes, que são calculados sobre 2019–2024.
- JSON: a chave `antes_2018` vale para 2019.

**V7 (baixa). Sensibilidade de 84 meses.**
Os meses de 2018 recebem o IPCA de dez/2018, porque a tabela começa nesse mês. Isso aumenta um pouco os custos de 2018 e o teto das vendas de 2018.
