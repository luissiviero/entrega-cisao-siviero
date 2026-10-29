# Revisão independente — planilha-mestre 01 Abastecedora / 08 Transportadora

- **Data:** 28/10/2026
- **Objeto:** `planilha-mestre-01-08.xlsx` (abas Mestre, Créditos e contingências, Resumo, Legenda), `_scripts/build_mestre_01_08.py`, `_scripts/extract_fase0.py`, `_extracao/fase0-itens-01-08.xlsx`, `decisoes.json`, `registro-data-base.md`.
- **Não lidos (por instrução):** `_diario/` (exceto a criação deste arquivo), `pacotes/`, `rovilio-quotas.md`.
- **Método:** cada número da planilha foi conferido direto na fonte (docx dos inventários 2025 e 2018 lidos por script próprio; PDFs de Rafael e arq. 1 lidos pela camada de texto; planilha Google lida em fórmulas e valores; contratos 2–4 pela transcrição conferida `_extracao/contratos-abastecedora.md`). Nenhum arquivo existente foi alterado.

## Execução do script

Para não sobrescrever a planilha, o fase0 nem o `hashes.txt` (o script regrava os três), copiei os dois scripts, `decisoes.json` e a transcrição do arq. 10 para o scratchpad e rodei `python -I` a partir de `C:\`, apontando as constantes de fonte (`INV25`, `RAFAEL`, `GSHEET` etc.) para as fontes reais. Resultado: o script roda sem erro e gera 20 linhas na Mestre e 12 na aba de créditos. A comparação célula a célula com os arquivos existentes deu **0 diferenças** em todas as 4 abas da mestre e nas 12 do fase0.

## Achados

Gravidade: **alta** = muda valor, total, par, destino ou situação; **média** = omissão relevante ou referência errada; **baixa** = forma.

### Alta

**A1. Destino 2018→2025 vazio nos 7 bens que mudaram de empresa.** Nas linhas 08-2025-01 a 05 (2018 na Oficina Mecânica, doc 10 itens 01–05) e 08-2025-07 e 08 (2018 na Agropecuária, doc 17 itens 07–08), a coluna "Destino 2018→2025" está vazia. São justamente os bens que trocaram de empresa. As fontes dizem algo sobre isso:
- a Oficina de 2018 não tinha CNPJ (doc 10, cabeçalho: "CNPJ/MF: ......");
- o grupo afirma que os imóveis estão "registrados em nome da empresa conforme consta no inventário" (arq. 1 p. 15, item 1);
- o contrato do Rovílio diz que a área de São Brás, em Cotiporã, está "em nome das COMPRADORAS" (arq. 10, cláusula 3ª, transcrição `_extracao/arq10-rovilio-transcricao.md`).

Nenhuma fonte traz um ato formal de transferência.
*Correção:* preencher, por exemplo, "Oficina Mecânica (2018, sem CNPJ) → 08 Transportadora (2025); sem documento de transferência; grupo afirma registro em nome da 08 (arq. 1 p. 15)" e "Agropecuária (2018, 'R$ NÃO DIVISÍVEL') → 08 (2025); em nome da 08 segundo o arq. 10, cl. 3ª; sem documento de transferência". Para as linhas que ficaram na mesma empresa (01-2025-01 e 08-2025-06, 09 a 13), escrever "permanece na 01" ou "permanece na 08".

**A2. O total "Bens de 2018 sem par em 2025 (R$ 2018)" do Resumo mistura um bem já deduzido em 2018.** Resumo, linha 13: 08 = R$ 399.000 e Total = R$ 544.000. Os R$ 399.000 incluem o apto 102 (mat. 7.395, R$ 220.000), que o próprio inventário de 2018 já deduz do total da empresa: doc 15, T4.R3, "R$ 1.784.000,00 (-) R$ 220.000,00 R$ 1.564.000,00". A planilha Google de Luís usa a mesma base líquida (8.Transportadora H2 = −179.000 = 149.000 + 30.000).
*Correção:* separar a linha em "sem par, não deduzidos em 2018" (01 R$ 145.000; 08 R$ 179.000; total R$ 324.000) e "doado e deduzido em 2018" (08 R$ 220.000), ou registrar a exclusão no rótulo.

### Média

**M1. A Mestre não tem a coluna "Existe em 2025?", que o plano exige.** Hoje a informação só se deduz de "Descrição 2025" e "Valor grupo 2025" estarem preenchidas.
*Correção:* incluir a coluna (Sim/Não, com a referência "doc 01/08 T.R").

**M2. O par por descrição dos itens 02 a 05 tem uma justificativa sem fonte, e a localização da borracharia e da rampa é contraditória nas fontes.** A célula "Chave registral" das linhas 08-2025-02 a 05 diz "benfeitoria sobre a mat. 3.753". Nenhuma fonte liga esses bens à mat. 3.753: o inventário 08 de 2025 diz só "instalados sobre uma das propriedades da empresa". Além disso:
- o item 03 são equipamentos, não benfeitoria;
- o item 02 não tem a "mesma descrição" de 2018 ("em fase de construção", R$ 746,60/m² em 2018, contra "em fase FINAL", R$ 1.100/m² em 2025).

As fontes divergem sobre onde ficam a borracharia e a rampa:
- a locação da Abastecedora à SIM trata "restaurante/lancheria, borracharia e rampa de lavagem" como espaços da locadora, locados a terceiros e a incluir por aditivo (arq. 4 p. 1 e 3; transcrição §3.3);
- o grupo diz que a Transportadora deu em comodato à Abastecedora "a área faltante" para fechar a área locada à SIM, e que os espaços da rampa e da borracharia "estão locados" (arq. 1 p. 16, item 6).

*Correção:* trocar o texto da chave por "sem chave; mesma posição na lista; localização não documentada (arq. 1 p. 16 × arq. 4 p. 1/3)". Registrar nas linhas 04 e 05 e na aba de créditos que esses bens estão locados a terceiros e possivelmente dentro da área cedida à SIM.

**M3. O comodato da 08 para a 01 (arq. 1 p. 16) falta, e contradiz o que a planilha registra.** A linha "Aluguel da locação à SIM" diz que a área locada "inclui 2.500 m² da mat. 3.754 (SEI)". Isso está correto pelo contrato (arq. 4 p. 1–2). Mas o grupo afirma que a área faltante veio da Transportadora, em comodato. Pelos inventários, a mat. 3.754 é da SEI (doc 11 item 03) e a 08 tem a mat. 3.753. A planilha não registra o comodato nem a divergência. O fato afeta a disponibilidade da linha 08-2025-01 (e talvez a 04 e a 05) até 31/12/2033 + 36 meses.
*Correção:* incluir uma linha na aba Créditos e contingências ("Comodato de área da 08 à 01 para a locação SIM; matrícula não identificada; contrato cita a mat. 3.754 da SEI") e citar essa linha nas linhas 08-2025-01, 04 e 05.

**M4. A hipoteca sobre o posto e a multa rescisória Petrobras não foram registradas.**
- A garantia dos contratos Petrobras é uma hipoteca (arq. 2 item VII; arq. 3 item XI). A locação diz que a hipoteca está "registrada no imóvel da Locadora" (arq. 4 p. 12, cl. 21ª §3º; transcrição §3.6). A locadora é a Abastecedora, e o único imóvel dela no inventário é a mat. 3.752. Mesmo assim, a linha 01-2025-01 não traz o ônus, e a aba de créditos fala em "hipoteca de imóvel não identificado".
- Falta a multa rescisória do arq. 2 (cl. 5.2, p. 4): tabela com R$ 300.000,00 na faixa do volume contratado, decrescente pelo volume adquirido. A obrigação é do revendedor (a Certo), mas está garantida pela hipoteca.

*Correção:* anotar na linha 01-2025-01 "hipoteca em favor da Petrobras, provavelmente sobre a mat. 3.752 (arq. 4 p. 12); confirmar na matrícula". Incluir a multa de R$ 300.000 como contingência garantida, com DB-04.

**M5. Faltam observações de Luís sobre 01/08 que estão na planilha Google.**
- `Resumo!B196:B199`, seção "TRANSPORTADORA IRMÃOS SIVIERO LTDA.": "Não entendi a parte da observação: 'FOI ACRESCENTADO … R$ 1.135.000,00 … ITENS 03 – 04 – 05'". Deveria constar na linha "Nota: acrescentado ao capital de bens" da aba de créditos, ou nas linhas 08-2025-03/04/05.
- `GERAL!H874` (repetida em `11.Empreendimentos!I723`), `Resumo!B53:E63` (C59 Abastecedora 2 % = 193.228,28; C60 Transportadora 2 % = 327.499,72; C63 "Diferença de 36.53%"), `Resumo!B67` e `Resumo!B4` ("questionar diferença entre a avaliação do patrimônio da cisão do Rovílio e a nossa"). Todas tratam dos 2 % de Rovílio na 01 e na 08. O próprio fase0 lista H874, C63 e B67 em "Notas do autor", mas a Mestre não as usa. Deveriam ir para DB-01, ou para uma linha geral da 01 e da 08.
- `Resumo!B16:B23` ("VALUATIONS NECESSÁRIAS — A íntegra dos documentos: 01. Abastecedora … 08. Transportadora") e `Resumo!B72:E72` ("Abastecedora", −2.555.437,89, que repete a nota da linha 01-2025-01).

*Correção:* incluir essas células em `NOTES` (no script), com situação própria.

**M6. O registro de data-base não traz o fato de Rafael nos veículos da 08.** Nas linhas 08-2025-11, 12 e 13 (Discorda), os valores de Rafael (50.000, 160.000 e 40.000; Rafael 08 p. 15) vêm de pesquisa de mercado atual: anúncios do Mercado Livre e da OLX e tabela FIPE, com links (Rafael 08 p. 8–12). O PDF foi gerado em 21/09/2026 (metadados). Não é um valor em 31.12.2024. O mesmo vale para a declaração "aproximadamente R$ 1.000.000,00" do item 03 (p. 5–6), que não tem data de referência.
*Correção:* criar DB-07, "Rafael 08, itens 03 e 11–13: valores de mercado e declaração de 2026, não em 31.12.2024; Quem fugiu: Rafael", e citá-lo nessas linhas.

**M7. Aluguéis a terceiros ficaram fora da aba Créditos e contingências.** O inventário 01 de 2025 (item 01) diz "Agrega ao valor do Posto, o aluguel da Lancharia/Restaurante". O inventário 08 de 2025 (itens 04 e 05) e o arq. 1 p. 16 dizem que a rampa e a borracharia "estão locados". A locação SIM prevê ainda o acréscimo do aluguel das sublocações (arq. 4 p. 4, cl. 3ª §3º). Só o aluguel SIM está na aba, e com o valor de 2022.
*Correção:* incluir as receitas de locação (valor desconhecido; pedir os contratos) e indicar que os R$ 15.000 são o valor de 2022 (ver B6).

### Baixa

**B1. A Situação foge do vocabulário do plano.** O plano prevê Concorda, Discorda, Só grupo, Só Rafael e Não abordado. A planilha usa "Só 2018" e "Fora do escopo (SEI)" nas linhas de 2018. Além disso, "Só grupo", que o script pode gerar, não está definido na Legenda.
*Correção:* usar "Não abordado" com um qualificador (por exemplo, "Não abordado — só 2018") ou registrar a extensão do vocabulário na Legenda.

**B2. Em DB-05, a página está imprecisa.** A frase "termo final … outubro/2025" está só na p. 1 do arq. 1. A p. 2 não a contém; a p. 4 contém, e está certa.

**B3. Em DB-03, a coluna "Data do fato" mistura fatos anteriores à data-base.** O acordo de ago/2020 e a desistência de fev/2022 são anteriores a 31.12.2024. Só a defesa de jul/2026 e o "processo está para réplica" (arq. 1 p. 16) são posteriores.
*Correção:* deixar só os fatos posteriores, ou separar contexto de fato posterior.

**B4. A coluna "Registro data-base" das linhas 08-2025-07 e 08 cita DB-01 para os "créditos hereditários sobre a terra de São Brás".** O texto de DB-01 em `registro-data-base.md` trata só do rateio de 6 × 16,66 % e não menciona São Brás.
*Correção:* incluir o ponto em DB-01 (arq. 10, cl. 3ª) ou criar um DB próprio.

**B5. A Legenda tem imprecisões.**
- Em "Créditos e contingências", diz "Valores escritos nas observações … e nos contratos (arq. 1, 3, 4)". Mas as duas reduções são calculadas pelo script (0,70 % × 9.661.414 e 1,50 % × 16.374.986), e o arq. 1 não é contrato.
- Não há verbete para "Destino 2018→2025", "Na lista de Rafael" nem "Situação da observação".
- "Par por descrição — confirmar (benfeitorias …)" também abrange equipamentos (item 03).

**B6. Na aba de créditos, o aluguel SIM está com o valor de 2022.** Os R$ 15.000 valem de 01/01/2022 a 31/12/2022. Em 31.12.2024 o valor já teria dois reajustes anuais (arq. 4 p. 1 e 4).
*Correção:* anotar "valor de 2022; valor em 31.12.2024 não informado".

**B7. Texto em coluna numérica.** "Valor 2018 (R$)" = "R$ NÃO DIVISÍVEL" nas linhas 08-2025-07 e 08, numa coluna com formato `#,##0.00`.
*Correção:* deixar a célula vazia e passar o texto para a Nota.

**B8. A chave da linha OFI-2018-07 está incompleta.** Ela diz "mat. 17409, mat. 6136". Em 2025 o bem está na mat. 29.954, que é a unificação das duas (doc 11 item 04).
*Correção:* acrescentar "→ mat. 29.954".

**B9. A "Situação da observação" da linha 01-2018-02 sugere uma alteração que não existe.** O texto "Alterado: … rótulo trocado" dá a entender que algo foi alterado; nenhuma planilha foi. O fato é que o rótulo de `Resumo!B73` ("Hotel Princesa dos Vales") está errado, porque C73 = GERAL!C8 é o caminhão-tanque da Abastecedora.
*Correção:* "Conferido: rótulo B73 errado na planilha Google; o valor é o caminhão-tanque (GERAL C8)".

**B10. A explicação da redução de 2.555.437,89 na linha 01-2025-01 omite uma parcela.** A planilha atribui a redução às "parcelas por m² e do equipamento". Ela também inclui a vantagem de 30 % (828.888,90 → 557.634,00, −271.254,90). Conferi: −1.380.000,00 − 191.550,26 − 362.632,73 − 350.000,00 − 271.254,90 = −2.555.437,89.

**B11. O destino dos móveis do escritório pode ter uma pista.** Na linha 08-2018-07 (móveis do escritório, R$ 30.000), "sem destino documentado" está correto, mas o item 06 de 2025 (sala 7.393) passou a dizer "CUB … + instalações".
*Correção:* anotar isso como hipótese a perguntar, sem parear.

**B12. Robustez do script.** A doação do apto 102 é detectada por código fixo (`num == "09"`), porque a descrição de 2018 vem cortada em 160 caracteres no fase0 e "doado" fica fora do corte (`build_mestre_01_08.py`, cálculo de `donated`). Também não é erro de conteúdo, mas convém registrar: rodar o script regrava `_extracao/fase0-itens-01-08.xlsx` e `_extracao/hashes.txt`.

## Conferido sem erro

- **Script:** roda com `python -I` fora da pasta e reproduz exatamente a planilha e o fase0 (0 diferenças célula a célula). `decisoes.json` está sem itens, regras e casamentos, e as colunas "Valor adotado", "Decisão (P5)" e "Motivo da decisão" estão vazias em todas as linhas.
- **Valores do grupo 2025 (14 itens):**
  - 01, item 01: célula de valor vazia; recomposto pelas parcelas da descrição, 7.245.000 + 439.275 + 769.505 + 650.000 + 557.634 = 9.661.414, igual ao total do doc 01 (T3.R3) e do doc 00 (T2.R2 e T3.R2).
  - 08, itens 01–13: iguais ao doc 08 (T3.R2–R7 e T4.R2–R8); a soma dá 16.374.986, igual ao doc 00.
- **Valores de Rafael:**
  - 01: R$ 12.361.850,00, adotado na p. 2 (texto) e na p. 3 (balanço).
  - 08: p. 13 (itens 1–3: 8.637.186; 1.427.800; 1.000.000), p. 14 (itens 4–7: 300.000; 220.000; 840.000; 3.500.000) e p. 15 (itens 8–13: 250.000; 360.000; 140.000; 50.000; 160.000; 40.000). Total 16.924.986 (p. 16). Páginas citadas na Mestre corretas.
  - Rafael declara, para 01 e 08, que não há bens selecionáveis (01 p. 5; 08 p. 17). Nenhum bem da 01 ou da 08 está em `ativos-avaliados-valter.xlsx` (108 linhas, empresas 05, 06, 07 e 10). Logo "Na lista de Rafael = Não" e "C0: fica na empresa, maior valor" estão certos, e só indicam a direção.
- **Diferenças e situações:** 5 Discorda (01-01; 08-03, 11, 12, 13) e 9 Concorda; diferenças em R$ e em % corretas (por exemplo, 2.700.436 / 9.661.414 = 27,95 %).
- **Valores de 2018:**
  - doc 12: posto 12.216.851,89; caminhão ICW 8342 145.000; total 12.361.851,89.
  - doc 15: 140.000; 30.000; 140.000; 30.000; 149.000; 685.000; 30.000; 360.000; 220.000; bruto 1.784.000 e líquido 1.564.000.
  - doc 10: 8.637.186; 1.259.812,84; 500.000; 300.000; 25.000; 4.950.000; 13.631.378,60; total 29.303.377,44.
  - doc 17, itens 07–08: "R$ NÃO DIVISÍVEL", fora do total de 24.112.390 (somei as parcelas).
- **Pares por chave:**
  - mat. 3.752 (posto); mat. 3.753 (Oficina 01 ↔ 08-01); mat. 7.393 (sala); mat. 7.400 (apto 401);
  - placas IKR 2054, IQB 0612, IJZ 5399 e IJZ 8207;
  - mat. 459, 460, 461 e 558, e registro 15.922 (Agropecuária 07/08 ↔ 08-07/08);
  - mat. 3.754 → SEI doc 11 item 03 (4.950.000); mat. 17.409 e 6.136 → SEI doc 11 item 04 (mat. 29.954; 13.631.378,60).
  - Os pares sem chave (Oficina 02–05 ↔ 08-02–05) estão marcados como "par por descrição — confirmar".
  - Nenhum bem faltando: todos os itens de 2018 (doc 12: 2; doc 15: 9; doc 10: 7; doc 17: 07–08) e de 2025 (01: 1; 08: 13) têm linha.
- **"Sem destino documentado":** caminhão ICW 8342, sala da mat. 18.952 (Palugana) e móveis do escritório. Procurei por matrícula, placa, endereço e nome em todos os inventários 2025 (00–11, versões atual e antiga), no arq. 1 inteiro (32 p.), na transcrição dos contratos e em todas as abas da planilha Google; não há destino. Apto 102: a doação a Geni Siviero Mattiuz está no doc 15 item 09.
- **Aba Créditos e contingências, valores:**
  - redução 0,70 % = 67.629,90 e redução 1,50 % = 245.624,79;
  - Lucchetta 1.032.648,00 (ajuizada em 31/07/2023; sentença em 14/08/2025; trânsito em maio/2026; arq. 1 p. 4);
  - bonificação 750.000,00 (arq. 3 item III; cessão à Abastecedora Siviero; restituição a cargo da Certo);
  - aluguel SIM 15.000 (2022) e mat. 3.754 com 2.500 m² (arq. 4 p. 1–2); prazo até 31/12/2033 + 36 meses;
  - INSS 122.254,73 e 62.555,51;
  - Joveni 340.000 (150.000 em 2018);
  - Sottili 364.320, 1.518.000 e 627.000, e acordo da seguradora em ago/2020 (arq. 1 p. 15);
  - nota de 1.135.000 contra a soma dos itens 03–05, 1.020.000.
- **Resumo:** todos os números recalculados conferem: 1/13/14 itens; 9.661.414 / 16.374.986 / 26.036.400; 12.361.850 / 16.924.986 / 29.286.836; Concorda 0/9 (15.674.986); Discorda 1/4 (grupo 9.661.414 / 700.000; Rafael 12.361.850 / 1.250.000); lista de Rafael 0; valor adotado 0; sem par 1/3. O problema de critério está em A2.
- **Observações de Luís citadas:** os textos batem com as células — `1.Abastecedora!I7` e `GERAL!H7` ("Por que reduziu R$2.555.437,89", que é 12.216.851,89 − 9.661.414); `8.Transportadora!C10/C11` ("?"); `GERAL!C124` ("Não divisível"); `GERAL!C125` ("?"); `Resumo!B73/C73` ("Hotel Princesa dos Vales", 145000 = GERAL!C8).
- **Registro de data-base:**
  - DB-01: contrato de 28/08/2025; versão antiga com Rovílio a 2 % e atual com 6 × 16,66 %.
  - DB-02: datas conferem com o arq. 1 p. 4.
  - DB-04: vigência de 01/01/2021 a 30/09/2025 (arq. 2 p. 1; arq. 3 p. 1).
  - DB-06: Rafael 01 p. 2–3.
- **Legenda:** descreve corretamente a origem, o ID, o critério de chave, a fonte do valor de Rafael e o C0 (só direção). Ressalvas em B5.
