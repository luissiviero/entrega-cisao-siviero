# Coordenação das demais empresas — levantamento e agrupamento

Sessão de coordenação, 28/10/2026. Não é a análise das empresas: é o levantamento mínimo para escrever um prompt específico para cada sessão. Os números abaixo vêm dos documentos, lidos por script (extração de texto dos docx e PDFs), e devem ser re-derivados pelas sessões. Diário desta sessão: `diario.md`, nesta pasta.

## 1. Tabela-resumo

Valores em R$. "Grupo" = total geral do inventário 2025 (31.12.2024, versão atual). "Rafael" = "TOTAL GERAL DA EMPRESA" do balanço executado no PDF da empresa (Estudo 2). "Lista de Rafael" = itens da empresa em `01_fontes/outros/ativos-avaliados-valter.xlsx` (critério C0).

| Nº | Empresa | Situação (inventário 2025) | Itens 2025 | Grupo 2025 | Rafael | Diferença | Lista de Rafael (C0) | 2018 (doc 27) | Sessão |
|---|---|---|---|---|---|---|---|---|---|
| 01 | Abastecedora de Combustíveis | ativa | 1 | 9.661.414,00 | 12.361.850,00 ¹ | +2.700.436,00 | nenhum | 12.361.851,89 | S3 |
| 02 | Agropecuária Fazenda São Crispin | sem bens; imóveis em nome do Hotel Princesa; Rafael registra incorporação ao Hotel "a validar" | 0 | 0,00 | sem balanço ³ | — | nenhum | 12.054.390,00 | S1 |
| 03 | Associação Cultural Princesa dos Vales | baixada em 09/07/2020 | 0 | 0,00 | 0,00 | 0,00 | nenhum | 10.000,00 | S4 |
| 04 | Ceval Terraplenagem | ativa | 3 | 315.000,00 | 385.000,00 | +70.000,00 | nenhum | 391.000,00 | S4 |
| 05 | Hotel Princesa dos Vales | ativa | 8 | 75.410.201,00 | 79.491.957,50 | +4.081.756,50 | 4 itens: Sede Campestre, Fazenda São Crispin (407,53 ha), 4 ha mat. 11.090, Mercedinha | 34.404.877,83 | S1 |
| 06 | Hotel Araçá | ativa | 22 | 58.076.823,20 | 76.230.000,00 | +18.153.176,80 | 5 lotes: 03/q19, 09 e 10/q38, 27 e 28/q38 | 49.684.243,20 | S2 |
| 07 | Rádio Comunidade Princesa dos Vales FM | ativa | 7 | 2.698.000,00 | 2.912.000,00 | +214.000,00 | 1 item: Toyota Bandeirante | 9.316.016,00 | S4 |
| 08 | Transportadora Irmãos Siviero | ativa | 13 | 16.374.986,00 | 16.924.986,00 | +550.000,00 | nenhum | 1.564.000,00 | S3 |
| 09 | Turismo Siviero | baixada em 04/09/2023 | 0 | 0,00 | 0,00 | 0,00 | nenhum | 0,00 | S4 |
| 10 | Siviero Transportes | baixada em 08/04/2024 | 0 | 0,00 | 0,00 | 0,00 | nenhum | 928.000,00 | S4 |
| — | Oficina Mecânica Siviero (só em 2018) ² | não aparece em 2025 | — | — | — | — | — | 29.303.377,44 | S3 (benchmark 2018) |
| | **Total das 10** | | **54** | **162.536.424,20** | **188.305.793,50** | **+25.769.369,30** | | | |

¹ Rafael não avalia o posto: adota o valor do balanço de 2018 apresentado por Valter (R$ 12.361.850,00), "sem prejuízo de eventual avaliação". É um valor de 2018 usado como valor atual, e é o total da empresa em 2018: posto R$ 12.216.851,89 mais o caminhão-tanque ICW 8342, R$ 145.000,00, que não existe em 2025.
³ O PDF de Rafael da Agropecuária (4 páginas) não tem balanço nem total: só transcreve as observações e registra a incorporação.
² A Oficina Mecânica tinha inventário próprio em 2018 (doc 10 de 2018). Em 2025 seus bens estão divididos: terreno da oficina (mat. 3.753), pavilhão, equipamentos, rampa de lavagem e borracharia na Transportadora (itens 01 a 05); terreno da antiga serraria (mat. 3.754) e imóvel Oleoplan na SEI. Proposta: a S3 usa o inventário 2018 da Oficina como benchmark dos itens 01 a 05 da Transportadora; os bens que foram para a SEI ficam com a sessão SEI.

Os totais das 10 somam 54 itens. A diferença Rafael − grupo das 10 empresas é de R$ 25,8 milhões, 70 % dela no Hotel Araçá.

## 2. Agrupamento proposto: 4 sessões

| Sessão | Empresas | Pasta | Grupo 2025 | Rafael | Por que juntas |
|---|---|---|---|---|---|
| S1 | 02 São Crispin + 05 Hotel Princesa | `02-05-sao-crispin-hotel-princesa/` | 75.410.201,00 | 79.491.957,50 | A Fazenda São Crispin está registrada em nome do Hotel e de terceiros (inventários 02 e 05 de 2025); em 2018 ela era avaliada no inventário da Agropecuária, em 2025 no do Hotel. A Agropecuária foi criada "para fins operacionais fiscais-tributários", com um arrendamento Hotel × Agropecuária documentado "pro forma" (resposta ao Dr. Pablo, p. 6) e Rafael registra uma incorporação ao Hotel. O inventário de 2018 das duas, somado, é o benchmark da fazenda. É também a sessão que usa os achados da análise de ACERTOS E ACORDOS e das áreas vendidas. |
| S2 | 06 Hotel Araçá | `06-hotel-araca/` | 58.076.823,20 | 76.230.000,00 | Maior divergência do grupo fora da SEI, 22 itens, 5 na lista de Rafael. Não divide bens com outra empresa. |
| S3 | 01 Abastecedora + 08 Transportadora | `01-08-abastecedora-transportadora/` | 26.036.400,00 | 29.286.836,00 | Complexo da BR-470, km 175: posto (mat. 3.752) ao lado do terreno da oficina (mat. 3.753). A locação à Rede SIM é um contrato conjunto: a Transportadora deu em comodato à Abastecedora parte do seu terreno para fechar a área locada (resposta ao Dr. Pablo, p. 16). As duas e a SEI são as empresas em que Rovílio tinha 2 %. Os itens 01 a 05 da Transportadora vêm da Oficina Mecânica de 2018. |
| S4 | 07 Rádio + 04 Ceval + 03 Associação + 09 Turismo + 10 Siviero Transportes | `03-04-07-09-10-radio-ceval-baixadas/` | 3.013.000,00 | 3.297.000,00 | Empresas pequenas: 10 itens em 2025, três empresas baixadas. Há bens que cruzam entre elas: a Toyota Bandeirante (placa IHW 0285) era da Siviero Transportes em 2018 e está na Rádio em 2025; a Associação foi criada para obter um canal de rádio comunitária, com custos pagos pela Rádio. |

Alternativas consideradas:
- Uma sessão por empresa (10 sessões): três delas sem bens em 2025, e os bens que cruzam (fazenda, terras de Cotiporã, Toyota, oficina) iriam todos para candidatos à Parte 6 em vez de serem resolvidos dentro de uma sessão.
- Rádio sozinha, Ceval com as baixadas (5 sessões): possível. A Rádio tem só 7 itens; o peso dela está na outorga e na participação, que cabem na S4 sem sobrecarga.

Cruzamentos que ficam entre sessões diferentes (cada sessão registra nos seus candidatos à Parte 6):

| Bem | Onde aparece | Sessões |
|---|---|---|
| Terras rurais de Cotiporã: lotes 63, 65, 66 e "N" (69,8 ha, mat. 459, 460, 461, 558) e parte do lote 63 (4,5 ha, registro 15.922) | 2018: Agropecuária São Crispin, itens 07 e 08, "NÃO DIVISÍVEL". 2025: Transportadora, itens 07 e 08 | S1 e S3 |
| Lotes rurais 12, 89 e 33 | 2018: Agropecuária São Crispin, itens 03 a 05. 2025: SEI, áreas rurais | S1 e SEI |
| Sede Campestre (mat. 1.941 e 7.104) e lote 52 Bassani (mat. 7.391) | 2018: Hotel / Agropecuária e lista "bens em outras empresas" da SEI; 2025 antigo: lote 52 também na SEI (retirado na versão atual). 2025: Hotel, itens 04 e 07 | S1 e SEI (candidato nº 2 da SEI) |
| Escavadeira Caterpillar 320D 2009 | 2025: Ceval item 01 (R$ 200.000) e SEI máquinas | S4 e SEI (candidato nº 4 da SEI) |
| Terreno da antena da Rádio (mat. 3.751) | 2018: Rádio, item 06 (R$ 7.148.016,00). 2025: SEI | S4 e SEI (candidato nº 1 da SEI) |
| Terreno da antiga serraria (mat. 3.754) e imóvel Oleoplan | 2018: Oficina Mecânica. 2025: SEI | SEI (a S3 só registra) |
| Venda das quotas de Rovílio (28/08/2025) | 01, 08 e SEI (2 %); 09 e 10 (1 %); texto também no doc 00 | S3 levanta o contrato; S4 usa para 09 e 10; SEI para a SEI |

## 3. Achados que mudam os prompts (confirmados nas fontes)

- **Data-base.** O grupo afirma que os inventários e o relatório financeiro se referem ao "termo final adotado em prévio consenso como sendo outubro/2025" (arquivo 1, p. 1–2; também p. 4). Isso conflita com a P2 (31.12.2024): alerta incluído nos quatro prompts.
- **Rovílio.** No texto das perguntas do Dr. Pablo, reproduzido no arquivo 1 (Anexo H, p. 17 e p. 21), consta venda das quotas em 28/08/2025 "com pagamento por lotes + dinheiro, totalizando R$ 3.440.082,00", com parcela ainda a pagar em lotes do Morada dos Ipês II, contra R$ 2.538.208,55 que seriam 2 % dos valores dos inventários de SEI, Abastecedora e Transportadora; diferença de 35,53 %. A resposta do grupo (p. 23) não confirma valor. Precisa ser conferido no contrato (arquivo 10, 11 páginas, escaneado). A venda é depois da data-base.
- **Arrendamento da São Crispin.** A receita do arrendamento à família Lorenzet é "depositada e saldada conforme fluxo financeiro necessário" para a SEI (arquivo 1, p. 6). Créditos de soja sem valor em R$: Rigo 37.694 sacas; Lorenzetti 13.085 sacas (planilha 2_0, aba GERAL, linhas 986 e 987).
- **Matrícula 00.760 e área da fazenda.** O grupo explica que as duas áreas (48 ha e 100 ha) foram compradas em momentos distintos e que, depois da venda de 83,47 ha, restam 64,53 ha nessa matrícula (arquivo 1, p. 9). O inventário 2018 da Agropecuária fecha a conta com o de 2025: os 190 ha vendidos aos Rigo são 83,47 ha da mat. 00.760 mais 106,53 ha da mat. 11.090 (a mesma matrícula dos 4 ha do item 06 do Hotel), e as quatro parcelas vendidas aos Pellegrini somam 85,22 ha, não os 85,32 declarados, o que explica os 0,10 ha entre 411,43 e 411,53 ha (achado do revisor).
- **Hotel Princesa, matrícula 1.412.** Itens 01 (1.800 m²) e 02 (1.080 m²) citam a mesma matrícula; o grupo explica que a matrícula tem 1.800 m² e a avaliação foi fracionada porque o edifício está sobre 720 m² (arquivo 1, p. 8–9). Conferir se 720 + 1.080 = 1.800 m² não é avaliado duas vezes.
- **Transportadora, nota de R$ 1.135.000,00.** Itens 03 a 05 somam R$ 1.020.000,00. Em 2018 a borracharia valia R$ 25.000,00 (Oficina) e em 2025 vale R$ 220.000,00. Coincidência numérica a testar, não conclusão: o acordo do ônibus IXO 6000 da Siviero Transportes também é de R$ 115.000,00.
- **Abastecedora.** O texto do inventário 2025 ainda fala em "contrato com a Companhia Ipiranga cumprido", mas os contratos entregues são com a Petrobras (arquivos 2 e 3) e a locação é com a Rede SIM (arquivo 4). O valor do item 01 foi apagado da célula na versão atual; o total continua R$ 9.661.414,00 (diff, seção 3). Processo Lucchetta: pedido de R$ 1.032.648,00, sentença de improcedência em 14/08/2025, trânsito em julgado em maio/2026 (fatos depois da data-base).
- **Hotel Araçá.** Matrícula 73.482 em dois lotes (27 e 28/q38), em 2018 e em 2025; a numeração dos itens mudou entre os anos. Em 2018 os itens 20 a 25 eram lotes com o nome de cada sócio no lugar do valor; o item 17 de 2025 (lote 15/q38, mat. 33.003, R$ 390.000) tem a mesma inscrição imobiliária (360.0038.015.000) do item 21 de 2018, o lote "IRINEU SIVIERO": um lote em nome de sócio em 2018 aparece valorado na empresa em 2025. Rafael: item 01 R$ 66.000.000,00 e item 20 (Condado) R$ 2.400.000,00 contra R$ 1.000.000,00; os itens do balanço dele somam R$ 75.870.000,00, mas o total declarado é R$ 76.230.000,00 (R$ 360.000,00 sem item). Ele também aponta duas residências fora do inventário (PDF, p. 24–25). A aba GERAL da planilha 2_0 tem o lote 28/q38 com 2018 = 0 (era R$ 200.000). A obra civil passou de 25 % para 15 % de "vantagem da coisa feita" entre as versões (−R$ 3.713.636,80). O poço tem outorga, mas não é usado (arquivo 1, p. 11).
- **Rádio.** A redução de 8 % inclui o custo de mudar a antena do terreno da SEI e de comprar as quotas do sócio minoritário Antônio Telmo de Toni (1 %) (arquivo 1, p. 13). O doc 00 calcula a cota de cada sócio como (total − R$ 26.980,00 de Telmo de Toni) / 6: a premissa P1 (1/6) precisa de uma decisão do Luís sobre essa parcela de terceiro.
- **Ceval.** O grupo corrige: os valores são da data do inventário, não da compra; a 320D custou R$ 470.000,00 em 12/11/2009 (arquivo 1, p. 7; arquivo 6). Em 2018 a Ceval tinha também um caminhão M. Benz 2423 K 2003 (R$ 93.000,00) que não aparece em 2025.
- **Siviero Transportes.** Em 2018: 3 ônibus em uso (IUY 8068 financiado, IAZ 0281, ISS 7019), Toyota IHW 0285 e 3 ônibus vendidos a prazo em 2015/2016 (IXO 6000, IMY 5000, IOM 9093). Em 2025: empresa baixada, processo do IXO 6000 com acordo de R$ 115.000,00 pago em parcelas desde nov/2024 (crédito sem valor no inventário).

- **Correção trazida pela S3 (28/10/2026), depois da abertura das sessões.** A área extra da locação do posto à Rede SIM é de 2.500 m² da **mat. 3.754, da SEI**, e não da mat. 3.753 da Transportadora (arquivo 4, p. 1, conferido pela S3 na imagem). A 3.753 aparece só como faixa de acesso (arquivo 4, p. 7). A resposta do grupo (arquivo 1, p. 16) fala em comodato da Transportadora: o que a coordenação escreveu na seção 2 e no prompt S3 repete a versão do grupo, que o contrato contradiz. O cruzamento passa a ser Abastecedora × SEI (renda e área da mat. 3.754), além de Abastecedora × Transportadora (acesso).
- **Revisão independente (P11).** Relatório no diário desta pasta, entrada "revisão dos prompts". Nenhum erro de gravidade alta; 6 médios e 10 baixos corrigidos; 11 omissões incorporadas aos prompts.

## 4. Fontes por empresa

Caminhos relativos a `C:\entrega-cisao-siviero\`. `[Drive2018]` = `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\`.

| Nº | Inventário 2025 (atual / antigo) | Inventário 2018 | PDF de Rafael | Aba das planilhas Google | Respostas ao Dr. Pablo |
|---|---|---|---|---|---|
| 01 | `01_fontes/inventario-2025/01. …ABASTECEDORA…docx` / `inventario-2025-antigo/ANTIGO 01. …` | `[Drive2018]12. 1. ABASTECEDORA…docx` | `01_fontes/estudo-rafael/ABASTECEDORA DE COMBUSTIVEIS SIVIERO LTDA, BELISE LTDA.pdf` | `1.Abastecedora`; GERAL linhas 5–10 | arq. 1 p. 3–4 e 21–23; arq. 2, 3, 4, 10 |
| 02 | `…/02. …SÃO CRISPIN…docx` / não há antigo | `[Drive2018]17. 1. AGROPECUÁRIA…docx` | `…/AGROPECUARIA E FAZENDA SÃO CRISPIN LTDA - BELISE LTDA.pdf` | `2.São Crispin`; GERAL 11–15, 941–950 e 986–987; aba Extras | arq. 1 p. 5–6; arq. 7, 8, 9 |
| 03 | `…/03. …ASSOCIAÇÃO…docx` / `ANTIGO 03. …` | `[Drive2018]14. 1. ASSOCIAÇÃO…docx` | `…/ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES LTDA - BELISE LTDA.pdf` | GERAL 16–19 (sem aba própria) | — |
| 04 | `…/04. …CEVAL…docx` / `ANTIGO 04. …` | `[Drive2018]11. 1. CEVAL…docx` | `…/CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA, BELISE LTDA.pdf` | `4.Ceval`; GERAL 20–27 | arq. 1 p. 7; arq. 5, 6 |
| 05 | `…/05. …HOTEL PRINCESA…docx` / `ANTIGO 05. …(2).docx`; docx `HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS…` | `[Drive2018]13. 1. HOTEL PRINCESA…docx` | `…/HOTEL PRINCESA DOS VALES LTDA- BELISE LTDA.pdf` | `5.Hotel Princesa`; GERAL 28–69 e 941; aba Extras | arq. 1 p. 8–9; arq. 11 |
| 06 | `…/06. …HOTEL ARAÇÁ…docx` / `ANTIGO 06. …` | `[Drive2018]16. 1. HOTEL ARAÇÁ…docx` | `…/HOTEL ARAÇÁ LTDA - BELISE LTDA.pdf` | `6.Hotel Araçá`; GERAL 70–104 | arq. 1 p. 10–11; arq. 12 |
| 07 | `…/07. …RÁDIO…docx` / `ANTIGO 07. …` | `[Drive2018]9. 1. RÁDIO…docx` | `…/RADIO COMUNIDADE PRINCESA DOS VALES FM LTDA, BELISE LTDA.pdf` | `7.Rádio`; GERAL 105–115 | arq. 1 p. 12–13; arq. 13, 14 |
| 08 | `…/08. …TRANSPORTADORA…docx` / `ANTIGO 08. …` | `[Drive2018]15. 1. TRANSPORTADORA…docx` e `10. 1. OFICINA MECÂNICA…docx` | `…/TRANSPORTADORA IRMÃOS SIVIERO LTDA, BELISE LTDA.pdf` | `8.Transportadora` (2 notas "?"); GERAL 116–134 | arq. 1 p. 14–16 e 21–23; arq. 10 |
| 09 | `…/09. …TURISMO…docx` / `ANTIGO 09. …` | `[Drive2018]8. 1. TURISMO…docx`; doc `23. 1. SIVIERO TRANSPORTES LTDA.  +  TURISMO SIVIERO LTDA.   =   D Í V I D A S …docx` (nome com letras espaçadas) | `…/TURISMO SIVIERO LTDA - BELISE LTDA.pdf` | GERAL 135–138 | arq. 1 p. 23; arq. 10 |
| 10 | `…/10. …SIVIERO TRANSPORTES…docx` / `ANTIGO 10. …` | `[Drive2018]7 .1. SIVIERO TRANSPORTES…docx`; doc 23 de 2018 | `…/SIVIERO TRANSPORTES LTDA - BELISE LTDA.pdf` | `10.Transportes`; GERAL 139–152 | arq. 1 p. 23; arq. 10 |

Fontes comuns a todas: `01_fontes/inventario-2025/00. PLANILHA DO VALOR DO CAPITAL…docx`; `01_fontes/inventario-2018/27. 1. PLANILHA DO VALOR DO CAPITAL…RESUMO…docx`; `02_trabalho/diff-inventario-2025-antigo-x-atual.md` (seções 2 a 12); `01_fontes/outros/ativos-avaliados-valter.xlsx` e o PDF Diretrizes; `01_fontes/outros/conferencia-discrepancias.md` e `inventario-2025-avaliacoes-corrigido.xlsx` (cobrem 03, 04, 06, 07, 08, 09 e 10; não cobrem 01, 02 e 05); `02_trabalho/pendencias-conferencia-anterior.md`; `02_trabalho/observacoes-planilhas-google.xlsx`; planilhas Google exportadas em `01_fontes/planilhas-google/`; `02_trabalho/parte6-candidatos-duplicidade.md` (SEI, só leitura).

Os inventários de 2018 das demais empresas não foram copiados para `01_fontes` e `00_FONTES.md` é da sessão SEI (P13). Cada sessão lê do Drive e registra caminho e hash no seu diário.

## 5. Arquivos desta pasta

- `analise-coordenacao.md` — este arquivo.
- `prompts/S1-02-05-sao-crispin-hotel-princesa.md`, `prompts/S2-06-hotel-araca.md`, `prompts/S3-01-08-abastecedora-transportadora.md`, `prompts/S4-03-04-07-09-10-radio-ceval-baixadas.md`.
- `diario.md` — diário da coordenação (P12).
- `sessoes.md` — registro das sessões abertas, depois da aprovação.
