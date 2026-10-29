# Índice de fontes da entrega

Gerado em 28/10/2026. Cópias em `01_fontes/`; os originais ficam no Drive (`[Drive]` = `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero`) e no repositório (`[Repo]` = `C:\inventario-siviero`). O hash MD5 (10 primeiros dígitos) identifica a versão copiada; se o original mudar, o hash deixa de bater.

## 1. Decisões sobre duplicados entre Drive e repositório

| Material | Onde está duplicado | Resultado da comparação | Versão canônica |
|---|---|---|---|
| Inventário 2025 (docs 00 a 11, docx Hotel Princesa, ACERTOS E ACORDOS.pdf) | `[Drive]/Material Siviero/2025/inventário-2025` e `[Repo]/2025/inventário-2025` | Todos os 14 arquivos idênticos por hash | Drive (o repositório é cópia). O Drive tem ainda `ACERTOS E ACORDOS.gdoc`, exportado para `01_fontes/planilhas-google`. |
| Planilhas Excel do Estudo 2 (empreendimentos) | `[Drive]/Material Rafael/Estudo 2/versão-excel` e `[Repo]/belise-empreendimentos` | 12 planilhas idênticas; o Drive tem a 13ª, `LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS` (31 loteamentos, 450 lotes), com o spec correspondente e `build_summary.py`/`README.md` atualizados. O `00 - RESUMO` do Drive tem 3 linhas a mais em Documentos/Observações e 17 divergências a mais. A alteração sem commit no repositório é só uma regeneração do mesmo conteúdo (14/43/14 linhas, igual ao HEAD). | **Drive `versão-excel`**. O repositório está atrasado em um documento; sincronizar depois (fora do escopo desta entrega). |
| `inventário-estudo-rafael.pdf` e `.xlsx` | `[Repo]` raiz e `[Drive]/Material Rafael` | O PDF do repositório é o mesmo arquivo que `DIRETRIZES PARA SELEÇÃO DE ATIVOS NA SEPARAÇÃO PATRIMONIAL E SOCIETÁRIA` do Estudo 2 (hash igual). O xlsx é idêntico nos dois lugares. | Drive, com o nome real (DIRETRIZES…). |
| Planilha 2025 de avaliações | `[Repo]/inventario-2025-avaliacoes.xlsx`, `…-corrigido.xlsx` e `[Repo]/inventario-irmaos/inventario-2025-siviero.xlsx` | `inventario-irmaos/` tem Ativos com 50 linhas (versão anterior); a raiz tem 152 linhas e a versão corrigida acrescenta a aba "Correções e pendências". `inventario-irmaos/ativos-avaliados-valter.xlsx` é cópia idêntica da raiz. | `inventario-2025-avaliacoes-corrigido.xlsx`. `inventario-irmaos/` tratada como obsoleta (pendência 5 em `pendencias-conferencia-anterior.md`). |
| Planilhas Google "Comparativo Patrimonial" | Drive: `2_0_…` (principal, 15 abas), `Comparativo … empreendimentos e Rafael` (18 abas, inclui "Rafael × Inventário", "Valor médio dos lotes" e "Lotes que saíram 2018-2025"), `1_0_…` (versão anterior, nov/2025) e duas "Cópia de 2_0_…" (uma de sivierow, uma de bentasiviero) | As cópias e a 1_0 não são usadas. | `2_0_…` para as observações do Luís (Parte 7) e `…e Rafael` para a SEI (abas exclusivas). Exportadas para `01_fontes/planilhas-google`. |
| Inventário 2025 "antigo" | `[Drive]/Material Siviero/Material antigo/2025 antigo` (11 docx, sem o 02) | Comparado com a versão atual em `02_trabalho/diff-inventario-2025-antigo-x-atual.md`. | Versão atual. A antiga serve só para registrar o que mudou. |

## 2. Arquivos copiados para `01_fontes/`

### `inventario-2018/` — Inventário de dez/2018 (benchmark de existência do patrimônio, P3)

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| 18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.docx` | 2026-03-02 | 369 KB | `a9097fa4ec` |
| 18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. = I N V E N T Á R I O = E M P R E S A   D A   S I V I E R O   -   Dezembro de 2018.xlsx` | 2026-10-26 | 39 KB | `8c63e21506` |
| 27. 1. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - RESUMO   -   Dezembro de 2018.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\27. 1. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - RESUMO   -   Dezembro de 2018.docx` | 2025-09-22 | 18 KB | `af5ec6c22e` |
| 21. 1. INVENTÁRIO - EMPRESAS DA SIVIERO - Minutas das Planilhas   -    Dezembro de 2018.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\21. 1. INVENTÁRIO - EMPRESAS DA SIVIERO - Minutas das Planilhas   -    Dezembro de 2018.docx` | 2025-09-22 | 21 KB | `0f4d83ebb2` |
| 25. 1. PLANILHA = DIVISÃO DE LOTES=ANO  2018 = SÓCIOS EMPRESAS SIVIERO - VÁLTER+PEDRO+IRINEU+RAUL+ÉLCIO+ADÉLIA - Em 31.12.2018   -   Dezembro de 2018.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2018\inventário-2018\25. 1. PLANILHA = DIVISÃO DE LOTES=ANO  2018 = SÓCIOS EMPRESAS SIVIERO - VÁLTER+PEDRO+IRINEU+RAUL+ÉLCIO+ADÉLIA - Em 31.12.2018   -   Dezembro de 2018.docx` | 2025-09-22 | 77 KB | `07fe4f03f4` |

### `inventario-2025/` — Inventário de 31.12.2024 do grupo Siviero (avaliação dos 5/6); idêntico por hash à cópia em C:\inventario-siviero

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| 00. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - 31.12.2024.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\00. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - 31.12.2024.docx` | 2026-01-28 | 323 KB | `a1670b10fe` |
| 01. INVENTÁRIO - 31.12.2024 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\01. INVENTÁRIO - 31.12.2024 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA..docx` | 2025-11-14 | 326 KB | `92e48032ae` |
| 02. INVENTÁRIO - 31.12.2024 - AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\02. INVENTÁRIO - 31.12.2024 - AGROPECUÁRIA FAZENDA SÃO CRISPIN LTDA..docx` | 2025-11-14 | 327 KB | `63f1037122` |
| 03. INVENTÁRIO - 31.12.2024 - ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\03. INVENTÁRIO - 31.12.2024 - ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES.docx` | 2025-11-10 | 29 KB | `51551f116e` |
| 04. INVENTÁRIO - 31.12.2024 - CEVAL TERRAPLENAGEM E ESCAVAÇÕES LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\04. INVENTÁRIO - 31.12.2024 - CEVAL TERRAPLENAGEM E ESCAVAÇÕES LTDA..docx` | 2025-11-10 | 30 KB | `39e3a75ec6` |
| 05. INVENTÁRIO - 31.12.2024 - HOTEL PRINCESA DOS VALES LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\05. INVENTÁRIO - 31.12.2024 - HOTEL PRINCESA DOS VALES LTDA..docx` | 2026-02-09 | 332 KB | `39bb7a345a` |
| 06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx` | 2025-11-10 | 43 KB | `4371e018c7` |
| 07. INVENTÁRIO - 31.12.2024 - RÁDIO COMUNIDADE PRINCESA DOS VALES FM LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\07. INVENTÁRIO - 31.12.2024 - RÁDIO COMUNIDADE PRINCESA DOS VALES FM LTDA..docx` | 2025-11-10 | 33 KB | `2fdea61145` |
| 08. INVENTÁRIO - 31.12.2024 - TRANSPORTADORA IRMÃOS SIVIERO LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\08. INVENTÁRIO - 31.12.2024 - TRANSPORTADORA IRMÃOS SIVIERO LTDA..docx` | 2025-11-10 | 39 KB | `d3882994f6` |
| 09. INVENTÁRIO - 31.12.2024 - TURISMO SIVIERO LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\09. INVENTÁRIO - 31.12.2024 - TURISMO SIVIERO LTDA..docx` | 2025-11-10 | 30 KB | `fedbb88ae1` |
| 10. INVENTÁRIO - 31.12.2024 - SIVIERO TRANSPORTES LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\10. INVENTÁRIO - 31.12.2024 - SIVIERO TRANSPORTES LTDA..docx` | 2025-11-10 | 31 KB | `1fcc0230c5` |
| 11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx` | 2026-10-26 | 1.086 KB | `5c9d13d0d9` |
| ACERTOS E ACORDOS.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\ACERTOS E ACORDOS.pdf` | 2025-11-10 | 2.342 KB | `5a6e856dcd` |
| HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\inventário-2025\HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx` | 2025-11-10 | 29 KB | `6cdf8807c9` |

### `inventario-2025-antigo/` — Versão anterior do inventário 2025 (para o diff da Fase 0.4)

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| ANTIGO 00. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - 31.12.2024.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 00. PLANILHA DO VALOR DO CAPITAL DE CADA EMPRESA - 31.12.2024.docx` | 2025-11-26 | 25 KB | `8ed9d82a32` |
| ANTIGO 01. INVENTÁRIO - 31.12.2024 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 01. INVENTÁRIO - 31.12.2024 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA..docx` | 2025-11-26 | 30 KB | `65b09c17e1` |
| ANTIGO 03. INVENTÁRIO - 31.12.2024 - ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 03. INVENTÁRIO - 31.12.2024 - ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES.docx` | 2025-11-26 | 28 KB | `53aa5e4999` |
| ANTIGO 04. INVENTÁRIO - 31.12.2024 - CEVAL TERRAPLENAGEM E ESCAVAÇÕES LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 04. INVENTÁRIO - 31.12.2024 - CEVAL TERRAPLENAGEM E ESCAVAÇÕES LTDA..docx` | 2025-11-26 | 29 KB | `d64cd0ff63` |
| ANTIGO 05. INVENTÁRIO - 31.12.2024 - HOTEL PRINCESA DOS VALES LTDA.(2).docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 05. INVENTÁRIO - 31.12.2024 - HOTEL PRINCESA DOS VALES LTDA.(2).docx` | 2025-11-26 | 42 KB | `d50daf8944` |
| ANTIGO 06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 06. INVENTÁRIO - 31.12.2024 - HOTEL ARAÇÁ LTDA..docx` | 2025-11-26 | 330 KB | `4d33595ff0` |
| ANTIGO 07. INVENTÁRIO - 31.12.2024 - RÁDIO COMUNIDADE PRINCESA DOS VALES FM LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 07. INVENTÁRIO - 31.12.2024 - RÁDIO COMUNIDADE PRINCESA DOS VALES FM LTDA..docx` | 2025-11-26 | 18 KB | `5d49c8f932` |
| ANTIGO 08. INVENTÁRIO - 31.12.2024 - TRANSPORTADORA IRMÃOS SIVIERO LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 08. INVENTÁRIO - 31.12.2024 - TRANSPORTADORA IRMÃOS SIVIERO LTDA..docx` | 2025-11-26 | 39 KB | `f3c4896281` |
| ANTIGO 09. INVENTÁRIO - 31.12.2024 - TURISMO SIVIERO LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 09. INVENTÁRIO - 31.12.2024 - TURISMO SIVIERO LTDA..docx` | 2025-11-26 | 29 KB | `52379cc2d8` |
| ANTIGO 10. INVENTÁRIO - 31.12.2024 - SIVIERO TRANSPORTES LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 10. INVENTÁRIO - 31.12.2024 - SIVIERO TRANSPORTES LTDA..docx` | 2025-11-26 | 30 KB | `3a2f7908ce` |
| ANTIGO 11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Material antigo\2025 antigo\ANTIGO 11. INVENTÁRIO - 31.12.2024 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA..docx` | 2025-11-26 | 375 KB | `c2897b89b8` |

### `estudo-rafael/` — Estudo 2 de Rafael (BELISE LTDA): PDFs gerais por empresa

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| ABASTECEDORA DE COMBUSTIVEIS SIVIERO LTDA, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\ABASTECEDORA DE COMBUSTIVEIS SIVIERO LTDA, BELISE LTDA.pdf` | 2026-09-23 | 332 KB | `84a05e9b14` |
| AGROPECUARIA E FAZENDA SÃO CRISPIN LTDA - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\AGROPECUARIA E FAZENDA SÃO CRISPIN LTDA - BELISE LTDA.pdf` | 2026-09-23 | 250 KB | `b3a7709f8d` |
| ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES LTDA - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\ASSOCIAÇÃO CULTURAL PRINCESA DOS VALES LTDA - BELISE LTDA.pdf` | 2026-09-23 | 239 KB | `e3ae5f5f07` |
| CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA, BELISE LTDA.pdf` | 2026-09-23 | 277 KB | `1e54cbbea4` |
| DIRETRIZES PARA SELEÇÃO DE ATIVOS NA SEPARAÇÃO PATRIMONIAL E SOCIETÁRIA , BENTA E V SIVIERO BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\DIRETRIZES PARA SELEÇÃO DE ATIVOS NA SEPARAÇÃO PATRIMONIAL E SOCIETÁRIA , BENTA E V SIVIERO BELISE LTDA.pdf` | 2026-09-23 | 384 KB | `aff2b08f2a` |
| HOTEL ARAÇÁ LTDA - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\HOTEL ARAÇÁ LTDA - BELISE LTDA.pdf` | 2026-09-23 | 4.357 KB | `f844ffc907` |
| HOTEL PRINCESA DOS VALES LTDA- BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\HOTEL PRINCESA DOS VALES LTDA- BELISE LTDA.pdf` | 2026-09-23 | 4.421 KB | `067bfc3721` |
| INTRODUÇÃO AO ESTUDO DO BALANÇO PATRIMONIAL DO GRUPO SIVIERO - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\INTRODUÇÃO AO ESTUDO DO BALANÇO PATRIMONIAL DO GRUPO SIVIERO - BELISE LTDA.pdf` | 2026-09-23 | 308 KB | `01a85c04c0` |
| LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.pdf` | 2026-09-23 | 5.343 KB | `8d5018a270` |
| RADIO COMUNIDADE PRINCESA DOS VALES FM LTDA, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\RADIO COMUNIDADE PRINCESA DOS VALES FM LTDA, BELISE LTDA.pdf` | 2026-09-23 | 321 KB | `b22568ded8` |
| SIVIERO TRANSPORTES LTDA - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\SIVIERO TRANSPORTES LTDA - BELISE LTDA.pdf` | 2026-09-23 | 251 KB | `56b624935a` |
| TRANSPORTADORA IRMÃOS SIVIERO LTDA, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\TRANSPORTADORA IRMÃOS SIVIERO LTDA, BELISE LTDA.pdf` | 2026-09-23 | 397 KB | `b09ef6ee08` |
| TURISMO SIVIERO LTDA - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\TURISMO SIVIERO LTDA - BELISE LTDA.pdf` | 2026-09-23 | 242 KB | `df2116a02f` |

### `estudo-rafael/empreendimentos-pdf/` — Estudo 2 de Rafael: PDFs dos empreendimentos da SEI

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| DEMIAS COMPROMISSOS A PAGAR, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\DEMIAS COMPROMISSOS A PAGAR, BELISE LTDA.pdf` | 2026-09-22 | 235 KB | `729bff137d` |
| LOTEAMENTO BONETTO  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\LOTEAMENTO BONETTO  BELISE LTDA.pdf` | 2026-09-22 | 333 KB | `9a1adbd2b1` |
| LOTEAMENTO COLINA DO RETIRO II BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\LOTEAMENTO COLINA DO RETIRO II BELISE LTDA.pdf` | 2026-09-22 | 392 KB | `586780a776` |
| LOTES DE TERCEIROS INDENIZAVEIS, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\LOTES DE TERCEIROS INDENIZAVEIS, BELISE LTDA.pdf` | 2026-09-22 | 319 KB | `341c8e6fec` |
| LOTES URBANOS EM VERANÓPOLIS RS - BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\LOTES URBANOS EM VERANÓPOLIS RS - BELISE LTDA.pdf` | 2026-09-22 | 4.328 KB | `cd818b176e` |
| ÁREAS CONTRATADAS LOTEAMENTO CAMINHOS DO SOL,  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO CAMINHOS DO SOL,  BELISE LTDA.pdf` | 2026-09-22 | 272 KB | `80b64a64ae` |
| ÁREAS CONTRATADAS LOTEAMENTO DA FAMÍLIA CECCONELLO – CAXIAS DO SUL, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO DA FAMÍLIA CECCONELLO – CAXIAS DO SUL, BELISE LTDA.pdf` | 2026-09-22 | 301 KB | `b7a78e6fc5` |
| ÁREAS CONTRATADAS LOTEAMENTO DAS FAMILIAS PEDROTTI, ANDREIS ,LORANDI,  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO DAS FAMILIAS PEDROTTI, ANDREIS ,LORANDI,  BELISE LTDA.pdf` | 2026-09-22 | 259 KB | `e568251847` |
| ÁREAS CONTRATADAS LOTEAMENTO DR LETTI E SOL NASCENTE III, BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO DR LETTI E SOL NASCENTE III, BELISE LTDA.pdf` | 2026-09-22 | 309 KB | `9bacfa5482` |
| ÁREAS CONTRATADAS LOTEAMENTO PASQUAL,  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO PASQUAL,  BELISE LTDA.pdf` | 2026-09-22 | 309 KB | `44d2306171` |
| ÁREAS CONTRATADAS LOTEAMENTO VALE DAS COLINAS,  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO VALE DAS COLINAS,  BELISE LTDA.pdf` | 2026-09-22 | 279 KB | `db5ce21dfd` |
| ÁREAS CONTRATADAS LOTEAMENTO VISTA DA COLINA,  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS CONTRATADAS LOTEAMENTO VISTA DA COLINA,  BELISE LTDA.pdf` | 2026-09-22 | 276 KB | `a6dec0a133` |
| ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL,  BELISE LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\Empreendimentos\ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL,  BELISE LTDA.pdf` | 2026-09-22 | 4.396 KB | `bab6104a63` |

### `estudo-rafael/empreendimentos-xlsx/` — Conversão para Excel dos PDFs de empreendimentos (versão canônica: Drive versão-excel, que inclui o 13º documento)

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| 00 - RESUMO BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\00 - RESUMO BELISE LTDA.xlsx` | 2026-10-27 | 14 KB | `caf72e6376` |
| DEMIAS COMPROMISSOS A PAGAR, BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\DEMIAS COMPROMISSOS A PAGAR, BELISE LTDA.xlsx` | 2026-10-26 | 12 KB | `03648fc4aa` |
| LOTEAMENTO BONETTO  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\LOTEAMENTO BONETTO  BELISE LTDA.xlsx` | 2026-10-26 | 18 KB | `fac7b04862` |
| LOTEAMENTO COLINA DO RETIRO II BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\LOTEAMENTO COLINA DO RETIRO II BELISE LTDA.xlsx` | 2026-10-26 | 18 KB | `876d5dc0b3` |
| LOTES DE TERCEIROS INDENIZAVEIS, BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\LOTES DE TERCEIROS INDENIZAVEIS, BELISE LTDA.xlsx` | 2026-10-26 | 15 KB | `6e70c914da` |
| LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\LOTES LOTEAMENTOS EXECUTADOS E REGISTRADOS, BELISE LTDA.xlsx` | 2026-10-27 | 92 KB | `ca58713b9c` |
| LOTES URBANOS EM VERANÓPOLIS RS - BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\LOTES URBANOS EM VERANÓPOLIS RS - BELISE LTDA.xlsx` | 2026-10-26 | 23 KB | `db1ad58136` |
| ÁREAS CONTRATADAS LOTEAMENTO CAMINHOS DO SOL,  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO CAMINHOS DO SOL,  BELISE LTDA.xlsx` | 2026-10-26 | 15 KB | `ce58270ef1` |
| ÁREAS CONTRATADAS LOTEAMENTO DA FAMÍLIA CECCONELLO – CAXIAS DO SUL, BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO DA FAMÍLIA CECCONELLO – CAXIAS DO SUL, BELISE LTDA.xlsx` | 2026-10-26 | 15 KB | `75cd53bf08` |
| ÁREAS CONTRATADAS LOTEAMENTO DAS FAMILIAS PEDROTTI, ANDREIS ,LORANDI,  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO DAS FAMILIAS PEDROTTI, ANDREIS ,LORANDI,  BELISE LTDA.xlsx` | 2026-10-26 | 15 KB | `0c7b9bf4a4` |
| ÁREAS CONTRATADAS LOTEAMENTO DR LETTI E SOL NASCENTE III, BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO DR LETTI E SOL NASCENTE III, BELISE LTDA.xlsx` | 2026-10-26 | 15 KB | `13af411e1c` |
| ÁREAS CONTRATADAS LOTEAMENTO PASQUAL,  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO PASQUAL,  BELISE LTDA.xlsx` | 2026-10-26 | 15 KB | `025d26bccf` |
| ÁREAS CONTRATADAS LOTEAMENTO VALE DAS COLINAS,  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO VALE DAS COLINAS,  BELISE LTDA.xlsx` | 2026-10-26 | 17 KB | `d4827994c2` |
| ÁREAS CONTRATADAS LOTEAMENTO VISTA DA COLINA,  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS CONTRATADAS LOTEAMENTO VISTA DA COLINA,  BELISE LTDA.xlsx` | 2026-10-26 | 16 KB | `bc4c784229` |
| ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL,  BELISE LTDA.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\Estudo 2\versão-excel\ÁREAS RURAIS EM VERANÓPOLIS, COTIPORÃ, FAGUNDES VARELA E CAXIAS DO SUL,  BELISE LTDA.xlsx` | 2026-10-26 | 44 KB | `67838bd992` |

### `respostas-docs-2026-08-13/` — Respostas do grupo Siviero ao Dr. Pablo (arquivos 1 a 18, 13.08.2026)

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| 1 - grupo siviero - encaminhamento esclarecimentos ao Dr. Pablo - 13.08.2026.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\1 - grupo siviero - encaminhamento esclarecimentos ao Dr. Pablo - 13.08.2026.pdf` | 2026-10-27 | 1.367 KB | `03968eb843` |
| 10 - CONTRATO DE COMPRA E VENDA DE QUOTAS DE CAPITAL - ROVÍLIO SIVIERO.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\10 - CONTRATO DE COMPRA E VENDA DE QUOTAS DE CAPITAL - ROVÍLIO SIVIERO.pdf` | 2026-10-27 | 2.382 KB | `23d9b1732d` |
| 11 - MATRÍCULA 1.412 - HOTEL PRINCESA DOS VALES LTDA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\11 - MATRÍCULA 1.412 - HOTEL PRINCESA DOS VALES LTDA.pdf` | 2026-10-27 | 708 KB | `d27c0a1be3` |
| 12 - OUTORGA_POCO_HOTEL_ARACA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\12 - OUTORGA_POCO_HOTEL_ARACA.pdf` | 2026-10-27 | 148 KB | `a8b7aa7ef5` |
| 13 - RÁDIO COMUNIDADE PRINCESA DOS VALES LTDA. - MATRÍCULAS DA SALA E BOX INSTALAÇÕES DOS ESTÚDIOS DA RÁDIO.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\13 - RÁDIO COMUNIDADE PRINCESA DOS VALES LTDA. - MATRÍCULAS DA SALA E BOX INSTALAÇÕES DOS ESTÚDIOS DA RÁDIO.pdf` | 2026-10-27 | 1.287 KB | `f43d974b6a` |
| 14 - RÁDIO COMUNIDADE PRINCESA DOS VALES LTDA. - OUTORGA DE FUNCIONAMENTO E PEDIDO DE RENOVAÇÃO - MINICON.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\14 - RÁDIO COMUNIDADE PRINCESA DOS VALES LTDA. - OUTORGA DE FUNCIONAMENTO E PEDIDO DE RENOVAÇÃO - MINICON.pdf` | 2026-10-27 | 641 KB | `7a9f214447` |
| 15 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - CONTRATO DE LOCAÇÃO DOS PAVILHÕES OLEOPLAN SA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\15 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - CONTRATO DE LOCAÇÃO DOS PAVILHÕES OLEOPLAN SA.pdf` | 2026-10-27 | 1.761 KB | `d522c07dab` |
| 16 - TERMOS DE ACORDO EXTRAJUDICIAIS - ALEXANDRE SUSIN E AGI EMPREENDIMENTOS  17 LOTES E 07 LOTES.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\16 - TERMOS DE ACORDO EXTRAJUDICIAIS - ALEXANDRE SUSIN E AGI EMPREENDIMENTOS  17 LOTES E 07 LOTES.pdf` | 2026-10-27 | 4.761 KB | `ca73deb381` |
| 17 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - LOTES VENDIDOS DE JANEIRO DE 2018 A OUTUBRO DE 2025 - REV.13.08.2026.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\17 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - LOTES VENDIDOS DE JANEIRO DE 2018 A OUTUBRO DE 2025 - REV.13.08.2026.pdf` | 2026-10-27 | 1.018 KB | `b88e91089e` |
| 17 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - LOTES VENDIDOS DE JANEIRO DE 2018 A OUTUBRO DE 2025 - REV.13.08.2026.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\17 - SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. - LOTES VENDIDOS DE JANEIRO DE 2018 A OUTUBRO DE 2025 - REV.13.08.2026.xlsx` | 2026-10-27 | 32 KB | `7ef0dc35dc` |
| 18 - PLANILHA DE LOTES VENDIDOS E VALORES PAGOS PARA AQUISIÇÃO DE BENS - VÁLTER MATERIAIS DE CONSTRUÇÃO LTDA. E VÁLTER SIVIERO.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\18 - PLANILHA DE LOTES VENDIDOS E VALORES PAGOS PARA AQUISIÇÃO DE BENS - VÁLTER MATERIAIS DE CONSTRUÇÃO LTDA. E VÁLTER SIVIERO.docx` | 2026-10-27 | 14 KB | `dcfb1bf9e1` |
| 18 - PLANILHA DE LOTES VENDIDOS E VALORES PAGOS PARA AQUISIÇÃO DE BENS - VÁLTER MATERIAIS DE CONSTRUÇÃO LTDA. E VÁLTER SIVIERO.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\18 - PLANILHA DE LOTES VENDIDOS E VALORES PAGOS PARA AQUISIÇÃO DE BENS - VÁLTER MATERIAIS DE CONSTRUÇÃO LTDA. E VÁLTER SIVIERO.pdf` | 2026-10-27 | 95 KB | `8da060b1a3` |
| 2 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. - CONTRATO COMPRA DE PRODUTOS E USO DA MARCA PETROBRÁS.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\2 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. - CONTRATO COMPRA DE PRODUTOS E USO DA MARCA PETROBRÁS.pdf` | 2026-10-27 | 2.547 KB | `03f8c4559f` |
| 3 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. - CONTRATO DE ANTECIPAÇÃO DE BONIFICAÇÃO PETROBRÁS.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\3 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. - CONTRATO DE ANTECIPAÇÃO DE BONIFICAÇÃO PETROBRÁS.pdf` | 2026-10-27 | 1.337 KB | `b83ae4ba92` |
| 4 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. - CONTRATO DE LOCAÇÃO SIM ADMINISTRADORA DE IMÓVEIS LTDA..pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\4 - ABASTECEDORA DE COMBUSTÍVEIS SIVIERO LTDA. - CONTRATO DE LOCAÇÃO SIM ADMINISTRADORA DE IMÓVEIS LTDA..pdf` | 2026-10-27 | 3.664 KB | `361f8b3a52` |
| 5 - CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA.- CASE DO BRASIL -NOTA FISCAL AQUISIÇÃO RETROESCAVADEIRA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\5 - CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA.- CASE DO BRASIL -NOTA FISCAL AQUISIÇÃO RETROESCAVADEIRA.pdf` | 2026-10-27 | 969 KB | `9ff5a5dadc` |
| 6 - CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA.- PARANÁ EQUIPAMENTOS -NOTA FISCAL AQUISIÇÃO ESCAVADEIRA HIDRÁULICA.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\6 - CEVAL TERRAPLANAGEM E ESCAVAÇÕES LTDA.- PARANÁ EQUIPAMENTOS -NOTA FISCAL AQUISIÇÃO ESCAVADEIRA HIDRÁULICA.pdf` | 2026-10-27 | 576 KB | `bfc853f6fc` |
| 7 - CONTRATO DE ARRENDAMENTO DA ÁREA DE TERRAS DA FAZENDA SÃO CRISPIN- FAMÍLIA LORENZET - 2017-2020.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\7 - CONTRATO DE ARRENDAMENTO DA ÁREA DE TERRAS DA FAZENDA SÃO CRISPIN- FAMÍLIA LORENZET - 2017-2020.pdf` | 2026-10-27 | 3.644 KB | `1af6d0b269` |
| 8 - CONTRATO DE ARRENDAMENTO DA ÁREA DE TERRAS DA FAZENDA SÃO CRISPIN- FAMÍLIA LORENZET - 2020-2023.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\8 - CONTRATO DE ARRENDAMENTO DA ÁREA DE TERRAS DA FAZENDA SÃO CRISPIN- FAMÍLIA LORENZET - 2020-2023.pdf` | 2026-10-27 | 3.159 KB | `99104ed82f` |
| 9 - CONTRATO DE ARRENDAMENTO DA ÁREA DE TERRAS DA FAZENDA SÃO CRISPIN- FAMÍLIA LORENZET - 2023-2026.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\2025\[siviero] respostas docs\9 - CONTRATO DE ARRENDAMENTO DA ÁREA DE TERRAS DA FAZENDA SÃO CRISPIN- FAMÍLIA LORENZET - 2023-2026.pdf` | 2026-10-27 | 1.372 KB | `7b519bed33` |

### `outros/` — Estudos e planilhas de apoio já existentes

| Arquivo | Origem | Data | Tamanho | MD5 |
|---|---|---|---|---|
| Estudo - lotes que saíram 2018-2025.md | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Estudo - lotes que saíram 2018-2025.md` | 2026-10-26 | 15 KB | `9a9f1b23b7` |
| Lotes vendidos 2018-2025 x GERAL.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Lotes vendidos 2018-2025 x GERAL.xlsx` | 2026-10-27 | 196 KB | `087986e972` |
| Contas_Pagas_2023-07_a_2024-10.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Relatórios Mensais\Contas_Pagas_2023-07_a_2024-10.xlsx` | 2026-10-26 | 519 KB | `a29fc2f2eb` |
| Contas_Pagas_e_Lotes_Vendidos_2023-07_a_2024-10.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Relatórios Mensais\Contas_Pagas_e_Lotes_Vendidos_2023-07_a_2024-10.xlsx` | 2026-10-26 | 570 KB | `80e181d0e1` |
| Contas_Pagas_Recorrentes_x_Extraordinarias_2023-07_a_2024-10.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Siviero\Relatórios Mensais\Contas_Pagas_Recorrentes_x_Extraordinarias_2023-07_a_2024-10.xlsx` | 2026-10-27 | 225 KB | `f2b2a5f940` |
| conferencia-discrepancias.md | `C:\inventario-siviero\conferencia-discrepancias.md` | 2026-10-26 | 14 KB | `bdd7c87e70` |
| inventario-2025-avaliacoes-corrigido.xlsx | `C:\inventario-siviero\inventario-2025-avaliacoes-corrigido.xlsx` | 2026-10-26 | 44 KB | `a565450170` |
| ativos-avaliados-valter.xlsx | `C:\inventario-siviero\ativos-avaliados-valter.xlsx` | 2026-10-26 | 22 KB | `935282c1aa` |
| inventário-estudo-rafael.xlsx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Material Rafael\inventário-estudo-rafael.xlsx` | 2026-10-26 | 17 KB | `e8ca17953f` |
| Emailing MEMORANDO BURITI X SIVIERO - LOTEAMENTOS DE CAXIAS DO SUL - 11-05-22 rev 13.05.2022.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Documentos\Emailing MEMORANDO BURITI X SIVIERO - LOTEAMENTOS DE CAXIAS DO SUL - 11-05-22 rev 13.05.2022.pdf` | 2025-11-14 | 579 KB | `9957040587` |
| Contrato Siviero X Buriti Holding Sul Participações.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Documentos\Contrato Siviero X Buriti Holding Sul Participações.pdf` | 2026-03-02 | 7.778 KB | `e06ff8d3d6` |
| Alexandre Susin - TERMO DE ACORDO.docx | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Documentos\Alexandre Susin - TERMO DE ACORDO.docx` | 2026-01-28 | 20 KB | `a98f32a4c6` |
| retiradas_socios_2018.pdf | `G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero\Documentos\retiradas_socios_2018.pdf` | 2026-07-01 | 2.501 KB | `96c512ef14` |

### `planilhas-google/` — exportações das planilhas Google e do gdoc ACERTOS E ACORDOS

Exportadas em 28/10/2026 pelo conector do Google Drive (somente leitura). Ver seção 1 para a decisão de uso.

| Arquivo | Origem (ID no Drive) | Conteúdo |
|---|---|---|
| 2_0_Comparativo Patrimonial Siviero 2018 x 2025.xlsx | `1-6Ym_773J419BlZ0g0uKrwoXilv0cHoDOUK9y2bxGgI` (Luís, modificada em 27/10/2026) | 15 abas; 6 comentários e 68 células de anotação do Luís (31 na aba 11.Empreendimentos). |
| Comparativo Patrimonial Siviero empreendimentos e Rafael 2018 x 2025.xlsx | `1WPku66w325IIcns09o1qQyniD8SWi0jCttDyE943-As` (modificada em 27/10/2026) | 18 abas, inclui "Rafael × Inventário", "Valor médio dos lotes" e "Lotes que saíram 2018-2025"; 87 notas, 7 comentários, 68 anotações. |
| ACERTOS E ACORDOS (gdoc).docx | `1_E0pkTbKofnCvu92JyfCvPuYgUNMJF4Zl4dFGZTmIJE` (nov/2025) | Versão em texto do PDF escaneado ACERTOS E ACORDOS (141 parágrafos). |

As notas, comentários e anotações foram extraídos para `02_trabalho/observacoes-planilhas-google.xlsx` (Parte 7).

### `geofarroupilha/` — mapa oficial de APP de Farroupilha (consulta de 28/10/2026)

Fonte pública da Prefeitura de Farroupilha (GEOFarroupilha, serviços ArcGIS): `https://geo.farroupilha.rs.gov.br/server/rest/services/Geo_Farroupilha/Marco_Regulatorio/MapServer` (mapa do Marco Regulatório das APPs, Lei Municipal 4.814/2023) e `.../_GEO_Farroupilha_Publico_Cadastro_Imobiliario/MapServer` (quadras e lotes). Consultada por `02_trabalho/_scripts/geofarroupilha_app.py`. É a versão do dia da consulta, não a de 31.12.2024 (registro de data-base nº 10). O texto da lei não foi obtido.

| Arquivo | Conteúdo | MD5 |
|---|---|---|
| lotes-sei-2026-10-28.json | Geometria e inscrição dos 95 lotes da SEI em Farroupilha | `d671ea8c61` |
| quadras-sei-2026-10-28.json | Inscrições encontradas por quadra | `85f61304c8` |

Resultado por lote: `02_trabalho/geofarroupilha-app-lotes.xlsx`.

### `respostas-docs-2026-08-13/` — OCR dos arquivos 2 a 16

Texto por OCR (Tesseract 5.4, português) em `02_trabalho/ocr-respostas-docs/`. Serve só para localizar informação: todo número é conferido na imagem da página. Guia e limitações: `02_trabalho/ocr-respostas-docs/LEIA-ME.md`.

## 3. Fontes referenciadas sem cópia

| Material | Caminho | Por que não foi copiado |
|---|---|---|
| Relatórios mensais (16 PDFs, jul/2023 a out/2024, ~70 MB) | `[Drive]/Material Siviero/Relatórios Mensais` | As três planilhas de contas pagas extraídas deles estão em `outros/`; os PDFs só serão abertos se a Parte 6 precisar de um lançamento específico. |
| Demais documentos do inventário 2018 (sócios, empresas que não a SEI, contratos de 2003) | `[Drive]/Material Siviero/2018/inventário-2018` | Só a SEI (doc 18), o resumo (27), as minutas (21) e a divisão de lotes (25) são usados nesta entrega. |
| Inventário 2015 e material 2018 desatualizado | `[Drive]/Material Siviero/Material antigo` | Contexto histórico; usado apenas pelo estudo de lotes que saíram, já copiado. |
| Contratos das holdings | `[Drive]/Holdings` | Não entram na avaliação patrimonial. |
| Estudo 1 de Rafael (docx de set/2025) | `[Drive]/Material Rafael/Estudo 1` | Superado pelo Estudo 2. |
| Projeto urbanístico Sol Nascente (DWG) | `[Drive]/Documentos/urbanísticos` | Candidato da Parte 5; formato não legível aqui, será citado por referência. |
