# Diário de trabalho — sessão S4 (03 Associação, 04 Ceval, 07 Rádio, 09 Turismo, 10 Siviero Transportes)

Só acréscimos (P12). Para Claude e para o revisor; não é material de leitura. Números sempre do script, nunca de memória.

## 2026-10-28 — abertura da sessão
- Brief: `02_trabalho/empresas/_orquestracao/prompts/S4-03-04-07-09-10-radio-ceval-baixadas.md` (aprovado pelo Luís em 28/10/2026). Lidos: `00_PLANO.md` (P1–P13, seção 7), `02_trabalho/log-decisoes.md`, fim de `02_trabalho/_diario/diario.md`, `_orquestracao/analise-coordenacao.md`, memória do projeto.
- Pasta própria criada: `_diario/`, `_scripts/`, `_extracao/`, `pacotes/`.
- Notas compartilhadas da família de sessões: pasta vazia ao abrir.

## 2026-10-28 — Fase 0: extração
- Script: `_scripts/fase0_extracao.py` (md5 67c981486ec7612137bf27b34deab3b6).
- Saídas: `_extracao/fase0-extracao.xlsx` (md5 f48aef4219eaf154b17394c8a783f93e), `_extracao/fase0-extracao.json` (md5 fb98b11a0a039f284d1251db4b8494ef). Hashes de todas as fontes na aba Fontes do xlsx e no json.

### Entradas (md5)
| Papel | Emp. | Arquivo | MD5 |
|---|---|---|---|
| Inventário 2018 (Drive) | 03 | 14. 1. ASSOCIAÇÃO… | 5a33775dab86c8f64774a025edff5727 |
| Inventário 2018 (Drive) | 04 | 11. 1. CEVAL… | be83069f0748c45cac8cdd77d1359886 |
| Inventário 2018 (Drive) | 07 | 9. 1. RÁDIO… | 5ebb34fab4a7a9279ced3cd62c7dd086 |
| Inventário 2018 (Drive) | 09 | 8. 1. TURISMO… | d1ba5895e128656b2bb3199b088a9af7 |
| Inventário 2018 (Drive) | 10 | 7 .1. SIVIERO TRANSPORTES… | c30b60cf4453422164d8a789e6886529 |
| Doc 23 de 2018 (Drive) | 09/10 | 23. 1. … D Í V I D A S … | 10183078378f5054916272d39700de21 |
| Doc 27 de 2018 | — | 27. 1. PLANILHA … RESUMO | af5ec6c22eda832bfa9b6b00f9082ebf |
| Inventário 2025 atual | 03/04/07/09/10 | 03./04./07./09./10. INVENTÁRIO… | 51551f116eb30e5ae40b96546d59dee2 / 39e3a75ec6f3360434f3228910ec55e7 / 2fdea6114549f6d1c75132b9070c5195 / fedbb88ae14b0ec278ae8e6321f91426 / 1fcc0230c50a4f0a5bafded7b9862b0e |
| Inventário 2025 antigo | 03/04/07/09/10 | ANTIGO … | 53aa5e49990365a5349c60e1cfa1d979 / d64cd0ff63c29286f4a9d129195f7fdc / 5d49c8f932a70a19934f35dcca111fc5 / 52379cc2d8f65dd7d5d914465dc357be / 3a2f7908ced9b2ba77bf6e6a640946ae |
| Doc 00 atual / antigo | — | 00. / ANTIGO 00. | a1670b10feadf7d4e38158915cc5982f / 8ed9d82a32d83dc7257f4ffa122a192a |
| Rafael PDF | 03/04/07/09/10 | ASSOCIAÇÃO… / CEVAL… / RADIO… / TURISMO… / SIVIERO TRANSPORTES… | e3ae5f5f070b1afd53995cecc2c68b38 / 1e54cbbea48b0dcd8aa8841598c4eba9 / b22568ded88b60aad547cc37847e9c28 / df2116a02f249f520d99b453a8ed85e5 / 56b624935a8dda1f5d62c3433fdf523c |
| Rafael Diretrizes / Introdução | — | | aff2b08f2ab247000f99288cd7f63d36 / 01a85c04c097f097d7f9c255aee0ae56 |
| Lista Valter | — | ativos-avaliados-valter.xlsx | 935282c1aa176ba914f5003d6012fddb |
| Planilha Google 2_0 / emp. e Rafael | — | | 11a81ef0aa1c897a35bd2aee50d43587 / c499486a2622eb81325b2201034adc75 |
| Observações planilhas Google | — | 02_trabalho/observacoes-planilhas-google.xlsx | 18bf9de495ef3e8512f19d88023c2bec |
| Respostas Dr. Pablo arq. 1 (texto, 32 p.) | — | | 03968eb843c05a5853790a4219a12a1f |
| Arq. 5, 6, 10, 13, 14 (escaneados, sem texto) | — | | 9ff5a5dadc0c70aac98065c832fe4a41 / bfc853f6fc77f17d2345377289a623b1 / 23d9b1732df6a96020aff24abd14c917 / f43d974b6adfcdbb0ffb3c00eb9c2333 / 7a9f21444770f76505aa0a33499c8196 |
| Apoio | — | conferencia-discrepancias.md / inventario-2025-avaliacoes-corrigido.xlsx / diff-inventario… / parte6-candidatos-duplicidade.md / achados-para-empresas.md | bdd7c87e7013b4213c58cb43126ac841 / a565450170c2889fc35d3757d8b34f5c / 7afde86365fd3f68703d5a25f0615e7f / 32ee3e62a48f34f421c25b1613ad5a7e / ca3c0b913bf4b53bf421efcd7b561d37 |

### Contagens
- Itens com valor: 2018 → 03: 0, 04: 4, 07: 6, 09: 0, 10: 7. 2025 atual = antigo → 03: 1 linha "-.-" R$ 0, 04: 3, 07: 7, 09: 3 linhas "-.-", 10: 2 linhas "-.-".
- Rafael (itens por seção): 04 e 07 com 3 seções completas (apresentado pelo grupo, avaliação executada por Valter com faixas, balanço executado por Valter); 03, 09, 10 só "apresentado pelo grupo" (R$ 0) + texto de baixa.
- Totais Rafael: 04 R$ 385.000,00 (p. 5); 07 R$ 2.912.000,00 (p. 10). Diretrizes: só a Toyota (p. 7). Lista Valter: só a Toyota (linha 10).
- 63 conferências; 57 "sim"; 6 "NÃO", todas reais nas fontes (não erro de extração):
  1. 03/2018: total R$ 10.000,00 sem nenhum item preenchido (5 linhas pontilhadas).
  2. 09/2018: tabela de bens e total em branco; doc 27 dá R$ 0,00.
  3. 09/2018 × doc 27: idem.
  4. 03 planilha 2_0 GERAL: "2018 apresentado" R$ 0 × doc 27 R$ 10.000.
  5. 07 planilha 2_0 GERAL: "2018 apresentado" R$ 2.338.000 × doc 27 R$ 9.316.016 (planilha tira o terreno da antena, mat. 3.751, R$ 7.148.016, e põe a Toyota R$ 20.000 e móveis R$ 150.000 na Rádio).
  6. 10 planilha 2_0 GERAL: R$ 908.000 × doc 27 R$ 928.000 (Toyota R$ 20.000 tirada da Siviero Transportes).
- Siviero Transportes 2018: bruto R$ 2.443.100,00 − R$ 1.515.100,00 (3 ônibus vendidos a prazo) = R$ 928.000,00, confere com doc 27.
- Antigo × atual: itens e valores iguais nas 5 empresas; mudam só a tabela de rateio (09 e 10: Rovílio 1 % sai, 16,50 % → 16,66 %) e o parágrafo da venda das quotas de Rovílio.
- Planilha Google: nenhuma observação do Luís nas linhas destas 5 empresas (2_0 e emp. e Rafael). Aba 7.Rádio tem nas linhas 2–33 uma cópia do bloco do Hotel Araçá.

### Fatos registrados na leitura (para as etapas seguintes; conferir de novo antes de usar)
- Rafael Rádio: avaliação executada por Valter repete R$ 210.000 da sala (p. 2) com faixa R$ 380.000–520.000 (p. 3); no balanço executado põe R$ 420.000 (p. 9). Fox: R$ 28.000 na avaliação (p. 6), faixa R$ 32.200–35.500, R$ 32.000 no balanço (p. 9). Toyota: faixa R$ 45.000–120.000 (p. 7), R$ 70.000 nas duas. Itens 02 a 05: "da concordância". Quadro-resumo (p. 10–11): 1/6 de R$ 2.912.000 = R$ 485.333,33 (sem excluir Telmo de Toni); selecionado R$ 70.000; crédito R$ 415.333,33.
- Rafael Ceval: 320D R$ 200.000 → R$ 240.000 (faixa 240–295 mil, p. 2); retro Case R$ 70.000 → R$ 100.000 (faixa 85–130 mil, p. 3); LK 1313 R$ 45.000 → R$ 45.000 no balanço apesar da faixa 90–150 mil (p. 4). Quadro-resumo p. 6: "0,5 bens disponíveis", 1/6 = R$ 64.166,00, texto fala em "0,16 bens divisíveis"; "ausência de interesse".
- Rafael 03: PDF transcreve o parágrafo de venda das quotas de Rovílio, que não está no inventário 03 do grupo (nem no antigo).
- Rafael 10: PDF não transcreve o parágrafo do processo do ônibus IXO 6000 (acordo R$ 115.000) que está no inventário 10.
- Ceval 2025 não menciona a dívida de 2018 com Celito Antônio Pierozan (NP R$ 150.000); doc 23 (2011) traz Celito R$ 25.000 como credor das transportadoras.
- Associação: capital 2018 R$ 100.000 (6 × 16.666,66); 2025 R$ 10.000 (6 × 1.666,66). Turismo 2018: quadro "não real" Geni 98 % + André Luiz 2 %. Siviero Transportes 2018: "não real" Adélia 98 % + Giovani 2 %.
- Achado E7 da análise ACERTOS (NF 4946, "66,66 Siviero / 33,33 Celito") está no pacote 2 daquela sessão, **sem aprovação registrada** (achados-para-empresas.md: só E1–E3 aprovados). Não usar até o Luís aprovar; perguntar.
- `rovilio-quotas.md` da sessão S3: ainda não existe (pasta 01-08 sem o arquivo). Pendência.

### Revisão independente (P11)
- Revisor disparado sobre a Fase 0 (fontes + script + saídas, sem conclusões). Saída esperada: `_diario/revisao-fase0.md`. Não apresentar a Fase 0 ao Luís antes do resultado.

## 2026-10-28 — aviso da coordenação: Tesseract OCR
- Tesseract 5.4 em `C:\Program Files\Tesseract-OCR\tesseract.exe` (fora do PATH), só `eng` e `osd`. Uso: localizar trechos nos arq. 5, 6, 10, 13, 14; valores, datas e matrículas sempre conferidos lendo a página como imagem.

## 2026-10-28 — revisão independente da Fase 0 e correções
- Relatório: `_diario/revisao-fase0.md`. 3 altas, 14 médias, 5 baixas. Nenhum valor de item de inventário errado; todos os "NÃO" anteriores eram reais.
- Conferido por mim nas fontes antes de corrigir: doc 18 (2018 SEI) tab. 5 itens 02 (320D 2009, R$ 250.000) e 06 (CASE 580L 2001, placa IKI 4827, R$ 60.000); tab. 28 (SEI pagou R$ 762.000 de dívidas e R$ 180.000 de combustíveis da Siviero Transportes); tab. 31 (terreno mat. 3.751, R$ 7.148.016, "relacionado no inventário da Rádio"); doc 11 (2025 SEI) máquinas item 03 (320D "– SEI", R$ 200.000) e item 07 (CASE 580L, IKI 4827, R$ 60.000); doc 25 par. 143–145 (Ceval: Celito 33,33 % / Valter 66,67 %; em 2013 a SEI comprou a parte de Celito, saldo NP R$ 150.000). Confirmados.
- Correções no script (md5 513fd286fa34eb6bd35845913a6a2983), todas por regra, nenhuma por valor:
  1. Planilhas Google lidas pelo texto do cabeçalho (antes: posição; abas 4.Ceval/7.Rádio/10.Transportes tinham "dez/2024" e coluna auxiliar trocadas). Lidas também as abas da 2ª planilha (7.Rádio tem "Estudo Rafael" = R$ 70.000, só o bem escolhido).
  2. Rafael: coluna "critério" (faixa de mercado / concordância expressa / sem avaliação) e coluna com todas as referências R$ do bloco de cada item (box R$ 35–50 mil e m² R$ 6.000–7.500 na p. 3; FIPE Toyota R$ 15–30 mil; Fox por versão e anúncios; FIPE LK 1313 R$ 30.600). Corrigida a perda de valores com "R$" no fim da linha (320D: faltava R$ 240.000 na lista de referências). Bem escolhido (Toyota, p. 11) virou linha estruturada.
  3. Fontes novas: doc 18 de 2018, doc 25 de 2018, doc 22 (caminho longo, prefixo \?\), doc 11 de 2025 (SEI), inventário-estudo-rafael.xlsx, ACERTOS E ACORDOS.pdf. Aba nova "Outras fontes (cruzamentos)" (61 linhas).
  4. Observações das planilhas lidas direto das duas planilhas (não do extrato da Parte 7): 8 textos (H221/I70/K69/J70 "Em 2018 estava na relação 9.1. RÁDIO…"; H874/I723/K722/J723 Rovílio), nenhum está no extrato `observacoes-planilhas-google.xlsx` → lacuna do extrato da Parte 7 (arquivo da sessão SEI; só registrar).
  5. Conferências novas: quadro-resumo de Rafael; cotas do doc 00; somas do doc 23 (1.733.700; 550.000; 36.900; 52.660; 1.656.500 — todas conferem); reconciliação da planilha 2018 (07: 9.316.016 − 7.148.016 + 20.000 + 150.000 de móveis inexistentes em 2018 = 2.338.000; 10: 928.000 − 20.000 = 908.000); sócios e notas antigo × atual; transcrição de Rafael com contagem e notas.
  6. Itens: coluna "valor citado na descrição" com alerta (ST 2018 item 05 IXO 6000: descrição 470.000 × coluna 585.000; item 07 IOM 9093: 450.000 × 500.000). Aba "Valores em notas" (25 linhas).
- Saídas: `_extracao/fase0-extracao.xlsx` (md5 8a54490814e39c6b74a340337fc6d21d), `_extracao/fase0-extracao.json` (md5 0b6d77ecfe821f2e8c737221ad444242). 47 fontes; 106 conferências; 94 sim, 10 NÃO, 2 "difere", todas reais nas fontes:
  - 03/2018 total 10.000 sem itens; 03 planilha zerou 2018; 03 PDF de Rafael omite "baixada em 09/07/2020" e acrescenta o parágrafo de Rovílio (que não está no inventário 03).
  - 04 Rafael: 1/6 = 64.166,00 (exato 64.166,67); "0,5 bens disponíveis" × "0,16 bens divisíveis" no mesmo quadro.
  - 09/2018 sem total (2 linhas).
  - 10 PDF de Rafael omite o parágrafo do processo IXO 6000 (acordo R$ 115.000).
  - 09 e 10 antigo × atual "difere": saída de Rovílio (1 %) da tabela de rateio + parágrafo da venda de quotas.
- Candidatos novos à Parte 6 (a registrar em `parte6-candidatos.md` depois do OK do Luís): retroescavadeira CASE 580L 2001 (Ceval IKI 4325 × SEI IKI 4827: placas diferentes, conferir chassi); NP Celito R$ 150.000 em 2018 no inventário da Ceval e no passivo da SEI (doc 18 tab. 23 item 07, "valor emprestado para a empresa Siviero"); SEI pagando dívidas da Siviero Transportes (R$ 762.000 + R$ 180.000, forma 9).
- Não refiz revisão completa após as correções: as mudanças são de leitura/estrutura; os números dos itens não mudaram (conferido: mesmas somas por empresa e versão).

## 2026-10-28 — Fase 0 aprovada pelo Luís ("Aprovo, siga"); planilha-mestre
- Criados: `log-decisoes.md` (Fase 0 aprovada; E7 continua sem resposta → não usado), `decisoes.json` (2 casamentos "proposto": 07-2018-02→07-2025-03 canal/equipamentos; 07-2018-03→07-2025-04 antena), `parte6-candidatos.md` (S4-1 a S4-6; S4-6 = dívida da Rádio de 1 lote no Colina do Retiro I, R$ 110.000, + lote permutado da sala, doc 25 par. 202).
- Ajuste no fase0_extracao.py (só a aba "Valores em notas"): fora as linhas "redução de 0,00 %"; dentro "PROCESSO JUDICIAL" sem valor (Operação Deus tá vendo, ST 2018). Conferências e itens sem mudança (106 conferências, 10 NÃO, 2 difere). md5: script 245e9851ebd3dd1b3a9cd750b906aa9b; json 9eecb6219ac91e13c23232cc7aa748da; xlsx d471ab1006ca638b85866ee722127b96.
- `_scripts/build_mestre.py` (md5 07d529f52cfc148f0bbcc6243eb5b9ee) → `planilha-mestre-03-04-07-09-10.xlsx` (md5 73cdcd6243d31c84e9319a21e2d6e9e3). Entradas: fase0-extracao.json, decisoes.json (md5 a526f3effad300a1596ba2db16aa3526), planilhas Google (só para a descrição completa das linhas com observação).
- Abas: Leia-me, Mestre (18 linhas), Créditos e dívidas (notas) (13 linhas), Resumo por empresa.
- Pares por chave: 07 mat. 11.382 (sala), mat. 11.324 (box), placa IRO2562 (Fox); placa IHW0285 (Toyota, ST 2018 item 04 → Rádio 2025 item 07); 04 NF 4946 (320D), chassi JHF0048344/placa IKI4325 (Case), chassi 34500312579515/placa IHM0033 (LK 1313); mat. 3.751 Rádio 2018 item 06 → SEI 2025 doc 11 tab. 7 item 02 (R$ 7.148.016). Sem chave (propostos): canal/equipamentos e antena. Sem par em 2018: 07-2025-02 móveis. Sem destino: 04-2018-03 (M.Benz 2423 K, ILD 3959, R$ 93.000); ST 2018 itens 01–03 (ônibus IUY 8068, IAZ 0281, ISS 7019); itens 05–07 vendidos a prazo em 2015/2016 (recebimento a verificar).
- Conferências do Resumo: soma grupo 2025 = total declarado nas 5; soma Rafael = total Rafael nas 5 (07: 2.912.000; 04: 385.000); soma 2018 = total declarado (07: 9.316.016; 04: 391.000; 10: 2.443.100 bruto); 03 NÃO (10.000 sem itens, real); 09 sem total.
- Situações: 07 Concorda 5, Discorda 2 (sala 210.000×420.000; Fox 28.000×32.000), Em 2025 fora destas empresas 1 (terreno 3.751); 04 Concorda 1 (LK 1313 45.000), Discorda 2 (320D 200.000×240.000; Case 70.000×100.000), só 2018 1; 10 só 2018 6.
- Revisor independente da planilha-mestre: disparado. Saída esperada `_diario/revisao-mestre.md`.

## 2026-10-28 — aviso da coordenação: OCR dos arquivos 2–16 e anotação na NF 4946
- OCR revisado (sessão SEI, só leitura): `02_trabalho/ocr-respostas-docs/` (um .md por arquivo; `ocr.jsonl` md5 começa 9b5b52b7c1 — conferido; LEIA-ME.md); revisão em `02_trabalho/_diario/revisao-ocr-respostas-docs.md`. Uso: localizar trechos; números com 3–4 % de erro → conferir na imagem e citar arquivo/página. NFs (arq. 5 e 6): ler só a imagem.
- Conferido por mim na imagem (arq. 6, p. 1, renderizada a 150 dpi no scratchpad): NF nº 0004946, Paraná Equipamentos S.A., emissão 12/11/2009, destinatário CEVAL TERRAPLANAGEM ESCAVACOES LTDA, CNPJ 04.129.392/0001-44; escavadeira hidráulica Caterpillar 320D, motor CAT A6F00902, nº de série CAT0320DAA6F00902; valor total R$ 470.000,00; alienação fiduciária a Caterpillar Financial; pagamento "à vista R$ 141.000,00 parc. cliente + R$ 329.000,00 parc. FINAME". Anotação manuscrita no campo "reservado ao fisco": "Maquina do Celito. Ceval" e "66,66 Siviero / 33,33 Celito". Confirmado.
- Chave nova para o candidato S4-2: nº de série CAT0320DAA6F00902 (os inventários da SEI 2018/2025 não trazem série da 320D deles).
- Situação do E7: o Luís pediu à coordenação que repassasse o achado; a leitura acima é da fonte primária, feita por mim. Ainda pergunto ao Luís como tratar a anotação (copropriedade 1/3 de Celito × rateio), antes de qualquer valor.

## 2026-10-28 — revisão independente da planilha-mestre e correções
- Relatório: `_diario/revisao-mestre.md`. 1 alta, 11 médias, 9 baixas. Nenhum valor, par por chave ou página errados; todos os 17 bens de 2018 e 10 de 2025 presentes uma vez.
- Conferido por mim antes de corrigir: doc 25 par. 202–204 (antena "num imóvel que pertence à empresa Siviero Empreendimentos"), 222–230 (lote industrial de Farroupilha "agregado ao capital" da ST, R$ 1.300.000 de mercado, negociado por R$ 875.000; 3 ônibus Millenium R$ 1.427.000, diferença R$ 552.000; GUIBOR 18 × R$ 8.000; a receber 24 × R$ 10.000 + R$ 100.000); arq. 1 p. 13 (8 % = transporte da antena, terreno, ITBI e emolumentos e compra das quotas de Telmo de Toni). Confirmados.
- Correções (todas por regra; nenhum valor de fonte mudou):
  1. Par proposto 1 → 2: 07-2018-02 (R$ 1.700.000, incluía móveis) → 07-2025-02 + 07-2025-03 (R$ 2.050.000); decisoes.json atualizado (continua "proposto").
  2. Terreno 3.751: situação "Bem de outra empresa listado em 2018", destino "sempre da SEI" com doc 18 tab. 31, doc 25 par. 204, Rádio 2025 par. 15, arq. 1 p. 12. Resumo: Rádio 2018 sem o terreno = R$ 2.168.000.
  3. Ônibus vendidos a prazo (ST 05–07): "Não como bem: já vendido (crédito a receber)"; Resumo com bruto 2.443.100 − 1.515.100 = 928.000 = doc 27. Destinos citam doc 25 par. 229–230 e 249 e a nota 10-2025-N22 (IXO 6000).
  4. Coluna "Alerta da fonte" (texto × coluna: IXO 6000, IOM 9093, sala, terreno).
  5. Toyota: destino "identidade confirmada pela placa; transferência ST → Rádio sem documento — pendente"; registro de data-base "verificar titular em 31.12.2024".
  6. Critério de Rafael: valor do balanço fora da faixa principal → dito na coluna (LK 1313: manteve o grupo; Fox: R$ 32.000 < R$ 32.200, dentro dos anúncios R$ 31.000–39.000). Nova coluna "Outras referências de Rafael".
  7. Registro de data-base preenchido ("Nenhum ajuste registrado" ou ponto a verificar), explicado no Leia-me.
  8. Notas: entram custo de mudar a antena (Rádio 2025 par. 15), subsídio da Rádio à Associação, ST sem pagar despesas/salários, venda das quotas de Rovílio e rescisão (09/10 2025), ônibus IUY 8068 financiado (doc 25 par. 227). Valores em formato brasileiro. 23 linhas.
  9. Linha 03-2018-TOTAL (R$ 10.000 sem itens, destino pendente).
  10. Resumo: Confere Rafael = "sem balanço próprio" em 03/09/10; colunas Doc 27 e Doc 00 conferidas; contagem de linhas por empresa de 2018 e de 2025.
  11. Parte 6: novos S4-7 (lote da SEI → ônibus da ST), S4-8 (Toyota ST → Rádio), S4-9 (sala + box no estudo de Rafael), S4-10 (parcela de Telmo de Toni descontada no doc 00 e dentro dos 8 %); textos de S4-1, S4-2 (NF conferida), S4-5 (divergência 269.300 × 762.000 + 180.000; credores do doc 23), S4-6 (lotes do Colina do Retiro na lista C0) refeitos. Vínculos por ID.
  12. Rótulos: "Só em 2018 (ausente em 2025)"; Ref. 2018 cita o arquivo de 2018 (arq. 9.1 etc.) e a tabela.
- md5: fase0_extracao.py 51565cf8effe9a27ffa2a7cf1c874414; fase0-extracao.json 19cfc9e7bfd956fdf3164d2abe5d8901; fase0-extracao.xlsx 2da3238d139f41d4d0b54363095defb7; build_mestre.py a036b2da642d35776eabcc447b8c03c9; decisoes.json 6dbe7c2c3d8ad468bd0e87982b8c7c79; parte6-candidatos.md d37de92669d4bb02a2fbcef68f933765; planilha-mestre 1b526f4946dbe39fa97b81e46efd7d71.
- Resumo final: 19 linhas no Mestre, 23 notas. Totais: 2025 grupo = inventário = doc 00 nas 5; Rafael 2.912.000 / 385.000; 2018 líquido = inventário = doc 27 nas 5 (Turismo 0).
- Não refiz revisão completa: as correções seguem os achados do revisor e não mudam valores; nova revisão de todas as decisões antes das fichas de fechamento (P11).

## 2026-10-28 — planilha-mestre e pares aprovados ("Aprovo os dois"); pacote 1
- decisoes.json: 07-2018-02 (→ 07-2025-02 + 03) e 07-2018-03 (→ 07-2025-04) "aprovado" (md5 196a1c8cc5e35a64e074992fc7ccfbd5). log-decisoes.md atualizado.
- build_mestre.py: rótulo do desmembramento segue a situação do par (md5 c7256d2d57c7af007d4c6882279af9c9). Planilha-mestre md5 b50578cd1d0d5bf73c1ef0881d1c4dbb.
- SEI: terreno 3.751 = SEI-2025-AURB-02, adotado pelo valor do grupo R$ 7.148.016 na regra R3 da SEI (planilha-mestre-sei.xlsx) → R3 deste pacote: sem valor aqui.
- _scripts/pacote01.py (md5 e7bba3b365bb49bae32375bb2bbeebce) → pacotes/pacote-01-tabelas.md (md5 a40656aa35d0785cc3352ae477e5d18d). R1: 6 bens, R$ 2.505.000 (Rádio 2.460.000; Ceval 45.000); R2: 03/09/10 R$ 0; R3: 07-2018-06. Exceções: 4 (grupo 508.000; Rafael 792.000). Destino: 8 linhas (R$ 2.526.100 em 2018). 19/19 linhas enquadradas. Conferido por script: 10 bens 2025 = R$ 3.013.000; teto LK 1313 → efeito no 1/6 = R$ 17.500.
- pacotes/pacote-01.md (md5 5359eeb2e591b5761c7e54f4df456ea4). Revisor independente disparado; saída _diario/revisao-pacote-01.md. Não apresentar antes.

## 2026-10-28 — aviso da S3: rovilio-quotas.md aprovado
- `02_trabalho/empresas/01-08-abastecedora-transportadora/rovilio-quotas.md` md5 c6ace6758de40dd6be77c71919485dc1; linha 3 do arquivo e log da S3 registram "Aprovado pelo Luís em 28/10/2026; liberado para as sessões SEI, S4 e S1". Lidas as seções 3, 4, 7, 9 e 10.
- Para 09 e 10: Rovílio 1 % em cada (contrato p. 3–4 e 6; 2018 docs 8 e 7); rateio no contrato 16,50 % por sócio; as duas figuram como compradoras apesar de baixadas; preço único R$ 3.440.082,00 pago com bens da SEI (cl. 4ª), sem separação por empresa; cl. 3ª inclui "créditos hereditários e/ou creditórios" em nome das compradoras (exemplo: terra de São Brás). Benchmark 2018 da S3 usa 1 % × R$ 928.000 = R$ 9.280 da Siviero Transportes.
- Criado `registro-data-base.md` desta sessão: S4-DB1 (venda das quotas, sem efeito em número), S4-DB2 (acordo IXO 6000, pendente), S4-DB3 (outorga/renovação, pendente), S4-DB4 (termo "outubro/2025", preventivo).
- Planilha-mestre não regenerada agora (revisor do pacote 1 em curso); a coluna "Registro de data-base" e as notas 09/10 recebem as referências S4-DB* na próxima regeneração.
- Pergunta ao grupo sobre 09/10 como compradoras: usar a nº 7 da S3 (não duplicar).

## 2026-10-28 — revisão independente do pacote 1 e correções
- Relatório: _diario/revisao-pacote-01.md. 1 alta, 8 médias, 8 baixas. Valores da R1, páginas de Rafael, totais e enquadramento confirmados pelo revisor.
- Conferido por mim: arq. 14 p. 1 (imagem): licença Anatel emitida 13/01/2022, FM 96,1 MHz, classe A3, validade da radiofrequência 08/04/2027, estação na "margem esq. da RS 470 – saída Nova Prata", estúdio Av. Osvaldo Aranha 1395 sala 311. Arq. 13 (OCR, datas): sala R.4 compra e venda, escritura 15/03/2013, registro 23/04/2013; box R.6 arrematação em 11/07/2013, registro e certidão 18/01/2016; box-garagem no subsolo, 17,28 m².
- Correções: R2 só bens + créditos/processos nomeados fora da regra (alta); LK 1313 fora da R1 → pergunta ao Luís no pacote 2 (regra no script: concordância com valor de Rafael fora da própria faixa não entra na R1); concordância de Rafael descrita como condicionada (p. 4–5, p. 12); ressalvas de titularidade (box, Toyota); C1 da Toyota; citação do canal com "no passado"; outorga vigente em 31.12.2024; destino separado (bens R$ 1.011.000; ônibus vendidos R$ 1.515.100); 23 notas citadas; valores brutos; S4-DB4; "declarado como da SEI"; páginas.
- registro-data-base.md: S4-DB2 com justificativa (parcelas depois da baixa da ST; credor indefinido); S4-DB3 com licença vigente (md5 73af014cd3a3a3f45b2e32f84d3392bb).
- build_mestre.py (md5 8013851e2ea9c2a56749a1097b648702): coluna de data-base por ID (07-2025-03 → S4-DB3; 10-2018-05 → S4-DB2) e nas notas (S4-DB1/DB2); resumo sem contar a Toyota na ST. Planilha-mestre md5 002fc0edf0cb3e08397b5d41919f6f3b.
- pacote01.py md5 605ffae3a7b9f80ec3bd0a4a95c1ff75 → pacote-01-tabelas.md md5 e83d2bde744fa00608bd831daa0382e8. R1: 5 bens, R$ 2.460.000. pacote-01.md md5 ace1ddbd92b4c450fb1802261e73ee59.

## 2026-10-28 — pacote 1 aprovado ("Aprovo as três"); pacote 2
- decisoes.json (md5 a6f74c9d7b0695f2fbf9128287e8a697): R1 em 07-2025-02/03/04/05/07 (valor adotado = valor comum); R3 em 07-2018-06 (sem valor aqui); regras_empresa.R2 (03/09/10). log-decisoes.md atualizado. Planilha-mestre regenerada (md5 4d8fdff1c06c4463d14f1691aca70cef).
- Conferido por mim (imagem): arq. 5 = NF 50.859, Case Brasil & Cia., 31/10/2001, destinatária Ceval CNPJ 04.129.392/0001-44, Case 580L 4x4 ano 2001, chassi JHF0048344, motor 46096951, total R$ 100.000,01 (R$ 95.238,10 + IPI R$ 4.761,91).
- Inventários 2025 (todos): nenhuma menção à NP de Celito (só "Pierozan" em Carmen Lourdes Pierozan Siviero, doc 00, e área comprada de Celito vendida em 2018, docx Hotel).
- _scripts/pacote02.py (md5 f625cd26d98fe7335cf2057d24980057) → pacotes/pacote-02-tabelas.md (md5 85fbe37ee1ce30f6b4b92ef8b0cd482e): exceções com C0 (sala 420.000; Fox 32.000; 320D 240.000; Case 100.000); totais com pacote 1 + C0: Rádio 2.912.000, Ceval 385.000; 8 %: 215.840 (grupo) / 232.960; Telmo: 1/6 × 16,50 % → 4.496,67 (grupo) / 4.853,33; doc 00 total 263.410.451,80 lido da fonte; 1/3 da 320D = 80.000 (13.333,33 no 1/6).
- pacotes/pacote-02.md (md5 d657779f3a5e4e65759849a3cd11aeb4): 8 itens. Recomendações: 8 % não aplicar; sala, 320D, Case, Fox por Rafael (C0); anotação NF 4946 não reduz a máquina; LK 1313 e Telmo sem recomendação (decisão do Luís). 3 perguntas ao grupo propostas.
- Revisor independente do pacote 2 disparado; saída _diario/revisao-pacote-02.md.

## 2026-10-28 — revisão independente do pacote 2 e correções
- Relatório: _diario/revisao-pacote-02.md. 1 alta, 9 médias, 6 baixas; nenhum valor, conta ou direção do C0 errado.
- Conferido por mim: arq. 1 p. 7 (texto): 320D e Case "equívoco", valores na data do inventário; Case R$ 95.238,10; FIPE out/2025 L-1313 3-Eixos R$ 41.201,00; "Nenhum dos bens guarda ônus financeiro, todos devidamente quitados". Doc 25 par. 146: máquinas e "03 caçambas" da Ceval prestam serviços para a SEI. Arq. 13 (OCR, linhas 48, 116): sala área útil 66,02 m², condominial 28,49 m²; compra pela Rádio por R$ 77.600,00. Diretrizes p. 15 (linha 868): "valor econômico líquido do ativo… passivos ou contingências que efetivamente incidam". Contas por script: 16,5 % × 232.960 = 38.438,40; 1 % × 2.912.000 = 29.120; 66,02 × 6.000/7.500 = 396.120/495.150; 420.000/66,02 = 6.361,71; 470.000/3 = 156.666,67.
- Correções em pacotes/pacote-02.md (md5 2bb3bedd22afe888191b02a6bb402120): item 4 separado em (4a) leitura própria e (4b) hipótese E7 (alta); item 1 com (b) "não é líquido", (g) condicional ao item 7, argumentos (d)–(f) e risco das Diretrizes; item 2 com matrícula e C1 por m²; item 3 recomendação manter R$ 45.000 (P5); item 5 par. 146 e valores do grupo; item 7 fatos neutros + precedente Rovílio; 2423 K fora do pacote (etapa 6); 7 perguntas propostas.

## 2026-10-28 — pacote 2 decidido
- Resposta do Luís: "Aprovo, decido 4b e 7" (escolhas de 4b e 7 não vieram) → perguntado por AskUserQuestion. Respostas: 4b "não é mais do Celito desde antes de 2018"; 7 "16,50 % (sem Telmo)"; perguntas ao grupo "Ainda não".
- decisoes.json (md5 8402f18b5be9c0e0ccdc5cf06446d43c): 07-2025-01 R$ 420.000; 07-2025-06 R$ 32.000; 04-2025-01 R$ 240.000; 04-2025-02 R$ 100.000; 04-2025-03 R$ 45.000; regras_empresa: 8 % não aplicar; NF 4946 4a e 4b; base 16,50 %. log-decisoes.md atualizado.
- Planilha-mestre md5 c1942a2e30e9119e8741b707386e14c9: Rádio 7/7 com decisão, soma adotada 2.912.000 (16,50 % = 480.480,00); Ceval 3/3, 385.000 (1/6 = 64.166,67).
- parte6-candidatos.md (md5 d6b3225f4e3249b2d32400a05532abe0): S4-2, S4-4 e S4-10 atualizados com as decisões.
- Próximo: etapa 6 (destino 2018 → 2025 e créditos/dívidas fora da tabela) — aguardando OK do Luís para começar.

## 2026-10-28 — etapa 6 (destino 2018 → 2025): busca e pacote 3
- OK do Luís ("Pode começar").
- _scripts/busca_destinos.py: 301 arquivos (Drive Cisao_Siviero + 01_fontes) + OCR revisado; 1 erro (arquivo de bloqueio ~$00 - RESUMO BELISE LTDA.xlsx, não é fonte); 900 achados por tema → _extracao/busca-destinos.json/.md.
- Fatos conferidos:
  - Contas pagas 2023–2024 (planilha Consolidado; relatórios mensais escaneados, jul/2023 p. 5 e jul/2024 p. 3 lidos na imagem — observações cortadas na própria fonte): ILD 3959 multa 04/07/2023 R$ 156,19; conserto elétrico 15/01/2024 R$ 120; IPVA 2024 23/01/2024 R$ 99,65 "M.BENZ 2003 - CEVA…" (pagos pela SEI). IHW 0285 IPVA 2024 07/02/2024 pago pela Transportadora. IKI 4827 IPVA 18/01/2024 e IKI 4325 IPVA 26/02/2024, ambos pela Transportadora. IHM 0033 IPVA 21/02/2024 pela SEI. Abastecedora → "SIVIERO TRANSPORTES LTDA" R$ 6.500 10/07/2023 "P. 28/28 - EMPRESTIMO CAIXA FEDERAL - RADIO C…"; SEI → ST R$ 328,49 24/07/2024 "CUSTAS AVALIACAO E VENDA IMOVEL PENHORA ALIJH…"; Mercato "SERVICOS MARCO/24-DIST…" 15/04/2024.
  - Doc 20 de 2018 (contrato de 01/12/2003, md5 8c3ebfb42b9cafdb21cadbe121508d95), par. 313–341: cessão gratuita das quotas da Rádio de Eliana (99 %) aos irmãos; Telmo de Toni continua com 1.000 quotas (1 %). Não muda a decisão do item 7 do pacote 2.
  - Doc 25 par. 201: sala comprada com permuta de lote do Colina do Retiro I "de propriedade da empresa Siviero Empreendimentos" → dívida de 2018 da Rádio é com a SEI.
  - Inventário 2015 da ST (Material antigo/2018_desatualizado, md5 7280791f09ddae5da4314f64663b4ebd): Operação Deus tá vendo — "passaram de credores/autores para RÉUS".
  - Nenhuma certidão de baixa (03, 09, 10) nas fontes.
- pacotes/pacote-03.md (md5 32697ce014ab155b96409e0deb3c32e2): 10 itens; 13 perguntas propostas; S4-3 (duas máquinas, indício forte), S4-6 (origem), S4-11 (novo, empréstimo Caixa "Rádio").
- Corrigido antes da revisão: "ônibus 2000" é número de frota (doc 23: Paradiso 1995), não a placa IJW 4230.
- Revisor independente do pacote 3 disparado; saída _diario/revisao-pacote-03.md.

## 2026-10-28 — revisão independente do pacote 3 e correções
- Relatório: _diario/revisao-pacote-03.md. 1 alta, 8 médias, 13 baixas; nenhum valor errado.
- Conferido por mim: doc 25 par. 211, 225–228, 237–238, 243–244; doc 22 par. 49, 50, 61; arq. 17 PDF p. 3 (26.02.2014, lote 04/1302, Alcides Basso Imóveis, 3 ônibus + R$ 552.000 em 60 parcelas "pagas pela Siviero" = R$ 1.427.000); arq. 17 xlsx: Colina do Retiro I "23 – saldo de 13 = 10", 10 saídas, nenhuma para a Rádio/Dal Pai; contas pagas: NF 20 CERAN R$ 126.590 (ago/2024), NF 22 CERAN R$ 286.940 (set/2024), NF 21 Renan Baronio (out/2024), folha Ceval ago/set 2024; 118 linhas de juros pagos pela SEI a Piccoli, André Luiz Siviero etc., nenhuma a Celito; ST 2025 quadro "não real": Adélia 98 % (R$ 303.800) + Giovani 2 % (R$ 6.200) = R$ 310.000.
- Correções: ICMS → destino declarado (recebido antes de 2018) (alta); IXO 6000 natureza como inferência + alternativa + sucessores formais; novo item Ceval 2023–2024 (bases R$ 446.244,50) e S4-12; dívidas ST 2018 com terceiros R$ 243.800; doc 18 tab. 28 citado por inteiro; empréstimo Caixa reescrito; arq. 17 no item 1/S4-7; lista única de 20 perguntas (pacotes 2 e 3, sem duplicatas); demais baixas.
- pacotes/pacote-03.md md5 d7cd5a2d1563f634af96a1d79c6e107b.

## 2026-10-28 — pacote 3 aprovado ("Aprovo")
- decisoes.json (md5 1762dffd8a31918ea6a7c0b6311d5689): 32 destinos (linhas de 2018 e as 23 notas); linhas só de 2018 com decisão "Sem bem em 2025 — ver destino". build_mestre.py (md5 263eb58464f24f1075b3ad1c78fa2a7b): "Declarado como da SEI". Planilha-mestre md5 22f8e2e0b2a39a3be0d297e1968e3067; nenhuma nota sem destino; coluna Decisão sem "Pendente.".
- parte6-candidatos.md reescrito (S4-1 a S4-12; md5 1fa07d878e6e718347790402fbb76dca). perguntas-ao-grupo.md criado (20 perguntas; md5 4ff0c457a916c0308261cb69566764a7). registro-data-base.md S4-DB2 atualizado (md5 b39317343951c486ab61d55d8ee46e11). log-decisoes.md atualizado (md5 c7eb1f959759dfa86f3d2d581205e674).
- Próximo: revisão independente de todas as decisões aprovadas (P11, antes das fichas de fechamento). Disparada; saída _diario/revisao-decisoes-final.md.

## 2026-10-28 — revisão independente final e fichas de fechamento
- Relatório: _diario/revisao-decisoes-final.md. 0 altas, 8 médias, 10 baixas; nenhum valor adotado errado, nenhuma decisão aplicada fora do aprovado; C0 conferido nas 10 linhas; totais e quotas conferem.
- Conferido por mim: planilha-mestre-sei.xlsx, aba Lotes linha 458: Industrial Vicentino 004/1302, R$ 1.355.000, "Lotes de terceiros passíveis de indenização", passivo adotado R$ 677.500 (R6 nível B). Contas pagas set/2024: "IMPOSTO SIMPLES S/NOTA N 6631 DA RADIO - RECE…" R$ 3.000 pago pela SEI.
- Correções: build_mestre.py (md5 122669eff70c7d9b984f236969db67b3): S4-DB2 sem "crédito em aberto"; critério "concordância condicionada"; alertas em formato brasileiro; observação do terreno "Respondida"; motivo na linha 03-2018-TOTAL; Leia-me (pares aprovados; "declarado como da SEI"); Resumo com Soma adotada, Base da quota e Quota; notas extras 04-2024-ATIV (S4-12) e 07-2023-CAIXA (S4-11). decisoes.json (md5 2a516822d32c1adc0ddc6655f19a1204): destinos 10-2018-06/07 "ônibus não identificado na fonte". Planilha md5 2e7a52ac0a9bf9a6d7331e596393105d (19 linhas, 25 notas). parte6-candidatos.md: S4-2 (indício de duas 320D; casamento deslocado na SEI), S4-7 (lote 004/1302 = passivo indenizável da SEI, R$ 677.500), S4-11 (leitura da S3). log-decisoes.md: notas de resolução do E7 e das perguntas. registro-data-base.md: S4-DB5 (FIPE out/2025 do LK 1313).
- Fichas de fechamento escritas: ficha-fechamento-07.md, -04.md, -10.md, -09.md, -03.md. Números: Rádio 2.912.000 / quota 16,50 % 480.480,00; Ceval 385.000 / 1/6 64.166,67; 03/09/10 R$ 0 em bens. Efeitos por script.
- Itens para a união com a SEI (P13): casamento deslocado das escavadeiras na planilha da SEI; lote 004/1302 passivo × cadeia da ST; resposta ao candidato nº 1 da SEI (3.751 só na SEI em 2025); leitura do empréstimo Caixa com a S3.
