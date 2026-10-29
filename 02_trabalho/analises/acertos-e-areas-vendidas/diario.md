# Diário — análise de ACERTOS E ACORDOS e das áreas vendidas

Só acréscimos (regra P12). Uso de Claude e do revisor. Brief: `00_prompts/4-analise-acertos-e-areas-vendidas.md`.

## 2026-10-28 — extração e cruzamentos

### Entradas (MD5)
- 5a6e856dcdf3836c4f2b23702de52d42 `01_fontes/inventario-2025/ACERTOS E ACORDOS.pdf` (6 p., sem camada de texto)
- 6cdf8807c9f929ea0f2b4090ba23f4b9 `01_fontes/inventario-2025/HOTEL PRINCESA DOS VALES ... SEM OBRAS.docx`
- 8c0f668319f9ce0973e1cecbe77f7c82 `01_fontes/planilhas-google/ACERTOS E ACORDOS (gdoc).docx`
- a565450170c2889fc35d3757d8b34f5c `01_fontes/outros/inventario-2025-avaliacoes-corrigido.xlsx`
- ca73deb381ee60dc2f39ff6e1cb47bbf arq. 16 (termos Susin/AGI, 9 p. escaneadas)
- 7ef0dc35dcc123bd1d356b0e1fa45164 arq. 17 (xlsx, 390 linhas)
- dcfb1bf9e1fa473b89da65023878fb0b arq. 18 (docx)
- 03968eb843c05a5853790a4219a12a1f arq. 1 (respostas ao Dr. Pablo, com texto)
- bfc853f6fc77f17d2345377289a623b1 arq. 6 (NF Ceval)
- 62ed740e2dca3c7ee6e041d4936ce330 `02_trabalho/sei-2025-itens.xlsx`
- dcabdf2209e918d39e1ef81f547c779e `02_trabalho/sei-2018-itens.xlsx`
- a98f32a4c686d083eec04d21903a40b9 `01_fontes/outros/Alexandre Susin - TERMO DE ACORDO.docx` (minuta de 04/02/2020; o termo assinado é o do arq. 16)
- 63f103712292821969f5ae87d581ebdc doc 02; 39bb7a345a4d373c8f885f7c726dccd7 doc 05
- ce58270ef14dd3b1f1d6cdaf3eeb056d Rafael, Caminhos do Sol

### Passos
1. `/pdf-to-excel` no ACERTOS: "no text layer found", nenhuma tabela, nenhum arquivo gravado. Páginas renderizadas a 200 dpi e lidas como imagem. Tudo legível.
2. Conferência linha a linha: imagem × gdoc × extração de 26/10. Sem divergência de número ou nome. O gdoc embaralha a numeração dos itens 10 a 32 (layout), mas o texto bate.
3. Totais (aba Conferência de `extracoes.xlsx`): 82 lotes, espécie 2.882.204,55, total 12.087.204,55, Boff 3.315.675,23, vendas 12.058.000,00, sacas 37.694 e 13.085, 411,43 ha e 411,4 ha: todos conferem. Arredondamentos: Rigo 190 × 49.726,30 = 9.447.997 (−3); Pellegrini 85,32 × 23.441,16 = 1.999.999,77.
4. Arq. 16 transcrito por subagente (9 páginas): Termo A, 13/07/2020, AGI + Susin, 17 lotes + R$ 200.000; cl. 1.1 famílias Piccoli, Boff, Andreis, Pedrotti, Wiltgen, Matioda/Girardi, Fiorese, Rizzo, Trentin/Camazzola, Pauletti, Zatera; cl. 6 quitação total; cl. 8 exclui intermediações novas. Termo B, 30/04/2020, 7 lotes (6 + 1 Pedrotti); lote 2.4a = Altos da Julieta Q1729 L12 (a AGI). Buriti não aparece. O docx em `outros` é minuta anterior (04/02/2020), com cláusulas diferentes.
5. Lotes: `match_lots.py` liga 77 linhas do arq. 17 a 15 grupos de itens do ACERTOS. Nenhum no estoque próprio 2025. Na tabela 47 (indenizáveis): Perini Q1469 L11 (Farina) e Jardim Itália Q1747 L02 (Ambitec). No arq. 18: Altos da Julieta Q1728 L03 (Verona/Rochedo) e Alto das Graças Q6369 L08 (Colunata/Basso).
6. Tabela 47 × arq. 17: Industrial Vicentino Q1301 L003 (venda 19/09/2025), Jardim Itália Q1747 L001 (10/03/2025), Primavera Q1199 L010 e L011 (dação à Safra, fev/2025, arq. 1 p. 26) estão como "lotes de terceiros" e fora do estoque próprio. O Reserva Paradiso Q104 L623, recebido na permuta de 2025, está no ativo do doc 11.
7. Doc 11 Obs. 15 (Susin, R$ 1.200.000): versão antiga com "......" (`diff-inventario-2025-antigo-x-atual.md`, linha 353).
8. Boff: inventário 2018, investimentos itens 10-12 = 1.858.000; ACERTOS p. 4-6 = 3.315.675,23; doc 11 = "R$ 00.000.000,00"; Rafael = 61 lotes × 141.000 = 8.601.000. Arq. 1 p. 30: os créditos serão usados "como moeda de postergação de contratos e compensação".
9. Ansolin: 55.000 + 156.666 + 375.000 = 586.666 = 7,0 ha do doc 11 área rural item 06 = arq. 17 linha 153.
10. São Crispin: doc 05 itens 05 + 06 = 407,53 + 4,0 = 411,53 ha × docx 411,43 ha (diferença 0,10 ha). O 491/83,47/407,53 é do próprio doc 05.
11. Escavadeira: só a SEI tem escavadeira de 2024 (busca nos docs 01 a 10). O arq. 6 é a 320D 2009 da Ceval (NF 4946, R$ 470.000, anotação à mão "Máquina do Celito... 66,66 Siviero / 33,33 Celito").

### Saídas
- `extracoes.xlsx` (MD5 926e164146bc35c9db3cd3b9398dbdd6) e `cruzamentos.xlsx` (MD5 6a8feccd2ccf91077f0bd793336c47e3), gerados por `_scripts/build_outputs.py`.

### Pendências
- Revisor independente: disparado sobre as duas planilhas, com as fontes.

## 2026-10-28 — revisor independente e correções

Revisor com contexto novo, sobre as duas planilhas, com as fontes. Transcrição confirmada sem erro. Achados do revisor e o que foi feito:
- ALTA 1. Chave de quadra "H - 6371" não casava com "6371" (match_lots.num). Corrigido. Resultado: Alto das Graças Q6371 L11 está no estoque do doc 11 (tabela 38 linha 16, R$ 141.000), no arq. 18 (2 lotes do Basso, R$ 260.000) e na lista de Rafael (ativos-avaliados-valter, Ativos, linha 101; conferido por mim). Nenhum lote dos acordos no estoque: confirmado de novo.
- ALTA 2. Termo B de Susin não é comissão: são lotes que a família Susin comprou de Bovo, C. Fioreze e L. Pedrotti. Os 6 pendentes da Obs. 15 = 5 do Alto das Graças + Pedrotti. Item 3 corrigido; os 6 lotes entraram na Composição do ACERTOS. Termo A: 14 lotes (15 equivalentes) contra 17.
- MÉDIA 3. Acrescentadas as cl. 5, 5.1 e 6 do Termo A e a minuta de 04/02/2020. Leitura: a comissão pelos contratos das famílias da cl. 1.1 está dentro do Termo A.
- MÉDIA 4. Soja Lorenzet 2024: o revisor apontou o Hotel como arrendador nos contratos 8 e 9. Conferido por mim, com OCR Tesseract da p. 1 do contrato 9: ARRENDADORA com CNPJ 89.471.270/0001-94 = Hotel (doc 05). Registrado como conflito com o docx.
- MÉDIA 5. Crédito Rigo de natureza indefinida; coincidência das 2.606 sacas com a Lorenzet 2021. Registrado.
- MÉDIA 6. 320D 2009 em dois inventários (04 e 11), com NF da Ceval e 1/3 de Celito. Registrado em S12/E7.
- MÉDIA 7. R$ 100.000 sem fonte para o lote Q1728 L03. Removido.
- MÉDIA 8. Lotes Q6371 L05, L06 e L12 trocados por 2,5 ha no Monte Bérico (arq. 18). Incluídos.
- MÉDIA 9. Leitura da resposta 19 do arq. 1 corrigida ("moeda de postergação", crédito tratado como usado). Acrescentada a permuta Via Nova de 2016 (2 % do Caminhos do Sol).
- MÉDIA 10. Disposições de 2025 fora do estoque (Barbiero Q05 L03 e L08; Falavigna Q174 L20, L21, L23; Q175 L12, L13; Q876 L16). Conferido por mim no sei-2025-itens: ausentes. Incluídas na aba Datas e em S6.
- BAIXA 11-15. Créditos x multas com critério único (só R$ escritos; Vale das Colinas 632.500). Itens 14 e 25 corrigidos. Basso como aproximação (186.902,68). Ansolin marcado como hipótese. Cabeçalho da p. 18 corrigido: é pergunta do lado consulente. Colombo como conflito. B4 = Prefeitura. Flávia como parte relacionada.
- Não incluído na extração: a lista "Observações / AFORA O TRABALHO DO DIA-A-DIA" da p. 3 (6 itens sem valor: retificações de áreas, negociações de contratos, escrituras de lotes antigos, acompanhamento de processos, viagens, "...").

Saídas finais: `extracoes.xlsx` (MD5 d3473bc5e69245765c99993f6187ec49), `cruzamentos.xlsx` (MD5 a6b3c816b70bd219a78373ee52cd9bf9), 76 lotes ligados aos acordos. Scripts em `_scripts/` (patch_review.py = correções do revisor).
Pacote 1 (10 itens) apresentado ao Luís. Aguardando aprovação.

## 2026-10-28 — pacote 1 aprovado
- Luís: "Aprovo o pacote 1". Registrado em log-decisoes.md; achados-para-sei (S1-S7, S9) e achados-para-empresas (E1-E3) marcados como aprovados.
- Próximo: pacote 2 (itens sem desconto direto), com revisor antes da apresentação.

## 2026-10-28 — nota da sessão 02/05 sobre E2
- Parcelas Pellegrini no inventário de 2018 somam 85,22 ha (não 85,32): a conciliação São Crispin fecha exata. Anotado em E2 e na aba São Crispin. cruzamentos.xlsx MD5 645f033e2a530e175ec9bde7348faea7.
- Nota da sessão SEI: Acordi (280.377,82 × ~150.000) e permutas do Lago Negro no arq. 17 só nas observações; incorporadas no rascunho do pacote 2.

## 2026-10-28 — revisor do pacote 2
- 2 altas, 6 médias, 6 baixas, todas corrigidas em pacote-2.md:
  - #1 reescrito: causa ambiental da tabela 47 × atraso de matrículas do acordo Farina; lotes do Vicentino nomeados no arq. 17; honorários Baruffi de 16/05/2022; R$ 2,75 mi marcado como cálculo.
  - #4: Indenização 03 inteira (L12 sem registro no arq. 17; L11 já no S5).
  - #6: dupla contagem como hipótese, conferir a série.
  - #7: "acordados", não "pagos".
  - #8: lotes consignados a Válter; área pode ser de Válter.
  - #9: Morada dos Ipês II = lotes prometidos a Rovílio.
  - #10: Termo B = entrega pendente; Lago Negro invertido; Q6371 L11 remete ao S4.
  - Flávia "Mônica" no arq. 18 (mesma pessoa, suposição).
- Pacote 2 apresentado ao Luís.
