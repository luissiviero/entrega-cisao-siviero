# Revisão independente da Fase 0: extração das empresas 03, 04, 07, 09 e 10

Revisor: agente independente, sem contexto anterior. Data: 28/10/2026.
Objeto: `_scripts/fase0_extracao.py`, `_extracao/fase0-extracao.xlsx` e `_extracao/fase0-extracao.json`.
Método: reli todas as fontes com scripts próprios, sem usar a saída do script revisado (python-docx célula por célula, pymupdf página por página, openpyxl com valores e com fórmulas, notas e comentários). Os scripts do revisor estão em `C:\Users\User\AppData\Local\Temp\claude\C--entrega-cisao-siviero\639aa2be-df0b-40be-bb0a-3182601981fe\scratchpad\revisor\`.

Gravidade: **alta** = número ou item errado ou faltando; **média** = conferência faltando ou rótulo enganoso; **baixa** = forma.

**Placar: 3 achados de gravidade alta, 14 de média e 5 de baixa.** Não achei nenhum valor de item de inventário errado: os problemas estão na planilha Google, nas seções secundárias de Rafael, nas fontes que faltam e nas conferências.

---

## Achados

### 1. Alta. Abas 4.Ceval, 7.Rádio e 10.Transportes: as colunas foram lidas trocadas
- **Evidência:** nas três abas do arquivo `2_0_Comparativo...xlsx` o cabeçalho é C "2018: apresentado", D "2018: atualizado", E "dezembro 2024", F (sem rótulo, coluna auxiliar), G "outubro 2025", H "Diferença", I "OBSERVAÇÕES". O script (`gsheet_rows`, linhas 517–519) grava a coluna F como "2025" e a coluna E como "Divergência". A coluna G ("outubro 2025") não é lida.
- **Efeitos na aba "Planilha 2_0" da extração:**
  - 7.Rádio, linha 36: aparece "2025 = 2.628.000". Na verdade é F36 `=SUM(F37:F42)`, uma soma auxiliar sem a Toyota, numa linha de rótulo.
  - 10.Transportes, linhas 5–7, e 4.Ceval, linha 7: a coluna "2025" sai vazia, mas G = 0 na fonte.
  - 4.Ceval, linha 3: a coluna "2025" mostra 315.000, que é a soma auxiliar F3.
  - A coluna "Divergência" dessas abas traz na verdade "dezembro 2024".
- **Correção:** escolher as colunas pelo texto do cabeçalho de cada aba, não pela posição. Gravar E como "dez/2024", F como "auxiliar (sem rótulo)" e G como "out/2025". Rodar de novo.

### 2. Alta. Rafael: valores das seções de avaliação ficaram de fora
A extração guarda só a primeira faixa de preço de cada item. Os valores abaixo estão no PDF e não aparecem em nenhuma aba:
- **Rádio, p. 3 (item 01, sala):**
  - m² privativo entre R$ 6.000 e R$ 7.500;
  - o box de garagem "agrega" entre R$ 35.000 e R$ 50.000. Essa faixa vale para o item 05 (Box), que ficou sem faixa na extração.
- **Rádio, p. 6 (item 06, Fox):**
  - valores por versão: R$ 32.217, R$ 35.052 e R$ 35.140;
  - preço de anúncio de usado entre R$ 31.000 e R$ 39.000.
- **Rádio, p. 7 (item 07, Toyota):** Tabela FIPE entre R$ 15.000 e R$ 30.000, além da faixa de R$ 45.000 a R$ 120.000, que foi extraída.
- **Ceval, p. 4 (item 03, caminhão LK 1313):** FIPE de cerca de R$ 30.600, além da faixa de R$ 90.000 a R$ 150.000, que foi extraída.

**Correção:** acrescentar em "Rafael itens" uma coluna "outras referências de valor" (valor, tipo e página), com todos os números `R$` de cada bloco do item na seção avaliacao_valter. A faixa do box (p. 3) vai para o item 05.

### 3. Alta. Fonte faltando: o inventário 2018 da Siviero Empreendimentos (doc 18) muda o total de 2018 da Rádio
- **Evidência:** arquivo `18. 1. SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS ... Dezembro de 2018.docx`, que está no Drive e em `01_fontes/inventario-2018`.
  - **Tabela 31:** o terreno da matrícula 3.751 (BR 470, antena), de R$ 7.148.016,00, é listado como bem da SEI "Relacionado no Inventário da empresa: RÁDIO". Esse valor é 77% do total da Rádio em 2018 (R$ 9.316.016).
  - **Tabela 5 (máquinas da SEI, 2018):** "Escavadeira Hidráulica, ano 2009, modelo 320D" por R$ 250.000 e "Retroescavadeira CASE 580L, ano 2001, PLACA IKI 4827" por R$ 60.000. São quase iguais aos itens 01 e 02 da Ceval (320D 2009/2010 por R$ 200.000; CASE 580L 2001, placa IKI 4325, por R$ 60.000).
  - **Tabela 28:** a SEI declara que pagou "dívidas da Siviero Transportes Ltda., de 2011 a 2015" (R$ 762.000) e "combustíveis da Siviero Transportes" (R$ 180.000).
- O arquivo não está na aba "Fontes" e nada dele foi extraído.
- **Correção:** incluir o arquivo na lista de fontes, com MD5, e extrair as tabelas 5, 28 e 31 com marcação de cruzamento para a Ceval (04), a Rádio (07) e a Siviero Transportes (10).

### 4. Média. Uma observação da planilha sobre um bem da Rádio não foi capturada
- **Evidência:** a frase "Em 2018 estava na relação '9.1. RÁDIO COMUNIDADE...'" está nas células abaixo, todas na linha do terreno da matrícula 3.751 (R$ 7.148.016):
  - planilha 2_0: aba GERAL, célula H221, e aba 11.Empreendimentos, célula I70;
  - planilha "empreendimentos e Rafael": aba GERAL, célula K69, e aba 11.Empreendimentos, célula J70.
- O script chegou a achar a linha 69 da segunda planilha (o JSON traz `gsheet2_rows_with_company_text: [69, 722]`), mas só procurou o texto no extrato `observacoes-planilhas-google.xlsx`, que não tem essas células.
- A aba "Observações Google" diz "(nenhuma observação nas linhas destas empresas)", o que engana.
- **Correção:** ler as colunas OBSERVAÇÕES direto das duas planilhas (não do extrato) em toda linha que cite as 5 empresas ou seus bens (matrículas 3.751, 11.382 e 11.324; placas). Registrar o texto. Avisar também que o extrato da Parte 7 tem essa lacuna.

### 5. Média. As abas por empresa da segunda planilha não foram lidas
- **Evidência:** "Comparativo ... empreendimentos e Rafael 2018 x 2025.xlsx", aba 7.Rádio. Essa aba tem uma coluna H a mais, "Estudo Rafael". H35 `=SUM(H36:H43)` = 70.000 e H43 = 70.000 (Toyota). Ou seja, o "Estudo Rafael" da Rádio na planilha é só o bem escolhido, não o balanço de Valter (R$ 2.912.000). Isso pode levar a erro e não foi registrado.
- Conferi as abas 4.Ceval e 10.Transportes: são idênticas nas duas planilhas.
- **Correção:** ler 7.Rádio da segunda planilha e anotar o que a coluna "Estudo Rafael" contém.

### 6. Média. Conferência "planilha 2018 apresentado × doc 27": a explicação genérica engana e esconde R$ 150.000
- **Evidência:** a observação "diferença esperada se a planilha reorganizou bens entre empresas" aparece nos três "NÃO". Reconciliando cada um:
  - **03:** a planilha (GERAL, linha 16) tem 0; o doc 27 tem 10.000. Não houve reorganização: a planilha zerou um total que no inventário de 2018 não tem itens.
  - **07:** 2.338.000 = 9.316.016 − 7.148.016 (terreno passado à SEI) + 20.000 (Toyota vinda da Siviero Transportes) + **150.000**. Esses R$ 150.000 são "Móveis, instalações e utensílios" (GERAL, linha 108, coluna C), item que **não existe no inventário 2018**: lá os móveis estavam dentro do item 02, de R$ 1.700.000.
  - **10:** 908.000 = 928.000 − 20.000 (Toyota passada para a Rádio). Confere.
- **Correção:** trocar a observação genérica pela reconciliação de cada empresa e marcar os R$ 150.000 da Rádio como divergência da planilha em relação ao inventário 2018.

### 7. Média. Siviero Transportes 2018: o valor da descrição não bate com a coluna de valor, e nada foi sinalizado
- **Evidência:** inventário 2018 da Siviero Transportes, tabela 3.
  - Item 05 (IXO 6000): a descrição diz "valor comercial de R$ 470.000,00"; a coluna de valor tem R$ 585.000,00.
  - Item 07 (IOM 9093): a descrição diz R$ 450.000,00; a coluna de valor tem R$ 500.000,00.
  - Item 06 bate (R$ 430.100).
- A extração usa a coluna de valor, o que está correto: a soma bate com 2.443.100 e a dedução de 1.515.100 com 928.000. Mas a divergência dentro do documento não aparece. O doc 25 de 2018 explica que os R$ 585.000 são o preço de venda a prazo ("foi vendido um ônibus... pelo valor de R$ 585.000,00, à prazo").
- **Correção:** acrescentar a coluna "valor citado na descrição" e um alerta quando ela for diferente da coluna de valor.

### 8. Média. Valores fora da tabela de bens ficaram só como texto solto
- **Evidência:** aparecem apenas como texto bruto na aba "Textos dos inventários", sem valor estruturado nem conferência:
  - **Siviero Transportes 2025:** acordo de R$ 115.000 pela venda do ônibus IXO 6000 (R$ 20.000 de despesas e parcelas de R$ 5.000 desde nov/2024). É um crédito, e o total da empresa é 0.
  - **Ceval 2018:** nota promissória de R$ 150.000 com Celito Pierozan.
  - **Rádio 2018:** "DÍVIDA: 01 lote no Colina do Retiro I = R$ 110.000".
  - **Turismo 2018:** dívida de cerca de R$ 77.000 com a Receita e ônibus vendido por R$ 95.000.
  - **Siviero Transportes 2018:** dívida de R$ 7.000 com a Receita; R$ 1.733.700 + R$ 269.300 = R$ 2.003.000.
  - **Rádio 2025:** redução estimada de 8% por regularização.
- Quem ler a conferência "2025 atual: soma dos itens 0 = 0, sim" da Siviero Transportes pode concluir que a empresa não tem nada.
- **Correção:** criar a aba "Valores em notas" (empresa, versão, parágrafo, valor, natureza: crédito, dívida, venda ou redução %).

### 9. Média. O quadro-resumo de Rafael não foi conferido
- **Evidência:**
  - **Ceval, p. 6:** 1/6 de 385.000 = 64.166,67; o PDF traz R$ 64.166,00 (truncado). O mesmo quadro diz "0,5 bens disponíveis" (em azul) e o parágrafo seguinte diz "0,16 bens divisíveis": os dois números se contradizem.
  - **Rádio, p. 10–11:** 2.912.000 / 6 = 485.333,33 (confere); 485.333,33 − 70.000 = 415.333,33 (confere); "1,16 bens".
  - A conferência "total do balanço de Valter = patrimônio total do quadro" também não foi feita. Ela confere nas duas empresas.
- **Correção:** acrescentar essas conferências em "Conferências" e anotar a contradição da Ceval.

### 10. Média. O bem escolhido (Toyota) não é um item estruturado em "Rafael itens"
- **Evidência:** na seção bens_escolhidos (Rádio, p. 11) a Toyota só aparece como linha bruta "R$ 70.000,00" (p. 10 e p. 11) na aba "Rafael totais e seções", sem descrição nem placa. O motivo: a expressão `ITEM_LINE` exige um número de item, e essa seção não tem número.
- O quadro (p. 10) chama a Toyota de "Conjunto de lotes escolhidos (07.)", um rótulo errado de Rafael que não foi registrado.
- **Correção:** criar uma linha bens_escolhidos com descrição, placa IHW 0285, R$ 70.000 e p. 11, e anotar o rótulo "lotes" do quadro.

### 11. Média. A "concordância" de Valter com os valores do grupo não foi registrada
- **Evidência:** Rádio, p. 4–5. Os itens 02 (móveis), 03 (equipamentos e canal), 04 (antena) e 05 (box) têm a cláusula "DA CONCORDÂNCIA COM OS VALORES DE AVALIAÇÃO PATRIMONIAL" no lugar de uma faixa de mercado. Na extração esses itens aparecem só sem faixa, como se faltasse dado.
- **Correção:** acrescentar a coluna "critério de Rafael" (faixa de mercado / concordância expressa / sem avaliação).

### 12. Média. A conferência "transcrição do grupo = inventário 2025" é fraca e deixou passar uma omissão
- **Evidência:** o script usa `zip` (linhas 763–774). Assim não confere se as duas listas têm o mesmo número de itens, nem as descrições e as notas. Diferenças reais entre o PDF de Rafael e o inventário:
  - **Siviero Transportes:** o PDF **omite** o parágrafo do processo do ônibus IXO 6000 (acordo de R$ 115.000).
  - **Associação:** o PDF omite "EMPRESA BAIXADA EM 09 DE JULHO DE 2020" e acrescenta o parágrafo do Rovílio, que não está no inventário 03.
- **Correção:** conferir o número de itens, os valores e as notas com valor, e mudar o rótulo para "valores dos itens iguais".

### 13. Média. Doc 23 (dívidas de 2011): só o hash, sem extração nem conferência
- **Evidência:** o doc 23 está em "Fontes", mas nenhum valor foi extraído. Refiz as somas:
  - 1ª lista (16/09/2011): 1.733.700 (confere);
  - 2ª lista (31/08/2011): 1.656.500 (confere);
  - "Capital em Veículos": 550.000 (confere; inclui "Toyota R$ 40.000");
  - receitas mensais: 36.900 (confere);
  - despesas mensais: 52.660 (confere).
- **Correção:** extrair o doc 23 e registrar essas conferências. A Toyota a R$ 40.000 em 2011 serve de histórico do bem escolhido.

### 14. Média. Fonte faltando: doc 25 de 2018 (divisão de lotes), com valores destas empresas
- **Evidência:** `25. 1. PLANILHA = DIVISÃO DE LOTES=ANO 2018 ...docx`, que está no Drive e em `01_fontes/inventario-2018`.
  - **Siviero Transportes:** "ainda tem a receber: 24 parcelas de R$ 10.000,00... e mais R$ 100.000,00, do Rodrigo"; capital de cerca de R$ 1.000.000; 3 ônibus da Milenium por R$ 1.427.000; venda por R$ 585.000.
  - **Ceval:** saldo de R$ 150.000 com Celito.
  - **Rádio:** sala comprada por R$ 100.000 (lote permutado por R$ 70.000 + R$ 30.000).
  - **Turismo:** ônibus vendido por R$ 95.000; dívida de R$ 77.000.
- **Correção:** incluir em "Fontes" e extrair os trechos das 5 empresas.

### 15. Média. Fonte faltando: inventário 2025 da Siviero Empreendimentos (doc 11), com possível dupla contagem da Ceval
- **Evidência:** na planilha 2_0, aba GERAL, linhas 185 e 189 (que vêm do doc 11):
  - "Escavadeira Hidráulica, ano 2009, modelo 320D, CATERPILLAR – SEI", R$ 200.000;
  - "Retroescavadeira CASE 580L, ano 2001, PLACA IKI 4827", R$ 60.000.
- A escavadeira já é o candidato nº 4 de `parte6-candidatos-duplicidade.md`. **A retroescavadeira CASE 580L de 2001 não está na lista** (Ceval: placa IKI 4325, R$ 70.000; SEI: placa IKI 4827, R$ 60.000).
- **Correção:** incluir o doc 11 em "Fontes" (o SEI é a outra ponta das duplicidades) e propor a retroescavadeira como candidato à Parte 6.

### 16. Média. Doc 00: as cotas por sócio não foram conferidas
- **Evidência:** a tabela 3 do doc 00 atual traz:
  - Ceval: "cota... para cada sócio: 16,66% = R$ 52.500,00";
  - Rádio: 16,50% = R$ 445.170,00 e Telmo 1,00% = R$ 26.980,00.
- Esses valores batem com as tabelas "valor real de cada sócio" dos inventários (conferi), mas a conferência não existe na extração.
- **Correção:** acrescentar a conferência.

### 17. Média. O rótulo "2025 antigo × atual: itens e valores iguais = sim" esconde mudanças nas tabelas de sócios e nas notas
- **Evidência:** comparei os docx inteiros.
  - **09 e 10:** a versão antiga traz a tabela de sócios com 16,50% × 6 + Rovílio 1,00%; a atual traz 16,66% × 6, sem Rovílio, e acrescenta o parágrafo da venda das cotas do Rovílio (28/08/2025).
  - **04:** só muda o espaçamento. **03 e 07:** não há diferença de conteúdo.
- A tabela de sócios da versão antiga também não foi conferida contra o total.
- **Correção:** deixar o rótulo como "itens iguais" e criar uma conferência separada para "tabela de sócios e notas".

### 18. Baixa. Linha da lista Valter que confunde
- **Evidência:** na aba "Rafael totais e seções" a linha da lista Valter começa com "10 | 07 - Rádio...". Esse "10" é o nº da linha na lista, mas parece "Empresa 10".
- **Correção:** escrever "nº 10 da lista" ou tirar esse número.

### 19. Baixa. Rafael usa 1/6, e o inventário da Rádio dá 16,50% por sócio
- **Evidência:** na p. 10, Rafael usa 1/6 (485.333,33 sobre 2.912.000). O inventário 07 e o doc 00 dão 16,50% por sócio, mais 1% do Telmo (445.170 sobre 2.698.000). Com 16,50% sobre 2.912.000, o resultado seria 480.480,00.
- O estudo tem a premissa de 1/6, mas a diferença deve ficar registrada.
- **Correção:** acrescentar uma nota na conferência do quadro.

### 20. Baixa. Observação do Rovílio na planilha não foi registrada
- **Evidência:** planilha 2_0, GERAL H874 (e K722 na outra planilha): "Rovílio tinha direito à 2% das empresas Empreendimentos, Abastecedora e Transportadora". Mas os inventários 09 e 10 mostram o Rovílio com 1% da Turismo e 1% da Siviero Transportes. A observação não cita essas duas empresas.
- **Correção:** registrar como contexto para as empresas 09 e 10.

### 21. Baixa. Fontes secundárias que não estão na lista
- `01_fontes/outros/inventário-estudo-rafael.xlsx`: aba "Ativos selecionados", linha 11, Toyota, R$ 70.000, p. 7. Repete a lista Valter.
- `01_fontes/inventario-2025/ACERTOS E ACORDOS.pdf`: é escaneado, então não consegui ler o texto. A aba "Acertos e Acordos" das duas planilhas não cita estas empresas.
- **Correção:** listar os dois com MD5, ou explicar por que ficaram de fora.

### 22. Baixa. Doc 22 de 2018: o caminho tem 263 caracteres e não abre por caminho normal
- **Evidência:** o arquivo "22. 1. SÓCIOS DAS EMPRESAS SIVIERO... Realizada dia 14 de janeiro de 2019" só abre com o prefixo `\\?\`. O conteúdo cita a Rádio, a Turismo, a Siviero Transportes, a Associação e a Ceval em termos gerais, e dá dois valores da área de transportes: "investimento de +/- R$ 1.200.000,00 – compra de 03 ônibus" e "injeção financeira de aproximadamente R$ 2.000.000,00".
- **Correção:** se o doc 22 entrar na lista de fontes, o script deve usar o prefixo `\\?\` para caminhos longos.

---

## O que conferi e está correto

1. **Itens dos inventários.** Refiz célula por célula todos os itens com valor:
   - 2018: Ceval 4 itens = 391.000; Rádio 6 itens = 9.316.016; Siviero Transportes 7 itens = 2.443.100, menos 1.515.100 = 928.000.
   - 2025 atual e antigo: Ceval 3 itens = 315.000; Rádio 7 itens = 2.698.000.
   - Associação, Turismo e Siviero Transportes em 2025: linhas "-.-" com R$ 0,00.
   - Associação e Turismo em 2018: linhas em branco.
   - Nenhum item sobrando, faltando ou inventado. Placas, matrículas, chassis e NF estão certos.
   - Nos ônibus vendidos a prazo (itens 05–07 de 2018), o valor da célula e a situação "Vendido em ..." foram separados corretamente.
2. **Conferências marcadas "NÃO".** Todas são reais nas fontes:
   - 03/2018: total de 10.000 sem nenhum item;
   - 09/2018: célula de total em branco ("R$");
   - 09: inventário × doc 27 (sem total × 0);
   - 03, 07 e 10 na planilha × doc 27 (veja o achado 6).
   Todas as conferências marcadas "sim" também estão certas.
3. **Resumos.**
   - Doc 27: 03 = 10.000; 04 = 391.000; 07 = 9.316.016; 09 = 0; 10 = 928.000.
   - Doc 00 atual e antigo: 03 = 0; 04 = 315.000; 07 = 2.698.000; 09 = 0; 10 = 0.
   - O doc 27 de `01_fontes` tem o mesmo MD5 da cópia do Drive.
4. **Rafael, valores principais.** Estão certos, com as páginas certas:
   - Balanço do grupo: Ceval 200/70/45 mil = 315.000 (p. 1); Rádio sete itens = 2.698.000 (p. 1–2).
   - Faixas de mercado: Ceval 240–295 mil (p. 2), 85–130 mil (p. 3), 90–150 mil (p. 4); Rádio sala 380–520 mil (p. 3), Fox 32.200–35.500 (p. 6), Toyota 45–120 mil (p. 7).
   - Balanço de Valter: Ceval 240.000 + 100.000 + 45.000 = 385.000 (p. 5); Rádio 420.000 + 150.000 + 1.900.000 + 300.000 + 40.000 + 32.000 + 70.000 = 2.912.000 (p. 8–10).
   - Linhas do quadro e Toyota escolhida (p. 11).
   - Associação, Turismo e Siviero Transportes: 2 páginas cada, só o balanço do grupo (zero) e a seção da baixa.
   - Diretrizes: a única menção a estas empresas é a Toyota (p. 7). Introdução: nenhum valor destas empresas.
   - Os PDFs não têm anotações. Há uma imagem por página (papel timbrado).
5. **Lista Valter.** Das 108 linhas de bens (05 HPV: 4; 06 Araçá: 5; 07 Rádio: 1; SEI: 98), só a Toyota Bandeirante é destas empresas: placa IHW 0285, R$ 70.000, p. 7, linha nº 10.
6. **Planilha 2_0, aba GERAL.** As linhas 16–19, 20–27, 105–115, 135–138 e 139–152 foram extraídas corretamente (colunas C a H). Essas linhas não têm observação na coluna H nem nota ou comentário.
   - Os 6 comentários da planilha 2_0 estão nas linhas 223–238 (SEI).
   - Os 94 da outra planilha estão nas abas GERAL e "Lotes que saíram", todos da SEI.
   - A cópia do Hotel Araçá nas linhas 2–33 da aba 7.Rádio foi confirmada.
7. **Hashes.** Conferi 9 MD5 e todos batem com a aba "Fontes": PDF Rafael Rádio, PDF Rafael Ceval, inventário 10 de 2025, ANTIGO 07, lista Valter, planilha 2_0, doc 00, inventário 2018 da Siviero Transportes e doc 23.
8. **Arquivos docx.** Nenhum dos 19 docx usados tem comentários, alterações controladas, texto oculto ou caixas de texto. Não há valores em cabeçalhos ou rodapés.
