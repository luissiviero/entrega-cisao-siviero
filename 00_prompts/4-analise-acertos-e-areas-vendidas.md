Você vai analisar dois documentos do inventário 2025 do grupo Siviero, no contexto da cisão. Eu, Luís, sou sócio de 1/6 e estou saindo. Trabalhe em português, na pasta `C:\entrega-cisao-siviero`. Outras sessões trabalham em paralelo: a da SEI, as das demais empresas e a do PowerPoint. Esta análise alimenta essas sessões, mas não altera os arquivos delas.

**Documentos analisados** (as cópias em `01_fontes` são idênticas por hash às de `C:\inventario-siviero\2025\inventário-2025\`):
1. **`01_fontes/inventario-2025/ACERTOS E ACORDOS.pdf`.** São 6 páginas escaneadas, só imagem. Pelas conferências anteriores:
   - as páginas 1 a 3 trazem cerca de 32 acertos e acordos com terceiros (comissões, serviços, distratos de loteamentos, acordos judiciais), pagos em lotes e em espécie. Os totais declarados são 82 lotes, R$ 2.882.204,55 em espécie e R$ 12.087.204,55 no total;
   - as páginas 4 a 6 trazem 8 pagamentos dos contratos Boff, em nome da SEI, num total de R$ 3.315.675,23.
2. **`01_fontes/inventario-2025/HOTEL PRINCESA DOS VALES LTDA. - ÁREA DE TERRAS VENDIDAS + LOTEAMENTOS CONTRATADOS PELA SEI - SEM OBRAS.docx`.** Contém:
   - a Fazenda São Crispin: 686,75 ha, menos 190,00 ha vendidos à família Rigo em 2018 e 85,32 ha vendidos à família Pellegrini, restando 411,43 ha;
   - créditos em soja sem valor no documento, com campos "R$ ......": 37.694 sacas da família Rigo e 13.085 da família Lorenzet;
   - 2.606 sacas vendidas em 2024 "para aquisição da escavadeira hidráulica";
   - o arrendamento à família Lorenzetti;
   - a lista de áreas contratadas pela SEI, sem obras ou com saldo zero, por cidade.

**Conversão do PDF.** Use a skill `/pdf-to-excel` para tirar as tabelas do PDF. Como é escaneado, confira o resultado linha a linha contra:
- a versão em texto `01_fontes/planilhas-google/ACERTOS E ACORDOS (gdoc).docx`;
- a extração anterior, nas abas "Acertos e acordos", "Contratos Boff" e "Vendas e créditos" de `01_fontes/outros/inventario-2025-avaliacoes-corrigido.xlsx`.

Se a conversão falhar ou divergir, leia as páginas como imagem. Nunca invente números. Onde estiver ilegível, registre como ilegível.

**Antes de começar, leia:**
- `00_PLANO.md`: premissas P1 a P13 e a seção 7, com as lições de método e a regra de arquivos;
- `02_trabalho/README.md`;
- `02_trabalho/log-decisoes.md`.

A memória do projeto traz as regras de trabalho. Siga todas.

**O que fazer:**
1. **Extração completa, uma linha por registro.**
   - **Acordos:** nº, contraparte, motivo (comissão, serviço, distrato, acordo judicial, honorários), empresa, lotes entregues (loteamento, quadra, lote), valor em espécie, valor total, data e texto exato.
   - **Contratos Boff:** pagamento, data, valor e forma.
   - **Áreas vendidas:** comprador, área, data, preço e forma de pagamento (dinheiro, sacas de soja, gado).
   - **Créditos em soja:** safra, sacas, origem e empresa (Hotel, Agropecuária, Oleoplan).
   - **Arrendamento:** áreas, sacas por hectare e prazo.
   - **Áreas contratadas sem obras:** nº, nome, cidade e situação.
   - **Conferência:** confira cada total declarado contra a soma.
2. **Cruzamentos, cada achado com fonte e página:**
   - **Lotes dados em acordo × estoque.** Confira cada lote contra o estoque da SEI em 2025 (`02_trabalho/sei-2025-itens.xlsx`, aba Lotes) e em 2018 (`02_trabalho/sei-2018-itens.xlsx`). Confira também contra o arquivo 17 (lotes vendidos 2018–2025) e o arquivo 18 (lotes consignados ao Válter), em `01_fontes/respostas-docs-2026-08-13/`. Procure lote que ainda conste em estoque, lote em mais de um acordo e lote ao mesmo tempo vendido e dado em acordo.
   - **Desconto em duplicidade.** Procure acordos cujo valor, em lotes ou em espécie, também esteja descontado como passivo no doc 11. Use as abas Compromissos e Passivos de `sei-2025-itens.xlsx`, com as observações 1 a 24 e os lotes de terceiros indenizáveis. Este é o exemplo típico que o Luís quer pegar: um lote sai do estoque e o mesmo valor é descontado de novo como quitação de dívida.
   - **Contratos Boff × Loteamento Caminhos do Sol.** O Caminhos do Sol é o loteamento contratado com a família Boff. Veja se os R$ 3,3 milhões pagos aparecem como ativo, como adiantamento ou como nada, no doc 11 e no estudo de Rafael (`01_fontes/estudo-rafael/empreendimentos-xlsx/ÁREAS CONTRATADAS LOTEAMENTO CAMINHOS DO SOL*`).
   - **Áreas contratadas sem obras.** Cruze esta lista com os loteamentos contratados do doc 11 (9 itens) e com os contratos de 2018 (25 itens, em `sei-2018-itens.xlsx`). Aponte os que aparecem só aqui, que são assuntos não abordados no inventário da SEI (Parte 5).
   - **Datas.** Marque os fatos posteriores a 31.12.2024 (data-base, premissa P2).
   - **Fazenda São Crispin, empresas 02 e 05.** Compare as áreas e as vendas com o estudo de Rafael (491 ha totais; 83,47 ha vendidos à família Rigo; 407,53 ha avaliados) e com `02_trabalho/pendencias-conferencia-anterior.md`, item 1.
   - **Créditos em soja sem valor.** São ativos não avaliados no inventário. Diga de qual empresa é cada crédito.
   - **Escavadeira paga com soja.** As 2.606 sacas vendidas para comprar a escavadeira hidráulica: qual empresa vendeu a soja e qual ficou com a máquina? A SEI tem uma escavadeira 320 de 2024, por R$ 887.000.
3. **Revisor.** Antes de me apresentar qualquer coisa, rode um revisor independente com contexto novo sobre as extrações e os cruzamentos. Diga o que ele encontrou e o que você corrigiu.
4. **Apresentação.** Mostre os achados em pacotes de até 10, do maior valor para o menor, cada um com evidência e recomendação. Eu aprovo cada pacote.

**Arquivos (regra P13).** Escreva só em `02_trabalho/analises/acertos-e-areas-vendidas/`:
- extrações em xlsx;
- `cruzamentos.xlsx`;
- `relatorio.md`;
- `diario.md`;
- `log-decisoes.md`;
- `achados-para-sei.md`, para as Partes 4, 5 e 6;
- `achados-para-empresas.md`, para a São Crispin, o Hotel Princesa e as demais;
- `perguntas-sugeridas.md`.

Não altere nenhum arquivo fora dessa pasta nem os arquivos de memória existentes. As sessões da SEI e das empresas vão incorporar os achados depois que eu aprovar.

Comece me dizendo, em até 5 linhas, o que entendeu e como vai converter o PDF.
