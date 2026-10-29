# Relatório — ACERTOS E ACORDOS e docx do Hotel (áreas vendidas e áreas sem obras)

Situação em 28/10/2026: pacote 1 aprovado pelo Luís; pacote 2 em revisão. Planilhas: `extracoes.xlsx` (extração) e `cruzamentos.xlsx` (cruzamentos). Método, hashes e revisão em `diario.md`.

## 1. Extração

- **ACERTOS E ACORDOS.pdf, p. 1-3.** 32 itens numerados, 36 linhas. Totais conferem: 82 lotes; espécie R$ 2.882.204,55; total R$ 12.087.204,55 (80 lotes × R$ 100.000 + R$ 830.000 + R$ 375.000 + espécie). As p. 1-3 não identificam a empresa. Nenhum trecho ilegível.
- **ACERTOS, p. 4-6 (cabeçalho da SEI).** 8 pagamentos ligados à família Boff (um deles é IPTU pago à Prefeitura). Total: R$ 3.315.675,23, confere.
- **Docx do Hotel.**
  - Vendas de 2018: R$ 12.058.000 (Rigo, 190 ha; Pellegrini, 85,32 ha; Angeli, 10,2 ha).
  - Créditos em soja: 37.694 + 13.085 sacas.
  - Arrendamento: 411,4 ha, de 05/2023 a 04/2026.
  - 12 áreas sem obras.
  - Todas as somas conferem.
- **Conversão.** O `/pdf-to-excel` não achou camada de texto. As páginas foram lidas como imagem e conferidas com o gdoc e com a extração de 26/10, sem divergência. O revisor reconferiu as imagens. O arrendador do contrato 9 foi conferido por OCR (Tesseract).

## 2. Pacote 1 — achados por valor em jogo (do maior para o menor)

| # | Achado | Valor em jogo | Para | Recomendação |
|---|---|---|---|---|
| 1 | Créditos em soja sem valor em nenhum inventário: 50.779 sacas, quase todas do Hotel. Os 3.682 de 2024 o docx atribui à Agropecuária, mas o contrato tem o Hotel como arrendador. | referência R$ 4.062.320 (R$ 80/saca, contrato de 2018; não é avaliação) | empresas 05/02 | Ativo não abordado: pedir quantidade e valor em 31/12/2024 |
| 2 | Crédito Boff (Caminhos do Sol) fora do doc 11 e de Rafael | R$ 3.315.675,23 nominais | SEI, Parte 5 | Assunto não abordado; pedir saldo e forma |
| 3 | Lotes da SEI em 31/12/2024 descontados como "de terceiros" (Vicentino Q1301 L003, Primavera Q1199 L010-011, Jardim Itália Q1747 L001) | R$ 1.887.000 | SEI, Parte 6 / P2 | Retirar da tabela 47 |
| 4 | Susin, Obs. 15: a comissão "Buriti" já está dentro do Termo A de 2020; os 6 lotes do Termo B são devidos | até R$ 1.200.000 menos os 6 lotes | SEI, Parte 6 | Aceitar os lotes, contestar a comissão |
| 5 | Créditos com proprietários não abatidos das multas cheias (Vale das Colinas, Vista da Colina, Pedrotti) | R$ 1.190.500 + R$ 270.000 a confirmar | SEI, Parte 6 | Abater ou trazer ao ativo |
| 6 | HECTA/Basso: custos de Válter contados como acertos da SEI; lote Q6371 L11 dado pelo Basso e ainda no estoque e na lista de Rafael; multa do Cecconello (HECTA) na SEI | R$ 583.695,51 + R$ 141.000; Cecconello R$ 1.300.000 | SEI, Parte 6 | Esclarecer a HECTA; segurar o Q6371 L11 |
| 7 | Lotes de acordos (Farina, Ambitec) que voltam como indenizáveis | R$ 360.000 | SEI, Parte 6 | Pedir base legal |
| 8 | Composição do ACERTOS: cerca de R$ 1,76 milhão não é acerto pago (Ansolin, Angest em dobro, 6 lotes Susin, parcelas Flávia e Basso) | sem efeito direto no doc 11 | SEI, Parte 4 | Usar contra a narrativa da saída de lotes |
| 9 | Lotes da SEI em 31/12/2024 dados ou vendidos em 2025 e ausentes do doc 11 (Barbiero, Falavigna) | sem valor no doc 11 | SEI, P2 | Registro de data-base |
| 10 | São Crispin: as áreas fecham (411,53 × 411,43 ha); falta o destino dos R$ 12.058.000 recebidos | R$ 12.058.000 (destino) | empresas 05/02 | Fechar a pendência 1 (área); perguntar o destino |

Detalhe por item: `achados-para-sei.md` (S1 a S13) e `achados-para-empresas.md` (E1 a E7).

## 3. Pacote 2 — depois da aprovação do pacote 1

Áreas sem obras (lista da Parte 5), escavadeiras (soja do Hotel na 320 de 2024 da SEI; NF da 320D 2009 da Ceval com 1/3 de Celito), contagem de lotes, Acordi, Perini Q1469 L18-19, Monte Bérico, Flávia Mânica (parte relacionada).

## 4. Limites

- O inventário de 2018 da SEI lista quantidades por loteamento, não lotes. A existência dos lotes dos acordos em 2018 foi conferida só pela data de saída no arq. 17 (fev/2018 a mai/2022).
- R$ 100.000/lote é o critério do próprio ACERTOS, e R$ 80/saca é o preço do contrato Rigo de 2018. Nenhum dos dois é avaliação (P5).
- O arq. 17 está incompleto: faltam o Primavera Q1199 L10, o Q6371 L11 e os lotes do Termo B, entre outros.
