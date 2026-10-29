# Fase 0 — Abastecedora (01) e Transportadora (08): resumo

28/10/2026. Números lidos das fontes por script (`_scripts/extract_fase0.py` → `_extracao/fase0-itens-01-08.xlsx`, aba Conferências). Passou por revisor independente (`_diario/revisao-fase0.md`): 18 achados, 2 de gravidade alta, todos na leitura do PDF de Rafael da Abastecedora. Todos foram corrigidos; nenhum valor de inventário estava errado.

## 1. Fontes e totais

São 23 fontes com hash (`_extracao/hashes.txt`). Os quatro contratos escaneados (arquivos 2, 3, 4 e 10) foram lidos como imagem, página a página. Todos os totais conferem com os documentos e com os resumos (doc 00 de 2025, atual e antigo; doc 27 de 2018).

| | Grupo 2025 | Rafael | Diferença | 2018 |
|---|---|---|---|---|
| 01 Abastecedora (1 item) | 9.661.414,00 | 12.361.850,00 | +2.700.436,00 | 12.361.851,89 (posto 12.216.851,89 + caminhão-tanque 145.000) |
| 08 Transportadora (13 itens) | 16.374.986,00 | 16.924.986,00 | +550.000,00 | 1.564.000,00 (Transportadora) + itens 01–05 da Oficina (10.721.998,84) |

- **Lista de Rafael:** nenhum bem da 01 ou da 08 está nela (C0: todos ficam com o maior valor). Os PDFs das duas empresas dizem "não há interesse na sua inclusão como ativo selecionável".
- **Versão antiga × atual:** os itens e valores são iguais. Muda só o rateio (sai o 2 % do Rovílio e entram 6 × 16,66 %). Na versão atual da 01, o valor do item 01 sumiu da célula, mas a descrição recompõe exatamente R$ 9.661.414,00.

## 2. Onde Rafael diverge

| Item | Grupo | Rafael | Fonte (PDF de Rafael) |
|---|---|---|---|
| 01-01 Posto (mat. 3.752) | 9.661.414 | 12.361.850 | p. 2–3. É o total da empresa em 2018, arredondado, e inclui o caminhão-tanque ICW 8342, que não existe em 2025. Está ao lado da descrição de 2025, cujas parcelas somam 9.661.414. |
| 08-03 Equipamentos da oficina | 500.000 | 1.000.000 | p. 13. Vem da "declaração expressa" de Benta, p. 6, que fala em "aproximadamente R$ 1.000.000,00". Na p. 5 o item ainda aparece com 500.000. |
| 08-12 Scania IJZ 5399 (prancha) | 140.000 | 160.000 | p. 15. Faixas: só o cavalo, 103–162 mil (FIPE) e 150–180 mil (anúncios); com prancha, 200–310 mil (p. 10–12). |
| 08-11 Carreta Randon IQB 0612 | 30.000 | 50.000 | p. 15. Faixa de 50–90 mil (p. 9–10). |
| 08-13 Carreta Randon IJZ 8207 | 30.000 | 40.000 | p. 15. Faixa de 40–85 mil (p. 12). |

Os outros 9 itens da 08 têm o mesmo valor nas duas avaliações. A conferência anterior (26/10) deu a 08 como "confere" porque comparou o balanço do grupo que Rafael transcreve (p. 1–3), e não o balanço dele (p. 13–15).

## 3. 2018 → 2025 (pares por matrícula, registro e placa)

- **10 pares por chave:** posto (3.752); terreno da oficina (3.753); sala do escritório (7.393); apto 401 (7.400); 4 veículos (IKR 2054, IQB 0612, IJZ 5399, IJZ 8207); terras de Cotiporã (459/460/461/558 e reg. 15.922).
- **4 pares sem chave**, só pela descrição: pavilhão, equipamentos, rampa e borracharia, que em 2018 eram da Oficina. Vou confirmar no pacote.
- **2 bens migraram para a SEI** (só registro): terreno da serraria (3.754) e imóvel Oleoplan (17.409 + 6.136 → 29.954).
- **4 bens de 2018 sem par em 2025**, que precisam de destino (P4):
  - caminhão-tanque ICW 8342 (R$ 145.000);
  - sala do consultório, mat. 18.952 (R$ 149.000);
  - móveis do escritório (R$ 30.000);
  - apto 102, mat. 7.395 (R$ 220.000, já registrado em 2018 como doado a Geni).
- O processo Posto Pértile, citado em 2018, não aparece em 2025.
- **Cotiporã:** em 2018 as terras estavam na Agropecuária com "NÃO DIVISÍVEL", sem valor e fora do total. Isso responde aos seus "?" (aba 8.Transportadora, C10 e C11).

## 4. Achados nos contratos (fatos novos)

1. **Locação à SIM** (arq. 4, p. 1 e 7):
   - A área locada é de 3.450 m² da mat. 3.752 **mais 2.500 m² da mat. 3.754**, o terreno da serraria, que hoje é da SEI, usados como estacionamento. A mat. 3.753 da Transportadora só aparece como faixa de acesso. Isso contradiz a resposta do grupo, segundo a qual a Transportadora teria dado parte do seu terreno em comodato (arq. 1, p. 16).
   - Aluguel de R$ 15.000/mês em 2022, com reajuste anual. Prazo de 01/01/2022 a 31/12/2033, mais 36 meses.
   - Restaurante, borracharia e rampa ficaram fora da locação, com distrato previsto até 31/05/2022. É o contexto do processo Lucchetta, ajuizado em 2023.
2. **Petrobras** (arq. 2 e 3):
   - Os contratos são da Abastecedora **Certo**, a operadora anterior, e vigoraram até 30/09/2025.
   - A Abastecedora Siviero recebeu, por cessão, a **bonificação antecipada de R$ 750.000,00**. Há uma hipoteca sobre um imóvel que os contratos não identificam.
   - Se o volume não foi comprado, a bonificação é devolvida na proporção, com juros e multa de 10 %. A obrigação é do revendedor; pela locação, a SIM assumiria esses contratos.
   - É uma possível contingência em 31.12.2024. O inventário 2025 ainda fala em "Ipiranga".
3. **Venda das quotas do Rovílio** (arq. 10): R$ 3.440.082,00 é a soma dos 10 itens de pagamento (confere).
   - **R$ 982.072 são "comissão"** pelo Morada do Ipê, e não preço das quotas. Sem essa parcela, o preço cai para R$ 2.458.010,00, contra os R$ 2.538.208,55 (2 % dos inventários) do Dr. Pablo.
   - O preço também inclui "créditos hereditários" sobre a **terra de Cotiporã (São Brás)**.
   - Tudo foi pago com bens da SEI.
   - Esses pontos ficam para o `rovilio-quotas.md`, que passa pelo revisor antes de chegar a você.

## 5. Contas que não fecham nos próprios documentos

- 08-07: 69,8 ha × R$ 50.000 = 3.490.000, mas o item vale 3.500.000.
- 08-08: 4,5 ha × R$ 50.000 = 225.000, mas o item vale 250.000.
- A nota de R$ 1.135.000 corresponde aos itens 03–05, que somam R$ 1.020.000 (sobram R$ 115.000).
- No quadro-resumo de Rafael da 01, o 1/6 está escrito como R$ 2.060.166; o correto seria R$ 2.060.308,33.
- Na sua planilha, a nota da Resumo C63 diz 36,53 %; a conta dá 35,53 %. O rótulo da Resumo B73 diz "Hotel Princesa", mas a célula aponta para o caminhão-tanque da Abastecedora.

## 6. Próximo passo (aguarda o seu OK)

- Planilha-mestre `planilha-mestre-01-08.xlsx`, gerada por script, com os pares acima.
- Em paralelo, o `rovilio-quotas.md`, com revisor antes de chegar a você.
