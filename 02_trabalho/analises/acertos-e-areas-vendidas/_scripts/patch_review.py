"""Apply reviewer corrections (28/10/2026) to build_outputs.py."""
import os

p = os.path.join(os.path.dirname(__file__), 'build_outputs.py')
s = open(p, encoding='utf-8').read()


def sub(a, b):
    global s
    assert a in s, a[:80]
    s = s.replace(a, b)


sub("""'03/04': 'Lotes 05 e 06/Q6364 "equivalem a 03 lotes" (16 equivalentes). Doc 11 Obs. 15: "restando ainda a entrega de 06 lotes".',""",
    """'03/04': 'Termo A (17 lotes): 14 lotes com data 13/07/2020, ou 15 equivalentes, porque os lotes 05 e 06/Q6364 "equivalem a 03 lotes"; faltam 2. Termo B (7 lotes): só o Q1729 L12 (30/04/2020); faltam 6, que são os "06 lotes" da Obs. 15 do doc 11. As datas são as dos termos, não as das escrituras.',""")
sub("""'14': 'arq. 17 tem 10 lotes em 07/07/2020 com o texto "num total de 09 lotes" (+1 lote em 2016, permuta com Caminhos do Sol).',""",
    """'14': 'arq. 17 tem 10 lotes em 07/07/2020 com o texto "num total de 09 lotes". O lote Q6373 L12 de 2016 (permuta por 1 lote, 2 % do Caminhos do Sol) ficou fora: não é deste acordo.',""")
sub("""'25': 'arq. 17: 2 lotes (Q6370 L06-07) pelo Loteamento Basso, 1 (Q6369 L08) pelo Reserva Falavigna e 1 dação no Falavigna (Q322 L020). Nenhum no Vale das Colinas.',""",
    """'25': 'arq. 17: Q6370 L06-07 pelo Loteamento Basso; Q6369 L08 e Falavigna Q322 L020 por serviços no Reserva Falavigna (não são deste item). Nenhum no Vale das Colinas. O arq. 18 dá Q6369 L08 e Q6371 L11 como os 2 lotes do Basso, e o Q6371 L11 está no estoque do doc 11.',""")
sub("""'Termo A (13/07/2020, arq. 16 p. 1-4): 17 lotes + R$ 200.000; cl. 1.1 cobre Piccoli, Boff, Andreis, Pedrotti, Fiorese e outros; cl. 6 dá "ampla, total e irrestrita quitação" de toda intermediação até a data; cl. 8 exclui só intermediações novas, que devem ser tratadas por escrito com objeto e valor. Termo B (30/04/2020, p. 5-9): 7 lotes. Arq. 17: 15 lotes (16 equivalentes) entregues a Susin/AGI em 2020. Versão antiga do doc 11: valores em branco ("......").',
         'Os 6 lotes pendentes podem ser devidos (os termos somam 24). A comissão sobre famílias já quitadas no Termo A não é devida sem documento novo escrito.'),""",
    """'Termo B (30/04/2020, arq. 16 p. 5-9): 7 lotes por lotes que a família Susin comprou de Bovo, C. Fioreze e L. Pedrotti, com sub-rogação da SEI; só 1 entregue (Q1729 L12); os 6 pendentes são os da Obs. 15. Termo A (13/07/2020, p. 1-4): 17 lotes + R$ 200.000 por toda a intermediação até a data; cl. 1.1 cita Piccoli, Boff, Andreis, Pedrotti e Fiorese (famílias dos contratos Buriti); cl. 5.1 inclui a comissão "ajustada com a outra parte" desses negócios; cl. 6: quitação total "cumprido integralmente este acordo"; cl. 5: vencimento antecipado após 90 dias; última etapa no 48º mês (07/2024); cl. 8 exclui só intermediações novas, por escrito. Arq. 17: 14 lotes do Termo A (15 equivalentes). Minuta não assinada de 04/02/2020 em 01_fontes/outros. Versão antiga do doc 11: valores em branco.',
         'Os 6 lotes do Termo B são devidos e não entram na quitação. A comissão pelos contratos das famílias do Termo A está dentro do Termo A (cl. 1.1 e 5.1), cuja quitação depende de a SEI cumprir o acordo (faltam 2 lotes equivalentes). Sem documento escrito de intermediação nova (cl. 8), a parte de comissão não tem base.'),""")
sub("""('Lote Altos da Julieta Q1728 L03', 'arq. 18, R005: "Pago Césio Verona, de Farroupilha – serviços prestados para a empresa Rochedo"', 100000.00,""",
    """('Lote Altos da Julieta Q1728 L03', 'arq. 18, R005: "Pago Césio Verona, de Farroupilha – serviços prestados para a empresa Rochedo"', None,""")
sub("""        ('Salas 201-203 (HECTA, VMC, URBAN)',""",
    """        ('Lote Alto das Graças Q6371 L11', 'arq. 18 (Basso, custo 2): "R$ 260.000,00, representados por 02 lotes ... (Lote 08–Quadra 6369 / Lote 11–Quadra 6371)". O lote está no estoque próprio do doc 11 (tabela 38 linha 16, R$ 141.000) e na lista de Rafael (ativos-avaliados-valter.xlsx, Ativos, linha 101). Não está no arq. 17, que dá o Q6369 L08 à Colunata pelo Reserva Falavigna.', 141000.00, 'lote possivelmente já entregue ou comprometido, na lista dos bens que viriam para nós'),
        ('Lotes Alto das Graças Q6371 L05, L06, L12 (consignados)', 'arq. 18, R004: "Área de terras de 2,5 ha em Monte Bérico – Caxias do Sul, com Alexandre Susin. Adquirida a parte do Alexandre Susin em troca de 03 lotes." Nenhum dos 3 está no arq. 17, e o Monte Bérico não aparece em nenhum inventário de 2018 ou 2025.', None, 'lotes da SEI trocados por área que não aparece em inventário'),
        ('Salas 201-203 (HECTA, VMC, URBAN)',""")
sub("""'Inventário 2018: "Pago despesas do Senhor Arlindo Piccoli – dívidas bancárias e de terceiros em troca de lotes" R$ 300.000; arq. 17: lote San Marino Q1342 L03 (2024, "Acordo – Arlindo Píccoli – Vale das Colinas. Dívida de R$ 270.000"), lote Alto das Graças Q6372 L13 a Evandro Zanrosso (2024, dação), R$ 62.500 a Zanrosso (Q6372 L14, 2019)', 300000.00),""",
    """'Inventário 2018: "Pago despesas do Senhor Arlindo Piccoli – dívidas bancárias e de terceiros em troca de lotes" R$ 300.000; arq. 17 linha 117: lote San Marino Q1342 L03 (2024, "Acordo – Arlindo Píccoli – Loteamento Vale das Colinas. Dívida de R$ 270.000,00"); linha 358: R$ 62.500 a Evandro Zanrosso "pela negociação de 01 lote no Loteamento Vale das Colinas" (2019); linha 357: lote Q6372 L13 a Zanrosso em dação (2024, sem valor)', 632500.00),""")
sub("""'Inventário 2018: "Compra de lotes da Família Fioreze" R$ 148.000; ACERTOS itens 22-24: Lori, Ezequiel e Cândido Fioreze R$ 410.000 ("compra"); ACERTOS itens 1-2 e 13 (Angest, Fochesatto): lotes dados pela não execução na área Fioreze/Zanandréa', 558000.00),""",
    """'Inventário 2018: "Compra de lotes da Família Fioreze" R$ 148.000; ACERTOS itens 22-24: Lori, Ezequiel e Cândido Fioreze R$ 410.000 ("compra"); Termo B de Susin (arq. 16): SEI sub-rogada nos lotes que C. Fioreze (4) e E. Bovo (contrato Zanandréa de 14/05/2008) venderam à família Susin, sem valor; ACERTOS itens 1-2 e 13 (Angest, Fochesatto): lotes dados pela não execução na área Fioreze/Zanandréa', 558000.00),""")
sub("""'ACERTOS item 31: Família Pedrotti R$ 270.000 (natureza não informada); Termo B de Susin: 7º lote comprado de Luiz Pedrotti', 270000.00),""",
    """'ACERTOS item 31: Família Pedrotti R$ 270.000 (natureza não informada, a confirmar); Termo B de Susin: SEI sub-rogada no 7º lote comprado de Luiz Pedrotti, sem valor', 270000.00),""")
sub("""note='Arq. 1, p. 30: o grupo diz que esses créditos serão usados "como moeda de postergação de contratos e compensação". O doc 11 tira os créditos do ativo e desconta a multa cheia.')""",
    """note='Critério: só valores em R$ escritos nas fontes. Arq. 1, p. 30: o grupo diz que esses créditos "foi e será acertado como moeda de postergação de contratos e compensação". O doc 11 não lista os créditos e desconta a multa cheia.')""")
sub("""'o grupo admite que o crédito existe'),""",
    """'o grupo trata os créditos como já usados ("moeda de postergação"), não como ativo; no exemplo Boff, prevê compensação a favor dos Boff'),""")
sub("""        ('Rafael – Caminhos do Sol, Quadro_Resumo', '61 lotes × R$ 141.000 = R$ 8.601.000 (10 % da SEI)', 8601000.00, 'não inclui o crédito nem os 13 lotes'),""",
    """        ('Rafael – Caminhos do Sol, Quadro_Resumo', '61 lotes × R$ 141.000 = R$ 8.601.000 (10 % da SEI)', 8601000.00, 'não inclui o crédito nem os 13 lotes'),
        ('Arq. 17, linha 364', 'lote Alto das Graças Q6373 L12 dado à Via Nova em 2016, "PERMUTADO com 01 lote do Loteamento Caminhos do Sol, no percentual de 2,00%"', None, 'mais um direito sobre os lotes do Caminhos do Sol'),""")
sub("""        ('Ansolin (itens 19-21)', 55000 + 156666 + 375000, 'É a compra""",
    """        ('Ansolin (itens 19-21) – hipótese', 55000 + 156666 + 375000, 'Hipótese (o item 19 diz "compra de 06 lotes"): é a compra""")
sub("""        ('Basso (item 30)', round((156100.41 - 11 * 4473.71) + (69060.32 - 11 * 1971.21), 2), 'Somada inteira; 29 de 40 parcelas a pagar (cálculo: valor − 11 parcelas pagas).'),""",
    """        ('Basso (item 30)', round(29 * 4473.71 + 29 * 1971.21, 2), 'Somada inteira; 29 de 40 parcelas a pagar = R$ 186.902,68. As 40 parcelas somam R$ 257.796,80, mais que os R$ 225.160,73 declarados; pela conta valor − 11 parcelas, R$ 154.266,61. Aproximação.'),
        ('Susin, 6 lotes do Termo B', 600000.00, 'Contados entre os 80 lotes × R$ 100.000 do resumo, mas não entregues (Obs. 15 do doc 11). Mais 2 lotes equivalentes do Termo A sem registro no arq. 17.'),""")
sub("""note='Parcelas dos R$ 12.087.204,55 que não são acertos já pagos. O ACERTOS não é deduzido no total do doc 11; serve ao grupo para explicar a saída de lotes (arq. 1, p. 18: 751 → 421 lotes).')""",
    """note='Parcelas dos R$ 12.087.204,55 que não são acertos já pagos. O ACERTOS não é deduzido no total do doc 11. O lado consulente cita os 82 lotes do ACERTOS ao perguntar pela queda de 751 para 421 lotes (arq. 1, p. 18).')""")
sub("""    ar = [(n, nome, cid, *AREAS_X[n], """,
    """    ar = [(n, nome, cid, *AREAS_X[n], 'conflito: distrato em 2018, listada como contratada' if n == '09' else """)
sub("""        ('Reserva Paradiso', 'Q104 L623 (no ativo do doc 11)', 'entrada 19/09/2025; venda 07/01/2026', 'sim'),""",
    """        ('Reserva Paradiso', 'Q104 L623 (no ativo do doc 11)', 'entrada 19/09/2025; venda 07/01/2026', 'sim'),
        ('Fora do estoque do doc 11', 'Barbiero Q05 L03 e L08 (acordo judicial, arq. 17 linhas 53-54)', '23/05/2025', 'sim'),
        ('Fora do estoque do doc 11', 'Falavigna Q174 L020, L021, L023 (Indústria Florense, arq. 17 linhas 387-389)', '20/03/2025', 'sim'),
        ('Fora do estoque do doc 11', 'Falavigna Q175 L012, L013 (Bett, arq. 17 linhas 390-391)', '10/07/2025', 'sim'),
        ('Fora do estoque do doc 11', 'Falavigna Q876 L016 (dação G3, arq. 17 linha 383)', '10/09/2025', 'sim'),""")
sub("""        ('Rigo 2021 + 2022 (líquido)', 37694, 'Hotel Princesa dos Vales', 'a área vendida está em nome do Hotel (doc 05 item 05; doc 02 diz que as escrituras estão em nome do Hotel)', 37694 * 80.0),""",
    """        ('Rigo 2021 + 2022 (líquido)', 37694, 'Hotel Princesa dos Vales (provável)', 'título do docx; doc 02 diz que as escrituras estão em nome do Hotel; só 83,47 ha dos 190 ha saem de matrícula listada no doc 05. Natureza indefinida: a tabela diz "FALTA RECEBER", a observação diz "AINDA NÃO FORAM VENDIDOS". Coincidência: as 2.606 sacas vendidas = Lorenzet 2021', 37694 * 80.0),""")
sub("""        ('Lorenzet 2021-2023', 2606 + 3206 + 3591, 'Hotel Princesa dos Vales', 'texto "HOTEL/OLEOPLAN"', (2606 + 3206 + 3591) * 80.0),
        ('Lorenzet 2024', 3682, 'Agropecuária Fazenda São Crispin', 'texto "AGROPECUÁRIA/OLEPOPLAN"', 3682 * 80.0),""",
    """        ('Lorenzet 2021-2023', 2606 + 3206 + 3591, 'Hotel Princesa dos Vales', 'texto "HOTEL/OLEOPLAN"; o arrendador nos contratos 8 e 9 é o Hotel (p. 1 de cada contrato, segundo o revisor)', (2606 + 3206 + 3591) * 80.0),
        ('Lorenzet 2024', 3682, 'conflito: o docx diz Agropecuária; o contrato 9 (2023-2026) tem o Hotel como arrendador', 'texto "AGROPECUÁRIA/OLEPOPLAN"', 3682 * 80.0),""")
open(p, 'w', encoding='utf-8').write(s)
print('patched')
