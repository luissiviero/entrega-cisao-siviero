# Revisão independente — Pacote 1 (01 Abastecedora e 08 Transportadora)

- **Data:** 28/10/2026
- **Revisor:** agente de contexto novo. Não li `_diario/` nem `log-decisoes.md`, e nenhum arquivo existente foi alterado.
- **Escopo:** `pacotes/pacote-01.md` e `_scripts/pacote_01.py` (rodado com `python -I`, só leitura; os horários de gravação da planilha-mestre e da fase 0 não mudaram), `perguntas-ao-grupo.md`, `parte6-candidatos.md` e `registro-data-base.md`, sobre a base `planilha-mestre-01-08.xlsx`.
- **Fontes conferidas:**
  - inventários 2025: docs 01, 08 e 11, pelas extrações em `_extracao/`;
  - inventários 2018: docs 10, 12, 15 e 17;
  - PDFs de Rafael (08, 01 e busca por texto nos demais PDFs da pasta `estudo-rafael`);
  - `ativos-avaliados-valter.xlsx`;
  - arq. 1, p. 1–4 e 14–16;
  - `contratos-abastecedora.md` (arq. 2, 3 e 4) e `arq10-rovilio-transcricao.md`;
  - planilha Google 2_0, células citadas;
  - lista da Parte 6 da SEI (`02_trabalho/parte6-candidatos-duplicidade.md`).

**Contagem:** 1 alta, 10 médias, 9 baixas.

---

## Achados

### Alta

**1. R3 fecha a busca de destino (P4) do apto 102 só com a declaração de 2018.** Gravidade **alta** (regra e cobertura).

- **Evidência.**
  - O doc 15 de 2018, item 09 (T4.R2), diz: "Conforme acordo dentre as partes, este imóvel foi doado para a Senhora Geni Siviero Mattiuz". O texto registra um acordo entre os sócios. Não traz escritura, registro nem data de transmissão.
  - Pela P3, 2018 é benchmark de existência. Prova que o bem existia e que houve a intenção de doar, mas não que a propriedade saiu da 08.
  - Nenhuma fonte de 2025 trata da mat. 7.395. O próprio pacote mantém a pergunta T-07 ("foi escriturada? Em que data?"), ou seja, admite que o destino não está provado.
  - Se a matrícula ainda estava em nome da 08 em 31.12.2024, falta um bem no inventário 2025 da 08 (R$ 220.000 em 2018), e isso afeta a quota de 1/6.
  - É também um fato que o cliente precisa decidir: o acordo de 2018 foi assinado por todos, inclusive Válter.
- **Correção sugerida.**
  - Reescrever a R3 como "destino declarado em 2018, pendente de prova". A linha fica aberta na P4 até a certidão da mat. 7.395, ou vai para o pacote 2 como fato a decidir.
  - Em vez de perguntar ao grupo (T-07), o lado do Luís pode tirar a certidão por conta própria. O grupo diz que não se opõe (arq. 1, p. 1 e p. 3, item 1).

### Média

**2. R1 omite que a concordância de Rafael é condicionada.** Gravidade **média** (ressalva omitida).

- **Evidência.**
  - Rafael 08, p. 4–5, "DA CONCORDÂNCIA COM OS VALORES DE AVALIAÇÃO PATRIMONIAL": a concordância vale "sem prejuízo da necessária conferência documental e jurídica quanto à titularidade, existência, regularidade registral". Fica ainda "sujeita à validação pelo escritório de advocacia contratado". Os itens 02 a 09 remetem a essa mesma cláusula (p. 5–8).
  - A conclusão da p. 17 também é condicionada.
  - Isso pesa justamente sobre os itens com dúvida de titularidade (07 e 08, São Brás) e de regularidade (02, 04 e 05).
- **Correção sugerida.** Acrescentar uma ressalva geral à R1: a regra adota o **valor**. A titularidade e a inclusão do bem na 08 ficam condicionadas a T-03 e T-04, a DB-08 e às certidões.

**3. "O grupo desconta ainda 1,5 % do total" e "descontada três vezes" exageram o que o documento faz.** Gravidade **média** (ressalva errada).

- **Evidência.**
  - O total do doc 08 de 2025 é bruto: R$ 16.374.986,00 (T4.R9) = soma dos itens.
  - O 1,5 % é uma redução anunciada, não aplicada: "O valor patrimonial lançado sofrerá redução pela regularização dos imóveis, o que se estima numa redução de 1,50%".
  - O INSS (R$ 122.254,73 e R$ 62.555,51) é só uma estimativa. Não há CNO ("Não há cadastro", arq. 1, p. 15, item 3).
  - Nos itens 04 (rampa) e 05 (borracharia), o inventário de 2018 nunca teve a linha dos 30 %. A rampa vale R$ 300.000 nos dois anos, e a borracharia subiu de R$ 25.000 para R$ 220.000. Para esses dois itens, "sem os 30 %" não é redução em relação a 2018.
  - Hoje, portanto, só a falta dos 30 % no pavilhão (item 02) está refletida nos valores.
  - O mesmo exagero aparece no texto do C-05 ("reduz o valor três vezes") e no resumo do pacote ("C-05: obra não regularizada descontada três vezes").
- **Correção sugerida.**
  - Ressalva: "o grupo anuncia uma redução estimada de 1,5 % sobre o total (R$ 245.624,79, não aplicada ao total) e lista o INSS estimado à parte".
  - C-05: "se a redução de 1,5 % e o INSS vierem a ser descontados, o mesmo fato reduzirá o valor até três vezes. No pavilhão, a falta dos 30 % já está no valor".

**4. Rampa e borracharia (itens 04 e 05): três fontes divergem, e isso não aparece na ressalva, nos candidatos nem em T-08.** Gravidade **média** (omissão relevante).

- **Evidência.**
  - O doc 08 de 2025 trata as obras como da 08, "instalados sobre uma das propriedades da empresa" (observações).
  - O arq. 4, p. 1 e 3, trata os mesmos espaços como da **Abastecedora** (a Locadora): "Os espaços hoje explorados para restaurante/lancheria, borracharia e rampa de lavagem são locados para terceiros". A Locadora se compromete a fazer os distratos e um aditivo para incluí-los na locação à SIM. Pela cl. 3ª, § 3º (p. 4), os aluguéis das sublocações são somados ao aluguel.
  - O arq. 1, p. 16, item 6, responde ao pedido de "contratos de locação/uso da rampa e da borracharia" com "Se trata de contrato conjunto com a REDE SIM".
  - A ressalva 1 do pacote diz apenas que a localização "não está documentada". Na verdade há documentos, e eles divergem sobre quem é o titular e quem recebe a renda.
- **Correção sugerida.**
  - Citar as três fontes na ressalva 1.
  - Criar o candidato **C-08**: rampa e borracharia avaliadas como ativo da 08 (R$ 520.000), com o aluguel possivelmente recebido pela 01, onde "Agrega ao valor do Posto, o aluguel". É uma suspeita de dupla contagem.
  - Reescrever T-08 a partir da resposta da p. 16: quem é o locador em cada contrato, se o aditivo SIM foi assinado e para qual empresa vai o aluguel.

**5. Falta o candidato "2 % do Rovílio sobre bens que em 2018 eram de empresas onde ele não tinha quota".** Gravidade **média** (omissão relevante; liga com S1).

- **Evidência.**
  - Em 2018, a Oficina (doc 10, T1) e a Agropecuária (doc 17, T1) tinham só os seis irmãos, com 16,666 % cada. Rovílio não aparece.
  - Em 2025, esses bens estão na 08, onde Rovílio tinha 2 %:
    - itens 01 a 05, vindos da Oficina: R$ 11.084.986;
    - itens 07 e 08, vindos da Agropecuária: R$ 3.750.000;
    - total: R$ 14.834.986. Os 2 % disso são R$ 296.699,72.
  - Esse valor é cerca de 90 % dos R$ 327.499,72 atribuídos à 08 na nota do Luís (Resumo B67). O preço foi pago com bens da SEI (arq. 10, cl. 4ª).
  - Busquei "Oficina" em `rovilio-quotas.md` e não há nenhuma ocorrência.
  - É uma suspeita, não uma conclusão: o grupo afirma que os imóveis estão "registrados em nome da empresa" (arq. 1, p. 15), e a Oficina não tinha CNPJ.
- **Correção sugerida.** Acrescentar o candidato (forma 9, intragrupo/S1) e citá-lo em "Por que importa" de T-03 e T-04. Encaminhar à sessão S1.

**6. T-02 pode entregar ao grupo um argumento contra o cliente.** Gravidade **média** (só aponto; a decisão é do cliente).

- **Evidência.**
  - Nos itens 07 e 08, os dois lados adotam R$ 3.500.000 e R$ 250.000, e a R1 do próprio pacote aprova esses valores.
  - Pela C0, o maior valor favorece o cliente, porque o bem fica na empresa.
  - Perguntar "Qual valor vale?" convida o grupo a reduzir R$ 35.000 (R$ 10.000 + R$ 25.000) e contradiz a R1.
- **Correção sugerida.** O cliente decide se retira T-02 ou se a mantém só como nota interna.

**7. T-04 tem efeito de dois gumes, e a fonte está incompleta.** Gravidade **média** (só aponto).

- **Evidência.**
  - "Em nome de quem estão as matrículas?" pode levar o grupo a tirar São Brás da 08. São R$ 3.750.000 no total; 1/6 disso é R$ 625.000.
  - O grupo já afirmou que os imóveis da 08 estão "devidamente registrado[s] em nome da empresa" (arq. 1, p. 15, itens 1 e 4). A pergunta não cita essa resposta.
  - Também não cita a observação do doc 17 de 2018: "É necessário transmitir todas as Escrituras Públicas das Propriedades Rurais, em nome da empresa".
  - O efeito liga-se ao C-03: se a terra é da 08, o pagamento ao Rovílio é suspeito; se não é, a 08 perde o valor.
- **Correção sugerida.**
  - Citar as duas fontes.
  - Considerar tirar as certidões diretamente, como na sugestão do achado 1, em vez de perguntar.
  - Apresentar ao cliente os dois efeitos antes de enviar.

**8. A-03 e T-10 convidam o grupo a quantificar passivos contra a 01 e a 08.** Gravidade **média** (só aponto; há também um fato impreciso).

- **Evidência.**
  - A segunda parte de T-10 ("qual é a posição da Transportadora caso o acordo não seja homologado") convida o grupo a reforçar a contingência de "mais de R$ 1.000.000", que reduz a 08.
  - A-03 trata a devolução da bonificação e a multa como possível passivo da Abastecedora Siviero. Pelo arq. 3, p. 1, porém, a obrigação de restituir continua com o **revendedor** (Abastecedora Certo). A Siviero é cessionária do crédito e possivelmente garantidora hipotecária.
  - "Recebida pela Abastecedora Siviero" não está documentado. O contrato diz que o crédito "será liberado [...] para a INTERVENIENTE ANUENTE".
- **Correção sugerida.**
  - Pedir documentos, sem pedir avaliação: a certidão da matrícula com a baixa da hipoteca (ou não), o comprovante do valor depositado pela seguradora e a situação da homologação.
  - Trocar "recebida" por "cedida".
  - O cliente decide sobre a segunda parte de T-10.

**9. Faltam perguntas que favorecem o cliente, ligadas à maior exceção (posto, R$ 2,7 milhões) e ao complexo da oficina.** Gravidade **média** (omissão relevante).

- **Evidência.**
  - **(a) Antiga oficina.** O arq. 1, p. 15, item 2, diz que a oficina "encontra-se em operação [...] na parte da 'antiga oficina'", e que o pavilhão novo ainda não está em uso. O doc 08 não lista a construção da antiga oficina. O terreno 3.753 vale R$ 992,78/m², o mesmo preço por metro do terreno da antiga serraria (mat. 3.754) em 2018. A construção pode estar fora do inventário.
  - **(b) Ampliação do posto.** O arq. 1, p. 3–4, diz que os projetos estão aprovados, "inclusive da ampliação que também pende de averbação". O doc 01 de 2025 repete as áreas de 2018 (292,85 m² e 699,55 m²).
  - **(c) Reforma depois de 2018.** A-01 cita só Rafael, p. 2. Há prova documental de reforma: a bonificação foi liberada por etapas de obra (arq. 3, p. 1, item V): "reforma cobertura pista", "Implantação da Imagem Pista" e "Conclusão das Obras".
- **Correção sugerida.**
  - Incluir as perguntas: (i) a antiga oficina é construção sobre a mat. 3.753 e por que não está no inventário; (ii) a ampliação aprovada está nas áreas do item 01.
  - Acrescentar o arq. 3, p. 1, como fonte de A-01.

**10. C-01 e C-02 estão formulados como "ônus × valor livre", o que tende a reduzir bens que ficam na empresa.** Gravidade **média** (só aponto).

- **Evidência.**
  - Nem a mat. 3.754 nem a mat. 3.753 estão na lista de Rafael (`ativos-avaliados-valter.xlsx`). Os dois terrenos ficam na empresa, e pela C0 recebem o maior valor.
  - O argumento do "ônus não declarado" (comodato até 2036, locação de longo prazo) só serve para baixar esses valores, contra o cliente.
  - Além disso, nenhum dos dois é "desconto em duplicidade". A lista da SEI separa "dupla contagem" de "desconto em duplicidade".
  - O argumento a favor do cliente é a renda: a área de outra empresa produz aluguel na 01, que "agrega ao valor do Posto".
- **Correção sugerida.** Reformular os dois candidatos pelo lado da renda, que é dupla contagem, e indicar o tipo. O cliente decide se o lado do ônus deve aparecer.

**11. Pacote 2: o item 1 (posto) e o item 5 (caminhão-tanque) interagem.** Gravidade **média** (omissão relevante).

- **Evidência.**
  - O valor de Rafael para o posto, R$ 12.361.850, é o total da empresa em 2018 (Rafael 01, p. 2), que inclui o caminhão-tanque de R$ 145.000. O DB-06 registra isso.
  - A lista do pacote 2 trata o caminhão como exceção à parte. Se o valor de Rafael for adotado e o caminhão também tiver destino ou valor, o mesmo valor entra duas vezes.
- **Correção sugerida.** Anotar no item 1 da lista: "inclui o caminhão-tanque (ver item 5)".

### Baixa

**12. A lista do pacote 2 não está "na ordem de valor".** Gravidade **baixa** (forma).

- **Evidência.** As carretas (R$ 50.000 e R$ 40.000) vêm antes da sala do consultório (R$ 149.000) e do caminhão-tanque (R$ 145.000). O script imprime outra ordem: posto, equipamentos, sala, caminhão, prancha, carretas, móveis.
- **Correção sugerida.** Reordenar pelo maior dos valores (posto, equipamentos, prancha, sala, caminhão, carretas, móveis), ou dizer que a lista está agrupada por tipo.

**13. O resumo do pacote chama os 7 candidatos de "suspeitas de desconto em dobro" e redige C-03 como conclusão.** Gravidade **baixa** (forma).

- **Evidência.**
  - C-01, C-02, C-03 e C-07 são dupla contagem ou ônus. C-04 é data-base.
  - "C-03: terra de São Brás contada como ativo da 08 e paga ao Rovílio" está escrito como fato. O arquivo da Parte 6 diz "pode estar".
  - Falta a coluna "Tipo", que a lista da SEI usa e de que a fusão (P13) vai precisar.
- **Correção sugerida.** Usar "suspeitas de dupla contagem ou de desconto em duplicidade", redigir C-03 no condicional e acrescentar a coluna "Tipo".

**14. DB-06: o total de 2018 é R$ 12.361.851,89, não R$ 12.361.850,00.** Gravidade **baixa** (diferença de R$ 1,89).

- **Evidência.** O doc 12 de 2018, T3.R4, dá R$ 12.361.851,89. R$ 12.361.850,00 é o número de Rafael (Rafael 01, p. 2–3). O pacote também diz "Rafael R$ 12.361.850 (valor de 2018)".
- **Correção sugerida.** Escrever: "Rafael adota R$ 12.361.850,00, que corresponde ao total de 2018 (R$ 12.361.851,89)".

**15. DB-07: o intervalo de páginas inclui um veículo da R1.** Gravidade **baixa** (referência).

- **Evidência.** As faixas dos itens 11 a 13 estão nas p. 9–12. A p. 8 é a faixa do item 10 (IKR 2054), que está na R1 e fora do DB-07.
- **Correção sugerida.** Trocar para "p. 9–12".

**16. A-07: o processo de 2025 pode ser o mesmo litígio, com outro autor.** Gravidade **baixa** (referência).

- **Evidência.** O doc 12 de 2018 traz "POSTO PÉRTILE LTDA. – Ação de Indenização = Ponto e pagamento de aluguéis". O doc 01 de 2025 traz "CRISTIANO RODRIGO LUCCHETTA – Ação de Indenização = Ponto e pagamento de aluguéis", com o mesmo texto. Os processos de 2018 não estão exatamente "ausentes": podem ter sido substituídos.
- **Correção sugerida.** Perguntar se há relação entre os processos de 2013/2014 e o de 2023. Antes de enviar, avaliar se a pergunta convida o grupo a quantificar uma contingência.

**17. T-06 sugere a resposta que favorece o grupo.** Gravidade **baixa** (forma).

- **Evidência.** "Foram incorporados às 'instalações' da sala 7.393?" oferece a explicação de que os R$ 30.000 já estão dentro dos R$ 840.000.
- **Correção sugerida.** Perguntar só o destino dos móveis e equipamentos (venda, baixa, transferência ou permanência), sem sugerir a resposta.

**18. R2 não diz que a OFI-07 (mat. 29.954) está na lista de Rafael.** Gravidade **baixa** (o tratamento é da SEI).

- **Evidência.** `ativos-avaliados-valter.xlsx`, linha 16: mat. 29.954, R$ 10.631.378,60. O grupo dá R$ 13.631.378,60 (doc 11, item 04). Por ser bem escolhido por Rafael, a C0 aponta o menor valor.
- **Correção sugerida.** Acrescentar à R2: "OFI-07 está na lista de Rafael (C0: menor valor), a cargo da SEI".

**19. Achado colateral para a SEI: três lotes a unificar, mas só duas matrículas.** Gravidade **baixa** (fora do escopo).

- **Evidência.**
  - O doc 10 de 2018, nas observações, diz: "É necessário unificar os 03 lotes onde está construído o PAVILHÃO".
  - A mat. 29.954 une só a 17.409 e a 6.136 (doc 11, item 04).
  - Os croquis do arq. 4 (p. 15–17) mostram "M.17.408" ao lado de "M.17.409".
  - Pode haver um terceiro lote sem destino.
- **Correção sugerida.** Encaminhar à sessão SEI como possível pendência P4.

**20. Borracharia: o valor subiu 8,8 vezes de 2018 (R$ 25.000) para 2025 (R$ 220.000), e isso não aparece na R1.** Gravidade **baixa**.

- **Evidência.** Doc 10 de 2018, item 05, × doc 08 de 2025, item 05. Pela P3 isso não muda o valor adotado. Mas reforça a dúvida sobre o par feito por descrição: pode ser outra obra, posterior a 2018, o que explicaria o INSS de R$ 62.555,51.
- **Correção sugerida.** Uma linha na ressalva 1 e em T-03.

---

## Conferido sem erro

### Script e cobertura

- `pacote_01.py` roda com `python -I`, só lê e confere com o pacote:
  - R1: 9 linhas, R$ 15.674.986,00 nas duas avaliações;
  - R2: 2 linhas, R$ 18.581.378,60 em 2018;
  - R3: 1 linha, R$ 220.000,00;
  - 8 linhas fora das regras;
  - 20 = 20 linhas na planilha.
- Nenhuma linha da planilha-mestre ficou fora das regras e fora da lista. As 8 da lista são exatamente as 8 "fora das regras".

### R1

- Os 9 valores são iguais nas três seções do PDF de Rafael (Transportadora):
  - balanço do grupo, p. 1–3;
  - "Avaliação patrimonial executada por Valter", p. 4–8;
  - "Balanço patrimonial executado por Valter", p. 13–15.
- O total de Rafael, R$ 16.924.986,00 (p. 16), é igual ao total do grupo mais as quatro diferenças. Não achei os bens da 08 em outros PDFs de Rafael. O Cotiporã do Hotel Princesa é outra matrícula (7.391).
- Nenhum bem da 01 ou da 08 está em `ativos-avaliados-valter.xlsx`.
- Amostra de 5 linhas: chaves, valores e páginas de Rafael (p. 13, 13, 14, 14 e 15) conferidos. Origem 2018 conferida (doc 10, itens 01 e 02; doc 15, itens 06 e 01; doc 17, item 07 "R$ NÃO DIVISÍVEL", fora do total de R$ 24.112.390,00). Tabelas e linhas do doc 08 de 2025 conferidas.
- Ressalva 3: 69,8 × 50.000 = 3.490.000 (+10.000) e 4,5 × 50.000 = 225.000 (+25.000).
- Ressalva 4: R$ 685.000 → R$ 840.000, "CUB [...] + instalações".
- Ressalva 5: faixa de R$ 106.000 a R$ 170.000 (p. 8–9).
- Ressalva 1: os dois lados omitem a "vantagem da coisa feita" de 30 % nos itens 02, 04 e 05.

### R2 e R3 (números)

- R2: mat. 3.754 (R$ 4.950.000) e mat. 17.409 + 6.136 → 29.954 (R$ 13.631.378,60), conferidas no doc 10 de 2018 (itens 06 e 07) e no doc 11 de 2025 (itens 03 e 04). A SEI já tem esses bens no candidato nº 1 da Parte 6.
- R3: R$ 1.784.000 − R$ 220.000 = R$ 1.564.000 (doc 15 de 2018, T4.R3).

### Lista do pacote 2

- Os valores dos 8 itens estão certos.

### Perguntas ao grupo (números e citações)

- **A-01:** preços unitários de 2018 e 2025 corretos; redução de R$ 2.555.437,89 = 12.216.851,89 − 9.661.414,00. Nota I7 confere.
- **A-02:** citação de "Companhia Ipiranga" confere; Petrobras com a Certo; locação com a SIM.
- **A-04:** arq. 4, p. 1, 3 e 4.
- **A-05:** arq. 4, p. 1–2 e 7, × arq. 1, p. 16.
- **A-06:** placa ICW 8342, L 1618, R$ 145.000.
- **A-07:** números dos dois processos.
- **T-01:** R$ 1.135.000 × R$ 1.020.000; notas B196–B199 conferem.
- **T-05:** mat. 18.952, Ed. Dona Ângela Siviero, R$ 149.000.
- **T-07:** mat. 7.395.
- **T-09:** R$ 340.000 (R$ 150.000 em 2018).
- **T-10:** "condenação superior a R$ 1.000.000,00".
- Contagem: 7 perguntas da Abastecedora e 10 da Transportadora.

### Candidatos à Parte 6 (fatos)

- **C-01:** item 03 da SEI, R$ 4.950.000, sem ônus citado; 2.500 m² da mat. 3.754 para estacionamento; prazo até 31/12/2033, com renovação de 36 meses.
- **C-03:** R$ 3.750.000; arq. 10, cl. 3ª, com a citação "créditos hereditários e/ou creditórios"; quitação com bens da SEI (cl. 4ª).
- **C-04:** três lotes de R$ 550.000 "À PAGAR" = R$ 1.650.000, entre eles a comissão (item 8 do contrato); R$ 16.127.652 em "demais compromissos" (doc 11, T50).
- **C-05:** R$ 245.624,79 = 1,5 % × R$ 16.374.986.
- **C-06:** R$ 67.629,90 = 0,7 % × R$ 9.661.414; prazo do INSS em março de 2027 (arq. 1, p. 3).

### Registro de data-base

- **DB-01:** 28/08/2025; 2 % em 31.12.2024.
- **DB-02:** ajuizamento em 31/07/2023; R$ 1.032.648,00; sentença em 14/08/2025; trânsito em julgado em maio/2026 (arq. 1, p. 4).
- **DB-03:** acordo em ago/2020, desistência em fev/2022, conciliação frustrada em maio/2024, defesa em jul/2026, processo "para réplica" (arq. 1, p. 15–16); "desde novembro de 2024" no doc 08.
- **DB-04:** R$ 750.000; multa de R$ 300.000 na faixa de 9.001 a 12.000 m³; volume de 10.847,10 m³; vigência até 30/09/2025.
- **DB-05:** "outubro/2025" (arq. 1, p. 1 e 4).
- **DB-07:** PDF gerado em 21/09/2026 (metadados do PDF) e declaração de Benta (p. 5–6).
- **DB-08:** arq. 10, cl. 3ª, p. 4, e citações.
