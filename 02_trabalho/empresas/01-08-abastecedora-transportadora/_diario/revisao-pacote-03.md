# Revisão independente — Pacote 3 (créditos, contingências e reduções, 01 e 08)

28/10/2026. Revisor sem contexto anterior. Não alterei nenhum arquivo além deste.

**O que foi lido:** `pacotes/pacote-03.md`; `_scripts/pacote_03.py`, rodado com `python -I` (saída reproduzida: 11 + 4 + 1 = 16 de 16 linhas; 313.254,69, 184.810,24, 52.209,11, 30.801,71 e 56.666,67); aba "Créditos e contingências" da planilha-mestre (16 linhas, todas as colunas); aba Conferências da Fase 0; `contratos-abastecedora.md`; `registro-data-base.md`, `parte6-candidatos.md`, `perguntas-ao-grupo.md`, `decisoes.json`. Fontes: inventários 2025 docs 00, 01 e 08 (docx e extração; versão antiga comparada por diff); inventários 2018 docs 12 e 15 (o 15 também no Drive); PDFs de Rafael 01 (6 p.) e 08 (17 p.), com texto de todas as páginas; arq. 1, p. 1–4 e 14–16; imagens do arq. 2 p. 4, do arq. 3 p. 1 e do arq. 4 p. 11–12.

**Contagem:** 2 de gravidade alta, 8 média e 7 baixa.

---

## Achados

### 1. ALTA — Cobertura: obrigações e contingências citadas nas fontes ficaram fora da aba e do pacote

O pacote diz que trata "os créditos, contingências e reduções citados nos inventários e nos contratos" e fecha em 16 itens. Faltam três:

- **(a) Processos do Posto Pértile Ltda.** (078/1.14.0000166-1 e 078/1.13.0001323-4, "Ação de Indenização = Ponto e pagamento de aluguéis"). Eles aparecem nas observações do inventário 2018 da 01 (doc 12; `_extracao/2018-12-abastecedora.txt`). Não estão no inventário 2025 e o desfecho é desconhecido. A pergunta A-07 foi retirada, mas a contingência não foi registrada como linha.
- **(b) Arq. 4, p. 12, cl. 22ª e § 1º.** Há um contrato de transporte de combustíveis de 22/03/2021 entre a **Transportadora Irmãos Siviero (08)** e a Transportes Zandoná, "com a intervenção da aqui Locadora". Nele, a 01 "se compromete a arcar isoladamente com quaisquer pendências que eventualmente existir com a contratada Transportes Zandona Ltda." É uma obrigação da 01, e a 08 é parte no contrato. Nenhum dos dois pontos está na aba.
- **(c) "Ônus de eventual dissolução da empresa em relação a parte de algum(ns) sócio(s)"** (docs 01 e 08 de 2025, duas vezes em cada). Ele está só na ressalva 3, sem linha. É o ônus que se dirige justamente à saída do cliente.

**Correção:** incluir as três linhas na aba (no próximo build), sem valor, na regra C1 ("nada deduzido; fato registrado"). Corrigir a contagem do pacote (19 itens, ou 16 + 3 registros). Os totais não mudam.

### 2. ALTA — Sottili: pensão e dano moral não são "cenários alternativos"

Na tabela C1, a linha do dano moral pedido diz: "São cenários alternativos do grupo; não se somam". Na aba, a linha da **pensão** diz "Cenários alternativos, não somar". O doc 08 diz outra coisa: "o pedido é 01 (um) salário mínimo até que o menino completasse 25 anos [...] e, + 1.000 salários mínimos de dano moral". Pensão e dano moral são **cumulativos**. Alternativos são só os dois cenários de dano moral (R$ 1.518.000 ou R$ 627.000).

A conclusão do próprio grupo, "condenação superior a R$ 1.000.000,00", só fecha somando: R$ 364.320 + R$ 627.000 = R$ 991.320, mais juros. A exposição citada vai, portanto, de cerca de R$ 991 mil a R$ 1.882.320, e não "até R$ 1.518.000".

**Correção:** na aba e no pacote, escrever "pensão + um dos dois cenários de dano moral (os de dano moral são alternativos)". Nada muda nos totais, porque nada é deduzido.

### 3. MÉDIA — Rafael ressalva expressamente os passivos (omissão na afirmação central)

A frase "Ele não deduz nenhum passivo" é verdadeira: nenhuma seção dos dois PDFs aplica dedução ou chega a valor líquido. Mas Rafael condiciona a conclusão:
- Rafael 01, p. 5, e Rafael 08, p. 17: "condicionada à conferência da documentação societária, contábil e patrimonial [...] especialmente quanto à inexistência de ativos, passivos ou direitos remanescentes que possam produzir efeitos na apuração patrimonial".
- Rafael 01, p. 6 (declaração de escopo): a apuração definitiva "deverá observar [...] obrigações fiscais, eventuais ônus e gravames".
- Rafael 08, p. 4: concordância "sem prejuízo da necessária conferência documental e jurídica quanto à [...] regularidade registral".

Se o estudo de Rafael chegar ao grupo, este pode citar o próprio lado do cliente para pedir a dedução do INSS, da hipoteca ou do Sottili.

**Correção:** acrescentar ao "Ponto de partida" a frase: "Rafael não deduz, mas ressalva a conferência de passivos, obrigações fiscais, ônus e gravames (01 p. 5–6; 08 p. 4 e 17)". Levar isso também à ressalva 3.

### 4. MÉDIA — O efeito em 1/6 usa os totais do grupo, e não os valores adotados; INSS e 1,5 % se sobrepõem

- As reduções são percentuais "do valor acima apontado". Os R$ 313.254,69 aplicam esses percentuais aos totais do grupo (R$ 9.661.414 e R$ 16.374.986). Os valores adotados nos pacotes 1 e 2 (`decisoes.json`; DB-06 e DB-07) são R$ 12.361.850 (01) e R$ 16.924.986 (08). Sobre eles, as reduções seriam R$ 86.532,95 + R$ 253.874,79 = **R$ 340.407,74**, ou **R$ 56.734,62** no 1/6, e não R$ 52.209,11.
- A frase "as duas reduções anunciadas [...] e o INSS estimado [...] tirariam R$ 52.209,11 e R$ 30.801,71" pode ser lida como soma (R$ 83.010,82). O próprio pacote trata o INSS como "mesmo fato" da redução de 1,5 % (C-05). Os números são alternativos, não cumulativos.

**Correção:** dizer a base de cálculo, dar também o número sobre os valores adotados e escrever "o INSS está contido no mesmo fato; não somar".

### 5. MÉDIA — Sottili: o argumento do "acordo depositado" omite fatos contrários que estão na mesma página; "a sua inclusão" está errado

- O arq. 1, p. 15–16, diz também: "Em fev/2022, a parte autora incluiu uma petição desistindo do acordo"; a conciliação de maio de 2024 foi "inexitosa"; o juiz mandou citar João Sotili por edital antes de examinar o acordo; a defesa foi apresentada em julho de 2026; e o processo "está para réplica". O DB-03 registra isso. Mesmo assim, a linha C1 e a ressalva 2 apresentam "o acordo depositado pela seguradora em 2020" como "a resposta". Se for usado com o grupo, ele responde com a desistência.
- "A sua inclusão no processo só como titular do bem": quem foi incluída é **a empresa**, "pelo fato do bem estar em seu nome" (arq. 1, p. 15, item 5), e não o Luís. Além disso, ser proprietária do veículo costuma ser o próprio fundamento da responsabilidade solidária. Não é uma defesa forte.

**Correção:** reescrever a ressalva 2 assim: "o grupo não deduziu; o acordo depositado em 2020 existe, mas a autora desistiu dele em 2022 e ele não foi homologado; a empresa responde como proprietária do veículo". Usar como argumento só a P5 (nenhum lado deduz) e a ausência de condenação na data-base.

### 6. MÉDIA — Petrobras: a SIM não "assumiu"; a 01 é garantidora; a hipoteca pode estar em mais de um imóvel

- **"Pela locação, a SIM assumiu esses contratos (cl. 21ª)".** O texto (arq. 4, p. 11) diz que a Locatária "**obriga-se a assumir**". O instrumento de cessão com anuência da Petrobras não está nos arquivos (lacuna "a" de `contratos-abastecedora.md`). Corrigir para "obrigou-se a assumir; a cessão não foi apresentada". A mesma cláusula traz um ponto a favor que o pacote não usa: a SIM assume "todo e qualquer ônus ou encargo direta ou indiretamente relacionado a referida cessão".
- **Quem garante.** As linhas C1 dizem só "obrigação do revendedor". A cl. 21ª diz que os contratos da Certo "são garantidos pela ora Locadora", ou seja, pela 01. No arq. 3, p. 2, a Petrobras pode opor à interveniente as exceções contra o revendedor (art. 294 do CC). A exposição da 01 é real, pela garantia, ainda que subsidiária.
- **Qual imóvel.** O § 2º fala em "o(s) imóvel(is) que garante(m)" e a p. 12 em "registrado nos imóveis em comento". O § 3º usa o singular ("no imóvel da Locadora"). O e-mail do "garantidor hipotecário" é o genérico do grupo. "Provavelmente a mat. 3.752" é hipótese.

**Correção:** ajustar a redação. Na ressalva 1, sugerir também a certidão da mat. 3.753 (08), ou ao menos registrar que a hipoteca pode estar em mais de uma matrícula.

### 7. MÉDIA — 01, redução de 0,70 %: o argumento depende de um fato posterior à data-base (P2) não registrado

"O grupo [...] diz que espera não pagar o INSS, cujo prazo vence em março de 2027" se apoia na expectativa de prescrição/decadência depois de 31.12.2024. Na data-base o INSS ainda era exigível. Além disso, os 0,70 % cobrem "a regularização dos imóveis", inclusive a "ampliação que também pende de averbação" (arq. 1, p. 4), e não só o INSS. E, ao contrário da 08, o posto da 01 **inclui** os 30 % de vantagem da coisa feita (R$ 557.634). Não há como alegar desconto em duplicidade.

**Correção:** usar como fundamento principal a P5 (o grupo e Rafael não aplicam a redução). Citar a expectativa de 2027 só como contexto e registrá-la no registro de data-base, ou dizer que ela não fundamenta o tratamento.

### 8. MÉDIA — C2: o aluguel da SIM não é um "ponto a favor na negociação"; pode ser usado contra o cliente

O pacote propõe "registrar como pontos a favor". O aluguel de R$ 15.000/mês dá R$ 180.000/ano, cerca de **1,5 %** do valor adotado da 01 (R$ 12.361.850) e 1,9 % do valor do grupo. A locação prende o imóvel até 31/12/2033, mais 36 meses de renovação, e obriga sucessores e cisões a respeitá-la (cl. 16ª e 17ª). Ela contradiz o "possibilidade de exploração direta" do inventário. Se a renda for levantada na negociação, o grupo ganha base para avaliar o posto pela renda ou para alegar ônus de locação de longo prazo, as duas coisas muito abaixo de R$ 12,36 milhões. É a mesma lógica da nota do C-01 em `parte6-candidatos.md`.

**Correção:** na C2, manter a linha "Aluguel da locação à SIM" só como registro, com a observação "não usar como argumento". "Ponto a favor" vale só para o restaurante e para o crédito Joveni.

### 9. MÉDIA — Perguntas aprovadas T-09 e T-01 podem dar argumento ao grupo, contra a diretriz

- **T-09 (Joveni):** se for adotada a opção (a), recomendada, o crédito fica sem valor qualquer que seja a resposta. A pergunta não traz ganho e pode revelar perda da ação ou sucumbência, que seria passivo da 08. Ela só faz sentido se o Luís escolher a (b) ou quiser usar o crédito na negociação.
- **T-01 (R$ 1.135.000):** pede a composição contábil dos itens 03–05. O item 03 agora está em R$ 1.000.000 (Rafael), o dobro do valor do grupo. A resposta pode convidar o grupo a "acertar" os itens 03–05 pelo valor contábil.

**Correção:** marcar as duas na coluna "Cuidado" e ligar a T-09 à escolha (a) ou (b). O Luís decide se mantém.

### 10. MÉDIA — "Depois deste pacote" diz que as decisões de valor estão completas, mas ainda há itens de data-base pendentes

No `registro-data-base.md`, os itens DB-01 (rateio sem o Rovílio), DB-05 (termo "outubro/2025") e DB-08 (créditos hereditários sobre São Brás) continuam "*pendente*" e não são tratados pelo pacote 3. O DB-08 pode afetar a titularidade ou o valor de R$ 3.750.000 na 08. A lista de pendências também não inclui a alternativa da T-04, retirada: as certidões de São Brás (mat. 459, 460, 461, 558 e reg. 15.922).

**Correção:** listar DB-01, DB-05 e DB-08 e essas certidões em "Pendências abertas", ou dizer onde elas serão decididas.

### 11. BAIXA — C-05 pode ser quantificado e não vale só para o pavilhão

Os 30 % ficaram fora também na rampa (item 04) e na borracharia (item 05), pelo mesmo motivo. O método é o do posto da 01, onde 30 % × (439.275 + 769.505 + 650.000) = R$ 557.634. Aplicado à 08, a vantagem que ficou fora é 30 % × (1.427.800 + 300.000 + 220.000) = **R$ 584.340**. Isso é mais que a redução de 1,5 % (R$ 245.624,79, ou R$ 253.874,79 sobre o valor adotado). **Correção:** trocar "o pavilhão" por "o pavilhão, a rampa e a borracharia" e citar o número.

### 12. BAIXA — Argumento a favor não usado: o doc 00 do grupo trata os totais como patrimônio líquido

O doc 00 de 2025 dá a cada sócio a participação "do (PL) patrimônio líquido de cada empresa" e usa os totais brutos (R$ 9.661.414 e R$ 16.374.986), sem dedução. Isso contradiz o "não é líquido" das observações. **Correção:** acrescentar ao "Ponto de partida".

### 13. BAIXA — "Relatório financeiro" do grupo não examinado

O grupo diz que os inventários "bem assim o relatório financeiro" retratam o grupo (arq. 1, p. 1 e p. 15, item 4). Esse relatório não está entre as fontes. A afirmação "nenhum dos dois lados desconta" vale para os inventários e para os PDFs de Rafael. **Correção:** delimitar a afirmação com essa ressalva.

### 14. BAIXA — 01: a soma "exata" usa o item recomposto

Na versão atual do doc 01, a célula de valor do item 01 está vazia ("R$"). A soma é recomposta pela descrição (nota da Fase 0; na versão antiga a célula trazia R$ 9.661.414,00). **Correção:** escrever "soma dos bens (item 01 recomposto pela descrição)".

### 15. BAIXA — Observação da linha do comodato desatualizada

A aba ainda diz "Divergência a perguntar", mas a A-05 foi retirada pela diretriz. **Correção:** atualizar no próximo build, que não deve ser rodado agora.

### 16. BAIXA — Referência "Resumo B196–B199" ambígua

A planilha-mestre também tem uma aba "Resumo", de 14 linhas. As células estão na aba "Resumo" do "Comparativo Patrimonial Siviero 2018 x 2025" (conferidas: B198–B199 trazem "Não entendi a parte da observação"). **Correção:** dizer o nome do arquivo.

### 17. BAIXA — DB-04: a vigência pode ter ido além de 30/09/2025; a bonificação recebida ajuda o valor de Rafael

- Pelo arq. 2, cl. 3.2 e 3.3, se pelo menos 80 % do volume tiver sido comprado, o contrato se prorroga por um quinto do prazo, ou por prazo indeterminado se a relação continuar. O fim do contrato pode ser posterior a 30/09/2025. Vale citar isso na ressalva 1.
- A 01 recebeu até R$ 750.000 como cessionária, para reforma da cobertura e da imagem da pista (arq. 3, p. 1). Isso apoia a "modernização" que Rafael usa para adotar R$ 12.361.850 (Rafael 01, p. 2). Pode ser usado internamente.

---

## Conferido sem erro

- **Totais e somas:** 01 do grupo R$ 9.661.414,00 (7.245.000 + 439.275 + 769.505 + 650.000 + 557.634); 08 do grupo R$ 16.374.986,00 (13 itens); 08 de Rafael R$ 16.924.986,00 (diferença de 550.000 = itens 03, 11, 12 e 13); Rafael 01 adota um valor único, R$ 12.361.850,00 (p. 2–3).
- **Nenhuma dedução aplicada:** nenhum dos dois lados desconta nada nos inventários 2025 (01 e 08, versão atual e antiga), em nenhuma das páginas dos PDFs de Rafael 01 e 08, nem no doc 00. O 1/6 de Rafael na 01, escrito "R$ 2.060,166", é 142,33 menor que 12.361.850/6. É erro aritmético (12.361.000/6), não dedução.
- **Reduções e INSS:** 0,70 % × 9.661.414 = 67.629,90; 1,5 % × 16.374.986 = 245.624,79; soma 313.254,69; 1/6 = 52.209,11. INSS 122.254,73 + 62.555,51 = 184.810,24; 1/6 = 30.801,71. Joveni: 340.000/6 = 56.666,67.
- **Citações literais:** "o valor real acima apontado não é líquido"; "NÃO significa que este será o valor ajustado"; "Agrega ao valor do Posto"; "condenação superior a R$ 1.000.000,00"; "registrada no imóvel da Locadora" (arq. 4, p. 12).
- **Sottili:** cálculos do doc 08 (1.518 × 12 × 20 = 364.320; 1.518 × 1.000; 150 SM = 627.000).
- **Joveni:** R$ 150.000 no doc 15 de 2018, conferido também no docx do Drive.
- **Páginas e fatos do arq. 1:** p. 1 e 3 (o grupo não se opõe às certidões); p. 3–4 (INSS, março de 2027); p. 4 (Lucchetta: R$ 1.032.648,00, ajuizada em 31/07/2023, improcedência em 14/08/2025, trânsito em maio/2026); p. 15 (sem cadastro de obra; acordo em ago/2020); p. 16 (comodato afirmado).
- **Contratos:** arq. 2, p. 4–5 (faixa de R$ 300.000 para 10.847,10 m³; redução proporcional em 5.2.1); arq. 3, p. 1 (R$ 750.000; Certo como revendedor; Siviero como interveniente anuente e cessionária; obrigação de restituir do revendedor); arq. 4, p. 1 e 4 (R$ 15.000 em 2022; 2.500 m² da mat. 3.754).
- **Classificação e cobertura do script:** 16 de 16 linhas classificadas corretamente nas regras C1, C2 e C3 pela natureza; nenhuma em duplicidade. O script só lê arquivos.
- **Opção Joveni:** (a) pela P5 está correta, porque os dois lados dão zero (o grupo fora do total, Rafael sem menção). (b) exige justificativa de caso extremo, como dito. O efeito de +R$ 56.666,67 está certo.
- **C3:** itens 03–05 somam R$ 1.020.000 pelo grupo e R$ 1.520.000 com o equipamento de Rafael; a diferença de R$ 115.000 não é explicada.
- **DB-02** como "fato posterior que confirma" está correto. **DB-03** e **DB-04** como ressalva são coerentes com o registro, com as ressalvas dos achados 5, 6 e 17.
