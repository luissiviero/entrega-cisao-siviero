# Revisão independente — Pacote 2 (exceções 01 e 08)

28/10/2026. Revisor de contexto novo. Objeto: `pacotes/pacote-02.md` e `_scripts/pacote_02.py`.

O script foi rodado com `python -I`. Ele só lê: as datas de gravação da planilha-mestre e da fase 0 não mudaram, e o `extract_fase0` importado tem guarda `__main__`. A saída do script bate com todos os números do pacote.

Fontes conferidas:
- inventários 2018, docs 12 e 15, lidos direto nos .docx do Drive;
- extrações 2025-00, 2025-01 e 2025-08;
- Rafael 01, p. 1–6, e Rafael 08, p. 4–17;
- arq. 1, p. 1–4 e 14–16;
- arq. 3, p. 1, na imagem `p01.png`;
- `ativos-avaliados-valter.xlsx`;
- abas "Rafael adotado" e "Rafael faixas";
- `perguntas-ao-grupo.md`, `registro-data-base.md` e `parte6-candidatos.md`.

Contagem: **1 alta, 10 médias, 6 baixas.**

---

## Achados

### 1. ALTA — "3 das 4 pendências de 2018 ficam abertas" contradiz a P4 e o próprio item 6

- **Evidência.** Em "Depois deste pacote", o pacote diz que ficam abertas três pendências (apto, sala e móveis) e que "o caminhão fica sem valor próprio". Mas a tabela B, item 6, classifica o caminhão como "Pendência P4, sem valor próprio" e manda a pergunta A-06 (destino). A P4 exige destino documentado. Decidir que o caminhão não tem valor próprio resolve o **valor**, não o **destino**. Nenhuma fonte documenta o destino do ICW 8342: não há nada nos inventários 2025, em Rafael, no arq. 1 nem nos contratos.
- **Correção.** Escrever: "ficam abertas as 4 pendências P4 de destino. O caminhão já tem tratamento de valor (dentro do número de Rafael), e o destino dele continua pendente (A-06 ou outra prova)."

### 2. MÉDIA — "A bonificação Petrobras foi liberada por etapas de obra" não tem base

- **Evidência.** O arq. 3, p. 1, item V, se chama "Forma de Liberação da Bonificação". Ele é um **cronograma contratual**: parcela "após" a reforma da cobertura da pista, "após" a implantação da imagem e "após" a conclusão das obras. O contrato é de 14/12/2020 e não comprova que as obras foram feitas nem que as parcelas foram pagas. A extração conferida diz o mesmo: "Liberação (item V)… após…". Nenhuma fonte traz recibo, TED ou termo de conclusão.
- **Correção.** Escrever: "O contrato de 14/12/2020 (arq. 3, p. 1, item V) condiciona a liberação a etapas de obra (…). É indício de obra prevista depois de 2018. Nas fontes não há prova de que as parcelas foram liberadas." Ver também o achado 9: o mesmo documento é a fonte das contingências de R$ 750.000 e R$ 300.000.

### 3. MÉDIA — Scania IJZ 5399: Rafael não fica "no piso ou dentro das próprias faixas"

- **Evidência.**
  - O bem é um caminhão que "Trabalha para a SEI – prancha" (inventário 2025 e 2018; Rafael p. 10).
  - Para o caminhão com prancha, as faixas do próprio Rafael são R$ 200.000–310.000 (p. 10 e 11) e, para o conjunto, R$ 240.000–280.000 (p. 12). Os R$ 160.000 dele ficam **abaixo** dessas faixas.
  - Os R$ 160.000 só cabem nas faixas do cavalo sozinho: R$ 103.000–162.000 e R$ 150.000–180.000 (p. 11).
  - Por isso, a frase "os valores dele ficam no piso ou dentro das próprias faixas" é exata para as carretas, mas não para o Scania.
- **Correção.** Escrever: "Scania: R$ 160.000 cabe na faixa do cavalo sozinho e fica **abaixo** da faixa do caminhão com prancha (R$ 200.000–310.000), que é a configuração descrita." Isso também reforça o C1 a favor do cliente: o valor de Rafael é conservador para o bem descrito.

### 4. MÉDIA — "Assinado por todos os sócios, inclusive Valter" não tem base, e o argumento vai contra o cliente

- **Evidência.**
  - O inventário de 2018 da Transportadora é um .docx (`15. 1. … Dezembro de 2018.docx`). Ao pé, ele traz os nomes **digitados** dos sete sócios. Não há imagem de assinatura nem versão assinada nas fontes.
  - O arq. 1, p. 1, mostra que o cliente é sucessor de Valter. A frase, portanto, diz que o lado do cliente concordou com a doação. Isso enfraquece a proposta de incluir o apartamento (achado 8).
- **Correção.** Escrever: "o inventário de 2018 (.docx, sem assinatura nas fontes) traz ao pé os nomes de todos os sócios, inclusive Valter." Deixar explícito que esse ponto é argumento **do outro lado**.

### 5. MÉDIA — A alternativa intermediária fere também a P3, não só a P5

- **Evidência.**
  - R$ 12.216.851,89 é **valor do inventário de 2018** (doc 12, item 01). A P3 diz que 2018 é benchmark de existência, não de valor.
  - O pacote marca a alternativa só como caso extremo da P5 e afirma que ela "tira do outro lado o argumento do caminhão". O argumento central do outro lado continua intacto: "2018 não é valor" (P3). A alternativa elimina só o argumento menor, de R$ 145.000.
  - A frase "Não é nenhum dos dois valores de 2025" também é imprecisa, porque o valor de Rafael também não é de 2025: ele vem do balanço de 2018.
- **Correção.** Escrever: "caso extremo da P5 **e** exceção à P3 (usa valor de 2018). Tira o argumento do caminhão, mas não o da P3." Trocar "dos dois valores de 2025" por "dos dois valores em comparação (grupo 2025 e Rafael)".

### 6. MÉDIA — O argumento dos R$ 1.135.000 (equipamentos) não sustenta o valor de R$ 1.000.000

- **Evidência.**
  - O doc 08 de 2025 diz só que se acrescentaram R$ 1.135.000 ao valor do balanço (código 1060), referentes aos itens 03–05.
  - A nota não está explicada: os itens 03–05 somam R$ 1.020.000 e sobram R$ 115.000, o que é justamente a pergunta T-01.
  - A nota fala de diferença contábil sobre o balanço, não do valor de mercado dos equipamentos. Com o equipamento a R$ 1.000.000, os itens 03–05 somariam R$ 1.520.000, que também não bate com a nota.
  - Citar no C1 uma nota que o próprio estudo ainda não entende é frágil.
- **Correção.** Tirar o argumento do C1, ou rebaixá-lo para "nota contábil não explicada (T-01), sem efeito no valor por ora". O C1 fica com estes pontos: nenhum dos dois lados tem nota ou laudo; o grupo repete o valor de 2018 sete anos depois; "estimativa por vistoria in loco" (arq. 1, p. 16).

### 7. MÉDIA — Falta um argumento forte do outro lado no posto: a locação à SIM

- **Evidência.**
  - O imóvel está locado à SIM de 01/01/2022 a 31/12/2033, com renovação única de 36 meses (arq. 4, p. 1 e 3).
  - O aluguel inicial é de R$ 15.000 por mês (arq. 4, p. 1 e 4), cerca de R$ 180.000 por ano. Isso é cerca de 1,5 % ao ano sobre R$ 12,36 mi e cerca de 1,9 % sobre R$ 9,66 mi.
  - A descrição que os dois lados usam ainda fala em "contrato com a Companhia Ipiranga… cumprido" e em "possibilidade de exploração direta", o que não vale com a locação em vigor até 2033/2036.
  - O grupo pode usar a renda e o ônus da locação longa contra qualquer valor acima do dele. O pacote não lista esse argumento.
  - A conta acima é só uma leitura da renda declarada, não uma reavaliação.
- **Correção.** Acrescentar o argumento em "Contra", com a resposta possível:
  - o aluguel é de 2022, com reajuste anual e revisão trienal a preço de mercado (arq. 4, p. 4);
  - restaurante, borracharia e rampa somam aluguel (cl. 3ª, § 3º);
  - os dois lados avaliaram por custo, não por renda.
  
  Fica só no documento interno; não levantar com o grupo.

### 8. MÉDIA — Apartamento 102: incluir o imóvel pode sair "líquido zero" e contraria um acordo com o nome de Valter

- **Evidência.**
  - Em 2018, o doc 15, item 09, registra: "Conforme acordo dentre as partes, este imóvel foi doado para a Senhora Geni Siviero Mattiuz…". O valor sai do total (R$ 1.784.000 − 220.000).
  - Se a matrícula ainda estiver em nome da 08, o grupo alegará que existe a **obrigação** de transferir o imóvel a Geni por acordo. Ativo e obrigação se anulariam.
  - O acordo leva o nome de Valter, antecessor do cliente.
  - O pacote não registra esse contra-argumento. Também não diz que o destino já está, em tese, "documentado" pelo próprio inventário de 2018: a doação foi declarada, e o que falta é a prova registral.
- **Correção.** No item 7, acrescentar o contra-argumento: "obrigação de doar por acordo assinado (nomes) pelos sócios, inclusive Valter: ativo e obrigação se anulam". Rebaixar a proposta de inclusão para "só se a certidão mostrar o imóvel na 08 **e** não houver instrumento de doação ou promessa". Para a P4, registrar o destino como "declarado em 2018, pendente de prova registral".

### 9. MÉDIA — Trechos que entregariam argumento ao grupo, e a pergunta A-06, que mudou de natureza

- **Evidência (o pacote não está marcado como interno).** Se o texto for enviado ou parafraseado ao grupo, entregaria:
  1. toda a seção "Contra" do posto, inclusive a concessão expressa "o caminhão já está dentro desse número";
  2. a alternativa de R$ 12.216.851,89, que mostra disposição de ceder R$ 144.998,11;
  3. "Rafael não fez avaliação própria";
  4. "declaração de Benta é de parte interessada";
  5. "anúncios de 2026";
  6. "provavelmente depreciados, e talvez já estejam nas instalações" (móveis);
  7. "assinado… inclusive Valter" (achado 4);
  8. o uso do arq. 3 como prova de modernização. O mesmo contrato é a fonte da restituição de R$ 750.000 (com juros de 1 % ao mês e multa de 10 %), da multa de R$ 300.000 e da hipoteca. Apresentá-lo ao grupo abre a discussão do passivo que a diretriz manda evitar. A pergunta A-01, já aprovada, cita as etapas da bonificação, e o Luís deve saber disso;
  9. a ampliação "pendente de averbação", que puxa INSS e regularização (passivo da 01).
- **A-06.** Ela foi aprovada sem "Cuidado" antes de o pacote recomendar o valor de Rafael. Agora, uma resposta do tipo "vendido" ou "baixado" vira prova direta de que o número de Rafael inclui um bem inexistente: −R$ 145.000 na 01 e −R$ 24.166,67 no 1/6. Pela diretriz ("não se pergunta… nada que possa reduzir valor"), a A-06 passou a ser pergunta de "Cuidado". Há conflito entre a P4 (destino) e a diretriz, e o pacote não o expõe.
- **Correção.**
  - Marcar o pacote como "uso interno — não enviar ao grupo".
  - Separar o que pode ir ao grupo (a recomendação e os argumentos a favor) do que é só para o Luís.
  - Levar ao Luís a decisão sobre a A-06: (a) manter, aceitando o risco de −R$ 145.000; (b) retirar e documentar o destino por outra via, como consulta de placa no DETRAN pelo próprio cliente; (c) manter sem citar o valor de 2018.

### 10. MÉDIA — P2: o pacote adota valores fora da data-base sem fechar DB-06 e DB-07

- **Evidência.**
  - Os valores recomendados do posto e dos itens 08-03, 11, 12 e 13 são exatamente os que o `registro-data-base.md` marca como fora de 31.12.2024. DB-06 é o balanço de 2018. DB-07 são os anúncios de 2026 e a declaração sem data.
  - As duas linhas estão com tratamento "*pendente*".
  - O pacote cita DB-06 e DB-07 só como argumento do outro lado. Não diz que a aprovação fecha o tratamento delas, e "Depois deste pacote" não as menciona.
- **Correção.** Acrescentar: "se aprovado, DB-06 e DB-07 recebem o tratamento 'valor de Rafael adotado pela P5 + diretriz; data do valor registrada como 2018 (DB-06) e 2026 (DB-07)'."

### 11. MÉDIA — Item 9 (móveis): tratamento inconsistente com os itens 7–8 e justificativa especulativa

- **Evidência.**
  - Para o apartamento e a sala, o pacote propõe incluir pelo valor de 2018 se o bem ainda for da 08 (caso extremo). Para os móveis, recomenda "sem valor", com dois motivos sem fonte: "provavelmente depreciados" e "talvez já estejam nas instalações".
  - A diretriz manda a versão mais alta. "Provavelmente depreciados" é argumento do outro lado.
  - O único motivo defensável é evitar dupla contagem com o "+ instalações" da sala 7.393 (Rafael p. 14; doc 08 de 2025, item 06), e mesmo esse é hipótese a perguntar (T-06).
- **Correção.** Escrever: "Pendência P4. Se a T-06 indicar que os móveis ainda são da 08 e **não** estão no '+ instalações' da sala, mesmo tratamento dos itens 7–8 (R$ 30.000, caso extremo). Se estiverem lá, sem valor próprio (evita dupla contagem)." Retirar "provavelmente depreciados".

### 12. BAIXA — Página da declaração de Benta

- **Evidência.** A tabela diz "declaração na p. 6", e o texto diz "(p. 5–6)". A declaração começa no fim da p. 5 e o valor está na p. 6.
- **Correção.** Uniformizar em "p. 5–6 (valor na p. 6)".

### 13. BAIXA — "arredondado"

- **Evidência.** De R$ 12.361.851,89 para R$ 12.361.850,00 não houve arredondamento ao real (daria R$ 12.361.852), e sim corte à dezena.
- **Correção.** Escrever "(R$ 12.361.851,89, sem os R$ 1,89)" ou "(truncado)".

### 14. BAIXA — Script: o total geral usa `max()` também nos itens já decididos

- **Evidência.** Em `pacote_02.py`, `rec_all` soma `max(grupo, Rafael)` em todas as linhas de 2025, em vez de usar "Valor adotado" quando ele existe. O resultado sai certo só porque os 9 itens decididos têm grupo = Rafael = adotado. O total por empresa já usa o valor adotado.
- **Correção.** Usar a mesma expressão do total por empresa.

### 15. BAIXA — Quota de 1/6 × data-base

- **Evidência.** Em 31.12.2024, Rovílio ainda tinha 2 % da 01 e da 08 (DB-01; doc 12 e doc 15 de 2018: 16,333 % para cada irmão). Por isso, "1/6 = +R$ 541.739,33" depende da P1.
- **Correção.** Acrescentar uma nota: "quota de 1/6 pela P1; ver DB-01".

### 16. BAIXA — Faixa ausente na aba "Rafael faixas" (item 08-13)

- **Evidência.** Rafael, p. 12, traz "Graneleiro: Entre R$ 50.000 e R$ 60.000", e a aba não tem essa linha. O pacote não é afetado, porque a faixa geral é R$ 40.000–85.000.
- **Correção.** Incluir a linha na próxima regeneração da fase 0.

### 17. BAIXA — O escopo da declaração de Benta é mais amplo que o item 03

- **Evidência.** A declaração cobre "os bens, equipamentos, materiais e demais ativos vinculados à oficina mecânica" (Rafael, p. 5–6), não só o item 03. O grupo pode tentar ler o R$ 1.000.000 como total da oficina, contra o pavilhão de R$ 1.427.800. O risco é pequeno, porque Rafael concorda com os itens 01, 02, 04 e 05 (p. 4–6).
- **Correção.** Uma linha em "o outro lado vai alegar", com a resposta: Rafael aplica a declaração só ao item 03 (p. 13).

---

## Conferido sem erro

- **Valores da tabela A** (grupo, Rafael, recomendação, efeito), conferidos contra a planilha-mestre, os .docx de 2018 e as extrações de Rafael (01 p. 3; 08 p. 13 e 15) e do inventário 2025: R$ 9.661.414 / 12.361.850; 500.000 / 1.000.000; 140.000 / 160.000; 30.000 / 50.000; 30.000 / 40.000. Efeitos: +2.700.436, +500.000, +20.000, +20.000, +10.000.
- **Diretriz e P5:** as 5 recomendações são o maior dos dois valores e coincidem com o valor de Rafael. Nenhum valor recomendado fica fora dos dois. Os únicos valores fora dos dois (a alternativa do posto e os valores de 2018 para o apartamento e a sala) estão marcados como caso extremo (ver achado 5 sobre a P3).
- **Totais.**
  - 01: R$ 9.661.414 → R$ 12.361.850.
  - 08: R$ 16.374.986 → R$ 16.924.986, igual ao total de Rafael, p. 16.
  - Juntas: +R$ 3.250.436,00. Ao 1/6: +R$ 541.739,33 (3.250.436 / 6 = 541.739,333…).
  - 14 itens decididos = R$ 29.286.836,00 (9 já decididos, R$ 15.674.986, + 5 deste pacote).
- **Posto, 2018 × 2025.**
  - 2018: posto R$ 12.216.851,89 + caminhão R$ 145.000 = R$ 12.361.851,89 (doc 12, conferido no .docx).
  - 2025: 7.245.000 + 439.275 + 769.505 + 650.000 + 557.634 = 9.661.414.
  - Preços unitários: terreno de R$ 2.500 para R$ 2.100/m². Construções e equipamentos também caíram: R$ 2.154,09 → R$ 1.500; R$ 1.618,38 → R$ 1.100; R$ 1.000.000 → R$ 650.000. A referência "CUB agosto de 2017" foi mantida nas duas versões.
  - As áreas de 292,85 m² e 699,55 m² são iguais em 2018 e 2025.
- **Alternativa:** R$ 144.998,11 abaixo de Rafael e R$ 2.555.437,89 acima do grupo. Conta certa.
- **Citações de Rafael 01, p. 2:** "o valor previamente estabelecido", "sem prejuízo de eventual avaliação técnica independente", "modernização e adequação estrutural, operacional e normativa" e Rede SIM. O balanço é o de 2018 apresentado por Valter. Na p. 3, ao lado de R$ 12.361.850, Rafael transcreve a descrição de 2025.
- **Arq. 3, p. 1:** as citações "reforma cobertura pista", "Implantação da Imagem Pista" e "Conclusão das Obras" são literais. O problema é só o verbo "foi liberada" (achado 2).
- **Arq. 1:**
  - p. 3–4: "inclusive da ampliação que também pende de averbação";
  - p. 1 e 3: o grupo não se opõe a que as certidões sejam tiradas pelos sucessores, que arcam com os custos;
  - p. 16: "estimativa em razão da quantidade e qualificação dos equipamentos por vistoria in loco… não dispondo de notas".
- **Equipamentos:** o grupo tem R$ 500.000 em 2018 (doc 10, item 03) e em 2025. Rafael mostra R$ 500.000 na seção de avaliação (p. 5) e adota R$ 1.000.000 no balanço (p. 13). A declaração de Benta fala em "aproximadamente R$ 1.000.000,00" (p. 6).
- **Faixas das carretas:** IQB 0612, R$ 50.000–90.000 (p. 9; subfaixas na p. 10), com Rafael no piso. IJZ 8207, R$ 40.000–85.000 (p. 12), com Rafael no piso. Faixas do cavalo do Scania na p. 11. Os três veículos repetem o valor de 2018 no inventário do grupo.
- **Bens de 2018** (valores e chaves): ICW 8342, 1995, R$ 145.000 (doc 12, item 02); sala mat. 18.952, Ed. Dona Ângela Siviero, R$ 149.000 (doc 15, item 05); móveis R$ 30.000 (doc 15, item 07); apartamento 102, mat. 7.395, R$ 220.000, doado a Geni "conforme acordo dentre as partes" e deduzido do total (R$ 1.784.000 − 220.000 = R$ 1.564.000). Nenhum dos quatro aparece no inventário 2025 da 01 ou da 08, nem em Rafael.
- **Sala 7.393** "+ instalações": consta em Rafael p. 14 e no doc 08 de 2025, item 06.
- **Lista de Rafael:** nenhum bem da 01 ou da 08 está em `ativos-avaliados-valter.xlsx` (C0 → maior valor). Rafael 01, p. 5: "não há interesse na sua inclusão como ativo selecionável".
- **Perguntas:** A-01, A-06, T-01, T-05, T-06 e T-07 têm a situação citada corretamente. A-03 e A-07, que poderiam criar passivo, constam como retiradas, como diz o pacote.
- **Lista do pacote 3** (Lucchetta, Sottili, Joveni, Petrobras, hipoteca, reduções, INSS, comodato, aluguéis, nota de R$ 1.135.000): coerente com a aba "Créditos e contingências" da planilha-mestre.
