# Revisão independente do pacote 5, versão 2 (regra R6 de "possível APP")

28/10/2026. Revisor com contexto novo (P11). Objeto: `pacotes/pacote-05.md` (v2), `_scripts/geofarroupilha_app.py`, `_scripts/pacote05.py`, `geofarroupilha-app-lotes.xlsx` e brutos em `01_fontes/geofarroupilha/`.

Método:
- rodei `pacote05.py`;
- consultei de novo os serviços públicos do GEOFarroupilha (Marco_Regulatorio, camadas 0 a 11; Cadastro, camadas 381 Loteamentos, 272 Quadras e 404 Lotes; Plano Diretor 2015 e 2023);
- remedi os 95 lotes com código próprio, que trata os buracos dos polígonos ESRI, e incluí as camadas 1 (Recursos_Hidricos), 6 e 7 (APP Balneário Santa Rita e Barragem Julieta), que o script não consulta;
- conferi Rafael (xlsx de terceiros e de loteamentos), o arquivo 1 (p.17-30), a planilha-mestre (Resumo e Passivos), o `decisoes-sei.json`, o `log-decisoes.md` e o `00_PLANO.md` (P2, P5, P10/C0).

Meus scripts ficaram no scratchpad da sessão, fora do projeto.

**Resumo: as contas fecham e a identificação dos lotes está certa. O problema é de leitura do mapa.** A camada "Sensibilidade ambiental" é a faixa de 30 m em volta de todos os cursos d'água, e o grupo marcou exatamente os lotes que a tocam. O pacote declara, sem base, que essa camada "não é APP", e é isso que sustenta o nível C de 7 lotes do passivo. Se a leitura do outro lado prevalecer, o efeito da R6 cai de + R$ 11.118.500 para + R$ 2.782.000 (1/6: de R$ 1.853.083,33 para R$ 463.666,67). Corrigir o achado 1 antes de levar o pacote ao Luís.

---

## Achados

### 1. ALTA. "Sensibilidade ambiental, que não é APP" não tem base, e o mapa indica o contrário

- **O que o pacote diz:** os 6 lotes de nível C do Vicentino (e o 003/1301 do S1) "estão só na camada 'Sensibilidade ambiental', que não é APP" (§2). A §0 diz que "o grupo marcou os lotes de acordo com o mapa" e que "os lotes marcados, em geral, tocam" as camadas 2 a 5.
- **Evidência:**
  - **O serviço não traz nenhum metadado.** A descrição do serviço, das camadas e do `iteminfo` está vazia. A camada 10 só tem `objectid` e `shape`. Nada no mapa diz o que é "Sensibilidade_Ambiental".
  - **A geometria diz o que ela é.** Medi a distância da borda da camada 10 até as linhas da camada 1 (Recursos_Hidricos) em 5 áreas: Vicentino 1302, San Marino 1650, Primavera 1199, Jardim Itália 1747 e Perini 1469. A mediana é de 30,0 m nas 5, com p10 de 29,9 m e p90 de 30,1 a 32,0 m. A "Sensibilidade ambiental" é, portanto, a faixa de 30 m de cada lado dos cursos d'água: a largura mínima de APP do Código Florestal (art. 4º, I, a). As camadas 3 (APP 30 m), 4 (faixa de 15 m) e 5 (faixa de 5 m) ficam dentro dela, e as bordas das faixas de 15 m e 5 m ficam a 15,0 m e 5,0 m dos cursos d'água. A leitura mais provável é que a lei municipal (área urbana consolidada, Lei 14.285/2021) reduziu a proteção dentro dessa faixa federal. A faixa continua no mapa como "sensibilidade".
  - **O curso d'água atravessa os lotes de nível C.** A linha da camada 1 cruza o 1302/009 (50 m dentro do lote), o 1302/010 (49 m) e o 003/1301 do S1 (36 m). Ela passa a 0,3 m a 7 m do 1302/011, do 1301/004, do 1302/004 e do 1303/012. Nesses trechos não há nenhuma faixa das camadas 2 a 5: o mapa municipal dispensou a faixa ali, sem explicar por quê (canalização? área consolidada?).
  - **O critério do grupo é a faixa de 30 m.** Os 33 lotes que o grupo marcou (7 próprios e 26 de terceiros) tocam a faixa de 30 m (camadas 2 a 5 ou 10). Dos 62 lotes próprios avaliados cheios, só 2 a tocam, e por lascas de 3 a 4 m² (San Marino 1650/010 e Jardim Itália 1737/003, menos de 1 % do lote), que são ruído de borda. Pelas camadas 2 a 5, só 22 dos 33 lotes marcados aparecem. Os 11 restantes somam R$ 9.651.000 do passivo (7 do Vicentino e o Primavera 006) e R$ 730.000 do ativo (San Marino). Logo, "o grupo marcou de acordo com o mapa" é verdade, mas o mapa que o grupo usou inclui a camada 10. O pacote adota uma leitura mais estreita que a do grupo e não diz isso.
  - **O grupo alega impedimento de construir:** "não sendo possível hoje aprovar projeto de construção sobre os lotes por serem considerados APP" (arq. 1, p.28). Ninguém conferiu se a faixa de "sensibilidade" impede ou condiciona a aprovação, porque o texto da lei não foi lido (§1, ressalvas).
- **Valor em jogo (mesma regra, 50 % e 10 %, contando a faixa de 30 m como APP):**

| Lote | Hoje (camadas 2 a 5) | Com a faixa de 30 m |
|---|---|---|
| Vicentino 1302/009, 1302/010, 1302/011 e 1301/004 (terceiros) | C, 0 % | A, de 54 % a 89 % |
| Vicentino 1302/004 e 1303/012 (terceiros) | C, 0 % | B, 24 % e 42 % |
| Vicentino 1303/014 (terceiro) | B, 40 % | A, 100 % |
| Primavera 1199/006 (terceiro) | C, 0 % | B, 29 % |
| San Marino 1650/012 e 013 (próprios) | B | A, 56 % e 95 % |
| Vicentino 1303/013 (próprio) | B, 33 % | A, 94 % |
| Vicentino 003/1301 (S1, lado do ativo) | C | A, 100 % |

  - Com a faixa de 30 m, o passivo sobe de R$ 4.422.000 para R$ 11.957.500, e o ativo cai de R$ 1.187.500 para R$ 386.500.
  - O efeito da R6 cai de + R$ 11.118.500 para **+ R$ 2.782.000**: menos R$ 8.336.500, ou R$ 1.389.416,67 no 1/6.
  - Leitura intermediária: a faixa de 30 m sozinha leva o lote no máximo ao nível B. Nesse caso, os 6 lotes C do Vicentino e o Primavera 006 vão para B. O passivo fica em R$ 8.570.000, e o efeito em + R$ 6.970.500 (1/6 = R$ 1.161.750).
- **Correção:**
  - Retirar "que não é APP". Escrever o que se sabe: a camada é a faixa de 30 m dos cursos d'água, fica fora das faixas que a lei municipal manteve, e o seu efeito legal não foi conferido.
  - Corrigir a §0. O grupo marcou os lotes que tocam a faixa de 30 m (33 de 33). A R6 conta só as faixas que a lei municipal manteve (22 de 33).
  - Apresentar ao Luís a escolha da leitura como **a** decisão central do pacote, com as três linhas de valor: camadas 2 a 5, intermediária e faixa de 30 m.
  - Pedir com prioridade ao grupo ou à prefeitura o texto da Lei 4.814/2023 e o significado da camada "Sensibilidade ambiental" (pergunta 2). Se a lei cair (risco MP-RS, §1), vale a faixa de 30 m, e a linha "faixa de 30 m" vira o cenário provável.

### 2. MÉDIA-ALTA. "Revisão MP com processo anterior à data-base" não sustenta o nível B do San Marino

- **Evidência:**
  - **A camada 11 (`Rev_MP_APPs`) não tem descrição.** O único campo é `processo_adm` (alias "Processo_Adm"). Os 71 polígonos trazem processos de 2022 a 2026, um deles é "OFÍCIO Nº 103/2024" e outro está vazio. Nada mostra que "MP" seja o Ministério Público, nem se o processo é do MP, do proprietário ou da prefeitura, nem se pede para incluir ou para retirar APP.
  - **No San Marino, o polígono da camada 11 coincide com a faixa de 30 m.** A borda fica a 30,0 m do curso d'água, como a da camada 10. Ali o curso d'água tem faixas de 5 e 15 m, mas elas não chegam aos 3 lotes. O registro traz dois processos, 0.012476/2024-79 e 0.002657/2026-59. É igualmente possível que a revisão tenha **retirado** ou **estreitado** a faixa sobre esses lotes, que estariam em APP em 31.12.2024 (nível A), ou que o processo de 2024 tenha terminado antes da data-base (nível C). O nível B como ponto médio pode ser mantido, mas pela incerteza, e não porque "há processo anterior à data-base".
  - **Isso também prova que o mapa foi alterado depois da data-base** (há processos de 2026 na camada). O serviço não informa data de versão nem histórico de edição. Nos polígonos de revisão, a geometria de hoje pode diferir da de 31.12.2024.
  - **A ressalva da §1 está incompleta.** Ela fala em processos "de 2024 (usado) e de 2026 (não usado)". Há também o 0.014308/2025-07, que cobre 100 % da quadra 1292 do Dalzóchio (lotes 002, 003 e 004) e parte do 0910/O09. Os 4 lotes do Dalzóchio de nível A estão, portanto, numa área em revisão desde 2025. A faixa de 15 m que os põe no nível A pode ser resultado dessa revisão ou ser alterada por ela.
  - **A regra não tem área mínima para a camada 11.** Pela letra da regra, o San Marino 1650/010, um lote cheio fora da lista, entraria no nível B por uma lasca de 3,4 m² de "Revisão MP" com processo de 2024. O Jardim Itália 1737/003 está no mesmo caso, com "OFÍCIO Nº 103/2024", que a expressão regular `(\d{4})-\d\d` do `pacote05.py` nem reconhece. Hoje não muda valor, porque a regra só se aplica aos lotes marcados pelo grupo, mas a regra escrita não diz isso.
- **Correção:**
  - Reescrever o critério: "lote na faixa de 30 m dentro de área em revisão do mapa; sentido e resultado da revisão desconhecidos; nível B pela incerteza".
  - Citar o processo de 2025 do Dalzóchio nas ressalvas.
  - Acrescentar à regra que ela vale só para os lotes listados pelo grupo como APP ou possível APP, com área mínima (por exemplo, 10 % do lote) para a camada 11.
  - Incluir os números dos processos na pergunta ao grupo e à prefeitura.

### 3. MÉDIA. Os limites de 10 % e 50 % criam saltos maiores que a precisão do mapa

- **Evidência:**
  - Perini 018: 236,3 m² em APP 50 m num lote de 466,7 m², ou seja, 50,6 %. Faltam 2,9 m² para o limite, muito menos que o erro de posição de um cadastro urbano ou do ponto da nascente, que está a 36,6 m do lote. O salto vale R$ 80.000 no passivo.
  - Vicentino 1303/016: 13,0 %, a 88 m² do limite de 10 %. O salto vale **R$ 677.500** no passivo, e o pacote não o menciona. Com a faixa de 30 m (achado 1), o lote fica em 48,4 %, perto do limite de 50 %.
  - O outro lado pode pedir o critério proporcional, ou seja, perda igual ao percentual da área do lote em faixa. Pelas camadas 2 a 5, esse critério é **melhor** para o Luís: passivo de R$ 2.783.807, contra R$ 4.422.000, e ativos de Farroupilha de R$ 1.616.909, contra R$ 927.500. Com a faixa de 30 m, seria pior.
- **Correção:**
  - Na §2, marcar os lotes a menos de 5 pontos percentuais de um limite (Perini 018, Vicentino 016) e o valor de cada salto.
  - Dizer por que a R6 usa níveis e não a proporção (simplicidade e incerteza jurídica, não só geométrica), ou mostrar a proporção como alternativa ao Luís.

### 4. MÉDIA. Para os lotes que ficam na empresa, a R6 contraria o C0, e o pacote não diz isso

- **Evidência:**
  - O pacote conclui que os lotes "ficam na empresa" (§2, C0). Para bem que fica, o C0 manda adotar o **maior** valor. Nos 7 lotes, o grupo exclui o valor do total e Rafael cobra o valor cheio (`decisoes-sei.json`, motivo dos 7 lotes; para o Simioni, Rafael dá R$ 520.000). Pelo C0 puro, os lotes ficariam em R$ 2.841.000. A R6 dá R$ 1.187.500.
  - Diferença: R$ 1.653.500, ou R$ 275.583,33 no 1/6 do Luís.
  - O pacote só discute o caso inverso, em que os lotes vêm para nós.
  - A precedência do C2 sobre o C0 foi aprovada para a APP **confirmada** do Lago Negro (log, linha 87), não para a "possível APP".
- **Correção:** acrescentar uma linha à §2: "pelo C0, que manda adotar o maior valor para o bem que fica, seriam R$ 2.841.000; a R6 aplica o C2 e reduz para R$ 1.187.500, o que custa R$ 275.583,33 ao Luís". Pedir que ele confirme que o C2 prevalece aqui.

### 5. BAIXA. Afirmações imprecisas no texto

- §0: "Os lotes marcados, em geral, tocam [as faixas]." São 22 de 33. Os 11 restantes concentram R$ 9.651.000 do passivo, o que não cabe em "em geral" (ver achado 1).
- §6: "os 3 lotes de terceiros e o lote próprio caem em faixa de APP". Os lotes 1292/002 e 1292/003 estão só na **faixa não edificante de 15 m**, não em camada de APP. O 0910/O09 e o 1292/004 estão em APP 30 m. Escrever "APP ou faixa não edificante".
- §1, ressalvas: "é a versão consultada hoje". Acrescentar que o serviço não informa data de versão e que a camada 11 prova edição posterior à data-base (achado 2).
- §2, linha 2 e amostra §4: o Vicentino 1303/013 está a 32,6 % em faixa, mas a 93,7 % na faixa de 30 m, e o curso d'água atravessa o lote por 98 m. Vale citar, pelo achado 1.

### 6. BAIXA. Simioni: a lógica da própria regra o levaria ao nível A

- O pacote reconhece que o TAC trata de uma nascente. Pela regra federal, sem lei municipal de área consolidada em Veranópolis nas fontes, o entorno da nascente é APP num raio de 50 m, cerca de 7.850 m². Isso é mais que a área dos 4 lotes juntos. Se a nascente estiver nos lotes ou junto deles, a R6 aplicada como nos lotes de Farroupilha (50 % ou mais do lote) daria nível A.
- O nível B se sustenta só porque falta a localização da nascente. É favorável a nós e defensável, mas o outro lado pode pedir a planta.
- **Correção:** escrever que o nível B depende de não haver mapa e que, com a localização da nascente, a mesma regra pode levar o Simioni ao nível A.

### 7. BAIXA. Robustez dos scripts (não muda nenhum número hoje)

- **`geofarroupilha_app.py`, filtro por setor:** o filtro só age quando a busca devolve mais de um lote. Se o mesmo número de quadra existir em outro setor e o lote só existir lá, o lote errado seria aceito sem aviso. Hoje a quadra 1747 só existe no setor 24, e o filtro do Jardim Itália é inócuo. A 0910 está nos setores 06 e 15, mas o lote 9 só existe no 06. A 0907 está quase toda no setor 15 (inscrição fiscal), mas o lote 13 fica dentro do polígono "Lot. Firminio Dalzochio". Sugestão: conferir cada lote pela camada 381 (Loteamentos), como fiz (achado verificado, abaixo).
- **Camadas não consultadas:** o script não lê as camadas 1 (cursos d'água), 6 e 7. As camadas 6 e 7 não tocam nenhum dos 95 lotes. A camada 1 é a base do achado 1.
- **Áreas por camada:** `poly()` une todos os anéis, inclusive os buracos ESRI, que têm sentido anti-horário. Há feições com buraco nas camadas 2 a 5 e 10 perto do Vicentino e do Primavera, mas refiz o cálculo tratando os buracos e nenhum percentual mudou. As colunas de área por camada somam as feições em vez de uni-las (o % final usa a união e está certo).
- **Registro das quadras:** o `quadras-sei-*.json` guarda, por quadra, só o resultado da última busca. Não serve de prova da busca por setor.
- **`pacote05.py`:** a expressão regular não lê "OFÍCIO Nº 103/2024" (ver achado 2).

---

## O que conferi e achei correto

- **Identificação, 95 de 95 lotes.** Para cada lote, consultei o ponto interior do polígono encontrado na camada 381 (Loteamentos) e na 272 (Quadras). Todos caem no loteamento certo, com a quadra certa.
  - Casos difíceis:
    - Dalzóchio 0910/O09 = 06.0910.00009, "Lot. Firminio Dalzochio";
    - Dalzóchio 0907/O13 = 15.0907.00013, no mesmo loteamento, apesar do setor fiscal 15;
    - Perini 1469/"19" = 17.1469.00019, "Lot. Perini";
    - Jardim Itália 1747/001, 002, 009, 015 e 016 no setor 24, "Lot. Jardim Itália III";
    - San Marino 1650/012 e 013 e 1676/004 (e o 1342 no setor fiscal 04), "Lot. San Marino";
    - Vicentino quadras 1301 a 1303, "Lot. Industrial Vicentino";
    - Primavera 1193 e 1199, "Lot. Primavera".
  - Nenhum lote foi casado com o lote errado. Contagem: 62 próprios cheios, 7 de possível APP e 26 de terceiros.
- **Medição.** Reproduzi o % (união das camadas 2 a 5) dos 33 lotes, com tratamento de buracos. Coincide até 0,1 ponto percentual com a planilha, por exemplo:
  - Perini 018: 50,6 %;
  - Vicentino 1303/016: 13,0 %;
  - Vicentino 1303/013: 32,6 %;
  - Dalzóchio 1292/004: 64,6 %.
- **Sobreposição de camadas.** O % usa a união: no Dalzóchio 1292/004, as camadas 3 e 4 somam 381,6 m², e a união dá 252 m² (64,6 %).
- **Unidades e sistema de coordenadas.** O serviço declara wkid 31982 (SIRGAS 2000 / UTM 22S) e unidade em metros. A geometria é pedida e devolvida em 31982.
- **Nascentes.** A camada APP 50 m é o raio de 50 m das nascentes: a borda fica a cerca de 49,5 m da nascente no Perini. A contagem de nascentes a até 50 m (Perini 018 e 19) confere.
- **Números do `pacote05.py`, iguais aos do pacote:**
  - ativos de R$ 2.841.000 cheios e R$ 1.187.500 recomendados;
  - passivo de R$ 16.240.000, S1 de R$ 1.887.000 e passivo recomendado de R$ 4.422.000;
  - soma adotada de R$ 153.970.060,60, que passa a R$ 155.157.560,60;
  - efeito da R6 de + R$ 11.118.500 (1/6 = R$ 1.853.083,33) e do S1 de + R$ 1.887.000 (1/6 = R$ 314.500).
  - Por loteamento: Vicentino R$ 2.710.000, Jardim Itália R$ 586.000, Perini R$ 560.000, Dalzóchio R$ 400.000 e Primavera R$ 166.000. A amostra da §4 confere.
- **Sem dupla contagem.** A R6 e o S1 estão separados. Os 4 lotes do S1 saem do passivo antes da regra. A planilha-mestre ainda traz R$ 16.240.000 no passivo, deduzido no total líquido, o que confirma que o S1 é aplicado uma única vez. A soma adotada atual não inclui os 7 lotes, e o LOT-04 (Simioni) continua sem valor adotado.
- **C0.** São 5 dos 7 lotes nas tabelas de "lotes escolhidos pela Valter" de Rafael (San Marino 3, Dalzóchio 1, Vicentino 1), sem valor no total.
- **Citações.** O arq. 1, p.26 (camada do Marco Regulatório no GEOFarroupilha) e p.28 ("faixas de APP de nascentes, córregos...") conferem. Os valores de Rafael por lote (terceiros e os 7 próprios) conferem.
- **Achados da v1 incorporados:** escrow, faixa de mercado do Vicentino, erro de R$ 7.700.000, exceção à P5, separação R6 e S1, lado do ativo do S1, três leituras do Simioni e redação sobre Rafael.

Fontes externas consultadas, sem achar o texto da Lei 4.814/2023: [lista do MP-RS de municípios que reduziram limites de APPs](https://www.mprs.mp.br/media/areas/imprensa/anexos_noticias/municipiosquereduziramlimitesdeapps.pdf); [Santa Maria, nota sobre a Lei 14.285/2021](https://www.santamaria.rs.gov.br/arquivos/baixar-arquivo/conteudo/D28-4550.pdf).
