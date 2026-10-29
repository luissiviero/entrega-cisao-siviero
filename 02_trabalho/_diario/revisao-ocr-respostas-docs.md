# Revisão independente – OCR das respostas de 13/08/2026 (arquivos 2 a 16)

Data: 28/10/2026. Revisor independente (contexto novo).

- PDFs: `01_fontes/respostas-docs-2026-08-13/`, arquivos "2 - " a "16 - " (15 arquivos, 100 páginas).
- OCR revisado: `02_trabalho/ocr-respostas-docs/ocr.jsonl` (Tesseract 5.4.0, por, 300 dpi, psm 1).
- Como conferi: cada página foi renderizada com PyMuPDF (110–150 dpi; recortes a 200–250 dpi onde o texto era pequeno, girado ou manuscrito). Comparei a imagem com o campo `texto` do jsonl, número por número. Os PNGs e o script ficaram fora do projeto, na pasta temporária `scratchpad/ocr-review`. Nenhum arquivo de origem foi alterado.
- Amostra: 22 páginas de 10 arquivos, com cerca de 450 números conferidos.

---

## 1. Tabela por página

"Números conferidos" é uma contagem aproximada: valores, áreas, datas, matrículas, CPF/CNPJ, lotes e quadras, selos, e também a numeração das cláusulas. Nas divergências, o formato é **imagem → OCR**.

| # | Arq. | Pág. | Conf. | Tipo de conteúdo | Núm. conf. | Divergências (imagem → OCR) | Busca? |
|---|---|---|---|---|---|---|---|
| 1 | 10 | p.1 | 91,1 | Contrato digitado: qualificação de Rovílio, da Abastecedora e da Transportadora (CPF, RG, CNPJ, NIRE, CEP) | ~30 | **Nenhuma.** Só erros no texto: "Pedro S�iviero", "Estadua!". O carimbo "FOLHA 1/9" não foi lido. | Sim |
| 2 | 10 | p.5 | 91,3 | Contrato digitado: forma de pagamento das quotas (8 valores, lotes e quadras, datas) | ~22 | **Nenhum valor divergente** (587.510,00; 550.000,00 ×3; 75.000,00; 60.000,00; 432.072,00; 110.000,00; lotes 10/Q02, 12/Q237, 15/Q252, 07/Q04; 01/11/2012, 25/10/2013). Os itens "3.)" a "9.)" foram perdidos (só aparece "2.) 4.)" no início). **"Ipê I" → "Ipê \|"** e **"Ipê II" → "Ipê Il"/"!l"**. | Sim, com ressalva sobre Ipê I/II |
| 3 | 10 | p.10 | – | Verso: só a transparência da página anterior e um carimbo parcial do Tabelionato | 0 | OCR vazio, o que está correto: a página não tem conteúdo legível próprio. | n/a |
| 4 | 10 | p.11 | 78,0 | Selos de reconhecimento de firma, impressos girados 180° | ~25 | O selo de Passo Fundo ficou quase todo perdido: "MARISTELA PRETTO SIVIERO", "0414.02.2300005.18785", R$ 10,30, R$ 2,90 e R$ 13,20 não aparecem; "23 de setembro de 2025" → "23 d�"; CEP 99010-021 e telefone perdidos. Selos de Veranópolis: "0736.02.0800004.46865" → "0/36,02.0800004 .46865"; "R$ 5,80" → "R$ 580"; "Elcio" → "FElcio". Os demais números estão corretos (46862, 46863, 46864, 46866, R$ 10,30, R$ 2,90, R$ 20,60, 24/09/2025). | Parcial: "Maristela" não é encontrável |
| 5 | 15 | p.2 | 90,9 | Contrato de locação Oleoplan: confrontações, áreas, prazo, aluguel | ~22 | **"1.942,03m²" → "1.,942,03mM?"**; **"4.792,03m²" → "4,792,03m?"** (separadores trocados). Em toda a página "m²" sai como "m?" e "§" sai como "&". Corretos: 67,00; 17.409; 29,60; 6.136; 2.850,00; 72,75; 01/08/2009; 31/07/2014; R$ 14.000,00; ag. 0450. | Sim |
| 6 | 15 | p.8 | 61,0 | Página de assinaturas com dois carimbos de autenticidade | ~25 | **O bloco da locatária foi perdido inteiro**: "OLEOPLAN S/A – Óleos Vegetais Planalto, Rodovia RSC 470, Km 175,3, nº 3.482, Bairro Valverde" não aparece, e a palavra OLEOPLAN não está na página. Carimbo 1: "0736.01.0800005.28582" → "0756,0f 0800005 28582"; 28583 perdido; "2 de outubro de 2009" perdido; **"R$ 7,60" → "R$ 7,50"**; R$ 0,40 perdido. Carimbo 2: "27427" → "0%000088%.274"; 27428, "24 de setembro de 2009" e R$ 7,60 perdidos. "Costella" → "Covleila"/"Costelia". Corretos: CNPJ 89.472.195/0001/86 e 88.676.127/0002-57; RG 4000780538; CPF 223.537.800/59; nº 230. | Parcial |
| 7 | 16 | p.2 | 90,2 | Acordo Susin: dação de lotes, parcelas, conta bancária | ~20 | **Nenhum valor divergente** (200.000,00; 12 × 5.000,00; 14 × 10.000,00; dia 18; julho/2020; conta 00915-6; ag. 0202; Sicredi 748; 6%; 2%). **Cláusulas "3.1/3.2/3.3/4.1" → "31/32/33/41"**: o ponto sumiu e o número da cláusula vira outro número. | Sim |
| 8 | 16 | p.5 | 91,0 | Acordo AGI/Venzon: qualificação (4 CPFs, RGs, endereços, CNPJ) | ~25 | **Nenhum valor divergente.** "1." perdido; "1.1" → "1.1/"; "344.550.190-49" → "344.550.190- 49" (por causa da hifenização da linha). Várias palavras ficam partidas: "ca- sado", "me- diante". | Sim |
| 9 | 16 | p.6 | 90,8 | Acordo AGI: lista de lotes por comprador (matrículas 16887 e 4021, datas) | ~15 | **Nenhum valor divergente.** "2.1/2.2/2.3" → "21/22/23". Palavras coladas: "Oscinco(5)lotesdoitem". "outorgadas" → "�gadas", por causa da assinatura sobreposta. | Sim |
| 10 | 16 | p.9 | 84,0 | Página de assinaturas com CPFs **manuscritos** das testemunhas | ~6 | "6.1" → "61". Os CPFs manuscritos das testemunhas (cerca de "423.996.780/78" e "618.735.600-44", leitura minha) e os nomes "Luciano Bertocco" e "Karin Mitteregger Susin" foram perdidos. Os nomes impressos "Alexandre Susin", "Alexandre Venzon" e "Bernardete L Susin Venzon" também foram perdidos, porque as assinaturas estão por cima. | Parcial |
| 11 | 7 | p.2 | 90,9 | Arrendamento 2017–2020: área, matrículas, INCRA, sacas/ha | ~22 | **Nenhuma.** Só erros no texto: "I.N.C.R.A." → "L.N.C.R.A."/"\|.N.C.R.A."; "N.I.R.F." → "N.LR.F."; "somente" → "soamente". | Sim |
| 12 | 9 | p.2 | 90,0 | Arrendamento 2023–2026: área, matrículas, sacas | ~24 | **"411,4 ha" → "4114 ha"** (a vírgula sumiu e o número fica 10 vezes maior). Corretos: 342,50 ha; 14.221; 12.473; 11.462; 155,0 ha × 18 sc = 2.790 sc; 30,0 ha × 30 sc = 900 sc; CPFs e RGs. | Sim |
| 13 | 9 | p.3 | 90,9 | Arrendamento 2023–2026: campo nativo, prazo | ~13 | **Nenhuma** (157,5 ha; 3.000,00 kg de boi; 100,0 ha; 68,9 ha; 02/05/2023; 30/04/2026; 90 dias). | Sim |
| 14 | 11 | p.1 | 84,1 | Matrícula 1.412 datilografada (1977–78), com algarismos antigos em que "0" parece "o" e "1" parece "l" | ~40 | "MATRÍCULA 1.412" → "1:412"; "nº 1.106" → "l1.106"; "30m" → "3om"; "01" → "ol" (zona; 01/07/1965); **vencimento "10.01.1982" → "lo.ol.1902"**; **terceiro vencimento "10.01.1988" → "lo.ol.1980"**; "Cr$ 3.920.000,00" → "G$ 3.920;000,00"; "Cr$ 30,00" → "s 30,00"; "Cr$ 920,00" → "U$ 920,00"; "Cr$ 1.140,00" → "Crb 1.140,00"; CGC "270/0001-94" → "270/0001-9%4"; "560/0001-37" → "560/000l=37"; CPF "009 718040-87" → "718040=87"; "HOTEL … LIMITADA" (R.2) → "HOrTCL … LIMIVADA". Corretos: lote 313; 1.800,00 m²; CPF 026523746/72; 28.10.1977; Cr$ 5.099.271,00; 5%; 2%. | Parcial: nomes sim, números não |
| 15 | 13 | p.2 | 80,6 | Matrícula 11.382 (Rádio): R.3, Av.3, R.4, selos digitais e carimbo de certidão | ~45 | **"R$ 80.000,00" → "80.000,D0"** e **"80.000,0O0"**; **"R.4/11.382" → "R.4/11.3882"**; **selo 07987 "(R$ 0,30)" → "(R$ 0,20)"**; selo 07985 "(R$ 0,30)" → "(R$ 0'30)"; "1200001.00805" → "1200D01.00805". No carimbo de certidão, "…05377" saiu truncado e "05378", "25 ABR 2013" e "M.11.382 ficha 01" foram perdidos. Corretos: R$ 77.600,00; 15/03/2013; Livro 230 fls. 194 nº 20.862-133; CNPJ 91 113 399/0001-72; CPF 114 987 900/97; cadastro 1020050468062; protocolos 063911 e 063912; R$ 328,90 e R$ 52,20. | Sim, no corpo do texto |
| 16 | 13 | p.6 | – | Página em branco, só com uma mancha no canto | 0 | OCR vazio, o que está correto. | n/a |
| 17 | 5 | p.1 | 73,2 | NF Case (metade esquerda; digitalizada cortada à direita): texto em matricial e tabela | ~35 | "ART. 52" → "ARI. O2"; "45.490/00" → "45 490/002"; "16/01/98" → "16/01/868"; "6118,00" → "6118,0PD"; "17.5X25-12L" → "17 , 5X25- 12h"; **"RETRO 30"" → "80""**; "580L" (código BIM) → "5806L"; **"2.401 KG" → "2.403 KG"**; **chassi "JHF0048344" → "JHFOQ48344"**; "12.000,00" → "12.000.00"; "4.761,91" → "4,761,91"; total "100…" (cortado) → "160"; peso bruto "6.118,00" → "G2110,.00"; quantidade 1,00 perdida. Corretos: pedido 229/01; parcelas 60000,01 / 24000,00 (14/12/2001) / 16000,00 (17/12/2001); base 100.000,01; NCM 84295190; motor 46096951; ano 2001. | Parcial: palavras sim, números não |
| 18 | 5 | p.2 | 62,0 | NF Case (metade direita), com carimbo de data | ~20 | CFOP 618 perdido; "58058" → "5BO58"; CNPJ "…/0001-4[4]" → "-4H"; **datas 31/10/2001 (emissão e saída) perdidas**; "1,00" → "1x00"; alíquotas 12 e 5 perdidas; **IPI "4.761,91" → "4./161:;83"**; totais 95.238,10 e 100.000,01 do rodapé, peso 6.118,000 e controle 055962 perdidos. Corretos: NF 50.859; CNPJ 60.850.617/0020-90; IE 669.074.054.118; 95.238,10 (linha do item); AIDF 2092. | Não para números |
| 19 | 6 | p.1 | 66,3 | NF Paraná Equipamentos (escavadeira CAT 320D) em matricial, com anotação **manuscrita** e carimbo da SEFAZ-RS | ~50 | **Quase tudo perdido ou errado.** Perdidos: NF 0004946; CNPJ emitente 76.527.951/0012-38; IE; CFOP 6102; NCM 84295219; descrição e série (CATERPILLAR, 320D, CAT0320DAA6F00902; as palavras "ESCAVADEIRA" e "CATERPILLAR" não aparecem); valor unitário e total 470.000,00 na linha do item; peso 19.754,000; condição "R$ 141.000,00 parc. cliente / R$ 329.000,00 parc. FINAME"; PAC 1596/09; carimbo 204371 e 21/11/09. Errados: CEP "13422-180" → "13922100"; "5478" → "3878"; **CNPJ Ceval "04.129.392/0001-44" → "04 199 302/0001-04"**; **"12/11/2009" → "2L 2008"**; telefone "(54) 4411317" → "158) 4941317"; ICMS "56.400,00" → "d6A 090"; **total dos produtos "470.000,00" → "479.000,00"**; Res. "SF 4/98" → "4/96". **Manuscrito perdido: "Maquina do Celito. Ceval – 66,66 Siviero / 33,33 Celito".** Corretos: total da nota (como "470.,000,00"), 20/01/1998, 006481. | **Não** |
| 20 | 4 | p.15 | – | Imagem Google Earth girada 90°, com rótulos "ESTACIONAMENTO SIM", "M. 3.752", "M. 17.408", "M. 17.409" | 4 | **OCR vazio, mas a página não está em branco: o OCR falhou.** Nenhum rótulo foi lido. | Não |
| 21 | 4 | p.16 | 60,3 | Mesma imagem, em outro enquadramento ("ESTACIONAMENTO SIM", "M. 3.752", "M. 17.40…") | 3 | OCR "g_y,i CIONAMENT …". Rótulos e matrículas perdidos. | Não |
| 22 | 4 | p.17 | 45,4 | Imagem com o polígono da M. 3.752 e as medidas 70,00 / 60,00 / 45,00 / 60,00 metros e "470" | 8 | OCR "Earth". Matrícula e medidas perdidas. | Não |

---

## 2. Páginas em que o OCR não serve (ilegíveis para o Tesseract, manuscritas ou só imagem)

Nestas páginas é preciso consultar a imagem. O texto do OCR não basta nem para busca.

1. **6 p.1**: NF da escavadeira CAT 320D. A fonte matricial não foi lida, e a página tem uma **anotação manuscrita relevante para a cisão** ("Maquina do Celito. Ceval 66,66 Siviero 33,33 Celito") que não existe no OCR.
2. **5 p.1 e 5 p.2**: NF da retroescavadeira Case 580L. As palavras-chave aparecem, mas valores, datas, chassi e peso estão errados ou perdidos. A p.2 é a metade direita da mesma nota (a digitalização dividiu a NF em duas imagens).
3. **4 p.15, p.16 e p.17**: imagens de satélite com rótulos de matrícula (M. 3.752, M. 17.408, M. 17.409), "ESTACIONAMENTO SIM" e medidas do terreno. **4 p.15 não está em branco**: o OCR falhou.
4. **11 p.1** (e provavelmente as outras páginas da matrícula 1.412 datilografada, que não conferi): os nomes se leem, mas os algarismos antigos ("o" no lugar de 0, "l" no lugar de 1) corrompem números e datas.
5. **Selos e carimbos de reconhecimento de firma e de certidão**, em todos os arquivos: 10 p.11, 15 p.8 e o rodapé de 13 p.2. Números de selo, datas e emolumentos saem truncados ou errados. Em 10 p.11 o selo de Passo Fundo (Maristela) se perde quase todo.
6. **Blocos de assinatura**: 15 p.8, onde o bloco "OLEOPLAN S/A" desaparece, e 16 p.9, onde se perdem os nomes impressos e os CPFs manuscritos das testemunhas.

**Páginas sem texto reconhecido:**

| Página | Situação |
|---|---|
| 4 p.15 | **Falha do OCR**: imagem Google Earth com rótulos de matrícula |
| 10 p.10 | Em branco de fato: verso com transparência da página anterior e um carimbo parcial do Tabelionato |
| 13 p.6 | Em branco de fato |

---

## 3. Conclusão

**O OCR é confiável para localizar informação?** Sim, nos contratos digitados (arquivos 7–10, 15 e 16, e o corpo das matrículas digitadas, como a 13 p.2). Nessas páginas o OCR acerta praticamente todos os nomes, as palavras-chave e os números. Fora delas, não. Notas fiscais (5, 6), a matrícula datilografada (11), imagens (4 p.15–17), selos, carimbos, blocos de assinatura e anotações manuscritas exigem consulta à imagem. Cuidados para a busca:
- "Morada do Ipê I" sai como "Ipê |" e "Ipê II" como "Ipê Il". Buscar por "Ipê" e conferir na imagem.
- "m²" sai como "m?", e "§" sai como "&".
- Números de cláusula perdem o ponto ("3.1" vira "31").
- Palavras hifenizadas no fim de linha ficam partidas ("ca- sado"), e algumas palavras ficam coladas ("Oscinco(5)lotesdoitem").
- "ARRENDATÁRIO" às vezes sai "ARRENDATAÁRIO" ou "ARRENDATARIO".
- Uma página inteira sem texto (4 p.15) não significa página em branco.

**Nível de erro nos números observado na amostra (cerca de 450 números):**
- **Corpo de contratos digitados modernos** (10 p.1 e p.5, 15 p.2, 16 p.2, p.5 e p.6, 7 p.2, 9 p.2 e p.3, e o corpo de 13 p.2): cerca de 250 números, com **cerca de 9 erros de valor (≈3–4%)**, além da numeração de cláusulas. Metade dos erros é visível ("80.000,D0", "1.,942,03", "1200D01"). A outra metade é **silenciosa**, porque o número errado parece válido: **"411,4 ha" → "4114 ha"** (9 p.2), **"R.4/11.382" → "11.3882"** e **"R$ 0,30" → "R$ 0,20"** (13 p.2), e **"4.792,03" → "4,792,03"** (15 p.2). Em 5 dessas 10 páginas não houve nenhum erro.
- **Selos, carimbos, assinaturas, matrícula datilografada de 1977 e NFs**: cerca de 200 números, com **mais de 50% de erros ou perdas**. Aqui também há erros silenciosos: "7,60" → "7,50", "1982" → "1902", "1988" → "1980", "2.401 kg" → "2.403 kg", "30"" → "80"", "470.000,00" → "479.000,00" e "12/11/2009" → "2L 2008".
- A confiança média do Tesseract por página ajuda a separar os dois grupos: abaixo de ~85 o OCR fica pouco confiável. Mas ela não protege contra erros pontuais: a 9 p.2 (confiança 90,0) tem o erro de 411,4 → 4114.

**Recomendação:** usar o OCR só como índice de busca. **Nenhum número usado em pacote, planilha ou parecer deve vir do OCR sem conferência na imagem da página** (arquivo + página, citando "p.N"). As NFs (5 e 6), a matrícula 1.412 (11) e as imagens do arquivo 4 devem ser transcritas manualmente se forem usadas. Na NF do arquivo 6, registrar a anotação manuscrita de rateio 66,66% Siviero / 33,33% Celito, que hoje só existe na imagem.
