Você continua o estudo da SIVIERO EMPREENDIMENTOS IMOBILIÁRIOS LTDA. (SEI) na cisão do grupo Siviero. Eu, Luís, sou sócio de 1/6 e estou saindo. A sessão anterior ficou longa e foi encerrada, mas todo o estado está em arquivos na pasta `C:\entrega-cisao-siviero`. Trabalhe em português.

**Antes de qualquer ação, leia nesta ordem:**
1. `00_PLANO.md`, inteiro. As premissas P1 a P13 estão na seção 1 e no fim da tabela. O estado atual, as lições de método e a divisão em sessões estão na seção 7.
2. `02_trabalho/_diario/diario.md`, inteiro.
3. `02_trabalho/log-decisoes.md` e `02_trabalho/decisoes-sei.json`.
4. `02_trabalho/README.md`.

A memória do projeto já traz as regras de trabalho. Siga todas.

**Como trabalhamos:**
- **Pacotes pequenos.** Eu aprovo cada fase. O que as fontes concordam vira uma regra em bloco, com amostra de 5 linhas para eu conferir. As exceções vêm em pacotes de até 10 itens, do maior valor para o menor, cada um com o valor do grupo, o valor de Rafael, sua recomendação, o motivo e a página da fonte.
- **Revisão independente obrigatória.** Antes de me mostrar qualquer pacote, rode um revisor com contexto novo que receba as fontes e os arquivos produzidos, mas não as suas conclusões. Diga o que ele encontrou e o que você corrigiu.
- **Números só das fontes.** Todo número é refeito das fontes por script, nunca tirado da memória. Toda decisão aprovada vai para `decisoes-sei.json` e para `log-decisoes.md`, e a planilha-mestre é regenerada por `_scripts/build_mestre_sei.py`.
- **Diário.** Registre cada passo em `_diario/diario.md`, só com acréscimos.
- **Decisões aprovadas.** Nunca mude uma decisão aprovada sem me consultar. Se encontrar erro, mostre, proponha a correção e espere.
- **Critérios C0, C1 e C2** conforme o plano e o log.

**Arquivos (regra P13).** Outras sessões trabalham em paralelo: uma por empresa e uma para o PowerPoint. Você escreve só na raiz de `02_trabalho`, em `_diario/`, `pacotes/`, `_scripts/`, `00_PLANO.md` e `00_FONTES.md`. Não escreva em `02_trabalho/empresas/` nem em `03_entrega/`.

**Próximos passos:**
1. **Refazer o levantamento de restrições ambientais da SEI.** O anterior foi cancelado. `levantamento-app.xlsx` é parcial e não revisado, então use só como pista. O escopo inclui:
   - os 7 lotes em "possível APP" e os 4 do Fracionamento Simioni;
   - os 7 lotes de APP do Lago Negro, já decididos em R$ 0;
   - os 26 lotes de terceiros indenizáveis, R$ 16.240.000;
   - as observações de compromissos ambientais;
   - o Bonetto e os processos de 2018.

   Para cada caso, verifique em todas as fontes, inclusive nas respostas ao Dr. Pablo e na pasta Documentos do Drive, se existe cópia de processo, decisão ou laudo. Monte também a lista de sobreposições, ou seja, o mesmo problema ambiental descontado em mais de um lugar. Revisor em seguida.
2. **Pacote 5: uma regra única para todos os casos de "possível APP",** em qualquer loteamento. Casos diferentes só podem ter tratamento diferente se houver documento que justifique. O Simioni entra nessa regra: em 2018 os 4 lotes valiam R$ 1.000.000 com o mesmo texto de restrição, e em 2025 nenhum processo é citado. Decidi que o valor dos lotes de "possível APP" será um percentual do lote cheio, conforme a situação do processo. Inclua os descontos duplos que encontrar, ou registre-os em `parte6-candidatos-duplicidade.md` se não forem decidíveis agora.
3. **Ficha de fechamento da Parte 3 da SEI,** com uma página de decisões, números e pendências, e nova revisão independente de todas as decisões aprovadas.

**Tarefa intercalada: OCR das respostas ao Dr. Pablo (aprovada em 28/10/2026).** Faça no próximo ponto de espera (por exemplo, enquanto eu analiso um pacote) e, no mais tardar, antes da Parte 4.
- **Escopo:** os arquivos 2 a 16 de `01_fontes/respostas-docs-2026-08-13/`, que somam 100 páginas escaneadas e sem camada de texto. Os arquivos 1, 17 e 18 já têm texto e ficam de fora. Confira o MD5 de cada arquivo contra o `00_FONTES.md` antes de começar.
- **Ferramenta:** Tesseract 5.4 em `C:\Program Files\Tesseract-OCR\` (fora do PATH), usado via PyMuPDF (`page.get_textpage_ocr`, 300 dpi, `TESSDATA_PREFIX=C:/Users/User/AppData/Local/tessdata`, onde estão os idiomas `por`, `eng` e `osd`; a pasta de Program Files só tem `eng` e `osd`). Use sempre o `por`: sem ele, os acentos se perdem ("LOCAGAO", "Verandpolis"). Script: `_scripts/ocr_respostas_docs.py`, já iniciado em 28/10/2026. Se não houver OCR, transcreva lendo as páginas como imagem, com um subagente por arquivo.
- **Saída:** um script novo, `_scripts/ocr_respostas_docs.py`, gera a pasta `02_trabalho/ocr-respostas-docs/` com:
  - `ocr.jsonl`, uma linha por página, no mesmo formato do corpus (`arquivo`, `local` = `p.N`, `texto`) e com os campos extras `md5` e `metodo`;
  - um `.md` por arquivo, para leitura.
  Em seguida, ajuste `_scripts/corpus_fontes.py` para usar esse texto nas páginas sem texto, sem tirá-las de `sem_texto.json`.
- **Números:** o texto do OCR serve para localizar informação, não para dar números. Todo valor, área, data ou matrícula usado numa decisão é conferido na imagem da página e citado com arquivo e página.
- **Revisor:** um revisor com contexto novo compara uma amostra de páginas com as imagens, priorizando as que têm números e as manuscritas, e lista as páginas ilegíveis.
- **Fechamento:** registre no diário e me avise, para que a coordenação acrescente o caminho aos prompts das empresas. Não escreva em `02_trabalho/empresas/`.
Uma sessão à parte analisa o arquivo ACERTOS E ACORDOS e o documento de áreas vendidas do Hotel. Os achados dela ficam em `02_trabalho/analises/acertos-e-areas-vendidas/`, no arquivo `achados-para-sei.md`. Use esses achados nas Partes 4, 5 e 6 só depois que eu aprovar, e não escreva nessa pasta.

4. **Parte 4: contratos, indenizações, compromissos e empréstimos, e destino dos bens que saíram desde 2018.** Antes de começar, peça que eu explique a abordagem, como combinado. Enquanto isso, pode adiantar a preparação: tabela dos compromissos do doc 11 cruzada com o arquivo ACERTOS E ACORDOS, os arquivos 16, 17 e 18 e as planilhas de Rafael.
5. **Depois, as Partes 5, 6 e 7,** na ordem do plano.

Comece me dizendo, em até 5 linhas, o estado que você entendeu e o que vai fazer primeiro. Em seguida, execute o passo 1.
