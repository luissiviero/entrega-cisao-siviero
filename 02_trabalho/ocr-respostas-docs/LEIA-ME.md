# OCR das respostas ao Dr. Pablo (arquivos 2 a 16)

Gerado em 28/10/2026 por `02_trabalho/_scripts/ocr_respostas_docs.py`, com Tesseract 5.4.0, idioma português, a 300 dpi. As fontes são `01_fontes/respostas-docs-2026-08-13/`, arquivos 2 a 16: 100 páginas. O MD5 de cada arquivo foi conferido com o `00_FONTES.md`.

- `ocr.jsonl`: uma linha por página, no formato do corpus (`arquivo`, `local` = `p.N`, `texto`), com `md5`, `metodo` e `confianca`. O `_scripts/corpus_fontes.py` usa esse texto nas páginas sem camada de texto.
- Um `.md` por arquivo, para leitura.
- Revisão por amostra (22 páginas, cerca de 450 números): `02_trabalho/_diario/revisao-ocr-respostas-docs.md`.

## Como usar

**O texto serve para localizar informação, não para tirar números.** Todo valor, área, data ou matrícula deve ser conferido na imagem da página e citado com arquivo e página.

- **Contratos digitados** (arquivos 2, 3, 4, 7 a 10, 15 e 16, e o corpo da matrícula 13): bons para busca. Cerca de 3 a 4% dos números saem errados, e às vezes o erro parece um número válido. Exemplos: 9 p.2, "411,4 ha" saiu "4114 ha"; 13 p.2, "R.4/11.382" saiu "11.3882"; 15 p.2, "4.792,03 m²" saiu "4,792,03".
- **O OCR não serve** (mais da metade dos números errados ou ausentes):
  - notas fiscais dos arquivos 5 e 6. A NF do arquivo 6 tem uma anotação à mão que o OCR não lê: "Maquina do Celito. Ceval 66,66 Siviero 33,33 Celito";
  - matrícula datilografada do arquivo 11, em que os anos 1982 e 1988 saíram 1902 e 1980;
  - selos, carimbos e assinaturas;
  - arquivo 15 p.8, em que o bloco "OLEOPLAN S/A" sumiu;
  - imagens do arquivo 4, p.15 a 17. A p.15 é uma imagem de satélite com rótulos de matrícula (M. 3.752, 17.408 e 17.409).

  Nesses casos, leia a imagem.
- **Páginas em branco:** 10 p.10 e 13 p.6.
- **Cuidados na busca:** "Ipê I" e "Ipê II" saem como "Ipê |" e "Ipê Il", então busque só "Ipê". Números de cláusula perdem o ponto ("3.1" vira "31"). "m²" sai como "m?".
