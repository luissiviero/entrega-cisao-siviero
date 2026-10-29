# Apresentação da cisão Siviero: estado para retomar

Última atualização: 28/10/2026, ao encerrar a primeira sessão de apresentação.
O brief que vale é `00_prompts/3-apresentacao-powerpoint.md`. Este arquivo só registra o estado e as lições.

## O que existe

| Arquivo | O que é |
|---|---|
| `cisao-siviero-esboco-v1.pptx` | Esboço v1, com 16 slides. Já foi entregue ao Luís (duas vezes: conteúdo e, depois, acabamento visual). Não sobrescrever. |
| `fontes-dos-numeros.md` | Origem de cada número da v1: arquivo, aba e célula ou linha. |
| `revisao-v1.md` | Relatório do revisor independente (20 achados), o que foi corrigido e a conferência visual. |
| `_build/build_deck.py` | Gerador (python-pptx). Os números estão digitados no código, com a origem registrada em `fontes-dos-numeros.md`. Uso: `python _build/build_deck.py <saída.pptx>`. |
| `_build/build_deck_v1.py` | Cópia congelada do gerador que produziu a v1. Não editar. |
| `_build/render.ps1` | Exporta os slides em PNG pelo PowerPoint instalado, para a conferência visual. |
| `_build/png/<deck>/sNN.png` | Imagens geradas pelo `render.ps1`. |

## Como gerar e conferir uma versão

1. Gerar o deck:
   ```
   python C:/entrega-cisao-siviero/03_entrega/apresentacao/_build/build_deck.py C:/entrega-cisao-siviero/03_entrega/apresentacao/cisao-siviero-esboco-vN.pptx
   ```
2. Exportar as imagens. Este comando já está liberado em `.claude/settings.local.json`; use-o exatamente assim, com barras normais:
   ```
   powershell.exe -NoProfile -File C:/entrega-cisao-siviero/03_entrega/apresentacao/_build/render.ps1 C:/entrega-cisao-siviero/03_entrega/apresentacao/cisao-siviero-esboco-vN.pptx
   ```
3. Abrir cada `sNN.png` com Read e procurar texto cortado, sobreposição, valor quebrado entre linhas e desalinhamento.

## Decisões pendentes do Luís (nenhuma respondida até agora)

1. **Critério C0.** Pergunta feita: o deck deve expor o C0 ou só os argumentos técnicos? Também ficaram sem resposta o tom (neutro?) e o nível de detalhe (por bloco?). Enquanto ele não responder, vale a opção conservadora: sem C0, tom neutro, detalhe por bloco.
2. **C0 dedutível.** O revisor apontou que os únicos valores menores adotados são justamente os itens da lista de Rafael: Oleoplan e lote 001 do Lago Negro. Na v1, a composição lote a lote do Lago Negro saiu do slide. O Oleoplan aparece como "ponto fraco", porque as parcelas do próprio Rafael somam R$ 13,63 mi.
3. **Contradição no Oleoplan.** `02_trabalho/pendencias-conferencia-anterior.md`, item 3, diz "adotar R$ 13.631.378,60". O log de decisões e a planilha-mestre mantêm R$ 10.631.378,60. O deck segue o log. O arquivo é da sessão SEI; quem corrige é ela, por pedido do Luís.
4. **Loteamentos contratados.** O brief fala em "até R$ 31 milhões", mas os dados dão R$ 34.143.600, porque o Bonetto tem R$ 3.111.000 segundo Rafael (aba Mestre, W91). O deck usa R$ 34 mi.

## Lições desta sessão

- **Ferramentas:** não há Node nem LibreOffice na máquina. Os decks são feitos com python-pptx, e a renderização é pelo PowerPoint via `render.ps1`.
- **Validação:** o `validate.py` da skill de pptx não funciona neste Windows (faltam os schemas) e dá falso erro. Basta reabrir o arquivo com python-pptx.
- **Gravação de arquivos:** não editar `build_deck.py` com heredoc do bash quando houver `\d`, `\b` ou `\u00a0`, porque o escape quebra. Use a ferramenta Edit.
- **Origem dos números:** todo número vem dos arquivos, por script, nunca da memória. O revisor independente achou erro real (Bonetto) e afirmações sem respaldo ("Rafael não contesta").
- **Limites de leitura:** para as demais empresas, o brief limita a fonte a `01_fontes/outros/conferencia-discrepancias.md`, até que existam arquivos em `02_trabalho/empresas/`. Não usar `02_trabalho/levantamento-app.xlsx`.
- **Arquivos fora do alcance (P13):** esta sessão só grava em `03_entrega/apresentacao/`. `.claude/` é protegido por um hook; o Luís faz essas mudanças.

## Para a v2 e as seguintes

- Copiar o gerador para `_build/build_deck_vN.py` ou editar `build_deck.py`, mantendo o `build_deck_v1.py` intacto.
- O texto "v1" e a data "28/10/2026" estão fixos no código: rodapé, capa, título do arquivo e slides 3 e 4. Trocar em todos os lugares.
- Reler `02_trabalho/log-decisoes.md`, `planilha-mestre-sei.xlsx` (abas Resumo e Mestre), `00_PLANO.md` (seção 7) e os arquivos de `02_trabalho/empresas/`.
- Gravar `fontes-dos-numeros-vN.md` e `revisao-vN.md` novos.
- Rodar o revisor independente, fazer a conferência visual e entregar com a lista do que mudou e do que continua provisório.
