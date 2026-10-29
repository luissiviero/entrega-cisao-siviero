# 02_trabalho — arquivos de análise

Situação em 28/10/2026, ao fim da primeira sessão. O estado do trabalho está na seção 7 de `../00_PLANO.md`.

## SEI (Siviero Empreendimentos Imobiliários)

| Arquivo | O que é |
|---|---|
| `planilha-mestre-sei.xlsx` | Planilha-mestre. Abas: Mestre (itens 2025 + itens 2018 sem par), Lotes, Loteamentos 2018x2025, Passivos, Resumo (totais grupo × adotado, pendências), Legenda. Gerada por `_scripts/build_mestre_sei.py`. Não editar à mão. |
| `decisoes-sei.json` | Decisões aprovadas: regras R1 a R5, critérios C0 a C2, itens, lotes e pares 2018 × 2025. O script aplica na planilha-mestre. |
| `log-decisoes.md` | Registro legível de cada pacote aprovado, das correções e das decisões do Luís. |
| `sei-2025-itens.xlsx` + `sei-2025-notas.md` | Extração do inventário 2025 da SEI (doc 11). |
| `sei-2018-itens.xlsx` + `sei-2018-notas.md` | Extração do inventário 2018 da SEI (doc 18). |
| `diff-inventario-2025-antigo-x-atual.md` | Diferenças entre as versões do inventário 2025, de todas as empresas. |
| `observacoes-planilhas-google.xlsx` | Notas, comentários e anotações do Luís nas planilhas Google, de todas as empresas (base da Parte 7). |
| `registro-data-base.md` | Ajustes por fuga da data-base 31.12.2024 (P2). |
| `perguntas-ao-grupo.md` | Perguntas aprovadas ao grupo. Nada foi enviado. |
| `parte6-candidatos-duplicidade.md` | Suspeitas de dupla contagem e de desconto em duplicidade. |
| `pendencias-conferencia-anterior.md` | Pendências da conferência de 26/10/2026 (estudo à parte). |
| `pacotes/` | Pacotes de revisão apresentados ao Luís. |
| `levantamento-app.xlsx` | **Parcial e não revisado** (levantamento cancelado). Não usar como base. |
| `_diario/diario.md` | Diário de trabalho (uso de Claude e do revisor). `_diario/revisao-*.md` são os relatórios dos revisores independentes. |
| `_scripts/` | Scripts reproduzíveis. |

## Demais empresas

`empresas/` será criada pela sessão de coordenação, com uma pasta por empresa (regra P13 do plano).
