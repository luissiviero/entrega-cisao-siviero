"""Stage 6 (destination 2018 -> 2025): search every readable source for traces of the assets,
credits and debts of companies 03, 04, 07, 09 and 10 that left between 2018 and 2025.

Reads .docx, text PDFs, .xlsx and .md under the Drive study folder and 01_fontes, plus the
reviewed OCR of the answers to Dr. Pablo (only to locate pages; numbers are checked on the
page image). Writes _extracao/busca-destinos.json and a readable busca-destinos.md.
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

import openpyxl
import pymupdf

sys.path.insert(0, str(Path(__file__).parent))
from fase0_extracao import docx_blocks

BASE = Path(
    r"C:\entrega-cisao-siviero\02_trabalho\empresas\03-04-07-09-10-radio-ceval-baixadas"
)
ROOTS = [
    Path(r"G:\Meu Drive\02_Patrimonio_e_Sociedades\Cisao_Siviero"),
    Path(r"C:\entrega-cisao-siviero\01_fontes"),
]
OCR = Path(r"C:\entrega-cisao-siviero\02_trabalho\ocr-respostas-docs\ocr.jsonl")

TOPICS = {
    "2423K ILD 3959 (Ceval)": r"ILD\s?-?3959|2423\s?K",
    "IUY 8068 Volare (ST)": r"IUY\s?-?8068|VOLARE",
    "IAZ 0281 (ST)": r"IAZ\s?-?0281",
    "ISS 7019 Scania (ST)": r"ISS\s?-?7019",
    "IXO 6000 (ST)": r"IXO\s?-?6000",
    "IMY 5000 (ST)": r"IMY\s?-?5000",
    "IOM 9093 (ST)": r"IOM\s?-?9093",
    "Toyota IHW 0285": r"IHW\s?-?0285|BANDEIRANTE",
    "Millenium / ônibus vendidos": r"MIL+EN+IUM",
    "Celito Pierozan": r"CELITO",
    "Colina do Retiro I (Rádio)": r"COLINA DO RETIRO I\b(?!I)",
    "Operação Deus tá vendo": r"DEUS T[ÁA] VENDO",
    "ICMS restituição (TS/ST)": r"RESTITUI|ICMS.{0,40}(TURISMO|TRANSPORTES)|DR\.? KOCH",
    "Baixa / certidão / extinção": r"BAIXAD[AO]|CERTID[ÃA]O DE BAIXA|EXTIN[ÇC][ÃA]O|DISTRATO",
    "Associação Cultural": r"ASSOCIA[ÇC][ÃA]O CULTURAL",
    "Siviero Transportes": r"SIVIERO TRANSPORTES?\b",
    "Turismo Siviero": r"TURISMO SIVIERO",
    "Ceval": r"CEVAL",
    "Rádio": r"R[ÁA]DIO COMUNIDADE|R[ÁA]DIO PRINCESA|R[ÁA]DIO 96",
    "Telmo de Toni": r"TELMO DE TONI|TELMO DETONI",
    "Parcelamento Receita (TS/ST)": r"PARCELAMENTO.{0,60}(TURISMO|TRANSPORTES)|(TURISMO|TRANSPORTES).{0,60}PARCELAMENTO",
}
# Skip the inventories of these companies themselves (already extracted in Phase 0) only in
# the generic company-name topics; registration keys are searched everywhere.
GENERIC = {
    "Associação Cultural",
    "Siviero Transportes",
    "Turismo Siviero",
    "Ceval",
    "Rádio",
}


def units(path: Path):
    """(location, text) units of a file."""
    suf = path.suffix.lower()
    try:
        if suf == ".docx":
            p = t = 0
            for kind, block in docx_blocks(
                Path("\\\\?\\" + str(path)) if len(str(path)) > 240 else path
            ):
                if kind == "p":
                    p += 1
                    yield f"par. {p}", block
                else:
                    t += 1
                    for n, r in enumerate(block, 1):
                        yield f"tab. {t} linha {n}", " | ".join(r)
        elif suf == ".pdf":
            for i, page in enumerate(pymupdf.open(path), 1):
                txt = page.get_text()
                if txt.strip():
                    yield f"p. {i}", txt
        elif suf == ".xlsx":
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            for ws in wb.worksheets:
                for r in ws.iter_rows(values_only=True):
                    s = " | ".join(str(v) for v in r if v not in (None, ""))
                    if s:
                        yield f"aba {ws.title}", s
        elif suf in (".md", ".txt"):
            for n, line in enumerate(
                path.read_text(encoding="utf-8", errors="replace").splitlines(), 1
            ):
                if line.strip():
                    yield f"linha {n}", line
    except Exception as e:  # noqa: BLE001 - unreadable file: reported, never skipped silently
        yield "ERRO", f"não abriu: {e}"


def main():
    files = []
    for root in ROOTS:
        files += [
            p
            for p in root.rglob("*")
            if p.is_file()
            and p.suffix.lower() in (".docx", ".pdf", ".xlsx", ".md", ".txt")
        ]
    hits, errors, scanned = [], [], 0
    pats = {k: re.compile(v, re.IGNORECASE) for k, v in TOPICS.items()}
    own = re.compile(
        r"(0[3479]|10)\. INVENT[ÁA]RIO|ANTIGO (0[3479]|10)\.|^(14|11|9|8|7 )\.\s?1\.",
        re.IGNORECASE,
    )
    for f in sorted(files):
        scanned += 1
        is_own = bool(own.search(f.name))
        for loc, txt in units(f):
            if loc == "ERRO":
                errors.append({"arquivo": str(f), "erro": txt})
                continue
            flat = re.sub(r"\s+", " ", txt)
            for topic, pat in pats.items():
                if is_own and topic in GENERIC:
                    continue
                for m in pat.finditer(flat):
                    hits.append(
                        {
                            "tema": topic,
                            "arquivo": str(f),
                            "local": loc,
                            "trecho": flat[max(0, m.start() - 220) : m.end() + 260],
                        }
                    )
                    break
    for line in OCR.read_text(encoding="utf-8").splitlines():
        rec = json.loads(line)
        txt = re.sub(r"\s+", " ", rec.get("text") or rec.get("texto") or "")
        src = rec.get("file") or rec.get("arquivo") or "?"
        pg = rec.get("local") or "?"
        for topic, pat in pats.items():
            m = pat.search(txt)
            if m:
                hits.append(
                    {
                        "tema": topic,
                        "arquivo": f"OCR: {src}",
                        "local": pg,
                        "trecho": txt[max(0, m.start() - 220) : m.end() + 260],
                    }
                )
    # De-duplicate identical excerpts (copies of the same file in Drive and 01_fontes).
    seen, uniq = set(), []
    for h in hits:
        k = (h["tema"], h["local"], hashlib.md5(h["trecho"].encode()).hexdigest())
        if k not in seen:
            seen.add(k)
            uniq.append(h)
    out = BASE / "_extracao" / "busca-destinos.json"
    out.write_text(
        json.dumps(
            {"arquivos_lidos": scanned, "erros": errors, "achados": uniq},
            ensure_ascii=False,
            indent=1,
        ),
        encoding="utf-8",
    )
    md = [
        "# Busca de destinos 2018 → 2025 (gerado por _scripts/busca_destinos.py)",
        "",
        f"Arquivos lidos: {scanned}. Erros: {len(errors)}. Achados: {len(uniq)}.",
        "",
    ]
    for topic in TOPICS:
        hs = [h for h in uniq if h["tema"] == topic]
        md += [f"## {topic} ({len(hs)})", ""]
        for h in hs:
            md.append(
                f"- **{Path(h['arquivo']).name if not h['arquivo'].startswith('OCR') else h['arquivo']}**, {h['local']}: …{h['trecho']}…"
            )
        md.append("")
    if errors:
        md += ["## Arquivos que não abriram", ""] + [
            f"- {e['arquivo']}: {e['erro']}" for e in errors
        ]
    (BASE / "_extracao" / "busca-destinos.md").write_text(
        "\n".join(md) + "\n", encoding="utf-8"
    )
    print("files", scanned, "errors", len(errors), "hits", len(uniq))
    for topic in TOPICS:
        print(f"  {topic}: {sum(1 for h in uniq if h['tema'] == topic)}")
    for e in errors:
        print("  ERR", e["arquivo"][-90:], e["erro"][:80])


if __name__ == "__main__":
    main()
