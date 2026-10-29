"""OCR of the scanned replies to Dr. Pablo (files 2 to 16).

Each page is rendered at 300 dpi and read by Tesseract (Portuguese). The text
only helps to find information: every number used in a decision must be
checked on the page image and cited with file and page.

Usage: python ocr_respostas_docs.py
Writes 02_trabalho/ocr-respostas-docs/ocr.jsonl (one line per page, corpus
format plus md5, metodo and confianca) and one .md per file.
Before running, the MD5 of each file is checked against 00_FONTES.md.
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import subprocess
import tempfile
from pathlib import Path

import fitz

BASE = Path(__file__).resolve().parents[2]
SRC = BASE / "01_fontes" / "respostas-docs-2026-08-13"
OUT = BASE / "02_trabalho" / "ocr-respostas-docs"
TESSERACT = Path("C:/Program Files/Tesseract-OCR/tesseract.exe")
TESSDATA = Path("C:/Users/User/AppData/Local/tessdata")
DPI = 300
FILES = range(2, 17)


def md5_10(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()[:10]


def tesseract_version() -> str:
    out = subprocess.run(
        [str(TESSERACT), "--version"], capture_output=True, text=True, check=False
    )
    return (out.stdout or out.stderr).splitlines()[0].strip()


def ocr_image(png: Path, psm: int) -> tuple[str, float | None]:
    """Run Tesseract once; return text and mean word confidence."""
    base = png.with_suffix("")
    env = dict(os.environ, TESSDATA_PREFIX=str(TESSDATA))
    subprocess.run(
        [
            str(TESSERACT),
            str(png),
            str(base),
            "-l",
            "por",
            "--psm",
            str(psm),
            "txt",
            "tsv",
        ],
        capture_output=True,
        env=env,
        check=True,
    )
    text = base.with_suffix(".txt").read_text(encoding="utf-8")
    confs = []
    with base.with_suffix(".tsv").open(encoding="utf-8") as fh:
        for row in csv.DictReader(fh, delimiter="\t", quoting=csv.QUOTE_NONE):
            if row.get("text", "").strip() and float(row["conf"]) >= 0:
                confs.append(float(row["conf"]))
    return text, (sum(confs) / len(confs) if confs else None)


def ocr_page(page, tmp: Path) -> tuple[str, float | None, str]:
    png = tmp / "page.png"
    page.get_pixmap(dpi=DPI).save(str(png))
    try:
        # psm 1 detects orientation (some scans are rotated)
        text, conf = ocr_image(png, 1)
        psm = 1
    except subprocess.CalledProcessError:
        text, conf = ocr_image(png, 3)
        psm = 3
    if conf is None or conf < 40:
        alt_text, alt_conf = ocr_image(png, 3)
        if alt_conf and (conf is None or alt_conf > conf):
            text, conf, psm = alt_text, alt_conf, 3
    return text, conf, f"psm {psm}"


def main() -> None:
    index = (BASE / "00_FONTES.md").read_text(encoding="utf-8")
    files = sorted(
        (p for p in SRC.glob("*.pdf") if int(p.name.split(" ")[0]) in FILES),
        key=lambda p: int(p.name.split(" ")[0]),
    )
    assert len(files) == len(FILES), [p.name for p in files]
    for p in files:
        if md5_10(p) not in index:
            raise SystemExit(f"MD5 de {p.name} não confere com 00_FONTES.md")

    OUT.mkdir(parents=True, exist_ok=True)
    version = tesseract_version()
    n_pages = 0
    low = []
    with (
        (OUT / "ocr.jsonl").open("w", encoding="utf-8") as jl,
        tempfile.TemporaryDirectory() as td,
    ):
        tmp = Path(td)
        for p in files:
            md5 = md5_10(p)
            doc = fitz.open(str(p))
            md = [
                f"# {p.stem}",
                "",
                f"OCR automático ({version}, por, {DPI} dpi). MD5 do PDF: `{md5}`.",
                "Texto só para localizar informação: números conferidos na imagem da página.",
                "",
            ]
            for i, page in enumerate(doc, 1):
                existing = page.get_text().strip()
                if len(existing) >= 30:
                    text, conf, how = existing, None, "camada de texto do PDF"
                else:
                    text, conf, how = ocr_page(page, tmp)
                    how = f"{version}, por, {DPI} dpi, {how}"
                text = re.sub(r"\n{3,}", "\n\n", text).strip()
                jl.write(
                    json.dumps(
                        {
                            "arquivo": str(p),
                            "local": f"p.{i}",
                            "texto": " ".join(text.split()),
                            "md5": md5,
                            "metodo": how,
                            "confianca": round(conf, 1) if conf is not None else None,
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
                n_pages += 1
                if conf is not None and conf < 70:
                    low.append(f"{p.name.split(' ')[0]} p.{i} ({conf:.0f})")
                conf_txt = f", confiança média {conf:.0f}" if conf is not None else ""
                md += [
                    f"## Página {i}",
                    "",
                    f"_{how}{conf_txt}_",
                    "",
                    text or "(sem texto reconhecido)",
                    "",
                ]
            (OUT / f"{p.stem}.md").write_text("\n".join(md), encoding="utf-8")
            print(f"{p.name[:70]}: {len(doc)} p.")
    print(
        f"{len(files)} arquivos, {n_pages} páginas; confiança < 70: {', '.join(low) or 'nenhuma'}"
    )


if __name__ == "__main__":
    main()
