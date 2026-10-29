"""Dump a .docx (paragraphs and tables, in document order) to a plain-text file.

Usage: python -I dump_docx.py <input.docx> <output.txt>
Tables are written row by row, cells separated by ' | ', prefixed by T<table>.R<row>.
"""

import sys

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def main(src: str, dst: str) -> None:
    doc = docx.Document(src)
    out = []
    t_idx = 0
    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                out.append(text)
        else:
            t_idx += 1
            for r_idx, row in enumerate(block.rows, 1):
                cells = []
                prev = None
                for cell in row.cells:
                    # merged cells repeat the same object; keep one copy
                    if prev is not None and cell._tc is prev:
                        continue
                    prev = cell._tc
                    cells.append(" ".join(cell.text.split()))
                out.append(f"T{t_idx}.R{r_idx}: " + " | ".join(cells))
    with open(dst, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out) + "\n")
    print(f"{dst}: {len(out)} lines, {t_idx} tables")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
