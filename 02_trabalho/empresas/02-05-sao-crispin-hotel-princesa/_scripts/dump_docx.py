"""Dump a .docx body (paragraphs and table rows, in document order) to plain text.

Usage: python dump_docx.py <in.docx> <out.txt>
Table rows are written as `T<table>R<row>| cell | cell | ...` so later
extraction scripts can be checked against the exact text.
"""

import sys

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def main(src: str, dst: str) -> None:
    doc = docx.Document(src)
    lines = []
    t_idx = 0
    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                lines.append(text)
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
                lines.append(f"T{t_idx}R{r_idx}| " + " | ".join(cells))
    with open(dst, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
