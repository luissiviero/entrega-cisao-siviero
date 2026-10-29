"""Dump a docx (paragraphs and tables, in document order) to text."""
import sys
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('}tbl'):
            yield Table(child, doc)


def main(path, out):
    doc = docx.Document(path)
    lines = []
    tcount = 0
    for b in iter_blocks(doc):
        if isinstance(b, Paragraph):
            t = b.text.strip()
            if t:
                lines.append(t)
        else:
            tcount += 1
            lines.append(f'[TABELA {tcount}]')
            for r in b.rows:
                cells = []
                prev = None
                for c in r.cells:
                    if c._tc is prev:
                        continue
                    prev = c._tc
                    cells.append(c.text.strip().replace('\n', ' / '))
                lines.append(' | '.join(cells))
            lines.append('[/TABELA]')
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(out, len(lines), 'lines', tcount, 'tables')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
