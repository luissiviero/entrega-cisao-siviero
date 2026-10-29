"""Dump PDF text page by page; report pages without text."""
import sys
import pymupdf


def main(path, out):
    d = pymupdf.open(path)
    empty = []
    with open(out, 'w', encoding='utf-8') as f:
        for i, p in enumerate(d):
            t = p.get_text()
            if len(t.strip()) < 30:
                empty.append(i + 1)
            f.write(f'\n=== PÁGINA {i + 1} ({len(t)} chars)\n' + t)
    print(out, len(d), 'pages; no-text pages:', empty)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
