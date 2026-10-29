import docx, sys
d = docx.Document(sys.argv[1])
out = open(sys.argv[2], 'w', encoding='utf-8')
n = 0
for el in d.element.body.iterchildren():
    tag = el.tag.split('}')[1]
    if tag == 'p':
        t = ''.join(x.text or '' for x in el.iter() if x.tag.endswith('}t'))
        if t.strip():
            n += 1; out.write(f'P{n:03d}| {t}\n')
    elif tag == 'tbl':
        for tr in el.iter():
            if tr.tag.endswith('}tr'):
                cells = [' '.join((x.text or '') for x in tc.iter() if x.tag.endswith('}t')).strip() for tc in tr if tc.tag.endswith('}tc')]
                n += 1; out.write(f'R{n:03d}| ' + ' || '.join(cells) + '\n')
