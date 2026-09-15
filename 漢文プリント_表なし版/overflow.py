# -*- coding: utf-8 -*-
"""Find text that has run outside the printable area.

    python3 overflow.py 早発白帝城_解答版.docx
"""
import os
import subprocess
import sys
import tempfile

import sheet as S

PT = 20.0
HERE = os.path.dirname(os.path.abspath(__file__))
EDGE = 20          # twips of tolerance


def main():
    import pymupdf
    out = tempfile.mkdtemp(prefix='ovf')
    for docx in sys.argv[1:]:
        subprocess.run([os.path.join(HERE, 'preview.sh'), os.path.abspath(docx), out],
                       check=True, stdout=subprocess.DEVNULL, cwd=HERE)
        pdf = os.path.join(out, os.path.splitext(os.path.basename(docx))[0] + '.pdf')
        doc = pymupdf.open(pdf)
        print('====', os.path.basename(docx))
        left, right = (S.M_LR - EDGE) / PT, (S.PG_W - S.M_LR + EDGE) / PT
        top, bottom = (S.M_TOP - EDGE) / PT, (S.PG_H - S.M_BOT + EDGE) / PT
        for pno, page in enumerate(doc, 1):
            sides = {'下': [], '左': [], '右': [], '上': []}
            for blk in page.get_text('rawdict')['blocks']:
                for ln in blk.get('lines', []):
                    for sp in ln.get('spans', []):
                        for ch in sp.get('chars', []):
                            if not ch['c'].strip():
                                continue
                            x0, y0, x1, y1 = ch['bbox']
                            if y1 > bottom:
                                sides['下'].append(ch['c'])
                            elif x0 < left:
                                sides['左'].append(ch['c'])
                            elif x1 > right:
                                sides['右'].append(ch['c'])
                            elif y0 < top - 4:
                                sides['上'].append(ch['c'])
            hit = {k: v for k, v in sides.items() if v}
            if hit:
                for side, chars in hit.items():
                    print('  page %d: %s にはみ出し %d 文字  %s'
                          % (pno, side, len(chars), ''.join(chars)[:50]))
            else:
                print('  page %d: 版面に収まっています' % pno)


if __name__ == '__main__':
    main()
