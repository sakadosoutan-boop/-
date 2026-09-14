# -*- coding: utf-8 -*-
"""Check that each sheet still fits on one printed side.

Without tables nothing can be clipped -- text that does not fit simply flows
on. The failure mode is therefore a sheet running onto a second page, which
is what this looks for. It also reports how much of the width each sheet
actually uses, so the blocks can be spread out to fill the paper.

    python3 check.py 早発白帝城_解答版.docx content_sohatsu.py
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))


def render(docx):
    import pymupdf
    out = tempfile.mkdtemp(prefix='check')
    subprocess.run([os.path.join(HERE, 'preview.sh'), os.path.abspath(docx), out],
                   check=True, stdout=subprocess.DEVNULL, cwd=HERE)
    return pymupdf.open(os.path.join(
        out, os.path.splitext(os.path.basename(docx))[0] + '.pdf'))


def used_width(page, minsz=7.0):
    xs = []
    for blk in page.get_text('rawdict')['blocks']:
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                if sp.get('size', 0) < minsz:
                    continue
                for ch in sp.get('chars', []):
                    if ch['c'].strip():
                        xs.append(ch['bbox'][0])
                        xs.append(ch['bbox'][2])
    return (min(xs), max(xs)) if xs else (0, 0)


def main():
    docx, content = sys.argv[1], sys.argv[2]
    ns = {}
    exec(compile(open(content, encoding='utf-8').read(), content, 'exec'), ns)
    want = len(ns['SHEETS'])
    doc = render(docx)
    ok = len(doc) == want
    print('%s: %d page(s), expected %d  %s'
          % (os.path.basename(docx), len(doc), want, 'OK' if ok else '<-- はみ出し'))
    for i, page in enumerate(doc, 1):
        lo, hi = used_width(page)
        left = lo - page.rect.x0 - 45.4        # left margin is 16mm = 45.4pt
        print('   page %d: 右端 %.0fpt から左端 %.0fpt まで、左に %.0fmm あまり'
              % (i, hi, lo, max(0.0, left) / 72 * 25.4))
    if not ok:
        print('   → 語注や設問を短くするか、sheet.py の PITCH_ASIDE を小さくしてください')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
