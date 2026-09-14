# -*- coding: utf-8 -*-
"""python3 build.py content_sohatsu.py  ->  <名前>_解答版.docx / _生徒版.docx

The frames are drawn at computed positions. How many lines the notes and the
questions really take can only be estimated without laying the text out, so if
LibreOffice is installed the sheet is rendered once, the blocks are measured,
and the file is written again with the frames sitting exactly on the text.
Without LibreOffice the estimate is used and a frame may stand a few
millimetres clear of its block, which prints perfectly well.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

import oxml as X
import sheet as S

PT = 20.0          # twips per point


def load(path):
    ns = {}
    exec(compile(open(path, encoding='utf-8').read(), path, 'exec'), ns)
    return ns


def _columns(page, tol=6.0, minsz=7.0):
    """x centre of every vertical line on the page, right to left"""
    xs = []
    for blk in page.get_text('rawdict')['blocks']:
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                if sp.get('size', 0) < minsz:
                    continue
                for ch in sp.get('chars', []):
                    if ch['c'].strip():
                        xs.append(((ch['bbox'][0] + ch['bbox'][2]) / 2,
                                   ch['bbox'][1], ch['c']))
    xs.sort()
    groups = []
    for item in xs:
        if groups and item[0] - groups[-1][-1][0] <= tol:
            groups[-1].append(item)
        else:
            groups.append([item])
    # within a line the characters run down the page, so order them by y
    return [sorted(g, key=lambda t: t[1]) for g in groups]


def measure(docx, sheets):
    """where each sheet's blocks actually landed, in twips"""
    if not shutil.which('soffice'):
        return None
    try:
        import pymupdf
    except ImportError:
        return None
    out = tempfile.mkdtemp(prefix='fit')
    here = os.path.dirname(os.path.abspath(__file__))
    subprocess.run([os.path.join(here, 'preview.sh'), os.path.abspath(docx), out],
                   check=True, stdout=subprocess.DEVNULL, cwd=here)
    doc = pymupdf.open(os.path.join(
        out, os.path.splitext(os.path.basename(docx))[0] + '.pdf'))
    if len(doc) != len(sheets):
        print('  ! %d ページになりました（%d 枚のはずです）。語注か設問を短くしてください'
              % (len(doc), len(sheets)))
        return None
    found = []
    for page, sheet in zip(doc, sheets):
        cols = _columns(page)
        if not cols:
            return None
        head = re.sub(r'\s+', '', X.plain(sheet['notes'][0][0]))[:4]
        note_x = None
        for col in cols:
            text = re.sub(r'\s+', '', ''.join(t[2] for t in col))
            if head and head in text:
                note_x = max(t[0] for t in col)
        if note_x is None:
            return None
        g = S.block_geometry(sheet)
        body_left = g['x_body'] / PT
        centres = [max(t[0] for t in col) for col in cols]
        # lines run right to left: the questions sit between the poem and the
        # first line of the notes, everything further left is a note
        pitch = S.PITCH_ASIDE / PT / 2.0
        q = [x for x in centres if note_x + pitch < x < body_left]
        n = [x for x in centres if x <= note_x + pitch]
        found.append((_count(q, S.PITCH_ASIDE / PT), _count(n, S.PITCH_ASIDE / PT)))
    return found


def _count(xs, pitch):
    """how many lines those glyph columns really are"""
    if not xs:
        return 0
    return max(1, int(round((max(xs) - min(xs)) / pitch)) + 1)


def main():
    path = sys.argv[1]
    data = load(path)
    outdir = os.path.dirname(os.path.abspath(path))
    frames = data.get('FRAMES', 'shape') == 'shape'
    fits = None
    for suffix, answers in (('解答版', True), ('生徒版', False)):
        out = os.path.join(outdir, '%s_%s.docx' % (data['NAME'], suffix))
        S.write(data['SHEETS'], out, answers, frames=frames)
        if fits is None:
            fits = measure(out, data['SHEETS'])
            if fits:
                for sheet, lines in zip(data['SHEETS'], fits):
                    sheet['lines'] = lines
        if fits:
            S.write(data['SHEETS'], out, answers, frames=frames)
        print('wrote', out)


if __name__ == '__main__':
    main()
