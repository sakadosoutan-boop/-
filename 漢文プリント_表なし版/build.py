# -*- coding: utf-8 -*-
"""python3 build.py content_sohatsu.py  ->  <名前>_解答版.docx / _生徒版.docx

Two things can only be known once the text is laid out: how far the notes and
questions reach across the paper, and where each block's edges are. So if
LibreOffice is installed the sheet is rendered twice -- once to widen the gaps
between the notes until the blocks fill the paper, once to read off the edges
and put the frames on them. Without LibreOffice the defaults are used; the
sheet prints correctly, it just leaves more white at the left.
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
HERE = os.path.dirname(os.path.abspath(__file__))


def load(path):
    ns = {}
    exec(compile(open(path, encoding='utf-8').read(), path, 'exec'), ns)
    return ns


def render(docx):
    if not shutil.which('soffice'):
        return None
    try:
        import pymupdf
    except ImportError:
        return None
    out = tempfile.mkdtemp(prefix='fit')
    subprocess.run([os.path.join(HERE, 'preview.sh'), os.path.abspath(docx), out],
                   check=True, stdout=subprocess.DEVNULL, cwd=HERE)
    return pymupdf.open(os.path.join(
        out, os.path.splitext(os.path.basename(docx))[0] + '.pdf'))


def _glyphs(page, minsz=7.0):
    out = []
    for blk in page.get_text('rawdict')['blocks']:
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                if sp.get('size', 0) < minsz:
                    continue
                for ch in sp.get('chars', []):
                    if ch['c'].strip():
                        out.append(ch)
    return out


def _columns(glyphs, tol=6.0):
    """glyphs grouped into vertical lines, then ordered down each line"""
    items = sorted(((g['bbox'][0] + g['bbox'][2]) / 2, g['bbox'][1], g['c'])
                   for g in glyphs)
    groups = []
    for item in items:
        if groups and item[0] - groups[-1][-1][0] <= tol:
            groups[-1].append(item)
        else:
            groups.append([item])
    return [sorted(g, key=lambda t: t[1]) for g in groups]


def _find(columns, needle):
    """x of the line that carries this heading"""
    needle = re.sub(r'\s+', '', needle)[:4]
    for col in columns:
        if needle and needle in re.sub(r'\s+', '', ''.join(t[2] for t in col)):
            return max(t[0] for t in col)
    return None


def widen(doc, sheets):
    """how much to open up the notes so the blocks reach the left margin"""
    gaps = {}
    for i, (page, sheet) in enumerate(zip(doc, sheets)):
        glyphs = _glyphs(page)
        if not glyphs:
            return None
        spare = min(g['bbox'][0] for g in glyphs) * PT - S.M_LR
        n = len(sheet['questions']) + len(sheet['notes'])
        gaps[i] = max(S.GAP_ASIDE,
                      min(S.GAP_ASIDE_MAX, int(S.GAP_ASIDE + spare / n)))
    return gaps


def frames(doc, sheets):
    """the rectangle each frame should sit on, in paper twips

    Three frames per sheet: the whole 本文 (訓読文 and 書き下し・訳 together,
    since they share the same columns, one above the other), then 設問, then
    語注 -- in that order, matching how `build_sheet` draws them.
    """
    pad, gap = S.FRAME_PAD, S.FRAME_GAP
    top, bottom = S.M_TOP - pad, S.M_TOP + S.TEXT_H + pad
    rects = {}
    for i, (page, sheet) in enumerate(zip(doc, sheets)):
        glyphs = _glyphs(page)
        cols = _columns(glyphs)
        q_x = _find(cols, X.plain(sheet['questions'][0][0]))
        n_x = _find(cols, X.plain(sheet['notes'][0][0]))
        head_x = _find(cols, X.plain(sheet['head'][-1][0].lstrip('@')))
        if q_x is None or n_x is None or head_x is None:
            return None
        # the heading and the notes both run the full height of the sheet, so
        # 本文 (both bands) is whatever lies between the two, full height
        body = [g for g in glyphs if q_x + 2 < g['bbox'][0] < head_x - 2]
        if not body:
            return None
        left = min(g['bbox'][0] for g in glyphs) * PT
        lo = min(g['bbox'][0] for g in body) * PT
        hi = max(g['bbox'][2] for g in body) * PT
        rects[i] = [
            (lo - pad, top, hi + pad, bottom),                   # 本文全体
            (n_x * PT + gap, top, q_x * PT + pad, bottom),       # 設問
            (left - pad, top, n_x * PT - gap, bottom),           # 語注
        ]
    return rects


STEP = 0.3          # pt shed per shrink step
FLOORS = {'aside': 8.0, 'lower': 8.0, 'head_bold': 10.5}


def fits_budget(sheets):
    return all(S.estimate_sheet_width(sh) <= S.TEXT_W - S.SLACK for sh in sheets)


def lower_fits(sheets):
    """Does every 書き下し／訳 line stay inside its own column?

    This one matters beyond just the page count: if a single line wraps into
    a second column, every line after it in that band shifts over by one --
    the answers stop lining up under the 訓読文 they belong to.
    """
    height = S.BAND_BOT - S.BODY_TAIL
    return all(S.fits(m, sh['lower_size'], height)
               for sh in sheets for m in sh['lower'] if m)


def shrink_lower(sheets):
    """Shed size only from 書き下し／訳, the minimal fix for a line that
    would otherwise wrap and throw off the alignment with 訓読文 above it."""
    changed = False
    height = S.BAND_BOT - S.BODY_TAIL
    for sh in sheets:
        if all(S.fits(m, sh['lower_size'], height) for m in sh['lower'] if m):
            continue
        new = max(FLOORS['lower'], sh['lower_size'] - STEP)
        if new != sh['lower_size']:
            sh['lower_size'] = new
            changed = True
    return changed


def shrink(sheets):
    """Shed a little size everywhere text can wrap into more 行 than the page
    has room for -- notes and questions first (least noticeable), then the
    書き下し／訳, then the bold title. 本文 (訓読文) is never touched here."""
    changed = False
    for sh in sheets:
        new_qs = [(m, max(FLOORS['aside'], sz - STEP)) for m, sz in sh['questions']]
        new_ns = [(m, max(FLOORS['aside'], sz - STEP)) for m, sz in sh['notes']]
        if new_qs != sh['questions'] or new_ns != sh['notes']:
            sh['questions'], sh['notes'] = new_qs, new_ns
            changed = True
            continue
        lower = max(FLOORS['lower'], sh['lower_size'] - STEP)
        if lower != sh['lower_size']:
            sh['lower_size'] = lower
            changed = True
            continue
        new_head = [(m, (max(FLOORS['head_bold'], sz - STEP) if bold else sz), bold)
                    for m, sz, bold in sh['head']]
        if new_head != sh['head']:
            sh['head'] = new_head
            changed = True
    return changed


def main():
    path = sys.argv[1]
    data = load(path)
    sheets, name = data['SHEETS'], data['NAME']
    outdir = os.path.dirname(os.path.abspath(path))
    want_frames = data.get('FRAMES', 'shape') == 'shape'

    # a cheap, render-free pass: first fix any 書き下し／訳 line that would
    # wrap and break the alignment with 訓読文, then shrink until the
    # arithmetic says each sheet fits the page width -- all before ever
    # calling LibreOffice
    budget_tries = 0
    while not lower_fits(sheets) and budget_tries < 10 and shrink_lower(sheets):
        budget_tries += 1
    while not fits_budget(sheets) and budget_tries < 12 and shrink(sheets):
        budget_tries += 1
    if budget_tries:
        print('  紙幅に収めるため文字をわずかに縮めました（%d 段階）' % budget_tries)

    probe = os.path.join(outdir, '%s_解答版.docx' % name)
    S.write(sheets, probe, True, frames=False)
    doc = render(probe)

    # the real-render safety net: LibreOffice's substitute fonts read a touch
    # narrower than Word's, so even a sheet that now fits the estimate can
    # still spill onto another page once actually laid out -- shrink further
    # if that happens
    render_tries = 0
    while (doc is not None and len(doc) != len(sheets)
           and render_tries < 8 and shrink(sheets)):
        render_tries += 1
        S.write(sheets, probe, True, frames=False)
        doc = render(probe)
    if render_tries:
        print('  実際に組んでもはみ出したため、さらに縮めました（%d 段階）' % render_tries)

    gaps = widen(doc, sheets) if doc and len(doc) == len(sheets) else None
    rects = None
    if gaps and want_frames:
        S.write(sheets, probe, True, frames=False, gaps=gaps)
        doc = render(probe)
        if doc and len(doc) == len(sheets):
            rects = frames(doc, sheets)
    if doc and len(doc) != len(sheets):
        print('  ! %d ページになりました（%d 枚のはずです）。語注か設問を短くしてください'
              % (len(doc), len(sheets)))

    for suffix, answers in (('解答版', True), ('生徒版', False)):
        out = os.path.join(outdir, '%s_%s.docx' % (name, suffix))
        S.write(sheets, out, answers, frames=want_frames, gaps=gaps, rects=rects)
        print('wrote', out)
    if want_frames and rects is None:
        print('  （枠は既定の位置です。LibreOffice があれば実測して合わせます）')


if __name__ == '__main__':
    main()
