# -*- coding: utf-8 -*-
"""Find text that collides on the rendered page.

Vertical text in a fixed-height table cell does not reflow -- when a column
holds more than it can take, Word/LibreOffice keep drawing and the glyphs
pile up on top of each other. That is the "真っ黒になる / 見切れる" failure,
and it is easy to detect once the file is rendered: look for characters whose
boxes overlap, and for characters drawn outside the printable area.

    python3 overlap.py rendered.pdf
"""
import sys
import pymupdf

MARGIN_MM = 8.0          # anything closer to the paper edge is off the sheet
MM = 72 / 25.4


def spans(page):
    out = []
    for blk in page.get_text('rawdict')['blocks']:
        for line in blk.get('lines', []):
            for sp in line.get('spans', []):
                for ch in sp.get('chars', []):
                    c = ch['c']
                    if c.strip():
                        out.append((pymupdf.Rect(ch['bbox']), c))
    return out


def report(path):
    doc = pymupdf.open(path)
    bad = 0
    for pno, page in enumerate(doc, 1):
        chars = spans(page)
        page_box = page.rect + (MARGIN_MM * MM, MARGIN_MM * MM,
                                -MARGIN_MM * MM, -MARGIN_MM * MM)
        off = [c for r, c in chars if not page_box.contains(r)]
        # bucket by rounded position so the scan stays linear-ish
        grid = {}
        for r, c in chars:
            grid.setdefault((int(r.x0 / 6), int(r.y0 / 6)), []).append((r, c))
        hits = []
        for key, items in grid.items():
            for i in range(len(items)):
                for j in range(i + 1, len(items)):
                    a, b = items[i][0], items[j][0]
                    inter = a & b
                    if inter.is_empty:
                        continue
                    small = min(abs(a.get_area()), abs(b.get_area())) or 1
                    if abs(inter.get_area()) / small > 0.55:
                        hits.append((items[i][1], items[j][1], round(a.x0), round(a.y0)))
        if off or hits:
            bad += 1
            print('page %d: %d char(s) off the sheet, %d collision(s)'
                  % (pno, len(off), len(hits)))
            if off:
                print('   off-sheet: %s' % ''.join(off[:40]))
            for a, b, x, y in hits[:12]:
                print('   overlap %r/%r at x=%d y=%d' % (a, b, x, y))
        else:
            print('page %d: clean' % pno)
    return bad


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print('====', p)
        report(p)
