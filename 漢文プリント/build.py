# -*- coding: utf-8 -*-
"""Generate the worksheet from template.docx + a content file.

    python3 build.py content_senju_kaishi.py

writes two files next to the content file:
    <名前>_解答版.docx   answers printed in red
    <名前>_生徒版.docx   the same sheet with the answers blanked out

The template carries the whole design -- page setup, the nested tables that
form the columns, every font size, the decorative scroll/frame shapes and
the illustrations. This script only rewrites the text inside the slots, so
nothing about the layout can drift between runs.
"""
import os
import re
import shutil
import sys
import zipfile
from lxml import etree

import tplkit as T
import content_io as C

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'template.docx')
DOC_PART = 'word/document.xml'

# Word substitutes look-alike names while you type; normalise them so the
# file asks for exactly the two fonts the design expects.
FONT_FIXES = {
    'UD デジタル 教科書体 NK': 'UD デジタル教科書体 N',
    'UD デジタル 教科書体 N': 'UD デジタル教科書体 N',
    'UDデジタル教科書体 N': 'UD デジタル教科書体 N',
    'ＭＳ 明朝': 'UD デジタル教科書体 N',
}
# near-blacks that crept in from hand-editing, unified to one ink colour
INK_ALIASES = {'000000', '262626', '2B2B26'}

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
V = '{urn:schemas-microsoft-com:vml}'


def drop_pictures(root):
    """Remove the photographs and illustrations, keeping the drawn shapes.

    The scrolls and rounded frames are part of the sheet's design and belong
    to every lesson; the portraits and the four-panel strip belong to one
    text, so a new lesson must not inherit them.
    """
    W = T.W
    gone = 0
    for r in list(root.iter(W + 'r')):
        if (r.find('.//' + A + 'blip') is not None
                or r.find('.//' + V + 'imagedata') is not None):
            r.getparent().remove(r)
            gone += 1
    return gone


def load_content(path):
    ns = {}
    with open(path, encoding='utf-8') as f:
        exec(compile(f.read(), path, 'exec'), ns)
    if 'SLOTS' not in ns:
        raise SystemExit('%s defines no SLOTS dict' % path)
    return ns


def normalise(root):
    W = T.W
    for rf in root.iter(W + 'rFonts'):
        for attr in ('ascii', 'eastAsia', 'hAnsi', 'cs'):
            v = rf.get(W + attr)
            if v in FONT_FIXES:
                rf.set(W + attr, FONT_FIXES[v])
    for col in root.iter(W + 'color'):
        v = (col.get(W + 'val') or '').upper()
        if v in INK_ALIASES:
            col.set(W + 'val', C.INK)
        elif v == 'FF0000':
            col.set(W + 'val', C.RED)


def build(content_path, show_answers, out_path):
    data = load_content(content_path)
    slots_text = data['SLOTS']

    with zipfile.ZipFile(TEMPLATE) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = etree.fromstring(parts[DOC_PART])
    normalise(root)
    if data.get('ART', 'shapes') == 'shapes':
        drop_pictures(root)

    slots = T.build_slots(root)
    missing = [k for k in slots_text if k not in slots]
    if missing:
        raise SystemExit('unknown slot(s) in content file: %s' % ', '.join(sorted(missing)[:8]))

    fallback = None
    for key, tc in slots.items():
        prof = C.probe(tc, fallback)
        if fallback is None and prof['base'] is not None:
            fallback = prof
        if key in slots_text:
            C.fill(tc, slots_text[key], prof, show_answers=show_answers)

    parts[DOC_PART] = etree.tostring(root, xml_declaration=True,
                                     encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, blob in parts.items():
            z.writestr(name, blob)
    return out_path


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    content_path = sys.argv[1]
    data = load_content(content_path)
    stem = data.get('NAME') or os.path.splitext(os.path.basename(content_path))[0]
    outdir = os.path.dirname(os.path.abspath(content_path))
    for suffix, show in (('解答版', True), ('生徒版', False)):
        out = os.path.join(outdir, '%s_%s.docx' % (stem, suffix))
        build(content_path, show, out)
        print('wrote', out)


if __name__ == '__main__':
    main()
