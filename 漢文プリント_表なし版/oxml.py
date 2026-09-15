# -*- coding: utf-8 -*-
"""Word XML helpers for a plain, table-free vertical-writing worksheet.

Everything the sheet needs is a paragraph. Layout comes from section breaks
(which stack right-to-left in vertical writing) and from an exact line pitch,
so the line spacing is one number per block instead of a property of a table
cell -- and text that does not fit reflows onto the next line instead of
being clipped away.
"""
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}
W = '{%s}' % NS['w']
WP = '{%s}' % NS['wp']
A = '{%s}' % NS['a']
WPS = '{%s}' % NS['wps']
EMU = 635                # EMU per twip
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

RUBY_SEP = '｜'
INK = '1A1A1A'
RED = 'EE0000'
INDIGO = '1B3A6B'
RULE = 'B9C4D6'          # the faint guide line students write along
FRAME = '8FA3C4'         # the decorative frames
PAGE_W = 16838           # A4 landscape, twips -- needed to place a frame

# Word measures a page-relative anchor in ordinary paper coordinates; in a
# vertical-writing section LibreOffice reads the two axes rotated. Files are
# written for Word; set this to 'preview' only to eyeball frame placement in
# a LibreOffice rendering.
SHAPE_AXES = 'word'
BRUSH = 'HG正楷書体-PRO'
TEXT = 'UD デジタル教科書体 N'

MM = 56.6929             # twips per millimetre


def document():
    """root <w:document> carrying every namespace the file uses"""
    return etree.Element(W + 'document', nsmap=NS)


def el(tag, **attrs):
    e = etree.Element(W + tag)
    for k, v in attrs.items():
        e.set(W + k, str(v))
    return e


def sub(parent, tag, **attrs):
    e = el(tag, **attrs)
    parent.append(e)
    return e


# ------------------------------------------------------------------ runs
def rpr(font=TEXT, size=21, color=INK, bold=False, underline=None):
    """size is in points; Word stores half-points"""
    pr = el('rPr')
    f = el('rFonts')
    for a in ('ascii', 'eastAsia', 'hAnsi', 'cs'):
        f.set(W + a, font)
    pr.append(f)
    if bold:
        pr.append(el('b'))
        pr.append(el('bCs'))
    pr.append(el('color', val=color))
    pr.append(el('sz', val=int(round(size * 2))))
    pr.append(el('szCs', val=int(round(size * 2))))
    if underline:
        # CT_RPr is an ordered sequence: w:u comes after w:sz / w:szCs
        pr.append(el('u', val='single', color=underline))
    return pr


def shape(p, x, y, w, h, ident, color=None, weight=12700, radius=4000):
    """A rounded frame floating behind the text.

    `x` and `y` are ordinary paper coordinates, twips from the left and top
    edges, and go into the file as they are -- that is how Word reads a
    page-relative anchor. (LibreOffice reads the two axes rotated in a
    vertical-writing section, so the preview puts these frames elsewhere;
    trust Word, not the preview, for where a frame lands.)
    """
    color = color or FRAME
    r = sub(p, 'r')
    d = sub(r, 'drawing')
    an = etree.SubElement(d, WP + 'anchor', distT='0', distB='0', distL='0',
                          distR='0', simplePos='0', relativeHeight='251658240',
                          behindDoc='1', locked='0', layoutInCell='1',
                          allowOverlap='1')
    etree.SubElement(an, WP + 'simplePos', x='0', y='0')
    ph = etree.SubElement(an, WP + 'positionH', relativeFrom='page')
    hx, vy = (x, y) if SHAPE_AXES == 'word' else (y, PAGE_W - x - w)
    etree.SubElement(ph, WP + 'posOffset').text = str(int(hx * EMU))
    pv = etree.SubElement(an, WP + 'positionV', relativeFrom='page')
    etree.SubElement(pv, WP + 'posOffset').text = str(int(vy * EMU))
    etree.SubElement(an, WP + 'extent', cx=str(int(w * EMU)), cy=str(int(h * EMU)))
    etree.SubElement(an, WP + 'effectExtent', l='0', t='0', r='0', b='0')
    etree.SubElement(an, WP + 'wrapNone')
    etree.SubElement(an, WP + 'docPr', id=str(ident), name='frame%d' % ident)
    etree.SubElement(an, WP + 'cNvGraphicFramePr')
    g = etree.SubElement(an, A + 'graphic')
    gd = etree.SubElement(g, A + 'graphicData', uri=NS['wps'])
    wsp = etree.SubElement(gd, WPS + 'wsp')
    etree.SubElement(wsp, WPS + 'cNvSpPr')
    sp = etree.SubElement(wsp, WPS + 'spPr')
    xf = etree.SubElement(sp, A + 'xfrm')
    etree.SubElement(xf, A + 'off', x='0', y='0')
    etree.SubElement(xf, A + 'ext', cx=str(int(w * EMU)), cy=str(int(h * EMU)))
    pg = etree.SubElement(sp, A + 'prstGeom', prst='roundRect')
    av = etree.SubElement(pg, A + 'avLst')
    etree.SubElement(av, A + 'gd', name='adj', fmla='val %d' % radius)
    etree.SubElement(sp, A + 'noFill')
    ln = etree.SubElement(sp, A + 'ln', w=str(weight))
    sf = etree.SubElement(ln, A + 'solidFill')
    etree.SubElement(sf, A + 'srgbClr', val=color)
    etree.SubElement(wsp, WPS + 'bodyPr')
    return r


def run(p, text, **style):
    r = sub(p, 'r')
    r.append(rpr(**style))
    t = sub(r, 't')
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return r


def kaeriten(p, mark, **style):
    """返り点・小書きの送り仮名: a subscript run beside the character"""
    st = dict(style)
    st['size'] = style.get('size', 21) * 0.62
    r = sub(p, 'r')
    pr = rpr(**st)
    pr.append(el('vertAlign', val='subscript'))
    r.append(pr)
    t = sub(r, 't')
    t.text = mark
    t.set(XML_SPACE, 'preserve')
    return r


def ruby(p, base, reading, **style):
    """furigana; the reading is sized by w:hps, which is what Word obeys"""
    size = style.get('size', 21)
    hps = max(4, int(round(size * 0.45 * 2)))
    r = sub(p, 'r')
    r.append(el('rPr'))
    rb = sub(r, 'ruby')
    pr = sub(rb, 'rubyPr')
    sub(pr, 'rubyAlign',
        val='distributeSpace' if len(reading) <= len(base) else 'distributeLetter')
    sub(pr, 'hps', val=hps)
    sub(pr, 'hpsRaise', val=int(round(size * 2 * 0.85)))
    sub(pr, 'hpsBaseText', val=int(round(size * 2)))
    sub(pr, 'lid', val='ja-JP')
    rt = sub(rb, 'rt')
    rr = sub(rt, 'r')
    st = dict(style)
    st['size'] = hps / 2.0
    rr.append(rpr(**st))
    t = sub(rr, 't')
    t.text = reading
    base_el = sub(rb, 'rubyBase')
    br = sub(base_el, 'r')
    br.append(rpr(**style))
    bt = sub(br, 't')
    bt.text = base
    bt.set(XML_SPACE, 'preserve')
    return r


# --------------------------------------------------------------- markup
def emit(p, markup, style, show_answers=True, blank_rule=RULE):
    """Write the little markup language into paragraph `p`.

        {漢｜かん}  furigana        ^{二}  kaeriten / small okurigana
        《…》      answer          *…*   bold          #…#  heading term
    """
    red = bold = term = small = False
    buf = []
    i = 0

    def flush():
        if not buf:
            return
        text = ''.join(buf)
        del buf[:]
        st = dict(style)
        if small:
            st['size'] = style.get('size', 10.5) * SMALL_RATIO
        if red:
            if not show_answers:
                # leave a ruled space the student writes on
                run(p, '　' * _blank_len(text), **dict(st, underline=blank_rule))
                return
            st['color'] = RED
        if bold:
            st['bold'] = True
        if term:
            st['color'] = INDIGO
            st['bold'] = True
            st['font'] = BRUSH
        run(p, text, **st)

    while i < len(markup):
        ch = markup[i]
        if ch == '^' and markup[i + 1:i + 2] == '{':
            j = markup.find('}', i)
            if j > 0:
                flush()
                kaeriten(p, markup[i + 2:j], **style)
                i = j + 1
                continue
        if ch == '{':
            j = markup.find('}', i)
            if j > 0 and RUBY_SEP in markup[i + 1:j]:
                base, rt = markup[i + 1:j].split(RUBY_SEP, 1)
                flush()
                st = dict(style)
                if red and show_answers:
                    st['color'] = RED
                if red and not show_answers:
                    run(p, '　' * _blank_len(base), **dict(st, underline=blank_rule))
                else:
                    ruby(p, base, rt, **st)
                i = j + 1
                continue
        if ch == '《':
            flush(); red = True; i += 1; continue
        if ch == '》':
            flush(); red = False; i += 1; continue
        if ch == '*':
            flush(); bold = not bold; i += 1; continue
        if ch == '~':
            flush(); small = not small; i += 1; continue
        if ch == '#':
            flush(); term = not term; i += 1; continue
        buf.append(ch)
        i += 1
    flush()
    return p


def _blank_len(text):
    """how much writing space to leave where an answer is hidden"""
    n = len(text.strip())
    return max(2, min(n, 40))


SMALL_RATIO = 0.72       # ~…~ runs, as a fraction of the base size
RUBY_RATIO = 0.45        # furigana size, as a fraction of the base size
KAERI_RATIO = 0.62       # 返り点 size, likewise


def advance(markup):
    """How long the line really is, counted in base characters.

    A character with furigana is widened until the reading fits beside it, so
    {朝｜あしたニ} takes nearly twice the room of a bare 朝. Counting plain
    characters under-measures those lines and they wrap unexpectedly.
    """
    total, i, scale = 0.0, 0, 1.0
    while i < len(markup):
        ch = markup[i]
        if ch == '~':
            scale = SMALL_RATIO if scale == 1.0 else 1.0
            i += 1
            continue
        if ch == '{':
            j = markup.find('}', i)
            if j > 0 and RUBY_SEP in markup[i + 1:j]:
                base, rt = markup[i + 1:j].split(RUBY_SEP, 1)
                total += max(len(base), len(rt) * RUBY_RATIO) * scale
                i = j + 1
                continue
        if ch == '^' and markup[i + 1:i + 2] == '{':
            j = markup.find('}', i)
            if j > 0:
                total += len(markup[i + 2:j]) * KAERI_RATIO * scale
                i = j + 1
                continue
        if ch in '《》*#':
            i += 1
            continue
        total += scale
        i += 1
    return total


def plain(markup):
    """the characters that take up room, ignoring the markup itself"""
    out, i = [], 0
    while i < len(markup):
        ch = markup[i]
        if ch == '{':
            j = markup.find('}', i)
            if j > 0:
                out.append(markup[i + 1:j].split(RUBY_SEP)[0])
                i = j + 1
                continue
        if ch == '^' and markup[i + 1:i + 2] == '{':
            j = markup.find('}', i)
            if j > 0:
                i = j + 1
                continue
        if ch in '《》*#~':
            i += 1
            continue
        out.append(ch)
        i += 1
    return ''.join(out)
