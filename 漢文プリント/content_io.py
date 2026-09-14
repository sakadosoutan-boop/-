# -*- coding: utf-8 -*-
"""Read and write the editable text of a worksheet slot.

Text is expressed in a small markup so a lesson's content can live in one
plain data file, while all sizing/spacing/fonts stay in the template:

    {漢｜かん}  ruby: base text and its reading, separated by a full-width bar
    ^{二}      kaeriten / small subscript kana set beside the character
    《…》      answer text (printed red on the answer key, blank for students)
    *…*        bold
    #…#        a kuho term (brush font + indigo + bold), e.g. #使役#
    newline    a column break inside the same cell

Run formatting is *not* stored in the content file. When filling a slot the
template's own runs are inspected first and their properties reused, so each
slot keeps exactly the font and size the teacher set for it in Word.
"""
import re
from copy import deepcopy
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

RUBY_SEP = '｜'       # separates base from reading inside {…｜…}
RED = 'EE0000'        # answer red used throughout the teacher's file
INK = '1A1A1A'        # body ink
INDIGO = '1B3A6B'     # headings, passage numbers
BRUSH = 'HG正楷書体-PRO'
TEXT = 'UD デジタル教科書体 N'


def _el(tag):
    return etree.SubElement if False else etree.Element(W + tag)


def sub_el(parent, tag):
    return etree.SubElement(parent, W + tag)


# ------------------------------------------------------------------ extract
def extract(tc):
    """slot cell -> markup string"""
    paras = tc.findall(W + 'p')
    out = []
    for p in paras:
        out.append(_extract_para(p))
    # drop trailing empty columns, they carry no content
    while out and not out[-1].strip():
        out.pop()
    return '\n'.join(out)


def _extract_para(p):
    buf = []
    for r in p.findall(W + 'r'):
        ruby = r.find(W + 'ruby')
        if ruby is not None:
            rt = ''.join(t.text or '' for t in ruby.find(W + 'rt').iter(W + 't'))
            base = ''.join(t.text or '' for t in ruby.find(W + 'rubyBase').iter(W + 't'))
            base_r = ruby.find(W + 'rubyBase').find(W + 'r')
            buf.append(_wrap('{' + base + RUBY_SEP + rt + '}', base_r))
            continue
        txt = ''.join(t.text or '' for t in r.iter(W + 't'))
        if not txt:
            continue
        rPr = r.find(W + 'rPr')
        if rPr is not None and rPr.find(W + 'vertAlign') is not None:
            buf.append('^{' + txt + '}')
            continue
        buf.append(_wrap(txt, r))
    return _merge(''.join(buf))


def _wrap(txt, r):
    """apply the markup for whatever styling this run carries"""
    rPr = r.find(W + 'rPr') if r is not None else None
    if rPr is None:
        return txt
    col = rPr.find(W + 'color')
    col = col.get(W + 'val') if col is not None else None
    bold = rPr.find(W + 'b') is not None
    fonts = rPr.find(W + 'rFonts')
    fam = fonts.get(W + 'eastAsia') if fonts is not None else None
    if col and col.upper() in ('EE0000', 'FF0000'):
        return '《' + txt + '》'
    if fam and fam.startswith('HG') and col == INDIGO:
        return '#' + txt + '#'
    if bold:
        return '*' + txt + '*'
    return txt


def _merge(s):
    """collapse adjacent identical markup spans: 《a》《b》 -> 《ab》"""
    for a, b in (('》《', ''), ('**', ''), ('##', '')):
        prev = None
        while prev != s:
            prev = s
            s = s.replace(a, b)
    return s


# --------------------------------------------------------------- style probe
def probe(tc, fallback=None):
    """Collect the run properties this slot already uses, so refilled text
    keeps the same look. Returns dict of deep-copied <w:rPr> templates."""
    prof = {'base': None, 'rt': None, 'kaeri': None, 'pPr': None, 'rubyPr': None}
    p0 = tc.find(W + 'p')
    if p0 is not None and p0.find(W + 'pPr') is not None:
        prof['pPr'] = deepcopy(p0.find(W + 'pPr'))
    # The base style is the one most of the slot's text is set in -- not
    # simply the first run, which is often a heading term (#...#) styled
    # differently from the body it introduces.
    rpr0 = tc.find('.//' + W + 'rubyPr')
    if rpr0 is not None:
        prof['rubyPr'] = deepcopy(rpr0)
    weights = {}
    for r in tc.iter(W + 'r'):
        rPr = r.find(W + 'rPr')
        if rPr is None:
            continue
        if rPr.find(W + 'vertAlign') is not None:
            if prof['kaeri'] is None:
                prof['kaeri'] = deepcopy(rPr)
            continue
        in_rt = any(a.tag == W + 'rt' for a in _ancestors(r))
        if in_rt:
            if prof['rt'] is None:
                prof['rt'] = deepcopy(rPr)
        elif rPr.find(W + 'rFonts') is not None:
            if _is_squeezed(rPr):
                continue          # 縦中横 / scaled run: never the body style
            n = len(''.join(t.text or '' for t in r.iter(W + 't')))
            bucket = 'answer' if _is_answer_red(rPr) else 'body'
            entry = weights.setdefault((bucket, _signature(rPr)), [0, rPr])
            entry[0] += n
    body = {k: v for k, v in weights.items() if k[0] == 'body'}
    pick = body or weights
    if pick:
        base = deepcopy(_dominant(pick)[1])
        # red and bold are carried by the markup (《…》 / *…*), never by the
        # slot's default, or a mostly-red slot would come back all red and
        # the student edition would have nothing left to blank out.
        if _is_answer_red(base):
            _set_color(base, INK)
        b = base.find(W + 'b')
        if b is not None:
            base.remove(b)
        bcs = base.find(W + 'bCs')
        if bcs is not None:
            base.remove(bcs)
        _unsqueeze(base)
        prof['base'] = base
    if fallback:
        for k in prof:
            if prof[k] is None:
                prof[k] = deepcopy(fallback[k]) if fallback.get(k) is not None else None
    return prof


def _is_squeezed(rPr):
    """True for a run that is horizontally scaled or set 縦中横.

    A kuho note prints "マスターＰ186" with the numerals rotated upright and
    squeezed (w:eastAsianLayout / w:w 42%). Such a run must never be taken as
    the slot's body style, or every character refilled into that slot comes
    out at 42% width.
    """
    if rPr.find(W + 'eastAsianLayout') is not None:
        return True
    wel = rPr.find(W + 'w')
    return wel is not None and (wel.get(W + 'val') or '100') != '100'


def _unsqueeze(rPr):
    for tag in ('w', 'eastAsianLayout', 'fitText'):
        el = rPr.find(W + tag)
        if el is not None:
            rPr.remove(el)


def _dominant(candidates):
    """Pick the slot's body style out of the styles its runs use.

    Not simply the heaviest one: a kuho note sets its 返り点 and 送り仮名 as
    ordinary small runs, and there can be more of those characters than of
    the sentence they annotate. Take the largest size that still carries a
    real share of the text, so refilled text is set in the body size rather
    than in the annotation size.
    """
    total = sum(e[0] for e in candidates.values()) or 1
    by_size = {}
    for (bucket, sig), entry in candidates.items():
        sz = int(sig[1]) if sig[1] else 0
        cur = by_size.get(sz)
        if cur is None or entry[0] > cur[0]:
            by_size[sz] = entry
        else:
            by_size[sz] = [cur[0] + entry[0], cur[1]]
    real = [(sz, e) for sz, e in by_size.items() if e[0] >= total * 0.25]
    if real:
        return max(real, key=lambda kv: kv[0])[1]
    return max(candidates.values(), key=lambda e: e[0])


def _is_answer_red(rPr):
    col = rPr.find(W + 'color')
    return col is not None and (col.get(W + 'val') or '').upper() in ('EE0000', 'FF0000')


def _signature(rPr):
    """what makes two runs 'the same style' for the purpose of picking the
    slot's dominant body style"""
    f = rPr.find(W + 'rFonts')
    sz = rPr.find(W + 'sz')
    col = rPr.find(W + 'color')
    return (f.get(W + 'eastAsia') if f is not None else None,
            sz.get(W + 'val') if sz is not None else None,
            col.get(W + 'val') if col is not None else None,
            rPr.find(W + 'b') is not None)


def _ancestors(el):
    cur = el.getparent()
    while cur is not None:
        yield cur
        cur = cur.getparent()


# --------------------------------------------------------------------- fill
TOKEN = re.compile(r'(\^.|《|》|\*|#|[^\^《》*#]+)')


MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'


def _floating_runs(tc):
    """Runs that carry a floating shape or picture. The decorative scrolls,
    frames and illustrations are anchored to a paragraph inside a slot, so
    they have to be carried across when that slot's text is replaced --
    otherwise refilling the sheet silently drops the artwork."""
    keep = []
    for p in tc.findall(W + 'p'):
        for r in p.findall(W + 'r'):
            if (r.find('.//' + W + 'drawing') is not None
                    or r.find('.//' + W + 'pict') is not None
                    or r.find('.//' + MC + 'AlternateContent') is not None):
                keep.append(deepcopy(r))
    return keep


def fill(tc, markup, prof, show_answers=True):
    """Replace the slot's text with `markup`, styled from `prof`."""
    art = _floating_runs(tc)
    for p in tc.findall(W + 'p'):
        tc.remove(p)
    first = None
    for line in markup.split('\n'):
        p = etree.SubElement(tc, W + 'p')
        if first is None:
            first = p
        if prof.get('pPr') is not None:
            p.append(deepcopy(prof['pPr']))
        _emit_line(p, line, prof, show_answers)
    for i, r in enumerate(art):
        first.insert(1 if prof.get('pPr') is not None else 0, r)
    # a table cell must end with a paragraph; split() always yields >=1 so ok
    return tc


def _emit_line(p, line, prof, show_answers):
    red = bold = term = False
    i = 0
    pending = ''

    def flush():
        nonlocal pending
        if not pending:
            return
        text = pending
        pending = ''
        if red and not show_answers:
            # student edition: leave the writing space blank
            text = '　' * _blank_len(text)
            if not text:
                return
            _add_run(p, text, prof, red=False, bold=False, term=False)
            return
        _add_run(p, text, prof, red=red, bold=bold, term=term)

    while i < len(line):
        ch = line[i]
        if ch == '^' and line[i + 1:i + 2] == '{':
            j = line.find('}', i)
            if j > 0:
                flush()
                _add_kaeri(p, line[i + 2:j], prof)
                i = j + 1
                continue
        if ch == '《':
            flush(); red = True; i += 1; continue
        if ch == '》':
            flush(); red = False; i += 1; continue
        if ch == '*':
            flush(); bold = not bold; i += 1; continue
        if ch == '#':
            flush(); term = not term; i += 1; continue
        if ch == '{':
            j = line.find('}', i)
            if j > 0 and RUBY_SEP in line[i + 1:j]:
                base, rt = line[i + 1:j].split(RUBY_SEP, 1)
                flush()
                _add_ruby(p, base, rt, prof, red=red, bold=bold, term=term,
                          show_answers=show_answers)
                i = j + 1
                continue
        pending += ch
        i += 1
    flush()


def _blank_len(text):
    """how much blank space a hidden answer should leave behind"""
    n = sum(1 for c in text if c.strip())
    return max(1, round(n * 0.9))


def _rpr(prof, kind='base', red=False, bold=False, term=False):
    src = prof.get(kind)
    if src is None:
        src = prof.get('base')
    rPr = deepcopy(src) if src is not None else etree.Element(W + 'rPr')
    if red:
        _set_color(rPr, RED)
    if term:
        _set_color(rPr, INDIGO)
        _set_font(rPr, BRUSH)
    if bold or term:
        if rPr.find(W + 'b') is None:
            b = etree.Element(W + 'b')
            rPr.insert(_ins_at(rPr, ('rFonts',)), b)
    return rPr


def _ins_at(rPr, after_tags):
    """index just past the given tags, keeping CT_RPr child order valid"""
    idx = 0
    for i, child in enumerate(rPr):
        if child.tag.split('}')[-1] in after_tags:
            idx = i + 1
    return idx


def _set_color(rPr, hexval):
    col = rPr.find(W + 'color')
    if col is None:
        col = etree.Element(W + 'color')
        order = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
                 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
                 'noProof', 'snapToGrid', 'vanish', 'webHidden']
        rPr.insert(_ins_at(rPr, order), col)
    col.set(W + 'val', hexval)


def _set_font(rPr, name):
    f = rPr.find(W + 'rFonts')
    if f is None:
        f = etree.Element(W + 'rFonts')
        rPr.insert(0, f)
    for a in ('ascii', 'eastAsia', 'hAnsi', 'cs'):
        f.set(W + a, name)


def _add_run(p, text, prof, red, bold, term):
    r = etree.SubElement(p, W + 'r')
    r.append(_rpr(prof, 'base', red, bold, term))
    t = etree.SubElement(r, W + 't')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text


def _add_kaeri(p, mark, prof):
    r = etree.SubElement(p, W + 'r')
    rPr = prof.get('kaeri')
    if rPr is None:
        rPr = _rpr(prof, 'base')
        va = etree.SubElement(rPr, W + 'vertAlign')
        va.set(W + 'val', 'subscript')
    else:
        rPr = deepcopy(rPr)
    r.append(rPr)
    t = etree.SubElement(r, W + 't')
    t.text = mark


def _add_ruby(p, base, rt, prof, red, bold, term, show_answers):
    if red and not show_answers:
        _add_run(p, '　' * _blank_len(base), prof, False, False, False)
        return
    r = etree.SubElement(p, W + 'r')
    outer = _rpr(prof, 'base')
    for tag in ('rFonts', 'color', 'b'):
        el = outer.find(W + tag)
        if el is not None:
            outer.remove(el)
    r.append(outer)
    ruby = etree.SubElement(r, W + 'ruby')
    if prof.get('rubyPr') is not None:
        # the slot already prints furigana; keep its exact sizing
        pr = deepcopy(prof['rubyPr'])
        ruby.append(pr)
        align = pr.find(W + 'rubyAlign')
        if align is None:
            align = etree.Element(W + 'rubyAlign')
            pr.insert(0, align)
        hps = _hps_of(prof)
    else:
        pr = etree.SubElement(ruby, W + 'rubyPr')
        align = etree.SubElement(pr, W + 'rubyAlign')
        base_hp = _size_of(prof.get('base')) or 20
        hps = max(8, round(base_hp * 0.45))
        for tag, val in (('hps', hps), ('hpsRaise', base_hp), ('hpsBaseText', base_hp)):
            e = etree.SubElement(pr, W + tag)
            e.set(W + 'val', str(int(val)))
        lid = etree.SubElement(pr, W + 'lid')
        lid.set(W + 'val', 'ja-JP')
    align.set(W + 'val', 'distributeSpace' if len(rt) <= len(base) else 'distributeLetter')
    rt_el = etree.SubElement(ruby, W + 'rt')
    rr = etree.SubElement(rt_el, W + 'r')
    # furigana follows the base's answer colouring, so a red kakikudashi
    # line reads as one red block rather than red text with grey readings
    rt_rpr = _rpr(prof, 'rt', red=red)
    _force_size(rt_rpr, hps)
    rr.append(rt_rpr)
    rtt = etree.SubElement(rr, W + 't')
    rtt.text = rt
    base_el = etree.SubElement(ruby, W + 'rubyBase')
    br = etree.SubElement(base_el, W + 'r')
    br.append(_rpr(prof, 'base', red, bold, term))
    bt = etree.SubElement(br, W + 't')
    bt.text = base


def _hps_of(prof):
    pr = prof.get('rubyPr')
    if pr is None:
        return None
    e = pr.find(W + 'hps')
    return int(e.get(W + 'val')) if e is not None else None


# CT_RPr child order up to w:sz -- anything listed here precedes it
_BEFORE_SZ = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
              'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
              'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
              'spacing', 'w', 'kern', 'position']


def _force_size(rPr, val):
    """Pin the run's point size, creating w:sz/w:szCs when the slot's runs
    inherit their size from the style. Ruby that inherits the body size is
    typeset full-height beside the base text instead of as furigana."""
    if val is None:
        return
    for tag in ('sz', 'szCs'):
        e = rPr.find(W + tag)
        if e is None:
            e = etree.Element(W + tag)
            rPr.insert(_ins_at(rPr, _BEFORE_SZ + (['sz'] if tag == 'szCs' else [])), e)
        e.set(W + 'val', str(int(val)))


def _size_of(rPr):
    if rPr is None:
        return None
    sz = rPr.find(W + 'sz')
    return int(sz.get(W + 'val')) if sz is not None else None
