# -*- coding: utf-8 -*-
"""Build the worksheet as plain paragraphs -- no tables anywhere.

The page is vertical writing (縦書き), so section breaks stack the blocks
right to left and an uneven-column section splits the page into stacked
bands:

    [見出し・記名] [訓読文 / 書き下し(訳)の2段] [設問] [語注]
     ← 1段        ← 2段組み                    ← 1段   ← 1段

Every paragraph is given an *exact* line pitch (`lineRule="exact"`), not
Word's "1 行" auto spacing. Auto spacing is computed from whatever font is
actually installed, so the same file measures differently on this preview
(LibreOffice + substitute fonts) and in a teacher's Word (the real UD デジタ
ル教科書体 N / HG正楷書体-PRO) -- which is what made the old auto-spaced
sheets drift and, once a paragraph lost its explicit override (e.g. a line
added by hand in Word, which falls back to the document default), balloon
open. Exact spacing is a fixed number of twips regardless of font, so a
sheet that fits here fits in Word too, and a hand-added paragraph that
happens to lose its override still falls back to a *tight* default (set in
STYLES below) instead of Word's roomy one.

設問・語注 used to have their 段落後の間隔 (the gap between 行) stretched
out to make the block reach the left margin -- but that stretch is exactly
what read as "行間が不自然に開いている" to a teacher looking at the file in
Word, so 設問・語注 now get the same small, fixed gap as everything else
(`GAP_ASIDE`) and simply leave blank paper below if they are short.
"""
import zipfile
from lxml import etree

import oxml as X
from oxml import W, el, sub, MM

# ---------------------------------------------------------------- paper
PG_W, PG_H = 16838, 11906          # A4 landscape, twips
# a deeper foot margin: a full line of notes used to finish flush against the
# bottom edge, which reads as text spilling off the sheet
M_TOP, M_BOT, M_LR = 850, 1020, 907
TEXT_W = PG_W - 2 * M_LR           # 15024
TEXT_H = PG_H - M_TOP - M_BOT      # 10036

BAND_TOP = 3402                    # 訓読文 band, 60mm (2-band sheet)
BAND_GAP = 284                     # 5mm, between any two stacked bands
BAND_BOT = TEXT_H - BAND_TOP - BAND_GAP    # 書き下し / 訳 band, 112mm

# Every paragraph's line spacing is exact -- a fixed number of twips per
# character, chosen once here rather than measured off whatever font happens
# to be installed. 1.45x the point size matches the density the previous
# (auto-spaced) sheets already had -- this is not new "extra" room, just the
# same density made deterministic; 段落後の間隔 (below) is the visible gap
# between 行, and is what actually reads as "line spacing" to a teacher.
LINE_FACTOR = 1.45


def pitch_of(size_pt):
    """the exact twips-per-character height for a run at this point size"""
    return int(round(size_pt * 20 * LINE_FACTOR))


# 段落後の間隔 -- in vertical writing, horizontal space to the left of a 行.
# Kept tight: this is what "行間" means when a teacher says the sheet looks
# too loose, and it is independent of each 行's own (now-fixed) line pitch.
GAP_HEAD = 140                     # 見出し・記名 (2.5mm)
GAP_UPPER = 170                    # 訓読文 (3mm)
GAP_ASIDE = 170                    # 語注・設問 -- 他と同じ、普通の行間 (3mm)

# A small residual margin, not for font-metric drift (exact spacing already
# removes that) but for `advance()`'s own estimate of how much room a ruby
# reading or a 返り点 needs -- a model, not a measurement.
SAFETY = 1.05

FRAME_PAD = 110                        # twips of air between text and frame
FRAME_GAP = 50                         # twips left clear between two frames

# Each section's first and last line come out a little wider than the exact
# pitch asks for, and there are three sections on a sheet. Hold that much
# back so the sheet still fits on one page.
SLACK = 700                            # twips


def sectpr(cols=1, bands=None, continuous=True, nextpage=False):
    sp = el('sectPr')
    if nextpage:
        sp.append(el('type', val='nextPage'))
    elif continuous:
        sp.append(el('type', val='continuous'))
    sp.append(el('pgSz', w=PG_W, h=PG_H, orient='landscape'))
    sp.append(el('pgMar', top=M_TOP, right=M_LR, bottom=M_BOT, left=M_LR,
                 header=567, footer=567, gutter=0))
    if bands:
        c = el('cols', num=len(bands), space=BAND_GAP, equalWidth=0)
        for w in bands:
            c.append(el('col', w=w, space=BAND_GAP))
        sp.append(c)
    else:
        sp.append(el('cols', num=cols, space=BAND_GAP, equalWidth=1))
    sp.append(el('textDirection', val='tbRl'))
    # no line grid: the grid would snap every line to its own pitch and
    # override the exact spacing each block asks for
    sp.append(el('docGrid', type='default', linePitch=360, charSpace=0))
    return sp


ASIDE_TAIL = 230                   # 4mm kept clear at the foot of a note
BODY_TAIL = 170                    # 3mm, likewise for the poem and the answers


def para(body, size, after=0, before=0, colbreak=False, bottom=False, tail=0,
         box=None):
    """One paragraph at an exact line pitch for `size` (see `pitch_of`).

    `after` is 段落後の間隔 -- in vertical writing that is horizontal space to
    the left of this line. `bottom` sets 下詰め and `tail` keeps that much of
    the column free below the text. `box` draws a ruled box around the whole
    paragraph (see `ruled_line`) -- a *paragraph* border, so it is immune to
    the run-level fragility that made the old per-character grid disappear
    the moment a teacher retyped a word inside it in Word: whatever text
    ends up in this paragraph, the border stays, because it belongs to the
    paragraph mark, not to any one run of text.
    """
    p = sub(body, 'p')
    pr = sub(p, 'pPr')
    if box:
        bd = el('pBdr')
        for side in ('top', 'left', 'bottom', 'right'):
            bd.append(el(side, val='single', sz=4, space=1, color=box))
        pr.append(bd)
    pr.append(el('snapToGrid', val=0))
    pr.append(el('spacing', before=before, after=after,
                 line=pitch_of(size), lineRule='exact'))
    if tail:
        pr.append(el('ind', right=tail))
    if bottom:
        pr.append(el('jc', val='right'))       # 縦書きでは「下詰め」
    if colbreak:
        r = sub(p, 'r')
        sub(r, 'br', type='column')
    return p


def line(body, markup, style, after=0, colbreak=False, answers=True, **kw):
    p = para(body, style.get('size', 10.5), after=after, colbreak=colbreak, **kw)
    X.emit(p, markup, style, show_answers=answers)
    return p


def ruled_line(body, markup, size, answers, band, after=0, colbreak=False):
    """A writing line boxed the full depth of its band -- a ruled box the
    student writes the answer into, the whole way down. Both editions get
    the same box (padded with blank characters to a uniform depth), so a
    blank sheet and its answer key line up exactly, and the border is on
    the paragraph itself so retyping the text later in Word can never lose
    it (see `para`).
    """
    p = para(body, size, after=after, colbreak=colbreak, tail=BODY_TAIL,
             box=X.RULE)
    style = dict(font=X.TEXT, size=size, color=X.INK)
    used = 0
    if markup:
        X.emit(p, markup, dict(style, color=X.RED), show_answers=answers)
        used = int(X.advance(markup) + 0.999)
    room = max(1, int((band - BODY_TAIL) // (size * 20)) - 1)
    n = max(2, room - used)
    X.run(p, '　' * n, **style)
    return p


def end_section(body, **kw):
    p = sub(body, 'p')
    pr = sub(p, 'pPr')
    pr.append(el('spacing', after=0, line=20, lineRule='exact'))
    pr.append(sectpr(**kw))
    return p


# ------------------------------------------------------------ assembly
FRAME = '8FA3C4'                   # the colour the decorative frames are drawn in


def gap_for(size_pt, pitch):
    """段落後の間隔 that puts a line of this size on the given pitch"""
    return max(0, pitch - pitch_of(size_pt))


def fits(markup, size_pt, height):
    """does this line stay inside its band, or will it wrap?

    A run of characters that never wraps advances at its own glyph size
    (measured: exactly 1.0x the point size, unaffected by this paragraph's
    exact line pitch -- that pitch only governs the gap *between* separate
    行, not the density of characters within one unbroken run). SAFETY pads
    only the residual uncertainty in `advance()`'s own model of how much
    room a ruby reading or a 返り点 needs.
    """
    return X.advance(markup) * size_pt * 20 * SAFETY <= height


def fits_raw(markup, size_pt, height):
    """does this line stay inside its band, with no residual padding --
    本文 (訓読文) is left exactly as the teacher set it, so a warning here
    has no fix to offer and should not cry wolf over the small SAFETY margin.
    """
    return X.advance(markup) * size_pt * 20 <= height


def n_lines(markup, size_pt, height):
    """how many wrapped 行 this note/question paragraph will need.

    An explicit \\n (see `oxml.emit`) always starts a fresh 行, so each
    segment between them is measured on its own rather than letting one
    long segment borrow room from a short one next to it.
    """
    per_col = max(1, int(height / (size_pt * 20 * SAFETY)))
    total = 0
    for seg in markup.split('\n'):
        total += max(1, -(-int(X.advance(seg) + 0.999) // per_col))
    return total


def _groups(sheet):
    """the (upper, upper_size, lower, lower_size) 訓読文/書き下し(訳) groups
    this sheet places side by side, right to left. Most sheets have exactly
    one; a sheet that puts 書き下し and 現代語訳 on the same page (instead of
    two separate sheets) supplies `groups` directly instead of a single
    `upper`/`lower` pair -- each group repeats 訓読文 above its own ruled
    band, so the two stay two bands each, never three stacked bands."""
    if 'groups' in sheet:
        return sheet['groups']
    return [(sheet['upper'], sheet['upper_size'], sheet['lower'], sheet['lower_size'])]


def estimate_sheet_width(sheet):
    """A cheap (no rendering) estimate of how much of the page a sheet asks
    for, so build.py can tell before ever calling LibreOffice whether it
    needs to shrink something to keep the sheet to the page(s) it should
    take."""
    head_w = 0
    for markup, size, _ in sheet['head']:
        n = n_lines(markup.lstrip('@'), size, TEXT_H)
        head_w += n * pitch_of(size) + GAP_HEAD
    body_w = 0
    for upper, upper_size, lower, lower_size in _groups(sheet):
        body_w += (pitch_of(upper_size) + GAP_UPPER) * len(upper)
    aside_w = 0
    for markup, size in sheet['questions'] + sheet['notes']:
        n = n_lines(markup, size, TEXT_H - ASIDE_TAIL)
        aside_w += n * pitch_of(size)
    return head_w + body_w + aside_w


NAME_TAIL = 2200                   # column left free below 氏名, twips (39mm)


def _head_and_frame(body, sheet, answers, frames, frame_rects):
    """見出し・記名, plus (optionally) the decorative frames anchored to the
    first heading paragraph. Shared by the 2-band and 3-band builders."""
    first = None
    for markup, size, bold in sheet['head']:
        name = markup.startswith('@')
        p = para(body, size, after=GAP_HEAD,
                 bottom=name, tail=NAME_TAIL if name else 0)
        X.emit(p, markup.lstrip('@'),
               dict(font=X.TEXT, size=size, bold=bold,
                    color=X.INDIGO if bold else X.INK),
               show_answers=answers)
        if first is None:
            first = p
    if frames and frame_rects:
        for i, (x0, y0, x1, y1) in enumerate(frame_rects, 1):
            X.shape(first, x0, y0, x1 - x0, y1 - y0, i)


def _aside(body, sheet, answers):
    """設問 → 語注（右から左へ）"""
    for markup, size in sheet['questions'] + sheet['notes']:
        line(body, markup, dict(font=X.TEXT, size=size, color=X.INK),
             after=GAP_ASIDE, answers=answers, tail=ASIDE_TAIL)


def build_sheet(body, sheet, answers, first_of_document, frames=True,
                frame_rects=None):
    """one printed side: one or more 訓読文/書き下し(訳) groups (2 bands
    each), side by side, right to left, then 設問 and 語注."""
    _head_and_frame(body, sheet, answers, frames, frame_rects)
    end_section(body, cols=1, nextpage=not first_of_document)

    # ---- 本文: 訓読文（上段、手を加えない） / 書き下し・訳（下段、罫線つき）。
    # 組ごとに段組みの区切り（continuous な sectPr）を閉じてから次の組を
    # 書く -- 1 つの区切りに全部まとめて詰め込もうとすると、Word 側が
    # うまく列を割り振れずに紙面からあふれてしまうため、1 組だけの
    # レイアウト（もとから正しく動くもの）をそのまま繰り返す形にしている。
    groups = _groups(sheet)
    for g, (upper, upper_size, lower, lower_size) in enumerate(groups):
        pitch = pitch_of(upper_size) + GAP_UPPER
        for markup in upper:
            if not fits_raw(markup, upper_size, BAND_TOP - BODY_TAIL):
                print('  ! 上段からはみ出します（折り返します）:', X.plain(markup))
            line(body, markup, dict(font=X.BRUSH, size=upper_size,
                                    color=X.INK), after=GAP_UPPER, answers=answers,
                 tail=BODY_TAIL)
        for i, markup in enumerate(lower):
            if markup is not None and not fits(markup, lower_size,
                                               BAND_BOT - BODY_TAIL):
                print('  ! 下段からはみ出します（折り返します）:', X.plain(markup)[:20])
            ruled_line(body, markup, lower_size, answers, BAND_BOT,
                       after=gap_for(lower_size, pitch), colbreak=(i == 0))
        end_section(body, bands=[BAND_TOP, BAND_BOT])

    _aside(body, sheet, answers)


CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

# The document default carries a *tight, exact* fallback (10.5pt-ish pitch),
# so a paragraph a teacher adds by hand later in Word -- which has no local
# w:spacing override -- still lands close to our spacing instead of Word's
# own roomy default (that mismatch, seen when a hand-edited copy came back
# with every newly-typed line spaced wide, was the "行間が不自然に開く" bug).
_DEFAULT_PITCH = pitch_of(10.5)

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="{t}" w:eastAsia="{t}" w:hAnsi="{t}" w:cs="{t}"/>
<w:color w:val="{ink}"/><w:sz w:val="21"/><w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="{pitch}" w:lineRule="exact"/>
<w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:spacing w:after="0" w:line="{pitch}" w:lineRule="exact"/></w:pPr>
<w:qFormat/></w:style>
</w:styles>'''.format(t=X.TEXT, ink=X.INK, pitch=_DEFAULT_PITCH)

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:ulTrailSpace/><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''


def write(sheets, path, answers, frames=True, rects=None):
    doc = X.document()
    body = sub(doc, 'body')
    for i, sh in enumerate(sheets):
        build_sheet(body, sh, answers, first_of_document=(i == 0),
                   frames=frames, frame_rects=(rects or {}).get(i))
        if i < len(sheets) - 1:
            end_section(body, cols=1)
    body.append(sectpr(cols=1))

    xml = etree.tostring(doc, xml_declaration=True, encoding='UTF-8',
                         standalone=True)

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', xml)
    return path
