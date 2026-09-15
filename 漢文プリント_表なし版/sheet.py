# -*- coding: utf-8 -*-
"""Build the worksheet as plain paragraphs -- no tables anywhere.

The page is vertical writing (縦書き), so section breaks stack the blocks
right to left and a two-column section splits the page into an upper and a
lower band. One sheet is therefore:

    [見出し・記名] [訓読文 / 書き下し(訳)の2段] [設問] [語注]
     ← 1段        ← 2段組み                    ← 1段   ← 1段

Every line has an exact pitch, so a block's width is just
(number of lines) x (pitch) -- which is what lets the decorative frames be
placed behind the text without measuring anything.
"""
import zipfile
from lxml import etree

import oxml as X
from oxml import W, el, sub, MM

# ---------------------------------------------------------------- paper
PG_W, PG_H = 16838, 11906          # A4 landscape, twips
M_TOP, M_BOT, M_LR = 850, 794, 907
TEXT_W = PG_W - 2 * M_LR           # 15024
TEXT_H = PG_H - M_TOP - M_BOT      # 10262

BAND_TOP = 3402                    # 訓読文 band, 60mm
BAND_GAP = 284                     # 5mm
BAND_BOT = TEXT_H - BAND_TOP - BAND_GAP    # 書き下し / 訳 band, 124mm

# Line spacing is left at Word's default "1 行" everywhere, so the paragraph
# dialog reads the way a teacher expects. The space between lines of the poem
# and between notes comes from 段落後の間隔 instead, which in vertical writing
# opens a gap to the left of the paragraph.
GAP_HEAD = 340                     # 段落後の間隔, twips (6mm)
GAP_UPPER = 280                    # 訓読文 (5mm)
GAP_ASIDE = 0                      # 語注・設問（refit が紙幅に合わせて広げる）
GAP_ASIDE_MAX = 620

# Word's "1 行" leaves about this much room per line, measured from a rendered
# sheet. Used to give the 書き下し the same line pitch as the 訓読文 above it,
# so each answer sits under the line it belongs to even though the two are set
# at different sizes.
LINE_FACTOR = 1.45

# How much of a column a paragraph really gets to use. Kinsoku pushes
# characters onto the next line and the frame eats a little at each end, so a
# line holds fewer characters than the bare arithmetic says. 0.80 was measured
# against the rendered sheet; it errs on the safe side, which costs a narrow
# strip of white at the left edge but never pushes text onto a second page.
FILL = 0.80

FRAME_PAD = 110                        # twips of air between text and frame
FRAME_GAP = 50                         # twips left clear between two frames

# Each section's first and last line come out a little wider than the exact
# pitch asks for (measured: about 7pt at each end, with or without a frame),
# and there are three sections on a sheet. Hold that much back so the sheet
# still fits on one page.
SLACK = 700                            # twips


def sectpr(cols=1, uneven=False, continuous=True, nextpage=False):
    sp = el('sectPr')
    if nextpage:
        sp.append(el('type', val='nextPage'))
    elif continuous:
        sp.append(el('type', val='continuous'))
    sp.append(el('pgSz', w=PG_W, h=PG_H, orient='landscape'))
    sp.append(el('pgMar', top=M_TOP, right=M_LR, bottom=M_BOT, left=M_LR,
                 header=567, footer=567, gutter=0))
    if uneven:
        c = el('cols', num=2, space=BAND_GAP, equalWidth=0)
        c.append(el('col', w=BAND_TOP, space=BAND_GAP))
        c.append(el('col', w=BAND_BOT))
        sp.append(c)
    else:
        sp.append(el('cols', num=cols, space=BAND_GAP, equalWidth=1))
    sp.append(el('textDirection', val='tbRl'))
    # no line grid: the grid would snap every line to its own pitch and
    # override the exact spacing each block asks for
    sp.append(el('docGrid', type='default', linePitch=360, charSpace=0))
    return sp


def para(body, after=0, before=0, colbreak=False, bottom=False, tail=0):
    """One paragraph at the default 1-line spacing.

    `after` is 段落後の間隔 -- in vertical writing that is horizontal space to
    the left of this line. `bottom` sets 下詰め and `tail` keeps that much of
    the column free below the text.
    """
    p = sub(body, 'p')
    pr = sub(p, 'pPr')
    pr.append(el('snapToGrid', val=0))
    pr.append(el('spacing', before=before, after=after,
                 line=240, lineRule='auto'))
    if tail:
        pr.append(el('ind', right=tail))
    if bottom:
        pr.append(el('jc', val='right'))       # 縦書きでは「下詰め」
    if colbreak:
        r = sub(p, 'r')
        sub(r, 'br', type='column')
    return p


def line(body, markup, style, after=0, colbreak=False, answers=True, **kw):
    p = para(body, after=after, colbreak=colbreak, **kw)
    X.emit(p, markup, style, show_answers=answers)
    return p


def ruled(body, markup, size, answers, after=0, colbreak=False):
    """A writing line: the answer, then a faint rule filling the rest.

    Both editions get the same rule, so the answer key and the sheet the
    students write on line up exactly.
    """
    p = para(body, after=after, colbreak=colbreak)
    room = max(1, int(BAND_BOT // (size * 20)) - 1)   # 1 spare so it never wraps
    used = 0
    if answers and markup:
        X.emit(p, markup, dict(font=X.TEXT, size=size, color=X.RED))
        used = int(X.advance(markup) + 0.999)
    n = max(2, room - used)
    X.run(p, '　' * n, font=X.TEXT, size=size, color=X.INK, underline=X.RULE)
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
    return max(0, int(pitch - size_pt * 20 * LINE_FACTOR))


def fits(markup, size_pt, height):
    """does this line stay inside its band, or will it wrap?"""
    return X.advance(markup) * size_pt * 20 <= height
NAME_TAIL = 2200                   # column left free below 氏名, twips (39mm)


def build_sheet(body, sheet, answers, first_of_document, frames=True,
                gap_aside=GAP_ASIDE, frame_rects=None):
    """one printed side"""
    # ---- 見出し（詩の題は本文側）・記名・指示文
    first = None
    for i, (markup, size, bold) in enumerate(sheet['head']):
        name = markup.startswith('@')
        p = para(body, after=GAP_HEAD,
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
    end_section(body, cols=1, nextpage=not first_of_document)

    # ---- 本文: 訓読文（上段、手を加えない） / 書き下し・訳（下段、罫線つき）
    pitch = int(sheet['upper_size'] * 20 * LINE_FACTOR) + GAP_UPPER
    for markup in sheet['upper']:
        if not fits(markup, sheet['upper_size'], BAND_TOP):
            print('  ! 上段からはみ出します（折り返します）:', X.plain(markup))
        line(body, markup, dict(font=X.BRUSH, size=sheet['upper_size'],
                                color=X.INK), after=GAP_UPPER, answers=answers)
    for i, markup in enumerate(sheet['lower']):
        if markup is not None and not fits(markup, sheet['lower_size'], BAND_BOT):
            print('  ! 下段からはみ出します（折り返します）:', X.plain(markup)[:20])
        ruled(body, markup, sheet['lower_size'], answers,
              after=gap_for(sheet['lower_size'], pitch), colbreak=(i == 0))
    end_section(body, uneven=True)

    # ---- 設問 → 語注（右から左へ）
    for markup, size in sheet['questions'] + sheet['notes']:
        line(body, markup, dict(font=X.TEXT, size=size, color=X.INK),
             after=gap_aside, answers=answers)


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

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="{t}" w:eastAsia="{t}" w:hAnsi="{t}" w:cs="{t}"/>
<w:color w:val="{ink}"/><w:sz w:val="21"/><w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
<w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:qFormat/></w:style>
</w:styles>'''.format(t=X.TEXT, ink=X.INK)

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''


def write(sheets, path, answers, frames=True, gaps=None, rects=None):
    doc = X.document()
    body = sub(doc, 'body')
    for i, sh in enumerate(sheets):
        build_sheet(body, sh, answers, first_of_document=(i == 0),
                    frames=frames,
                    gap_aside=(gaps or {}).get(i, GAP_ASIDE),
                    frame_rects=(rects or {}).get(i))
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
