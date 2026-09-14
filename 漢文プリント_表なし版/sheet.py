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

BAND_TOP = 3175                    # 訓読文 band, 56mm
BAND_GAP = 284                     # 5mm
BAND_BOT = TEXT_H - BAND_TOP - BAND_GAP    # 書き下し / 訳 band, 124mm

PITCH_HEAD = 560
PITCH_ASIDE = 340
PITCH_BODY_MIN, PITCH_BODY_MAX = 700, 1800

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


def para(body, pitch, before=0, colbreak=False):
    p = sub(body, 'p')
    pr = sub(p, 'pPr')
    pr.append(el('snapToGrid', val=0))
    sp = el('spacing', before=before, after=0, line=pitch, lineRule='exact')
    pr.append(sp)
    if colbreak:
        r = sub(p, 'r')
        sub(r, 'br', type='column')
    return p


def line(body, markup, pitch, style, answers=True, colbreak=False):
    p = para(body, pitch, colbreak=colbreak)
    X.emit(p, markup, style, show_answers=answers)
    return p


def ruled(body, pitch, size, colbreak=False):
    """an empty writing line, ruled the full depth of the band"""
    p = para(body, pitch, colbreak=colbreak)
    n = max(1, int(BAND_BOT // (size * 20)))
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


def fits(markup, size_pt, height):
    """does this line stay inside its band, or will it wrap?"""
    return X.advance(markup) * size_pt * 20 <= height


def n_lines(markup, size_pt, height=TEXT_H):
    """how many vertical lines a paragraph takes in a column of `height`"""
    per = max(1, int(height // (size_pt * 20) * FILL))
    return max(1, -(-int(X.advance(markup)) // per))


def block_geometry(sheet):
    """where each block sits, in twips from the left edge of the paper

    How many lines the questions and the notes take can only be estimated
    until the text is laid out, so `sheet['lines']` may carry the counts
    measured from a first rendering; then the blocks fill the paper exactly.
    """
    if sheet.get('lines'):
        q_lines, note_lines = sheet['lines']
        aside_lines = q_lines + note_lines
    else:
        q_lines = sum(n_lines(t, sz) for t, sz in sheet['questions'])
        aside_lines = q_lines + sum(n_lines(t, sz) for t, sz in sheet['notes'])
    w_head = len(sheet['head']) * PITCH_HEAD
    avail = TEXT_W - w_head - SLACK
    pitch_body = (avail - aside_lines * PITCH_ASIDE) // max(1, len(sheet['upper']))
    pitch_body = max(PITCH_BODY_MIN, min(PITCH_BODY_MAX, pitch_body))
    w_body = pitch_body * len(sheet['upper'])
    right = PG_W - M_LR
    x_body = right - w_head - w_body
    x_q = x_body - q_lines * PITCH_ASIDE
    x_note = x_q - (aside_lines - q_lines) * PITCH_ASIDE
    return dict(pitch_body=pitch_body, x_body=x_body, w_body=w_body,
                x_q=x_q, x_note=x_note)


def build_sheet(body, sheet, answers, first_of_document, frames=True):
    """one printed side"""
    aside = [(t, sz, True) for t, sz in sheet['questions']]
    aside += [(t, sz, False) for t, sz in sheet['notes']]
    g = dict(block_geometry(sheet))
    g.update(sheet.get('frame_x') or {})     # measured edges, when we have them
    pitch_body = g['pitch_body']

    # ---- 見出し・指示文・記名
    first = None
    for markup, size, bold in sheet['head']:
        p = para(body, PITCH_HEAD)
        X.emit(p, markup, dict(font=X.TEXT, size=size, bold=bold,
                               color=X.INDIGO if bold else X.INK),
               show_answers=answers)
        if first is None:
            first = p
    if frames:
        # 語注 | 設問 | 本文 sit side by side; pad the outer edges only, and
        # leave a hairline gap where two frames meet so they do not overlap
        top, height = M_TOP - FRAME_PAD, TEXT_H + 2 * FRAME_PAD
        cuts = [g['x_note'] - FRAME_PAD, g['x_q'], g['x_body'],
                g['x_body'] + g['w_body'] + FRAME_PAD]
        for i in range(3):
            lo = cuts[i] + (FRAME_GAP if i else 0)
            hi = cuts[i + 1] - (FRAME_GAP if i < 2 else 0)
            X.shape(first, lo, top, hi - lo, height, i + 1)
    end_section(body, cols=1, nextpage=not first_of_document)

    # ---- 本文: 訓読文（上段） / 書き下し・訳（下段）
    for markup in sheet['upper']:
        if not fits(markup, sheet['upper_size'], BAND_TOP):
            print('  ! 上段からはみ出します（折り返します）:', X.plain(markup))
        line(body, markup, pitch_body,
             dict(font=X.BRUSH, size=sheet['upper_size'], color=X.INK), answers)
    for i, markup in enumerate(sheet['lower']):
        if markup is not None and not fits(markup, sheet['lower_size'], BAND_BOT):
            print('  ! 下段からはみ出します（折り返します）:', X.plain(markup)[:20])
        if answers:
            if markup is None:
                para(body, pitch_body, colbreak=(i == 0))
            else:
                line(body, markup, pitch_body,
                     dict(font=X.TEXT, size=sheet['lower_size'], color=X.RED),
                     True, colbreak=(i == 0))
        else:
            # every row gets a rule, including the one beside the poet's name
            ruled(body, pitch_body, sheet['lower_size'], colbreak=(i == 0))
    end_section(body, uneven=True)

    # ---- 設問 → 語注（右から左へ）
    for markup, size, _ in aside:
        line(body, markup, PITCH_ASIDE,
             dict(font=X.TEXT, size=size, color=X.INK), answers)
    return pitch_body


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


def write(sheets, path, answers, frames=True):
    doc = X.document()
    body = sub(doc, 'body')
    for i, sh in enumerate(sheets):
        build_sheet(body, sh, answers, first_of_document=(i == 0),
                    frames=frames)
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
