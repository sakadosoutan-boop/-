# -*- coding: utf-8 -*-
"""Template kit for the kanbun worksheet.

The worksheet's skeleton (nested tables, column widths, decorative shapes,
illustrations) lives in a template .docx that was finished by hand in Word
and is therefore known to render correctly. This module does not rebuild
that skeleton -- it only finds the leaf table cells that hold editable text
("slots") and rewrites their runs. Everything else in the file is carried
over untouched, so the layout can't regress.

Slot ids look like:
    p1.kanbun.3          kundoku column for passage (3) on sheet 1
    p1.kakikudashi.3     its answer column underneath
    p2.yaku.3            modern-Japanese column on sheet 2
    p1.note.2.def        the 2nd kuho note's definition column
    p1.note.2.supp       its "*" supplement column
    p1.q.1.text          question 1's prompt
    p1.q.1.answer        question 1's write-in box
    p1.title / p1.instruction / p1.name / p1.source ...
"""
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def qn(tag):
    return '{%s}%s' % (NS_W, tag.split(':', 1)[1])


# ---------------------------------------------------------------- geometry
def cells(tbl):
    """the <w:tc> of a single-row layout table, left to right"""
    tr = tbl.find(W + 'tr')
    return tr.findall(W + 'tc')


def rows_of(tbl):
    return tbl.findall(W + 'tr')


def sub(tc, i=0):
    """the i-th table nested inside a cell"""
    return tc.findall(W + 'tbl')[i]


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(W + 't'))


def is_leaf(tc):
    return tc.find(W + 'tbl') is None


def leaf_under(tc):
    """descend through single-cell wrapper tables to the cell that holds text"""
    while not is_leaf(tc):
        inner = tc.find(W + 'tbl')
        cs = []
        for tr in rows_of(inner):
            cs += tr.findall(W + 'tc')
        if len(cs) != 1:
            return tc
        tc = cs[0]
    return tc


# ---------------------------------------------------------------- slot map
def page_tables(doc_root):
    return doc_root.find(W + 'body').findall(W + 'tbl')


def build_slots(doc_root):
    """Walk the known skeleton and return {slot_id: <w:tc>}.

    The walk mirrors how the sheet is composed: each page is one table whose
    single row holds [aside | gap | body | gap | header]; the aside holds the
    note group and the question group; the body holds the kundoku zone above
    and the answer zone below. Cells that only space things out are skipped.
    """
    slots = {}
    for pi, ptbl in enumerate(page_tables(doc_root)):
        P = 'p%d' % (pi + 1)
        top = cells(ptbl)
        aside, body, header = top[0], top[2], top[4]

        # ---- aside: note group + question group
        ac = cells(sub(aside))
        note_group, q_group = ac[0], ac[2]
        _map_notes(slots, P, note_group)
        _map_questions(slots, P, q_group)

        # ---- body: kundoku zone (row 0) and answer zone (last row)
        body_rows = rows_of(sub(body))
        topzone = body_rows[0].findall(W + 'tc')[0]
        botzone = body_rows[-1].findall(W + 'tc')[0]
        _map_zone(slots, P, 'kanbun', topzone, with_source=True)
        _map_zone(slots, P, 'answer', botzone, with_source=False)

        # ---- header: [instruction/name column | gap | title]
        hc = cells(sub(header))
        hrows = rows_of(sub(hc[0]))
        slots[P + '.instruction'] = hrows[0].findall(W + 'tc')[0]
        slots[P + '.name'] = hrows[1].findall(W + 'tc')[0]
        slots[P + '.title'] = hc[2]
    return slots


def _map_notes(slots, P, note_group):
    """Note columns run right-to-left, so the rightmost cell is note 1.

    Each note column is either a leaf (one block of text) or a small table
    of [supplement | gap | definition] -- again right-to-left, so the
    definition is the rightmost cell of the pair.
    """
    cs = cells(sub(note_group))
    idx = 0
    for tc in reversed(cs):
        w = _width_mm(tc)
        if w is not None and w < 2.0:
            continue  # gap column
        if not is_leaf(tc) and _looks_like_heading(tc):
            hr = rows_of(sub(tc))
            slots[P + '.note.heading'] = hr[0].findall(W + 'tc')[0]
            if len(hr) > 1:
                slots[P + '.note.subheading'] = hr[1].findall(W + 'tc')[0]
            continue
        if is_leaf(tc):
            if not text_of(tc).strip():
                continue
            idx += 1
            slots['%s.note.%d.def' % (P, idx)] = tc
            continue
        inner = [c for c in cells(sub(tc)) if (_width_mm(c) or 9) >= 2.0]
        inner = [leaf_under(c) for c in inner]
        inner = [c for c in inner if text_of(c).strip()]
        if not inner:
            continue
        idx += 1
        slots['%s.note.%d.def' % (P, idx)] = inner[-1]
        if len(inner) > 1:
            slots['%s.note.%d.supp' % (P, idx)] = inner[0]


def _map_questions(slots, P, q_group):
    """Question items also run right-to-left; each item is
    [write-in answer | gap | prompt]."""
    cs = cells(sub(q_group))
    idx = 0
    for tc in reversed(cs):
        w = _width_mm(tc)
        if w is not None and w < 2.0:
            continue
        if _looks_like_heading(tc):
            hr = rows_of(sub(tc))
            slots[P + '.q.heading'] = hr[0].findall(W + 'tc')[0]
            if len(hr) > 1:
                slots[P + '.q.subheading'] = hr[1].findall(W + 'tc')[0]
            continue
        inner = [c for c in cells(sub(tc)) if (_width_mm(c) or 9) >= 2.0]
        if len(inner) < 2:
            continue
        idx += 1
        ans, prompt = inner[0], inner[-1]
        ans_leaf = leaf_under(ans)
        if is_leaf(ans_leaf):
            slots['%s.q.%d.answer' % (P, idx)] = ans_leaf
        else:
            # the write-in box was subdivided by hand; expose the pieces
            # right-to-left as .answer.c1, .c2, ... like the passage columns
            for k, c in enumerate(_leaves_rtl(ans_leaf), 1):
                slots['%s.q.%d.answer.c%d' % (P, idx, k)] = c
        prompt_leaf = leaf_under(prompt)
        if is_leaf(prompt_leaf):
            slots['%s.q.%d.text' % (P, idx)] = prompt_leaf
        else:
            # prompt column was split further by hand (extra note beside it)
            pcs = [leaf_under(c) for c in cells(sub(prompt_leaf))
                   if (_width_mm(c) or 9) >= 2.0]
            slots['%s.q.%d.text' % (P, idx)] = pcs[0]
            if len(pcs) > 1:
                slots['%s.q.%d.extra' % (P, idx)] = pcs[-1]


def _leaves_rtl(tc):
    """every text-bearing cell below `tc`, rightmost first"""
    if is_leaf(tc):
        return [tc]
    out = []
    for tbl in tc.findall(W + 'tbl'):
        for tr in rows_of(tbl):
            for c in tr.findall(W + 'tc'):
                if (_width_mm(c) or 9) < 1.0:
                    continue
                out += _leaves_rtl(c)
    return list(reversed(out))


def _map_zone(slots, P, kind, zone, with_source):
    """A zone is one row of passage columns (right-to-left) plus, on the
    kundoku side, the source label at the far right. Each passage column is
    [number row / text row]."""
    cs = cells(sub(zone))
    idx = 0
    source_taken = False
    for tc in reversed(cs):
        w = _width_mm(tc)
        if w is not None and w < 2.0:
            continue
        if not is_leaf(tc):
            rws = rows_of(sub(tc))
            if len(rws) != 2:
                continue
            num_tc = rws[0].findall(W + 'tc')[0]
            txt_tc = leaf_under(rws[1].findall(W + 'tc')[0])
            if (with_source and not source_taken and idx == 0
                    and not text_of(num_tc).strip()):
                # only the outermost column is the source label; a passage
                # column may legitimately carry no number (a poem's title)
                source_taken = True
                slots[P + '.source'] = txt_tc
                continue
            idx += 1
            slots['%s.%s.%d.num' % (P, kind, idx)] = num_tc
            if is_leaf(txt_tc):
                slots['%s.%s.%d' % (P, kind, idx)] = txt_tc
            else:
                # hand-split into several columns: keep them right-to-left
                parts = [leaf_under(c) for c in cells(sub(txt_tc))
                         if (_width_mm(c) or 9) >= 1.0]
                parts = list(reversed(parts))
                for k, p in enumerate(parts, 1):
                    slots['%s.%s.%d.c%d' % (P, kind, idx, k)] = p


def _looks_like_heading(tc):
    """the group heading column is a 2-row stack (title over subtitle)"""
    if is_leaf(tc):
        return False
    inner = tc.find(W + 'tbl')
    return len(rows_of(inner)) == 2 and all(
        len(tr.findall(W + 'tc')) == 1 for tr in rows_of(inner))


def _width_mm(tc):
    el = tc.find(W + 'tcPr/' + W + 'tcW')
    if el is None:
        return None
    try:
        return int(el.get(W + 'w')) / 56.6929
    except (TypeError, ValueError):
        return None
