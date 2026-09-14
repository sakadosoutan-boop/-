# -*- coding: utf-8 -*-
"""Check that every slot's text actually reached the printed page.

A vertical text box that is too small does not complain -- Word and
LibreOffice just stop drawing where the box ends, so the tail of a note or
an answer disappears silently. Rendering the sheet and looking for the end
of each slot's text catches exactly that.

    python3 verify.py 早発白帝城_解答版.docx
"""
import os
import re
import subprocess
import sys
import tempfile

import content_io as C

HERE = os.path.dirname(os.path.abspath(__file__))
PREVIEW = os.path.join(HERE, 'preview.sh')
TAIL = 8            # how many trailing characters must be found


def is_subseq(needle, hay, slack=4):
    """`needle` in order inside a short stretch of `hay`.

    Furigana is interleaved with the base text in the extracted page text
    (「白帝城」 comes back as 「はくていじやう白帝城」), so a plain substring
    test would report text that printed perfectly well as missing. Gaps are
    therefore allowed, but the whole match must still sit inside a window of
    `slack` times its length -- otherwise stray characters from elsewhere on
    the sheet would spell out a tail that was never printed.
    """
    span = len(needle) * slack
    for start in range(len(hay)):
        if hay[start] != needle[0]:
            continue
        k, i = 1, start + 1
        stop = start + span
        while i < len(hay) and i < stop and k < len(needle):
            if hay[i] == needle[k]:
                k += 1
            i += 1
        if k == len(needle):
            return True
    return False


def plain(markup, answers=True):
    if not answers:
        markup = re.sub(r'《[^》]*》', '', markup)
    out = []
    i = 0
    while i < len(markup):
        ch = markup[i]
        if ch == '{':
            j = markup.find('}', i)
            if j > 0:
                out.append(markup[i + 1:j].split(C.RUBY_SEP)[0])
                i = j + 1
                continue
        if ch == '^' and markup[i + 1:i + 2] == '{':
            j = markup.find('}', i)
            if j > 0:
                out.append(markup[i + 2:j])
                i = j + 1
                continue
        if ch in '《》*#\n':
            i += 1
            continue
        out.append(ch)
        i += 1
    return re.sub(r'\s+', '', ''.join(out))


def rendered_text(docx):
    import pymupdf
    out = tempfile.mkdtemp(prefix='verify')
    subprocess.run([PREVIEW, os.path.abspath(docx), out],
                   check=True, stdout=subprocess.DEVNULL, cwd=HERE)
    pdf = os.path.join(out, os.path.splitext(os.path.basename(docx))[0] + '.pdf')
    doc = pymupdf.open(pdf)
    return re.sub(r'\s+', '', ''.join(p.get_text() for p in doc))


def main():
    docx, content = sys.argv[1], sys.argv[2]
    answers = '生徒' not in os.path.basename(docx)
    ns = {}
    exec(compile(open(content, encoding='utf-8').read(), content, 'exec'), ns)
    page = rendered_text(docx)
    missing = []
    for key, markup in sorted(ns['SLOTS'].items()):
        txt = plain(markup, answers)
        if len(txt) < 3:
            continue
        tail = txt[-TAIL:]
        if not is_subseq(tail, page):
            missing.append((key, txt[-24:]))
    for key, tail in missing:
        print('CUT OFF  %-22s ...%s' % (key, tail))
    print('%d of %d slots lose their tail' % (len(missing), len(ns['SLOTS'])))


if __name__ == '__main__':
    main()
