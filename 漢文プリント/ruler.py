# -*- coding: utf-8 -*-
"""Measure, once, how many characters each slot of the template can print.

Every slot is filled with a long run of one distinctive kanji, the sheet is
rendered, and the glyphs actually drawn are counted. Characters that do not
fit are simply not drawn, so the count is the slot's real capacity -- no
guessing at Word's line breaking.

    python3 ruler.py            # writes ruler.json next to the template
"""
import json, os, subprocess, sys, tempfile
import zipfile
from lxml import etree
import tplkit as T

HERE = os.path.dirname(os.path.abspath(__file__))
MARKS = ('一二三四五六七八九十百千万円日月火水木金土年時分秒上下左右前後内外東西南北中'
         '大小高低多少新古長短明暗春夏秋冬花鳥風雲山川海空天地人口手足目耳鼻舌歯心身体'
         '力気水火田米麦茶酒肉魚鳥犬猫馬牛羊虫貝石金銀銅鉄糸布紙木林森竹松梅桜草葉根実'
         '道路橋門戸窓壁床屋家村町市県国州都府郡区寺社宮城塔橋船車輪馬駅港空港')


def build(path_out):
    slots_probe = T.build_slots(etree.fromstring(
        zipfile.ZipFile(os.path.join(HERE, 'template.docx')).read('word/document.xml')))
    keys = sorted(slots_probe)
    marks = {}
    seen = set()
    pool = [c for c in MARKS if not (c in seen or seen.add(c))]
    if len(pool) < len(keys):
        raise SystemExit('need %d distinct marks, have %d' % (len(keys), len(pool)))
    lines = ['# -*- coding: utf-8 -*-', "NAME = '_ruler'", 'SLOTS = {']
    for k, c in zip(keys, pool):
        marks[k] = c
        lines.append("    %r: %r," % (k, c * 400))
    lines.append('}')
    open(path_out, 'w', encoding='utf-8').write('\n'.join(lines))
    return marks


def main():
    content = os.path.join(HERE, '_ruler_content.py')
    marks = build(content)
    subprocess.run([sys.executable, os.path.join(HERE, 'build.py'), content],
                   check=True, stdout=subprocess.DEVNULL)
    docx = os.path.join(HERE, '_ruler_解答版.docx')
    out = tempfile.mkdtemp()
    subprocess.run([os.path.join(HERE, 'preview.sh'), docx, out],
                   check=True, stdout=subprocess.DEVNULL)
    import pymupdf
    pdf = os.path.join(out, '_ruler_解答版.pdf')
    doc = pymupdf.open(pdf)
    counts = {}
    for page in doc:
        for ch in page.get_text():
            if ch.strip():
                counts[ch] = counts.get(ch, 0) + 1
    cap = {k: counts.get(c, 0) for k, c in marks.items()}
    json.dump(cap, open(os.path.join(HERE, 'ruler.json'), 'w'),
              ensure_ascii=False, indent=1, sort_keys=True)
    for k in sorted(cap, key=lambda x: cap[x]):
        print('%-24s %4d' % (k, cap[k]))


if __name__ == '__main__':
    main()
