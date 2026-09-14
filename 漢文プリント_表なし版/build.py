# -*- coding: utf-8 -*-
"""python3 build.py content_sohatsu.py  ->  <名前>_解答版.docx / _生徒版.docx"""
import os
import sys
import sheet


def main():
    path = sys.argv[1]
    ns = {}
    exec(compile(open(path, encoding='utf-8').read(), path, 'exec'), ns)
    outdir = os.path.dirname(os.path.abspath(path))
    for suffix, answers in (('解答版', True), ('生徒版', False)):
        out = os.path.join(outdir, '%s_%s.docx' % (ns['NAME'], suffix))
        sheet.write(ns['SHEETS'], out, answers)
        print('wrote', out)


if __name__ == '__main__':
    main()
