# -*- coding: utf-8 -*-
"""Check a content file against each slot's measured capacity.

Vertical text in a fixed-size box is not reflowed onto another page -- what
does not fit is simply not printed. ruler.json holds the number of plain
characters each slot can take, measured by rendering a sheet whose every
slot is packed with a marker character and counting what came out.

Ruby costs width, so a line carrying furigana fits fewer columns than a
plain one; RUBY_COST charges for that.

    python3 budget.py content_xxx.py
"""
import json
import os
import sys
import content_io as C

RULER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ruler.json')
RUBY_COST = 0.55      # a ruby-bearing column is ~1.5x as wide as a plain one


def cost(markup):
    """how much of a slot's capacity this markup consumes, in characters"""
    s, i, n, ruby = markup, 0, 0, 0
    while i < len(s):
        ch = s[i]
        if ch == '{':
            j = s.find('}', i)
            if j > 0:
                base = s[i + 1:j].split(C.RUBY_SEP)[0]
                n += len(base)
                ruby += len(base)
                i = j + 1
                continue
        if ch == '^' and s[i + 1:i + 2] == '{':
            j = s.find('}', i)
            if j > 0:
                i = j + 1
                continue
        if ch in '《》*#':
            i += 1
            continue
        n += 1
        i += 1
    return n + round(ruby * RUBY_COST)


def load(path):
    ns = {}
    exec(compile(open(path, encoding='utf-8').read(), path, 'exec'), ns)
    return ns['SLOTS']


def main():
    mine = load(sys.argv[1])
    cap = json.load(open(RULER, encoding='utf-8'))
    rows = []
    for k, v in mine.items():
        room = cap.get(k)
        if not room:
            continue
        used = 0
        for line in v.split('\n'):
            used += -(-cost(line) // 1)      # each newline starts a new column
        rows.append((used / room, k, used, room))
    rows.sort(reverse=True)
    over = [r for r in rows if r[0] > 1.0]
    print('%-22s %5s %5s  %s' % ('slot', 'used', 'room', 'fill'))
    for f, k, a, b in rows[:12]:
        print('%-22s %5d %5d  %4.0f%%%s' % (k, a, b, f * 100,
                                            '  <-- OVER' if f > 1.0 else
                                            ('  <- tight' if f > 0.85 else '')))
    print('\n%d of %d slots exceed the measured capacity' % (len(over), len(rows)))


if __name__ == '__main__':
    main()
