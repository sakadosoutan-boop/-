#!/bin/sh
# Render a .docx to page PNGs so the layout can be looked at before printing.
#
#   ./preview.sh 早発白帝城_解答版.docx [出力フォルダ]
#
# Needs LibreOffice Writer and PyMuPDF. The two Japanese fonts the design
# asks for (UD デジタル教科書体 N / HG正楷書体-PRO) are Windows fonts, so on
# another machine LibreOffice substitutes something close; line breaks can
# therefore differ a little from Word.
set -e
DOC=$(cd "$(dirname "$1")" && pwd)/$(basename "$1")
OUT=${2:-preview}
mkdir -p "$OUT"
HOME_DIR=$(mktemp -d)
soffice --headless --norestore -env:UserInstallation="file://$HOME_DIR" \
        --convert-to pdf --outdir "$OUT" "$DOC" >/dev/null 2>&1
BASE=$(basename "${DOC%.*}")
python3 - "$OUT/$BASE.pdf" "$OUT" <<'PY'
import sys, pymupdf
pdf, out = sys.argv[1], sys.argv[2]
doc = pymupdf.open(pdf)
print('%s -> %d page(s)' % (pdf.rsplit('/', 1)[-1], len(doc)))
for i, page in enumerate(doc, 1):
    page.get_pixmap(dpi=110).save('%s/p%d.png' % (out, i))
PY
