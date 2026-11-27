#!/bin/sh
# Builds the Word file and the PDF, with real page numbers in the contents,
# the list of tables and the list of figures.
# Needs node, LibreOffice (soffice) and Python with PyMuPDF (pip install pymupdf).
set -e
# LibreOffice on Windows is not on PATH by default.
[ -d "/c/Program Files/LibreOffice/program" ] && PATH="$PATH:/c/Program Files/LibreOffice/program"
cd "$(dirname "$0")"
# The first Python that runs ("python3" on Windows can be the Microsoft Store stub).
PYTHON=$(for p in python3 python; do "$p" -c 'import pymupdf' >/dev/null 2>&1 && { echo "$p"; break; }; done)
[ -n "$PYTHON" ] || { echo "make_pdf.sh: needs Python with PyMuPDF (pip install pymupdf)" >&2; exit 1; }
T=$(mktemp -d)
# Git Bash: give Windows programs (node, python, soffice) a path they understand.
command -v cygpath >/dev/null && T=$(cygpath -m "$T")

# Pass 1: build, render, and find the page of everything the lists print.
FIND_OUT="$T/find.json" OUT="$T/draft.docx" node build.js
soffice --headless --convert-to pdf --outdir "$T" "$T/draft.docx" >/dev/null 2>&1
"$PYTHON" - "$T" <<'PY'
import json, re, sys
import pymupdf
t = sys.argv[1]
wanted = json.load(open(f'{t}/find.json', encoding='utf-8'))
doc = pymupdf.open(f'{t}/draft.pdf')
norm = lambda s: ' '.join(s.split())
roman = re.compile(r'^[ivxlcdm]+$')
pages = []  # (folio, is_front, lines)
for p in doc:
    blocks = [b for b in p.get_text('blocks') if b[4].strip()]
    if not blocks:
        continue
    folio = max(blocks, key=lambda b: b[3])[4].strip()  # the footer is the lowest text
    if not (folio.isdigit() or roman.match(folio)):
        continue  # the title page has no number
    lines = [norm(l) for l in p.get_text('text').splitlines() if l.strip()]
    pages.append((folio, bool(roman.match(folio)), lines))
found, missing = {}, []
for w in wanted:
    key = norm(w['find'])
    for folio, front, lines in pages:
        if front != w['front']:
            continue
        # A heading can wrap: try each line, and each line with the next.
        if any(l.startswith(key) or (l + ' ' + n).startswith(key) for l, n in zip(lines, lines[1:] + [''])):
            found[w['find']] = folio
            break
    else:
        missing.append(w['find'])
if missing:
    sys.exit('make_pdf.sh: not found in the PDF: %s' % missing)
json.dump(found, open(f'{t}/pages.json', 'w', encoding='utf-8'))
print('found the page of', len(found), 'entries')
PY

# Pass 2: the same report with the page numbers printed.
PAGES="$T/pages.json" OUT="$T/LifeTrack_Health_NTCC_Report.docx" node build.js
soffice --headless --convert-to pdf --outdir "$T" "$T/LifeTrack_Health_NTCC_Report.docx" >/dev/null 2>&1
# No folder entries in the Word file: Word has refused a file with them when it was
# opened through a cloud service, and the notice asks for the report on Google Drive.
"$PYTHON" - "$T/LifeTrack_Health_NTCC_Report.docx" <<'PY'
import os, sys, zipfile
src = sys.argv[1]
tmp = src + '.tmp'
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if not item.filename.endswith('/'):
            zout.writestr(item, zin.read(item.filename))
os.replace(tmp, src)
PY
cp "$T/LifeTrack_Health_NTCC_Report.docx" "$T/LifeTrack_Health_NTCC_Report.pdf" .
rm -rf "$T"
echo "wrote LifeTrack_Health_NTCC_Report.docx and LifeTrack_Health_NTCC_Report.pdf"
