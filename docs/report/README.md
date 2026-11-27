# NTCC project report

`LifeTrack_Health_NTCC_Report.pdf` is the version to submit; `LifeTrack_Health_NTCC_Report.docx` is
the editable copy. Both follow the university's NTCC templates (title page, project certificate,
acknowledgement, self-certificate, abstract, contents, lists of tables and figures, eight chapters,
LNCS-style references).

Rebuild both after editing `build.js` (needs node, LibreOffice and Python with PyMuPDF):

    cd docs/report && npm install && ./make_pdf.sh

`make_pdf.sh` builds the Word file, renders it once to find the page of every heading, table and
figure, and prints those numbers into the contents and the two lists.

- `fig/` — diagrams (`.dot` sources rendered with Graphviz: `dot -Gdpi=200 -Tpng`; `dfd1.dot` is
  laid out by hand, so render it with `neato`; `seq.svg` → `seq.png`). `amity-logo.png` is taken
  from the university's title-page template.
- `shots/` — app screenshots used in Chapter 7. The patient-app ones are rendered from the code by
  `app/test/tool/report_screens_test.dart` and `photo_screens_test.dart` (light mode, 2.625), then
  scaled to 600 px wide; the doctor-app ones are browser captures.
- Personal details (name, enrolment number, programme, guide, session) are filled in `build.js`;
  edit them there so a rebuild keeps them.
