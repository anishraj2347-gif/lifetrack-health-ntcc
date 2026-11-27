const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, Header, TextRun, HeadingLevel, AlignmentType, ImageRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, PageBreak, TableOfContents, Footer, PageNumber, LevelFormat,
  NumberFormat, TabStopType, VerticalAlign, TableLayoutType,
} = require('docx');

const FIG = path.join(__dirname, 'fig');
const SHOTS = path.join(__dirname, 'shots');
const DEMO = SHOTS;
const FONT = 'Times New Roman';
const W = 9026; // A4 text width at 1" margins (DXA)

// Every number the report repeats lives here, so one edit changes every mention.
// Regenerate before each rebuild: scripts/report_stats.sh prints these values.
const STATS = {
  date: '6 November 2026', patient: '1.48.2', doctor: '1.9.1',
  tests: { app: 495, doctor: 28, server: 57, protocol: 8, integration: 6, total: 594, skipped: 1 },
  code: { app: [109, '22,056'], doctor: [16, '3,326'], server: [9, '1,956'], protocol: [8, '713'], total: [142, '28,051'], approx: '28,000', tests: [87, '11,273'] },
};

// ---------- helpers ----------
// Text with **bold**, *italic*, `code` and [[placeholder]] (highlighted for the author to fill in).
function runs(text, opts = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*\s][^*]*\*|`[^`]+`|\[\[[^\]]+\]\])/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...opts }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ text: t.slice(2, -2), bold: true, ...opts }));
    else if (t.startsWith('*')) out.push(new TextRun({ text: t.slice(1, -1), italics: true, ...opts }));
    else if (t.startsWith('`')) out.push(new TextRun({ text: t.slice(1, -1), font: 'Courier New', size: 21, ...opts }));
    else out.push(new TextRun({ text: '[' + t.slice(2, -2) + ']', highlight: 'yellow', ...opts }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...opts }));
  return out;
}
const P = (text, o = {}) => new Paragraph({ children: runs(text, o.run || {}), alignment: o.align ?? AlignmentType.JUSTIFIED, spacing: { after: 160, line: 360 }, ...o.para });
const C = (text, o = {}) => P(text, { ...o, align: AlignmentType.CENTER });
const headings = []; // [level, text] in order, for the printed contents page
const H1 = (text) => { headings.push([1, text]); return new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(text)] }); };
const H2 = (text) => { headings.push([2, text]); return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); };
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
const B = (text, level = 0) => new Paragraph({ numbering: { reference: 'bul', level }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 340 } });
const N = (text, ref = 'num') => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 340 } });
const blank = (n = 1) => Array.from({ length: n }, () => new Paragraph({ children: [] }));
const brk = () => new Paragraph({ children: [new PageBreak()] });

let figNo = 0, tabNo = 0;
const figures = [], tables = [];
function imgSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20), data: b };
}
function figure(file, caption, maxW = 600, maxH = 640) {
  const { w, h, data } = imgSize(file);
  const s = Math.min(maxW / w, maxH / h);
  figNo++; figures.push(`Figure ${chap}.${figNo}  ${caption}`);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120 }, keepNext: true, children: [new ImageRun({ type: 'png', data, transformation: { width: Math.round(w * s), height: Math.round(h * s) } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 240 }, children: [new TextRun({ text: `Figure ${chap}.${figNo}: ${caption}`, italics: true, size: 20 })] }),
  ];
}
// Several phone screenshots side by side.
function shotRow(files, caption, each = 140) {
  const cells = files.map((f) => {
    const { w, h, data } = imgSize(f);
    return new ImageRun({ type: 'png', data, transformation: { width: each, height: Math.round(h * each / w) } });
  });
  const kids = [];
  cells.forEach((c, i) => { kids.push(c); if (i < cells.length - 1) kids.push(new TextRun('   ')); });
  figNo++; figures.push(`Figure ${chap}.${figNo}  ${caption}`);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120 }, keepNext: true, children: kids }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 240 }, children: [new TextRun({ text: `Figure ${chap}.${figNo}: ${caption}`, italics: true, size: 20 })] }),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: '808080' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(header, rows, widths, caption) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (t, i, head) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: head ? { fill: 'D9E2B6', type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: runs(t, { size: 20, bold: head || undefined }) })],
  });
  tabNo++; tables.push(`Table ${chap}.${tabNo}  ${caption}`);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }, keepNext: true, children: [new TextRun({ text: `Table ${chap}.${tabNo}: ${caption}`, bold: true, size: 20 })] }),
    new Table({
      width: { size: total, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
      rows: [new TableRow({ tableHeader: true, children: header.map((t, i) => cell(t, i, true)) }),
        ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, false)) }))],
    }),
    new Paragraph({ spacing: { after: 160 }, children: [] }),
  ];
}
// A Gantt chart as a table: one row per activity, one column per week. Weeks up
// to `doneThrough` are shaded dark (completed), later ones light (planned). A
// row's optional third entry names weeks that differ: { done: [...] } shaded
// dark, { progress: [...] } shaded mid-green (in progress).
function gantt(rows, doneThrough, caption, weeks = 12) {
  const first = 2826, col = Math.floor((W - first) / weeks);
  const widths = [first, ...Array(weeks).fill(col)];
  const cell = (i, kids, fill) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: i ? 0 : 100, right: i ? 0 : 100 },
    children: [new Paragraph({ alignment: i ? AlignmentType.CENTER : AlignmentType.LEFT, children: kids })],
  });
  const head = new TableRow({ tableHeader: true, children: [
    cell(0, [new TextRun({ text: 'Activity', bold: true, size: 18 })], 'D9E2B6'),
    ...Array.from({ length: weeks }, (_, w) => cell(w + 1, [new TextRun({ text: `W${w + 1}`, bold: true, size: 16 })], 'D9E2B6')),
  ] });
  const shade = (w, on, s = {}) => !on.includes(w) ? undefined
    : (s.progress ?? []).includes(w) ? '8EAA4E'
    : (s.done ?? []).includes(w) || w <= doneThrough ? '4F6228' : 'C4D79B';
  const body = rows.map(([name, on, s]) => new TableRow({ cantSplit: true, children: [
    cell(0, [new TextRun({ text: name, size: 18 })]),
    ...Array.from({ length: weeks }, (_, w) => cell(w + 1, [], shade(w + 1, on, s))),
  ] }));
  tabNo++; tables.push(`Table ${chap}.${tabNo}  ${caption}`);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, pageBreakBefore: true, spacing: { before: 160, after: 80 }, keepNext: true, children: [new TextRun({ text: `Table ${chap}.${tabNo}: ${caption}`, bold: true, size: 20 })] }),
    new Table({ width: { size: first + col * weeks, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED, rows: [head, ...body] }),
    new Paragraph({ spacing: { after: 160 }, children: [] }),
  ];
}
let chap = 0;
let lstNo = 0;
// A code listing: monospace, shaded, kept on one page where possible.
function listing(lines, caption) {
  lstNo++;
  const kids = lines.map((l, i) => new TextRun({ text: l.length ? l : ' ', font: 'Courier New', size: 16, break: i ? 1 : 0 }));
  return [
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED, rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      borders, shading: { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ spacing: { line: 240 }, children: kids })] })] })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 240 }, children: [new TextRun({ text: `Listing ${chap}.${lstNo}: ${caption}`, italics: true, size: 20 })] }),
  ];
}
function chapter(n, title) { chap = n; figNo = 0; tabNo = 0; lstNo = 0; return H1(`Chapter ${n}: ${title}`); }

// ---------- front matter ----------
const cover = [
  ...blank(1),
  C('**LIFETRACK HEALTH**', { run: { size: 40 } }),
  C('A Patient-Owned, Local-First Health Record with Consent-Controlled Sharing', { run: { size: 28, italics: true } }),
  ...blank(1),
  C('NTCC Project Report', { run: { size: 26 } }),
  C('Twelve weeks: 1 September – 23 November 2026', { run: { size: 22, italics: true } }),
  C('submitted in partial fulfilment of the requirements for the degree of', { run: { size: 24 } }),
  C('Bachelor of Science in Information Technology (B.Sc. IT)', { run: { size: 24, bold: true } }),
  ...blank(1),
  C('Submitted by', { run: { size: 24 } }),
  C('**Anish Raj**', { run: { size: 24 } }),
  C('Enrolment No. A45304925031', { run: { size: 24 } }),
  ...blank(1),
  C('Under the supervision of', { run: { size: 24 } }),
  C('**Dr. Priyadarshini Gupta**', { run: { size: 24 } }),
  C('Assistant Professor, Amity Institute of Information Technology', { run: { size: 24 } }),
  ...blank(2),
  C('**AMITY INSTITUTE OF INFORMATION TECHNOLOGY**', { run: { size: 26 } }),
  C('**AMITY UNIVERSITY PATNA**', { run: { size: 26 } }),
  C('Patna – 801503', { run: { size: 24 } }),
  C('Academic Session 2025–2028  ·  November 2026', { run: { size: 24 } }),
];

const declaration = [
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('DECLARATION')] }),
  P('I hereby declare that the project report entitled **“LifeTrack Health: A Patient-Owned, Local-First Health Record with Consent-Controlled Sharing”**, submitted to Amity Institute of Information Technology, Amity University Patna, in partial fulfilment of the requirements for the degree of Bachelor of Science in Information Technology (B.Sc. IT), is a record of original work carried out by me under the supervision of Dr. Priyadarshini Gupta.'),
  P('The software described in this report was designed, built and tested by me. Published standards, laws and libraries used in the work are acknowledged in the References, and the software tools used, including an AI coding assistant, in the Acknowledgement. This report has not been submitted, in part or in full, to any other university or institution for the award of any degree or diploma.'),
  ...blank(2),
  P('**Anish Raj**', { align: AlignmentType.LEFT }),
  P('Enrolment No. A45304925031', { align: AlignmentType.LEFT }),
  P('Signature: ____________________', { align: AlignmentType.LEFT }),
  P('Place: Patna\t\t\t\t\t\tDate: ____________', { align: AlignmentType.LEFT }),
];

const certificate = [
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('CERTIFICATE')] }),
  P('This is to certify that the project report entitled **“LifeTrack Health: A Patient-Owned, Local-First Health Record with Consent-Controlled Sharing”**, submitted by Anish Raj (Enrolment No. A45304925031) in partial fulfilment of the requirements for the degree of Bachelor of Science in Information Technology (B.Sc. IT) at Amity Institute of Information Technology, Amity University Patna, is a bona fide record of work carried out under my supervision during the academic session 2025–2028.'),
  P('To the best of my knowledge, the work presented in this report is original and has not been submitted elsewhere for the award of any degree or diploma.'),
  ...blank(4),
  P('**Dr. Priyadarshini Gupta**', { align: AlignmentType.LEFT }),
  P('Assistant Professor', { align: AlignmentType.LEFT }),
  P('Amity Institute of Information Technology, Amity University Patna', { align: AlignmentType.LEFT }),
  P('Signature: ____________________\t\tDate: ____________', { align: AlignmentType.LEFT }),
];

const AB = { para: { spacing: { after: 120, line: 320 } } };
const ack = [
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('ACKNOWLEDGEMENT')] }),
  P('I express my sincere gratitude to my supervisor, **Dr. Priyadarshini Gupta**, for her guidance, encouragement and critical feedback throughout this project. Her questions about how an analytical or clinical claim can be proven correct shaped the testing approach described in this report.'),
  P('I thank the Director and faculty of Amity Institute of Information Technology, Amity University Patna, for providing the environment and resources to carry out this work,.'),
  P('I also acknowledge the open-source communities behind Flutter, Dart, the `cryptography` package, Tesseract OCR and the many other libraries this project builds on.'),
  P('Parts of the code and of this report were drafted with the help of an AI coding assistant (Claude, by Anthropic). I directed the work, reviewed what it produced, and ran and checked the tests reported here.'),
  P('Finally, I thank my family for their patience and support.'),
  ...blank(2),
  P('Anish Raj', { align: AlignmentType.RIGHT }),
];

const abstract = [
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('ABSTRACT')] }),
  P('In India a patient’s medical history is usually spread across the clinics, hospitals and laboratories they have visited, with the patient carrying paper reports between them. When the history is not available, investigations are repeated, medicines are duplicated and family patterns are missed. The cause is usually scattered data, not poor clinical judgement. The patient is the only person present at every encounter, yet rarely holds a complete record.', AB),
  P('LifeTrack Health is a working prototype of a health record that the patient owns and keeps on their own phone. It has four parts: an Android patient app built in Flutter; LifeTrack Clinic, a doctor app; a small Dart server that acts as a doctor registry, an encrypted relay and a backup store; and a shared protocol package. Records live on the patient’s phone, sealed with AES-256-GCM under a device key protected by the Android Keystore. A doctor sees records only after the patient scans the doctor’s registry-verified code and approves a level of access (Summary, All reports, Family illnesses or Home readings) for a chosen time between one hour and two days. The share is encrypted on the phone to that doctor’s device key alone (X25519 with AES-GCM); the server carries it but holds no key that opens it. Prescriptions return to the patient signed with the doctor’s Ed25519 key.', AB),
  P('The patient app also reads photographed or PDF lab reports on the device with Tesseract OCR, compares results with reference ranges without stating a diagnosis, checks every saved value against a table of danger limits and prompts the patient to seek care, tracks medicines with reminders, keeps a child vaccination schedule, manages a family under a head-of-family model, offers an Emergency ID readable from the lock screen, and makes an end-to-end encrypted backup protected by an app-generated password and Shamir 2-of-3 recovery codes. Every read, share, refusal and export is written to an append-only audit log, and anything that leaves the phone is logged before it leaves.', AB),
  P(`The system is written in Dart and Flutter (about ${STATS.code.approx} lines of application code). It has ${STATS.tests.total} automated tests at unit, widget, invariant, layout and end-to-end integration level, all passing with a clean static analysis (${STATS.tests.skipped} further test is skipped). The patient app’s interface was reviewed screen by screen twice and then restyled, and tests lay out every main screen at three phone widths and three text sizes and keep the health-status colours distinguishable for colour-blind users. The invariant tests turn the project’s main privacy rules into checks that fail the test run: no screen may claim the server “cannot read” the data, for example, and no agent may issue itself a grant. The system has not been reviewed by a doctor or tried with real users. The main open items are that clinical review, a usability check, OCR accuracy on real reports, ABDM/ABHA integration and a legal review before any pilot.`, AB),
  P('**Keywords:** personal health record, patient consent, end-to-end encryption, local-first software, Flutter, digital health India, DPDP Act 2023.', AB),
];

// The Word file carries a contents field that Word fills in on opening. For the
// PDF, make_pdf.sh measures where each heading lands and passes TOC_PAGES
// (a JSON map of heading → page), and the contents are printed as text.
const tocPages = process.env.TOC_PAGES ? JSON.parse(fs.readFileSync(process.env.TOC_PAGES, 'utf8')) : null;
const tocTitle = new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('TABLE OF CONTENTS')] });
const toc = [tocTitle, new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-2' })];
const printedToc = () => [tocTitle, ...headings.map(([level, text]) => new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: W, leader: 'dot' }],
  indent: { left: level === 1 ? 0 : 400 },
  spacing: { before: level === 1 ? 120 : 0, after: 40 },
  children: [new TextRun({ text, bold: level === 1, size: 22 }), new TextRun({ text: `\t${tocPages[text] ?? ''}`, bold: level === 1, size: 22 })],
}))];

// ---------- Chapter 1 ----------
const ch1 = [
  chapter(1, 'Introduction'),
  H2('1.1 Background'),
  P('In India, health care comes from a very large number of independent providers: government hospitals, private hospitals, single-doctor clinics, diagnostic laboratories and pharmacies. Each keeps its own records, if it keeps them digitally at all, and very few exchange them. For the patient, the practical result is a carry bag of paper (prescriptions, lab printouts, discharge summaries and scan reports) that is incomplete, easily lost and hard for a new doctor to read in a ten-minute consultation.'),
  P('The Government of India’s Ayushman Bharat Digital Mission (ABDM) is building national rails for health data exchange, including the ABHA health ID and consent managers. At the same time, the Digital Personal Data Protection Act, 2023 (DPDP Act) makes consent, purpose limitation and the right to access one’s data legal obligations. Both point in the same direction: the patient should be at the centre of their own record and should decide who sees it.'),
  H2('1.2 Problem statement'),
  P('A patient is the only person present at every one of their medical encounters, but has no single, trustworthy, consolidated view of their own history and no safe way to show the relevant part of it to a new doctor. Three consequences follow:'),
  B('**Repeated investigations.** When a previous test result is not available, or its source cannot be trusted, the test is ordered again, costing the patient money and time.'),
  B('**Duplicate or conflicting medication.** A doctor who does not know what the patient already takes may prescribe the same ingredient under a different brand name.'),
  B('**Missed hereditary patterns.** Illnesses in the family are rarely recorded in a way a doctor can use.'),
  P('Existing solutions either keep the data inside one institution (hospital EMRs), are designed for large organisations (enterprise systems), establish identity without holding the record (government health IDs), or store documents without making them useful or safely shareable (document-locker apps). The project therefore asks: **can a patient keep a complete, useful health record on their own phone and share exactly the right part with exactly the right doctor, for a limited time, without any server holding a key that opens it?**'),
  H2('1.3 Motivation'),
  P('Health data is among the most sensitive personal data. A design in which a company’s servers hold readable copies of millions of records creates a single point of failure for breaches, misuse and function creep. In a local-first design the main copy lives on the patient’s device and leaves it only encrypted, and only when the patient asks. That removes the central store of readable records. Building such a system end to end, with the cryptography, the consent rules and a doctor-side app, requires work across mobile development, security, distributed systems, user-interface design and law, which made it a suitable scope for this project.'),
  H2('1.4 Objectives'),
  N('Build a patient-owned record on the phone that holds prescriptions, lab reports, scans, notes, allergies, conditions, vitals and vaccinations for every member of a family, with all data encrypted at rest.'),
  N('Make clinician access explicit, time-limited and revocable: the patient approves a level of access for one verified doctor for a chosen time, and can shorten, extend or end it.'),
  N('Give no server a key that opens shared records or backups: shares are encrypted to the doctor’s device key and backups to a key derived from a password only the patient holds.'),
  N('Provide useful, safe insights on the device: compare results with published reference ranges, notice movement between reports and spot duplicate medicines, without ever stating a diagnosis or a forecast.'),
  N('Reduce typing by reading photographed and PDF lab reports on the device (OCR), with the patient confirming uncertain fields.'),
  N('Build a doctor app with a verified identity, so that a patient knows who they are sharing with, and so that prescriptions returned to the patient are signed.'),
  N('Make the main privacy promises verifiable: express the system’s rules as automated tests that fail when broken.'),
  N('Design an interface that older patients and caregivers can use: readable on small screens and at large text sizes, with WCAG AA contrast in light and dark mode, and health-status colours that colour-blind users can tell apart.'),
  H2('1.5 Scope'),
  P('**In scope.** An Android patient app; a doctor app (Flutter, runs in a browser or on a phone); a local server providing a doctor registry, an encrypted relay, a LifeTrack ID directory, a backup vault and an identity-check flow; and a shared protocol package. Sample (synthetic) patients and doctors are used throughout.'),
  P('**Out of scope for this phase.** Real patient data; live ABDM/ABHA exchange; real DigiLocker partner integration (a “practice” stand-in drives the same flow); cloud-based AI analysis (designed as “Plane C” but disabled); iOS OCR; clinical decision support and predictive modelling; deployment on the public internet.'),
  H2('1.6 Methodology'),
  P('Development followed an incremental model in six stages. Each stage ended with working, tested software rather than documents alone.'),
  N('Requirements and threat analysis: the patient, doctor and family needs listed in Chapter 3, together with the ways health data can leak (a readable server copy, an over-broad grant, a misleading claim on screen).', 'num3'),
  N('Architecture and invariants: the three-plane structure, the single consent authority and eighteen written invariants (rules the code must never break), tested wherever a test can check them.', 'num3'),
  N('Incremental feature releases: each feature (records, insights, sharing, reminders, family, backup, OCR and so on) was built as a separate version with its own tests and a changelog line.', 'num3'),
  N('Security review rounds: two dedicated hardening releases, and a full review of the server, consent, sharing and backup code on 2 November 2026, fixed issues found by reviewing the code against the invariants. An audit on 6 November 2026, done with the AI assistant named in the Acknowledgement, led to the clinical-safety and reliability fixes in patient 1.43.0–1.48.2 and doctor 1.8.2–1.9.1. It was not a clinical sign-off: no doctor has reviewed the system.', 'num3'),
  N('Interface design and review: two screen-by-screen reviews of the patient app, each followed by a series of small releases, then a visual restyle. Layout, contrast and colour-vision tests keep the result from regressing.', 'num3'),
  N('Integration and device testing: both apps were run against the real server handler in automated integration tests, and the release build was tested by hand on an Android phone.', 'num3'),
  P('The guiding rule across all stages, taken from the project proposal, was that a claim about the system is accepted only when a test can prove it.'),
  H2('1.7 Project schedule'),
  P(`The project is planned over twelve weeks, from 1 September to 23 November 2026. Table 1.1 lists the work for each week and its status on ${STATS.date}, when this report was last updated. Weeks 1 to 7 are complete. In week 8 frame timing was measured on one phone only, on 2 November, so that work is only partly done. In week 9 the third security review was done, but the usability check was not, and it has moved to weeks 10 and 11 with the remaining phone tests. On 2 November the server also moved to a Windows PC. Week 10 added an audit and the fixes that followed it, and the user guide. The OCR check on real reports, the final review and submission remain. Table 1.2 shows the same plan as a Gantt chart.`),
  P('The Git history begins on 29 September 2026, so it does not show weeks 1 to 4. The original repository was never pushed to a remote and was lost around 23 September. [[Anish to confirm the date and cause]] The code was rebuilt from the project blueprint in week 5, which is why the first releases of both apps fall within a few days.'),
  ...table(['Week', 'Dates', 'Work', 'Status'], [
    ['1', '1–7 Sep', 'Problem study; review of existing systems, standards and Indian law (DPDP Act, ABDM, Aadhaar Act)', 'Completed'],
    ['2', '8–14 Sep', 'Requirements: stakeholders, functional and non-functional requirements, use cases, threat analysis', 'Completed'],
    ['3', '15–21 Sep', 'System design: architecture and the three planes, consent model, data design, written invariants', 'Completed'],
    ['4', '22–28 Sep', 'Security design (sealed storage, backup key hierarchy, clinical-share protocol); HTML prototype of the screens', 'Completed'],
    ['5', '29 Sep – 5 Oct', 'Patient app 1.0.0–1.19.0, LifeTrack Clinic 1.0.0–1.8.0 and the server; two security review rounds; integration tests; first draft of this report', 'Completed'],
    ['6', '6–12 Oct', 'Two interface reviews and their fixes; visual restyle; colour-vision check; family tree; server over the home Wi-Fi (patient 1.19.1–1.42.0)', 'Completed'],
    ['7', '13–19 Oct', 'Layout tests at three widths and three text sizes; faster drawing; full test run; report update (patient 1.42.1–1.42.2)', 'Completed'],
    ['8', '20–26 Oct', 'Frame timing and manual tests on more Android phones; fixes', 'Partly done: one phone, on 2 Nov'],
    ['9', '27 Oct – 2 Nov', 'Third security review; usability check with family members and classmates, including older users; fixes', 'Review done; usability check moved'],
    ['10', '3–9 Nov', 'Audit and fixes (patient 1.43.0–1.48.2, doctor 1.8.2–1.9.1); user guide; remaining phone tests, usability check and OCR check on real reports', 'In progress'],
    ['11', '10–16 Nov', 'Final review with the guide; report revisions; presentation slides', 'Planned'],
    ['12', '17–23 Nov', 'Final build, report submission and presentation', 'Planned'],
  ], [700, 1500, 5626, 1200], 'Twelve-week project schedule'),
  ...gantt([
    ['Problem study and literature review', [1, 2]],
    ['Requirements analysis', [2, 3]],
    ['System and security design', [3, 4]],
    ['Screen prototype', [4]],
    ['Implementation (apps and server)', [5, 6, 7, 10], { progress: [10] }],
    ['Security reviews', [5, 9], { done: [9] }],
    ['Interface design and review', [6, 7]],
    ['Automated and device testing', [5, 6, 7, 8, 10], { progress: [8, 10] }],
    ['Usability check', [9, 10, 11]],
    ['OCR check and user guide', [10], { progress: [10] }],
    ['Report writing', [5, 7, 10, 11], { progress: [10] }],
    ['Final review, presentation, submission', [11, 12]],
  ], 7, 'Gantt chart of the twelve-week schedule (dark: completed; mid: in progress or partly done; light: planned, not yet done)'),
  H2('1.8 Organisation of the report'),
  P('Chapter 2 reviews existing systems, standards and laws. Chapter 3 sets out the requirements and feasibility. Chapter 4 describes the architecture, data design and security design. Chapter 5 explains the implementation, module by module, with screenshots. Chapter 6 covers testing and results. Chapter 7 discusses outcomes against the objectives, and Chapter 8 concludes with limitations and future work. Appendices give the version history and instructions to run the system.'),
];

// ---------- Chapter 2 ----------
const ch2 = [
  chapter(2, 'Literature Review and Existing Systems'),
  H2('2.1 Categories of existing solutions'),
  ...table(['Category', 'Examples', 'Strengths', 'Gaps for the patient'], [
    ['Hospital EMR / HIS', 'Systems used inside a hospital chain', 'Complete within the institution; integrated with billing and labs', 'Data stays inside the institution; patient has no consolidated view'],
    ['Enterprise health platforms', 'Large vendor suites', 'Powerful, standards-based', 'Expensive and complex; built for organisations, not individuals'],
    ['Government health ID and exchange', 'ABHA, ABDM consent managers, PHR apps', 'National identity and consent rails; legal backing', 'Identity and exchange rather than a patient-held, analysed record; depends on provider adoption'],
    ['Document lockers', 'DigiLocker and general cloud storage', 'Easy storage of issued documents', 'Stores files, does not interpret results; provider holds readable copies'],
    ['Personal record apps', 'Consumer health-record apps', 'Convenient capture and reminders', 'Usually cloud-first with readable server copies; sharing is often an open link'],
  ], [1700, 2000, 2400, 2926], 'Existing solution categories compared'),
  H2('2.2 Relevant standards'),
  B('**HL7 FHIR R4** [8]. HL7’s interoperability standard models clinical data as typed resources such as Patient, Observation, MedicationRequest and Provenance. ABDM uses FHIR-based profiles. LifeTrack’s record model keeps provenance and source on every record so that a later FHIR export can say where each value came from; FHIR export itself is future work (Section 8.3).'),
  B('**Ayushman Bharat Digital Mission (ABDM)** [6]. Defines the ABHA number, health information providers and users, and consent artefacts. LifeTrack records an ABHA number as optional, typed in, and shown as “not verified” until ABDM integration exists.'),
  B('**Cryptographic standards.** AES-GCM (NIST SP 800-38D) [9] for authenticated encryption; X25519 (RFC 7748) [10] for key agreement; Ed25519 (RFC 8032) [11] for signatures; HKDF (RFC 5869) [12] for key derivation; Argon2id (RFC 9106) [13] for password hashing, using the OWASP-recommended minimum parameters [14]; Shamir’s secret sharing [15] for recovery codes.'),
  B('**WCAG 2.1 AA** [19]. Contrast requirements for readable text; the apps are tested against them in both light and dark mode.'),
  H2('2.3 Legal and regulatory context'),
  B('**Digital Personal Data Protection Act, 2023** [1]. Under the Act, a person must agree knowingly and voluntarily to each use of their data, data may be used only for the purpose agreed, people may see and erase what is held about them, and a parent must consent for a child (Section 9). LifeTrack implements consent per purpose and time, an export of all data (the access right), erasure, and parent-declared consent for children that stops at 18. The DPDP Rules, 2025 [23] were notified in November 2025 and take effect in phases; most duties on those who handle personal data apply from about May 2027, 18 months after publication.'),
  B('**Aadhaar Act, 2016, Section 29** [2] restricts storing and sharing Aadhaar numbers. LifeTrack never asks for or stores an Aadhaar number; LifeTrack IDs always begin with 1 so that they can never be, or validate as, an Aadhaar number.'),
  B('**Justice K.S. Puttaswamy v. Union of India (2018)** [3] and the 2019 Aadhaar amendment [2] limit mandatory Aadhaar authentication by private entities. The project’s optional “identity required” mode, which verifies every profile through DigiLocker, is therefore flagged for legal review before any pilot (Section 8.2).'),
  B('**Telemedicine Practice Guidelines, 2020** [5] require a doctor to keep a record of the consultation and prescription. LifeTrack therefore never promises that a doctor “keeps no copy”; it says that shared records close when access ends.'),
  B('**Information Technology (SPDI) Rules, 2011** [4] cover sensitive personal data, including health data. They rest on Section 43A of the IT Act, which Section 44(2) of the DPDP Act removes; that provision is timed for about May 2027, and legal commentators differ on the exact position until then. This is one of the points for the legal review in Section 8.2.'),
  H2('2.4 Personal health records in the research literature'),
  P('Tang et al. [24] set out what a personal health record is and what stands in the way of its adoption. They concluded that records connected to a provider’s electronic system give much greater benefit than stand-alone ones. That is a direct challenge to this project: LifeTrack is stand-alone until ABDM integration exists, and its value depends on the patient entering records and on doctors trusting them. The provenance shown on every result (Section 4.5) responds to the second problem but does not remove the first.'),
  P('Roehrs et al. [25] screened more than 5,000 studies from a ten-year period in a systematic review, and organised the field into a taxonomy and a list of open challenges. Azaria et al. [26] proposed MedRec, which keeps access permissions for medical records on a blockchain so that no single institution controls them. LifeTrack shares that aim of patient control but needs no ledger, because the record stays on the phone and each share is a time-limited grant between one patient and one doctor. The trade-off is that LifeTrack has no shared log that a third party can inspect: its audit log lives on the patient’s phone.'),
  H2('2.5 Design principles drawn from the literature'),
  P('Kleppmann et al. [17] describe local-first software, in which the main copy of the data stays on the user’s own device and the network is used only to sync it and work with others. End-to-end encryption, familiar from messaging, shows that a relay can carry data it cannot read. Clinical-safety guidance for software as a medical device [21] distinguishes between displaying information (low risk) and diagnosing or predicting (regulated). LifeTrack combines these: a local-first record, an end-to-end encrypted relay, and insights that state only position and movement and route the patient to a person.'),
  H2('2.6 Research gap'),
  P('Among the categories of systems reviewed here, which is not an exhaustive survey of products, none was found that combines (a) a patient-held, encrypted, family-aware record, (b) on-device interpretation of results that is deliberately non-diagnostic, (c) sharing to a registry-verified doctor that the server cannot read, with a time limit the patient sets, and (d) privacy promises enforced by automated tests. LifeTrack Health is a prototype of that combination.'),
];

// ---------- Chapter 3 ----------
const ch3 = [
  chapter(3, 'Requirements Analysis'),
  H2('3.1 Stakeholders'),
  ...table(['Stakeholder', 'Needs'], [
    ['Patient / head of family', 'Keep everyone’s records in one place; understand results; share safely with a doctor; get reminders; recover the record if the phone is lost'],
    ['Doctor', 'See a trustworthy summary quickly, with where each result came from; ask a patient for access remotely; send a prescription back'],
    ['Registry administrator', 'Approve, reject and revoke doctors; keep the registry key safe'],
    ['Linked family member', 'See a parent’s Emergency ID and current medicines when the parent chooses to share them'],
    ['Paramedic / bystander', 'Read the Emergency ID from the lock screen in an emergency'],
  ], [2600, 6426], 'Stakeholders and their needs'),
  H2('3.2 Functional requirements'),
  H3('Patient app'),
  ...table(['ID', 'Requirement'], [
    ['FR-1', 'Create profiles for the phone owner (head of family) and family members with roles; children only as the head’s son or daughter'],
    ['FR-2', 'Add, view, edit and delete prescriptions, reports, scans and notes; keep the original image or PDF with each record; enter allergies, conditions and family illnesses, where “no known allergies” is a stated choice and an empty list shows as “not recorded”'],
    ['FR-3', 'Read a photographed or PDF lab report on the device and turn it into a draft record; uncertain fields must be confirmed by the patient'],
    ['FR-4', 'Show each test’s results over time; compare with published reference ranges for age and sex; mark unit mismatches as “not checked”; check every saved lab value and home reading against danger limits and prompt the patient to seek care'],
    ['FR-5', 'Detect duplicate medicines (same ingredient, different brand) and meaningful changes between reports'],
    ['FR-6', 'Medicine reminders from active prescriptions, plus the patient’s own reminders; names hidden on the lock screen by default'],
    ['FR-7', 'Child vaccination schedule (India Universal Immunisation Programme) with due dates from date of birth'],
    ['FR-8', 'Home readings diary for blood pressure and sugar, with charts'],
    ['FR-9', 'Emergency ID (blood type, conditions, medicines, contact) readable from the lock screen, every read logged; plain-text QR'],
    ['FR-10', 'Share with a doctor by scanning the doctor’s verified code; choose level and time (1 h, 4 h, 12 h, 1 day, 2 days); change or end it later'],
    ['FR-11', 'Receive and approve or decline access requests a doctor sends by LifeTrack ID'],
    ['FR-12', 'Collect signed prescriptions from the doctor and file them as doctor-issued'],
    ['FR-13', 'Encrypted backup of all profiles to the server; restore onto an empty phone with the password or recovery codes; retry a restore that was interrupted'],
    ['FR-14', 'Export all data as a zip (HTML, JSON, originals, access log), password-protected by default'],
    ['FR-15', 'Access log showing every read, share, refusal, export and Emergency ID open'],
    ['FR-16', 'Family link: share Emergency ID and/or medicine list with a linked adult’s phone, each off by default'],
  ], [900, 8126], 'Functional requirements: patient app'),
  H3('Doctor app and server'),
  ...table(['ID', 'Requirement'], [
    ['FR-17', 'Doctor enrols with name, registration number and specialty; the device generates its own keys; the administrator approves'],
    ['FR-18', 'Doctor shows a verified QR code for the patient to scan; certificate lasts 30 days and renews automatically'],
    ['FR-19', 'Doctor requests access by a patient’s 12-digit LifeTrack ID, asking for levels and a time'],
    ['FR-20', 'Doctor sees shared records with provenance and a countdown; access closes at the end time'],
    ['FR-21', 'Doctor writes, signs, prints and sends a prescription back to the patient'],
    ['FR-22', 'Doctor keeps their own sealed notes and a patient register'],
    ['FR-23', 'Server: registry with admin approve/reject/revoke; relay that deletes on pickup or expiry; ID directory; backup vault with rate-limited proofs; identity-check flow'],
  ], [900, 8126], 'Functional requirements: doctor app and server'),
  H2('3.3 Non-functional requirements'),
  ...table(['Quality', 'Requirement and how it is met'], [
    ['Confidentiality', 'Every blob on disk sealed with AES-256-GCM; shares and backups end-to-end encrypted; server holds no key that opens records'],
    ['Integrity', 'Authenticated encryption binds each blob to its name; prescriptions and certificates are signed; audit log is append-only'],
    ['Accountability', 'Every read, change, grant, refusal, egress, restore and erasure is audited; egress is logged before it happens'],
    ['Availability / offline', 'The patient app works fully offline; the network is needed only to share, back up, or check a doctor’s status. Saves do not wait for the network, and every network call has a time limit'],
    ['Usability', 'Plain-language screens; one confirmation for scanned reports; levels as presets; labelled tabs; swipe to tick a medicine; touch targets of at least 48 dp; dark mode'],
    ['Accessibility', 'All visible text meets WCAG AA contrast in light and dark mode (tested); secondary text at least 14 sp; every main screen laid out at 320, 360 and 411 dp wide with text at up to twice the normal size (tested); health-status colours kept at least 25 ΔE apart from the accent under every simulated colour-vision deficiency (tested)'],
    ['Performance', 'Shared widgets use the renderer’s fast paths (rounded-rectangle clips and shadows, one shared background blur); photos decoded at the size shown; a frame-timing test runs on the phone in a profile build'],
    ['Safety', 'Insights never diagnose or forecast; a finding with a banned phrase is dropped, except that an urgent alert falls back to fixed wording rather than disappearing; no dose is ever filled in by the app; adult ranges never applied to children or in pregnancy'],
    ['Maintainability', `Strict static analysis; ${STATS.tests.total} automated tests; every update raises the version and adds a changelog line`],
    ['Portability', 'Flutter: one code base for the patient app (Android) and the doctor app (web or mobile)'],
  ], [1900, 7126], 'Non-functional requirements'),
  H2('3.4 Hardware and software requirements'),
  ...table(['Item', 'Development', 'Deployment (prototype)'], [
    ['Hardware', 'Laptop with 8 GB+ RAM; Android phone with USB debugging', 'Android 8.0+ phone with a screen lock; any modern browser for the doctor app'],
    ['Operating system', 'macOS / Windows / Linux', 'Android; server runs on any OS with Dart'],
    ['Languages', 'Dart 3, Kotlin (Android platform code)', 'Same'],
    ['Frameworks', 'Flutter, shelf (server)', 'Same'],
    ['Tools', 'Android SDK, adb, Git and GitHub, VS Code / Android Studio', 'Not needed'],
    ['Key libraries', 'cryptography, flutter_secure_storage, local_auth, flutter_zxing, qr_flutter, Tesseract4Android, flutter_local_notifications, archive, printing', 'Bundled in the apps'],
  ], [1800, 3800, 3426], 'Hardware and software requirements'),
  H2('3.5 Use cases'),
  P('Figure 3.1 shows the actors and the use cases, numbered as in Table 3.6. The patient (as head of the family) has the most; a doctor reaches the system only through verified enrolment and patient-approved shares.'),
  ...figure(path.join(FIG, 'usecase.png'), 'Use-case diagram', 600, 520),
  ...table(['ID', 'Use case', 'Actor', 'Main flow'], [["UC-1", "Add or scan a record", "Patient", "Types a record, or photographs/imports a report; OCR fills a draft; the patient ticks uncertain fields and confirms."], ["UC-2", "Share with a doctor", "Patient", "Scans the doctor’s code; the phone verifies the certificate; the patient picks a level and time; the encrypted bundle is sent."], ["UC-3", "Respond to an access request", "Patient", "A signed request from a verified doctor appears; the patient approves (possibly at a lower level or shorter time) or declines."], ["UC-4", "End or change access", "Patient", "From “Who has access now”, shortens, extends or ends a share; the doctor’s copy is replaced or closed."], ["UC-5", "Back up and restore", "Patient", "Turns on backup and receives the generated password and recovery codes; on a new, empty phone restores with them."], ["UC-6", "Read the Emergency ID", "Paramedic", "Opens the card from the lock screen without unlocking; the read is logged."], ["UC-7", "Enrol and get verified", "Doctor", "Enters registration details; the administrator approves; the device receives a signed 30-day certificate."], ["UC-8", "View records and prescribe", "Doctor", "Opens an arrived share, reviews the summary, writes and signs a prescription that returns to the patient."], ["UC-9", "Approve or revoke a doctor", "Administrator", "Reviews pending enrolments and approves, rejects or revokes them with the admin token."], ["UC-10", "View insights", "Patient", "Turns insights on for 24 hours; reads each finding with its severity and the lab’s printed range; every read is logged."], ["UC-11", "Medicine reminders", "Patient", "Gets reminders from current prescriptions by morning, afternoon and night, ticks a dose by swiping, and adds own reminders; names stay off the lock screen unless turned on."], ["UC-12", "Manage family", "Patient", "Adds a child as son or daughter (consent declared), an adult who delegates in person, or invites an adult with their own phone; sees the family tree."], ["UC-13", "Export data", "Patient", "Exports a zip (readable pages, data, originals, access log), password-protected by default, to the app the patient picks; logged before it is built."], ["UC-14", "Share Emergency ID or medicines", "Patient, linked family member", "Two adults link phones and compare a safety code; each ticks the Emergency ID and/or medicine list to share; unlinking deletes the other phone’s copy."], ["UC-15", "Request access by ID", "Doctor", "Sends a signed request to a LifeTrack ID; the phone shows it only after the certificate and signature verify; nothing is shared until the patient approves."]], [800, 1900, 1550, 4776], 'Use cases (the same IDs and names as Figure 3.1)'),
  H2('3.6 Feasibility study'),
  B('**Technical.** Flutter and Dart [20] provide mature cryptography, camera, QR and notification support; Tesseract runs fully on-device. Every required flow runs end to end in the automated integration tests.'),
  B('**Economic.** All tools and libraries are free and open source. The server is a single small Dart process; hosting costs would be minimal because it stores only ciphertext and short-lived messages.'),
  B('**Operational.** Patients already use smartphones and QR codes; doctors already use computers in clinics. The share flow requires one scan and one approval.'),
  B('**Legal.** The design follows DPDP principles and avoids storing Aadhaar numbers. Mandatory identity verification and the clinical status of insights need confirmation from legal counsel before a pilot.'),
];

// ---------- Chapter 4 ----------
const ch4 = [
  chapter(4, 'System Design'),
  H2('4.1 Overall architecture'),
  P('LifeTrack Health has four parts (Figure 4.1). The patient app holds the record. The doctor app, LifeTrack Clinic, is the only other place records can be read, and only while a share is open. The server is kept small and is never trusted with readable content: it signs doctor certificates, carries sealed messages, maps LifeTrack IDs to mailboxes, stores encrypted backups and runs the identity-check flow. A shared Dart package, lifetrack_protocol, defines the message formats, certificates and ID rules used by all three, so that the apps and the server cannot drift apart.'),
  ...figure(path.join(FIG, 'arch.png'), 'System architecture', 560),
  H2('4.2 Data flow (Level 0)'),
  P('Figure 4.2 shows the system as a single process with its external entities. Records enter only from the patient (typed, scanned, imported) or from a doctor (signed prescriptions). Records leave only in three ways: a clinical share the patient approved, encrypted backup ciphertext, or a family-link message the patient ticked. The patient can also export their own data to an app they choose (Section 3.2, FR-14).'),
  ...figure(path.join(FIG, 'dfd0.png'), 'Level 0 data flow diagram', 600),
  H2('4.3 Patient app architecture'),
  P('Inside the patient app (Figure 4.3), screens never touch storage or compare values with ranges themselves. They call AppServices, the single entry point, and render Insight objects. Any reader other than the patient (on-device insights, a share, the disabled cloud plane) must go through RecordAccess, which asks ConsentStore whether a stored grant allows the read and writes an audit entry for every read and every refusal. Below that, SealedStore encrypts every blob and FileByteStore writes it atomically (write to a temporary file, flush, rename).'),
  ...figure(path.join(FIG, 'layers.png'), 'Layers of the patient app', 600),
  H3('The three planes'),
  B('**Plane A: on-device insights (Tier 1).** Fixed rules that state where a value sits relative to a published range, whether it moved meaningfully since the last report, and whether two prescriptions share an ingredient. It reads records only under its own actor and a 24-hour grant the patient turns on, and it is forbidden from calling the clock directly so that its output is reproducible in tests.'),
  B('**Plane B: capture.** Turns a photographed or PDF page into a draft record using Tesseract OCR on the phone. Fields read with low confidence must be ticked or corrected before one final confirmation.'),
  B('**Plane C: cloud analysis (Tier 2).** Designed but switched off. It can only cross on a fresh human tap bound to the disclosure the patient read, with its own one-time grant. It stays dark until a certified service exists.'),
  H2('4.4 Consent model'),
  P('A grant (consent token) records the patient, who granted it (the patient, or a parent or custodian for a dependent), the grantee, the purpose, the level and filter, the validity window, and whether it is revoked or spent. Authorisation always looks up the stored grant by id and re-checks every field; a token presented by a caller is never trusted on its own. Further rules:'),
  B('No grant exceeds 48 hours. A clinical share lasts one hour unless the patient chooses 4 h, 12 h, 1 day or 2 days; a doctor may ask for a time, and the patient may give less or change it later.'),
  B('A live grant cannot be silently re-issued with any difference; changing the time replaces the doctor’s copy with one that ends at the new time.'),
  B('No component under the AI planes may create, extend or exercise a grant it was not given.'),
  B('Proxy consent applies only to dependents: a parent for a child under 18 (sharing stops at 18), or a custodian for an adult who delegated in person. One function, actingGrantor(), makes that decision for every share.'),
  ...table(['Level', 'What the doctor sees'], [
    ['Summary (default)', 'Allergies, conditions, active medicines and the latest result for each test; the level aimed at preventing repeat tests'],
    ['All my reports', 'Every report, prescription and scan, optionally filtered by date or test'],
    ['Family illnesses', 'Illnesses in the family as the patient reports them, marked “patient-reported”'],
    ['Home readings', 'Blood pressure and sugar readings from the last 30 days; off unless ticked'],
  ], [2400, 6626], 'Access levels'),
  H2('4.5 Data design'),
  P('The main entities are Patient (profile, date of birth, sex, ABHA, role), Custody (who may consent for whom, and the family role), HealthRecord (kind: prescription, report, scan or note; source: typed by patient, smart capture, doctor-issued, restored, lab-signed; provenance badge; test results with the lab’s printed range; medicines with status), ClinicalProfile (allergies, conditions, family illnesses), VitalReading, EmergencyCard (fixed, bounded fields), ConsentToken, and AuditEntry. Every record keeps where it came from, because a doctor will reuse a result instead of re-ordering the test only if it can be traced to a laboratory. Results show their age and never “valid” or “expired”.'),
  P('Each entity is serialised to JSON as `{schemaVersion, body}` and sealed into a frame `“LT” ‖ version ‖ nonce ‖ MAC ‖ ciphertext`. Version 2 frames bind the blob’s name as associated data, so a blob copied under another name will not open. The index of blob names is itself sealed.'),
  H2('4.6 Security design'),
  H3('Clinical share'),
  P('Figure 4.4 shows the complete share sequence. The doctor’s device generates its own key pairs; the registry signs a certificate binding the doctor’s registration number to those keys. The patient’s phone verifies the certificate against the registry key pinned in the app and, when online, asks the registry that the doctor is still approved with the same keys. Only then does it show the doctor’s name and registration number. The share is encrypted on the phone to the doctor’s X25519 key; the audit entry is written before the bundle leaves; the relay deletes it after pickup or at expiry.'),
  ...figure(path.join(FIG, 'seq.png'), 'Clinical share and prescription sequence', 600, 600),
  H3('Backup vault'),
  P('User-chosen passwords are typically weak, and in an encrypted backup the password is the weakest link. LifeTrack therefore generates the password: 20 characters, about 100 bits of entropy, shown once with an account code and recovery codes. Figure 4.5 shows the key hierarchy. Argon2id stretches the password; HKDF derives a vault key and a login verifier. The vault key wraps a random data key, which encrypts the backup. The data key is also split into three Shamir shares, any two of which recover it. The server keeps only a hash of the verifier and rate-limits wrong proofs to ten per hour, a limit that survives restarts.'),
  ...figure(path.join(FIG, 'keys.png'), 'Backup key hierarchy', 380),
  H3('Other protections'),
  B('**App lock.** The app opens behind the phone’s own lock (fingerprint, face or PIN). A phone without a screen lock is asked to set one and the records stay closed. The app re-locks after one minute away.'),
  B('**Screen privacy.** Screenshots and the recent-apps preview are blocked by default.'),
  B('**No plaintext on disk.** Camera files, picked PDFs and share-sheet files are deleted after use and swept at start-up.'),
  B('**Network discipline.** One file in the app is allowed to use the network; a test fails if any other file becomes socket-capable.'),
  B('**Sender proof.** A share that carries a LifeTrack ID also proves it came from the phone holding that ID’s reply key, so a public ID lookup cannot be used to forge a share.'),
  B('**Not protected: the connection itself.** Over Wi-Fi the prototype uses plain HTTP. Records and backups are already ciphertext, but the backup login proof and ID tokens are not, so someone on the same Wi-Fi could read those in transit. A hosted server would need TLS (Section 8.2).'),
  B('**Server hardening.** Bound to 127.0.0.1 by default; DNS-rebinding and cross-origin pages refused; body size limits; per-caller rate limits; ID lookups throttled so the number space cannot be walked; enrolments awaiting review cannot be replaced; revocation follows keys.'),
  H2('4.7 LifeTrack ID'),
  P('Each profile gets a random 12-digit LifeTrack ID shown as `XXXX XXXX XXXX`. The last digit is a Verhoeff check digit, which catches every single-digit error and every swap of neighbouring digits. IDs always start with 1, so they can never be confused with an Aadhaar number. The server maps an ID only to a reply mailbox and key. It holds no name and no records. The token held on the phone moves or deletes the ID. If that token is lost, for example after a reinstall with no backup, a fresh identity check less than 15 minutes old, together with an explicit choice in the app, moves the ID to the new phone. The server accepts only a signed check newer than the last issue or move of that ID, so an old check cannot be replayed. This can take an ID from another phone, which is the price of not leaving a person locked out for good.'),
  H2('4.8 Entity-relationship model'),
  P('Figure 4.6 shows how the entities in Section 4.5 relate. Everything belongs to a Patient profile. A profile has many records, vital readings, grants and audit entries, and exactly one clinical profile and one Emergency ID. Custody links a grantor (a parent or custodian) to a dependent profile. Each grant has its own trail of audit entries, so every read made under a grant can be listed.'),
  ...figure(path.join(FIG, 'er.png'), 'Entity-relationship diagram', 560),
  P('Because the data lives on the phone as sealed JSON documents rather than in a relational database, the relationships are held as identifiers inside the documents. The audit log is append-only and is also sealed on disk.'),
  H2('4.9 Key algorithms'),
  H3('Verhoeff check digit'),
  P('The twelfth digit of a LifeTrack ID is computed with the Verhoeff algorithm [16], which uses the multiplication table of the dihedral group D5, a permutation table and an inverse table. Unlike a simple sum-of-digits check, it detects every single-digit error and every transposition of two adjacent digits, the two most common mistakes when an ID is read out or typed. The first digit is fixed at 1 so that no LifeTrack ID can be mistaken for, or validate as, an Aadhaar number.'),
  H3('Password derivation'),
  P('The backup password is stretched with Argon2id [13] using the OWASP minimum parameters [14]: 19 MiB of memory (19,456 KiB), 2 passes and 1 lane, producing a 32-byte root key. At this cost the derivation is slow enough to drop frames, so it runs on a background isolate. HKDF-SHA256 [12] then derives two independent 32-byte keys from the root using different labels: the vault key (lifetrack/v1/vault-key), which never leaves the phone, and the login verifier (lifetrack/v1/login-verifier), which is sent to the server so that it can release the locked backup. Test builds use weaker parameters, and that code path throws an error in a release build.'),
  H3('Shamir secret sharing'),
  P('The random data key that encrypts the backup is split with Shamir’s scheme [15] over the finite field GF(2^8) into three shares, any two of which rebuild it; one share alone reveals nothing. Each share carries its index, the threshold and an identifier of the split, and is printed as a recovery code of the form LT1-index-threshold-splitId-hex. Combining fewer shares than the threshold, or shares from different splits, is refused instead of producing a wrong key.'),
  H3('Insight rules'),
  P('Plane A compares the latest result for each test with a reference interval for the patient’s age, sex and pregnancy status, checks it against the danger limits of Section 5.3, and looks for movement between consecutive results. A change counts as a trend only when it is larger than a fixed threshold for that test (Table 4.2). Each finding has one of four severities: a note, discuss at the next visit, contact a doctor today, or urgent care.'),
  ...table(['Test', 'Unit', 'Minimum change', 'Worse direction'], [
    ['HbA1c', '%', '0.5', 'Up'], ['Fasting glucose', 'mg/dL', '15', 'Up'], ['LDL cholesterol', 'mg/dL', '20', 'Up'],
    ['HDL cholesterol', 'mg/dL', '8', 'Down'], ['Haemoglobin', 'g/dL', '1.0', 'Down'], ['Creatinine', 'mg/dL', '0.3', 'Up'],
  ], [3000, 1800, 2100, 2126], 'Trend thresholds used by Plane A'),
];

// ---------- Chapter 5 ----------
const ch5 = [
  chapter(5, 'Implementation'),
  H2('5.1 Technology stack'),
  ...table(['Layer', 'Technology', 'Reason'], [
    ['Patient app', 'Flutter (Dart), Kotlin for Android platform code', 'One code base, native performance, strong typing'],
    ['Doctor app', 'Flutter (web and mobile)', 'Shares the protocol package and design language with the patient app'],
    ['Server', 'Dart with shelf and shelf_router', 'Same language and protocol package as the apps; small and testable'],
    ['Cryptography', 'cryptography package: AES-256-GCM, X25519, Ed25519, HKDF, HMAC-SHA256, Argon2id', 'Maintained library, so primitives are not written here. Shamir splitting is the exception: `shamir.dart` was written for this project from the published scheme [15] and is unit-tested but not independently audited'],
    ['Key storage', 'Android Keystore via flutter_secure_storage', 'Device key kept out of app storage; hardware-backed on phones that support it'],
    ['OCR', 'Tesseract 4 (Tesseract4Android), English “best” model, checksum pinned', 'Fully offline; chosen over a cloud SDK that sends usage statistics'],
    ['QR', 'flutter_zxing (scan), qr_flutter (draw)', 'Offline, no network permission'],
    ['Interface', 'Flutter widgets with the project’s own component set; Sora typeface and Lucide icons bundled in the app', 'Consistent look on every screen; no fonts or images fetched from the internet'],
    ['Other', 'local_auth, flutter_local_notifications, archive (encrypted zip), printing / pdf, share_plus', 'App lock, reminders, export, prescriptions'],
  ], [1700, 3900, 3426], 'Technology stack'),
  H2('5.2 Code organisation and size'),
  ...table(['Part', 'Folder', 'Source files', 'Lines of code'], [
    ['Patient app', 'app/lib', String(STATS.code.app[0]), STATS.code.app[1]], ['Doctor app (LifeTrack Clinic)', 'doctor/lib', String(STATS.code.doctor[0]), STATS.code.doctor[1]],
    ['Server', 'server/lib, server/bin', String(STATS.code.server[0]), STATS.code.server[1]], ['Shared protocol', 'packages/lifetrack_protocol/lib', String(STATS.code.protocol[0]), STATS.code.protocol[1]],
    ['**Total application code**', '', `**${STATS.code.total[0]}**`, `**${STATS.code.total[1]}**`], ['Tests (all parts)', '*/test', String(STATS.code.tests[0]), STATS.code.tests[1]],
  ], [2900, 3000, 1400, 1726], `Code size by part (${STATS.date})`),
  P('The patient app is divided into modules under `app/lib/src/`: ai (planes A, B, C and the Insight contract), app (AppServices and wiring), audit, capture (OCR and image handling), consent, crypto, emergency, export, family, identity, model, reminders, security, share, storage, ui, vaccines and vault.'),
  H2('5.3 Patient app'),
  P('The app has five tabs (Reminders, Records, Insights, Share and Me) on a floating bar at the bottom, and a person button at the top right for switching between the people on the phone (Figures 5.1 and 5.2). The screenshots in this chapter show the synthetic sample family that demonstration builds can load. They were taken from version 1.42, so the screens added after the audit, such as the urgent prompt and allergy entry, are described but not shown.'),
  ...shotRow([path.join(SHOTS, '1-reminders.png'), path.join(SHOTS, '2-records.png'), path.join(SHOTS, '3-insights.png')], 'Patient app: Reminders, Records and Insights', 150),
  B('**Reminders** opens with a dark card showing today’s progress. Below it, today’s medicines from active prescriptions are grouped by morning, afternoon and night, with the current slot marked “Now”. A medicine is ticked by swiping it sideways. The patient’s own reminders (a check-up, a refill, a home reading) follow.'),
  B('**Records** has a search box, filters for prescriptions, reports and scans, and a “results over time” card that plots one test across reports. Smaller cards show the latest value of every other test with a sparkline. Tapping a test opens every result with its laboratory, sample date and the lab’s printed range. A floating button adds a record.'),
  B('**Insights** first explains what the on-device insights do and asks the patient to turn them on for 24 hours. Once on, it counts the findings to discuss with a doctor and the values that could not be checked, then lists each finding with its severity and the lab’s printed range. Every read the insights make appears in the access log.'),
  ...shotRow([path.join(SHOTS, '4-share.png'), path.join(SHOTS, 'scan-doctor.png'), path.join(SHOTS, '5-me.png')], 'Patient app: Share, scanning a doctor’s code, and Me', 150),
  B('**Share** offers “Show on my phone” (nothing is sent, and leaving the summary needs the phone’s lock), sharing with a doctor by scanning the code on the doctor’s screen, “Tests you already have” to show a doctor before they order a repeat, and “Who has access now”, where any share can be shortened, extended or ended.'),
  B('**Me** starts with the person’s own card and LifeTrack ID, then holds the Emergency ID, ABHA (optional, “not verified”), reminders, backup, the access log, data rights (export and erasure), app lock and screenshot privacy, appearance and the family.'),
  H3('Scanning a report (Plane B)'),
  P('A page arrives from the camera, the photo library, a PDF or another app’s share sheet. The image is straightened, resized and stripped of metadata, then passed to Tesseract [18] in the Android activity. A parser recognises test names, values, units and printed ranges. Any field read with confidence below 0.85, and every handwritten field, must be ticked or corrected before the patient confirms that the draft matches the report. The original image is kept, encrypted, with the record. Figure 5.3 shows the three steps: choosing how to add a record, adding the pages, and the saved record with its original and the values as reported.'),
  ...shotRow([path.join(SHOTS, 'add-record.png'), path.join(SHOTS, 'scan.png'), path.join(SHOTS, 'record.png')], 'Adding a report: choose, add pages, and the saved record', 150),
  H3('Insights (Plane A)'),
  P('The engine takes records already obtained through RecordAccess and a clock, and returns Insight objects. A value is compared with a reference interval only when the units match and the interval applies to the patient’s age, sex and pregnancy status; otherwise it is reported as “not checked”, never as normal. Critical values change only the severity and suggested action (discuss soon, contact a doctor today, urgent care), never the wording of the finding. Every output is screened against a list of banned phrases covering diagnoses (“you have”) and forecasts (“risk of”, “may indicate”); a banned output is dropped, not rephrased.'),
  H3('Danger-level check'),
  P('Insights are optional, so they cannot be the only safeguard. Since version 1.43.0 a separate check runs on every lab value as it is saved, with insights on or off. It covers potassium, sodium, calcium, platelets, white cells, haemoglobin, creatinine, HbA1c and every kind of glucose, normalises the units first, and treats a value it cannot judge as “not checked”. Home readings are checked the same way: sugar below 54 mg/dL or blood pressure at or above 180/120 asks for care now. When a value crosses a limit, a prompt with a Call 108 button replaces the usual “saved” message. A result more than a week old asks the patient to contact a doctor today instead. Pure code in the patient’s own save path does this, so it reads no store and needs no grant.'),
  P('This check has one exception to the banned-phrase rule. An alert quotes the unit as printed on the report, so an odd unit could contain a banned phrase and drop the alert. An urgent alert is too important to lose, so here the app shows fixed wording with nothing taken from the report; the banned text is still never displayed. All the limits sit in one file, `critical_values.dart`, marked as needing clinical review. They are commonly cited figures, and published surveys [27] show that such limits vary between institutions. No doctor has approved them.'),
  H3('Doses and allergies'),
  P('The app never fills in a dose. A scanned medicine with no printed dose has none until a person says how it is taken, and weekly, part-tablet, as-needed and insulin doses are flagged and kept as written. Methotrexate and warfarin always wait for a person. An allergy list is entered under Me and feeds both the Emergency ID and clinical shares. An empty list is shown as “not recorded”, because silence must not read as “no allergies”.'),
  H3('Family'),
  P('The phone owner is the head of the family. Each other profile has a role describing who the person is to the head. The family overview opens with a drawn family tree, one row per generation, with couples joined and the head marked (Figure 5.4). Adding a person starts with a choice: a child under 18, an adult who wants the head to manage their records, or an adult with their own phone, who is sent an invite instead. A child under 18 can be added only as the head’s son or daughter, with consent declared by the parent; sharing for that child stops at 18. An adult may delegate to a custodian in person, and can end it. Adults with their own phones can be linked. Each may share their Emergency ID and/or current medicine list with the other (both are off by default), sealed to the other phone’s key via the relay, with a safety code shown on both phones.'),
  H3('Emergency ID'),
  P('The Emergency ID holds a fixed set of bounded fields (blood type, conditions, medicines, emergency contact, last four digits of ABHA). It is sealed like every other blob but can be opened from the app’s lock screen, and every such read is logged as an unauthenticated lock-screen read. On screen it is a dark card; tapping the card turns it over to its QR code (Figure 5.5). The QR code is plain text so any camera can read it without the internet.'),
  ...shotRow([path.join(SHOTS, 'family-tree.png'), path.join(SHOTS, 'add-person.png')], 'Family overview with the family tree, and adding a person', 150),
  ...shotRow([path.join(SHOTS, 'emergency.png'), path.join(SHOTS, 'emergency-qr.png')], 'Emergency ID card and its plain-text QR code', 150),
  H2('5.4 User-interface design'),
  P('The interface was built in three rounds during weeks 6 and 7. Two screen-by-screen reviews of the patient app looked at layout, wording, touch targets and contrast, with older users and caregivers in mind, and each was followed by a series of small releases (patient 1.19.1 to 1.37.3). The app was then restyled (1.38.0 to 1.40.1) and, in week 7, tested for fit on small screens and made faster to draw.'),
  H3('Visual language'),
  P('Pages are near-white with white cards. The one thing that matters most on a screen, such as today’s progress, the Emergency ID or the person’s own card, sits on a near-black card. Text is set in Sora, which is bundled with the app, as are the icons, so nothing is fetched from the internet. The welcome and lock screens open on a full-width photograph of a family (Figure 5.6). The photographs, and the sample family’s faces cropped from them, were generated with an AI image tool (Canva) for the project and are bundled with the app. Text is never placed over a photograph, because the contrast test cannot measure text on an image.'),
  P('Colour carries a fixed meaning. The caramel accent (#A65E30) means only “selected” or “now”: the active tab, the chosen filter, the current medicine slot. Health status never uses it. Out-of-range values use purple and trends use teal, and the status badges also carry a warning icon, so colour is never the only signal. The first version of the restyle used an orange for out-of-range values; measured with a simulation of colour-vision deficiency [22], it was only 12 to 16 ΔE from the caramel accent for red- and green-blind viewers, so a selected chip could be mistaken for a warning. Purple and teal replaced it, and a test now keeps every pair of these colours at least 25 ΔE apart for normal, protan, deutan and tritan vision, in light and dark mode.'),
  ...shotRow([path.join(SHOTS, 'welcome.png'), path.join(SHOTS, 'lock.png'), path.join(SHOTS, 'first-profile.png')], 'Welcome screen, lock screen and the first step of creating a profile', 150),
  H3('Interaction'),
  P('Every tab is labelled, and the selected one sits in a caramel pill that slides along the bar. The top and bottom bars are frosted glass, so the list can be seen scrolling beneath them. A medicine is ticked by swiping it sideways, Share refreshes when pulled down, and saving a record, a reading or a vaccination ends with a short tick animation. Buttons and cards respond to touch with a soft glow and a light vibration. Dangerous actions, such as erasing data or leaving the summary shown on the phone, ask for the phone’s own lock first.'),
  H3('Dark mode'),
  P('Every colour in the app is defined for light and dark mode, and the patient picks one or follows the phone under Me → Appearance (Figure 5.7). The contrast test walks every visible text on the main screens in both modes and checks it against WCAG AA. QR codes stay black on white in both modes so that any camera can read them.'),
  ...shotRow([path.join(SHOTS, 'dark-reminders.png'), path.join(SHOTS, 'dark-records.png')], 'Dark mode: Reminders and Records', 150),
  H3('Fit and speed'),
  P('Layouts break when a phone is set to a large text size or display size, and those are the settings older users are most likely to change. A layout test therefore lays out every main tab and 18 other screens at widths of 320, 360 and 411 logical pixels, each with text at 1×, 1.3× and 2×, and fails on any text that runs out of its box or any tab name that leaves its pill. It found overflows in tab names, vaccine dates, prescription rows, the Emergency ID’s blood type, dose pills, month headers and test results, which versions 1.42.1 and 1.42.2 fixed. Secondary text is at least 14 sp and no text is smaller than 12 sp.'),
  P('Version 1.42.2 reduced the work the phone does for every frame without changing the look, which I confirmed by comparing screens pixel by pixel. The glass bars share one background blur, cards and buttons draw their shadows, borders and touch ripple as rounded rectangles (the renderer’s fast path), list items fade in only when a page opens, a tab switch draws only the incoming tab, and photographs are decoded at the size they are shown. A frame-timing test (`integration_test/scroll_perf_test.dart`) scrolls the main tabs on a real phone in a profile build to measure the result.'),
  H2('5.5 Doctor app: LifeTrack Clinic'),
  P('On first launch the doctor enrols with their name, registration number, council and specialty; the device generates its keys and the administrator approves the enrolment. The approved doctor’s home screen shows their verified code for patients to scan, a form to request access by LifeTrack ID, and the shares that have arrived, each with a countdown (Figure 5.8).'),
  ...figure(path.join(DEMO, 'doctor-05-share-arrived.png'), 'LifeTrack Clinic: verified code, request by ID and an arrived share', 560),
  P('Opening a share shows the patient’s allergies prominently, current and past medicines with their status, conditions and the latest tests with their source and age, and a link to the original report image. Beside it, the doctor writes a prescription with the medicine, active ingredient, morning-afternoon-night dose, timing relative to food, duration and a note, then signs and sends it (Figure 5.9). The dose, days and timing start empty, so a prescription cannot go out on a default. The prescription can also be printed with the registration number and a signature fingerprint. The doctor app keeps its own notes and a patient register, sealed in the browser.'),
  ...figure(path.join(DEMO, 'doctor-07-prescription.png'), 'LifeTrack Clinic: a shared summary and a prescription being written', 560),
  P('The app has three tabs: Patients, Prescriptions and Profile. The Profile tab (Figure 5.10) shows the licence details from the approved enrolment, which the doctor cannot edit, and when the certificate next renews. It also holds the practice details printed on paper prescriptions, the privacy settings (the app lock is always on, and hiding from screenshots is a switch), the appearance setting and a count of what this device keeps.'),
  ...figure(path.join(DEMO, 'doctor-09-profile.png'), 'LifeTrack Clinic: the Profile tab', 560),
  H2('5.6 Server'),
  P('The server is a single Dart process using shelf. Its routes are grouped under `/v1/`:'),
  ...table(['Route group', 'Purpose'], [
    ['/v1/health, /v1/registry/key', 'Health check; registry public key'],
    ['/v1/doctors, /v1/admin/doctors/…', 'Enrolment, status lookup; administrator approve, reject, revoke (admin token required)'],
    ['/v1/relay/<mailbox>', 'Deposit and signed pickup of sealed messages; deleted after pickup or at expiry; never beyond 48 hours'],
    ['/v1/ids, /v1/ids/reclaim, /v1/ids/<id>', 'Register, move, delete and look up LifeTrack IDs (rate-limited); reclaim one after a fresh identity check'],
    ['/v1/vault/<account>/…', 'Salt, open, replace and delete backups with a password proof; ten wrong proofs per hour'],
    ['/v1/verify/…', 'Identity check: start, callback, result released only with the hand-back code'],
  ], [3200, 5826], 'Server routes'),
  P('By default the server binds only to 127.0.0.1, and the phone reaches it over USB with `adb reverse tcp:8787 tcp:8787`. Since week 6 it can also listen on the laptop’s Wi-Fi address (`scripts/run_server_lan.sh`), so the phone works without a cable. The release build then has that address built in and allows plain HTTP to that one address only; records and backups cross the connection only as ciphertext, though the backup login proof and ID tokens do not (Section 4.6). Data is kept in an owner-only folder: the registry signing key, the admin token, doctor enrolments, the ID map and the backup files.'),
  H2('5.7 Development practices'),
  B('**Version control.** Git with a GitHub repository, in use since the rebuild of 29 September 2026 (Section 1.7). [[Anish to state whether commit 8f7a9d2, authored as Phanishwar Lal, was someone else’s work]]'),
  B(`**Versioning rule.** Every update raises the version (a feature raises the middle number, a fix the last) and adds a line to the changelog in the same commit. At the time of writing the patient app is at ${STATS.patient} and the doctor app at ${STATS.doctor} (Appendix A).`),
  B('**Interface reviews.** Two screen-by-screen reviews of the patient app, each followed by small releases that fixed what it found, then a restyle (Section 5.4).'),
  B('**Security reviews.** Two dedicated hardening releases (patient 1.13.0 and 1.18.1/1.18.2, doctor 1.4.1/1.4.2) fixed issues found in review, such as binding the identity result to a hand-back code and proving the sender of a share. A third, full review of the server, the consent and sharing code and the backup vault on 2 November 2026 fixed a race in which two registrations arriving at once could give one person two LifeTrack IDs, and capped what a single sender may leave waiting on the relay, so that one device cannot fill it for everyone else. [[Anish to state who carried out the interface and security reviews]]'),
  B('**Audit and fixes.** The audit of 6 November 2026 found problems the tests had not, and they were fixed in the releases of Appendix A: a missing dose defaulted to 1-0-1, allergies had no entry screen, the danger check ran only with insights on, and ending access crashed with the server off.'),
  B('**Release build.** A script runs the analysis and tests, then builds an obfuscated APK; the symbol map is kept on the development machine. `scripts/check_all.sh` runs every package’s analysis and tests. There is no hosted continuous integration, so these checks run only when started on the development machine.'),
  H2('5.8 Selected implementation details'),
  P('This section shows four short excerpts from the patient app’s source code. They are the places where the main rules of Chapter 4 are enforced. Listings are shortened only where marked with “...”.'),
  H3('Consent check'),
  P('Every read by anyone other than the patient ends in ConsentStore.authorize (Listing 5.1). The caller passes only the grant id; the method looks up the stored grant and checks every field, in a fixed order, against the request. Each failure throws a typed denial, which RecordAccess writes to the audit log as a refusal. A one-time grant, such as the one Plane C would use, is marked spent on its first use.'),
  ...listing(["ConsentToken authorize({required String tokenId, required String actorId,", " required String patientId, required DataClass dataClass,", " required GrantPurpose purpose, required DateTime at}) {", "  final t = _issued[tokenId];", "  if (t == null) throw ConsentDenied(DenialReason.unknown, tokenId);", "  if (_revoked.contains(tokenId)) throw ConsentDenied(DenialReason.revoked, tokenId);", "  if (_spent.contains(tokenId)) throw ConsentDenied(DenialReason.spent, tokenId);", "  if (at.isBefore(t.issuedAt)) throw ConsentDenied(DenialReason.notYetValid, tokenId);", "  if (!at.isBefore(t.expiresAt)) throw ConsentDenied(DenialReason.expired, tokenId);", "  if (t.patientId != patientId) throw ConsentDenied(DenialReason.wrongPatient, tokenId);", "  if (t.grantedToActorId != actorId) throw ConsentDenied(DenialReason.wrongActor, tokenId);", "  if (t.purpose != purpose) throw ConsentDenied(DenialReason.wrongPurpose, tokenId);", "  if (!t.covers(dataClass)) throw ConsentDenied(DenialReason.outOfScope, tokenId);", "  if (t.type == GrantType.oneTime) { ... }", "  return t;", "}"], 'ConsentStore.authorize (lib/src/consent/consent_store.dart)'),
  H3('Sealed storage'),
  P('All records, grants and audit entries pass through SealedStore (Listing 5.2). Each blob is encrypted under the device key and bound to its own name, so a blob copied to another name fails to open. A blob in the older version-1 frame is upgraded to version 2 the first time it is read.'),
  ...listing(["Future<Object?> read(String name) async {", "  final frame = await bytes.read(name);", "  if (frame == null) return null;", "  final body = await VaultCodec.open(frame, key, name: name);", "  if (VaultCodec.versionOf(frame) == 1) await write(name, body);", "  return body;", "}", "", "Future<void> write(String name, Object? body) async =>", "    bytes.write(name, await VaultCodec.seal(body, key, name: name));"], 'SealedStore read and write (lib/src/storage/sealed_store.dart)'),
  H3('Audit before egress'),
  P('Listing 5.3 is the end of the clinical-share flow. The bundle is sealed to the doctor’s key first, then the audit entry is written, and only then is the envelope handed to the relay. If the network call fails, the attempt is still on record and the grant is ended, so a failed share leaves no live access.'),
  ...listing(["final envelope = await ShareEnvelope.seal(bundle: bundle,", "    recipientKey: doctor.encryptionKey, keyId: keyId,", "    grantId: grant.id, expiresAt: grant.expiresAt);", "// Audited before it leaves: a failed send is still recorded.", "await audit.append(AuditEntry(at: at, actorId: grant.grantorId,", "    patientId: patientId, action: AuditAction.clinicalShare, tokenId: grant.id,", "    detail: 'Sent to ${doctor.name} (${doctor.registrationNo}), encrypted to their device'));", "...", "try {", "  await relay.deposit(mailbox: mailbox, grantId: grant.id,", "      envelope: envelope, expiresAt: grant.expiresAt);", "} on Object {", "  await _end(grant, at, 'share not delivered');", "  rethrow;", "}"], 'Writing the audit entry before the share is sent (lib/src/share/clinical_share_service.dart)'),
  H3('Banned-phrase filter'),
  P('Every Plane A output is checked against a list of banned phrases covering diagnoses, forecasts and treatment advice (Listing 5.4). An insight that contains one is dropped, not reworded, so the filter can only remove output and never invent new wording. The danger-level alert of Section 5.3 is the single exception.'),
  ...listing(["const bannedPhrases = <String>[", "  // assertions", "  'diagnos', 'you have', 'you are', 'disease', 'is normal', 'healthy', ...", "  // forecasts and hedges", "  'risk of', 'likely', 'may indicate', 'suggests', 'progression to', 'predict', ...", "  // treatment advice", "  'stop taking', 'increase your dose', 'reduce your dose', 'skip',", "];", "", "String? firstBannedPhrase(String text) {", "  final t = text.toLowerCase();", "  for (final p in bannedPhrases) {", "    if (t.contains(p)) return p;", "  }", "  return null;", "}"], 'Banned-phrase list and check, shortened (lib/src/ai/plane_a/banned_phrases.dart)'),
  H3('Vaccination schedule'),
  P('The child vaccination schedule follows India’s Universal Immunisation Programme as published by the Ministry of Health and Family Welfare [7]: 26 doses from birth (BCG, OPV 0, hepatitis B) to 16 years. Each dose has an age at which it falls due, and the app computes due dates from the child’s date of birth. The screen advises following the child’s doctor where a private schedule adds vaccines.'),
];

// ---------- Chapter 6 ----------
const ch6 = [
  chapter(6, 'Testing and Results'),
  H2('6.1 Testing strategy'),
  P('Testing followed one principle from the project proposal: a claim is accepted only if a test can prove it. Seven kinds of tests were developed:'),
  B('**Unit tests** for pure logic: reference ranges, trend thresholds, Verhoeff digits, Shamir splitting, frame encoding, grant checks.'),
  B('**Widget tests** that drive real screens: adding records, sharing, back-button behaviour, dark mode.'),
  B('**Invariant tests** that read the source code itself and fail if a rule is broken. For example, the test run fails if any file other than the one network file can open a socket, if the release manifest gains the INTERNET permission without review, if code under `lib/src/ai/` calls `ConsentStore.issue` or `DateTime.now()`, or if any UI string claims that LifeTrack or the server “cannot read” the patient’s data. These tests strip comments first, so a comment cannot hide or fake a violation.'),
  B('**Contrast tests** that walk every visible text on every main screen, in light and dark mode, to the solid colour behind it and check WCAG AA contrast.'),
  B('**Layout tests** that lay out every main tab and 18 other screens at 320, 360 and 411 logical pixels wide with text at 1×, 1.3× and 2×, and fail on any overflow or any tab name outside its pill.'),
  B('**Colour-vision tests** that simulate protan, deutan and tritan vision [22] and fail if the accent colour comes within 25 ΔE of any health-status colour, in light or dark mode.'),
  B('**Integration tests** that run the real code of both apps against the real server handler: scan, share, confirmed sender, prescription back, request by ID, decline, time change and revocation.'),
  H2('6.2 Results'),
  P(`All tests were run on ${STATS.date} against the current version (patient app ${STATS.patient}, doctor app ${STATS.doctor}) with \`scripts/check_all.sh\`. Static analysis (\`flutter analyze\` / \`dart analyze\`) reported no issues in any package. The figures show that the code matches what the tests check; they do not show that the tests cover everything, as the audit in Section 5.7 found.`),
  ...table(['Part', 'Static analysis', 'Tests passed', 'Failed'], [
    ['Patient app (app/)', 'No issues', `${STATS.tests.app} (${STATS.tests.skipped} skipped)`, '0'], ['Doctor app (doctor/)', 'No issues', String(STATS.tests.doctor), '0'],
    ['Server (server/)', 'No issues', String(STATS.tests.server), '0'], ['Shared protocol', 'No issues', String(STATS.tests.protocol), '0'],
    ['Both apps + server (integration/)', 'No issues', String(STATS.tests.integration), '0'], ['**Total**', '', `**${STATS.tests.total}**`, '**0**'],
  ], [3600, 2000, 1800, 1626], 'Test results by part'),
  P('The patient app also has tools that render screens to images for visual review; they run only when asked to save screenshots and are not counted. The patient app’s total rose from 285 on 5 October to 318 by 17 October, as the interface work of weeks 6 and 7 added features and tests, to 319 on 2 November with the test for the bug the frame-timing run found (Section 6.4), and to ' + STATS.tests.app + ' on 6 November with the audit fixes.'),
  H2('6.3 Selected test cases'),
  ...table(['#', 'Test case', 'Expected result', 'Result'], [
    ['1', 'A pending (unapproved) doctor tries to show a code', 'No scannable code is produced', 'Pass'],
    ['2', 'Patient scans, shares; doctor opens; doctor prescribes; patient collects', 'Share opens only with the doctor’s key; prescription verified as signed and filed as doctor-issued', 'Pass'],
    ['3', 'Doctor requests by LifeTrack ID; patient approves once, declines once', 'Approval produces a share; decline is logged under Refused and the doctor is told', 'Pass'],
    ['4', 'A doctor is revoked after enrolment', 'Patient app refuses the doctor; doctor app stops showing as approved', 'Pass'],
    ['5', 'Doctor asks for 12 h; patient gives 4 h, later changes it', 'Doctor sees each new end time; old copy replaced', 'Pass'],
    ['6', 'Share made without choosing a time', 'Access lasts exactly 1 hour', 'Pass'],
    ['7', 'A doctor tries to pick up another doctor’s mailbox', 'Refused', 'Pass'],
    ['8', 'A signed pickup is replayed', 'Refused', 'Pass'],
    ['9', 'Ten wrong backup passwords', 'Account locked for an hour; server keeps only a hash', 'Pass'],
    ['10', 'An ID is mistyped by one digit or two neighbours swapped', 'Rejected by the Verhoeff check before any lookup', 'Pass'],
    ['11', 'A sealed blob is copied under another name', 'Fails to open (name bound as associated data)', 'Pass'],
    ['12', 'A unit on a report does not match the reference range’s unit', 'Shown as “not checked”, never as out of range or normal', 'Pass'],
    ['13', 'An insight’s wording contains a banned diagnostic or forecast phrase', 'Insight dropped', 'Pass'],
    ['14', 'A DigiLocker sign-in link is forwarded to someone else', 'Result released only with the hand-back code on the finishing page; five wrong codes end the check', 'Pass'],
    ['15', 'Any visible text in light or dark mode', 'Meets WCAG AA contrast', 'Pass'],
    ['16', 'Every main tab and 18 screens at 320 wide with text at 2×', 'No text runs out of its box; every tab name stays inside its pill', 'Pass'],
    ['17', 'Accent and health-status colours under simulated protan, deutan and tritan vision', 'Every checked pair at least 25 ΔE apart, in light and dark mode', 'Pass'],
    ['18', 'Potassium 6.5 is saved with insights off', 'An urgent prompt replaces “saved”', 'Pass'],
    ['19', 'A scanned medicine has no printed dose', 'No dose is filled in; a person must state it', 'Pass'],
    ['20', 'The patient ends access with the server off', 'Access ends on the phone and is logged; the doctor is told on a later retry; the screen says they have not been told yet', 'Pass'],
    ['21', 'The owner erases their profile while the server cannot delete the backup', 'The log and the screen say the backup stayed on the server', 'Pass'],
    ['22', 'A person lost their ID token and tries to reclaim the ID', 'A fresh signed identity check moves it; an old or replayed check is refused', 'Pass'],
  ], [500, 3300, 4000, 1226], 'Selected test cases'),
  H2('6.4 Manual testing on a device'),
  P('Device results are recorded in `docs/testing/results/`. So far that record holds one phone, a Samsung Galaxy S25, on 2 November 2026. Hand tests on the phone during weeks 5 to 7 were not logged, so they are not counted as results here. OCR has been tested mainly on synthetic report images; accuracy on real photographed reports has not been measured.'),
  P('On 2 November 2026 the server moved from the Mac to a Windows PC and started with a fresh registry key and a new release signing key. The patient app was rebuilt, and a sample doctor enrolled again from LifeTrack Clinic 1.8.0 and was approved. Frame timing was then measured on the same phone with an integration test that sets up the ten-person sample family by itself, scrolls each of the five tabs and switches between them while the phone records how long every frame takes to build and to draw (Table 6.3). At 60 frames a second each frame has 16.7 ms.'),
  ...table(['Action', 'Frames', 'Build: average / 90th percentile (ms)', 'Draw: average / 90th percentile (ms)', 'Frames over budget'], [
    ['Scrolling all five tabs', '1,265', '1.2 / 2.1', '3.3 / 5.8', '19 (1.5%)'],
    ['Switching tabs, two rounds', '82', '3.7 / 7.3', '10.7 / 21.8', '14 (17%)'],
  ], [2300, 900, 2100, 2100, 1626], 'Frame timing on a Samsung Galaxy S25 (patient 1.42.3, profile build)'),
  P('Scrolling stays within the frame budget for 98.5% of frames. Switching tabs is slower to draw: about one frame in six is late, on the drawing (GPU) side rather than in building the widgets, most likely when a tab’s blurred bars and photos are drawn for the first time. This is a polish item rather than a defect. The run also found a bug in the patient app, fixed with a test: a finger lifted after its card had left the screen raised an error (1.42.3). A doctor-app bug fixed the same day made an enrolment wait for approval forever when the server no longer knew it (1.8.1).'),
  P('Three manual checks from the field test kit are recorded, and all three passed: the app installs and opens behind the phone’s lock, a share from the patient app reaches the doctor app, and screenshots are blocked on a normal build. The other checks in the kit, and timing on more phones, have not been done. **No change from version 1.43.0 onwards has been run on a real phone.** That includes the urgent prompt and its Call 108 button, background backup, the “Move my ID” dialog and the test step in the release script.'),
  H2('6.5 Example of an invariant test'),
  P('Listing 6.1 is taken from the invariant test that guards the AI planes. It loads the source of every file under lib/src/ai, strips the comments (codeOf), and fails if any of them calls ConsentStore.issue, imports the grant code, or reads the wall-clock time. If a future change gave an agent a way to grant itself access, this test would fail, and the release script stops on a failed test.'),
  ...listing(["group('agents (lib/src/ai) never', () {", "  final ai = [for (final f in lib) if (rel(f).startsWith('lib/src/ai/')) f];", "", "  test('issue a grant (CLAUDE.md #5)', () {", "    for (final f in ai) {", "      final code = codeOf(f);", "      expect(code.contains('.issue('), isFalse, reason: rel(f));", "      expect(code.contains('grants.dart'), isFalse, reason: rel(f));", "    }", "  });", "", "  test('read wall-clock time - the clock is always a parameter', () {", "    for (final f in ai) {", "      expect(codeOf(f).contains('DateTime.now()'), isFalse, reason: rel(f));", "    }", "  });", "});"], 'Part of plane_boundary_test.dart (app/test/invariants)'),
];

// ---------- Chapter 7 ----------
const ch7 = [
  chapter(7, 'Results and Discussion'),
  H2('7.1 Objectives achieved'),
  ...table(['Objective', 'Outcome'], [
    ['1. Patient-owned, encrypted record for the family', 'Achieved. All data sealed with AES-256-GCM under a Keystore-protected key; family with head-of-family roles; allergies and conditions entered since 1.45.0'],
    ['2. Explicit, time-limited, revocable clinician access', 'Achieved. Four levels; 1 h default up to 2 days; change or end at any time; every share audited before it leaves'],
    ['3. No server can read shares or backups', 'Achieved by design and tested: shares sealed to the doctor’s key; backups to a key from an app-generated password. The prototype’s plain-HTTP connection is a limit (Section 8.2)'],
    ['4. Safe on-device insights', 'Achieved as a prototype for ranges, movement and duplicate medicines, with non-diagnostic wording enforced by tests. No doctor has reviewed the rules or the danger limits'],
    ['5. On-device reading of reports', 'Partly achieved. Works on Android with Tesseract, English only, on synthetic images; accuracy on real photos has not been measured'],
    ['6. Verified doctor app with signed prescriptions', 'Achieved: registry certificates, status checks, signed and printable prescriptions'],
    ['7. Verifiable privacy promises', `Achieved for the rules a test can check: invariant tests, run by the release script; ${STATS.tests.total} tests passing`],
    ['8. Usable, accessible interface', 'Partly achieved. Tests show WCAG AA contrast in both modes, every main screen at three widths and three text sizes, and status colours checked for colour-blind viewers. Usability with older users has not been tested'],
  ], [3800, 5226], 'Objectives and outcomes'),
  H2('7.2 Changes from the original proposal'),
  P('The proposal planned a mobile app, a separate web analytics dashboard and a clinician portal, with a FHIR R4 data model and a synthetic-dataset approach to verifying analytics. Three changes were made to the proposal during development:'),
  B('**Analytics moved onto the phone.** A web dashboard would need a server that can read records, which contradicts the project’s central promise. The analytics were therefore built into the patient app as on-device insights.'),
  B('**The clinician portal became LifeTrack Clinic**, a full doctor app with verified identity, request by ID and signed prescriptions, because safe sharing turned out to be the most useful feature for reducing repeat tests.'),
  B('**FHIR alignment was reduced to provenance.** Every record keeps its source and provenance so that a later FHIR export is possible, but FHIR resources and export were not implemented in this phase.'),
  P('The proposal’s “ground-truth” principle was retained and extended. Instead of one synthetic dataset, every rule the system promises is written as a test with a known expected result.'),
  H2('7.3 Discussion'),
  P('The principal lesson of the project is that privacy must be part of the structure of the code from the outset; it cannot be added later as a feature. Making ConsentStore the only grant authority and RecordAccess the only non-patient read path meant that sharing, insights and the cloud plane all inherited the same checks and the same audit trail. Expressing the rules as source-scanning tests turned code review questions (“does anything else open a socket?”) into automatic answers.'),
  P('A second lesson is that accurate wording is itself a requirement. Early screens said the doctor “keeps no copy” and that the server “cannot read” the data. Both claims were removed: a doctor is legally required to keep their consultation record, and an offline password-guessing attack on a human-chosen password would make “cannot read” untrue. The app now generates the password and the UI makes only claims the code can back, enforced by a test.'),
  P('Finally, the work on provenance showed that a doctor will reuse a result instead of ordering it again only if they trust where it came from. Access control alone is not enough; the doctor needs the laboratory, the sample date, the printed range and the original image. These are shown on every result.'),
  P('The interface work taught a similar lesson. Colour on a health screen is a claim: when the restyle first used an orange for out-of-range values close to the caramel used for “selected”, a colour-blind patient could have read a selected filter as a warning. Giving each colour one meaning fixed it, and a test that measures the distance between the colours keeps it fixed. Text that ran out of its box at large text sizes was found the same way, by a test that lays out every screen at the settings older users choose, not by looking at the app on one phone.'),
  P('The audit of week 10 taught the limit of that approach. All 400 tests were passing while a missing dose became 1-0-1 and allergies had no entry screen. The tests were correct, but nobody had written a test for those cases. A passing test run shows that the written rules hold, not that the right rules were written, and that is why review by a doctor and trials with real users are still needed.'),
  H2('7.4 Comparison with existing solutions'),
  P('Table 7.2 compares LifeTrack Health with the two most common existing categories from Chapter 2. The entries for existing categories describe typical systems in each category, not any single product.'),
  ...table(['Feature', 'Hospital EMR', 'Document locker', 'LifeTrack Health'], [["Patient holds the complete record", "No", "Documents only", "Yes"], ["Works offline", "Inside the hospital", "No", "Yes"], ["Results compared with ranges", "Within one hospital", "No", "Yes, on the device"], ["Time-limited sharing with one doctor", "No", "Link or consent artefact", "Yes, 1 hour to 2 days"], ["Doctor verified before sharing", "Staff only", "Varies", "Yes, certificate"], ["Server can read shared data", "Yes", "Yes", "No (encrypted)"], ["Family profiles, proxy consent", "No", "Limited", "Yes"], ["Privacy rules tested automatically", "Not published", "Not published", "Yes (invariant tests)"]], [3500, 1700, 1700, 2126], 'Comparison with existing categories'),
];

// ---------- Chapter 8 ----------
const ch8 = [
  chapter(8, 'Conclusion and Future Scope'),
  H2('8.1 Conclusion'),
  P(`LifeTrack Health shows, as a working prototype, that a health record can be kept on the patient’s own phone and that the right part of it can go to a verified doctor for a limited time, through a server that holds no key to open it. The prototype covers the full loop (capture, understanding, sharing, prescribing back, reminders, emergencies, family, backup and export) across a patient app, a doctor app and a server, and its main privacy and safety rules are checked by automated tests. The interface was reviewed, restyled and tested for contrast, colour vision and fit on small screens with large text. All ${STATS.tests.total} tests pass and static analysis is clean. Whether real patients and doctors can use it safely is not yet shown: that needs the clinical review, phone tests and usability check listed below.`),
  H2('8.2 Limitations'),
  B('Only synthetic patients and doctors have been used; no real patient data has been processed.'),
  B('The identity check uses a practice stand-in for DigiLocker until partner registration; the optional mode that requires Aadhaar-based verification for every profile must be reviewed by legal counsel in light of *Puttaswamy* (2018) and the 2019 Aadhaar amendment before any pilot.'),
  B('ABHA numbers are typed in and shown as “not verified”; ABDM integration is not yet built.'),
  B('OCR is English-only, Android-only, and not yet measured on a large set of real photographed reports.'),
  B('No doctor has reviewed the app. The danger limits and dose rules are commonly cited figures that vary between institutions [27]; they need clinical review and must be replaced by a partner laboratory’s published list before a pilot.'),
  B('Device testing covers one phone and three recorded manual checks, and nothing from version 1.43.0 onwards has been run on a phone (Section 6.4).'),
  B('If a profile is erased while the server cannot be reached, the backup stays on the server. The patient is told, and can remove it by restoring and turning backup off. Turning backup off with the server away simply fails.'),
  B('A failed update to a linked family phone is retried but not shown on screen, and restoring an old backup after the ID has moved leaves that profile with no ID and no warning.'),
  B('The server has run only on a laptop or PC, reached over USB or the home Wi-Fi, and has not been deployed on the public internet. Over Wi-Fi the connection is plain HTTP. Records and backups are already encrypted, but someone on the same Wi-Fi could read the backup login proof and ID tokens, and could in theory interfere with an ID reclaim. A hosted server would need TLS.'),
  B('The app is in English only, and its usability has not been tested with older users.'),
  H2('8.3 Future scope'),
  N('ABDM / ABHA integration: verify ABHA, link records as a health information user, and map ABDM consent artefacts as a separate long-lived type.', 'num2'),
  N('DigiLocker partner registration to replace the practice stand-in.', 'num2'),
  N('FHIR R4 export of records with Provenance resources.', 'num2'),
  N('Hosted deployment of the server behind a reverse proxy, with push notifications instead of polling.', 'num2'),
  N('iOS support, including on-device OCR with Apple Vision.', 'num2'),
  N('Hindi and other Indian languages for the screens and for reading reports.', 'num2'),
  N('OCR accuracy study on real photographed reports from Indian laboratories.', 'num2'),
  N('Relative-contributed family history, with that relative’s own consent per request.', 'num2'),
  N('Plane C (Tier 2) cloud analysis, only with a certified service and the human-tap safeguards already designed.', 'num2'),
  N('Pilot with a partner clinic and laboratory after legal and clinical review.', 'num2'),
];

// ---------- References ----------
const refs = [
  'Government of India. *The Digital Personal Data Protection Act, 2023* (Act No. 22 of 2023). Ministry of Law and Justice, 2023.',
  'Government of India. *The Aadhaar (Targeted Delivery of Financial and Other Subsidies, Benefits and Services) Act, 2016*, Section 29; and the Aadhaar and Other Laws (Amendment) Act, 2019.',
  'Supreme Court of India. *Justice K.S. Puttaswamy (Retd.) v. Union of India*, (2019) 1 SCC 1 (Aadhaar judgment, 26 September 2018).',
  'Ministry of Electronics and Information Technology. *Information Technology (Reasonable Security Practices and Procedures and Sensitive Personal Data or Information) Rules, 2011*.',
  'Board of Governors, Medical Council of India. *Telemedicine Practice Guidelines*, March 2020.',
  'National Health Authority. *Ayushman Bharat Digital Mission (ABDM)*: ABHA and Health Information Exchange & Consent Manager specifications. https://abdm.gov.in',
  'Ministry of Health and Family Welfare. *Universal Immunization Programme*: National Immunization Schedule.',
  'HL7 International. *FHIR Release 4 (R4)*. https://hl7.org/fhir/R4/',
  'M. Dworkin. *Recommendation for Block Cipher Modes of Operation: Galois/Counter Mode (GCM) and GMAC*. NIST SP 800-38D, 2007.',
  'A. Langley, M. Hamburg, S. Turner. *Elliptic Curves for Security*. RFC 7748, IETF, 2016.',
  'S. Josefsson, I. Liusvaara. *Edwards-Curve Digital Signature Algorithm (EdDSA)*. RFC 8032, IETF, 2017.',
  'H. Krawczyk, P. Eronen. *HMAC-based Extract-and-Expand Key Derivation Function (HKDF)*. RFC 5869, IETF, 2010.',
  'A. Biryukov, D. Dinu, D. Khovratovich, S. Josefsson. *Argon2 Memory-Hard Function for Password Hashing and Proof-of-Work Applications*. RFC 9106, IETF, 2021.',
  'OWASP Foundation. *Password Storage Cheat Sheet* (Argon2id parameters). https://cheatsheetseries.owasp.org',
  'A. Shamir. “How to Share a Secret.” *Communications of the ACM*, 22(11), 612–613, 1979.',
  'J. Verhoeff. *Error Detecting Decimal Codes*. Mathematical Centre Tracts 29, Mathematisch Centrum, Amsterdam, 1969.',
  'M. Kleppmann, A. Wiggins, P. van Hardenberg, M. McGranaghan. “Local-First Software: You Own Your Data, in Spite of the Cloud.” *Proc. ACM Onward! 2019*.',
  'R. Smith. “An Overview of the Tesseract OCR Engine.” *Proc. ICDAR 2007*, IEEE.',
  'W3C. *Web Content Accessibility Guidelines (WCAG) 2.1*, 2018.',
  'Google. *Flutter documentation* and *Dart language documentation*. https://docs.flutter.dev, https://dart.dev',
  'IMDRF SaMD Working Group. *Software as a Medical Device: Possible Framework for Risk Categorization and Corresponding Considerations*, 2014.',
  'G. M. Machado, M. M. Oliveira, L. A. F. Fernandes. “A Physiologically-based Model for Simulation of Color Vision Deficiency.” *IEEE Transactions on Visualization and Computer Graphics*, 15(6), 1291–1298, 2009.',
  'Ministry of Electronics and Information Technology. *Digital Personal Data Protection Rules, 2025*, G.S.R. 846(E), 13 November 2025.',
  'P. C. Tang, J. S. Ash, D. W. Bates, J. M. Overhage, D. Z. Sands. “Personal Health Records: Definitions, Benefits, and Strategies for Overcoming Barriers to Adoption.” *Journal of the American Medical Informatics Association*, 13(2), 121–126, 2006.',
  'A. Roehrs, C. A. da Costa, R. da Rosa Righi, K. S. F. de Oliveira. “Personal Health Records: A Systematic Literature Review.” *Journal of Medical Internet Research*, 19(1), e13, 2017.',
  'A. Azaria, A. Ekblaw, T. Vieira, A. Lippman. “MedRec: Using Blockchain for Medical Data Access and Permission Management.” *Proc. 2nd International Conference on Open and Big Data*, IEEE, 25–30, 2016.',
  'G. J. Kost. “Critical Limits for Urgent Clinician Notification at US Medical Centers.” *JAMA*, 263(5), 704–707, 1990.',
];
function refPara(t, i) {
  const parts = t.split('*');
  const kids = [new TextRun(`[${i + 1}]  `)];
  parts.forEach((s, j) => { if (s) kids.push(new TextRun({ text: s, italics: j % 2 === 1 })); });
  return new Paragraph({ children: kids, spacing: { after: 120, line: 300 }, indent: { left: 567, hanging: 567 } });
}
const references = [H1('References'), ...refs.map(refPara)];

// ---------- Appendices ----------
chap = 'A';
figNo = 0; tabNo = 0;
const appA = [
  H1('Appendix A: Version History'),
  P('Selected releases from the project changelog. A feature raises the middle number, a fix the last.'),
  ...table(['Version', 'Date', 'Highlights'], [
    ['Patient 1.0.0', '29 Sep 2026', 'App scaffold, consent core, sealed storage, backup vault, Emergency ID'],
    ['Patient 1.1–1.6', '30 Sep – 1 Oct', 'Report capture with originals, family profiles, doctor sharing, edit and delete records'],
    ['Patient 1.7–1.12', '1 Oct', 'Medicine reminders, vaccinations, PDF import, “Tests you already have”, home readings'],
    ['Patient 1.13.0', '1 Oct', 'Security hardening'],
    ['Patient 1.14–1.16', '1 Oct', 'LifeTrack ID, request by ID, family roles, dark mode'],
    ['Patient 1.17.0', '2 Oct', 'On-device OCR (Tesseract) with confirm step'],
    ['Patient 1.18.1–1.18.3', '3 Oct', 'Security review fixes: hand-back code, proven family links, sender proof, back button'],
    ['Patient 1.19.0', '3 Oct', 'Choose and change how long a doctor can see the records'],
    ['Patient 1.19.1–1.26.0', '10 Oct', 'First interface review: whose records on every tab, labelled tabs, swipe to tick, photos, “done” animation, larger secondary text, frosted glass bars'],
    ['Patient 1.27.0–1.37.3', '10–11 Oct', 'Second interface review: branded launch screen, slimmer top bar, today’s progress card, welcome screen, touch effect, richer Records cards with sparklines, sliding tab pill'],
    ['Patient 1.38.0–1.40.1', '12 Oct', 'Restyle: Sora typeface, caramel accent, white cards, photo welcome and lock screens, new add pages, profile photos; server over the home Wi-Fi; purple and teal status colours'],
    ['Patient 1.41.0–1.42.0', '12 Oct', 'Family tree; ten-person sample family; reminders editable where shown; LifeTrack IDs for sample people (practice server only)'],
    ['Patient 1.42.1–1.42.2', '14 Oct', 'No text overflow on narrow screens or at large text sizes; smoother scrolling and tab switches'],
    ['Patient 1.42.3', '2 Nov', 'Fix: a finger lifted after its card left the screen; frame-timing test runs unattended'],
    ['Patient 1.43.0–1.44.1', '6 Nov', 'Danger-level check on every saved lab value and home reading; no dose ever defaulted'],
    ['Patient 1.45.0–1.46.1', '6 Nov', 'Allergy and condition entry, one allergy list; real Emergency ID date; report parser reads more and lists what it could not read; Emergency ID edits logged'],
    ['Patient 1.46.2–1.46.4', '6 Nov', 'End access with the server off; erasure ends live shares; family links logged before sending; saves no longer wait for the network'],
    ['Patient 1.47.0–1.47.1', '6 Nov', 'LifeTrack ID reclaim after a fresh identity check; retry of an interrupted restore'],
    ['Patient 1.48.0–1.48.2', '6 Nov', 'Typed weekly medicines; urgent alert survives an odd unit; release build runs the tests; erasure entry says what happened to the backup'],
    ['Doctor 1.0.0', '30 Sep', 'Enrolment, verified code, view shares'],
    ['Doctor 1.1–1.4', '1 Oct', 'Prescriptions, printable PDF, request by ID, dark mode'],
    ['Doctor 1.4.1–1.6.1', '3 Oct', 'Security fixes, patient register, notes, “LifeTrack Clinic” branding, back button'],
    ['Doctor 1.7.0–1.8.0', '3 Oct', 'Ask for a time; Profile tab and screenshot setting'],
    ['Doctor 1.8.1', '2 Nov', 'Fix: an enrolment the server has no record of returns to the form instead of waiting forever'],
    ['Doctor 1.8.2–1.9.1', '6 Nov', 'New prescriptions start with no dose; the summary always has an allergy line; a lab value reported as a limit keeps its sign'],
  ], [2200, 1700, 5126], 'Version history'),
];
const appB = [
  H1('Appendix B: Running the System'),
  H3('Start the server'),
  P('`cd server && dart run bin/server.dart --practice-digilocker`', { align: AlignmentType.LEFT }),
  P('The server prints its address (http://127.0.0.1:8787), the registry public key to pin in the apps, and where the admin token is kept.'),
  H3('Connect an Android phone'),
  P('Enable USB debugging on the phone, connect it, and run `adb reverse tcp:8787 tcp:8787`. The phone’s 127.0.0.1:8787 then reaches the server on the computer.'),
  P('To work without a cable, start the server with `scripts/run_server_lan.sh` instead. It listens on the computer’s Wi-Fi address, which the phone must share.'),
  H3('Build and install the patient app'),
  P('`scripts/build_apk.sh` builds an obfuscated release APK with the computer’s current Wi-Fi address and the registry key built in (`LT_SERVER=usb` builds the cable version); install it with `adb install -r`, which keeps the records already on the phone. Rebuild when the computer’s Wi-Fi address changes.'),
  H3('Run the doctor app'),
  P('`cd doctor && flutter run -d chrome` (or on a phone). Enrol, then approve the enrolment with the admin token.'),
  H3('Run all tests'),
  P('`scripts/check_all.sh` runs analysis and tests in all five packages and reports every failure.'),
];

// ---------- document ----------
const docOpts = ({
  creator: 'LifeTrack Health',
  title: 'LifeTrack Health: NTCC Project Report',
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { size: 32, bold: true, font: FONT }, paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 360 } } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, font: FONT }, paragraph: { spacing: { before: 240, after: 360 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: FONT }, paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1, keepNext: true } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, italics: true, font: FONT }, paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: 'num2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: 'num3', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ] },
});

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } };
const footer = (fmt) => ({ default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })] })] }) });

// Body pages carry a small running header with a rule under it.
const runningHeader = { default: new Header({ children: [new Paragraph({
  alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '999999', space: 4 } },
  children: [new TextRun({ text: 'LifeTrack Health  |  NTCC Project Report', size: 18, color: '666666', italics: true })],
})] }) };

// Build chapters first so figure/table lists are complete.
const body = [...ch1, ...ch2, ...ch3, ...ch4, ...ch5, ...ch6, ...ch7, ...ch8, ...references, ...appA, ...appB];
const lists = [
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('LIST OF FIGURES')] }),
  ...figures.map((f) => new Paragraph({ children: [new TextRun({ text: f, size: 22 })], spacing: { after: 80 } })),
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('LIST OF TABLES')] }),
  ...tables.filter((t) => !t.includes('Project team')).map((t) => new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 80 } })),
  new Paragraph({ heading: HeadingLevel.TITLE, pageBreakBefore: true, children: [new TextRun('LIST OF ABBREVIATIONS')] }),
  ...[['ABDM', 'Ayushman Bharat Digital Mission'], ['ABHA', 'Ayushman Bharat Health Account'], ['AES-GCM', 'Advanced Encryption Standard, Galois/Counter Mode'], ['APK', 'Android Package'], ['DFD', 'Data Flow Diagram'], ['dp', 'Density-independent pixel (Android layout unit)'], ['DPDP', 'Digital Personal Data Protection (Act, 2023)'], ['EMR', 'Electronic Medical Record'], ['ER', 'Entity-Relationship'], ['GF', 'Galois Field'], ['FHIR', 'Fast Healthcare Interoperability Resources'], ['HKDF', 'HMAC-based Key Derivation Function'], ['OCR', 'Optical Character Recognition'], ['QR', 'Quick Response (code)'], ['SaMD', 'Software as a Medical Device'], ['sp', 'Scale-independent pixel (Android text-size unit)'], ['UI', 'User Interface'], ['UIP', 'Universal Immunisation Programme'], ['WCAG', 'Web Content Accessibility Guidelines'], ['ΔE', 'Colour difference measured in the CIELAB colour space']]
    .map(([a, b]) => new Paragraph({ tabStops: [{ type: TabStopType.LEFT, position: 1800 }], children: [new TextRun({ text: a, bold: true }), new TextRun(`\t${b}`)], spacing: { after: 60 } })),
];

const doc = new Document({ ...docOpts, features: { updateFields: true }, sections: [
  { properties: { page: { ...page, borders: {
      pageBorders: { display: 'allPages', offsetFrom: 'page' },
      pageBorderTop: { style: BorderStyle.DOUBLE, size: 12, color: '333333', space: 24 },
      pageBorderBottom: { style: BorderStyle.DOUBLE, size: 12, color: '333333', space: 24 },
      pageBorderLeft: { style: BorderStyle.DOUBLE, size: 12, color: '333333', space: 24 },
      pageBorderRight: { style: BorderStyle.DOUBLE, size: 12, color: '333333', space: 24 },
    } } }, children: cover },
  { properties: { page: { ...page, pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } } }, footers: footer(), children: [...declaration, ...certificate, ...ack, ...abstract, ...(tocPages ? printedToc() : toc), ...lists] },
  { properties: { page: { ...page, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } } }, headers: runningHeader, footers: footer(), children: body },
] });

Packer.toBuffer(doc).then((buf) => {
  const out = process.env.OUT || path.join(__dirname, 'LifeTrack_Health_NTCC_Report.docx');
  fs.writeFileSync(out, buf);
  if (process.env.HEADINGS_OUT) fs.writeFileSync(process.env.HEADINGS_OUT, JSON.stringify(headings));
  console.log('wrote', out, figures.length, 'figures', tables.length, 'tables');
});
