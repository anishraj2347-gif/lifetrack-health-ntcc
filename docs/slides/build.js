// Builds LifeTrack_Health_NTCC_Slides.pptx: the NTCC evaluation deck, on the
// header of the university's "PPT Guidelines" template (Amity logo, a gold and
// a black rule, a navy line along the foot) and in the order the notice of
// 26 Nov 2026 asks for: title, agenda, problem, (objectives), conceptual
// framework and two slides explaining it, three on methodology, two on
// results, then advantages, conclusion and references: 15 slides at most.
//
//   cd docs/slides && npm install && node build.js
//
// Screenshots and diagrams come from docs/report; icons from make_icons.py.
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const REPORT = path.join(__dirname, '..', 'report');
const SHOT = (f) => path.join(REPORT, 'shots', f);
const FIG = (f) => path.join(REPORT, 'fig', f);
const ICON = (name, colour = 'navy') => path.join(__dirname, 'icons', `${name}-${colour}.png`);
const OUT = path.join(__dirname, 'LifeTrack_Health_NTCC_Slides.pptx');

// The template's own colours: gold and black rules, a navy foot line.
const THEME = {
  name: 'Amity NTCC',
  headFontFace: 'Calibri',
  bodyFontFace: 'Calibri',
  colors: {
    dk1: '1A1A1A', lt1: 'FFFFFF', dk2: '002060', lt2: 'EEF2F8',
    accent1: 'FFCA08', accent2: '1F4E99', accent3: 'A65E30', accent4: '7E22CE', accent5: '115E59', accent6: '5B6472',
    hlink: '1F4E99', folHlink: '7E22CE',
  },
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 × 7.5 in, as the template
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'LifeTrack Health: NTCC Term Paper';
pres.author = 'Anish Raj';
const C = pres.SchemeColor;

// ---------- layouts ----------
const header = [
  { image: { x: 0.797, y: 0.18, w: 2.323, h: 0.875, path: path.join(__dirname, 'assets', 'amity-logo.png'), altText: 'Amity University Patna logo' } },
  { line: { x: 0.917, y: 1.153, w: 5.75, h: 0, line: { color: C.accent1, width: 4.5 } } },
  { line: { x: 6.667, y: 1.153, w: 5.75, h: 0, line: { color: C.text1, width: 4.5 } } },
  { line: { x: 0, y: 7.47, w: 13.333, h: 0, line: { color: C.text2, width: 4.5 } } },
];
pres.defineSlideMaster({
  title: 'AMITY_TITLE',
  background: { color: C.background1 },
  objects: [
    ...header,
    { placeholder: { options: { name: 'title', type: 'title', x: 0.9, y: 1.45, w: 11.53, h: 0.85, align: 'center', valign: 'middle', fontSize: 40, bold: true, color: C.text1 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'AMITY_CONTENT',
  background: { color: C.background1 },
  objects: [
    ...header,
    { placeholder: { options: { name: 'title', type: 'title', x: 0.9, y: 1.28, w: 11.53, h: 0.62, align: 'center', valign: 'middle', fontSize: 28, bold: true, color: C.text1 }, text: '' } },
    { text: { text: 'LifeTrack Health  ·  NTCC Term Paper  ·  AIIT, Amity University Patna', options: { x: 0.9, y: 7.02, w: 8, h: 0.3, fontSize: 10, color: C.accent6 } } },
  ],
  slideNumber: { x: 11.93, y: 7.02, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: 'right' },
});

// ---------- helpers ----------
let n = 0;
const id = (what) => `${what} ${++n}`;
function content(title, section) {
  const s = pres.addSlide({ masterName: 'AMITY_CONTENT', sectionTitle: section });
  s.addText(title, { placeholder: 'title' });
  return s;
}
const text = (s, t, o) => s.addText(t, { isTextBox: true, color: C.text1, fontSize: 16, valign: 'top', margin: 0, objectName: id('text'), ...o });
// An icon in a filled circle.
function badge(s, icon, x, y, d = 0.72, fill = C.text2) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: id('badge') });
  s.addImage({ path: ICON(icon, 'white'), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, altText: icon, objectName: id('icon') });
}
function card(s, x, y, w, h, fill = C.background2) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: id('card') });
}
// A phone screenshot with a thin dark frame.
function phone(s, file, x, y, h, caption) {
  const w = h * 600 / 1300;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x - 0.06, y: y - 0.06, w: w + 0.12, h: h + 0.12, rectRadius: 0.12, fill: { color: '2B2B2B' }, line: { color: '2B2B2B', width: 0 }, objectName: id('frame') });
  s.addImage({ path: file, x, y, w, h, altText: caption, objectName: id('screenshot') });
  if (caption) text(s, caption, { x: x - 0.4, y: y + h + 0.14, w: w + 0.8, h: 0.32, fontSize: 14, bold: true, align: 'center' });
}

// ---------- 1. Title ----------
pres.addSection({ title: 'Introduction' });
{
  const s = pres.addSlide({ masterName: 'AMITY_TITLE', sectionTitle: 'Introduction' });
  s.addText('LifeTrack Health', { placeholder: 'title' });
  text(s, 'A Patient-Owned, Local-First Health Record with Consent-Controlled Sharing', { x: 0.9, y: 2.3, w: 11.53, h: 0.5, fontSize: 22, color: C.text2, align: 'center' });
  text(s, [
    { text: 'Presented By', options: { bold: true, fontSize: 20, breakLine: true } },
    { text: 'Anish Raj (A45304925031)', options: { fontSize: 18 } },
  ], { x: 0.9, y: 3.35, w: 5.2, h: 1.0, align: 'center', paraSpaceAfter: 6 });
  text(s, [
    { text: 'Under the Supervision of', options: { bold: true, fontSize: 20, breakLine: true } },
    { text: 'Dr. Priyadarshini Gupta', options: { fontSize: 18, breakLine: true } },
    { text: 'Assistant Professor', options: { fontSize: 18 } },
  ], { x: 7.23, y: 3.35, w: 5.2, h: 1.3, align: 'center', paraSpaceAfter: 6 });
  s.addShape(pres.shapes.LINE, { x: 6.667, y: 3.4, w: 0, h: 1.2, line: { color: C.accent1, width: 2 }, objectName: id('divider') });
  text(s, [
    { text: 'NTCC Term Paper (ETTP101)  ·  B.Sc. IT, Semester III', options: { fontSize: 16, breakLine: true, color: C.accent6 } },
    { text: 'Amity Institute of Information Technology, Patna', options: { fontSize: 18, breakLine: true } },
    { text: '8 December 2026', options: { fontSize: 18 } },
  ], { x: 0.9, y: 5.25, w: 11.53, h: 1.2, align: 'center', paraSpaceAfter: 4 });
  s.addNotes('Introduce yourself and the project in one sentence: LifeTrack Health keeps a family\'s health records on the patient\'s own phone and lets the patient share exactly the right part with a verified doctor for a limited time.');
}

// ---------- 2. Agenda ----------
{
  const s = content('Agenda', 'Introduction');
  const items = ['Problem statement', 'Objectives', 'Conceptual framework', 'How the framework works', 'Methodology', 'Results', 'Advantages and limitations', 'Conclusion and future scope', 'References'];
  items.forEach((t, i) => {
    const col = i < 5 ? 0 : 1, row = i < 5 ? i : i - 5;
    const x = col ? 7.0 : 1.6, y = 2.2 + row * 0.92;
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.6, h: 0.6, fill: { color: i % 2 ? C.accent2 : C.text2 }, line: { color: C.background1, width: 0 }, objectName: id('number') });
    text(s, String(i + 1), { x, y, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    text(s, t, { x: x + 0.85, y, w: 4.6, h: 0.6, fontSize: 22, valign: 'middle' });
  });
  s.addNotes('The talk follows the order the evaluation notice asks for: the problem, the framework and how it works, the method, the results, then the limits and what comes next.');
}

// ---------- 3. Problem statement ----------
pres.addSection({ title: 'Problem and objectives' });
{
  const s = content('Problem Statement', 'Problem and objectives');
  text(s, 'A patient is present at every medical visit, yet holds no complete, trustworthy record of their own health, and no safe way to show the right part of it to a new doctor.', { x: 0.9, y: 2.05, w: 11.53, h: 0.8, fontSize: 19, align: 'center' });
  const cards = [
    ['repeat', 'Repeated tests', 'A past result is missing or cannot be trusted, so the test is ordered again: more cost and delay for the patient.'],
    ['pill', 'Duplicate medicines', 'A doctor who cannot see current medicines may prescribe the same ingredient again.'],
    ['users', 'Missed family patterns', 'Illnesses in the family are rarely recorded in a way a doctor can use.'],
  ];
  cards.forEach(([icon, head, body], i) => {
    const x = 0.9 + i * 3.88;
    card(s, x, 3.0, 3.65, 2.15);
    badge(s, icon, x + 0.25, 3.2, 0.7);
    text(s, head, { x: x + 1.12, y: 3.25, w: 2.4, h: 0.6, fontSize: 20, bold: true, valign: 'middle' });
    text(s, body, { x: x + 0.25, y: 4.02, w: 3.15, h: 1.05, fontSize: 15 });
  });
  card(s, 0.9, 5.35, 11.53, 1.3, C.text2);
  text(s, [
    { text: 'Research question:  ', options: { bold: true, color: C.accent1 } },
    { text: 'can a patient keep a complete, useful record on their own phone and share exactly the right part with exactly the right doctor, for a limited time, without any server holding a key that opens it?', options: { color: C.background1 } },
  ], { x: 1.2, y: 5.45, w: 10.93, h: 1.1, fontSize: 18, valign: 'middle' });
  s.addNotes('Today the record is scattered across clinics and labs, and the patient carries paper. Existing solutions keep data inside one hospital, hold identity but not the record, or store files a provider can read. The question is whether the patient can hold the record and share it safely.');
}

// ---------- 4. Objectives ----------
{
  const s = content('Objectives', 'Problem and objectives');
  const goals = [
    ['lock', 'A patient-owned, encrypted record for the whole family'],
    ['timer', 'Time-limited, revocable access for one verified doctor'],
    ['shield-check', 'No server holds a key to shares or backups'],
    ['sparkles', 'Safe insights on the phone, never a diagnosis'],
    ['scan-text', 'Read lab reports on the phone (OCR)'],
    ['stethoscope', 'A verified doctor app with signed prescriptions'],
    ['clipboard-check', 'Privacy rules checked by automated tests'],
    ['accessibility', 'Readable and accessible for older users'],
  ];
  goals.forEach(([icon, t], i) => {
    const x = 0.9 + (i % 4) * 2.93, y = i < 4 ? 2.1 : 4.45;
    card(s, x, y, 2.75, 2.15);
    text(s, String(i + 1).padStart(2, '0'), { x: x + 1.85, y: y + 0.18, w: 0.7, h: 0.4, fontSize: 18, bold: true, color: C.accent3, align: 'right' });
    s.addImage({ path: ICON(icon), x: x + 0.25, y: y + 0.22, w: 0.62, h: 0.62, altText: icon, objectName: id('icon') });
    text(s, t, { x: x + 0.25, y: y + 1.0, w: 2.3, h: 1.0, fontSize: 16 });
  });
  s.addNotes('Eight objectives. The ones to stress: the record stays on the phone, sharing is time-limited to one verified doctor, and no server can open shared data or backups. Objectives 5 and 8 are only partly met, as the results slide shows.');
}

// ---------- 5. Conceptual framework ----------
pres.addSection({ title: 'Framework' });
{
  const s = content('Conceptual Framework', 'Framework');
  s.addImage({ path: FIG('arch.png'), x: 0.9, y: 2.05, w: 7.1, h: 7.1 * 993 / 1824, altText: 'System architecture: patient app, doctor app, server and shared protocol', objectName: id('diagram') });
  const ideas = [
    ['smartphone', 'Local-first', 'The record lives on the patient’s phone, sealed with AES-256-GCM.'],
    ['key-round', 'Consent-controlled', 'Nothing is read without a stored grant; every read is logged.'],
    ['server', 'Server-blind', 'The server relays and stores only ciphertext it cannot open.'],
  ];
  ideas.forEach(([icon, head, body], i) => {
    const y = 2.1 + i * 1.45;
    badge(s, icon, 8.35, y, 0.72);
    text(s, head, { x: 9.25, y, w: 3.2, h: 0.4, fontSize: 19, bold: true });
    text(s, body, { x: 9.25, y: y + 0.42, w: 3.2, h: 0.85, fontSize: 15 });
  });
  text(s, 'Four parts: the patient app  ·  LifeTrack Clinic, the doctor app  ·  a small Dart server  ·  a shared protocol package', { x: 0.9, y: 6.4, w: 11.53, h: 0.4, fontSize: 15, color: C.accent6, align: 'center' });
  s.addNotes('Three ideas hold the design together. Local-first: the main copy is on the phone. Consent-controlled: one authority decides every read and the audit log records it. Server-blind: the server is needed to carry and store data, but it never holds a key that opens it.');
}

// ---------- 6. Framework: inside the patient app ----------
{
  const s = content('How It Works: Inside the Patient App', 'Framework');
  s.addImage({ path: FIG('layers.png'), x: 0.9, y: 2.1, w: 6.7, h: 6.7 * 1127 / 2010, altText: 'Layers of the patient app', objectName: id('diagram') });
  text(s, [
    { text: 'Screens call one entry point, AppServices', options: { bullet: true, breakLine: true } },
    { text: 'Anyone but the patient reads only through RecordAccess, which asks ConsentStore for a stored grant', options: { bullet: true, breakLine: true } },
    { text: 'Three planes: A on-device insights, B report capture (OCR), C cloud analysis (designed, switched off)', options: { bullet: true, breakLine: true } },
    { text: 'SealedStore encrypts every record and binds it to its name', options: { bullet: true, breakLine: true } },
    { text: 'The audit entry is written before anything leaves the phone', options: { bullet: true } },
  ], { x: 7.95, y: 2.15, w: 4.5, h: 4.4, fontSize: 17, paraSpaceAfter: 10 });
  s.addNotes('Privacy is part of the structure, not a feature. Because ConsentStore is the only grant authority and RecordAccess the only read path for anyone but the patient, sharing, insights and the switched-off cloud plane all inherit the same checks and the same audit trail.');
}

// ---------- 7. Framework: sharing with a doctor ----------
{
  const s = content('How It Works: Sharing With a Doctor', 'Framework');
  const steps = [
    'Patient scans the doctor’s verified code',
    'Phone checks the registry certificate and the doctor’s status',
    'Patient picks the level and the time (1 hour to 2 days)',
    'Records sealed to the doctor’s key (X25519 + AES-GCM)',
    'Audit entry written, then sent through the relay',
    'Access closes at the end time; the relay deletes the copy',
  ];
  steps.forEach((t, i) => {
    const x = 0.9 + (i % 3) * 3.95, y = i < 3 ? 2.1 : 3.75;
    card(s, x, y, 3.55, 1.3);
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.32, w: 0.62, h: 0.62, fill: { color: C.text2 }, line: { color: C.text2, width: 0 }, objectName: id('step') });
    text(s, String(i + 1), { x: x + 0.2, y: y + 0.32, w: 0.62, h: 0.62, fontSize: 18, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    text(s, t, { x: x + 1.0, y: y + 0.12, w: 2.4, h: 1.06, fontSize: 15, valign: 'middle' });
    if (i % 3 < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.6, y: y + 0.5, w: 0.3, h: 0.3, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0 }, objectName: id('arrow') });
  });
  card(s, 0.9, 5.3, 5.65, 1.45, C.text2);
  text(s, [
    { text: 'Backup vault', options: { bold: true, color: C.accent1, breakLine: true } },
    { text: 'App-generated password → Argon2id → vault key; Shamir 2-of-3 recovery codes; the server keeps only ciphertext.', options: { color: C.background1 } },
  ], { x: 1.15, y: 5.42, w: 5.2, h: 1.25, fontSize: 15 });
  card(s, 6.78, 5.3, 5.65, 1.45, C.text2);
  text(s, [
    { text: 'Access levels', options: { bold: true, color: C.accent1, breakLine: true } },
    { text: 'Summary (default)  ·  All reports  ·  Family illnesses  ·  Home readings. The patient can shorten, extend or end a share.', options: { color: C.background1 } },
  ], { x: 7.03, y: 5.42, w: 5.2, h: 1.25, fontSize: 15 });
  s.addNotes('Each step is a check that can stop the share: an unsigned code is refused, a revoked doctor is refused, the patient can cancel. A share the relay did not accept ends its grant at once. The audit entry is written before the bundle leaves, so even a failed attempt is on record.');
}

// ---------- 8. Methodology: incremental development ----------
pres.addSection({ title: 'Methodology' });
{
  const s = content('Methodology: Incremental Development', 'Methodology');
  const stages = ['Requirements and threats', 'Architecture and invariants', 'Feature releases', 'Security reviews', 'Interface reviews', 'Integration and phone tests'];
  stages.forEach((t, i) => {
    s.addShape(pres.shapes.CHEVRON, { x: 0.9 + i * 1.93, y: 2.05, w: 2.05, h: 1.0, fill: { color: i % 2 ? C.accent2 : C.text2 }, line: { color: C.background1, width: 1 }, objectName: id('stage') });
    // Inside the solid part of the chevron: its left notch shows the white page.
    text(s, t, { x: 1.36 + i * 1.93, y: 2.1, w: 1.2, h: 0.9, fontSize: 12, bold: true, color: C.background1, align: 'center', valign: 'middle' });
  });
  // A simple timeline over the fifteen weeks (1 Sep – 8 Dec 2026).
  const x0 = 3.6, wk = (12.43 - x0) / 15, y0 = 3.35;
  for (let w = 1; w <= 15; w++) text(s, `W${w}`, { x: x0 + (w - 1) * wk, y: y0, w: wk, h: 0.28, fontSize: 10, color: C.accent6, align: 'center' });
  const rows = [
    ['Study and design', 1, 4, C.text2],
    ['Build: patient 1.0 → 1.42, doctor 1.0 → 1.8', 5, 7, C.text2],
    ['Timing, third review, audit (→ 1.48)', 8, 10, C.text2],
    ['Glass interface, phone test (1.49)', 13, 13, C.accent2],
    ['Report in university format, slides', 13, 14, C.accent2],
    ['External evaluation, 8 December', 15, 15, C.accent1],
  ];
  rows.forEach(([label, a, b, fill], i) => {
    const y = y0 + 0.36 + i * 0.42;
    text(s, label, { x: 0.9, y, w: 2.6, h: 0.36, fontSize: 13, valign: 'middle' });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x0 + (a - 1) * wk + 0.03, y: y + 0.04, w: (b - a + 1) * wk - 0.06, h: 0.28, rectRadius: 0.06, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: id('bar') });
  });
  text(s, 'Each stage ended with working, tested software. Every update raised the version and added a changelog line: patient app 1.0.0 → 1.49.2, doctor app 1.0.0 → 1.9.1.', { x: 0.9, y: 6.4, w: 11.53, h: 0.5, fontSize: 14, color: C.accent6, align: 'center' });
  s.addNotes('Development was incremental: each feature was its own version with its own tests. Weeks 11 and 12 slipped; the evaluation date moved the plan to week 15. The usability check with older users has not been done yet, and the report says so.');
}

// ---------- 9. Methodology: technology stack ----------
{
  const s = content('Methodology: Technology Stack', 'Methodology');
  const tiles = [
    ['code', 'Flutter and Dart', 'One code base for the patient app, the doctor app, the server and the protocol package'],
    ['key-round', 'Android Keystore', 'Keeps the device key out of the app’s storage, in hardware where the phone supports it'],
    ['shield-check', 'cryptography package', 'AES-256-GCM, X25519, Ed25519, HKDF and Argon2id: standard primitives, none written here'],
    ['scan-text', 'Tesseract OCR', 'Reads report photographs on the phone, fully offline'],
    ['server', 'Dart server (shelf)', 'Doctor registry, encrypted relay, ID directory and backup vault'],
    ['git-branch', 'Git and test scripts', 'check_all.sh runs every analysis and test; build_apk.sh tests, then builds the release'],
  ];
  tiles.forEach(([icon, head, body], i) => {
    const x = 0.9 + (i % 3) * 3.88, y = i < 3 ? 2.1 : 4.45;
    card(s, x, y, 3.65, 2.1);
    badge(s, icon, x + 0.25, y + 0.25, 0.68);
    text(s, head, { x: x + 1.08, y: y + 0.3, w: 2.45, h: 0.6, fontSize: 18, bold: true, valign: 'middle' });
    text(s, body, { x: x + 0.25, y: y + 1.05, w: 3.2, h: 0.95, fontSize: 15 });
  });
  s.addNotes('One language everywhere means the apps and the server share one set of message formats, so they cannot drift apart. The only cryptography written for the project is Shamir splitting, which is unit-tested but not independently audited.');
}

// ---------- 10. Methodology: testing ----------
{
  const s = content('Methodology: Testing', 'Methodology');
  text(s, '“A claim is accepted only if a test can prove it.”', { x: 0.9, y: 2.05, w: 6.0, h: 0.5, fontSize: 20, bold: true, italic: true, color: C.text2 });
  const kinds = [
    ['Unit', 'ranges, thresholds, Verhoeff, Shamir'],
    ['Widget', 'real screens, vibrations, error shake'],
    ['Invariant', 'scan the source: one network file, no self-granting agent'],
    ['Contrast', 'every text meets WCAG AA, light and dark'],
    ['Layout', '3 widths × 3 text sizes, no overflow'],
    ['Colour vision', 'status colours stay ≥ 25 ΔE apart'],
    ['Photo crop', 'every photograph keeps its subject in view'],
    ['Integration', 'both apps against the real server'],
  ];
  text(s, kinds.map(([k, v], i) => ({ text: `${k}: `, options: { bold: true, bullet: true, breakLine: false } })).flatMap((head, i) => [head, { text: kinds[i][1], options: { breakLine: i < kinds.length - 1 } }]),
    { x: 0.9, y: 2.7, w: 6.1, h: 3.4, fontSize: 15, paraSpaceAfter: 5 });
  s.addChart(pres.charts.BAR, [{ name: 'Tests', labels: ['Patient app', 'Server', 'Doctor app', 'Protocol', 'Integration'], values: [522, 57, 28, 8, 6] }], {
    x: 7.2, y: 2.0, w: 5.25, h: 4.0, barDir: 'bar', catAxisOrientation: 'maxMin',
    chartColors: ['1F4E99'], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 13, dataLabelFontFace: '+mn-lt', dataLabelColor: '1A1A1A',
    showTitle: true, title: 'Automated tests by package', titleFontSize: 15, titleFontFace: '+mn-lt', titleColor: '1A1A1A',
    catAxisLabelFontSize: 13, catAxisLabelFontFace: '+mn-lt', catAxisLabelColor: '1A1A1A', valAxisHidden: true,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false, barGapWidthPct: 60,
  });
  card(s, 0.9, 6.25, 11.53, 0.6, C.text2);
  text(s, [
    { text: '621 tests passing', options: { bold: true, color: C.accent1 } },
    { text: '   ·   0 failing   ·   clean static analysis in all five packages   ·   5 skipped (render tools)', options: { color: C.background1 } },
  ], { x: 1.1, y: 6.25, w: 11.13, h: 0.6, fontSize: 16, align: 'center', valign: 'middle' });
  s.addNotes('The invariant tests are the unusual part: they read the source code and fail the build if, for example, any file but the one network file can open a socket, or an agent could grant itself access. Passing tests show the written rules hold, not that the right rules were written: an audit and a phone test still found gaps, which were fixed with new tests.');
}

// ---------- 11. Result overview: the patient app ----------
pres.addSection({ title: 'Results' });
{
  const s = content('Result Overview: The Patient App', 'Results');
  const shots = [['1-reminders.png', 'Reminders'], ['2-records.png', 'Records'], ['3-insights.png', 'Insights'], ['emergency.png', 'Emergency ID']];
  const h = 4.25, w = h * 600 / 1300, gap = 0.95;
  const x0 = (13.333 - (shots.length * w + (shots.length - 1) * gap)) / 2;
  shots.forEach(([f, cap], i) => phone(s, SHOT(f), x0 + i * (w + gap), 2.1, h, cap));
  s.addNotes('Rendered from the current code, version 1.49.2, with the synthetic sample family. Reminders shows today\'s medicines; Records plots one test across reports; Insights, once turned on, lists findings with the lab\'s range and never a diagnosis; the Emergency ID opens from the lock screen and turns over to a plain-text QR code.');
}

// ---------- 12. Result overview: outcomes ----------
{
  const s = content('Result Overview: Objectives and Numbers', 'Results');
  const rows = [
    ['Patient-owned, encrypted family record', 'Achieved'],
    ['Time-limited, revocable doctor access', 'Achieved'],
    ['No server can open shares or backups', 'Achieved'],
    ['Safe on-device insights', 'Achieved*'],
    ['Reading reports on the phone (OCR)', 'Partly'],
    ['Verified doctor app, signed prescriptions', 'Achieved'],
    ['Privacy rules checked by tests', 'Achieved'],
    ['Accessible for older users', 'Partly'],
  ];
  const head = { bold: true, color: 'FFFFFF', fill: { color: '002060' } };
  s.addTable([
    [{ text: 'Objective', options: head }, { text: 'Outcome', options: { ...head, align: 'center' } }],
    ...rows.map(([o, r]) => [{ text: o }, { text: r, options: { bold: true, align: 'center', color: r.startsWith('Partly') ? 'A65E30' : '115E59' } }]),
  ], { x: 0.9, y: 2.05, w: 7.1, colW: [5.3, 1.8], fontSize: 15, rowH: 0.42, border: { type: 'solid', pt: 0.75, color: 'C9D1DE' }, color: '1A1A1A', valign: 'middle', objectName: id('outcomes') });
  text(s, '* No doctor has reviewed the rules yet. Partly: OCR is not yet measured on real reports; usability with older users is not yet tested.', { x: 0.9, y: 6.0, w: 7.1, h: 0.7, fontSize: 12, color: C.accent6 });
  const stats = [['28,400', 'lines of Dart in four packages'], ['621', 'automated tests, all passing'], ['1.49.2', 'patient app, run on a Galaxy S25']];
  stats.forEach(([big, small], i) => {
    const y = 2.05 + i * 1.55;
    card(s, 8.4, y, 4.03, 1.35);
    text(s, big, { x: 8.6, y: y + 0.12, w: 3.6, h: 0.7, fontSize: 36, bold: true, color: C.text2 });
    text(s, small, { x: 8.6, y: y + 0.82, w: 3.6, h: 0.4, fontSize: 15, color: C.accent6 });
  });
  s.addNotes('Six objectives are met and two partly. Be ready to say what is missing: OCR accuracy on real photographed reports, a usability check with older users, and a doctor\'s review of the clinical rules.');
}

// ---------- 13. Advantages and limitations ----------
pres.addSection({ title: 'Conclusion' });
{
  const s = content('Advantages and Limitations', 'Conclusion');
  const pros = [
    'The patient holds the complete record; it works offline',
    'No server can open shared records or backups',
    'Sharing is time-limited, to one verified doctor, and logged',
    'Insights never diagnose; urgent values prompt for care',
    'Privacy rules are checked by automated tests',
    'Family profiles, Emergency ID and encrypted backup',
  ];
  const cons = [
    'Synthetic data only; no doctor has reviewed it',
    'OCR is English-only and not measured on real reports',
    'Tested on one phone; not tried with older users',
    'Prototype server: plain HTTP on home Wi-Fi, no TLS',
    'No ABDM / ABHA integration yet',
    'English only',
  ];
  [[pros, 'Advantages', 'circle-check', C.accent5, 0.9], [cons, 'Limitations', 'triangle-alert', C.accent3, 6.78]].forEach(([list, title, icon, colour, x]) => {
    card(s, x, 2.05, 5.65, 4.7);
    badge(s, icon, x + 0.3, 2.25, 0.62, colour);
    text(s, title, { x: x + 1.1, y: 2.25, w: 4.2, h: 0.62, fontSize: 22, bold: true, valign: 'middle', color: colour });
    text(s, list.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < list.length - 1 } })), { x: x + 0.3, y: 3.1, w: 5.1, h: 3.5, fontSize: 16, paraSpaceAfter: 8 });
  });
  s.addNotes('Be candid about the limits: they are also in the report. The prototype has only been used with synthetic data, and a legal review is needed before any pilot.');
}

// ---------- 14. Conclusion and future scope ----------
{
  const s = content('Conclusion and Future Scope', 'Conclusion');
  card(s, 0.9, 2.05, 11.53, 1.45, C.background2);
  text(s, 'LifeTrack Health shows, as a working prototype, that a family’s health record can live on the patient’s own phone and reach a verified doctor for a limited time, through a server that holds no key to open it, with its privacy and safety rules checked by 621 automated tests.', { x: 1.2, y: 2.15, w: 10.93, h: 1.25, fontSize: 18, valign: 'middle' });
  const next = [
    ['hospital', 'ABDM / ABHA integration'],
    ['badge-check', 'DigiLocker partner registration'],
    ['file-text', 'FHIR R4 export with provenance'],
    ['globe', 'Hosted server with TLS'],
    ['languages', 'Hindi and other Indian languages'],
    ['scan-line', 'OCR study on real reports; pilot after clinical and legal review'],
  ];
  text(s, 'Future scope', { x: 0.9, y: 3.7, w: 6, h: 0.45, fontSize: 20, bold: true, color: C.text2 });
  next.forEach(([icon, t], i) => {
    const x = 0.9 + (i % 3) * 3.88, y = i < 3 ? 4.3 : 5.5;
    badge(s, icon, x, y, 0.62, C.accent2);
    text(s, t, { x: x + 0.8, y: y - 0.05, w: 2.95, h: 0.75, fontSize: 15, valign: 'middle' });
  });
  s.addNotes('Close with the one-sentence answer to the research question, then what it would take to go further: ABDM integration, a hosted server with TLS, Indian languages, and a pilot only after clinical and legal review.');
}

// ---------- 15. References ----------
{
  const s = content('References', 'Conclusion');
  const refs = [
    'Government of India: The Digital Personal Data Protection Act, 2023 (Act No. 22 of 2023). Ministry of Law and Justice (2023).',
    'National Health Authority: Ayushman Bharat Digital Mission (ABDM), https://abdm.gov.in, last accessed 2026/11/27.',
    'Board of Governors in supersession of the MCI: Telemedicine Practice Guidelines. New Delhi (2020).',
    'Kleppmann, M., et al.: Local-first software: you own your data, in spite of the cloud. In: Onward! 2019, pp. 154–178. ACM (2019).',
    'Tang, P.C., et al.: Personal health records: definitions, benefits, and strategies for overcoming barriers to adoption. JAMIA 13(2), 121–126 (2006).',
    'Shamir, A.: How to share a secret. Communications of the ACM 22(11), 612–613 (1979).',
    'Biryukov, A., et al.: Argon2 Memory-Hard Function for Password Hashing. RFC 9106, IETF (2021).',
    'Langley, A., et al.: Elliptic Curves for Security. RFC 7748, IETF (2016).',
    'W3C: Web Content Accessibility Guidelines (WCAG) 2.1, https://www.w3.org/TR/WCAG21/, last accessed 2026/11/27.',
  ];
  text(s, refs.map((r, i) => ({ text: `[${i + 1}]  ${r}`, options: { breakLine: i < refs.length - 1 } })), { x: 0.9, y: 2.05, w: 11.53, h: 4.2, fontSize: 14, paraSpaceAfter: 6 });
  text(s, 'The full list of 27 references is in the report. Parts of the code and of the report were drafted with the help of an AI coding assistant (Claude, by Anthropic); I directed the work, reviewed what it produced, and ran and checked the tests.', { x: 0.9, y: 6.3, w: 11.53, h: 0.6, fontSize: 12, italic: true, color: C.accent6 });
  s.addNotes('Thank the examiners and invite questions.');
}

if (pres.slides.length > 15) throw new Error(`${pres.slides.length} slides: the notice allows 15`);

// pptxgenjs writes Office's palette into the theme; write this deck's colours instead,
// and mend the chart axis attribute PowerPoint would otherwise offer to repair.
async function applyTheme(file, theme) {
  const JSZip = require(require.resolve('jszip', { paths: [require.resolve('pptxgenjs')] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file), { createFolders: false });
  // No folder entries: a file with them can fail to open through a cloud service (the
  // notice asks for the deck on Google Drive). Only the entries go; the files inside stay.
  for (const [name, entry] of Object.entries(zip.files)) if (entry.dir) delete zip.files[name];
  const SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const part = 'ppt/theme/theme1.xml';
  const xml = await zip.file(part).async('string');
  const scheme = `<a:clrScheme name="${theme.name}">` + SLOTS.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k]}"/></a:${k}>`).join('') + '</a:clrScheme>';
  const themed = xml.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme).replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${theme.name}"`);
  if (!themed.includes(scheme)) throw new Error('no colour scheme to replace in ' + part);
  zip.file(part, themed);
  for (const name of Object.keys(zip.files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))) {
    const chart = await zip.file(name).async('string');
    zip.file(name, chart.replace(/<c:crossesAt val="autoZero"\/>/g, '<c:crosses val="autoZero"/>'));
  }
  for (const [name, entry] of Object.entries(zip.files)) if (entry.dir) delete zip.files[name]; // zip.file() made some again
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

pres.writeFile({ fileName: OUT }).then(async () => {
  await applyTheme(OUT, THEME);
  console.log('wrote', OUT, pres.slides.length, 'slides');
});
