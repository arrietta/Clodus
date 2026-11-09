// Builds Unveil_Agile_Presentation.pptx
// Run: NODE_PATH=<dir with pptxgenjs> node build/build_deck.js
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = path.join(__dirname, "..", "Unveil_Agile_Presentation.pptx");

const HEX = {
  ink: "111111", white: "FFFFFF", muted: "5F5F5F", card: "EFEFEF",
  yellow: "F2C318", line: "D4D4D4", card2: "E3E3E3", green: "2E7D32",
  dark2: "2A2A2A", grayOnDark: "A8A8A8",
};
const THEME = {
  name: "Unveil",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.muted, lt2: HEX.card,
    accent1: HEX.yellow, accent2: HEX.card2, accent3: HEX.line, accent4: HEX.green,
    accent5: HEX.dark2, accent6: HEX.grayOnDark, hlink: "1F5FBF", folHlink: "6B3FA0",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Agile Project Management: Unveil";
pres.author = "Marlen Kalbayev, Diyor Khabibullayev, Yeldos Bakayev, Farkhad Amanbayev";
pres.subject = "IT Project Management - Agile report";
const C = pres.SchemeColor;
// C.text1 = ink, C.text2 = muted, C.background1 = white, C.background2 = card,
// accent1 = yellow, accent2 = card2, accent3 = line, accent4 = green, accent5 = dark2, accent6 = grayOnDark

const X = 0.6, W = 8.8;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "Content",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: X, y: 0.32, w: W, h: 0.3, fontSize: 11, bold: true, color: C.text2, charSpacing: 2, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: X, y: 0.62, w: W, h: 0.62, align: "left", fontSize: 28, bold: true, color: C.text1, margin: 0, valign: "top" }, text: "" } },
    { text: { text: "Unveil · Agile report", options: { x: X, y: 5.17, w: 4, h: 0.25, fontSize: 10, color: C.text2, margin: 0 } } },
  ],
  slideNumber: { x: 8.9, y: 5.17, w: 0.5, h: 0.25, fontSize: 10, color: C.text2, align: "right", margin: 0 },
});
pres.defineSlideMaster({
  title: "Dark",
  background: { color: C.text1 },
  objects: [],
});
pres.defineSlideMaster({
  title: "Yellow",
  background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: X, y: 0.32, w: W, h: 0.3, fontSize: 11, bold: true, color: C.text1, charSpacing: 2, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: X, y: 0.7, w: W, h: 1.0, align: "left", fontSize: 36, bold: true, color: C.text1, margin: 0, valign: "top" }, text: "" } },
    { text: { text: "Unveil · Agile report", options: { x: X, y: 5.17, w: 4, h: 0.25, fontSize: 10, color: C.text1, margin: 0 } } },
  ],
  slideNumber: { x: 8.9, y: 5.17, w: 0.5, h: 0.25, fontSize: 10, color: C.text1, align: "right", margin: 0 },
});

// ---------- helpers ----------
let n = 0;
function card(slide, x, y, w, h, fill, extra = {}) {
  n += 1;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color: fill },
    line: extra.line ? { color: extra.line, width: 1 } : { type: "none" },
    objectName: extra.name || `card-${n}`,
  });
}
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontSize: 14, color: C.text1 }, o));
}
function content(kicker, title, section) {
  const s = pres.addSlide({ masterName: "Content", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function table(slide, rows, opts) {
  const header = rows[0].map((t) => ({ text: t, options: { bold: true, color: C.text1, fill: { color: C.background2 } } }));
  const body = rows.slice(1).map((r) => r.map((cell) => (typeof cell === "string" ? { text: cell } : cell)));
  slide.addTable([header, ...body], Object.assign({
    x: X, w: W, fontSize: 12, color: C.text1, fontFace: "Arial",
    border: { type: "solid", pt: 0.75, color: HEX.line }, margin: [0.04, 0.08, 0.04, 0.08], valign: "middle",
  }, opts));
}

// ===================== 1. Title =====================
pres.addSection({ title: "Introduction" });
{
  const s = pres.addSlide({ masterName: "Dark", sectionTitle: "Introduction" });
  txt(s, "IT PROJECT MANAGEMENT · AGILE REPORT", { x: X, y: 0.55, w: 6, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2 });
  txt(s, [
    { text: "Agile, the way", options: { breakLine: true } },
    { text: "we build ", options: {} },
    { text: "Unveil", options: { color: C.accent1 } },
  ], { x: X, y: 1.0, w: 6, h: 1.55, fontSize: 40, bold: true, color: C.background1 });
  txt(s, "A dating web service where you meet the person first, built in six two-week sprints inside five stage gates.",
    { x: X, y: 2.75, w: 5.6, h: 0.8, fontSize: 15, color: C.accent6 });
  txt(s, [
    { text: "Marlen Kalbayev · Diyor Khabibullayev · Yeldos Bakayev · Farkhad Amanbayev", options: { breakLine: true } },
    { text: "Lecturer: Muhammad Khalil Shahid · Group presentation" },
  ], { x: X, y: 4.3, w: 8.5, h: 0.6, fontSize: 12, color: C.accent6, paraSpaceAfter: 4 });

  // sprint chips motif
  const cx = 6.95, cy = 1.05, sz = 0.7, gap = 0.18;
  ["S1", "S2", "S3", "S4", "S5", "S6"].forEach((lab, i) => {
    const x = cx + (i % 3) * (sz + gap), y = cy + Math.floor(i / 3) * (sz + gap);
    const done = i < 2, now = i === 2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: sz, h: sz, rectRadius: 0.1,
      fill: { color: done ? C.accent1 : C.accent5 },
      line: now ? { color: C.accent1, width: 2 } : { type: "none" },
      objectName: `chip-${lab}`,
    });
    txt(s, lab, { x, y, w: sz, h: sz, align: "center", valign: "middle", fontSize: 16, bold: true,
      color: done ? C.text1 : now ? C.accent1 : C.accent6 });
  });
  txt(s, "We are in Sprint 3 of 6", { x: cx, y: cy + 2 * (sz + gap) + 0.02, w: 2.5, h: 0.3, fontSize: 11, color: C.accent6 });
  s.addNotes("Speaker: Marlen. Hello, we are the Unveil team. Last time we showed our progress report: a hybrid plan with stage gates and Scrum sprints. Today we explain Agile project management and show how we apply it to Unveil, a dating web service where you meet the person first. The six squares are our six two-week sprints: two are done, the third is running now.");
}

// ===================== 2. Why Agile =====================
{
  const s = content("// 01 · WHY AGILE", "We cannot know in advance if the idea works", "Introduction");
  const y = 1.55, h = 2.35, w = 2.8, g = 0.2;
  const items = [
    ["24-hour reveal", "Will people chat for a whole day before they see a photo? Only real users can tell us.", C.background2, C.text1],
    ["AI co-pilot", "How good the games and questions are, and the limits of the free AI tier, we learn only by building.", C.background2, C.text1],
    ["14 wk · 448 h · $0", "Time and budget are fixed. 4 people x 8 hours a week. The only thing that can move is scope.", C.accent1, C.text1],
  ];
  items.forEach(([head, body, fill, col], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, fill);
    txt(s, head, { x: x + 0.22, y: y + 0.25, w: w - 0.44, h: 0.5, fontSize: 20, bold: true, color: col });
    txt(s, body, { x: x + 0.22, y: y + 0.9, w: w - 0.44, h: 1.3, fontSize: 14, color: col });
  });
  txt(s, "A new product with fixed time and budget: we fix the dates and let the scope flex. This is what Agile is for.",
    { x: X, y: 4.2, w: W, h: 0.6, fontSize: 15, bold: true });
  s.addNotes("Speaker: Marlen. Why do we need Agile at all? Unveil is a new idea. We do not know if users will accept waiting 24 hours to see a photo, and we do not know how good our AI co-pilot will be until we build it. At the same time our time and budget are fixed: 14 weeks, 448 team-hours and zero dollars. When the product is uncertain but time and money are fixed, the safest approach is to deliver small working parts, get feedback early and let the scope change. That is Agile.");
}

// ===================== 3. Manifesto =====================
pres.addSection({ title: "Agile theory" });
{
  const s = content("// 02 · WHAT IS AGILE", "Four values of the Agile Manifesto", "Agile theory");
  const vals = [
    ["Individuals and interactions", "over processes and tools", "Us: a weekly stand-up and a team chat instead of long status reports."],
    ["Working software", "over comprehensive documentation", "Us: every sprint ends with a demo that really runs."],
    ["Customer collaboration", "over contract negotiation", "Us: the lecturer and test users see every Sprint Review."],
    ["Responding to change", "over following a plan", "Us: MoSCoW lets us drop a “Could” instead of moving a date."],
  ];
  const y = 1.5, h = 3.15, w = 2.05, g = 0.2;
  vals.forEach(([a, b, us], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, C.background2);
    txt(s, `0${i + 1}`, { x: x + 0.18, y: y + 0.2, w: 1, h: 0.3, fontSize: 12, bold: true, color: C.text2 });
    txt(s, a, { x: x + 0.18, y: y + 0.5, w: w - 0.36, h: 0.75, fontSize: 15, bold: true });
    txt(s, b, { x: x + 0.18, y: y + 1.25, w: w - 0.36, h: 0.55, fontSize: 12, color: C.text2 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.12, y: y + 1.9, w: w - 0.24, h: 1.12, rectRadius: 0.06, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `us-${i}` });
    txt(s, us, { x: x + 0.22, y: y + 1.97, w: w - 0.44, h: 1.0, fontSize: 12, valign: "middle" });
  });
  txt(s, "Agile Manifesto, 2001 · the items on the right still matter, we value the left more · plus 12 principles",
    { x: X, y: 4.78, w: W, h: 0.25, fontSize: 10, color: C.text2 });
  s.addNotes("Speaker: Marlen. Agile started with the Agile Manifesto in 2001. It has four values and twelve principles. Each value says we prefer the left side over the right side, but the right side still matters. Under each value you see how we use it on Unveil: short talks instead of reports, a working demo every sprint, the lecturer and test users giving feedback, and MoSCoW priorities that let us change scope without moving the deadline.");
}

// ===================== 4. Waterfall vs Agile =====================
{
  const s = content("// 03 · PREDICTIVE VS AGILE", "Waterfall vs Agile on our project", "Agile theory");
  const hl = (t) => ({ text: t, options: { bold: true } });
  table(s, [
    ["Aspect", "Waterfall: plan, then build", "Agile: build in small steps (Unveil)"],
    [hl("Planning"), "Whole project upfront", "Gates fixed, next sprint planned in detail"],
    [hl("Scope"), "Frozen in the charter", "Product Backlog in MoSCoW order"],
    [hl("Delivery"), "One release in week 14", "A working increment every 2 weeks"],
    [hl("Feedback"), "At the final demo", "At every Sprint Review"],
    [hl("Change"), "Formal change request", "PO reorders the backlog; new scope needs sponsor OK"],
    [hl("Progress"), "Earned value: SPI, CPI", "Earned value + velocity and burndown"],
  ], { y: 1.45, colW: [1.6, 3.0, 4.2], rowH: 0.36, fontSize: 13 });
  txt(s, "With pure Waterfall, the first real test of the 24-hour reveal would be in week 14. Too late to fix anything.",
    { x: X, y: 4.25, w: W, h: 0.6, fontSize: 15, bold: true });
  s.addNotes("Speaker: Marlen. Here is the difference between the classic predictive approach, Waterfall, and Agile, using our project. In Waterfall we plan everything at the start, freeze scope and release once at the end. In Agile we keep the big milestones but plan only the next sprint in detail, keep a prioritised backlog and release something that works every two weeks. The key risk for Unveil: in Waterfall nobody would try the 24-hour photo reveal until week 14, when it is too late to change it. Now Yeldos will explain Scrum.");
}

// ===================== 5. Scrum loop =====================
pres.addSection({ title: "Scrum and our team" });
{
  const s = content("// 04 · SCRUM", "One sprint is a two-week loop", "Scrum and our team");
  const ev = [
    ["DAY 1", "Sprint Planning", "45 min. Pick the Sprint Goal and the stories."],
    ["DAILY + WEEKLY", "Stand-up", "Chat update every day; 15-min call once a week."],
    ["DAY 10", "Sprint Review", "30 min. Live demo, feedback from the lecturer."],
    ["DAY 10", "Retrospective", "20 min. One thing we change next sprint."],
  ];
  const y = 1.5, h = 1.6, w = 1.85, g = 0.47;
  ev.forEach(([when, name, body], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, i === 0 ? C.accent1 : C.background2);
    txt(s, when, { x: x + 0.15, y: y + 0.15, w: w - 0.3, h: 0.25, fontSize: 10, bold: true, color: i === 0 ? C.text1 : C.text2 });
    txt(s, name, { x: x + 0.15, y: y + 0.42, w: w - 0.3, h: 0.35, fontSize: 15, bold: true });
    txt(s, body, { x: x + 0.15, y: y + 0.8, w: w - 0.3, h: 0.75, fontSize: 12 });
    if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + w + 0.1, y: y + h / 2 - 0.12, w: 0.27, h: 0.24, fill: { color: C.text1 }, line: { type: "none" }, objectName: `arrow-${i}` });
  });
  // roles / artifacts / events
  const cols = [
    ["3 roles", "Product Owner, Scrum Master, Developers"],
    ["5 events", "Sprint, Planning, Daily Scrum, Review, Retrospective"],
    ["3 artifacts", "Product Backlog, Sprint Backlog, Increment"],
  ];
  const w2 = 2.8, g2 = 0.2, y2 = 3.4;
  cols.forEach(([a, b], i) => {
    const x = X + i * (w2 + g2);
    card(s, x, y2, w2, 1.25, C.background1, { line: HEX.line });
    txt(s, a, { x: x + 0.2, y: y2 + 0.15, w: w2 - 0.4, h: 0.4, fontSize: 18, bold: true });
    txt(s, b, { x: x + 0.2, y: y2 + 0.58, w: w2 - 0.4, h: 0.6, fontSize: 12, color: C.text2 });
  });
  txt(s, "Source: The Scrum Guide, 2020", { x: X, y: 4.82, w: W, h: 0.25, fontSize: 10, color: C.text2 });
  s.addNotes("Speaker: Yeldos. The most popular Agile framework is Scrum. Work is split into sprints; ours are two weeks long. Each sprint starts with Sprint Planning where we choose a Sprint Goal. During the sprint we post a short daily update in our chat and have a 15-minute call every week. On the last day we show a live demo in the Sprint Review and then hold a Retrospective to agree one improvement. Scrum has three roles, five events and three artifacts: the Product Backlog, the Sprint Backlog and the Increment, the working result of the sprint.");
}

// ===================== 6. Team roles =====================
{
  const s = content("// 05 · OUR TEAM IN SCRUM", "Four people, three Scrum roles", "Scrum and our team");
  const y = 1.45, h = 2.75;
  card(s, X, y, 2.7, h, C.accent1);
  txt(s, "PRODUCT OWNER", { x: X + 0.2, y: y + 0.2, w: 2.3, h: 0.25, fontSize: 10, bold: true });
  txt(s, "Yeldos", { x: X + 0.2, y: y + 0.5, w: 2.3, h: 0.45, fontSize: 22, bold: true });
  txt(s, "Owns requirements and user stories, orders the backlog, talks to the sponsor.", { x: X + 0.2, y: y + 1.05, w: 2.3, h: 1.5, fontSize: 13 });

  const x2 = X + 2.9;
  card(s, x2, y, 2.7, h, C.background2);
  txt(s, "SCRUM MASTER", { x: x2 + 0.2, y: y + 0.2, w: 2.3, h: 0.25, fontSize: 10, bold: true, color: C.text2 });
  txt(s, "Marlen", { x: x2 + 0.2, y: y + 0.5, w: 2.3, h: 0.45, fontSize: 22, bold: true });
  txt(s, "Runs the events, removes blockers, keeps the WBS, CPM and risk register up to date.", { x: x2 + 0.2, y: y + 1.05, w: 2.3, h: 1.5, fontSize: 13 });

  const x3 = X + 5.8, w3 = W - 5.8;
  card(s, x3, y, w3, h, C.background1, { line: HEX.ink });
  txt(s, "DEVELOPERS · ALL FOUR", { x: x3 + 0.2, y: y + 0.2, w: w3 - 0.4, h: 0.25, fontSize: 10, bold: true, color: C.text2 });
  txt(s, [
    { text: "Diyor", options: { bold: true } }, { text: " · data model, API, matching", options: { breakLine: true } },
    { text: "Farkhad", options: { bold: true } }, { text: " · sign-up UI, tests, chat", options: { breakLine: true } },
    { text: "Marlen", options: { bold: true } }, { text: " · AI co-pilot", options: { breakLine: true } },
    { text: "Yeldos", options: { bold: true } }, { text: " · privacy, testing" },
  ], { x: x3 + 0.2, y: y + 0.55, w: w3 - 0.4, h: 2.0, fontSize: 13, paraSpaceAfter: 8 });
  txt(s, "In a team of four the PO and the Scrum Master also write code. Stakeholder: our lecturer as sponsor; test users give feedback.",
    { x: X, y: 4.4, w: W, h: 0.5, fontSize: 13, color: C.text2 });
  s.addNotes("Speaker: Yeldos. This is how the Scrum roles map onto our team. I am the Product Owner because I wrote the requirements and user stories: I decide what we build first and talk to the sponsor, our lecturer. Marlen is the Scrum Master: he already keeps the WBS, the critical path and the risk register, so he runs the events and removes blockers. All four of us are Developers: Diyor does the data model, API and matching algorithm, Farkhad the sign-up, onboarding tests and chat, Marlen the AI co-pilot, and I do privacy settings and testing.");
}

// ===================== 7. Product backlog =====================
{
  const s = content("// 06 · PRODUCT BACKLOG", "The backlog: user stories in MoSCoW order", "Scrum and our team");
  const pr = (t) => ({ text: t, options: { bold: true, color: t === "Must" ? C.text1 : C.text2 } });
  table(s, [
    ["#", "As a user, I want to…", "MoSCoW", "SP", "Sprint"],
    ["1", "Sign up and log in", pr("Must"), "3", "S3"],
    ["2", "Take personality and values tests", pr("Must"), "8", "S3"],
    ["3", "Get a match from the algorithm, no swipes", pr("Must"), "13", "S4"],
    ["4", "Chat with my match in real time", pr("Must"), "8", "S4"],
    ["5", "See their photo after 24 hours of chat", pr("Must"), "5", "S4"],
    ["6", "Edit my profile and see my test results", pr("Should"), "5", "S3"],
    ["7", "Play AI co-pilot games and questions in chat", pr("Should"), "8", "S5"],
    ["8", "Choose who sees my data (privacy settings)", pr("Should"), "5", "S5"],
    ["9", "Get an AI insight report after a week", pr("Could"), "5", "S6"],
  ], { y: 1.4, colW: [0.5, 5.0, 1.2, 0.9, 1.2], rowH: 0.3, fontSize: 12 });
  txt(s, "Total 60 story points (SP) · Won’t have this term: swipes, mobile app, video calls, payments",
    { x: X, y: 4.65, w: W, h: 0.3, fontSize: 11, color: C.text2 });
  s.addNotes("Speaker: Yeldos. The Product Backlog is our ordered to-do list. We took the MoSCoW scope from our charter and rewrote it as user stories: as a user, I want to... Must-have stories come first: sign-up, the onboarding tests, the matching algorithm, the chat and the 24-hour photo reveal. Then the Should and Could items. Each story has an estimate in story points; the whole backlog is 60 points. The Won't list protects us from scope creep. Now Diyor will show how this is split into sprints.");
}

// ===================== 8. Sprint plan =====================
pres.addSection({ title: "Agile in Unveil" });
{
  const s = content("// 07 · SPRINT PLAN", "Six sprints inside five stage gates", "Agile in Unveil");
  // top band
  const by = 1.4, bh = 0.4;
  card(s, X, by, 1.25, bh, C.background2);
  txt(s, "Wk 1–2 · Gate 1", { x: X + 0.1, y: by, w: 1.1, h: bh, fontSize: 10, valign: "middle", bold: true });
  card(s, X + 1.4, by, W - 1.4, bh, C.text1);
  txt(s, "Weeks 3–14 · six two-week Scrum sprints · 448 team-hours, $0", { x: X + 1.55, y: by, w: W - 1.7, h: bh, fontSize: 12, bold: true, color: C.background1, valign: "middle" });

  const sp = [
    ["S1", "wk 3–4", "Requirements, user stories, WBS", "Gate 2 ✓", "done"],
    ["S2", "wk 5–6", "Data model, API design, test design", "✓ done", "done"],
    ["S3", "wk 7–8", "Sign-up, onboarding tests, profiles · 16 SP", "Gate 3 · now", "now"],
    ["S4", "wk 9–10", "Matching, chat, 24-h reveal · 26 SP", "Gate 4", "plan"],
    ["S5", "wk 11–12", "AI co-pilot, privacy, testing · 13 SP", "Gate 5", "plan"],
    ["S6", "wk 13–14", "Insight report, fixes, final demo · 5 SP", "Final demo", "plan"],
  ];
  const y = 1.98, h = 2.25, g = 0.15, w = (W - 5 * g) / 6;
  sp.forEach(([id, wk, body, gate, st], i) => {
    const x = X + i * (w + g);
    const fill = st === "now" ? C.accent1 : st === "done" ? C.background2 : C.background1;
    card(s, x, y, w, h, fill, st === "plan" ? { line: HEX.line } : {});
    txt(s, id, { x: x + 0.12, y: y + 0.12, w: w - 0.24, h: 0.4, fontSize: 20, bold: true, color: st === "done" ? C.text2 : C.text1 });
    txt(s, wk, { x: x + 0.12, y: y + 0.55, w: w - 0.24, h: 0.25, fontSize: 11, bold: true });
    const [main, pts] = body.split(" · ");
    txt(s, main, { x: x + 0.12, y: y + 0.85, w: w - 0.24, h: 0.95, fontSize: 11 });
    if (pts) txt(s, pts, { x: x + 0.12, y: y + h - 0.4, w: w - 0.24, h: 0.28, fontSize: 12, bold: true });
    txt(s, gate, { x, y: y + h + 0.1, w, h: 0.3, fontSize: 11, bold: true, align: "center", color: st === "plan" ? C.text2 : C.text1 });
  });
  txt(s, "Each gate is checked at a Sprint Review. Sprint 4 is the heaviest: it carries the critical path (G matching, I chat).",
    { x: X, y: 4.72, w: W, h: 0.4, fontSize: 12, color: C.text2 });
  s.addNotes("Speaker: Diyor. Our 14 weeks start with two weeks for the charter, which is Gate 1. After that we run six two-week sprints. Sprints 1 and 2 are done: requirements, user stories, WBS, the data model and API design. Sprint 3 is running now: sign-up, onboarding tests and profiles, 16 story points. Sprint 4 brings the matching, the chat and the 24-hour photo reveal. Sprint 5 is the AI co-pilot, privacy and testing, and Sprint 6 is the insight report, fixes and the final demo. The stage gates from our plan are checked at the Sprint Reviews, so the two models work together.");
}

// ===================== 9. Kanban =====================
{
  const s = content("// 08 · KANBAN", "One board shows all the work", "Agile in Unveil");
  const cols = [
    ["TO DO · WK 8–10", C.background1, ["Matching algorithm · Diyor", "Chat + 24-h reveal · Farkhad", "Privacy settings · Yeldos"]],
    ["IN PROGRESS · WIP 2", C.accent1, ["Sign-up and test UI · Farkhad", "Profiles, test scoring · Diyor", "AI co-pilot prototype · Marlen"]],
    ["DONE", C.background2, ["Charter and roles · Marlen", "User stories · Yeldos", "Data model, API · Diyor"]],
  ];
  const y = 1.4, h = 2.4, g = 0.2, w = (W - 2 * g) / 3;
  cols.forEach(([head, fill, items], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, fill, fill === C.background1 ? { line: HEX.line } : {});
    txt(s, head, { x: x + 0.18, y: y + 0.15, w: w - 0.36, h: 0.25, fontSize: 10, bold: true });
    items.forEach((it, j) => {
      const iy = y + 0.5 + j * 0.6;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: iy, w: w - 0.3, h: 0.5, rectRadius: 0.05, fill: { color: C.background1 }, line: { color: HEX.line, width: 0.75 }, objectName: `ticket-${i}-${j}` });
      txt(s, it, { x: x + 0.25, y: iy, w: w - 0.5, h: 0.5, fontSize: 12, valign: "middle" });
    });
  });
  const rules = [
    ["WIP limit 2", "No one holds more than 2 cards. Finish before you start."],
    ["Pull, don’t push", "Take the next card from the top of To do."],
    ["Done = reviewed", "A card moves only after code review of its pull request."],
  ];
  const rw = (W - 2 * g) / 3;
  rules.forEach(([a, b], i) => {
    const x = X + i * (rw + g);
    txt(s, [{ text: a, options: { bold: true, breakLine: true } }, { text: b, options: { color: C.text2 } }],
      { x, y: 3.98, w: rw, h: 0.9, fontSize: 12 });
  });
  s.addNotes("Speaker: Diyor. Inside each sprint we use a Kanban board. It shows all the work in three columns: to do, in progress and done. This is our board at week 7. We have three simple rules: a work-in-progress limit of two cards per person, so we finish things before starting new ones; we pull the next card from the top of the list; and a card is done only after code review of its pull request. Scrum gives us the rhythm, Kanban makes the flow visible.");
}

// ===================== 10. Estimation & DoD =====================
{
  const s = content("// 09 · ESTIMATION AND DONE", "How we size work and when it counts as done", "Agile in Unveil");
  // left: planning poker
  const lw = 4.1, y = 1.45;
  card(s, X, y, lw, 3.3, C.background2);
  txt(s, "Planning poker", { x: X + 0.2, y: y + 0.18, w: lw - 0.4, h: 0.4, fontSize: 18, bold: true });
  const fib = ["1", "2", "3", "5", "8", "13"];
  fib.forEach((f, i) => {
    const cx = X + 0.2 + i * 0.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: y + 0.7, w: 0.5, h: 0.7, rectRadius: 0.06, fill: { color: f === "13" ? C.accent1 : C.background1 }, line: { color: HEX.ink, width: 1 }, objectName: `poker-${f}` });
    txt(s, f, { x: cx, y: y + 0.7, w: 0.5, h: 0.7, fontSize: 16, bold: true, align: "center", valign: "middle" });
  });
  txt(s, "Everyone shows a card at once; big gaps are discussed. Matching got 13 points: it has the most unknowns. Story points measure effort and risk, not hours.",
    { x: X + 0.2, y: y + 1.6, w: lw - 0.4, h: 1.5, fontSize: 13 });

  // right: DoD
  const rx = X + lw + 0.3, rw = W - lw - 0.3;
  card(s, rx, y, rw, 3.3, C.background1, { line: HEX.ink });
  txt(s, "Definition of Done", { x: rx + 0.2, y: y + 0.18, w: rw - 0.4, h: 0.4, fontSize: 18, bold: true });
  const dod = [
    "Code reviewed and merged via pull request",
    "Meets the story’s acceptance criteria",
    "Tested on test accounts only, no real photos",
    "Runs in the demo build",
    "Hours logged against the task (for EVM)",
    "Board and backlog updated",
  ];
  dod.forEach((d, i) => {
    const iy = y + 0.72 + i * 0.42;
    s.addShape(pres.shapes.OVAL, { x: rx + 0.2, y: iy + 0.05, w: 0.24, h: 0.24, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `dod-dot-${i}` });
    txt(s, "✓", { x: rx + 0.2, y: iy + 0.05, w: 0.24, h: 0.24, fontSize: 10, bold: true, align: "center", valign: "middle" });
    txt(s, d, { x: rx + 0.55, y: iy, w: rw - 0.75, h: 0.34, fontSize: 13, valign: "middle" });
  });
  s.addNotes("Speaker: Diyor. To estimate the backlog we use planning poker with Fibonacci numbers. Everyone shows a card at the same time, and if the numbers are far apart we discuss why. For example the matching algorithm got 13 points because it has the most unknowns. On the right is our Definition of Done: a story counts as done only when the code is reviewed and merged, it meets the acceptance criteria, it is tested on test accounts only, it runs in the demo build, the hours are logged and the board is updated.");
}

// ===================== 11. Metrics =====================
{
  const s = content("// 10 · METRICS", "What we track every sprint", "Agile in Unveil");
  s.addChart(pres.charts.BAR, [{ name: "Story points", labels: ["S3", "S4", "S5", "S6"], values: [16, 26, 13, 5] }], {
    x: X, y: 1.35, w: 4.6, h: 3.1, barDir: "col",
    chartColors: [HEX.yellow], barGapWidthPct: 60,
    showTitle: true, title: "Story points planned per sprint", titleFontSize: 12, titleColor: HEX.ink, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.ink, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt", dataLabelFontBold: true,
    catAxisLabelColor: HEX.muted, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLineShow: true, catAxisLineColor: HEX.line, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 30,
  });
  const tx = X + 4.9, tw = W - 4.9;
  const tiles = [
    ["224 / 448 h", "Team-hours used at week 7: exactly half", C.background2],
    ["SPI 1.00 · CPI 1.00", "Earned value stays our cost and time check", C.background2],
    ["Alarm < 90%", "If a sprint ends under 90% of planned points, the PO cuts a Could or Should", C.accent1],
  ];
  tiles.forEach(([a, b, fill], i) => {
    const y = 1.35 + i * 1.07;
    card(s, tx, y, tw, 0.95, fill);
    txt(s, a, { x: tx + 0.2, y: y + 0.1, w: tw - 0.4, h: 0.38, fontSize: 18, bold: true });
    txt(s, b, { x: tx + 0.2, y: y + 0.48, w: tw - 0.4, h: 0.42, fontSize: 11 });
  });
  txt(s, "Sprint 4 holds 26 of 60 points. Fix: Marlen and Yeldos move from tasks with float (H, J) to help.",
    { x: X, y: 4.6, w: W, h: 0.4, fontSize: 12, bold: true });
  s.addNotes("Speaker: Farkhad. What do we measure? The chart shows the story points planned for each of the remaining sprints. Velocity, the points we finish per sprint, and a burndown chart of the remaining points show if we are on track. We keep earned value too: at week 7 we used 224 of 448 team-hours, and both SPI and CPI are 1.00. Our alarm rule: if a sprint finishes under 90 percent of its planned points, the Product Owner cuts a Could or Should item. The chart shows a real problem: Sprint 4 holds 26 of 60 points and it is on the critical path. So Marlen and Yeldos will move from tasks that have float to help with matching and chat.");
}

// ===================== 12. Hybrid =====================
pres.addSection({ title: "Conclusion" });
{
  const s = content("// 11 · HYBRID MODEL", "Keep the gates, run sprints inside them", "Conclusion");
  const hl = (t) => ({ text: t, options: { bold: true } });
  table(s, [
    ["Plan-driven tool we keep", "Agile partner we add"],
    [hl("Stage gates (wk 2, 4, 7, 10, 12)"), "Sprint Review where the gate is checked"],
    [hl("WBS: all the work"), "Product Backlog of user stories"],
    [hl("Critical path A→B→D→E→G→I→K→L"), "Sprint order: critical stories first"],
    [hl("Earned value: SPI, CPI"), "Velocity and burndown"],
    [hl("Risk register"), "Blockers at stand-up, lessons at the retro"],
    [hl("Change control: sponsor approval"), "MoSCoW: the PO reorders inside the fixed scope"],
  ], { y: 1.4, colW: [4.4, 4.4], rowH: 0.37, fontSize: 13 });
  txt(s, "Time (14 weeks) and budget (448 h, $0) stay fixed. What we build can change inside them.",
    { x: X, y: 4.35, w: W, h: 0.5, fontSize: 15, bold: true });
  s.addNotes("Speaker: Farkhad. We do not throw away our plan-driven tools. We use a hybrid model: each classic tool gets an Agile partner. The stage gates are checked at Sprint Reviews, the WBS becomes the Product Backlog, the critical path decides the order of sprints, earned value works together with velocity and burndown, and the risk register is fed by stand-ups and retros. Time and budget stay fixed; only the scope can move inside them.");
}

// ===================== 13. Risks =====================
{
  const s = content("// 12 · RISKS", "Agile has costs too, each with a fix", "Conclusion");
  const hl = (t) => ({ text: t, options: { bold: true } });
  table(s, [
    ["Risk (score from register)", "Our Agile response"],
    [hl("Team overload, exams (16, High)"), "Plan capacity per sprint; core features before exams; WIP limit 2"],
    [hl("Scope creep (15, High)"), "Won’t list; new ideas go to the bottom of the backlog"],
    [hl("Privacy of photos and tests (10)"), "In the Definition of Done: test accounts only, photo locked 24 h"],
    [hl("AI quality and cost (9)"), "Prototype with mock prompts in S3; AI co-pilot stays a Should"],
    [hl("Estimates are new to us"), "Planning poker; compare estimate and actual at every retro"],
  ], { y: 1.4, colW: [3.5, 5.3], rowH: 0.42, fontSize: 13 });
  txt(s, "Benefit: the sponsor and test users see a working part of Unveil every two weeks.",
    { x: X, y: 4.35, w: W, h: 0.5, fontSize: 15, bold: true });
  s.addNotes("Speaker: Farkhad. Agile is not free. It needs discipline and meetings, and there is a risk of scope creep. We linked the Agile practices to the risks in our risk register. Team overload during exams is handled by planning capacity per sprint and the WIP limit. Scope creep is handled by the Won't list and the backlog. Privacy is part of our Definition of Done. AI quality is tested early with a prototype. And because estimation is new to us, we compare estimates with actual results at every retrospective. The main benefit: every two weeks there is a working part of Unveil to show.");
}

// ===================== 14. Next step =====================
{
  const s = pres.addSlide({ masterName: "Yellow", sectionTitle: "Conclusion" });
  s.addText("// 13 · NEXT STEP", { placeholder: "kicker" });
  s.addText("Sprint 4: the heart of Unveil", { placeholder: "title" });
  const items = [
    ["WEEK 9 · PLANNING", "Goal: two test users get matched, chat, and see the photo after 24 hours"],
    ["SPRINT BACKLOG", "Matching (13) + chat (8) + 24-h reveal (5) = 26 SP"],
    ["WEEK 10 · REVIEW", "Live demo at Gate 4, then the retrospective"],
  ];
  const y = 2.0, h = 1.75, g = 0.2, w = (W - 2 * g) / 3;
  items.forEach(([a, b], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, C.text1);
    txt(s, a, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.25, fontSize: 10, bold: true, color: C.accent1 });
    txt(s, b, { x: x + 0.2, y: y + 0.55, w: w - 0.4, h: 1.1, fontSize: 14, color: C.background1 });
  });
  txt(s, "Questions?", { x: X, y: 4.15, w: 4, h: 0.5, fontSize: 22, bold: true });
  s.addNotes("Speaker: Farkhad. Our next step is Sprint 4, the heart of Unveil. The Sprint Goal: two test users get matched by the algorithm, chat with each other and see each other's photo after 24 hours. That is 26 story points on the critical path, and we will show it as a live demo at Gate 4 in week 10. Thank you. We are happy to answer your questions.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
