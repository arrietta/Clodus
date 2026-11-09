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
  s.addNotes("Speaker: Marlen. Hello, we are the Unveil team. Today: Marlen covers Agile basics, Yeldos Scrum and the backlog, Diyor our sprints and board, Farkhad metrics, risks and the next step. Last time we showed our progress report: a hybrid plan with stage gates and Scrum sprints. Today we explain Agile project management and show how we apply it to Unveil, a dating web service where you meet the person first. The six squares are our six two-week sprints: two are done, the third is running now.");
}

// ===================== 3. Manifesto =====================
pres.addSection({ title: "Agile theory" });
{
  const s = content("// 01 · WHAT IS AGILE", "Four values of the Agile Manifesto", "Agile theory");
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
  const s = content("// 02 · WHY AGILE FOR UNVEIL", "Waterfall vs Agile on our project", "Agile theory");
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
  txt(s, "Will users wait 24 hours for a photo? Is a free AI good enough? With Waterfall we learn it in week 14, too late. With Agile, every two weeks.",
    { x: X, y: 4.25, w: W, h: 0.6, fontSize: 15, bold: true });
  s.addNotes("Speaker: Marlen. Why does Unveil need Agile? It is a new idea: we do not know if users will wait 24 hours to see a photo, or how good a free AI co-pilot will be, while our time and budget are fixed: 14 weeks, 448 team-hours and zero dollars. Here is the difference between the classic predictive approach, Waterfall, and Agile, using our project. In Waterfall we plan everything at the start, freeze scope and release once at the end. In Agile we keep the big milestones but plan only the next sprint in detail, keep a prioritised backlog and release something that works every two weeks. The key risk for Unveil: in Waterfall nobody would try the 24-hour photo reveal until week 14, when it is too late to change it. Now Yeldos will explain Scrum.");
}

// ===================== Scrum + team =====================
pres.addSection({ title: "Scrum and our team" });
{
  const s = content("// 03 · SCRUM ON OUR TEAM", "A two-week loop and three roles", "Scrum and our team");
  const ev = [
    ["DAY 1", "Sprint Planning", "45 min. Pick the Sprint Goal and the stories."],
    ["DAILY + WEEKLY", "Stand-up", "Chat update every day, 15-min call weekly."],
    ["DAY 10", "Sprint Review", "30 min. Live demo, feedback from the lecturer."],
    ["DAY 10", "Retrospective", "20 min. One thing we change next sprint."],
  ];
  const y = 1.4, h = 1.35, w = 1.85, g = 0.47;
  ev.forEach(([when, name, body], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, C.background2);
    txt(s, when, { x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.22, fontSize: 10, bold: true, color: C.text2 });
    txt(s, name, { x: x + 0.15, y: y + 0.35, w: w - 0.3, h: 0.3, fontSize: 14, bold: true });
    txt(s, body, { x: x + 0.15, y: y + 0.68, w: w - 0.3, h: 0.62, fontSize: 11 });
    if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + w + 0.1, y: y + h / 2 - 0.12, w: 0.27, h: 0.24, fill: { color: C.text1 }, line: { type: "none" }, objectName: `arrow-${i}` });
  });
  const ry = 2.95, rh = 1.8, rw = (W - 0.4) / 3;
  const roles = [
    ["PRODUCT OWNER", "Yeldos", "Owns user stories, orders the backlog, talks to the sponsor.", C.accent1],
    ["SCRUM MASTER", "Marlen", "Runs the events, removes blockers, keeps WBS, CPM and risks.", C.background2],
  ];
  roles.forEach(([lab, name, body, fill], i) => {
    const x = X + i * (rw + 0.2);
    card(s, x, ry, rw, rh, fill);
    txt(s, lab, { x: x + 0.2, y: ry + 0.15, w: rw - 0.4, h: 0.22, fontSize: 10, bold: true, color: i === 0 ? C.text1 : C.text2 });
    txt(s, name, { x: x + 0.2, y: ry + 0.4, w: rw - 0.4, h: 0.4, fontSize: 20, bold: true });
    txt(s, body, { x: x + 0.2, y: ry + 0.9, w: rw - 0.4, h: 0.8, fontSize: 12 });
  });
  const dx = X + 2 * (rw + 0.2);
  card(s, dx, ry, rw, rh, C.background1, { line: HEX.ink });
  txt(s, "DEVELOPERS · ALL FOUR", { x: dx + 0.2, y: ry + 0.15, w: rw - 0.4, h: 0.22, fontSize: 10, bold: true, color: C.text2 });
  txt(s, [
    { text: "Diyor", options: { bold: true } }, { text: " · API, matching", options: { breakLine: true } },
    { text: "Farkhad", options: { bold: true } }, { text: " · sign-up, tests, chat", options: { breakLine: true } },
    { text: "Marlen", options: { bold: true } }, { text: " · AI co-pilot", options: { breakLine: true } },
    { text: "Yeldos", options: { bold: true } }, { text: " · privacy, testing" },
  ], { x: dx + 0.2, y: ry + 0.45, w: rw - 0.4, h: 1.3, fontSize: 12, paraSpaceAfter: 4 });
  txt(s, "Scrum: 3 roles, 5 events, 3 artifacts (Product Backlog, Sprint Backlog, Increment) · Source: The Scrum Guide, 2020",
    { x: X, y: 4.85, w: W, h: 0.25, fontSize: 10, color: C.text2 });
  s.addNotes("Speaker: Yeldos. The most popular Agile framework is Scrum. Work is split into sprints; ours are two weeks long. Each sprint starts with Sprint Planning, where we choose a Sprint Goal. During the sprint we post a short daily update in our chat and have a 15-minute call every week. On the last day we show a live demo in the Sprint Review and then hold a Retrospective to agree one improvement. Scrum has three roles. I am the Product Owner, because I wrote the requirements and user stories: I decide what we build first and talk to the sponsor, our lecturer. Marlen is the Scrum Master: he already keeps the WBS, critical path and risk register, so he runs the events and removes blockers. All four of us are Developers: Diyor does the API and matching, Farkhad sign-up, tests and chat, Marlen the AI co-pilot, and I do privacy and testing.");
}

// ===================== 7. Product backlog =====================
{
  const s = content("// 04 · PRODUCT BACKLOG", "The backlog: user stories in MoSCoW order", "Scrum and our team");
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
  txt(s, "60 story points (SP), sized by planning poker (1, 2, 3, 5, 8, 13) · Won’t have this term: swipes, mobile app, video calls, payments",
    { x: X, y: 4.65, w: W, h: 0.3, fontSize: 11, color: C.text2 });
  s.addNotes("Speaker: Yeldos. The Product Backlog is our ordered to-do list. We took the MoSCoW scope from our charter and rewrote it as user stories: as a user, I want to... Must-have stories come first: sign-up, the onboarding tests, the matching algorithm, the chat and the 24-hour photo reveal. Then the Should and Could items. Each story has an estimate in story points; we estimate them with planning poker: everyone shows a Fibonacci card at once, and big differences are discussed. The matching algorithm got 13 because it has the most unknowns. The whole backlog is 60 points. The Won't list protects us from scope creep. Now Diyor will show how this is split into sprints.");
}

// ===================== 8. Sprint plan =====================
pres.addSection({ title: "Agile in Unveil" });
{
  const s = content("// 05 · SPRINT PLAN", "Six sprints inside five stage gates", "Agile in Unveil");
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

// ===================== Kanban + DoD =====================
{
  const s = content("// 06 · KANBAN AND DONE", "One board, one rule for “done”", "Agile in Unveil");
  const cols = [
    ["TO DO", C.background1, ["Matching · Diyor", "Chat + reveal · Farkhad", "Privacy · Yeldos"]],
    ["IN PROGRESS", C.accent1, ["Sign-up, test UI · Farkhad", "Profiles, scoring · Diyor", "AI co-pilot · Marlen"]],
    ["DONE", C.background2, ["Charter · Marlen", "User stories · Yeldos", "Data model, API · Diyor"]],
  ];
  const BW = 5.6, y = 1.4, h = 2.45, g = 0.15, w = (BW - 2 * g) / 3;
  cols.forEach(([head, fill, items], i) => {
    const x = X + i * (w + g);
    card(s, x, y, w, h, fill, fill === C.background1 ? { line: HEX.line } : {});
    txt(s, head, { x: x + 0.12, y: y + 0.14, w: w - 0.24, h: 0.22, fontSize: 10, bold: true });
    items.forEach((it, j) => {
      const iy = y + 0.48 + j * 0.64;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: iy, w: w - 0.2, h: 0.54, rectRadius: 0.05, fill: { color: C.background1 }, line: { color: HEX.line, width: 0.75 }, objectName: `ticket-${i}-${j}` });
      txt(s, it, { x: x + 0.18, y: iy, w: w - 0.36, h: 0.54, fontSize: 11, valign: "middle" });
    });
  });
  txt(s, [
    { text: "WIP limit 2", options: { bold: true } }, { text: " per person: finish before you start. ", options: {} },
    { text: "Pull", options: { bold: true } }, { text: " the next card from the top of To do. Board at week 7." },
  ], { x: X, y: 4.05, w: BW, h: 0.7, fontSize: 12 });

  const rx = X + BW + 0.25, rw = W - BW - 0.25, rh = 3.35;
  card(s, rx, y, rw, rh, C.background1, { line: HEX.ink });
  txt(s, "Definition of Done", { x: rx + 0.2, y: y + 0.15, w: rw - 0.4, h: 0.35, fontSize: 16, bold: true });
  const dod = ["Reviewed, merged by PR", "Meets acceptance criteria", "Test accounts only", "Runs in the demo build", "Hours logged (for EVM)", "Board and backlog updated"];
  dod.forEach((d, i) => {
    const iy = y + 0.65 + i * 0.44;
    s.addShape(pres.shapes.OVAL, { x: rx + 0.2, y: iy + 0.06, w: 0.24, h: 0.24, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `dod-dot-${i}` });
    txt(s, "✓", { x: rx + 0.2, y: iy + 0.06, w: 0.24, h: 0.24, fontSize: 10, bold: true, align: "center", valign: "middle" });
    txt(s, d, { x: rx + 0.55, y: iy, w: rw - 0.7, h: 0.36, fontSize: 12, valign: "middle" });
  });
  s.addNotes("Speaker: Diyor. Inside each sprint we use a Kanban board with three columns: to do, in progress and done. This is our board at week 7. Two rules keep the flow: a work-in-progress limit of two cards per person, so we finish things before starting new ones, and we pull the next card from the top of the list. A card moves to Done only when it meets our Definition of Done: the code is reviewed and merged through a pull request, it meets the acceptance criteria, it is tested on test accounts only, it runs in the demo build, the hours are logged for earned value, and the board is updated.");
}

// ===================== 11. Metrics =====================
{
  const s = content("// 07 · METRICS", "What we track every sprint", "Agile in Unveil");
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
  s.addNotes("Speaker: Diyor. What do we measure? The chart shows the story points planned for each of the remaining sprints. Velocity, the points we finish per sprint, and a burndown chart of the remaining points show if we are on track. We keep earned value too: at week 7 we used 224 of 448 team-hours, and both SPI and CPI are 1.00. Our alarm rule: if a sprint finishes under 90 percent of its planned points, the Product Owner cuts a Could or Should item. The chart shows a real problem: Sprint 4 holds 26 of 60 points and it is on the critical path. So Marlen and Yeldos will move from tasks that have float to help with matching and chat.");
}

// ===================== 12. Hybrid =====================
pres.addSection({ title: "Conclusion" });
{
  const s = content("// 08 · HYBRID MODEL", "Keep the gates, run sprints inside them", "Conclusion");
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
  const s = content("// 09 · RISKS", "Agile has costs too, each with a fix", "Conclusion");
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
  s.addText("// 10 · NEXT STEP", { placeholder: "kicker" });
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
