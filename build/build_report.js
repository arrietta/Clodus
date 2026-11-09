// Builds Unveil_Agile_Report.docx
// Run: NODE_PATH=<dir with docx> node build/build_report.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, LevelFormat, BorderStyle, Footer, PageNumber,
} = require("docx");

const OUT = path.join(__dirname, "..", "Unveil_Agile_Report.docx");
const FONT = "Arial";
const YEL = "F2C318", GRAY = "EFEFEF", LINE = "BFBFBF";
const PAGE_W = 11906, MARGIN = 1134, TW = PAGE_W - 2 * MARGIN; // A4, 2 cm margins

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 300 }, ...opts,
  children: (Array.isArray(text) ? text : [text]).map((t) => (typeof t === "string" ? new TextRun(t) : t)),
});
const B = (t) => new TextRun({ text: t, bold: true });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const bullet = (parts) => P(parts, { numbering: { reference: "bul", level: 0 }, spacing: { after: 60, line: 288 } });
const num = (parts) => P(parts, { numbering: { reference: "num", level: 0 }, spacing: { after: 60, line: 288 } });

function table(rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      children: r.map((cell, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA },
        borders,
        shading: ri === 0 ? { type: ShadingType.CLEAR, color: "auto", fill: GRAY } : undefined,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: cell, bold: ri === 0, size: 19 })] })],
      })),
    })),
  });
}
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });
const caption = (t) => P([new TextRun({ text: t, italics: true, size: 18, color: "5F5F5F" })], { spacing: { before: 60, after: 200 } });

const children = [
  // ---------- title block ----------
  new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "IT PROJECT MANAGEMENT · AGILE REPORT", bold: true, size: 18, color: "5F5F5F" })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("Agile Project Management, Applied to Unveil")] }),
  P([new TextRun({ text: "Unveil: a dating web service where you meet the person first", size: 24 })], { spacing: { after: 200 } }),
  P([B("Team: "), "Marlen Kalbayev, Diyor Khabibullayev, Yeldos Bakayev, Farkhad Amanbayev"], { spacing: { after: 40 } }),
  P([B("Course: "), "IT Project Management · ", B("Lecturer: "), "Muhammad Khalil Shahid"], { spacing: { after: 40 } }),
  P([B("Status of the project: "), "week 7 of 14, Sprint 3 of 6, on track (SPI 1.00, CPI 1.00)"], {
    spacing: { after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: YEL, space: 8 } },
  }),

  // ---------- 1 ----------
  H1("1. Introduction"),
  P("This report explains Agile project management and shows how our group applies it to our course project, Unveil. Unveil is a dating web service without swipes: new users take personality and values tests, an algorithm with AI help picks a match, and the match’s photo unlocks only after 24 hours of chatting. An AI co-pilot offers games and questions in the chat to help people learn about each other."),
  P("The project runs for 14 weeks with a team of four students, 8 hours per person per week (448 team-hours in total) and a budget of $0, because we use free tiers and open-source tools. Our progress report already described a hybrid plan: fixed stage gates plus two-week Scrum sprints. In this report we describe the Agile part in detail: the theory behind it, the Scrum and Kanban practices we use, our backlog and sprint plan, and the metrics and risks."),

  // ---------- 2 ----------
  H1("2. What is Agile"),
  P("Agile is an approach to managing projects in which work is delivered in small, working pieces, and the plan is updated often based on feedback. Instead of defining all requirements at the start and delivering one large result at the end, an Agile team works in short cycles (iterations or sprints), shows a working result after each cycle, and decides what to do next."),
  P("The approach was formalised in the Manifesto for Agile Software Development (2001). It states four values. In each, the items on the left are valued more, but the items on the right still matter:"),
  num([B("Individuals and interactions"), " over processes and tools. For us: a weekly stand-up call and a team chat instead of long status reports."]),
  num([B("Working software"), " over comprehensive documentation. For us: every sprint ends with a demo that really runs."]),
  num([B("Customer collaboration"), " over contract negotiation. For us: the lecturer (our sponsor) and test users see every Sprint Review."]),
  num([B("Responding to change"), " over following a plan. For us: MoSCoW priorities let us drop a “Could have” item instead of moving the deadline."]),
  P("The Manifesto also lists 12 principles. The most important for our project are: deliver working software frequently (every two weeks); welcome changing requirements, even late; keep a sustainable pace (important during exams); simplicity, maximising the amount of work not done; and regular reflection, where the team stops at intervals to improve how it works.", { spacing: { before: 120, after: 120, line: 300 } }),

  // ---------- 3 ----------
  H1("3. Predictive (Waterfall) vs Agile"),
  P("Predictive project management, often called Waterfall, plans the whole project upfront and moves through phases in order: requirements, design, build, test and release. It works well when requirements are clear and stable. Agile is better when the product is new and uncertain. The table compares the two using our own project."),
  table([
    ["Aspect", "Waterfall: plan, then build", "Agile: build in small steps (Unveil)"],
    ["Planning", "Whole project planned upfront", "Stage gates fixed; next sprint planned in detail"],
    ["Scope", "Frozen in the charter", "Product Backlog ordered by MoSCoW"],
    ["Delivery", "One release in week 14", "A working increment every 2 weeks"],
    ["Feedback", "At the final demo", "At every Sprint Review"],
    ["Change", "Formal change request", "PO reorders the backlog; new scope needs sponsor approval"],
    ["Progress", "Earned value (SPI, CPI)", "Earned value plus velocity and burndown"],
  ], [1800, 3519, 4319]),
  caption("Table 1. Waterfall and Agile compared on the Unveil project"),
  P("For Unveil the main argument is risk. With pure Waterfall, the first real test of the 24-hour photo reveal and of the AI co-pilot would happen in week 14, when it is too late to change anything. With Agile, these ideas are tested by users within the sprint in which they are built."),

  // ---------- 4 ----------
  H1("4. Agile frameworks: Scrum and Kanban"),
  H2("4.1 Scrum"),
  P("Scrum is the most widely used Agile framework (The Scrum Guide, 2020). Work happens in sprints of fixed length, at most one month; ours are two weeks. Scrum defines three roles, five events and three artifacts."),
  bullet([B("Roles. "), "The Product Owner decides what is built and in what order and owns the Product Backlog. The Scrum Master helps the team follow Scrum, runs the events and removes blockers. The Developers build the increment."]),
  bullet([B("Events. "), "The Sprint itself; Sprint Planning (choose the Sprint Goal and the items); the Daily Scrum (a short sync: done, next, blocked); the Sprint Review (demo of the increment and feedback); and the Sprint Retrospective (how to work better next sprint)."]),
  bullet([B("Artifacts. "), "The Product Backlog (ordered list of everything the product needs, with the Product Goal), the Sprint Backlog (items chosen for this sprint, with the Sprint Goal) and the Increment (the usable result that meets the Definition of Done)."]),
  H2("4.2 Kanban"),
  P("Kanban is a method for visualising and limiting work in progress. Tasks are cards on a board with columns such as To do, In progress and Done. The key rule is a work-in-progress (WIP) limit: a person or column may hold only a few cards at once, so the team finishes work before starting new work. Kanban has no fixed iterations, so it combines well with Scrum: Scrum gives the rhythm, Kanban makes the flow of work visible."),

  // ---------- 5 ----------
  H1("5. Applying Agile to Unveil"),
  H2("5.1 Why Unveil needs Agile"),
  P("Unveil combines a fixed frame with an uncertain product. Time (14 weeks) and budget (448 team-hours, $0) cannot change. But we do not know in advance whether users will accept waiting 24 hours to see a photo, or how good the AI co-pilot games will be on a free AI tier. Agile handles this by fixing time and cost and letting scope flex: the most valuable features are built first, and lower-priority features are cut if time runs out."),
  H2("5.2 Scrum roles in our team"),
  table([
    ["Role", "Person", "Responsibilities on Unveil"],
    ["Product Owner", "Yeldos Bakayev", "Owns requirements and user stories, orders the backlog, talks to the sponsor; also works on privacy settings and testing"],
    ["Scrum Master", "Marlen Kalbayev", "Runs the events, removes blockers, keeps the WBS, critical path and risk register; also builds the AI co-pilot"],
    ["Developer", "Diyor Khabibullayev", "Data model, API design, user profiles and test scoring, matching algorithm"],
    ["Developer", "Farkhad Amanbayev", "Sign-up and onboarding test UI, chat with the 24-hour photo reveal"],
    ["Sponsor / stakeholder", "Muhammad Khalil Shahid", "Approves scope changes, gives feedback at Sprint Reviews and stage gates"],
  ], [1900, 2300, 5438]),
  caption("Table 2. Scrum roles mapped to the Unveil team. In a team of four, the PO and Scrum Master also develop."),
  H2("5.3 Product Backlog"),
  P("We rewrote the MoSCoW scope from our charter as user stories in the form “As a user, I want to…”. Each story is estimated in story points (SP) and assigned to a sprint. The whole backlog is 60 story points."),
  table([
    ["#", "As a user, I want to…", "MoSCoW", "SP", "Sprint"],
    ["1", "Sign up and log in", "Must", "3", "S3"],
    ["2", "Take personality and values tests", "Must", "8", "S3"],
    ["3", "Get a match from the algorithm, with no swipes", "Must", "13", "S4"],
    ["4", "Chat with my match in real time", "Must", "8", "S4"],
    ["5", "See my match’s photo after 24 hours of chat", "Must", "5", "S4"],
    ["6", "Edit my profile and see my test results", "Should", "5", "S3"],
    ["7", "Play AI co-pilot games and questions in the chat", "Should", "8", "S5"],
    ["8", "Choose who sees my data (privacy settings)", "Should", "5", "S5"],
    ["9", "Get an AI insight report after a week of chatting", "Could", "5", "S6"],
  ], [600, 5438, 1300, 1000, 1300]),
  caption("Table 3. Product Backlog. Won’t have this term: swipes, mobile app, video calls, payments."),
  H2("5.4 Sprint plan"),
  P("Weeks 1–2 were used for the Project Charter (Gate 1). After that we run six two-week sprints. The stage gates from our plan are checked during Sprint Reviews, so the predictive milestones and the Agile cycle work together."),
  table([
    ["Sprint", "Weeks", "Sprint Goal", "SP", "Gate / status"],
    ["S1", "3–4", "Requirements, user stories, WBS, schedule", "–", "Gate 2 · done"],
    ["S2", "5–6", "Data model, API design, onboarding test design", "–", "Done"],
    ["S3", "7–8", "Sign-up, onboarding tests, profiles and scoring", "16", "Gate 3 · in progress"],
    ["S4", "9–10", "Matching algorithm, chat, 24-hour photo reveal", "26", "Gate 4"],
    ["S5", "11–12", "AI co-pilot games, privacy settings, testing", "13", "Gate 5"],
    ["S6", "13–14", "AI insight report, bug fixes, final demo", "5", "Final demo"],
  ], [900, 1000, 4638, 800, 2300]),
  caption("Table 4. Sprint plan inside the stage gates. Sprints 1 and 2 delivered analysis and design work, not user stories."),
  H2("5.5 Sprint events and Kanban board"),
  P("Each sprint follows the same loop. On day 1 we hold Sprint Planning (45 minutes) and agree the Sprint Goal. During the sprint each member posts a daily update in the team chat (done, next, blocked) and we hold a 15-minute stand-up call once a week. On day 10 we hold the Sprint Review (30 minutes, live demo to the lecturer) and the Retrospective (20 minutes, one concrete improvement for the next sprint)."),
  P("Work inside the sprint is tracked on a Kanban board with three columns. At week 7 the board shows: Done – charter and roles (Marlen), user stories (Yeldos), data model and API (Diyor); In progress – sign-up and test UI (Farkhad), profiles and test scoring (Diyor), AI co-pilot prototype (Marlen); To do – matching algorithm (Diyor), chat with 24-hour reveal (Farkhad), privacy settings (Yeldos). We use three rules:"),
  bullet([B("WIP limit of 2. "), "Nobody holds more than two cards at once."]),
  bullet([B("Pull, don’t push. "), "A free member takes the next card from the top of To do."]),
  bullet([B("Done means reviewed. "), "A card moves to Done only after code review of its pull request."]),
  H2("5.6 Estimation and Definition of Done"),
  P("We estimate stories with planning poker using Fibonacci numbers (1, 2, 3, 5, 8, 13). Everyone shows a card at the same time; if the numbers are far apart, the highest and lowest voters explain their reasons and we vote again. Story points measure effort and uncertainty, not hours. For example, the matching algorithm got 13 points because it has the most unknowns."),
  P("A story counts as done only when it meets our Definition of Done:"),
  bullet("the code is reviewed and merged through a pull request;"),
  bullet("it meets the acceptance criteria of the story;"),
  bullet("it is tested on test accounts only, with no real photos;"),
  bullet("it runs in the demo build;"),
  bullet("the hours are logged against the task (used for earned value);"),
  bullet("the Kanban board and the backlog are updated."),
  H2("5.7 Metrics"),
  P("We track Agile metrics next to the earned value metrics from our progress report:"),
  bullet([B("Velocity: "), "story points finished per sprint. Planned load: S3 16, S4 26, S5 13, S6 5 points."]),
  bullet([B("Burndown: "), "remaining story points after each sprint, compared with an ideal straight line from 60 to 0."]),
  bullet([B("Earned value: "), "at week 7, PV = EV = AC = 224 team-hours of 448, so SPI = 1.00 and CPI = 1.00."]),
  bullet([B("Alarm rule: "), "if a sprint finishes under 90% of its planned points (or SPI/CPI falls below 0.90), the Product Owner removes a Could or Should item at the next planning."]),
  P("The plan shows one clear issue: Sprint 4 holds 26 of 60 points and contains the critical-path tasks G (matching) and I (chat with reveal). Our response is to move Marlen and Yeldos from tasks with float (H, AI co-pilot, 1 week; J, privacy, 2 weeks) to help on matching and chat during Sprint 4.", { spacing: { before: 120, after: 120, line: 300 } }),

  // ---------- 6 ----------
  H1("6. Hybrid model: keeping the plan, adding the loop"),
  P("We do not replace our plan-driven tools. Each one gets an Agile partner:"),
  table([
    ["Plan-driven tool we keep", "Agile partner we add"],
    ["Stage gates (weeks 2, 4, 7, 10, 12)", "Sprint Review where the gate is checked"],
    ["WBS: all the work", "Product Backlog of user stories"],
    ["Critical path A→B→D→E→G→I→K→L", "Sprint order: critical stories first"],
    ["Earned value: SPI, CPI", "Velocity and burndown"],
    ["Risk register", "Blockers raised at stand-up, lessons at the retrospective"],
    ["Change control: sponsor approval", "MoSCoW: the PO reorders inside the fixed scope"],
  ], [4819, 4819]),
  caption("Table 5. Hybrid model used on Unveil"),
  P("Time (14 weeks) and budget (448 team-hours, $0) stay fixed. What we build can change inside them."),

  // ---------- 7 ----------
  H1("7. Risks, benefits and limitations"),
  P("Agile is not free: it needs discipline, regular meetings and good estimates, and without a strong Product Owner it can lead to scope creep. We linked our Agile practices to the risks in our register:"),
  table([
    ["Risk (score from register)", "Agile response"],
    ["Team overload during exams (P4 × I4 = 16, High)", "Plan capacity per sprint; core features before exams; WIP limit of 2"],
    ["Scope creep (P3 × I5 = 15, High)", "Won’t list; new ideas go to the bottom of the backlog; sponsor approval for new scope"],
    ["Privacy of photos and tests (P2 × I5 = 10, Medium)", "Part of the Definition of Done: test accounts only, photos locked on the server for 24 hours"],
    ["AI quality and cost (P3 × I3 = 9, Medium)", "Prototype with mock prompts in Sprint 3; AI co-pilot stays a Should"],
    ["Estimation is new to us", "Planning poker; compare estimates with actuals at every retrospective"],
  ], [4200, 5438]),
  caption("Table 6. Risks and Agile responses"),
  P([B("Benefits for Unveil: "), "early user feedback on the core ideas (no swipes, 24-hour reveal), a working product at every stage gate, visible progress for the sponsor, and the ability to cut low-priority features instead of missing the deadline."]),
  P([B("Limitations: "), "the four of us are part-time students, so a true daily Scrum is not realistic and we use a chat update instead; the roles of PO and Scrum Master are combined with development work; and story-point estimates will be rough until we have a few sprints of velocity data."]),

  // ---------- 8 ----------
  H1("8. Conclusion"),
  P("Agile project management delivers value in small, working pieces and uses regular feedback to steer the project. For Unveil, a new product with uncertain user behaviour but fixed time and budget, it is a natural fit. We keep the stage gates, WBS, critical path and earned value from our plan, and run Scrum sprints with a Kanban board inside them. The next step is Sprint 4 (weeks 9–10), with the Sprint Goal: two test users get matched by the algorithm, chat with each other and see each other’s photo after 24 hours, shown as a live demo at Gate 4."),

  // ---------- 9 ----------
  H1("9. Presentation plan"),
  table([
    ["Speaker", "Slides", "Topic"],
    ["Marlen Kalbayev", "1–4", "Introduction, why Agile, Agile Manifesto, Waterfall vs Agile"],
    ["Yeldos Bakayev", "5–7", "Scrum, our team roles, Product Backlog"],
    ["Diyor Khabibullayev", "8–10", "Sprint plan, Kanban board, estimation and Definition of Done"],
    ["Farkhad Amanbayev", "11–14", "Metrics, hybrid model, risks, next step and questions"],
  ], [2600, 1200, 5838]),
  caption("Table 7. Who presents what (about 10–12 minutes in total)"),

  // ---------- refs ----------
  H1("References"),
  bullet("Beck, K. et al. (2001). Manifesto for Agile Software Development. agilemanifesto.org"),
  bullet("Schwaber, K. and Sutherland, J. (2020). The Scrum Guide. scrumguides.org"),
  bullet("Anderson, D. J. (2010). Kanban: Successful Evolutionary Change for Your Technology Business. Blue Hole Press."),
  bullet("Project Management Institute (2017). Agile Practice Guide. PMI."),
  bullet("Unveil team (2026). Unveil: Project Progress Report, Week 7."),
];

const doc = new Document({
  creator: "Unveil team",
  title: "Agile Project Management, Applied to Unveil",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", run: { font: FONT, size: 44, bold: true }, paragraph: { spacing: { before: 0, after: 120 }, keepNext: true } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 30, bold: true }, paragraph: { spacing: { before: 320, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 24, bold: true }, paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
        new TextRun({ text: "Unveil · Agile report · ", size: 16, color: "5F5F5F" }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "5F5F5F" }),
      ] })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("wrote", OUT); });
