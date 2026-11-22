// Builds reports/NSUK_Connectivity_Deck.pptx, which is uploaded to Google Slides.
// Run: python scripts/deck_assets.py  then  node scripts/deck/build_deck.js
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..", "..");
const D = JSON.parse(fs.readFileSync(path.join(__dirname, "deck_data.json"), "utf8"));
const OUT = path.join(ROOT, "reports", "NSUK_Connectivity_Deck.pptx");
const APPLY_THEME = process.env.APPLY_THEME;

// ---------- one colour family: teal, light -> dark ----------
const T = ["F4F9FA", "E6F4F6", "B9E2E8", "6CC1CF", "2A9DB0", "137C8F", "0E5A6B", "0B3C49"]; // T[0] lightest ... T[7] darkest
const INK = "1B2B30", INK2 = "5B6B70", MUTED = "8A9A9F", GRID = "E3E9EA";
const THEME = {
  name: "NSUK Teal", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: INK, lt1: "FFFFFF", dk2: T[7], lt2: "EEF6F7", accent1: T[5], accent2: T[6], accent3: T[4],
            accent4: T[3], accent5: T[2], accent6: T[7], hlink: T[5], folHlink: T[6] },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Internet Connectivity and Students' Learning Experience at NSUK";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333, M = 0.6, CW = W - 2 * M;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK", background: { color: T[7] },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 1.9, w: 8.2, h: 2.2, fontFace: "Cambria", fontSize: 40, bold: true, color: "FFFFFF", valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.3, w: 8.2, h: 1.0, fontFace: "Calibri", fontSize: 18, color: T[2], valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: CW, h: 1.05, fontFace: "Cambria", fontSize: 34, bold: true, color: T[7], valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Internet Connectivity at NSUK   |   NS-DTS 223 Group Practical Assignment", options: { x: M, y: 7.0, w: 9, h: 0.3, fontFace: "Calibri", fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.55, h: 0.3, fontFace: "Calibri", fontSize: 10, color: MUTED, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING_DARK", background: { color: T[7] },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: CW, h: 1.05, fontFace: "Cambria", fontSize: 34, bold: true, color: "FFFFFF", valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: "Calibri", color: INK, margin: 0, valign: "top" }, o));
const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
const caption = (s, text, x, y, w) => txt(s, text, { x, y, w, h: 0.35, fontSize: 12, color: INK2, italic: true });
const numCircle = (s, n, x, y, d, fill, color) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: `num-${n}` });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: d > 0.5 ? 18 : 14, bold: true, color, align: "center", valign: "middle" });
};
// signal-strength bars: the deck's motif
const signal = (s, x, y, scale, tones) => {
  const bw = 0.32 * scale, gap = 0.14 * scale;
  [0.45, 0.85, 1.3, 1.8].forEach((h, i) => {
    const hh = h * scale;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + i * (bw + gap), y: y + 1.8 * scale - hh, w: bw, h: hh, rectRadius: 0.05 * scale,
      fill: { color: tones[i] }, line: { color: tones[i] }, objectName: `signal-bar-${i + 1}` });
  });
};
const stat = (s, x, y, w, big, label, opts = {}) => {
  txt(s, big, { x, y, w, h: 0.9, fontFace: "Cambria", fontSize: opts.size || 48, bold: true, color: opts.color || T[6] });
  txt(s, label, { x, y: y + (opts.size ? opts.size / 60 : 0.85), w, h: opts.lh || 0.8, fontSize: 14, color: opts.labelColor || INK2 });
};
// Charts are images rendered by scripts/deck_assets.py in the same teal ramp: Google Slides drops the
// category labels and number formats of imported native PowerPoint charts.
const img = (s, file, x, y, w, h, altText) => s.addImage({ path: path.join(__dirname, file), x, y, w, h, altText });
const slide = (master, section) => pres.addSlide({ masterName: master, sectionTitle: section });

// =====================================================================
pres.addSection({ title: "Introduction" });
// 1. Title
let s = slide("TITLE_DARK", "Introduction");
s.addText("Internet Connectivity and Students' Learning Experience at NSUK", { placeholder: "title" });
s.addText("NS-DTS 223 Data Exploration and Visualization  |  Group Practical Assignment, Topic 2", { placeholder: "body" });
txt(s, `Survey of ${D.n} students  ·  18–22 November 2026  ·  Nasarawa State University, Keffi`, { x: M, y: 5.6, w: 8.2, h: 0.4, fontSize: 14, color: T[3] });
signal(s, 9.6, 2.1, 1.55, [T[3], T[4], T[5], T[2]]);
s.addNotes("Introduce the group and the topic. We investigated how students experience internet connectivity on campus and whether poorer connectivity is linked to disruption of their learning. Everything in this deck comes from our own survey of 84 students.");

// 2. Questions
s = slide("CONTENT", "Introduction");
s.addText("What we set out to answer", { placeholder: "title" });
card(s, M, 1.7, 4.6, 4.9, "EEF6F7", "central-question-card");
txt(s, "CENTRAL QUESTION", { x: M + 0.35, y: 2.0, w: 3.9, h: 0.3, fontSize: 12, bold: true, color: T[5], charSpacing: 2 });
txt(s, "How do students experience internet connectivity across the NSUK campus, and how is a poorer experience associated with disruption to their learning?",
    { x: M + 0.35, y: 2.45, w: 3.9, h: 2.6, fontFace: "Cambria", fontSize: 21, italic: true, color: T[7] });
txt(s, "Audience: university management, the ICT directorate, the Students' Union and mobile network operators.", { x: M + 0.35, y: 5.35, w: 3.9, h: 0.9, fontSize: 13, color: INK2 });
const rqs = [["Access", "How do students get online on campus?"], ["Location", "Is the experience worse in some campus locations?"],
  ["Time", "At what time of day is the connection best?"], ["Relationship", "Is poorer connectivity linked to greater academic impact?"],
  ["Deadlines", "Are students with poorer connectivity more likely to miss deadlines?"], ["Action", "What do students want the university to do?"]];
rqs.forEach(([k, q], i) => {
  const y = 1.75 + i * 0.8;
  numCircle(s, i + 1, 5.75, y, 0.52, T[5], "FFFFFF");
  txt(s, [{ text: k, options: { bold: true, color: T[7], breakLine: true } }, { text: q, options: { color: INK2 } }], { x: 6.5, y: y - 0.06, w: 6.2, h: 0.72, fontSize: 15 });
});
s.addNotes("Our central question has two parts: the experience itself, and its link to learning. We broke it into six analytical questions. Each later chart answers one of them, in this order.");

// 3. Method
s = slide("CONTENT", "Introduction");
s.addText("84 students surveyed over five days", { placeholder: "title" });
[["84", "valid responses"], ["5", "days of collection, 18–22 Nov"], ["20", "questions in the online form"], ["0", "respondents removed in cleaning"]].forEach(([b, l], i) => {
  const x = M + i * 3.07;
  card(s, x, 1.7, 2.85, 1.75, "EEF6F7", `method-stat-${i + 1}`);
  txt(s, b, { x: x + 0.3, y: 1.85, w: 2.3, h: 0.85, fontFace: "Cambria", fontSize: 44, bold: true, color: T[6] });
  txt(s, l, { x: x + 0.3, y: 2.75, w: 2.4, h: 0.55, fontSize: 14, color: INK2 });
});
const steps = ["Collect", "Inspect & audit", "Clean & validate", "Explore", "Tell the story"];
steps.forEach((t, i) => {
  const x = M + i * 2.45, fill = T[2 + i], dark = i >= 2;
  s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x, y: 3.95, w: 2.55, h: 0.95, fill: { color: fill }, line: { color: fill }, objectName: `step-${i + 1}` });
  txt(s, t, { x: x + (i === 0 ? 0.2 : 0.45), y: 3.95, w: 1.85, h: 0.95, fontSize: 15, bold: true, color: dark ? "FFFFFF" : T[7], valign: "middle", align: "center" });
});
txt(s, [
  { text: "Instrument: ", options: { bold: true, color: T[7] } }, { text: "anonymous Google Form covering profile, access, connectivity experience, academic impact, cost and suggestions.", options: { breakLine: true } },
  { text: "Sampling: ", options: { bold: true, color: T[7] } }, { text: "convenience sample of NSUK students, shared online; voluntary, no names or ID numbers collected.", options: { breakLine: true } },
  { text: "Tools: ", options: { bold: true, color: T[7] } }, { text: "Python (pandas, SciPy, Matplotlib, Seaborn) in a Jupyter notebook; every step is reproducible." },
], { x: M, y: 5.25, w: CW, h: 1.5, fontSize: 15, color: INK2, paraSpaceAfter: 6 });
s.addNotes("We designed a 20-question Google Form and collected 84 valid responses between 18 and 22 November. The sample is a convenience sample, which matters for how far we can generalise. The workflow followed the course stages: collect, audit, clean, explore, and build the story.");

// 4. Cleaning
s = slide("CONTENT", "Introduction");
s.addText("Cleaning fixed six problems without dropping anyone", { placeholder: "title" });
const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: T[6] }, fontSize: 14 } });
const issues = [
  ["Q8 rating switched from labels to a 1–5 scale during collection", "31", "Mapped 1 = Very poor … 5 = Excellent, after checking the direction against related questions"],
  ["Two options ticked on single-choice questions", "8", "Kept the option nearer the midpoint; contradictory pairs set to missing"],
  ["Department typed 27 different ways; 6 left blank", "84", "Standardised to 19 departments; blanks marked \"Not stated\""],
  ["Faculty did not match department", "1", "Corrected to Natural and Applied Sciences and flagged"],
  ["\"No difference\" ticked together with a specific time", "4", "Kept the specific time slot"],
  ["Non-answers in open questions (\"Yes\", \"Nil\")", "8", "Flagged and left out of theme coding"],
];
const rows = [[hdr("Problem found"), hdr("Rows"), hdr("What we did")]].concat(issues.map((r, i) => r.map((c, j) => ({
  text: c, options: { fill: { color: i % 2 ? "FFFFFF" : "EEF6F7" }, align: j === 1 ? "center" : "left", bold: j === 1, color: j === 1 ? T[6] : INK } }))));
s.addTable(rows, { x: M, y: 1.65, w: CW, colW: [4.9, 0.9, 6.3], fontFace: "Calibri", fontSize: 14, valign: "middle",
                   rowH: 0.62, border: { type: "solid", pt: 0.75, color: GRID }, margin: [0.06, 0.12, 0.06, 0.12] });
caption(s, "No duplicate submissions were found. Every change is recorded in a cleaning log generated by the notebook.", M, 6.25, CW);
s.addNotes("These are the quality problems we found and how we handled each one. The most important was Q8: the form changed from text labels to a 1 to 5 scale halfway through. We confirmed that 1 means Very poor by checking that the numeric ratings correlate with satisfaction and disconnection in the same direction as the text ratings. Nobody was removed; missing answers are handled question by question.");

// 5. Profile
pres.addSection({ title: "Findings" });
s = slide("CONTENT", "Findings");
s.addText("Most respondents are 200–300 Level science students", { placeholder: "title" });
caption(s, "Respondents by level of study (number of students)", M, 1.6, 7.4);
img(s, "chart_level.png", M, 1.95, 7.4, 4.7, "Column chart of respondents by level: 100 Level 4, 200 Level 28, 300 Level 45, 400 Level 6, 500 Level 1.");
const nas = Math.round(D.faculty["Natural & Applied Sciences"] / D.n * 100);
card(s, 8.45, 1.95, 4.28, 4.7, "EEF6F7", "profile-card");
stat(s, 8.8, 2.15, 3.7, `${nas}%`, "from the Faculty of Natural and Applied Sciences");
stat(s, 8.8, 3.75, 3.7, "68%", "from computing departments (Computer Science, Data Science, Cyber Security)", { lh: 0.9 });
txt(s, "Results describe this group best. They are indicative, not campus-wide estimates.", { x: 8.8, y: 5.55, w: 3.7, h: 0.9, fontSize: 13, italic: true, color: T[6] });
s.addNotes("Before the findings, who answered. Most respondents are in 200 or 300 Level, and almost four in five are from Natural and Applied Sciences, mainly computing departments. So read the results as describing this group, not the whole university.");

// 6. Access (V1)
s = slide("CONTENT", "Findings");
s.addText("Most students get online with their own mobile data", { placeholder: "title" });
caption(s, "Network providers used on campus (% of students; some use more than one)", M, 1.6, 7.2);
img(s, "chart_providers.png", M, 1.95, 7.2, 4.6, "Bar chart of network providers used: MTN 54%, Glo 32%, Airtel 27%, 9mobile 1%.");
const k = D.kpi;
txt(s, `${Math.round(D.conn_type["Mobile data"])}%`, { x: 8.3, y: 1.75, w: 4.4, h: 1.5, fontFace: "Cambria", fontSize: 88, bold: true, color: T[7] });
txt(s, "rely mainly on their own mobile data", { x: 8.35, y: 3.25, w: 4.3, h: 0.5, fontSize: 18, color: INK });
s.addShape(pres.shapes.LINE, { x: 8.35, y: 4.0, w: 4.3, h: 0, line: { color: GRID, width: 1 } });
stat(s, 8.35, 4.15, 2.1, `${Math.round(D.conn_type["Campus Wi-Fi"])}%`, "mainly use campus Wi-Fi", { size: 36, color: T[5] });
stat(s, 10.55, 4.15, 2.2, `${Math.round(k.academic_use)}%`, "use the internet very often or always for study", { size: 36, color: T[5], lh: 0.9 });
caption(s, "Visualization 1: Overview", 8.35, 6.25, 4.3);
s.addNotes("Visualization 1, the overview. MTN is the most used network, followed by Glo and Airtel. The key number is 82%: most students rely on their own mobile data, and only 10% mainly use campus Wi-Fi. At the same time 89% use the internet very often or always for academic work. So learning depends on networks students pay for themselves.");

// 7. Experience
s = slide("CONTENT", "Findings");
s.addText("Problems are frequent and satisfaction is low", { placeholder: "title" });
const cards4 = [[`${Math.round(k.poor_often)}%`, "experience poor connectivity often or always"],
  [`${Math.round(k.disconnect_often)}%`, "are disconnected unexpectedly often or always"],
  [`${Math.round(k.satisfied)}%`, "are satisfied with their provider's speed"],
  [`${Math.round(k.median_cei)}`, "median Connectivity Experience Index, out of 100"]];
cards4.forEach(([b, l], i) => {
  const x = M + i * 3.07, fill = [T[6], T[5], T[4], T[7]][i];
  card(s, x, 1.75, 2.85, 2.9, fill, `experience-card-${i + 1}`);
  txt(s, b, { x: x + 0.3, y: 2.0, w: 2.3, h: 1.1, fontFace: "Cambria", fontSize: 56, bold: true, color: "FFFFFF" });
  txt(s, l, { x: x + 0.3, y: 3.2, w: 2.3, h: 1.2, fontSize: 15, color: "FFFFFF" });
});
card(s, M, 4.95, CW, 1.4, "EEF6F7", "cei-explainer");
txt(s, [{ text: "How the index works. ", options: { bold: true, color: T[7] } },
  { text: "The Connectivity Experience Index combines four questions: the connection rating, how often connectivity is poor, how often it drops, and satisfaction with speed. 0 is the worst possible experience and 100 the best. The four items hang together well (Cronbach's alpha = 0.78). The typical student scores below the midpoint of 50." }],
  { x: M + 0.35, y: 5.15, w: CW - 0.7, h: 1.05, fontSize: 15, color: INK2 });
s.addNotes("Problems are frequent. Half of students face poor connectivity often or always, a third are cut off unexpectedly that often, and only a third are satisfied with their provider's speed. To compare groups we combined these four questions into one index from 0 to 100. The median student scores 44, below the midpoint.");

// 8. Location (V2)
s = slide("CONTENT", "Findings");
s.addText("The experience is middling in every campus location", { placeholder: "title" });
caption(s, "Connectivity Experience Index by main access location (each dot is one student; number of students in brackets)", M, 1.6, 8.2);
s.addImage({ path: path.join(__dirname, "boxplot_location.png"), x: M, y: 1.95, w: 8.0, h: 4.78, altText: "Box plot of Connectivity Experience Index by access location. Medians range from 44 to 56 with wide overlapping spread." });
const side = [["44–56", "range of median scores across the six locations"], ["p = 0.60", "differences between locations are not statistically significant"], ["LT1 & LT2", "FNAS lecture theatres named in comments as weak spots"]];
side.forEach(([b, l], i) => stat(s, 9.1, 1.85 + i * 1.6, 3.6, b, l, { size: 34, color: T[6] }));
caption(s, "Visualization 2: Distribution (box plot)", 9.1, 6.45, 3.6);
s.addNotes("Visualization 2 is a box plot of the index by where students mostly go online. All the boxes overlap, and the medians only range from 44 to 56. The difference is not statistically significant. The spread inside each location is much bigger than the gap between them. So the problem is campus-wide, not one bad building, although students did name the FNAS lecture theatres in their comments.");

// 9. Time (V3)
s = slide("CONTENT", "Findings");
s.addText("Connectivity is best at night and worst in the afternoon", { placeholder: "title" });
caption(s, `Students naming each time slot as the best for connectivity (% of the ${D.time_named} who named a slot)`, M, 1.6, 8.2);
img(s, "chart_time.png", M, 1.95, 8.0, 4.7, "Line chart of best time for connectivity: morning 21%, afternoon 19%, evening 26%, night 45%.");
stat(s, 9.1, 1.85, 3.6, `${Math.round(D.time.Night)}%`, "say night (after 8 PM) is best", { size: 54, color: T[7] });
stat(s, 9.1, 3.4, 3.6, `${Math.round(D.time.Afternoon)}%`, "say the afternoon is best, the lowest of any slot", { size: 34, color: T[5] });
txt(s, "This fits congestion during teaching hours, when classes, portal work and research all happen on campus.", { x: 9.1, y: 4.85, w: 3.6, h: 1.2, fontSize: 14, color: INK2 });
caption(s, `Visualization 3: Pattern. ${D.time_nodiff} students noticed no difference.`, 9.1, 6.3, 3.6);
s.addNotes("Visualization 3 shows when students say the connection is best. Night is clearly ahead at 45%, and the afternoon is lowest. This is consistent with congestion during the day, when most people on campus are online. The best connection comes when students are least able to use it for scheduled learning.");

// 10. Relationship (V4) - heatmap as a native table
s = slide("CONTENT", "Findings");
s.addText("Poorer connectivity goes with greater academic impact", { placeholder: "title" });
caption(s, "Effect of poor connectivity on academic performance, by rating of campus connection (% of each row; number of students)", M, 1.6, 8.6);
const shade = (p) => p === 0 ? T[0] : p < 0.15 ? T[1] : p < 0.25 ? T[2] : p < 0.35 ? T[3] : p < 0.45 ? T[4] : p < 0.55 ? T[5] : T[6];
const hrow = [{ text: "Rating  \\  Impact", options: { bold: true, color: INK2, fill: { color: "FFFFFF" }, align: "left" } }]
  .concat(D.heat_cols.map((c) => ({ text: c, options: { bold: true, color: INK2, fill: { color: "FFFFFF" }, align: "center" } })));
const hrows = [hrow].concat(D.heat_rows.map((r, i) => [{ text: [{ text: r, options: { bold: true, breakLine: true } }, { text: `${D.heat_n[i]} students`, options: { fontSize: 11, color: INK2 } }], options: { fill: { color: "FFFFFF" }, color: INK, align: "left" } }]
  .concat(D.heat_pct[i].map((p, j) => {
    const fg = p >= 0.35 ? "FFFFFF" : INK;
    return { text: [{ text: `${Math.round(p * 100)}%`, options: { bold: true, fontSize: 16, color: fg, breakLine: true } }, { text: `${D.heat_counts[i][j]}`, options: { fontSize: 11, color: fg } }],
             options: { fill: { color: shade(p) }, align: "center" } };
  }))));
s.addTable(hrows, { x: M, y: 2.0, w: 8.3, colW: [1.8, 1.3, 1.3, 1.3, 1.3, 1.3], rowH: [0.5, 0.82, 0.82, 0.82, 0.82, 0.82], fontFace: "Calibri", fontSize: 13,
                    valign: "middle", border: { type: "solid", pt: 2, color: "FFFFFF" }, margin: [0.04, 0.1, 0.04, 0.1] });
stat(s, 9.35, 1.85, 3.4, "ρ = −0.31", "correlation between the index and academic impact (p = 0.005)", { size: 36, color: T[7], lh: 0.9 });
txt(s, "The link holds after allowing for data spend, level of study and how often students use the internet (partial ρ = −0.28, p = 0.011).", { x: 9.35, y: 3.65, w: 3.4, h: 1.5, fontSize: 14, color: INK2 });
txt(s, "This is an association, not proof that connectivity causes the impact.", { x: 9.35, y: 5.2, w: 3.4, h: 0.9, fontSize: 14, italic: true, color: T[6] });
caption(s, "Visualization 4: Relationship (heatmap)", 9.35, 6.3, 3.4);
s.addNotes("Visualization 4 is a heatmap. Each row is a connection rating and the shading shows how that group rated the academic impact. The dark cells run from top right to bottom left: half of those with a very poor connection say the impact is very high, while nobody rating it good or excellent says so. The correlation is negative and statistically significant, and it holds after controlling for spending, level and usage. We are careful to call this an association.");

// 11. Key insight (V5)
s = slide("CONTENT", "Findings");
s.addText("High impact is 2.6× as common with a poor connection", { placeholder: "title" });
const G = D.groups, gl = ["Poor", "Average", "Good"];
caption(s, "Share of students in each connectivity group (from the Q8 rating)", M, 1.6, 8.0);
img(s, "chart_groups.png", M, 1.95, 8.0, 4.75, `Grouped column chart. Missed a deadline: Poor ${G.Poor.missed}%, Average ${G.Average.missed}%, Good ${G.Good.missed}%. High impact: Poor ${G.Poor.high}%, Average ${G.Average.high}%, Good ${G.Good.high}%.`);
stat(s, 9.1, 1.85, 3.6, "65% vs 25%", "of students with a poor connection report high academic impact, compared with the rest (p = 0.005)", { size: 40, color: T[7], lh: 1.2 });
stat(s, 9.1, 3.95, 3.6, `${Math.round(k.missed)}%`, "of all students have missed a deadline because of connectivity: 85% (poor) to 59% (good)", { size: 40, color: T[4], lh: 1.2 });
caption(s, "Visualization 5: Key insight. The deadline trend is not statistically significant (p = 0.15).", 9.1, 6.05, 3.6);
s.addNotes("Visualization 5 is the key insight. Students with a poor connection are 2.6 times as likely to say their performance is highly affected: 65% against 25%. Missed deadlines are common everywhere, 71% overall. The share drops from 85% to 59% as connectivity improves, but with 84 responses that trend is not statistically significant, so we present it as a pattern, not a proven difference.");

// 12. What students want
s = slide("CONTENT", "Findings");
s.addText("Students want campus Wi-Fi and better infrastructure", { placeholder: "title" });
caption(s, "What the university should do (% of students; more than one answer allowed)", M, 1.6, 7.2);
img(s, "chart_uni.png", M, 1.95, 7.2, 4.6, "Bar chart of what the university should do: more Wi-Fi hotspots 50%, improve infrastructure 40%, partner with telecom providers 29%, affordable access 29%.");
const quotes = ["Lt1 and Lt2 in FNAS have very terrible network traffic and should be fixed please",
  "Immediately i enter the campus, my network disappears", "I could not download an educational file for 1hr because of bad network."];
txt(s, "IN STUDENTS' WORDS", { x: 8.3, y: 1.65, w: 4.4, h: 0.3, fontSize: 12, bold: true, color: T[5], charSpacing: 2 });
quotes.forEach((q, i) => {
  const y = 2.05 + i * 1.5;
  card(s, 8.3, y, 4.43, 1.3, "EEF6F7", `quote-${i + 1}`);
  txt(s, `"${q}"`, { x: 8.55, y: y + 0.15, w: 3.95, h: 1.0, fontFace: "Cambria", fontSize: 14, italic: true, color: T[7], valign: "middle" });
});
caption(s, "Quotes reproduced as written.", 8.3, 6.6, 4.4);
s.addNotes("Students' own priorities match the evidence. Half want more Wi-Fi hotspots and 40% want better infrastructure. Their comments name specific places, such as the FNAS lecture theatres, and specific failures, such as downloads that take an hour.");

// 13. Recommendations
pres.addSection({ title: "Recommendations" });
s = slide("CONTENT", "Recommendations");
s.addText("Five recommendations supported by the evidence", { placeholder: "title" });
const recs = [
  ["Expand campus Wi-Fi in academic spaces", "Faculty buildings, FNAS LT1 and LT2, the library and lecture halls, sized for daytime peaks.", "10% use campus Wi-Fi; 50% want hotspots"],
  ["Agree capacity upgrades with operators", "Work with MTN, Glo and Airtel on capacity for the campus during teaching hours.", "52% face poor connectivity often; night is best"],
  ["Make teaching tolerant of outages", "Record online classes, share downloadable materials, allow grace periods for outage-related late work.", "71% have missed a deadline"],
  ["Reduce students' data costs", "Zero-rated access to the portal, LMS and e-library; student data bundles.", "65% spend over ₦3,000 a week"],
  ["Measure connectivity regularly", "Scheduled speed tests at fixed campus points each semester, with published results.", "This study relies on perceptions only"],
];
recs.forEach(([t, d, e], i) => {
  const row = i < 3 ? 0 : 1, w = row ? 5.95 : 3.93, x = M + (row ? (i - 3) * 6.15 : i * 4.08), y = row ? 4.3 : 1.65, h = 2.4;
  card(s, x, y, w, h, "EEF6F7", `rec-card-${i + 1}`);
  numCircle(s, i + 1, x + 0.3, y + 0.3, 0.55, T[6], "FFFFFF");
  txt(s, t, { x: x + 1.05, y: y + 0.28, w: w - 1.3, h: 0.65, fontSize: 16, bold: true, color: T[7], valign: "middle" });
  txt(s, d, { x: x + 0.3, y: y + 1.0, w: w - 0.6, h: 0.85, fontSize: 14, color: INK });
  txt(s, e, { x: x + 0.3, y: y + 1.9, w: w - 0.6, h: 0.35, fontSize: 12, italic: true, color: T[5] });
});
s.addNotes("Our five recommendations, each tied to a finding. First, campus Wi-Fi in academic spaces, because almost nobody uses it now and the problem is everywhere. Second, work with the operators on daytime capacity. Third, lecturers can reduce the damage with recordings, offline materials and grace periods. Fourth, cut data costs. Fifth, measure connectivity properly, which also addresses our own main limitation.");

// 14. Conclusion
s = slide("CLOSING_DARK", "Recommendations");
s.addText("Conclusion", { placeholder: "title" });
const concl = ["Students depend on the internet for learning but mostly pay for their own mobile data.",
  "The experience is middling across the whole campus and best only at night.",
  "Students with poorer connectivity report much greater academic harm, and most have missed a deadline."];
concl.forEach((c, i) => {
  numCircle(s, i + 1, M, 1.75 + i * 1.25, 0.55, T[3], T[7]);
  txt(s, c, { x: M + 0.85, y: 1.7 + i * 1.25, w: 6.4, h: 1.0, fontSize: 18, color: "FFFFFF", valign: "middle" });
});
card(s, 8.15, 1.65, 4.58, 3.05, T[6], "limitations-card");
txt(s, "LIMITATIONS", { x: 8.45, y: 1.9, w: 4.0, h: 0.3, fontSize: 12, bold: true, color: T[2], charSpacing: 2 });
txt(s, [
  { text: "84 responses, below the 100 target; mostly science students", options: { bullet: true, breakLine: true } },
  { text: "Self-reported only; no speed tests were run", options: { bullet: true, breakLine: true } },
  { text: "The Q8 format changed during collection", options: { bullet: true, breakLine: true } },
  { text: "Associations, not causes", options: { bullet: true } },
], { x: 8.45, y: 2.3, w: 4.0, h: 2.25, fontSize: 15, color: "FFFFFF", paraSpaceAfter: 8 });
signal(s, M, 5.75, 0.45, [T[3], T[4], T[5], T[2]]);
txt(s, "Thank you. Questions?", { x: M + 1.0, y: 5.85, w: 6, h: 0.6, fontFace: "Cambria", fontSize: 26, bold: true, color: "FFFFFF", valign: "middle" });
s.addNotes("To conclude: students rely on the internet and on their own data, the experience is middling everywhere and best only at night, and poorer connectivity goes with greater academic harm. Our main limitations are the sample size and make-up, the lack of measured speed data, and that we can show association but not cause. Thank you, we welcome questions.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (APPLY_THEME) { const { applyTheme } = require(APPLY_THEME); await applyTheme(OUT, THEME); }
  console.log("saved", OUT);
})();
