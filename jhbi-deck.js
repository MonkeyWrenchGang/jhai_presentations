"use strict";
const pptxgen = require("pptxgenjs");

// ─── Jack Henry Brand Colors ────────────────────────────────────
const JH = {
  navy:     "06185F",
  cobalt:   "085CE5",
  techBlue: "76DCFD",
  openSky:  "E8F7F7",
  ivory:    "FEFDF8",
  dkGray:   "575A5D",
  mdGray:   "B6BBC0",
  ltGray:   "E7ECF0",
  white:    "FFFFFF",
  teal:     "0A5CA8",
  green:    "0A5C47",
};

// ─── Pillar definitions ─────────────────────────────────────────
const PILLARS = [
  { name: "Data Acquisition",       short: "Data Acq.",    color: JH.navy,   text: JH.white  },
  { name: "Reporting & Dashboards", short: "Reporting",    color: JH.cobalt, text: JH.white  },
  { name: "Portal / Self Service",  short: "Portal",       color: JH.teal,   text: JH.white  },
  { name: "Data Science Apps",      short: "Data Science", color: JH.green,  text: JH.white  },
];

// ─── Team ───────────────────────────────────────────────────────
const TEAM_LEAD = { name: "Mike Ames", title: "Head of Data Science", initials: "MA" };
const TEAM_REPORTS = [
  { name: "Michael Goolsby", title: "Director of Software Development", initials: "MG", focus: "Software Dev"  },
  { name: "[TBD]",           title: "Mgr, JHBI Product Mgmt & GTM",    initials: "PM", focus: "Product & GTM" },
  { name: "[TBD]",           title: "Data Science Lead",                initials: "DS", focus: "Data Science"  },
];

// ─── Quarterly initiatives [quarter][pillar]  (FY27: Jul 2026 – Jun 2027) ──
const Q_INITIATIVES = [
  [  // Q1 FY27  Jul–Sep 2026
    "Core Banking Data Pipeline — FI Connectivity Layer",
    "Executive KPI Dashboard — Platform MVP Launch",
    "Churn Mediation MVP — JH CRM Integration",
    "ML Model Serving Infrastructure & Feature Store",
  ],
  [  // Q2 FY27  Oct–Dec 2026
    "Data Quality Framework & Lineage Tracking",
    "Risk Score Explorer — Banker-Facing Interface",
    "Zelle Memo Intelligence — NLP Pipeline Live",
    "Anomaly Detection — Streaming Production Launch",
  ],
  [  // Q3 FY27  Jan–Mar 2027
    "Real-time Kafka Streaming Infrastructure",
    "Model Drift Monitoring & Alert Engine",
    "Call Report AI — Full CU Client Rollout",
    "MLOps Pipeline — Full CI/CD & Governance",
  ],
  [  // Q4 FY27  Apr–Jun 2027
    "Enterprise Data Catalog & Governance Framework",
    "Self-Serve Analytics Surface — Business Users",
    "Account Opening LTV — JH Platform Launch",
    "Cross-Sell Propensity Engine (Horizon Phase 1)",
  ],
];

// ─── Gantt items  (18-month view: month 0 = Jan 2026) ───────────
// Foundation: months 0–5 (Jan–Jun 2026)
// FY27 Q1: months 6–8 · Q2: 9–11 · Q3: 12–14 · Q4: 15–17
const GANTT = [
  { p: 0, name: "FI Data Pipeline Architecture & Core Banking Integration", s: 0,  e: 5,  foundation: true  },
  { p: 1, name: "Dashboard Requirements & Design Baseline",                 s: 0,  e: 5,  foundation: true  },
  { p: 2, name: "AI App Architecture & FI Data Access Agreements",          s: 0,  e: 5,  foundation: true  },
  { p: 3, name: "ML Platform: Feature Store, Experiment Tracking & Registry", s: 0, e: 5, foundation: true  },
  { p: 0, name: "Core Banking Pipeline Integration",                        s: 6,  e: 8,  foundation: false },
  { p: 0, name: "Data Quality Framework & Enterprise Catalog",              s: 9,  e: 17, foundation: false },
  { p: 1, name: "Executive KPI Dashboard — Platform Launch",                s: 6,  e: 9,  foundation: false },
  { p: 1, name: "Risk Score Explorer + Drift Monitoring",                   s: 9,  e: 17, foundation: false },
  { p: 2, name: "Churn Mediation + Zelle Memo Intelligence",                s: 6,  e: 11, foundation: false },
  { p: 2, name: "Call Report AI Scale + Account Opening LTV",               s: 12, e: 17, foundation: false },
  { p: 3, name: "Anomaly Detection — Streaming Production",                 s: 6,  e: 11, foundation: false },
  { p: 3, name: "MLOps Pipeline + Horizon App Scoping",                     s: 12, e: 17, foundation: false },
];

// ─── Foundation phase deliverables per pillar (Jan–Jun 2026) ────
const Q_FOUNDATION = [
  "FI Data Pipeline Architecture & Core Banking Connectivity",
  "Dashboard Design Baseline & KPI Framework",
  "AI App Architecture & FI Data Access Agreements",
  "ML Platform: Feature Store, Experiment Tracking & Model Registry",
];

// ─── Program deliverables for overview grid ─────────────────────
const DELIVERABLES = [
  ["FI behavioral data pipelines from JH core systems", "Real-time streaming infrastructure (Kafka / event bus)", "Data quality monitoring, lineage & enterprise catalog"],
  ["Executive KPI dashboards embedded in JH platform", "Model performance & drift monitoring dashboards", "FI risk score explorer & self-serve analytics surface"],
  ["AI apps in JH ecosystem (Churn, Anomaly, Zelle Memo)", "REST APIs distributing model scores to JH products", "Banno + JH Enterprise integration & alerting workflows"],
  ["AutoGluon, CatBoost, XGBoost model training at scale", "SHAP explainability layer across all deployed models", "MLOps: CI/CD, model registry, drift monitoring & governance"],
];

// ─── Shadow factory (avoids pptxgenjs object mutation bug) ──────
const mkShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.08 });
const mkShadowSm = () => ({ type: "outer", color: "000000", blur: 3, offset: 1, angle: 135, opacity: 0.07 });

// ─── Global helpers ─────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title  = "JHBI Analytics — FY27 Strategic Roadmap";

function hLine(slide, x1, x2, y, color, w) {
  slide.addShape(pres.shapes.LINE, { x: x1, y, w: x2 - x1, h: 0, line: { color: color || JH.mdGray, width: w || 1 } });
}
function vLine(slide, x, y1, y2, color, w) {
  slide.addShape(pres.shapes.LINE, { x, y: y1, w: 0, h: y2 - y1, line: { color: color || JH.mdGray, width: w || 1 } });
}

function headerBar(slide, title, subtitle) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.52, fill: { color: JH.navy }, line: { color: JH.navy } });
  slide.addText(title, {
    x: 0.4, y: 0, w: subtitle ? 6 : 9.2, h: 0.52,
    fontSize: 13, bold: true, color: JH.white, fontFace: "Calibri", valign: "middle", margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 6.5, y: 0, w: 3.2, h: 0.52,
      fontSize: 10, color: JH.techBlue, fontFace: "Calibri", valign: "middle", align: "right", margin: 0,
    });
  }
}

function footerBar(slide, note) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.42, w: 10, h: 0.205, fill: { color: JH.ltGray }, line: { color: JH.ltGray } });
  slide.addText(note || "jack henry\u2122  \u00b7  JHBI Analytics  \u00b7  Confidential  \u00b7  FY27  \u00b7  Jul 2026 \u2013 Jun 2027", {
    x: 0.4, y: 5.42, w: 9.2, h: 0.205, fontSize: 8, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
  });
}

// ─── Build slides ────────────────────────────────────────────────
addCoverSlide();
addTeamSlide();
addProgramOverviewSlide();
addQuarterlyRoadmapSlide();
addGanttSlide();
addFinancialImpactSlide();
addNowNextLaterSlide();

pres.writeFile({ fileName: "/sessions/clever-stoic-dirac/mnt/outputs/jhbi-roadmap-2026.pptx" })
  .then(() => console.log("✅  Written: jhbi-roadmap-2026.pptx"))
  .catch(err  => { console.error("❌", err); process.exit(1); });


// ═══════════════════════════════════════════════════════════════
//  SLIDE 1 — Cover
// ═══════════════════════════════════════════════════════════════
function addCoverSlide() {
  const slide = pres.addSlide();
  slide.background = { color: JH.navy };

  // Top accent bar (cobalt)
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.09, fill: { color: JH.cobalt }, line: { color: JH.cobalt } });
  // Bottom accent bar (tech blue)
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.535, w: 10, h: 0.09, fill: { color: JH.techBlue }, line: { color: JH.techBlue } });
  // Right accent bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 9.88, y: 0, w: 0.12, h: 5.625, fill: { color: JH.cobalt }, line: { color: JH.cobalt } });

  slide.addText("jack henry\u2122", {
    x: 0.55, y: 0.5, w: 5, h: 0.32,
    fontSize: 11, color: JH.techBlue, fontFace: "Calibri", charSpacing: 2,
  });
  slide.addText("JHBI Analytics", {
    x: 0.55, y: 1.05, w: 9, h: 1.0,
    fontSize: 46, bold: true, color: JH.white, fontFace: "Calibri",
  });
  slide.addText("FY27 Strategic Roadmap", {
    x: 0.55, y: 2.1, w: 8.5, h: 0.55,
    fontSize: 24, color: JH.techBlue, fontFace: "Calibri",
  });
  slide.addText("Q1\u2013Q4  FY27  \u00b7  Jul 2026 \u2013 Jun 2027", {
    x: 0.55, y: 2.72, w: 4, h: 0.36,
    fontSize: 13, color: JH.mdGray, fontFace: "Calibri", charSpacing: 2,
  });

  // Divider
  hLine(slide, 0.55, 5.5, 3.28, JH.cobalt, 1.5);

  // 4 pillar tiles
  PILLARS.forEach((p, i) => {
    const px = 0.55 + i * 2.3;
    slide.addShape(pres.shapes.RECTANGLE, { x: px, y: 3.52, w: 2.18, h: 0.52, fill: { color: p.color }, line: { color: JH.techBlue, width: 0.75 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: px, y: 3.52, w: 0.07, h: 0.52, fill: { color: JH.techBlue }, line: { color: JH.techBlue } });
    slide.addText(p.name, {
      x: px + 0.12, y: 3.52, w: 2.03, h: 0.52,
      fontSize: 8.5, bold: true, color: JH.white, fontFace: "Calibri", valign: "middle",
    });
  });

  slide.addText("For internal use  \u00b7  Executive briefing", {
    x: 0.55, y: 4.58, w: 5, h: 0.26,
    fontSize: 8.5, color: JH.mdGray, fontFace: "Calibri", italic: true,
  });
}


// ═══════════════════════════════════════════════════════════════
//  SLIDE 2 — Team / Org Structure  (3-level hierarchy)
// ═══════════════════════════════════════════════════════════════
function addTeamSlide() {
  const slide = pres.addSlide();
  slide.background = { color: JH.ivory };
  headerBar(slide, "JHBI Analytics Team", "FY27 Leadership Structure");
  footerBar(slide);

  // ── Column layout constants ─────────────────────────────────
  const MARGIN  = 0.1;
  const COL_GAP = 0.12;
  const COL_W   = (10 - 2 * MARGIN - 2 * COL_GAP) / 3;   // ≈ 3.19"
  const COL_X   = [0, 1, 2].map(i => MARGIN + i * (COL_W + COL_GAP));
  const COL_CX  = COL_X.map(x => x + COL_W / 2);
  const COL_COLORS = [JH.cobalt, JH.teal, JH.green];

  // ── Level 1: Mike Ames ──────────────────────────────────────
  const L1_W = 2.6, L1_H = 0.64;
  const L1_X = (10 - L1_W) / 2, L1_Y = 0.63;
  const L1_CX = L1_X + L1_W / 2, L1_BOT = L1_Y + L1_H;

  slide.addShape(pres.shapes.RECTANGLE, { x: L1_X, y: L1_Y, w: L1_W, h: L1_H, fill: { color: JH.navy }, line: { color: JH.navy } });
  slide.addShape(pres.shapes.RECTANGLE, { x: L1_X, y: L1_Y, w: L1_W, h: 0.055, fill: { color: JH.cobalt }, line: { color: JH.cobalt } });
  slide.addText("Mike Ames", {
    x: L1_X + 0.1, y: L1_Y + 0.06, w: L1_W - 0.2, h: 0.26,
    fontSize: 12, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
  });
  slide.addText("Head of Data Science", {
    x: L1_X + 0.1, y: L1_Y + 0.34, w: L1_W - 0.2, h: 0.22,
    fontSize: 9, color: JH.techBlue, fontFace: "Calibri", align: "center", valign: "middle",
  });

  // ── Net New Headcount callout (top-right, beside Mike Ames) ─
  const NB_X = L1_X + L1_W + 0.22, NB_Y = L1_Y;
  const NB_W = 9.82 - NB_X, NB_H = 0.98;
  const THIRD = NB_W / 3;
  // Card shell
  slide.addShape(pres.shapes.RECTANGLE, { x: NB_X, y: NB_Y, w: NB_W, h: NB_H, fill: { color: JH.white }, line: { color: JH.ltGray, width: 0.75 } });
  // Navy header strip
  slide.addShape(pres.shapes.RECTANGLE, { x: NB_X, y: NB_Y, w: NB_W, h: 0.20, fill: { color: JH.navy }, line: { color: JH.navy } });
  slide.addText("NET NEW HEADCOUNT", {
    x: NB_X, y: NB_Y, w: NB_W, h: 0.20,
    fontSize: 6.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });
  // Three stat columns: Software Dev | Management | Data Science
  const NB_STATS = [
    { label: "Software Dev", count: "7", color: JH.cobalt },
    { label: "Management",   count: "2", color: JH.teal   },
    { label: "Data Science", count: "5", color: JH.green  },
  ];
  const STAT_H = NB_H - 0.20 - 0.38;  // space between header and two-row FY split band
  NB_STATS.forEach(({ label, count, color }, si) => {
    const sx = NB_X + si * THIRD;
    if (si > 0) vLine(slide, sx, NB_Y + 0.20, NB_Y + NB_H - 0.24, JH.ltGray, 0.5);
    slide.addText(count, {
      x: sx + 0.04, y: NB_Y + 0.20, w: THIRD * 0.40, h: STAT_H,
      fontSize: 22, bold: true, color, fontFace: "Calibri", align: "center", valign: "middle",
    });
    slide.addText(label, {
      x: sx + THIRD * 0.42, y: NB_Y + 0.20, w: THIRD * 0.54, h: STAT_H,
      fontSize: 7, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
    });
  });
  // FY26 sub-band (1 position — MLOps Engineer)
  const FY26_Y = NB_Y + NB_H - 0.38;
  slide.addShape(pres.shapes.RECTANGLE, { x: NB_X, y: FY26_Y, w: NB_W, h: 0.18, fill: { color: JH.teal, transparency: 75 }, line: { color: JH.ltGray, width: 0.5 } });
  slide.addText("FY26  ·  1 Position  (MLOps Engineer)", {
    x: NB_X, y: FY26_Y, w: NB_W, h: 0.18,
    fontSize: 7, bold: false, color: JH.teal, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });
  // FY27 sub-band (13 positions)
  const FY27_Y = NB_Y + NB_H - 0.20;
  slide.addShape(pres.shapes.RECTANGLE, { x: NB_X, y: FY27_Y, w: NB_W, h: 0.20, fill: { color: JH.navy }, line: { color: JH.navy, width: 0 } });
  slide.addText("FY27  ·  13 Positions", {
    x: NB_X, y: FY27_Y, w: NB_W, h: 0.20,
    fontSize: 8.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });

  // ── Connectors L1 → L2 ─────────────────────────────────────
  const L2_Y  = 1.58;
  const JOIN_Y = L1_BOT + (L2_Y - L1_BOT) * 0.5;
  vLine(slide, L1_CX, L1_BOT, JOIN_Y, JH.mdGray, 1.5);
  hLine(slide, COL_CX[0], COL_CX[2], JOIN_Y, JH.mdGray, 1.5);
  COL_CX.forEach(cx => vLine(slide, cx, JOIN_Y, L2_Y, JH.mdGray, 1.5));

  // ── Level 2: 3 managers ─────────────────────────────────────
  const L2_H = 0.64, L2_BOT = L2_Y + L2_H;
  const L2_DATA = [
    { name: "Michael Goolsby", title: "Director, Software Development"   },
    { name: "[TBD]",           title: "Mgr, JHBI Product Mgmt & GTM"     },
    { name: "[TBD]",           title: "Data Science Lead"                 },
  ];
  L2_DATA.forEach((m, i) => {
    const x = COL_X[i], cc = COL_COLORS[i];
    slide.addShape(pres.shapes.RECTANGLE, { x, y: L2_Y, w: COL_W, h: L2_H, fill: { color: cc }, line: { color: cc } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y: L2_Y, w: COL_W, h: 0.055, fill: { color: JH.white, transparency: 40 }, line: { color: JH.white, transparency: 40 } });
    slide.addText(m.name, {
      x: x + 0.1, y: L2_Y + 0.06, w: COL_W - 0.2, h: 0.26,
      fontSize: 10.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
    });
    slide.addText(m.title, {
      x: x + 0.1, y: L2_Y + 0.35, w: COL_W - 0.2, h: 0.22,
      fontSize: 8, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
    });
  });

  // ── Level 3: Direct reports ─────────────────────────────────
  const L3_Y0   = L2_BOT + 0.14;
  const L3_H    = 0.62;
  const L3_GAP  = 0.06;
  const L3_DATA = [
    // Under Goolsby — open positions by function
    [
      { name: "Data Acq. & Engineering", title: "3 Open Positions", open: true, pillars: [0] },
      { name: "Visualization",           title: "2 Open Positions", open: true, pillars: [1] },
      { name: "App Dev",                 title: "2 Open Positions", open: true, pillars: [2] },
    ],
    // Under Product TBD — 3 named PMs
    [
      { name: "Elspeth Bloodgood", title: "PM, Go-to-Market & Visualization",  tbd: false, pillars: [1] },
      { name: "Ashley Greenhaw",   title: "PM, Application Development",        tbd: false, pillars: [2] },
      { name: "Mike Saunders",     title: "PM, Data Acq. & Engineering",        tbd: false, pillars: [0] },
    ],
    // Under Data Science TBD — open roles
    [
      { name: "Data Scientist",   title: "4 Open Positions",   open: true, pillars: [3] },
      { name: "MLOps Engineer",   title: "1 Open Position",    open: true, fy: "FY26", pillars: [3] },
    ],
  ];

  L3_DATA.forEach((reports, ci) => {
    const cc  = COL_COLORS[ci];
    const col_cx = COL_CX[ci];

    // Drop line from L2 to L3 start
    vLine(slide, col_cx, L2_BOT, L3_Y0, cc, 1.2);

    reports.forEach((rep, ri) => {
      const ry    = L3_Y0 + ri * (L3_H + L3_GAP);
      const cardX = COL_X[ci] + 0.06;
      const cardW = COL_W - 0.12;

      if (rep.open) {
        // Open position — outlined card with count badge
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX, y: ry, w: cardW, h: L3_H, fill: { color: JH.ltGray, transparency: 20 }, line: { color: cc, width: 0.75 } });
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX, y: ry, w: 0.055, h: L3_H, fill: { color: cc, transparency: 30 }, line: { color: cc, transparency: 30 } });
        slide.addText(rep.name, {
          x: cardX + 0.1, y: ry + 0.04, w: cardW - 0.52, h: 0.22,
          fontSize: 8.5, bold: true, color: JH.navy, fontFace: "Calibri", valign: "middle",
        });
        // Count badge
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX + cardW - 0.44, y: ry + 0.06, w: 0.38, h: 0.20, fill: { color: cc }, line: { color: cc } });
        slide.addText(rep.title.split(" ")[0], {
          x: cardX + cardW - 0.44, y: ry + 0.06, w: 0.38, h: 0.20,
          fontSize: 7.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
        });
        slide.addText("Open Position" + (rep.title.startsWith("1") ? "" : "s") + (rep.fy ? ` · ${rep.fy}` : ""), {
          x: cardX + 0.1, y: ry + 0.26, w: cardW - 0.16, h: 0.18,
          fontSize: 7, color: rep.fy ? JH.teal : JH.dkGray, fontFace: "Calibri", italic: true, valign: "middle",
        });
      } else if (rep.tbd) {
        // TBD lead — muted style
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX, y: ry, w: cardW, h: L3_H, fill: { color: JH.white }, line: { color: JH.mdGray, width: 0.5 } });
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX, y: ry, w: 0.055, h: L3_H, fill: { color: cc, transparency: 20 }, line: { color: cc, transparency: 20 } });
        slide.addText("[TBD]", {
          x: cardX + 0.1, y: ry + 0.04, w: cardW - 0.16, h: 0.20,
          fontSize: 8, bold: true, color: JH.mdGray, fontFace: "Calibri", italic: true, valign: "middle",
        });
        slide.addText(rep.title, {
          x: cardX + 0.1, y: ry + 0.26, w: cardW - 0.16, h: 0.20,
          fontSize: 8, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
        });
      } else {
        // Named hire — full style
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX, y: ry, w: cardW, h: L3_H, fill: { color: JH.white }, line: { color: JH.mdGray, width: 0.5 } });
        slide.addShape(pres.shapes.RECTANGLE, { x: cardX, y: ry, w: 0.055, h: L3_H, fill: { color: cc }, line: { color: cc } });
        slide.addText(rep.name, {
          x: cardX + 0.1, y: ry + 0.04, w: cardW - 0.16, h: 0.22,
          fontSize: 8.5, bold: true, color: JH.navy, fontFace: "Calibri", valign: "middle",
        });
        slide.addText(rep.title, {
          x: cardX + 0.1, y: ry + 0.26, w: cardW - 0.16, h: 0.20,
          fontSize: 7.5, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
        });
      }

      // ── Pillar tag pills (shared across all card types) ────────
      if (rep.pillars && rep.pillars.length > 0) {
        const TAG_Y = ry + L3_H - 0.14;
        let tx = cardX + 0.1;
        rep.pillars.forEach(pi => {
          const pp = PILLARS[pi];
          const tw = pp.short.length * 0.048 + 0.16;
          slide.addShape(pres.shapes.RECTANGLE, { x: tx, y: TAG_Y, w: tw, h: 0.11, fill: { color: pp.color, transparency: 70 }, line: { color: pp.color, width: 0.5 } });
          slide.addText(pp.short, { x: tx, y: TAG_Y, w: tw, h: 0.11, fontSize: 5.5, bold: true, color: pp.color, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
          tx += tw + 0.07;
        });
      }
    });
  });

  // ── Pillar Ownership Row ─────────────────────────────────────
  const PR_Y   = 4.52;
  const PR_H   = 0.36;
  const PR_GAP = 0.12;
  const PR_W   = (10 - 0.4 * 2 - PR_GAP * 3) / 4;
  slide.addText("PILLAR OWNERSHIP", {
    x: 0.4, y: PR_Y - 0.18, w: 3, h: 0.16,
    fontSize: 6, bold: true, color: JH.mdGray, fontFace: "Calibri", valign: "middle", characterSpacing: 1,
  });
  PILLARS.forEach((p, pi) => {
    const px = 0.4 + pi * (PR_W + PR_GAP);
    slide.addShape(pres.shapes.RECTANGLE, { x: px, y: PR_Y, w: PR_W, h: PR_H, fill: { color: p.color, transparency: 18 }, line: { color: p.color, width: 0 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: px, y: PR_Y, w: PR_W, h: 0.05, fill: { color: JH.white, transparency: 50 }, line: { color: JH.white, transparency: 50 } });
    slide.addText(p.name, {
      x: px, y: PR_Y, w: PR_W, h: PR_H,
      fontSize: 8.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
    });
  });

  // Legend
  const LEG_Y = 5.06;
  [
    { color: JH.cobalt, label: "Named hire" },
    { color: JH.mdGray, label: "[TBD] — planned role" },
    { color: JH.ltGray, label: "Open position" },
  ].forEach((li, idx) => {
    const lx = 0.4 + idx * 2.6;
    slide.addShape(pres.shapes.RECTANGLE, { x: lx, y: LEG_Y + 0.03, w: 0.18, h: 0.14, fill: { color: li.color }, line: { color: li.color } });
    slide.addText(li.label, { x: lx + 0.24, y: LEG_Y, w: 2.3, h: 0.22, fontSize: 7.5, color: JH.dkGray, fontFace: "Calibri", valign: "middle" });
  });
}


// ═══════════════════════════════════════════════════════════════
//  SLIDE 3 — Program Overview (2×2 pillar grid)
// ═══════════════════════════════════════════════════════════════
function addProgramOverviewSlide() {
  const slide = pres.addSlide();
  slide.background = { color: JH.ivory };
  headerBar(slide, "JHBI Analytics — Program Overview", "Strategic Pillars \u00b7 FY27");
  footerBar(slide);

  const CARD_W = 4.62, CARD_H = 2.26;
  const GAP_X = 0.2, GAP_Y = 0.18;
  const START_X = (10 - 2 * CARD_W - GAP_X) / 2;
  const START_Y = 0.65;

  PILLARS.forEach((p, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = START_X + col * (CARD_W + GAP_X);
    const y = START_Y + row * (CARD_H + GAP_Y);

    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: CARD_W, h: CARD_H, fill: { color: JH.white }, line: { color: JH.ltGray, width: 0.75 }, shadow: mkShadow() });
    // Left accent strip
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: CARD_H, fill: { color: p.color }, line: { color: p.color } });
    // Header band
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: CARD_W, h: 0.46, fill: { color: p.color }, line: { color: p.color } });
    slide.addText(p.name, {
      x: x + 0.16, y, w: CARD_W - 0.2, h: 0.46,
      fontSize: 13, bold: true, color: JH.white, fontFace: "Calibri", valign: "middle",
    });

    DELIVERABLES[i].forEach((d, j) => {
      const dy = y + 0.54 + j * 0.5;
      // Dot and text share the same vertical center line
      const rowCenter = dy + 0.20;
      slide.addShape(pres.shapes.OVAL, { x: x + 0.2, y: rowCenter - 0.06, w: 0.12, h: 0.12, fill: { color: p.color }, line: { color: p.color } });
      slide.addText(d, {
        x: x + 0.38, y: rowCenter - 0.16, w: CARD_W - 0.52, h: 0.32,
        fontSize: 9.5, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
      });
    });

    // Source system tags — Data Acquisition card only
    if (i === 0) {
      const TAG_Y = y + 1.98, TAG_H = 0.18;
      hLine(slide, x + 0.16, x + CARD_W - 0.16, TAG_Y - 0.06, JH.ltGray, 0.5);
      const SOURCES = ["SilverLake", "Symitar", "Banno", "Zelle", "ACH", "NCUA 5300"];
      let tx = x + 0.16;
      SOURCES.forEach(src => {
        const tw = src.length * 0.050 + 0.16;  // approx width at 6pt bold
        slide.addShape(pres.shapes.RECTANGLE, { x: tx, y: TAG_Y, w: tw, h: TAG_H, fill: { color: p.color, transparency: 72 }, line: { color: p.color, width: 0.5 } });
        slide.addText(src, {
          x: tx, y: TAG_Y, w: tw, h: TAG_H,
          fontSize: 6.5, bold: true, color: p.color, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
        });
        tx += tw + 0.07;
      });
    }
  });
}


// ═══════════════════════════════════════════════════════════════
//  SLIDE 4 — Quarterly Roadmap  (Foundation + FY27 Q1–Q4)
// ═══════════════════════════════════════════════════════════════
function addQuarterlyRoadmapSlide() {
  const slide = pres.addSlide();
  slide.background = { color: "FAFBFC" };
  headerBar(slide, "JHBI Analytics — Quarterly Roadmap", "FY26 H2 Foundation + FY27  \u00b7  Jan 2026 \u2013 Jun 2027");
  footerBar(slide);

  const MARGIN_L = 0.1;
  const LABEL_W  = 1.0;
  const COL_GAP  = 0.08;
  const FOUND_W  = 1.40;
  // Remaining width split 4 ways for Q1–Q4
  const COL_W    = (10 - MARGIN_L - LABEL_W - FOUND_W - 5 * COL_GAP - 0.1) / 4;  // ~1.73"
  const QHEADER_Y = 0.60, QHEADER_H = 0.33;
  const ROW_Y0   = QHEADER_Y + QHEADER_H + 0.06;
  const ROW_GAP  = 0.07;
  const ROW_H    = (5.35 - ROW_Y0 - 3 * ROW_GAP) / 4;

  // ── Foundation column header ────────────────────────────────
  const FOUND_X = MARGIN_L + LABEL_W + COL_GAP;
  // Dark gray background
  slide.addShape(pres.shapes.RECTANGLE, { x: FOUND_X, y: QHEADER_Y, w: FOUND_W, h: QHEADER_H, fill: { color: JH.dkGray }, line: { color: JH.dkGray } });
  // Left label "FY26 H2"
  slide.addText("FY26 H2", {
    x: FOUND_X + 0.04, y: QHEADER_Y, w: FOUND_W * 0.50, h: QHEADER_H,
    fontSize: 8.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });
  // Right sub-label "Jan–Jun '26"
  slide.addText("Jan\u2013Jun '26", {
    x: FOUND_X + FOUND_W * 0.50, y: QHEADER_Y, w: FOUND_W * 0.50, h: QHEADER_H,
    fontSize: 7, color: JH.techBlue, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });

  // ── FY27 quarter column headers ────────────────────────────
  const Q_LABELS = ["Q1  FY27", "Q2  FY27", "Q3  FY27", "Q4  FY27"];
  const Q_MONTHS = ["Jul\u2013Sep '26", "Oct\u2013Dec '26", "Jan\u2013Mar '27", "Apr\u2013Jun '27"];
  const Q_COLORS = [JH.navy, JH.cobalt, JH.teal, JH.green];
  const Q_START_X = FOUND_X + FOUND_W + COL_GAP;
  Q_LABELS.forEach((ql, qi) => {
    const qx = Q_START_X + qi * (COL_W + COL_GAP);
    slide.addShape(pres.shapes.RECTANGLE, { x: qx, y: QHEADER_Y, w: COL_W, h: QHEADER_H, fill: { color: Q_COLORS[qi] }, line: { color: Q_COLORS[qi] } });
    slide.addText(ql, {
      x: qx, y: QHEADER_Y, w: COL_W * 0.52, h: QHEADER_H,
      fontSize: 8.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
    slide.addText(Q_MONTHS[qi], {
      x: qx + COL_W * 0.52, y: QHEADER_Y, w: COL_W * 0.48, h: QHEADER_H,
      fontSize: 6.5, color: JH.techBlue, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
  });

  // ── Pillar rows ─────────────────────────────────────────────
  PILLARS.forEach((p, pi) => {
    const ry = ROW_Y0 + pi * (ROW_H + ROW_GAP);

    // Alternating row tint
    if (pi % 2 === 1) {
      slide.addShape(pres.shapes.RECTANGLE, { x: MARGIN_L, y: ry, w: 9.8, h: ROW_H, fill: { color: JH.ltGray, transparency: 60 }, line: { color: JH.ltGray, width: 0 } });
    }

    // Pillar label tile
    slide.addShape(pres.shapes.RECTANGLE, { x: MARGIN_L, y: ry, w: LABEL_W, h: ROW_H, fill: { color: p.color }, line: { color: p.color } });
    slide.addText(p.name, {
      x: MARGIN_L, y: ry, w: LABEL_W, h: ROW_H,
      fontSize: 7.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
    });

    // Foundation card (in-progress style)
    const fcx = FOUND_X + 0.07, fcw = FOUND_W - 0.14;
    const fcy = ry + 0.09, fch = ROW_H - 0.18;
    // Card background — muted
    slide.addShape(pres.shapes.RECTANGLE, { x: fcx, y: fcy, w: fcw, h: fch, fill: { color: JH.white }, line: { color: JH.dkGray, width: 0.75 } });
    // Left color strip (semi-transparent)
    slide.addShape(pres.shapes.RECTANGLE, { x: fcx, y: fcy, w: 0.055, h: fch, fill: { color: p.color, transparency: 30 }, line: { color: p.color, transparency: 30 } });
    // Top gray band
    slide.addShape(pres.shapes.RECTANGLE, { x: fcx, y: fcy, w: fcw, h: 0.055, fill: { color: JH.dkGray, transparency: 20 }, line: { color: JH.dkGray, transparency: 20 } });
    // "▶ Active" badge
    slide.addShape(pres.shapes.RECTANGLE, { x: fcx + 0.1, y: fcy + 0.07, w: 0.52, h: 0.16, fill: { color: p.color }, line: { color: p.color } });
    slide.addText("\u25b6 Active", {
      x: fcx + 0.1, y: fcy + 0.07, w: 0.52, h: 0.16,
      fontSize: 6, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
    // Foundation initiative text
    slide.addText(Q_FOUNDATION[pi], {
      x: fcx + 0.1, y: fcy + 0.28, w: fcw - 0.16, h: fch - 0.33,
      fontSize: 7.5, color: JH.dkGray, fontFace: "Calibri", valign: "top",
    });

    // FY27 quarter cards
    Q_INITIATIVES.forEach((qArr, qi) => {
      const qx = Q_START_X + qi * (COL_W + COL_GAP);
      const cx = qx + 0.07, cw = COL_W - 0.14;
      const cy = ry + 0.09, ch = ROW_H - 0.18;

      slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cw, h: ch, fill: { color: JH.white }, line: { color: p.color, width: 1 }, shadow: mkShadowSm() });
      // Top color strip
      slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cw, h: 0.055, fill: { color: p.color }, line: { color: p.color } });
      slide.addText(qArr[pi], {
        x: cx + 0.09, y: cy + 0.09, w: cw - 0.16, h: ch - 0.13,
        fontSize: 8.5, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
      });
    });
  });
}


// ═══════════════════════════════════════════════════════════════
//  SLIDE 5 — Gantt Chart  (18-month: Jan 2026 – Jun 2027)
// ═══════════════════════════════════════════════════════════════
function addGanttSlide() {
  const slide = pres.addSlide();
  slide.background = { color: JH.ivory };
  headerBar(slide, "JHBI Analytics — Gantt Chart", "18-Month View  \u00b7  Jan 2026 \u2013 Jun 2027");
  footerBar(slide);

  // 18 months: Jan 2026 (idx 0) → Jun 2027 (idx 17)
  const MONTHS   = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"];
  const LABEL_X  = 0.15, LABEL_W = 2.18;
  const GRID_X   = LABEL_X + LABEL_W + 0.05;
  const GRID_W   = 10 - GRID_X - 0.1;
  const CELL_W   = GRID_W / 18;
  const HEADER_Y = 0.62, HEADER_H = 0.37;
  const ROW_Y0   = HEADER_Y + HEADER_H + 0.05;
  const SECTION_H = 0.17, ROW_H = 0.25, SECTION_GAP = 0.04, ROW_GAP = 0.02;

  // TODAY: March 30, 2026 = month index 2 + 29/31 days elapsed
  const TODAY_MI = 2 + 29 / 31;
  const TODAY_X  = GRID_X + TODAY_MI * CELL_W;

  // Foundation zone shading (months 0–5, inclusive)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: GRID_X, y: HEADER_Y + HEADER_H,
    w: 6 * CELL_W, h: 5.35 - (HEADER_Y + HEADER_H),
    fill: { color: "D8DBE0", transparency: 55 },
    line: { color: "D8DBE0", width: 0 },
  });

  // Phase header bands
  const PHASES = [
    { mi: 0,  span: 6, label: "FY26 H2  \u00b7  Foundation", color: JH.dkGray },
    { mi: 6,  span: 3, label: "Q1  FY27",                    color: JH.navy   },
    { mi: 9,  span: 3, label: "Q2  FY27",                    color: JH.cobalt },
    { mi: 12, span: 3, label: "Q3  FY27",                    color: JH.teal   },
    { mi: 15, span: 3, label: "Q4  FY27",                    color: JH.green  },
  ];
  PHASES.forEach(({ mi, span, label, color }) => {
    const qx = GRID_X + mi * CELL_W, qw = CELL_W * span;
    slide.addShape(pres.shapes.RECTANGLE, { x: qx, y: HEADER_Y, w: qw, h: HEADER_H * 0.5, fill: { color }, line: { color: JH.white, width: 0.5 } });
    slide.addText(label, {
      x: qx + 0.02, y: HEADER_Y, w: qw - 0.04, h: HEADER_H * 0.5,
      fontSize: 7, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
  });

  // Month sub-headers
  const MCOLS = [
    JH.dkGray,JH.dkGray,JH.dkGray,JH.dkGray,JH.dkGray,JH.dkGray,
    JH.navy,JH.navy,JH.navy,
    JH.cobalt,JH.cobalt,JH.cobalt,
    JH.teal,JH.teal,JH.teal,
    JH.green,JH.green,JH.green,
  ];
  MONTHS.forEach((m, mi) => {
    const mx = GRID_X + mi * CELL_W;
    slide.addShape(pres.shapes.RECTANGLE, { x: mx, y: HEADER_Y + HEADER_H * 0.5, w: CELL_W, h: HEADER_H * 0.5, fill: { color: MCOLS[mi], transparency: 28 }, line: { color: JH.white, width: 0.4 } });
    slide.addText(m, {
      x: mx, y: HEADER_Y + HEADER_H * 0.5, w: CELL_W, h: HEADER_H * 0.5,
      fontSize: 5.5, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
  });

  // Vertical grid lines — phase boundaries are heavier
  const PHASE_BOUNDS = new Set([0, 6, 9, 12, 15, 18]);
  for (let mi = 0; mi <= 18; mi++) {
    const lx = GRID_X + mi * CELL_W;
    const heavy = PHASE_BOUNDS.has(mi);
    vLine(slide, lx, HEADER_Y + HEADER_H, 5.35, heavy ? JH.dkGray : JH.mdGray, heavy ? 1.0 : 0.4);
  }

  // Gantt rows, grouped by pillar
  let curY = ROW_Y0;
  PILLARS.forEach((p, pi) => {
    const allItems = GANTT.filter(g => g.p === pi);   // foundation item first, then FY27

    // Section header band
    slide.addShape(pres.shapes.RECTANGLE, { x: LABEL_X, y: curY, w: 9.75, h: SECTION_H, fill: { color: p.color, transparency: 10 }, line: { color: p.color, width: 0 } });
    slide.addText(p.name.toUpperCase(), {
      x: LABEL_X + 0.1, y: curY, w: 9.5, h: SECTION_H,
      fontSize: 7, bold: true, color: p.text, fontFace: "Calibri", valign: "middle", charSpacing: 0.5,
    });
    curY += SECTION_H;

    allItems.forEach((item, rowIdx) => {
      // Alternating row tint
      slide.addShape(pres.shapes.RECTANGLE, {
        x: LABEL_X, y: curY, w: 9.75, h: ROW_H,
        fill: { color: rowIdx % 2 === 0 ? JH.ltGray : JH.white, transparency: rowIdx % 2 === 0 ? 70 : 0 },
        line: { color: "E0E4E8", width: 0.3 },
      });

      // Row label
      slide.addText(item.name, {
        x: LABEL_X + 0.06, y: curY + 0.01, w: LABEL_W - 0.08, h: ROW_H - 0.02,
        fontSize: 6.5, color: JH.dkGray, fontFace: "Calibri", valign: "middle",
      });

      // Gantt bar
      const barX = GRID_X + item.s * CELL_W + 0.02;
      const barW = (item.e - item.s + 1) * CELL_W - 0.04;
      const barY = curY + 0.04, barH = ROW_H - 0.08;

      if (item.foundation) {
        // Foundation bar — muted fill + "▶ In Progress" label
        slide.addShape(pres.shapes.RECTANGLE, { x: barX, y: barY, w: barW, h: barH, fill: { color: p.color, transparency: 40 }, line: { color: p.color, width: 1.0 } });
        slide.addText("\u25b6 In Progress", {
          x: barX + 0.04, y: barY, w: barW - 0.06, h: barH,
          fontSize: 6, bold: true, color: p.color, fontFace: "Calibri", valign: "middle",
        });
      } else {
        // FY27 bar — solid fill
        slide.addShape(pres.shapes.RECTANGLE, { x: barX, y: barY, w: barW, h: barH, fill: { color: p.color, transparency: 12 }, line: { color: p.color, width: 1.2 } });
        const shortName = item.name.length > 24 ? item.name.slice(0, 22) + "\u2026" : item.name;
        slide.addText(shortName, {
          x: barX + 0.04, y: barY, w: barW - 0.06, h: barH,
          fontSize: 6, color: p.text, fontFace: "Calibri", valign: "middle",
        });
      }

      curY += ROW_H + ROW_GAP;
    });

    curY += SECTION_GAP;
  });

  // TODAY marker — red vertical line + label in phase header zone
  vLine(slide, TODAY_X, HEADER_Y, 5.35, "D03030", 1.5);
  slide.addText("TODAY", {
    x: TODAY_X - 0.24, y: HEADER_Y + 0.005, w: 0.48, h: HEADER_H * 0.5,
    fontSize: 5.5, bold: true, color: "D03030", fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });
}


// ═══════════════════════════════════════════════════════════════
//  SLIDE 6 — Phase 1 Financial Impact & ARR Scaling
// ═══════════════════════════════════════════════════════════════
function addFinancialImpactSlide() {
  const slide = pres.addSlide();
  slide.background = { color: JH.ivory };
  headerBar(slide, "JHBI Analytics — Phase 1 Business Case", "FI Value  \u00b7  ARR Scaling  \u00b7  8,000+ FI Clients");
  footerBar(slide);

  // ── 5 application value cards ───────────────────────────────
  const APPS = [
    { name: "Churn\nMediation",        value: "$1\u20133M",      unit: "/yr per FI", metric: "15\u201325% churn reduction",       launch: "Q1 FY27",  color: JH.teal   },
    { name: "Zelle Memo\nIntelligence", value: "$250\u2013600K",  unit: "/yr per FI", metric: "60% manual review reduction",      launch: "Q2 FY27",  color: JH.cobalt },
    { name: "Anomaly\nDetection",       value: "$1\u20132M",      unit: "/yr per FI", metric: "70% false-positive reduction",     launch: "Q2 FY27",  color: JH.green  },
    { name: "Call\nReport AI",          value: "$100\u2013300K",  unit: "/yr per FI", metric: "Board prep & peer benchmarking",   launch: "Q1 FY27",  color: JH.navy   },
    { name: "Account\nOpening LTV",     value: "$500K\u20131.5M", unit: "/yr per FI", metric: "3\u20135\u00d7 cross-sell lift",  launch: "Q4 FY27",  color: JH.teal   },
  ];

  const CARD_W = 1.76, CARD_H = 2.22, CARD_GAP = 0.10;
  const CARDS_X0 = (10 - APPS.length * CARD_W - (APPS.length - 1) * CARD_GAP) / 2;
  const CARD_Y  = 0.64;

  APPS.forEach((app, i) => {
    const cx = CARDS_X0 + i * (CARD_W + CARD_GAP);

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: CARD_Y, w: CARD_W, h: CARD_H, fill: { color: JH.white }, line: { color: JH.ltGray, width: 0.75 }, shadow: mkShadow() });
    // Top color header band
    slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: CARD_Y, w: CARD_W, h: 0.44, fill: { color: app.color }, line: { color: app.color } });
    // App name in header
    slide.addText(app.name, {
      x: cx + 0.06, y: CARD_Y + 0.03, w: CARD_W - 0.12, h: 0.40,
      fontSize: 9.5, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
    });

    // Value (large)
    slide.addText(app.value, {
      x: cx + 0.06, y: CARD_Y + 0.50, w: CARD_W - 0.12, h: 0.50,
      fontSize: 19, bold: true, color: app.color, fontFace: "Calibri", align: "center", valign: "middle",
    });
    // Unit label
    slide.addText(app.unit, {
      x: cx + 0.06, y: CARD_Y + 0.99, w: CARD_W - 0.12, h: 0.20,
      fontSize: 7.5, color: JH.mdGray, fontFace: "Calibri", align: "center", valign: "middle",
    });

    // Divider
    hLine(slide, cx + 0.14, cx + CARD_W - 0.14, CARD_Y + 1.22, JH.ltGray, 0.5);

    // Key metric text
    slide.addText(app.metric, {
      x: cx + 0.08, y: CARD_Y + 1.26, w: CARD_W - 0.16, h: 0.48,
      fontSize: 8, color: JH.dkGray, fontFace: "Calibri", align: "center", valign: "top",
    });

    // Launch badge
    slide.addShape(pres.shapes.RECTANGLE, { x: cx + 0.08, y: CARD_Y + CARD_H - 0.28, w: CARD_W - 0.16, h: 0.22, fill: { color: app.color, transparency: 14 }, line: { color: app.color, width: 0 } });
    slide.addText(app.launch, {
      x: cx + 0.08, y: CARD_Y + CARD_H - 0.28, w: CARD_W - 0.16, h: 0.22,
      fontSize: 7.5, bold: true, color: app.color, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
  });

  // ── Combined value callout banner ───────────────────────────
  const COMB_Y = CARD_Y + CARD_H + 0.14;
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: COMB_Y, w: 9.2, h: 0.46, fill: { color: JH.navy }, line: { color: JH.navy } });
  slide.addText("$3\u20137.5M / year combined value per FI client \u00b7 conservative estimate for a mid-size $1B deposits community bank", {
    x: 0.4, y: COMB_Y, w: 9.2, h: 0.46,
    fontSize: 10, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle",
  });

  // ── ARR Scaling at full JH platform adoption ────────────────
  const ARR_Y = COMB_Y + 0.56;
  const ARR_DATA = [
    { label: "500 FIs",    arr: "$19.5M ARR",  sub: "6% of JH client base",  color: JH.cobalt },
    { label: "2,000 FIs",  arr: "$78M ARR",    sub: "25% of JH client base", color: JH.teal   },
    { label: "8,000+ FIs", arr: "$312M+ ARR",  sub: "Full JH client base",   color: JH.green  },
  ];
  const BOX_W = 2.8, BOX_H = 0.74, BOX_GAP = 0.3;
  const ARR_X0 = (10 - ARR_DATA.length * BOX_W - (ARR_DATA.length - 1) * BOX_GAP) / 2;

  ARR_DATA.forEach((a, i) => {
    const bx = ARR_X0 + i * (BOX_W + BOX_GAP);
    slide.addShape(pres.shapes.RECTANGLE, { x: bx, y: ARR_Y, w: BOX_W, h: BOX_H, fill: { color: a.color, transparency: 8 }, line: { color: a.color, width: 0 } });
    slide.addText(a.arr, {
      x: bx, y: ARR_Y + 0.05, w: BOX_W, h: 0.38,
      fontSize: 18, bold: true, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
    slide.addText(a.label + "  \u00b7  " + a.sub, {
      x: bx, y: ARR_Y + 0.42, w: BOX_W, h: 0.26,
      fontSize: 8, color: JH.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
  });

  // Attribution note
  slide.addText("ARR estimates based on avg. subscription value across tiers. Adoption rates modeled at 6%, 25%, and 100% of JH\u2019s 8,000+ FI client base. Sources: Mordor Intelligence, ICBA / Cornerstone 2025 survey.", {
    x: 0.4, y: ARR_Y + BOX_H + 0.09, w: 9.2, h: 0.22,
    fontSize: 6.5, color: JH.mdGray, fontFace: "Calibri", italic: true, align: "center",
  });
}


// ═══════════════════════════════════════════════════════════════
//  SLIDE 7 — Alt Layout: Now / Next / Later
// ═══════════════════════════════════════════════════════════════
function addNowNextLaterSlide() {
  const slide = pres.addSlide();
  slide.background = { color: "FAFBFC" };
  headerBar(slide, "JHBI Analytics — FY27 Quarterly View", "Q1\u2013Q4 FY27  |  Alternate Layout");
  footerBar(slide);

  const COLS = [
    { label: "Q1", sub: "Jul\u2013Sep '26", color: JH.navy,   qi: 0 },
    { label: "Q2", sub: "Oct\u2013Dec '26", color: JH.cobalt, qi: 1 },
    { label: "Q3", sub: "Jan\u2013Mar '27", color: JH.teal,   qi: 2 },
    { label: "Q4", sub: "Apr\u2013Jun '27", color: JH.green,  qi: 3 },
  ];

  const COL_W = 2.22, COL_GAP = 0.1;
  const START_X = (10 - 4 * COL_W - 3 * COL_GAP) / 2;
  const COL_HEADER_H = 0.48;
  const COL_Y = 0.63;

  COLS.forEach((col, ci) => {
    const cx = START_X + ci * (COL_W + COL_GAP);

    // Column header
    slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: COL_Y, w: COL_W, h: COL_HEADER_H, fill: { color: col.color }, line: { color: col.color } });
    slide.addText("FY27  " + col.label, {
      x: cx, y: COL_Y, w: COL_W * 0.46, h: COL_HEADER_H,
      fontSize: 12, bold: true, color: JH.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(col.sub, {
      x: cx + COL_W * 0.46, y: COL_Y, w: COL_W * 0.54, h: COL_HEADER_H,
      fontSize: 8, color: JH.techBlue, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0,
    });

    // Initiative cards (one per pillar)
    const CARD_H = 0.82, CARD_GAP = 0.08;
    const CARD_Y0 = COL_Y + COL_HEADER_H + 0.1;

    PILLARS.forEach((p, pi) => {
      const cardY = CARD_Y0 + pi * (CARD_H + CARD_GAP);
      const text = Q_INITIATIVES[col.qi][pi];

      slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cardY, w: COL_W, h: CARD_H, fill: { color: JH.white }, line: { color: p.color, width: 0.75 }, shadow: mkShadowSm() });
      slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cardY, w: 0.07, h: CARD_H, fill: { color: p.color }, line: { color: p.color } });
      slide.addText(p.name, {
        x: cx + 0.12, y: cardY + 0.04, w: COL_W - 0.16, h: 0.20,
        fontSize: 7, bold: true, color: p.color, fontFace: "Calibri", valign: "middle",
      });
      hLine(slide, cx + 0.12, cx + COL_W - 0.07, cardY + 0.26, JH.ltGray, 0.5);
      slide.addText(text, {
        x: cx + 0.12, y: cardY + 0.29, w: COL_W - 0.18, h: CARD_H - 0.34,
        fontSize: 8.5, color: JH.dkGray, fontFace: "Calibri", valign: "top",
      });
    });
  });
}
