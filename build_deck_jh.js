const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const {
  FaUserMinus, FaMobileAlt, FaExclamationTriangle,
  FaFileAlt, FaChartLine, FaDatabase, FaBrain,
  FaChartBar, FaCode, FaRocket, FaUsers,
  FaShieldAlt, FaMoneyBillWave, FaBalanceScale, FaHeadset,
  FaNetworkWired, FaSlidersH, FaBriefcase, FaUserFriends,
} = require("react-icons/fa");

// ============================================================
// JACK HENRY BRAND PALETTE
// ============================================================
const NAVY    = "06185F"; // Heritage Navy
const COBALT  = "085CE5"; // Vibrant Cobalt
const TECH    = "76DCFD"; // Tech Blue
const SKY_BG  = "E8F7F7"; // Open Sky
const IVORY   = "FEFDF8"; // Warm Ivory
const DK_GRAY = "575A5D"; // Dark Cool Gray
const MD_GRAY = "B6BBC0"; // Medium Cool Gray
const LT_GRAY = "E7ECF0"; // Light Cool Gray
const WHITE   = "FFFFFF";

// Blue-spectrum variants for app card variety (within JH brand family)
const COBALT_D  = "073FA8"; // deep cobalt (between navy and cobalt)
const COBALT_M  = "0B68D4"; // medium cobalt
const NAVY_MED  = "0D2E7A"; // lighter heritage navy

// ============================================================
// APP ACCENT COLORS (all within JH blue/gray spectrum)
// ============================================================
const A1 = COBALT;    // Vibrant Cobalt
const A2 = NAVY;      // Heritage Navy
const A3 = COBALT_D;  // Deep Cobalt
const A4 = DK_GRAY;   // Dark Cool Gray
const A5 = COBALT_M;  // Medium Cobalt
const A6 = NAVY_MED;  // Decision Studio — Lighter Navy
const A7 = "0A52C4";  // Rich Blue — CommercialSignal
const A8 = "0F3D8A";  // Deep Navy-Blue — Generational Wealth Deflection

const FONT_H = "Calibri";
const FONT_B = "Calibri";

// ============================================================
// HELPERS
// ============================================================
async function iconPng(Icon, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color: color || "#FFFFFF", size: "256" })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const mkSh = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.10 });

// ============================================================
// BUILD DECK
// ============================================================
async function buildDeck() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10" × 5.625"

  // ==========================================================
  // SLIDE 1 — TITLE
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };

    // Left Heritage Navy accent bar in Cobalt
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: COBALT }, line: { color: COBALT } });

    // Decorative circles — Tech Blue palette
    s.addShape(pres.shapes.OVAL, { x: 7.0, y: 2.3, w: 4.5, h: 4.5, fill: { color: COBALT, transparency: 85 }, line: { color: COBALT, transparency: 75 } });
    s.addShape(pres.shapes.OVAL, { x: 8.0, y: 3.1, w: 2.8, h: 2.8, fill: { color: TECH, transparency: 75 }, line: { color: TECH, transparency: 65 } });

    // Eyebrow
    s.addText("JHBI  ·  Data Science Platform Investment Brief", {
      x: 0.4, y: 0.55, w: 9.2, h: 0.35,
      fontSize: 11, fontFace: FONT_B, color: TECH, margin: 0, charSpacing: 1
    });

    // Main title
    s.addText("Data Science\nPlatform Strategy", {
      x: 0.4, y: 1.05, w: 8.0, h: 2.1,
      fontSize: 48, fontFace: FONT_H, bold: true, color: WHITE, margin: 0
    });

    // Subtitle
    s.addText("Building AI-Powered Applications for Jack Henry's Financial Institution Clients", {
      x: 0.4, y: 3.25, w: 8.4, h: 0.5,
      fontSize: 16, fontFace: FONT_B, color: TECH, margin: 0, italic: true
    });

    // Divider
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.88, w: 2.5, h: 0.05, fill: { color: COBALT }, line: { color: COBALT } });

    // Date / team
    s.addText("Data & Analytics Team  |  March 2026", {
      x: 0.4, y: 4.05, w: 6, h: 0.35,
      fontSize: 12, fontFace: FONT_B, color: MD_GRAY, margin: 0
    });

    // JH tagline + confidential
    s.addText("Powering the Financial World™", {
      x: 0.4, y: 5.25, w: 4, h: 0.26,
      fontSize: 9, fontFace: FONT_B, color: MD_GRAY, margin: 0, italic: true
    });
    s.addText("CONFIDENTIAL", {
      x: 7.8, y: 5.28, w: 2.0, h: 0.25,
      fontSize: 9, fontFace: FONT_B, color: MD_GRAY, align: "right", margin: 0, charSpacing: 1
    });
  }

  // ==========================================================
  // SLIDE 2 — THE BUSINESS CASE
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("The Market Opportunity for Jack Henry", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    s.addText("Jack Henry serves 1,660 financial institutions on core platforms and 7,500+ across its full ecosystem — institutions that collectively manage trillions in assets. By embedding AI-driven analytics directly into our platform, we can transform how every FI client makes decisions on retention, risk, compliance, and growth.", {
      x: 0.4, y: 0.85, w: 9.2, h: 0.62,
      fontSize: 12.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    const stats = [
      { num: "1,660", sub: "Core FI clients — banks & CUs\nwith deepest JH data access", color: COBALT, numColor: COBALT, icon: FaUsers },
      { num: "$6.2B", sub: "projected AI banking analytics\nmarket by 2028 (Mordor Intel.)", color: NAVY, numColor: NAVY, icon: FaChartLine },
      { num: "74%", sub: "of FIs cite AI/ML tools as a\ntop 3 technology priority*", color: COBALT_D, numColor: COBALT_D, icon: FaBrain },
    ];

    for (let i = 0; i < stats.length; i++) {
      const st = stats[i];
      const x = 0.4 + i * 3.08;
      const icoData = await iconPng(st.icon, `#${st.color}`);

      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: 2.9, h: 2.7, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.6, w: 2.9, h: 0.1, fill: { color: st.color }, line: { color: st.color } });
      s.addImage({ data: icoData, x: x + 1.18, y: 1.77, w: 0.52, h: 0.52 });
      s.addText(st.num, {
        x: x + 0.1, y: 2.34, w: 2.7, h: 0.88,
        fontSize: 42, fontFace: FONT_H, bold: true, color: st.numColor, align: "center", margin: 0
      });
      s.addText(st.sub, {
        x: x + 0.1, y: 3.24, w: 2.7, h: 0.75,
        fontSize: 11.5, fontFace: FONT_B, color: DK_GRAY, align: "center", margin: 0
      });
    }

    s.addText("* ICBA / Cornerstone Advisors 2025 What's Going On In Banking survey", {
      x: 0.4, y: 4.38, w: 9.2, h: 0.28,
      fontSize: 9, fontFace: FONT_B, color: MD_GRAY, margin: 0, italic: true
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.72, w: 9.2, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("We are requesting budget and headcount to build the team and platform to deliver this value — starting with our 1,660 core FI clients, expanding across the full JH ecosystem.", {
      x: 0.6, y: 4.76, w: 8.8, h: 0.64,
      fontSize: 13.5, fontFace: FONT_B, bold: true, color: WHITE, valign: "middle", margin: 0
    });
  }

  // ==========================================================
  // SLIDE — MARKET OPPORTUNITY
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: IVORY };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("The Real Market Opportunity — Grounded in JH's Actual Footprint", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 20, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    // Subtitle
    s.addText("Jack Henry's FY2024 10-K reports 1,660 FIs on core processing — 940 banks and 720 credit unions. The remaining ~5,870 clients use payment and complementary products only. This is the true addressable base for the AI platform.", {
      x: 0.4, y: 0.82, w: 9.2, h: 0.40,
      fontSize: 10, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Three segment cards
    const segments = [
      {
        color: COBALT,
        label: "CORE AI TARGET",
        title: "Core Platform FIs",
        num: "1,660",
        sub: "940 banks  ·  720 credit unions",
        detail: "Full data access: core transactions, deposits, lending, digital. Highest AI value — all 5 Phase 1 apps apply. JH has 25% of target bank market and ~17% of US CU market on core.",
        badge: "PRIMARY",
      },
      {
        color: COBALT_D,
        label: "PAYMENTS EXPANSION",
        title: "Payment / Complementary FIs",
        num: "5,870",
        sub: "Payments, digital, card, bill pay",
        detail: "Payment data only — no core relationship. Zelle Memo Intelligence and Anomaly Detection apply directly. Broader sales motion; lighter data integration per FI.",
        badge: "SECONDARY",
      },
      {
        color: DK_GRAY,
        label: "OPEN DATA MARKET",
        title: "All U.S. Credit Unions",
        num: "4,374",
        sub: "NCUA 5300 — public regulatory data",
        detail: "Call Report AI has no JH relationship requirement. Any of the 4,374 US credit unions with $2.5T in assets is addressable — JH relationship accelerates but doesn't gate distribution.",
        badge: "CALLRPT AI",
      },
    ];

    segments.forEach((seg, i) => {
      const x = 0.3 + i * 3.22;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.30, w: 3.1, h: 3.62, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.30, w: 3.1, h: 0.06, fill: { color: seg.color }, line: { color: seg.color } });

      // Badge
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.16, y: 1.42, w: 1.1, h: 0.22, fill: { color: seg.color }, line: { color: seg.color } });
      s.addText(seg.badge, { x: x + 0.16, y: 1.42, w: 1.1, h: 0.22, fontSize: 7.5, fontFace: FONT_H, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });

      // Title
      s.addText(seg.title, { x: x + 0.16, y: 1.70, w: 2.78, h: 0.28, fontSize: 12, fontFace: FONT_H, bold: true, color: NAVY, margin: 0 });

      // Big number
      s.addText(seg.num, { x: x + 0.16, y: 2.00, w: 2.78, h: 0.76, fontSize: 46, fontFace: FONT_H, bold: true, color: seg.color, margin: 0 });
      s.addText(seg.sub, { x: x + 0.16, y: 2.76, w: 2.78, h: 0.24, fontSize: 9, fontFace: FONT_B, color: MD_GRAY, margin: 0, italic: true });

      // Divider
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.16, y: 3.06, w: 2.78, h: 0.02, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });

      // Detail text
      s.addText(seg.detail, { x: x + 0.16, y: 3.14, w: 2.78, h: 1.60, fontSize: 9.5, fontFace: FONT_B, color: DK_GRAY, margin: 0, wrap: true });
    });

    // Bottom summary bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.00, w: 9.4, h: 0.50, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Phase 1 targets the 1,660 core FIs first — then expands payment apps to the 5,870 complementary clients and Call Report AI to all 4,374 U.S. CUs.", {
      x: 0.5, y: 5.03, w: 9.0, h: 0.44,
      fontSize: 11, fontFace: FONT_B, bold: true, color: WHITE, valign: "middle", margin: 0
    });
  }

  // ==========================================================
  // SLIDE 3 — STRATEGIC VISION (dark)
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };

    // Decorative circles
    s.addShape(pres.shapes.OVAL, { x: 5.8, y: -2.0, w: 7, h: 7, fill: { color: COBALT, transparency: 90 }, line: { color: COBALT, transparency: 82 } });
    s.addShape(pres.shapes.OVAL, { x: 7.2, y: 1.0, w: 3.5, h: 3.5, fill: { color: TECH, transparency: 88 }, line: { color: TECH, transparency: 80 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: COBALT }, line: { color: COBALT } });

    s.addText("OUR VISION", {
      x: 0.4, y: 0.5, w: 6, h: 0.38,
      fontSize: 11, fontFace: FONT_H, bold: true, color: TECH, margin: 0, charSpacing: 4
    });
    s.addText("Make Jack Henry the AI-First\nPlatform Partner for Every\nCommunity FI in America", {
      x: 0.4, y: 0.95, w: 7.0, h: 1.85,
      fontSize: 34, fontFace: FONT_H, bold: true, color: WHITE, margin: 0
    });
    s.addText("By embedding AI-powered analytics and decision intelligence directly into the JH platform ecosystem, we deepen client stickiness, create new revenue streams, and establish an insurmountable competitive moat versus FIS, Fiserv, and niche challengers.", {
      x: 0.4, y: 2.9, w: 6.4, h: 1.05,
      fontSize: 13, fontFace: FONT_B, color: LT_GRAY, margin: 0
    });

    s.addText("FOUR INVESTMENT PILLARS", {
      x: 0.4, y: 4.08, w: 6, h: 0.28,
      fontSize: 9, fontFace: FONT_B, bold: true, color: MD_GRAY, margin: 0, charSpacing: 1.5
    });

    const caps = [
      { label: "Data Acquisition", color: COBALT },
      { label: "DS Buildout",      color: NAVY_MED },
      { label: "Visualizations",   color: COBALT_D },
      { label: "App / API Dev",    color: DK_GRAY },
    ];
    caps.forEach((c, i) => {
      const x = 0.4 + i * 2.32;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.42, w: 2.18, h: 0.5, fill: { color: c.color, transparency: 65 }, line: { color: c.color } });
      s.addText(c.label, {
        x: x + 0.05, y: 4.42, w: 2.08, h: 0.5,
        fontSize: 11.5, fontFace: FONT_B, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0
      });
    });
  }

  // ==========================================================
  // SLIDE 4 — APPLICATION PORTFOLIO OVERVIEW
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("JH Platform AI — Six Core Applications", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    const apps = [
      { name: "Churn\nSentinel",       desc: "ML scoring to identify at-risk FI customers weekly and power proactive retention workflows before churn occurs.", color: A1, icon: FaUserMinus },
      { name: "Zelle Memo\nIntelligence",             desc: "NLP pipeline to classify Zelle memos for fraud patterns, elder abuse, and compliance flags — automatically.", color: A2, icon: FaMobileAlt },
      { name: "Anomaly\nDetection",                  desc: "Real-time ML detection of unusual account behavior and transaction patterns across JH-powered institutions.", color: A3, icon: FaExclamationTriangle },
      { name: "FI Decision\nStudio",                 desc: "No-code AutoML for FIs: upload data, train custom models, explain with SHAP, author rules, deploy as API, monitor drift.", color: A6, icon: FaSlidersH },
      { name: "CommercialSignal",  desc: "ML classifier surfacing personal accounts with commercial transaction signatures — converting hidden SMB relationships to business banking.", color: A7, icon: FaBriefcase },
      { name: "Generational\nWealth Deflection",     desc: "Household coverage scoring to identify single-holder wealth concentration and expand FI relationships to next-generation family members.", color: A8, icon: FaUserFriends },
    ];

    // 3-column × 2-row grid (6 core apps)
    for (let i = 0; i < apps.length; i++) {
      const app = apps[i];
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.35 + col * 3.12;
      const y = 0.85 + row * 2.38;
      const icoData = await iconPng(app.icon, "#FFFFFF");

      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.96, h: 2.20, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.96, h: 0.72, fill: { color: app.color }, line: { color: app.color } });
      s.addImage({ data: icoData, x: x + 0.15, y: y + 0.14, w: 0.42, h: 0.42 });
      s.addText(app.name, {
        x: x + 0.65, y: y + 0.08, w: 2.21, h: 0.62,
        fontSize: 11.5, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
      });
      s.addText(app.desc, {
        x: x + 0.12, y: y + 0.80, w: 2.72, h: 1.28,
        fontSize: 10.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
      });
    }

    // Horizon callout strip
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 5.16, w: 9.3, h: 0.36, fill: { color: SKY_BG }, line: { color: COBALT } });
    s.addText("Horizon Pipeline includes Year 2 opportunities: Cross-Sell Propensity  ·  Fraud Ring Detection  ·  Deposit Runoff  ·  Overdraft Prediction  ·  and more →  see Horizon slide", {
      x: 0.52, y: 5.18, w: 9.0, h: 0.30,
      fontSize: 9.5, fontFace: FONT_B, color: COBALT, valign: "middle", margin: 0, italic: true
    });
  }

  // ==========================================================
  // SLIDES 5–9 — APP DETAIL SLIDES
  // ==========================================================
  const appDetails = [
    {
      name: "Churn Sentinel",
      icon: FaUserMinus,
      color: A1,
      problem: "Community banks and credit unions lose 8–12% of their customer base annually — but the signal is hiding in plain sight. Direct deposit cadence changes, ACH originator shifts, and balance migration patterns in payments data telegraph defection weeks before an account closes. By the time staff react, the relationship is already gone.",
      painPoints: [
        "Direct deposit moves are the #1 churn signal — caught only after the fact",
        "ACH cadence breaks (frequency, amount, originator) go unmonitored",
        "Relationship managers have no behavioral risk score to prioritize outreach",
        "Reactive retention is expensive — proactive saves cost a fraction",
      ],
      solution: "Churn Sentinel is a two-phase signal model built on JH's network data — Phase 1 launches on ACH payments signals (direct deposit monitoring, cadence anomalies, originator changes) and Phase 2 enriches with core banking and Banno digital signals as those pipelines come online. The model trains on JH's full FI network regardless of per-FI adoption.",
      approach: [
        "Phase 1 — ACH Sentinel: direct deposit cadence + originator change detection",
        "Phase 2 — Core + Digital: balance trends, product thinning, Banno session signals",
        "CatBoost survival model · SHAP explainability for banker-facing scores",
        "Weekly scores pushed to JH CRM → automated outreach workflow triggers",
      ],
      metrics: [
        { val: "$1–3M", label: "Annual revenue protected\nper FI ($1B deposits)" },
        { val: "15–25%", label: "Reduction in churn\nrate for FI clients" },
        { val: "Q3 FY26", label: "Phase 1 ACH Sentinel\ntarget launch" },
      ],
    },
    {
      name: "Zelle Memo Intelligence",
      icon: FaMobileAlt,
      color: A2,
      problem: "Zelle transaction volumes across JH-powered institutions have grown 40%+ year-over-year. Free-text memos contain fraud signals, elder abuse patterns, and compliance triggers — but manual review at scale is impossible and static rules miss adaptive language.",
      painPoints: [
        "Manual memo review is expensive and impossible to scale",
        "Pattern-based rules miss novel language and obfuscation",
        "No systemic link between memo text and SAR filings",
        "Compliance exposure growing with every Zelle transaction",
      ],
      solution: "An NLP classification and entity-extraction pipeline using fine-tuned transformer models — deployed within JH's compliance infrastructure to classify Zelle memos in near-real-time and surface risk-ranked alerts for examiner review.",
      approach: [
        "Fine-tuned DistilBERT / RoBERTa on labeled memo corpus",
        "Entity extraction: amounts, counterparties, intent signals",
        "Risk score → automated alert queue integration",
        "Attention visualization for compliance examiner review",
      ],
      metrics: [
        { val: "60%", label: "Reduction in manual\nmemo review hours per FI" },
        { val: "$250K+", label: "Annual compliance cost\nsavings per FI (est.)" },
        { val: "Q4 2026", label: "Target JH platform\ndelivery" },
      ],
    },
    {
      name: "Anomaly Detection",
      icon: FaExclamationTriangle,
      color: A3,
      problem: "Fraud and operational anomalies across JH-powered FIs are caught only by static rules — generating massive false positive rates that exhaust alert teams, while sophisticated multi-account patterns, account takeover signals, and operational errors slip through undetected.",
      painPoints: [
        "Static rules generate excessive false positive alert volume",
        "Novel fraud patterns evade threshold-based logic",
        "No unsupervised learning across multi-account behavior",
        "Operational anomalies (data gaps, system errors) undetected",
      ],
      solution: "An unsupervised anomaly detection platform using Isolation Forest, DBSCAN, and autoencoder neural networks — deployed as a JH platform service scoring transactions and behavioral signals across all connected FI institutions in near real-time.",
      approach: [
        "Isolation Forest + autoencoder ensemble architecture",
        "Streaming Kafka pipeline for near-real-time scoring",
        "Multi-entity context: account, customer, branch, channel",
        "SHAP anomaly scores surfaced in JH analyst dashboard",
      ],
      metrics: [
        { val: "$1–2M", label: "Annual fraud loss saved\nper FI (est.)" },
        { val: "70%", label: "False positive\nreduction target" },
        { val: "Q1 FY27", label: "Target JH platform\ndelivery" },
      ],
    },
    {
      name: "Call Report AI",
      icon: FaFileAlt,
      color: A4,
      problem: "Credit union executives lack real-time, actionable intelligence from NCUA 5300 call report data. Extracting peer benchmarks, spotting risk trends, and generating board-ready summaries requires manual data pulls, Excel gymnastics, and hours of analyst time every quarter.",
      painPoints: [
        "NCUA 5300 data is public but difficult to query at scale",
        "Peer benchmarking requires manual multi-institution comparison",
        "Board and exec briefs take hours to prepare from raw data",
        "No early warning system for emerging credit quality or capital risk",
      ],
      solution: "CallRpt AI — in active development, targeting Aug/Sept 2026 as a fast follow to JHBI approval — is a Claude-powered executive intelligence platform ingesting NCUA 5300 data for all 4,374 U.S. credit unions, with Payments and Core data modules to follow.",
      approach: [
        "Market Pulse: industry health score, KPIs, 8-quarter trend charts",
        "Ask: Claude-powered NL Q&A on full NCUA 5300 dataset",
        "Compare: percentile benchmarking + similar CU discovery",
        "AI Executive Brief: one-click Claude summary per institution",
      ],
      metrics: [
        { val: "Aug '26", label: "Target launch — fast\nfollow to JHBI approval" },
        { val: "4,374", label: "U.S. credit unions\nindexed in the platform" },
        { val: "$2.5T", label: "Total assets covered\nacross all institutions" },
      ],
    },
    {
      name: "Account Opening Lifetime Value",
      icon: FaChartLine,
      color: A5,
      problem: "FI clients using JH's account opening platform treat every new applicant identically — the same onboarding experience, the same cross-sell timing, regardless of predicted value. High-LTV customers are underinvested; resources are wasted on low-value accounts.",
      painPoints: [
        "No LTV signal available at the account opening stage",
        "Cross-sell sequencing is time-based, not value-informed",
        "Acquisition ROI unmeasured by customer cohort value",
        "Onboarding resources undifferentiated across channels",
      ],
      solution: "A CatBoost + AutoGluon LTV regression model scoring new accounts at opening — using behavioral archetypes, product signals, and demographic proxies to route applicants to differentiated service tiers and trigger ML-guided cross-sell sequencing.",
      approach: [
        "CatBoost regression + AutoGluon ensemble LTV scoring",
        "SHAP feature attribution for banker-facing explanations",
        "Integration with JH CRM for tiered onboarding routing",
        "Cohort tracking dashboard to validate model accuracy over time",
      ],
      metrics: [
        { val: "$800K+", label: "Incremental revenue\nper FI in Year 1 (est.)" },
        { val: "3–5×", label: "Cross-sell conversion\nimprovement" },
        { val: "Q1 FY27", label: "Target JH platform\ndelivery" },
      ],
    },
    {
      name: "CommercialSignal",
      icon: FaBriefcase,
      color: A7,
      problem: "Business owners run commercial operations through personal accounts daily — creating BSA/AML blind spots, Reg E gaps, and hiding SMB relationships that belong in business banking generating fee income, treasury products, and SBA loan revenue.",
      painPoints: [
        "Commercial cash flows in consumer accounts invisible to business banking teams",
        "BSA/AML programs miss commercial-risk patterns governed by Reg E frameworks",
        "FIs lose fee income, treasury products, and SBA loan opportunities from hidden SMBs",
        "Sole proprietors remain in consumer tiers with no differentiated service or depth",
      ],
      solution: "An ML classification model analyzing transaction velocity, ACH vendor/payroll signatures, payee diversity, and cash flow patterns to score personal accounts for commercial usage probability — surfacing them for business banking outreach and revenue conversion.",
      approach: [
        "CatBoost classifier: transaction velocity, ACH pattern, payee diversity, cash flow features",
        "SHAP feature attribution for banker review and Regulation E compliance documentation",
        "Integration with JH CRM for segmented SMB conversion outreach campaigns",
        "Automated monitoring for commercial activity thresholds with alert triggers",
      ],
      metrics: [
        { val: "$22T", label: "Annual business payments\nflowing through personal accounts" },
        { val: "30–40%", label: "Small business owners use\npersonal accounts commercially" },
        { val: "Q1 FY27", label: "Target JH platform\ndelivery" },
      ],
      sources: "Bank Director — 'The Untapped Market Hiding in Consumer Bank Accounts' (bankdirector.com/article/the-untapped-market-hiding-in-consumer-bank-accounts)  ·  Datos Insights — 'Bank Wealth Managers: Business Owner Client Opportunity' (datos-insights.com)",
    },
    {
      name: "Generational Wealth Deflection",
      icon: FaUserFriends,
      color: A8,
      problem: "FIs manage individual accounts but have no household model — so when wealth concentrates in a single-holder account with no family ties at the FI, a life event sends those assets to a competitor who already has the relationship with the heirs.",
      painPoints: [
        "FI relationships are account-level, not household- or family-level",
        "No model for identifying households with single-point-of-relationship risk",
        "Wealth transfer triggers reactive outreach — typically too late to retain assets",
        "Next-gen account holders are captured by digital-first competitors before FI engages",
      ],
      solution: "A household coverage scoring model that identifies existing accounts with concentrated wealth and no generational relationship depth — scoring deflection risk and triggering proactive family outreach campaigns to build multi-generational FI relationships before wealth transitions.",
      approach: [
        "Household linkage model: address, name, and account relationship graph from core data",
        "Wealth concentration + relationship density scoring per household",
        "Proactive outreach triggers: next-gen Banno digital invites, relationship banker introductions",
        "Household coverage rate tracked as FI performance KPI in JH executive dashboard",
      ],
      metrics: [
        { val: "$84T", label: "U.S. household wealth\ntransferring by 2045 (Cerulli)" },
        { val: "40%", label: "Of transferred wealth changes\nfinancial institutions" },
        { val: "Q1 FY27", label: "Target JH platform\ndelivery" },
      ],
      sources: "Cerulli Associates — 'U.S. High-Net-Worth and Ultra-High-Net-Worth Markets 2023'  ·  McKinsey — 'The Great Wealth Transfer' (2023)  ·  LIMRA — Intergenerational Wealth Transfer Research",
    },
  ];

  for (const app of appDetails) {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };
    const icoWhite = await iconPng(app.icon, "#FFFFFF");

    // Top header bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.76, fill: { color: app.color }, line: { color: app.color } });
    s.addImage({ data: icoWhite, x: 0.2, y: 0.16, w: 0.42, h: 0.42 });
    s.addText(app.name, {
      x: 0.72, y: 0.1, w: 7.2, h: 0.58,
      fontSize: 20, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });
    s.addText("JH Platform  |  AI Application", {
      x: 8.0, y: 0.18, w: 1.85, h: 0.4,
      fontSize: 10, fontFace: FONT_B, color: WHITE, align: "right", valign: "middle", margin: 0
    });

    // LEFT CARD — Problem
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 0.9, w: 4.58, h: 3.02, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 0.9, w: 0.08, h: 3.02, fill: { color: app.color }, line: { color: app.color } });

    s.addText("THE CHALLENGE FOR FI CLIENTS", {
      x: 0.48, y: 0.97, w: 4.24, h: 0.28,
      fontSize: 8.5, fontFace: FONT_H, bold: true, color: app.color, margin: 0, charSpacing: 2
    });
    s.addText(app.problem, {
      x: 0.48, y: 1.28, w: 4.22, h: 0.88,
      fontSize: 11, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });
    s.addText("Key Pain Points", {
      x: 0.48, y: 2.22, w: 4.0, h: 0.28,
      fontSize: 10, fontFace: FONT_H, bold: true, color: NAVY, margin: 0
    });
    const painBullets = app.painPoints.map((p, idx) => ({
      text: p,
      options: { bullet: true, breakLine: idx < app.painPoints.length - 1, paraSpaceAfter: 5 }
    }));
    s.addText(painBullets, {
      x: 0.48, y: 2.52, w: 4.22, h: 1.28,
      fontSize: 10.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // RIGHT CARD — Solution
    s.addShape(pres.shapes.RECTANGLE, { x: 5.14, y: 0.9, w: 4.58, h: 3.02, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.14, y: 0.9, w: 0.08, h: 3.02, fill: { color: COBALT }, line: { color: COBALT } });

    s.addText("WHAT JACK HENRY DELIVERS", {
      x: 5.34, y: 0.97, w: 4.24, h: 0.28,
      fontSize: 8.5, fontFace: FONT_H, bold: true, color: COBALT, margin: 0, charSpacing: 2
    });
    s.addText(app.solution, {
      x: 5.34, y: 1.28, w: 4.22, h: 0.88,
      fontSize: 11, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });
    s.addText("Technical Implementation", {
      x: 5.34, y: 2.22, w: 4.22, h: 0.28,
      fontSize: 10, fontFace: FONT_H, bold: true, color: NAVY, margin: 0
    });
    const approachBullets = app.approach.map((a, idx) => ({
      text: a,
      options: { bullet: true, breakLine: idx < app.approach.length - 1, paraSpaceAfter: 5 }
    }));
    s.addText(approachBullets, {
      x: 5.34, y: 2.52, w: 4.22, h: 1.28,
      fontSize: 10.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // BOTTOM METRIC CARDS (3) — shrink slightly when sources present to make room
    const metricY  = app.sources ? 4.00 : 4.09;
    const metricH  = app.sources ? 1.20 : 1.32;
    app.metrics.forEach((m, i) => {
      const x = 0.4 + i * 3.2;
      s.addShape(pres.shapes.RECTANGLE, { x, y: metricY, w: 2.8, h: metricH, fill: { color: NAVY }, line: { color: NAVY }, shadow: mkSh() });
      s.addText(m.val, {
        x: x + 0.08, y: metricY + 0.05, w: 2.64, h: 0.60,
        fontSize: 30, fontFace: FONT_H, bold: true, color: TECH, align: "center", margin: 0
      });
      s.addText(m.label, {
        x: x + 0.08, y: metricY + 0.66, w: 2.64, h: 0.48,
        fontSize: 10.5, fontFace: FONT_B, color: LT_GRAY, align: "center", margin: 0
      });
    });

    // SOURCE FOOTNOTE (optional) — thin strip at very bottom
    if (app.sources) {
      const srcY = metricY + metricH + 0.04;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: srcY, w: 9.44, h: 0.16, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: srcY, w: 0.06, h: 0.16, fill: { color: app.color }, line: { color: app.color } });
      s.addText(app.sources, {
        x: 0.42, y: srcY, w: 9.26, h: 0.16,
        fontSize: 7, fontFace: FONT_B, color: DK_GRAY, italic: true, valign: "middle", margin: 0
      });
    }
  }

  // ==========================================================
  // SLIDE — FI DECISION STUDIO (App 6)
  // ==========================================================
  {
    const appColor = A6;
    const s = pres.addSlide();
    s.background = { color: IVORY };

    // Top color bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.06, fill: { color: appColor }, line: { color: appColor } });
    const icoData = await iconPng(FaSlidersH, "#FFFFFF");
    s.addImage({ data: icoData, x: 0.38, y: 0.16, w: 0.68, h: 0.68 });
    s.addText("FI Decision Studio", {
      x: 1.18, y: 0.08, w: 7.5, h: 0.55,
      fontSize: 28, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });
    s.addText("AutoML · SHAP Explainability · Rule Authoring · API Hosting · Drift Monitoring", {
      x: 1.18, y: 0.63, w: 8.5, h: 0.32,
      fontSize: 10.5, fontFace: FONT_B, color: WHITE, margin: 0, italic: true
    });

    // Problem
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.18, w: 0.06, h: 0.28, fill: { color: appColor }, line: { color: appColor } });
    s.addText("THE PROBLEM", { x: 0.48, y: 1.18, w: 2.5, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: appColor, valign: "middle", margin: 0, charSpacing: 1 });
    s.addText("Community FIs have transaction data, behavioral signals, and fraud intuition built up over decades — but no way to turn it into models. Vendor AI is a black box that doesn't reflect local market dynamics and can't be explained to regulators under SR 11-7 model risk management guidelines. Building custom ML requires data science talent most FIs can't hire or afford.", {
      x: 0.35, y: 1.48, w: 4.55, h: 1.18,
      fontSize: 10.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Pain points
    const pains = [
      "No labeled data to train supervised models — FIs can't even start",
      "Vendor scores can't be explained to the board or examiners",
      "No rule authoring — compliance teams can't tune or override ML output",
      "Model deployment requires engineering resources FIs don't have",
    ];
    pains.forEach((pt, i) => {
      s.addShape(pres.shapes.OVAL, { x: 0.38, y: 2.76 + i * 0.36, w: 0.16, h: 0.16, fill: { color: appColor }, line: { color: appColor } });
      s.addText(pt, { x: 0.62, y: 2.70 + i * 0.36, w: 4.25, h: 0.34, fontSize: 10, fontFace: FONT_B, color: DK_GRAY, valign: "middle", margin: 0 });
    });

    // Divider
    s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.18, w: 0.04, h: 3.58, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });

    // Solution
    s.addShape(pres.shapes.RECTANGLE, { x: 5.28, y: 1.18, w: 0.06, h: 0.28, fill: { color: appColor }, line: { color: appColor } });
    s.addText("THE SOLUTION", { x: 5.42, y: 1.18, w: 2.5, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: appColor, valign: "middle", margin: 0, charSpacing: 1 });
    s.addText("FI Decision Studio meets every FI at their label maturity. Start with zero labels — the platform detects anomalies with unsupervised ML. As analysts review and tag cases, it tracks label maturity and offers a one-click upgrade to a full AutoGluon supervised ensemble. Tree SHAP explains every prediction, a rule authoring layer converts model insights into editable human decisions, and one-click API deployment puts models into production without engineering support.", {
      x: 5.28, y: 1.48, w: 4.35, h: 1.28,
      fontSize: 10.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Approach bullets
    s.addShape(pres.shapes.RECTANGLE, { x: 5.28, y: 2.84, w: 0.06, h: 0.28, fill: { color: appColor }, line: { color: appColor } });
    s.addText("PLATFORM PILLARS", { x: 5.42, y: 2.84, w: 3.5, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: appColor, valign: "middle", margin: 0, charSpacing: 1 });
    const pillars = [
      "Unsupervised → Supervised: Isolation Forest until labels mature; AutoGluon (CatBoost + XGBoost) ensemble when ready",
      "Tree SHAP on every prediction: top feature drivers, directional impact, plain-English output — SR 11-7 defensible",
      "Rule Authoring: model insights → editable decision rules with visual threshold sliders and backtest on historical data",
      "One-click API: versioned REST endpoint deployment; JH hosts, manages, and monitors — no FI engineering required",
      "Drift Monitoring: PSI on feature distributions weekly; prediction drift alerts; automated retraining recommendations",
    ];
    pillars.forEach((pt, i) => {
      const buls = [{ text: pt, options: { bullet: true } }];
      s.addText(buls, { x: 5.28, y: 3.18 + i * 0.38, w: 4.35, h: 0.36, fontSize: 9.5, fontFace: FONT_B, color: DK_GRAY, margin: 0 });
    });

    // Metric cards
    const metrics = [
      { val: "< 2 hrs", label: "Data upload → first\ntrained model deployed" },
      { val: "SR 11-7", label: "Model risk compliance\nout of the box with SHAP" },
      { val: "3 paths", label: "Unsupervised · supervised\n· hybrid label progression" },
    ];
    metrics.forEach((m, i) => {
      const x = 0.4 + i * 3.2;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 4.62, w: 2.8, h: 0.96, fill: { color: appColor }, line: { color: appColor } });
      s.addText(m.val, { x, y: 4.64, w: 2.8, h: 0.48, fontSize: 24, fontFace: FONT_H, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
      s.addText(m.label, { x, y: 5.10, w: 2.8, h: 0.46, fontSize: 9.5, fontFace: FONT_B, color: WHITE, align: "center", valign: "middle", margin: 0 });
    });
  }

  // ==========================================================
  // SLIDE 10 — FI ROI SUMMARY
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Financial Impact for FI Clients — Across the JH Platform", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    s.addText("Conservative estimates for a mid-size community bank ($1B deposits / $800M assets). Across 1,660 core FI clients — and 7,500+ across the broader JH ecosystem — these represent transformative market value.", {
      x: 0.4, y: 0.84, w: 9.2, h: 0.45,
      fontSize: 12, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    const roiRows = [
      {
        app: "Churn Sentinel",
        savings: "$1–3M / year",
        how: "15–25% churn reduction × avg. customer revenue of $400–$600/yr",
        color: A1
      },
      {
        app: "Zelle Memo Intelligence",
        savings: "$250K–$600K / year",
        how: "60% reduction in manual review hours + SAR accuracy / regulatory risk avoidance",
        color: A2
      },
      {
        app: "Anomaly Detection",
        savings: "$1–2M / year",
        how: "Fraud loss prevention + 70% false positive reduction (analyst labor saved)",
        color: A3
      },
      {
        app: "Call Report AI",
        savings: "$100–$300K / year",
        how: "Analyst hours eliminated for peer benchmarking, board prep, and quarterly NCUA 5300 review; Aug/Sept 2026 launch",
        color: A4
      },
      {
        app: "Account Opening LTV",
        savings: "$500K–$1.5M / year",
        how: "Tiered onboarding + 3–5× cross-sell improvement across 5,000–10,000 new accounts",
        color: A5
      },
      {
        app: "CommercialSignal",
        savings: "$400K–$1.2M / year",
        how: "SMB conversion revenue from business checking, treasury, and SBA loans on identified commercial-use accounts",
        color: A7
      },
      {
        app: "Generational Wealth Deflection",
        savings: "$500K–$2M+ / year",
        how: "AUM retained at generational transition × avg. household deposit balance; one retained high-net-worth household can exceed full-year subscription cost",
        color: A8
      },
    ];

    // Table header
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.36, w: 9.2, h: 0.38, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Application", { x: 0.5, y: 1.4, w: 2.8, h: 0.3, fontSize: 11, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    s.addText("Est. Annual FI Value", { x: 3.4, y: 1.4, w: 2.0, h: 0.3, fontSize: 11, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    s.addText("Value Driver", { x: 5.5, y: 1.4, w: 4.0, h: 0.3, fontSize: 11, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });

    roiRows.forEach((row, i) => {
      const y = 1.78 + i * 0.44;
      const rowBg = i % 2 === 0 ? WHITE : LT_GRAY;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.41, fill: { color: rowBg }, line: { color: LT_GRAY } });
      // App color chip
      s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.1, h: 0.41, fill: { color: row.color }, line: { color: row.color } });
      s.addText(row.app, { x: 0.55, y: y + 0.03, w: 2.75, h: 0.35, fontSize: 9.5, fontFace: FONT_B, bold: true, color: NAVY, valign: "middle", margin: 0 });
      s.addText(row.savings, { x: 3.4, y: y + 0.03, w: 1.95, h: 0.35, fontSize: 9.5, fontFace: FONT_H, bold: true, color: COBALT, valign: "middle", margin: 0 });
      s.addText(row.how, { x: 5.5, y: y + 0.03, w: 4.0, h: 0.35, fontSize: 8.5, fontFace: FONT_B, color: DK_GRAY, valign: "middle", margin: 0 });
    });

    // Total impact line — totalY = 1.78 + 7*0.44 = 4.86
    const totalY = 1.78 + 7 * 0.44;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: totalY, w: 9.2, h: 0.46, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("COMBINED POTENTIAL VALUE PER FI", { x: 0.55, y: totalY + 0.03, w: 2.75, h: 0.40, fontSize: 10, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    s.addText("$4–10M+ / year", { x: 3.4, y: totalY + 0.03, w: 1.95, h: 0.40, fontSize: 11, fontFace: FONT_H, bold: true, color: TECH, valign: "middle", margin: 0 });
    s.addText("Conservative aggregate across all 7 Phase 1–2 applications", { x: 5.5, y: totalY + 0.03, w: 4.0, h: 0.40, fontSize: 9.5, fontFace: FONT_B, color: LT_GRAY, valign: "middle", margin: 0 });
  }

  // ==========================================================
  // SLIDE — PRICING TIERS (3-tier SaaS model)
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("JH Intelligence Platform — Pricing Architecture", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });
    s.addText("Asset-tiered SaaS subscription — priced to fit every institution from $50M community CUs to $200B+ tier-1 FIs, with modules that expand as JH adds data sources.", {
      x: 0.4, y: 0.84, w: 9.2, h: 0.38,
      fontSize: 12, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Three tier cards
    const tiers = [
      {
        name: "Insight",
        tagline: "Entry point for emerging\nand community FIs",
        price: "$750 – $2,500",
        unit: "/ month per institution",
        assets: "Institutions < $500M assets",
        color: COBALT_M,
        features: [
          "CallRpt AI — NCUA 5300 intelligence",
          "Market Pulse dashboard (industry KPIs)",
          "Peer benchmarking (10 comparisons / mo)",
          "AI Q&A — 100 queries / month",
          "Standard executive brief templates",
          "Email support",
        ],
        highlight: false,
      },
      {
        name: "Intelligence",
        tagline: "Full DS platform for\ngrowth-stage FIs",
        price: "$2,500 – $8,000",
        unit: "/ month per institution",
        assets: "$500M – $5B assets",
        color: COBALT,
        features: [
          "Everything in Insight",
          "Churn / Attrition Mediation scores",
          "Account Opening Lifetime Value model",
          "Unlimited peer benchmarking + API",
          "AI Q&A — unlimited queries",
          "CRM integration + outreach triggers",
          "Dedicated onboarding support",
        ],
        highlight: true,
      },
      {
        name: "Enterprise",
        tagline: "Custom platform for\nlarge & complex FIs",
        price: "$10,000+",
        unit: "/ month — custom scoped",
        assets: "$5B+ assets",
        color: NAVY,
        features: [
          "Everything in Intelligence",
          "Zelle Memo Intelligence pipeline",
          "Anomaly Detection (real-time streaming)",
          "White-labeled client-facing portal",
          "Full REST API suite + webhooks",
          "Custom model training on FI data",
          "Dedicated success manager",
        ],
        highlight: false,
      },
    ];

    for (let i = 0; i < tiers.length; i++) {
      const t = tiers[i];
      const x = 0.3 + i * 3.18;
      const cardBg = t.highlight ? NAVY : WHITE;
      const textColor = t.highlight ? WHITE : DK_GRAY;
      const mutedColor = t.highlight ? LT_GRAY : MD_GRAY;
      const borderColor = t.highlight ? COBALT : LT_GRAY;

      // Card
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.32, w: 3.05, h: 4.15, fill: { color: cardBg }, line: { color: borderColor }, shadow: mkSh() });
      // Top color band
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.32, w: 3.05, h: 0.52, fill: { color: t.color }, line: { color: t.color } });

      // Tier name + "RECOMMENDED" badge on middle card
      s.addText(t.name.toUpperCase(), {
        x: x + 0.15, y: 1.37, w: t.highlight ? 1.8 : 2.75, h: 0.42,
        fontSize: 14, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0, charSpacing: 1
      });
      if (t.highlight) {
        s.addShape(pres.shapes.RECTANGLE, { x: x + 1.98, y: 1.42, w: 1.0, h: 0.3, fill: { color: TECH }, line: { color: TECH } });
        s.addText("RECOMMENDED", {
          x: x + 1.98, y: 1.42, w: 1.0, h: 0.3,
          fontSize: 7, fontFace: FONT_B, bold: true, color: NAVY, align: "center", valign: "middle", margin: 0, charSpacing: 0.5
        });
      }

      // Asset band
      s.addText(t.assets, {
        x: x + 0.12, y: 1.9, w: 2.82, h: 0.28,
        fontSize: 9.5, fontFace: FONT_B, color: t.color, margin: 0, bold: true
      });

      // Price
      s.addText(t.price, {
        x: x + 0.12, y: 2.2, w: 2.82, h: 0.55,
        fontSize: 26, fontFace: FONT_H, bold: true, color: t.highlight ? TECH : COBALT, margin: 0
      });
      s.addText(t.unit, {
        x: x + 0.12, y: 2.76, w: 2.82, h: 0.25,
        fontSize: 9.5, fontFace: FONT_B, color: mutedColor, margin: 0
      });

      // Divider
      s.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.12, y: 3.06, w: 2.82, h: 0.04,
        fill: { color: t.highlight ? COBALT_M : LT_GRAY }, line: { color: t.highlight ? COBALT_M : LT_GRAY }
      });

      // Features
      const featBullets = t.features.map((f, idx) => ({
        text: f,
        options: { bullet: true, breakLine: idx < t.features.length - 1, paraSpaceAfter: 4 }
      }));
      s.addText(featBullets, {
        x: x + 0.12, y: 3.14, w: 2.82, h: 2.22,
        fontSize: 9.5, fontFace: FONT_B, color: textColor, margin: 0
      });
    }

    // Bottom note
    s.addText("All tiers include: SOC 2 compliant data handling · NCUA 5300 + FFIEC base dataset · 99.9% SLA · Quarterly model updates", {
      x: 0.4, y: 5.46, w: 9.2, h: 0.22,
      fontSize: 8.5, fontFace: FONT_B, color: MD_GRAY, align: "center", margin: 0, italic: true
    });
  }

  // ==========================================================
  // SLIDE — PER-APP PRICING BREAKDOWN
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: IVORY };

    // Header bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Application Pricing — Full Platform Suite", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    // Subtitle
    s.addText("Each application can be licensed standalone or bundled. Bundle pricing reflects the Intelligence tier — 20–30% below standalone sum.", {
      x: 0.4, y: 0.82, w: 9.2, h: 0.28,
      fontSize: 10.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Column headers
    const colHdrs = [
      { label: "APPLICATION", x: 0.35, w: 3.4 },
      { label: "VALUE DRIVER", x: 3.85, w: 3.0 },
      { label: "STANDALONE / MO", x: 6.95, w: 1.55 },
      { label: "BUNDLE CREDIT", x: 8.55, w: 1.35 },
    ];
    colHdrs.forEach(h => {
      s.addText(h.label, { x: h.x, y: 1.18, w: h.w, h: 0.24, fontSize: 8.5, fontFace: FONT_H, bold: true, color: MD_GRAY, margin: 0, charSpacing: 0.5 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.43, w: 9.55, h: 0.02, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });

    const appRows = [
      {
        icon: FaChartLine,
        name: "Churn Sentinel",
        tagline: "ML churn scores + proactive retention triggers per member",
        driver: "15–25% reduction in member attrition · avg. $400–$600 LTV saved per retained account",
        price: "$800 – $2,500",
        credit: "Included in Intelligence+",
        color: COBALT,
      },
      {
        icon: FaMoneyBillWave,
        name: "Zelle Memo Intelligence",
        tagline: "NLP + compliance risk scoring on every Zelle transaction",
        driver: "60% reduction in manual SAR review hours · improved FinCEN defensibility",
        price: "$400 – $1,200",
        credit: "Included in Intelligence+",
        color: COBALT_D,
      },
      {
        icon: FaShieldAlt,
        name: "Anomaly Detection",
        tagline: "Real-time fraud + operational anomaly alerting across channels",
        driver: "70%+ false-positive reduction · analyst hours recaptured from alert triage",
        price: "$600 – $2,000",
        credit: "Included in Intelligence+",
        color: NAVY_MED,
      },
      {
        icon: FaFileAlt,
        name: "Call Report AI (CallRpt AI)",
        tagline: "NCUA 5300 intelligence — Pulse, Ask, Compare + AI briefs",
        driver: "Eliminates hours of peer benchmarking & board prep per quarter · Aug/Sept 2026",
        price: "$300 – $1,000",
        credit: "Included in all tiers",
        color: COBALT_M,
      },
      {
        icon: FaUsers,
        name: "Account Opening LTV Model",
        tagline: "Lifetime value scoring at account open · tiered onboarding triggers",
        driver: "3–5× cross-sell lift · 20–30% improvement in 90-day product adoption",
        price: "$500 – $1,500",
        credit: "Included in Intelligence+",
        color: A5,
      },
      {
        name: "FI Decision Studio",
        tagline: "No-code AutoML — unsupervised → supervised + rule authoring + API deploy",
        driver: "SR 11-7 defensible SHAP explainability · custom models on FI's own data · no ML team required",
        price: "$1,000 – $3,500",
        credit: "Enterprise tier only",
        color: A6,
      },
      {
        icon: FaBriefcase,
        name: "CommercialSignal",
        tagline: "ML scoring of personal accounts exhibiting commercial transaction signatures",
        driver: "Converts hidden SMB accounts to business banking — fee income, treasury products, SBA loans",
        price: "$400 – $1,200",
        credit: "Included in Intelligence+",
        color: A7,
      },
      {
        icon: FaUserFriends,
        name: "Generational Wealth Deflection",
        tagline: "Household coverage scoring · next-gen relationship expansion before wealth transfers",
        driver: "AUM retained at generational transition · multi-generational household relationships built proactively",
        price: "$500 – $1,500",
        credit: "Included in Intelligence+",
        color: A8,
      },
    ];

    appRows.forEach((row, i) => {
      const y = 1.42 + i * 0.48;
      const rowBg = i % 2 === 0 ? WHITE : LT_GRAY;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 9.55, h: 0.45, fill: { color: rowBg }, line: { color: LT_GRAY } });
      // Left accent
      s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.08, h: 0.45, fill: { color: row.color }, line: { color: row.color } });

      // App name + tagline
      s.addText(row.name, { x: 0.52, y: y + 0.03, w: 3.25, h: 0.19, fontSize: 9.5, fontFace: FONT_H, bold: true, color: NAVY, margin: 0 });
      s.addText(row.tagline, { x: 0.52, y: y + 0.24, w: 3.25, h: 0.18, fontSize: 7.5, fontFace: FONT_B, color: DK_GRAY, margin: 0 });

      // Value driver
      s.addText(row.driver, { x: 3.85, y: y + 0.04, w: 3.0, h: 0.37, fontSize: 7.5, fontFace: FONT_B, color: DK_GRAY, margin: 0, wrap: true });

      // Price
      s.addText(row.price, { x: 6.95, y: y + 0.03, w: 1.55, h: 0.24, fontSize: 10.5, fontFace: FONT_H, bold: true, color: COBALT, align: "center", valign: "middle", margin: 0 });
      s.addText("/ month per FI", { x: 6.95, y: y + 0.30, w: 1.55, h: 0.13, fontSize: 6.5, fontFace: FONT_B, color: MD_GRAY, align: "center", margin: 0 });

      // Bundle credit badge
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.58, y: y + 0.12, w: 1.28, h: 0.21, fill: { color: rowBg === WHITE ? SKY_BG : WHITE }, line: { color: COBALT }, rectRadius: 0.04 });
      s.addText(row.credit, { x: 8.58, y: y + 0.12, w: 1.28, h: 0.21, fontSize: 6.5, fontFace: FONT_B, color: COBALT, align: "center", valign: "middle", margin: 0 });
    });

    // Total standalone vs bundle bar — totY = 1.42 + 8*0.48 = 5.26
    const totY = 1.42 + 8 * 0.48;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: totY, w: 9.55, h: 0.32, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("STANDALONE TOTAL — ALL EIGHT APPS (MID-MARKET FI)", { x: 0.52, y: totY + 0.03, w: 4.5, h: 0.26, fontSize: 9, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    s.addText("~$8,100 / mo", { x: 6.95, y: totY + 0.03, w: 1.55, h: 0.26, fontSize: 12, fontFace: FONT_H, bold: true, color: TECH, align: "center", valign: "middle", margin: 0 });
    s.addText("vs. ~$6,000 bundled", { x: 8.35, y: totY + 0.04, w: 1.52, h: 0.22, fontSize: 8.5, fontFace: FONT_B, color: TECH, align: "center", valign: "middle", margin: 0 });
  }

  // ==========================================================
  // SLIDE — REVENUE EXPANSION MODEL (data source modules)
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Revenue Scales With Every Data Source We Add", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });
    s.addText("The JH Intelligence Platform is a modular expansion engine. Each new data source (Core, Payments, Digital) unlocks new applications and deepens revenue per FI client.", {
      x: 0.4, y: 0.84, w: 9.2, h: 0.38,
      fontSize: 12, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Module stack — 4 horizontal rows showing expansion
    const modules = [
      {
        phase: "NOW",
        module: "Regulatory Intelligence",
        sources: "NCUA 5300  ·  FFIEC Call Reports",
        apps: "CallRpt AI  ·  Compliance Monitoring",
        arr500: "$4.5M", arr2k: "$9M", arr8k: "$15M",
        color: COBALT,
        desc: "Public regulatory data — lowest friction entry point. Foundation of the platform.",
      },
      {
        phase: "NEXT",
        module: "+ Core Banking Intelligence",
        sources: "jhaEnterprise  ·  Symitar  ·  Core Transactions",
        apps: "Churn Sentinel  ·  Cross-Sell Propensity",
        arr500: "+$6M", arr2k: "+$12M", arr8k: "+$20M",
        color: COBALT_D,
        desc: "JH's proprietary core data creates a durable competitive moat — no competitor can replicate this.",
      },
      {
        phase: "NEXT",
        module: "+ Payments Intelligence",
        sources: "Zelle  ·  ACH  ·  Card  ·  Wire",
        apps: "Zelle Memo Intelligence  ·  Anomaly Detection  ·  Fraud Ring Detection",
        arr500: "+$5M", arr2k: "+$10M", arr8k: "+$17M",
        color: DK_GRAY,
        desc: "Payment behavioral signals unlock real-time risk detection at a depth rule-based systems cannot match.",
      },
      {
        phase: "FUTURE",
        module: "+ Digital / Banno Intelligence",
        sources: "Banno  ·  Digital Engagement  ·  Session Behavior",
        apps: "Sentiment Engine  ·  Next Best Action  ·  Overdraft Prediction",
        arr500: "+$4M", arr2k: "+$8M", arr8k: "+$13M",
        color: COBALT_M,
        desc: "Digital behavioral data closes the loop — every interaction becomes an AI-powered touchpoint.",
      },
    ];

    // Column headers
    s.addText("MODULE", { x: 0.35, y: 1.3, w: 4.9, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: MD_GRAY, margin: 0, charSpacing: 1 });
    s.addText("ARR @ 500 FIs", { x: 5.38, y: 1.3, w: 1.42, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: MD_GRAY, align: "center", margin: 0, charSpacing: 0.5 });
    s.addText("ARR @ 1,000 FIs", { x: 6.88, y: 1.3, w: 1.5, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: MD_GRAY, align: "center", margin: 0, charSpacing: 0.5 });
    s.addText("ARR @ 1,660 FIs", { x: 8.45, y: 1.3, w: 1.45, h: 0.28, fontSize: 9, fontFace: FONT_H, bold: true, color: MD_GRAY, align: "center", margin: 0, charSpacing: 0.5 });

    modules.forEach((m, i) => {
      const y = 1.60 + i * 0.84;
      const rowBg = i % 2 === 0 ? WHITE : LT_GRAY;

      s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 9.55, h: 0.78, fill: { color: rowBg }, line: { color: LT_GRAY } });
      // Left color accent
      s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.1, h: 0.78, fill: { color: m.color }, line: { color: m.color } });

      // Phase badge
      const badgeBg = m.phase === "NOW" ? COBALT : m.phase === "NEXT" ? COBALT_D : MD_GRAY;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.52, y: y + 0.12, w: 0.52, h: 0.20, fill: { color: badgeBg }, line: { color: badgeBg } });
      s.addText(m.phase, { x: 0.52, y: y + 0.12, w: 0.52, h: 0.20, fontSize: 7.5, fontFace: FONT_B, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });

      // Module name + sources + apps
      s.addText(m.module, { x: 1.12, y: y + 0.06, w: 4.1, h: 0.26, fontSize: 11, fontFace: FONT_H, bold: true, color: NAVY, margin: 0 });
      s.addText(m.sources, { x: 1.12, y: y + 0.31, w: 4.1, h: 0.20, fontSize: 8.5, fontFace: FONT_B, color: DK_GRAY, margin: 0 });
      s.addText(m.apps, { x: 1.12, y: y + 0.50, w: 4.1, h: 0.20, fontSize: 8.5, fontFace: FONT_B, color: m.color, margin: 0, italic: true });

      // ARR columns
      const arrVals = [m.arr500, m.arr2k, m.arr8k];
      const arrX = [5.38, 6.88, 8.45];
      const arrW = [1.42, 1.5, 1.45];
      arrVals.forEach((v, j) => {
        const isPositive = v.startsWith("+");
        const numColor = isPositive ? COBALT_D : COBALT;
        s.addText(v, {
          x: arrX[j], y: y + 0.18, w: arrW[j], h: 0.38,
          fontSize: 17, fontFace: FONT_H, bold: true, color: numColor, align: "center", margin: 0
        });
      });
    });

    // Total row
    const totY = 1.60 + 4 * 0.84;  // = 4.96
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: totY, w: 9.55, h: 0.44, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("FULL PLATFORM ARR POTENTIAL", { x: 0.52, y: totY + 0.03, w: 4.7, h: 0.38, fontSize: 11, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    [{ v: "$19.5M", x: 5.38, w: 1.42 }, { v: "$39M", x: 6.88, w: 1.5 }, { v: "$65M+", x: 8.45, w: 1.45 }].forEach(col => {
      s.addText(col.v, { x: col.x, y: totY + 0.03, w: col.w, h: 0.38, fontSize: 14, fontFace: FONT_H, bold: true, color: TECH, align: "center", valign: "middle", margin: 0 });
    });

    // Footnote
    s.addText("ARR estimates based on avg. subscription value across tiers. Columns reflect 30%, 60%, and 100% penetration of JH's 1,660 core FI clients (940 banks + 720 CUs). Source: JH 10-K FY2024.", {
      x: 0.4, y: 5.44, w: 9.2, h: 0.18,
      fontSize: 8, fontFace: FONT_B, color: MD_GRAY, italic: true, margin: 0
    });
  }

  // ==========================================================
  // SLIDE 11 — HORIZON PIPELINE
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Horizon Pipeline — Phase 2 & Year 2 Platform Opportunities", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 20, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    const horizonApps = [
      // Year 2 horizon (6 apps, 3×2 grid)
      { name: "Cross-Sell Propensity",       desc: "ML model scoring each FI customer's likelihood to open savings, CDs, credit cards, or mortgage — enabling timed, personalized offers embedded in Banno.", icon: FaUsers,        color: COBALT,   badge: "Year 2" },
      { name: "Fraud Ring Detection",        desc: "Graph-based ML to identify coordinated fraud across linked accounts, shared devices, and IP clusters — across the entire JH network for cross-institution signals.", icon: FaShieldAlt,   color: NAVY,     badge: "Year 2" },
      { name: "Deposit Runoff Forecasting",  desc: "Portfolio-level deposit attrition prediction to support FI clients' ALM, liquidity stress testing, and treasury planning with 90-day forward visibility.", icon: FaMoneyBillWave,color: COBALT_D, badge: "Year 2" },
      { name: "Overdraft Prediction",        desc: "Anticipate overdraft events before they occur and surface proactive micro-loan offers or alerts in Banno — improving CX while reducing NSF fee risk for FIs.", icon: FaBalanceScale, color: DK_GRAY,  badge: "Year 2" },
      { name: "Customer Sentiment Engine",   desc: "NLP analysis of call transcripts and CFPB complaints to generate real-time CX health scores — helping FIs identify and route at-risk relationships before escalation.", icon: FaHeadset,     color: COBALT_M, badge: "Year 2" },
      { name: "Next Best Action Platform",   desc: "Unified orchestration engine synthesizing all model outputs — churn score, LTV, propensity, sentiment — into a single ranked recommendation per customer per channel.", icon: FaRocket,      color: NAVY_MED, badge: "Year 2" },
    ];

    // 3-column × 2-row grid (6 items)
    for (let i = 0; i < horizonApps.length; i++) {
      const app = horizonApps[i];
      const col = i % 3;
      const row = Math.floor(i / 3);
      const x = 0.14 + col * 3.22;
      const y = 0.86 + row * 2.30;
      const icoData = await iconPng(app.icon, `#${app.color}`);

      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.06, h: 2.12, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.06, h: 0.08, fill: { color: app.color }, line: { color: app.color } });

      // Phase/timing badge
      s.addShape(pres.shapes.RECTANGLE, { x: x + 2.40, y: y + 0.14, w: 0.60, h: 0.20, fill: { color: app.color }, line: { color: app.color } });
      s.addText(app.badge, { x: x + 2.40, y: y + 0.14, w: 0.60, h: 0.20, fontSize: 6.5, fontFace: FONT_B, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });

      s.addImage({ data: icoData, x: x + 0.14, y: y + 0.20, w: 0.34, h: 0.34 });
      s.addText(app.name, {
        x: x + 0.54, y: y + 0.14, w: 1.80, h: 0.44,
        fontSize: 9.5, fontFace: FONT_H, bold: true, color: NAVY, valign: "middle", margin: 0
      });
      s.addText(app.desc, {
        x: x + 0.12, y: y + 0.66, w: 2.82, h: 1.34,
        fontSize: 9, fontFace: FONT_B, color: DK_GRAY, margin: 0
      });
    }

    s.addText("Year 2 opportunities contingent on Phase 1 & Phase 2 delivery and continued investment approval.", {
      x: 0.4, y: 5.22, w: 9.2, h: 0.28,
      fontSize: 9, fontFace: FONT_B, color: MD_GRAY, italic: true, margin: 0
    });
  }

  // ==========================================================
  // SLIDE 12 — FOUR CAPABILITY PILLARS
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Four Investment Pillars — What We Are Building", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    const pillars = [
      {
        title: "Data Acquisition",
        tag: "THE FOUNDATION",
        color: COBALT,
        icon: FaDatabase,
        bullets: [
          "FI behavioral data pipelines from JH core systems",
          "Real-time streaming infrastructure (Kafka / event bus)",
          "Data quality monitoring & lineage tracking",
          "Enterprise data catalog and governance framework",
          "Aggregated cross-FI benchmarking data (anonymized)",
        ],
      },
      {
        title: "DS Buildout",
        tag: "THE ENGINE",
        color: COBALT_D,
        icon: FaBrain,
        bullets: [
          "ML platform: feature store + experiment tracking",
          "AutoGluon, CatBoost, XGBoost model training at scale",
          "SHAP explainability layer across all deployed models",
          "MLOps: CI/CD, model registry, drift monitoring",
          "Model governance, bias review, and model card process",
        ],
      },
      {
        title: "Visualizations",
        tag: "THE LENS",
        color: NAVY_MED,
        icon: FaChartBar,
        bullets: [
          "Executive KPI dashboards embedded in JH platform",
          "Model performance and drift monitoring views",
          "FI client-facing risk score explorer (banker tool)",
          "Anomaly alert investigation and triage interface",
          "Self-serve analytics surface for business stakeholders",
        ],
      },
      {
        title: "App / API Development",
        tag: "THE INTERFACE",
        color: DK_GRAY,
        icon: FaCode,
        bullets: [
          "Lightweight internal AI apps embedded in JH ecosystem",
          "REST APIs distributing model scores to JH products",
          "Banno + JH Enterprise integration for score consumption",
          "Alerting and workflow orchestration service layer",
          "Role-based access control and full audit logging",
        ],
      },
    ];

    for (let i = 0; i < pillars.length; i++) {
      const p = pillars[i];
      const x = 0.28 + i * 2.38;
      const icoData = await iconPng(p.icon, "#FFFFFF");

      s.addShape(pres.shapes.RECTANGLE, { x, y: 0.82, w: 2.24, h: 4.62, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 0.82, w: 2.24, h: 0.52, fill: { color: p.color }, line: { color: p.color } });
      s.addImage({ data: icoData, x: x + 0.9, y: 0.87, w: 0.42, h: 0.42 });
      s.addText(p.tag, {
        x: x + 0.06, y: 1.4, w: 2.12, h: 0.28,
        fontSize: 7.5, fontFace: FONT_B, bold: true, color: p.color, align: "center", margin: 0, charSpacing: 1.2
      });
      s.addText(p.title, {
        x: x + 0.06, y: 1.7, w: 2.12, h: 0.42,
        fontSize: 13.5, fontFace: FONT_H, bold: true, color: NAVY, align: "center", margin: 0
      });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.55, y: 2.16, w: 1.12, h: 0.04, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });
      const bulletItems = p.bullets.map((b, idx) => ({
        text: b,
        options: { bullet: true, breakLine: idx < p.bullets.length - 1, paraSpaceAfter: 6 }
      }));
      s.addText(bulletItems, {
        x: x + 0.1, y: 2.25, w: 2.08, h: 3.1,
        fontSize: 9.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
      });
    }
  }

  // ==========================================================
  // SLIDE 13 — RESOURCE REQUEST
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("What We're Asking For: JHBI Investment", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    s.addText("To execute this roadmap and embed AI capabilities across the JH platform, we are requesting the following headcount and infrastructure investment:", {
      x: 0.4, y: 0.84, w: 9.2, h: 0.38,
      fontSize: 12.5, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    const tableRows = [
      [
        { text: "Role",         options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11.5, fontFace: FONT_H } },
        { text: "FTEs",         options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11.5, fontFace: FONT_H, align: "center" } },
        { text: "Pillar",       options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11.5, fontFace: FONT_H } },
        { text: "Rationale",    options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11.5, fontFace: FONT_H } },
      ],
      [
        { text: "Data Engineer",               options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "2",                           options: { fontFace: FONT_B, fontSize: 11, align: "center" } },
        { text: "Data Acquisition",            options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "Build and maintain FI data pipelines from JH core banking systems",          options: { fontFace: FONT_B, fontSize: 11 } },
      ],
      [
        { text: "ML Engineer",                 options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "2",                           options: { fontFace: FONT_B, fontSize: 11, align: "center" } },
        { text: "DS Buildout",                 options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "Model training, MLOps pipeline, feature store, drift monitoring",     options: { fontFace: FONT_B, fontSize: 11 } },
      ],
      [
        { text: "Applied Data Scientist",      options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "3",                           options: { fontFace: FONT_B, fontSize: 11, align: "center" } },
        { text: "DS Buildout",                 options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "Feature engineering, SHAP explainability, model research and validation",    options: { fontFace: FONT_B, fontSize: 11 } },
      ],
      [
        { text: "Full-Stack / API Developer",  options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "2",                           options: { fontFace: FONT_B, fontSize: 11, align: "center" } },
        { text: "App / API Dev",               options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "JH platform AI app development and model API integration layer",             options: { fontFace: FONT_B, fontSize: 11 } },
      ],
      [
        { text: "BI / Viz Developer",          options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "1",                           options: { fontFace: FONT_B, fontSize: 11, align: "center" } },
        { text: "Visualizations",              options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "Embedded analytics surfaces in JH platform and client-facing dashboards",    options: { fontFace: FONT_B, fontSize: 11 } },
      ],
      [
        { text: "DS Product Manager",          options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "1",                           options: { fontFace: FONT_B, fontSize: 11, align: "center" } },
        { text: "All Pillars",                 options: { fontFace: FONT_B, fontSize: 11 } },
        { text: "Roadmap ownership, FI client discovery, stakeholder alignment",              options: { fontFace: FONT_B, fontSize: 11 } },
      ],
      [
        { text: "TOTAL",                       options: { bold: true, fontFace: FONT_H, fontSize: 11.5, fill: { color: LT_GRAY } } },
        { text: "11",                          options: { bold: true, fontFace: FONT_H, fontSize: 12, color: NAVY, align: "center", fill: { color: LT_GRAY } } },
        { text: "",                            options: { fill: { color: LT_GRAY } } },
        { text: "Full-time equivalents across all four investment pillars",                    options: { italic: true, fontFace: FONT_B, fontSize: 10.5, color: MD_GRAY, fill: { color: LT_GRAY } } },
      ],
    ];

    s.addTable(tableRows, {
      x: 0.4, y: 1.28, w: 9.2,
      fontSize: 11, fontFace: FONT_B, color: DK_GRAY,
      border: { pt: 0.5, color: LT_GRAY },
      rowH: 0.41,
      colW: [2.35, 0.92, 1.78, 4.15],
    });

    // Infrastructure note
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.05, w: 9.2, h: 0.44, fill: { color: SKY_BG }, line: { color: COBALT } });
    s.addText("Infrastructure ask (separate from headcount): Cloud data platform (Snowflake / Databricks), ML tooling (MLflow / SageMaker), BI licensing (Tableau / Power BI), DevOps / CI-CD environment.", {
      x: 0.55, y: 5.08, w: 8.9, h: 0.38,
      fontSize: 10.5, fontFace: FONT_B, color: NAVY, valign: "middle", margin: 0
    });
  }

  // ==========================================================
  // SLIDE — PAYMENT DATA SOURCES GANTT (minus CPS)
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: IVORY };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Data Source Ingestion Roadmap — Payments First, Core & Digital Next", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 20, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    s.addText("Phase 1 activates all JH payment rails (CPS excluded). Core banking, Banno digital, and lending data follow in Phase 2–3 once the pipeline is proven.", {
      x: 0.4, y: 0.78, w: 9.2, h: 0.28,
      fontSize: 10, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // Layout constants
    const labelW = 2.25;
    const timelineX = labelW + 0.45;   // x where timeline starts
    const timelineW = 9.6 - timelineX; // ~6.9"
    const nQ = 8;                       // Q3 FY26 → Q2 FY28
    const qW = timelineW / nQ;
    const quarters = ["Q3 FY26","Q4 FY26","Q1 FY27","Q2 FY27","Q3 FY27","Q4 FY27","Q1 FY28","Q2 FY28"];
    const rowH = 0.42;
    const rowsStartY = 1.82;           // pushed down so phase bands clear the subtitle

    // Phase bands — sit between subtitle and quarter headers
    const phaseBands = [
      { label: "PHASE 1 — All Payment Rails  (Q3–Q4 FY26)", startQ: 0, endQ: 1, color: COBALT },
      { label: "PHASE 2 — Core Banking  (Q1–Q2 FY27)", startQ: 2, endQ: 3, color: COBALT_D },
      { label: "PHASE 3 — Digital & Lending  (Q3 FY27+)", startQ: 4, endQ: 7, color: DK_GRAY },
    ];
    phaseBands.forEach(pb => {
      const bx = timelineX + pb.startQ * qW;
      const bw = (pb.endQ - pb.startQ + 1) * qW;
      s.addShape(pres.shapes.RECTANGLE, { x: bx, y: rowsStartY - 0.58, w: bw, h: 0.26, fill: { color: pb.color }, line: { color: pb.color } });
      s.addText(pb.label, { x: bx + 0.05, y: rowsStartY - 0.58, w: bw - 0.1, h: 0.26, fontSize: 7.5, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    });

    // Quarter header background
    s.addShape(pres.shapes.RECTANGLE, { x: timelineX, y: rowsStartY - 0.30, w: timelineW, h: 0.28, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });
    quarters.forEach((q, qi) => {
      const qx = timelineX + qi * qW;
      // Alternating column shade
      if (qi % 2 === 0) {
        s.addShape(pres.shapes.RECTANGLE, { x: qx, y: rowsStartY, w: qW, h: 9 * rowH, fill: { color: "F4F6F8" }, line: { color: "F4F6F8" } });
      }
      s.addText(q, { x: qx, y: rowsStartY - 0.30, w: qW, h: 0.26, fontSize: 7.5, fontFace: FONT_H, bold: true, color: DK_GRAY, align: "center", valign: "middle", margin: 0 });
      // Vertical gridline
      s.addShape(pres.shapes.RECTANGLE, { x: qx, y: rowsStartY - 0.58, w: 0.015, h: 9 * rowH + 0.58, fill: { color: MD_GRAY }, line: { color: MD_GRAY } });
    });

    // Data source rows
    // startQ/endQ are 0-indexed (0 = Q3 FY26, 7 = Q2 FY28)
    // PHASE 1 = all JH payment rails (Q3–Q4 FY26)
    // PHASE 2/3 = Core banking, Banno digital, Lending
    const sources = [
      { name: "ACH / EPS",                     detail: "ACH credits, debits, RDC",        startQ: 0, endQ: 1, color: COBALT,   apps: "Anomaly Detection" },
      { name: "Zelle / JHA PayCenter",         detail: "P2P, Zelle, RTP, FedNow",         startQ: 0, endQ: 1, color: COBALT,   apps: "Zelle Memo Intelligence · Anomaly" },
      { name: "Jack Henry Wires",              detail: "Domestic FedWire + Intl",         startQ: 0, endQ: 1, color: COBALT,   apps: "Anomaly Detection (wire fraud)" },
      { name: "iPay Bill Pay",                 detail: "Consumer + Business bill pay",    startQ: 0, endQ: 1, color: COBALT,   apps: "Cross-Sell Propensity" },
      { name: "JHA SmartPay / Biz Payments",   detail: "B2B ACH, vendor, remittance",     startQ: 0, endQ: 1, color: COBALT,   apps: "Overdraft Prediction" },
      { name: "Payrailz (P2P / A2A)",         detail: "Pay a Person, Transfer Money",    startQ: 0, endQ: 1, color: COBALT,   apps: "Fraud Ring Detection" },
      { name: "Core Banking (jhaEnterprise / Symitar)", detail: "Core txns, deposits, loans", startQ: 2, endQ: 3, color: COBALT_D, apps: "Churn Sentinel · Acct Opening LTV · CommercialSignal · Wealth Deflection" },
      { name: "Banno Digital + Lending",       detail: "Session behavior, loan data",     startQ: 4, endQ: 6, color: DK_GRAY,  apps: "Sentiment · Next Best Action · FI Decision Studio" },
    ];

    sources.forEach((src, ri) => {
      const y = rowsStartY + ri * rowH;
      const rowBg = ri % 2 === 0 ? WHITE : "F8F9FA";
      s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.3, h: rowH - 0.03, fill: { color: rowBg }, line: { color: LT_GRAY } });

      // Source name label
      s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 0.06, h: rowH - 0.03, fill: { color: src.color }, line: { color: src.color } });
      s.addText(src.name, { x: 0.42, y: y + 0.03, w: labelW - 0.18, h: 0.22, fontSize: 8.5, fontFace: FONT_H, bold: true, color: NAVY, margin: 0 });
      s.addText(src.detail, { x: 0.42, y: y + 0.24, w: labelW - 0.18, h: 0.15, fontSize: 7.5, fontFace: FONT_B, color: MD_GRAY, margin: 0, italic: true });

      // Gantt bar
      const barX = timelineX + src.startQ * qW + 0.04;
      const barW = (src.endQ - src.startQ + 1) * qW - 0.08;
      s.addShape(pres.shapes.RECTANGLE, { x: barX, y: y + 0.08, w: barW, h: rowH - 0.22, fill: { color: src.color }, line: { color: src.color } });

      // App label: inside bar if wide enough, otherwise to the right
      if (barW >= 1.4) {
        s.addText(src.apps, { x: barX + 0.07, y: y + 0.08, w: barW - 0.12, h: rowH - 0.22, fontSize: 7.5, fontFace: FONT_B, color: WHITE, valign: "middle", margin: 0 });
      } else {
        const appX = barX + barW + 0.06;
        const appAvailW = 9.55 - appX;
        if (appAvailW > 0.4) {
          s.addText(src.apps, { x: appX, y: y + 0.08, w: Math.min(appAvailW, 2.4), h: rowH - 0.22, fontSize: 7.5, fontFace: FONT_B, color: src.color, valign: "middle", margin: 0, italic: true });
        }
      }
    });

    // Today marker (Q1 2026 start)
    s.addShape(pres.shapes.RECTANGLE, { x: timelineX + 0.015, y: rowsStartY - 0.60, w: 0.04, h: 9 * rowH + 0.60, fill: { color: TECH }, line: { color: TECH } });
    s.addText("NOW", { x: timelineX - 0.12, y: rowsStartY - 0.62, w: 0.55, h: 0.24, fontSize: 7, fontFace: FONT_H, bold: true, color: TECH, align: "center", margin: 0 });
  }

  // ==========================================================
  // SLIDE — APP DELIVERY GANTT
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: IVORY };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Application Delivery Roadmap — Build & Launch Timeline", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 20, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });
    s.addText("Eight applications sequenced across three phases. Each bar spans the active build and validation window; wider bars reflect higher integration complexity.", {
      x: 0.4, y: 0.78, w: 9.2, h: 0.28,
      fontSize: 10, fontFace: FONT_B, color: DK_GRAY, margin: 0
    });

    // ── Layout constants (same spine as Data Source Gantt) ──
    const aLabelW    = 2.30;
    const aTimelineX = aLabelW + 0.44;
    const aTimelineW = 9.6 - aTimelineX;
    const aNQ        = 8;
    const aQW        = aTimelineW / aNQ;
    const aQtrs      = ["Q3 FY26","Q4 FY26","Q1 FY27","Q2 FY27","Q3 FY27","Q4 FY27","Q1 FY28","Q2 FY28"];
    const aRowH      = 0.44;
    const aStartY    = 1.82;

    // ── Phase bands ──
    const aPhaseBands = [
      { label: "PHASE 1 — Foundation  (Q3–Q4 FY26)", startQ: 0, endQ: 1, color: COBALT },
      { label: "PHASE 2 — Core Applications  (Q1–Q2 FY27)", startQ: 2, endQ: 3, color: COBALT_D },
      { label: "PHASE 3 — Scale & Expand  (Q3 FY27+)", startQ: 4, endQ: 7, color: DK_GRAY },
    ];
    aPhaseBands.forEach(pb => {
      const bx = aTimelineX + pb.startQ * aQW;
      const bw = (pb.endQ - pb.startQ + 1) * aQW;
      s.addShape(pres.shapes.RECTANGLE, { x: bx, y: aStartY - 0.58, w: bw, h: 0.26, fill: { color: pb.color }, line: { color: pb.color } });
      s.addText(pb.label, { x: bx + 0.05, y: aStartY - 0.58, w: bw - 0.1, h: 0.26, fontSize: 7.5, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0 });
    });

    // ── Quarter headers & column shading ──
    s.addShape(pres.shapes.RECTANGLE, { x: aTimelineX, y: aStartY - 0.30, w: aTimelineW, h: 0.28, fill: { color: LT_GRAY }, line: { color: LT_GRAY } });
    aQtrs.forEach((q, qi) => {
      const qx = aTimelineX + qi * aQW;
      if (qi % 2 === 0) {
        s.addShape(pres.shapes.RECTANGLE, { x: qx, y: aStartY, w: aQW, h: 8 * aRowH, fill: { color: "F4F6F8" }, line: { color: "F4F6F8" } });
      }
      s.addText(q, { x: qx, y: aStartY - 0.30, w: aQW, h: 0.26, fontSize: 7.5, fontFace: FONT_H, bold: true, color: DK_GRAY, align: "center", valign: "middle", margin: 0 });
      s.addShape(pres.shapes.RECTANGLE, { x: qx, y: aStartY - 0.58, w: 0.015, h: 8 * aRowH + 0.58, fill: { color: MD_GRAY }, line: { color: MD_GRAY } });
    });

    // ── App rows ──
    // startQ/endQ = 0-indexed quarters (0 = Q3 FY26 … 7 = Q2 FY28)
    // Bar spans active build → validate → launch window
    const appGanttRows = [
      { name: "Churn Sentinel",          detail: "ACH signals  ·  Phase 1",  startQ: 0, endQ: 1, color: A1,  label: "ACH Sentinel  ·  Phase 1 launch" },
      { name: "Churn Sentinel +",        detail: "Core + digital enrichment", startQ: 2, endQ: 3, color: A1,  label: "Core & Banno signals  →  Phase 2" },
      { name: "Zelle Memo Intelligence",  detail: "NLP compliance pipeline",  startQ: 2, endQ: 3, color: A2,  label: "Build  →  Phase 2 launch" },
      { name: "CommercialSignal",   detail: "SMB conversion model",     startQ: 2, endQ: 3, color: A7,  label: "Build  →  Phase 2 launch" },
      { name: "Gen. Wealth Deflection",   detail: "Household coverage model", startQ: 2, endQ: 4, color: A8,  label: "Phase 2 build  →  Phase 3" },
      { name: "Anomaly Detection",        detail: "Real-time fraud engine",   startQ: 4, endQ: 6, color: A3,  label: "Build  →  Phase 3 launch" },
      { name: "FI Decision Studio",       detail: "AutoML platform  ·  GA",   startQ: 4, endQ: 7, color: A6,  label: "Platform build  →  Phase 3 GA" },
    ];

    appGanttRows.forEach((row, ri) => {
      const y      = aStartY + ri * aRowH;
      const rowBg  = ri % 2 === 0 ? WHITE : "F8F9FA";

      // Row background
      s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.3, h: aRowH - 0.03, fill: { color: rowBg }, line: { color: LT_GRAY } });

      // Left color accent + label
      s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 0.06, h: aRowH - 0.03, fill: { color: row.color }, line: { color: row.color } });
      s.addText(row.name,   { x: 0.42, y: y + 0.03, w: aLabelW - 0.18, h: 0.22, fontSize: 8.5, fontFace: FONT_H, bold: true, color: NAVY,    margin: 0 });
      s.addText(row.detail, { x: 0.42, y: y + 0.25, w: aLabelW - 0.18, h: 0.15, fontSize: 7.5, fontFace: FONT_B, color: MD_GRAY, margin: 0, italic: true });

      // Gantt bar (with a slightly lighter "build" tail)
      const barX = aTimelineX + row.startQ * aQW + 0.04;
      const barW = (row.endQ - row.startQ + 1) * aQW - 0.08;
      const barH = aRowH - 0.24;
      const barY = y + 0.10;

      // Draw: lighter "build" region (full bar), then darker "launch" cap on final quarter
      s.addShape(pres.shapes.RECTANGLE, { x: barX,                   y: barY, w: barW,       h: barH, fill: { color: row.color, transparency: 30 }, line: { color: row.color, transparency: 10 } });
      const launchCapW = Math.min(aQW - 0.08, barW * 0.28);
      s.addShape(pres.shapes.RECTANGLE, { x: barX + barW - launchCapW, y: barY, w: launchCapW, h: barH, fill: { color: row.color }, line: { color: row.color } });

      // Label — inside if wide, outside if narrow
      if (barW >= 1.4) {
        s.addText(row.label, { x: barX + 0.07, y: barY, w: barW - 0.14, h: barH, fontSize: 7.5, fontFace: FONT_B, color: WHITE, valign: "middle", margin: 0 });
      } else {
        const lblX = barX + barW + 0.06;
        const lblW = 9.55 - lblX;
        if (lblW > 0.5) {
          s.addText(row.label, { x: lblX, y: barY, w: Math.min(lblW, 2.4), h: barH, fontSize: 7.5, fontFace: FONT_B, color: row.color, valign: "middle", margin: 0, italic: true });
        }
      }
    });

    // ── Legend: build vs. launch shading ──
    const legX = 0.42;
    const legY = aStartY + 8 * aRowH + 0.04;
    s.addShape(pres.shapes.RECTANGLE, { x: legX,       y: legY, w: 0.50, h: 0.14, fill: { color: COBALT, transparency: 30 }, line: { color: COBALT } });
    s.addText("Build / validate",  { x: legX + 0.56, y: legY, w: 1.4, h: 0.14, fontSize: 7.5, fontFace: FONT_B, color: DK_GRAY, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: legX + 2.1, y: legY, w: 0.50, h: 0.14, fill: { color: COBALT }, line: { color: COBALT } });
    s.addText("Launch window",     { x: legX + 2.66, y: legY, w: 1.4, h: 0.14, fontSize: 7.5, fontFace: FONT_B, color: DK_GRAY, valign: "middle", margin: 0 });

    // ── Today marker ──
    s.addShape(pres.shapes.RECTANGLE, { x: aTimelineX + 0.015, y: aStartY - 0.60, w: 0.04, h: 8 * aRowH + 0.60, fill: { color: TECH }, line: { color: TECH } });
    s.addText("NOW", { x: aTimelineX - 0.12, y: aStartY - 0.62, w: 0.55, h: 0.24, fontSize: 7, fontFace: FONT_H, bold: true, color: TECH, align: "center", margin: 0 });
  }

  // ==========================================================
  // SLIDE 14 — PHASED ROADMAP
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: SKY_BG };

    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.72, fill: { color: NAVY }, line: { color: NAVY } });
    s.addText("Phased Delivery Roadmap — FY26 Q3 through FY28 Q2", {
      x: 0.4, y: 0.08, w: 9.2, h: 0.56,
      fontSize: 22, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
    });

    const phases = [
      {
        num: "01", title: "Foundation", period: "Q3–Q4 FY26  (Jan–Jun 2026)", color: COBALT,
        items: [
          "Hire core team: Data Engineers + ML Engineers",
          "Stand up cloud data platform and ML infrastructure",
          "Ingest ACH payment rails — direct deposit monitoring live",
          "Churn Sentinel Phase 1: ACH cadence + originator change alerts",
        ],
        apps: [
          { label: "Churn Sentinel", color: A1, badge: "ACH" },
        ],
      },
      {
        num: "02", title: "Core Applications", period: "Q1–Q2 FY27  (Jul–Dec 2026)", color: COBALT_D,
        items: [
          "Churn Sentinel Phase 2: core banking + Banno digital signals",
          "CommercialSignal: personal-to-business conversion model",
          "Generational Wealth household relationship engine",
          "Exec dashboards + FI banker score explorer live",
        ],
        apps: [
          { label: "Churn Sentinel +",               color: A1, badge: "Core+Digital" },
          { label: "Zelle Memo Intelligence",        color: A2 },
          { label: "CommercialSignal",               color: A7 },
          { label: "Generational Wealth Deflection", color: A8 },
        ],
      },
      {
        num: "03", title: "Scale & Expand", period: "Q3 FY27–Q2 FY28  (Jan–Jun 2027+)", color: DK_GRAY,
        items: [
          "Anomaly Detection: real-time streaming fraud engine",
          "FI Decision Studio: AutoML platform GA release",
          "Core + Payments data modules fully activated",
          "Horizon apps: Cross-Sell, Overdraft, Sentiment (pilot)",
        ],
        apps: [
          { label: "Anomaly Detection",  color: A3 },
          { label: "FI Decision Studio", color: A6, badge: "Platform" },
          { label: "+ Horizon Pipeline", color: COBALT_D, badge: "Pilot" },
        ],
      },
    ];

    for (let i = 0; i < phases.length; i++) {
      const p = phases[i];
      const x = 0.28 + i * 3.22;

      s.addShape(pres.shapes.RECTANGLE, { x, y: 0.82, w: 3.08, h: 4.52, fill: { color: WHITE }, line: { color: LT_GRAY }, shadow: mkSh() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 0.82, w: 3.08, h: 0.88, fill: { color: p.color }, line: { color: p.color } });

      s.addText(p.num, {
        x: x + 0.12, y: 0.88, w: 0.68, h: 0.72,
        fontSize: 36, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
      });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.84, y: 0.98, w: 0.04, h: 0.52, fill: { color: WHITE, transparency: 40 }, line: { color: WHITE, transparency: 40 } });
      s.addText(p.title, {
        x: x + 0.96, y: 0.9, w: 2.0, h: 0.38,
        fontSize: 15, fontFace: FONT_H, bold: true, color: WHITE, margin: 0
      });
      s.addText(p.period, {
        x: x + 0.96, y: 1.3, w: 2.0, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: WHITE, margin: 0, italic: true
      });

      // Capability bullets
      const bulletItems = p.items.map((it, idx) => ({
        text: it,
        options: { bullet: true, breakLine: idx < p.items.length - 1, paraSpaceAfter: 8 }
      }));
      s.addText(bulletItems, {
        x: x + 0.18, y: 1.82, w: 2.78, h: 2.12,
        fontSize: 10, fontFace: FONT_B, color: DK_GRAY, margin: 0
      });

      // "APPS DELIVERING" divider bar — move up if many chips to avoid overflow
      const chipH   = p.apps.length > 3 ? 0.24 : 0.28;
      const chipSpc = p.apps.length > 3 ? 0.28 : 0.36;
      const divY    = p.apps.length > 3 ? 3.76 : 4.02;
      s.addShape(pres.shapes.RECTANGLE, { x, y: divY, w: 3.08, h: 0.26, fill: { color: p.color, transparency: 88 }, line: { color: p.color, transparency: 70 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y: divY, w: 0.06, h: 0.26, fill: { color: p.color }, line: { color: p.color } });
      s.addText("APPS DELIVERING THIS PHASE", {
        x: x + 0.14, y: divY, w: 2.9, h: 0.26,
        fontSize: 7.5, fontFace: FONT_H, bold: true, color: p.color, valign: "middle", margin: 0, charSpacing: 0.5
      });

      // App chips
      p.apps.forEach((app, ai) => {
        const chipY = divY + 0.30 + ai * chipSpc;
        s.addShape(pres.shapes.RECTANGLE, {
          x: x + 0.14, y: chipY, w: 2.76, h: chipH,
          fill: { color: app.color }, line: { color: app.color },
        });
        // Dot accent
        s.addShape(pres.shapes.OVAL, { x: x + 0.20, y: chipY + (chipH - 0.14) / 2, w: 0.14, h: 0.14, fill: { color: WHITE, transparency: 30 }, line: { color: WHITE, transparency: 30 } });
        // App name
        const nameW = app.badge ? 1.70 : 2.44;
        s.addText(app.label, {
          x: x + 0.40, y: chipY, w: nameW, h: chipH,
          fontSize: p.apps.length > 3 ? 8 : 9, fontFace: FONT_H, bold: true, color: WHITE, valign: "middle", margin: 0
        });
        // Badge if present
        if (app.badge) {
          s.addShape(pres.shapes.RECTANGLE, {
            x: x + 0.14 + nameW + 0.10, y: chipY + 0.03, w: 0.82, h: chipH - 0.06,
            fill: { color: WHITE, transparency: 20 }, line: { color: WHITE, transparency: 40 },
          });
          s.addText(app.badge, {
            x: x + 0.14 + nameW + 0.10, y: chipY + 0.03, w: 0.82, h: chipH - 0.06,
            fontSize: 7, fontFace: FONT_B, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0
          });
        }
      });
    }
  }

  // ==========================================================
  // SLIDE 15 — NEXT STEPS (dark closing)
  // ==========================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };

    // Decorative circles
    s.addShape(pres.shapes.OVAL, { x: 5.5, y: -2.0, w: 7.5, h: 7.5, fill: { color: COBALT, transparency: 90 }, line: { color: COBALT, transparency: 82 } });
    s.addShape(pres.shapes.OVAL, { x: 7.0, y: 1.2, w: 4.0, h: 4.0, fill: { color: TECH, transparency: 90 }, line: { color: TECH, transparency: 82 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: COBALT }, line: { color: COBALT } });

    s.addText("NEXT STEPS", {
      x: 0.4, y: 0.48, w: 6, h: 0.38,
      fontSize: 11, fontFace: FONT_H, bold: true, color: TECH, margin: 0, charSpacing: 4
    });
    s.addText("Approve the Roadmap.\nFund the Team.", {
      x: 0.4, y: 0.95, w: 7.2, h: 1.5,
      fontSize: 40, fontFace: FONT_H, bold: true, color: WHITE, margin: 0
    });

    const steps = [
      { num: "1", text: "JHBI approves headcount budget and infrastructure allocation for Phase 1" },
      { num: "2", text: "DS team begins recruiting: Data Engineers + ML Engineers (Month 1)" },
      { num: "3", text: "Cloud platform procurement and ML environment setup (Months 1–2)" },
      { num: "4", text: "Churn Sentinel Phase 1 scoping: ACH pipeline access + direct deposit monitoring spec (Month 2)" },
      { num: "5", text: "Bi-weekly DS steering committee established for roadmap governance" },
    ];

    steps.forEach((st, i) => {
      const y = 2.62 + i * 0.52;
      s.addShape(pres.shapes.OVAL, { x: 0.4, y: y - 0.01, w: 0.36, h: 0.36, fill: { color: COBALT }, line: { color: COBALT } });
      s.addText(st.num, {
        x: 0.4, y: y - 0.01, w: 0.36, h: 0.36,
        fontSize: 11, fontFace: FONT_H, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0
      });
      s.addText(st.text, {
        x: 0.88, y: y + 0.04, w: 6.5, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: LT_GRAY, margin: 0
      });
    });

    s.addText("Powering the Financial World™  |  Jack Henry Data & Analytics Team  |  March 2026", {
      x: 0.4, y: 5.28, w: 9.0, h: 0.26,
      fontSize: 9.5, fontFace: FONT_B, color: MD_GRAY, margin: 0
    });
  }

  // ==========================================================
  // WRITE FILE
  // ==========================================================
  const outPath = "/sessions/laughing-keen-fermat/mnt/outputs/JH_DS_Strategy_JHBI.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`✅ Written to: ${outPath}`);
}

buildDeck().catch(err => { console.error(err); process.exit(1); });
