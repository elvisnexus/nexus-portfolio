const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  BorderStyle, TabStopType, LevelFormat, AlignmentType,
} = require("docx");

// ---- palette (matches the existing CV) ----
const NAVY = "0C2340";
const AMBER = "B5720A";
const SLATE = "44515E";
const RULE_FAINT = "C7D0D8";
const FONT = "Calibri";
const RIGHT_TAB = 10900;

// ---- helpers ----
const t = (text, opts = {}) => new TextRun({ text, font: FONT, ...opts });

const link = (url, label, size = 16) =>
  new ExternalHyperlink({ link: url, children: [t(label, { color: AMBER, size, underline: {} })] });

const sectionHeader = (text) =>
  new Paragraph({
    spacing: { before: 190, after: 70 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, space: 3, color: AMBER } },
    children: [t(text, { bold: true, color: NAVY, size: 20, characterSpacing: 20 })],
  });

const groupLabel = (text) =>
  new Paragraph({
    spacing: { before: 130, after: 20 },
    children: [t(text, { bold: true, color: AMBER, size: 16, characterSpacing: 15 })],
  });

const bodyPara = (text) =>
  new Paragraph({ spacing: { after: 40 }, children: [t(text, { color: SLATE, size: 19 })] });

const skillLine = (label, value) =>
  new Paragraph({
    spacing: { after: 36 },
    children: [
      t(label + "  ", { bold: true, color: NAVY, size: 18 }),
      t(value, { color: SLATE, size: 18 }),
    ],
  });

const bullet = (text) =>
  new Paragraph({
    numbering: { reference: "cv-bullets", level: 0 },
    spacing: { after: 24 },
    children: [t(text, { color: SLATE, size: 18 })],
  });

const roleLine = (title, org, dates) =>
  new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: RIGHT_TAB }],
    spacing: { before: 120, after: 24 },
    children: [
      t(title, { bold: true, color: NAVY, size: 20 }),
      t("  —  " + org, { color: SLATE, size: 20 }),
      new TextRun({ text: "\t" + dates, font: FONT, italics: true, color: SLATE, size: 17 }),
    ],
  });

// compact project: one line (name — link/suffix — description) + one italic stack line
const project = ({ name, url, urlText, suffix, desc, stack }) => {
  const head = [t(name, { bold: true, color: NAVY, size: 19 })];
  if (url) {
    head.push(t("  —  ", { color: SLATE, size: 16 }), link(url, urlText, 16), t("   ", { size: 16 }));
  } else {
    head.push(t("  —  " + suffix + ".   ", { color: SLATE, size: 16 }));
  }
  head.push(t(desc, { color: SLATE, size: 18 }));
  return [
    new Paragraph({ spacing: { before: 90, after: 6 }, children: head }),
    new Paragraph({ spacing: { after: 42 }, children: [t(stack, { italics: true, color: AMBER, size: 15 })] }),
  ];
};

const eduLine = (bold, rest) =>
  new Paragraph({
    spacing: { after: 30 },
    children: [
      t(bold, { bold: true, color: NAVY, size: 19 }),
      t(rest, { color: SLATE, size: 18 }),
    ],
  });

// ---- build children ----
const kids = [];

// header
kids.push(
  new Paragraph({
    spacing: { after: 20 }, alignment: AlignmentType.CENTER,
    children: [t("ELVIS NOSAKHARE", { bold: true, color: NAVY, size: 40, characterSpacing: 10 })],
  }),
  new Paragraph({
    spacing: { after: 60 }, alignment: AlignmentType.CENTER,
    children: [t("IT Technical Lead & Full-Stack Developer", { bold: true, color: AMBER, size: 22 })],
  }),
  new Paragraph({
    spacing: { after: 30 }, alignment: AlignmentType.CENTER,
    children: [t("Benin City, Nigeria   |   elvisnosakhare@gmail.com   |   +234 816 178 4341   |   +250 798 972 967", { color: SLATE, size: 17 })],
  }),
  new Paragraph({
    spacing: { after: 30 }, alignment: AlignmentType.CENTER,
    children: [
      link("https://inexus.space/", "inexus.space", 17),
      t("   |   ", { color: SLATE, size: 17 }),
      link("https://github.com/elvisnexus", "github.com/elvisnexus", 17),
      t("   |   ", { color: SLATE, size: 17 }),
      link("https://www.terawork.com/services/profile/26660", "Terawork Profile", 17),
    ],
  }),
  new Paragraph({
    spacing: { before: 30, after: 90 }, alignment: AlignmentType.CENTER,
    children: [t("Most repositories are private under client NDAs — happy to share a live preview or walkthrough on request.", { italics: true, color: SLATE, size: 15 })],
  }),
  new Paragraph({
    spacing: { before: 20, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1, color: RULE_FAINT } },
    children: [],
  }),
);

// profile
kids.push(
  sectionHeader("PROFILE"),
  bodyPara(
    "IT technical lead and full-stack developer, currently IT Director for a Nigeria- and Rwanda-based financial-services group. I lead an eight-person engineering, design, and QA team, own the technical roadmap for every project, and manage the vendors and regulators that come with banking and insurance work. I stay hands-on — building production web and mobile products end-to-end in React, Next.js, TanStack Start, and TypeScript, with Claude Code as my daily driver and AI features on the Anthropic Claude API. Trained in Agile DSDM; pursuing PMP."
  ),
);

// leadership & delivery
kids.push(
  sectionHeader("LEADERSHIP & DELIVERY"),
  skillLine("Team:", "Lead 8 — 5 engineers, 2 designers, 1 QA — across Nigeria & Rwanda; own the technical roadmap for every project"),
  skillLine("Reporting:", "Report to the Executive Director; work directly with the Managing Director, CFO, and board of directors"),
  skillLine("External stakeholders:", "National Bank of Rwanda (BNR), Opportunity International, external auditors, legal teams, micro-insurance partners"),
  skillLine("Vendors & fleet:", "Manage core-banking / loan-management vendors (Lendsqr, Presta, Musoni); administer Hexnode MDM across 600+ staff devices"),
  skillLine("Method:", "Agile delivery (DSDM), Jira & sprint planning, prompt engineering, workflow thinking & automation, PMP (in progress)"),
);

// technical skills
kids.push(
  sectionHeader("TECHNICAL SKILLS"),
  skillLine("Front-End:", "HTML5, CSS3, JavaScript, TypeScript, React, Next.js, TanStack Start, responsive & mobile-first design"),
  skillLine("Back-End & Data:", "Node.js, REST APIs, PHP, Laravel, SQL, PostgreSQL, Supabase (RLS, RPCs), Prisma"),
  skillLine("AI:", "Claude Code (daily driver), Codex, Copilot, Lovable; Anthropic Claude API integration; LLM-powered product features; prompt engineering"),
  skillLine("Integrations:", "Anthropic Claude API, Paystack, Flutterwave, Stripe, Daily.co, Mux, CharmHealth EHR, Google Maps Platform"),
  skillLine("Cloud & Delivery:", "Vercel, Cloudflare Workers, AWS Lambda, pg_cron; Core Web Vitals, Open Graph/SEO, cross-browser QA; WordPress for landing pages"),
);

// professional experience
kids.push(
  sectionHeader("PROFESSIONAL EXPERIENCE"),
  roleLine("IT Director", "Standard Life — Nigeria & Rwanda", "2023 – Present"),
  bullet("Lead an 8-person engineering, design, and QA team (5 engineers, 2 designers, 1 QA) across Nigeria and Rwanda, owning the technical roadmap for every project; delivered 3 enterprise systems in 3 years — micro-insurance, an assets portal, and core-banking & digital lending."),
  bullet("Report to the Executive Director; work directly with the MD, CFO, and board, and externally with the National Bank of Rwanda (BNR), Opportunity International, external auditors, and legal teams."),
  bullet("Select and manage core-banking and loan-management vendors — Lendsqr, Presta, and Musoni — from evaluation through integration and support."),
  bullet("Manage and administer Hexnode MDM for 600+ staff — device enrolment, security policy, compliance, and app distribution across the fleet."),
  bullet("Run Agile/DSDM delivery with Jira sprint planning; progressed from IT Support to IT Director in four years and was awarded Best Staff, Standard Life (2022)."),
  new Paragraph({
    spacing: { before: 110, after: 24 },
    children: [
      t("Earlier — Standard Life:", { bold: true, color: NAVY, size: 19 }),
      t("  IT Manager (2022–2023) · Full-Stack Developer (2020–2022) · IT Support (2019–2020)", { color: SLATE, size: 18 }),
    ],
  }),
  bullet("Progressed from first-line technical support through hands-on full-stack development of internal web applications into IT leadership."),
);

// projects
kids.push(
  sectionHeader("PROJECTS"),
  groupLabel("AS IT TECHNICAL LEAD  —  STANDARD LIFE ENTERPRISE SYSTEMS"),
  ...project({
    name: "Core-Banking & Digital Loans Platform", suffix: "Standard Life, Rwanda",
    desc: "Rwanda's first digital loans application — origination, disbursement, and repayment on a core-banking backbone with reporting to the National Bank of Rwanda. Led the team and delivery end to end; 1,000+ loans processed in the first 3 months live.",
    stack: "Core banking · Digital lending · Lendsqr / Musoni / Presta",
  }),
  ...project({
    name: "Digital Micro-Insurance Platform", suffix: "Standard Life",
    desc: "Policy issuance, premium collection, and claims processing for a regulated micro-insurance product, delivered with micro-insurance partners and Opportunity International.",
    stack: "Insurance · Payments · Regulatory reporting",
  }),
  ...project({
    name: "Assets Portal", suffix: "Standard Life",
    desc: "Enterprise portal for tracking and managing company assets across Nigeria and Rwanda, replacing manual registers and spreadsheets.",
    stack: "Internal platform",
  }),
  groupLabel("AS FULL-STACK DEVELOPER  —  CLIENT & PERSONAL PRODUCTS"),
  ...project({
    name: "CityTaska", url: "https://citytaska.com", urlText: "citytaska.com",
    desc: "Trust-based local services marketplace for Africa — post a request, compare quotes, book and pay in-platform.",
    stack: "Next.js · Supabase · Paystack · Vercel",
  }),
  ...project({
    name: "NoWahala", url: "https://oyanowahala.com", urlText: "oyanowahala.com",
    desc: "Multi-vendor marketplace with escrow-protected payments, vendor KYC, fraud monitoring, and an admin back-office for payouts and disputes.",
    stack: "React · Supabase · Paystack · Flutterwave",
  }),
  ...project({
    name: "Koursa", url: "https://koursalearn.com", urlText: "koursalearn.com",
    desc: "Multi-tenant course platform — video courses with quizzes and certificates, bulk team enrolment, and live classes alongside on-demand lessons.",
    stack: "Next.js · Prisma · Stripe · Mux · Daily.co",
  }),
  ...project({
    name: "Rotary Club of Oregbeni", url: "https://rotarycluboforegbeni.vercel.app", urlText: "rotarycluboforegbeni.vercel.app",
    desc: "Club-management platform — live video meeting rooms, dues and contribution payments, and a member/admin portal replacing manual administration.",
    stack: "React · Supabase · Daily.co · Paystack",
  }),
  ...project({
    name: "HSIM — Healing Solutions & Integrative Medicine (Seattle, USA)", url: "https://app.hsim.org", urlText: "app.hsim.org",
    desc: "Lead implementation engineer. One-way CharmHealth (EHR) integration and provider operations app — ingests appointments via signed webhooks, lets providers log office and travel visits, computes routes and mileage via Google Maps, and runs semi-monthly payroll from effective-dated rates. HIPAA-conscious: row-level security on every table, no PHI in email or logs.",
    stack: "TanStack Start (React 19) · TypeScript · Supabase (PostgreSQL, RLS) · CharmHealth EHR API · Google Maps Routes API · Cloudflare Workers + AWS Lambda",
  }),
  ...project({
    name: "Hedeo Foods", url: "https://hedeofoods.store", urlText: "hedeofoods.store",
    desc: "UK food e-commerce and wholesale platform — online ordering, delivery tracking, and order history for consumers, restaurants, and hotels.",
    stack: "TanStack Start · Supabase · Stripe",
  }),
  ...project({
    name: "MicroVault", url: "https://microvaultmfi.com", urlText: "microvaultmfi.com",
    desc: "Digital microfinance platform for a Nigerian lending client — loan management, repayment tracking, savings accounts, and CBN-compliant audit reporting. In development, launching December 2026.",
    stack: "Full-Stack · Fintech · Compliance",
  }),
  ...project({
    name: "Profilo", url: "https://getprofilo.com", urlText: "getprofilo.com",
    desc: "My own AI-powered career toolkit — tailored CVs, cover letters, LinkedIn and professional bios, salary-negotiation letters, and interview prep from a user's profile, built on the Anthropic Claude API. In progress.",
    stack: "Anthropic Claude API · AI product · Career tools",
  }),
);

// education & certifications
kids.push(
  sectionHeader("EDUCATION & CERTIFICATIONS"),
  eduLine("BSc, Computer Science", "  —  University of Benin, Benin City, Edo State  ·  Second Class Upper Division  ·  2018"),
  eduLine("OND, Computer Science", "  —  Federal Polytechnic Auchi, Edo State  ·  Distinction  ·  2013"),
  new Paragraph({
    spacing: { before: 20, after: 0 },
    children: [
      t("Training:  ", { bold: true, color: NAVY, size: 18 }),
      t("Agile DSDM (trained) · Jira & sprint planning · Prompt engineering · Workflow thinking & automation · PMP (in progress)", { color: SLATE, size: 18 }),
    ],
  }),
);

// references
kids.push(
  sectionHeader("REFERENCES"),
  new Paragraph({
    spacing: { after: 26 },
    children: [
      t("Mr. Adedeji Olowe", { bold: true, color: NAVY, size: 19 }),
      t("  —  CEO, Lendsqr Technologies", { color: SLATE, size: 18 }),
      t("   ·   ", { color: SLATE, size: 17 }),
      link("mailto:adedeji@lendsqr.com", "adedeji@lendsqr.com", 17),
      t("  ·  +234 703 413 8291", { color: SLATE, size: 17 }),
    ],
  }),
  new Paragraph({
    spacing: { after: 26 },
    children: [
      t("Mr. Ugo Kelvin", { bold: true, color: NAVY, size: 19 }),
      t("  —  Co-Founder, Egis Technologies Ltd.", { color: SLATE, size: 18 }),
      t("   ·   +234 703 020 2868", { color: SLATE, size: 17 }),
    ],
  }),
  new Paragraph({
    spacing: { before: 4 },
    children: [t("Additional references available on request.", { italics: true, color: SLATE, size: 17 })],
  }),
);

// ---- document ----
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "cv-bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 230, hanging: 180 } } },
          },
        ],
      },
    ],
  },
  styles: { default: { document: { run: { font: FONT, size: 19, color: SLATE } } } },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 540, right: 660, bottom: 500, left: 660 },
        },
      },
      children: kids,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const out = process.argv[2] || "Elvis_Nosakhare_CV_Technical_Lead.docx";
  fs.writeFileSync(out, buf);
  console.log("wrote " + out);
});
