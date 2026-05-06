const PptxGenJS = require("pptxgenjs");

const pptx = new PptxGenJS();

// Theme colors
const BLUE_DARK = "1F3864";
const BLUE_MID = "2E75B6";
const BLUE_LIGHT = "BDD7EE";
const WHITE = "FFFFFF";
const YELLOW = "FFD966";
const RED = "C00000";
const ORANGE = "ED7D31";
const GREEN = "70AD47";
const GREY_BG = "F2F2F2";
const TEXT_DARK = "1A1A1A";
const TEXT_MID = "404040";

pptx.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 inches

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────
function addSlideBg(slide, colorHex) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: "100%", fill: { color: colorHex } });
}

function addHeader(slide, title, y = 0, h = 0.75) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y, w: "100%", h, fill: { color: BLUE_DARK } });
  slide.addText(title, {
    x: 0.3, y, w: 12.7, h,
    fontSize: 22, bold: true, color: WHITE,
    valign: "middle", align: "left",
  });
}

function addFooter(slide) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.1, w: "100%", h: 0.4, fill: { color: BLUE_MID } });
  slide.addText("Outstation Redundancy Program  |  Confidential  |  May 2026", {
    x: 0.3, y: 7.1, w: 12.7, h: 0.4,
    fontSize: 9, color: WHITE, valign: "middle", align: "left",
  });
}

function bullet(text, level = 0, color = TEXT_DARK, size = 14) {
  return { text, options: { bullet: { indent: 12 + level * 15 }, fontSize: size, color, indentLevel: level } };
}

// ─────────────────────────────────────────────
// SLIDE 1 — TITLE
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, BLUE_DARK);

  // Decorative accent bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 2.6, w: 0.18, h: 2.2, fill: { color: YELLOW } });

  slide.addText("Outstation Redundancy Program", {
    x: 0.5, y: 2.0, w: 12.3, h: 1.0,
    fontSize: 36, bold: true, color: WHITE, align: "left",
  });
  slide.addText("Execution & Legal Alignment Overview", {
    x: 0.5, y: 3.05, w: 12.3, h: 0.6,
    fontSize: 22, color: BLUE_LIGHT, align: "left",
  });
  slide.addText("May 2026", {
    x: 0.5, y: 3.7, w: 12.3, h: 0.45,
    fontSize: 16, color: YELLOW, align: "left", italic: true,
  });

  slide.addText("CONFIDENTIAL", {
    x: 0.5, y: 6.8, w: 12.3, h: 0.35,
    fontSize: 11, color: "AAAAAA", align: "left",
  });
}

// ─────────────────────────────────────────────
// SLIDE 2 — OBJECTIVE
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 2 — Objective");
  addFooter(slide);

  // Left panel
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.95, w: 5.8, h: 5.8, fill: { color: WHITE }, line: { color: "D0D0D0", width: 1 } });
  slide.addText("Objective", { x: 0.5, y: 1.0, w: 5.4, h: 0.5, fontSize: 16, bold: true, color: BLUE_DARK });
  slide.addText([
    bullet("Execute the outstation redundancy program across multiple jurisdictions"),
    bullet("Ensure:", 0, BLUE_MID, 13),
    bullet("Full legal compliance by country", 1, TEXT_MID, 13),
    bullet("Controlled and confidential execution", 1, TEXT_MID, 13),
    bullet("Minimal operational disruption", 1, TEXT_MID, 13),
  ], { x: 0.5, y: 1.55, w: 5.4, h: 4.8, fontSize: 14, color: TEXT_DARK, valign: "top", paraSpaceAfter: 6 });

  // Right panel
  slide.addShape(pptx.ShapeType.rect, { x: 6.5, y: 0.95, w: 6.5, h: 5.8, fill: { color: BLUE_DARK }, line: { color: BLUE_MID, width: 1 } });
  slide.addText("Scope", { x: 6.7, y: 1.1, w: 6.1, h: 0.5, fontSize: 16, bold: true, color: YELLOW });
  slide.addShape(pptx.ShapeType.rect, { x: 6.7, y: 1.7, w: 5.8, h: 1.1, fill: { color: BLUE_MID } });
  slide.addText("50 Employees", { x: 6.7, y: 1.75, w: 5.8, h: 0.55, fontSize: 28, bold: true, color: WHITE, align: "center" });
  slide.addText("across multiple regions", { x: 6.7, y: 2.3, w: 5.8, h: 0.4, fontSize: 14, color: BLUE_LIGHT, align: "center" });
  slide.addText("Multi-jurisdiction program requiring coordinated legal and HR execution across all affected countries.", {
    x: 6.7, y: 2.9, w: 5.8, h: 2.5,
    fontSize: 13, color: WHITE, valign: "top", wrap: true,
  });
}

// ─────────────────────────────────────────────
// SLIDE 3 — PROJECT TIMELINE
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 3 — Project Timeline (Corrected Reality View)");
  addFooter(slide);

  const phases = [
    { num: "1", label: "Legal Engagement", date: "Oct 2025", desc: "Dentons formally engaged. Employee data, contracts & cases shared.\nTrue project initiation (legal start).", color: BLUE_DARK },
    { num: "2", label: "Legal Structuring & Review", date: "Oct 2025 → Apr 2026", desc: "Multi-country legal assessment.\nSeparation structuring & risk evaluation.", color: BLUE_MID },
    { num: "3", label: "Formal Project Plan", date: "1 Apr \u2192 20 Apr 2026", desc: 'Internal "Legal Review" phase.', color: "2E86C1" },
    { num: "4", label: "Governance / Approval", date: "21 Apr → 6 May 2026", desc: "EXCOM approval window.", color: ORANGE },
    { num: "5", label: "Execution (Planned)", date: "Starting 7 May 2026 onward", desc: "Full execution begins upon legal clearance + EXCOM approval.", color: GREEN },
  ];

  const startX = 0.3;
  const boxW = 2.45;
  const gap = 0.12;
  const boxY = 1.0;
  const boxH = 5.6;

  phases.forEach((p, i) => {
    const x = startX + i * (boxW + gap);
    slide.addShape(pptx.ShapeType.rect, { x, y: boxY, w: boxW, h: boxH, fill: { color: p.color } });
    // Phase number circle-ish
    slide.addShape(pptx.ShapeType.ellipse, { x: x + boxW / 2 - 0.3, y: boxY + 0.15, w: 0.6, h: 0.6, fill: { color: YELLOW } });
    slide.addText(p.num, { x: x + boxW / 2 - 0.3, y: boxY + 0.15, w: 0.6, h: 0.6, fontSize: 16, bold: true, color: BLUE_DARK, align: "center", valign: "middle" });
    slide.addText("Phase " + p.num, { x, y: boxY + 0.85, w: boxW, h: 0.35, fontSize: 11, bold: true, color: YELLOW, align: "center" });
    slide.addText(p.label, { x: x + 0.1, y: boxY + 1.25, w: boxW - 0.2, h: 0.75, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.1, y: boxY + 2.1, w: boxW - 0.2, h: 0.45, fill: { color: "2A4A6A" } });
    slide.addText(p.date, { x: x + 0.1, y: boxY + 2.1, w: boxW - 0.2, h: 0.45, fontSize: 11, color: WHITE, bold: true, align: "center", valign: "middle" });
    slide.addText(p.desc, { x: x + 0.1, y: boxY + 2.7, w: boxW - 0.2, h: 3.0, fontSize: 11, color: "E0E0E0", align: "left", valign: "top", wrap: true });
  });
}

// ─────────────────────────────────────────────
// SLIDE 4 — KEY MILESTONE GAP
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 4 — Key Milestone Gap (Critical)");
  addFooter(slide);

  // Key Observation box
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.95, w: 12.7, h: 1.3, fill: { color: BLUE_LIGHT }, line: { color: BLUE_MID, width: 1 } });
  slide.addText("Key Observation", { x: 0.5, y: 0.98, w: 12.3, h: 0.4, fontSize: 13, bold: true, color: BLUE_DARK });
  slide.addText([
    { text: "Legal engagement started:  ", options: { bold: false, color: TEXT_DARK } },
    { text: "Oct 2025", options: { bold: true, color: BLUE_DARK } },
    { text: "        Legal review planned completion:  ", options: { bold: false, color: TEXT_DARK } },
    { text: "20 Apr 2026", options: { bold: true, color: BLUE_DARK } },
  ], { x: 0.5, y: 1.4, w: 12.3, h: 0.7, fontSize: 13 });

  // Gap box
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 2.5, w: 6.0, h: 2.6, fill: { color: "FFF2CC" }, line: { color: ORANGE, width: 2 } });
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 2.5, w: 6.0, h: 0.42, fill: { color: ORANGE } });
  slide.addText("⚠  Gap Identified", { x: 0.5, y: 2.5, w: 5.8, h: 0.42, fontSize: 13, bold: true, color: WHITE, valign: "middle" });
  slide.addText("No formal documented communication confirming:", { x: 0.5, y: 3.0, w: 5.5, h: 0.45, fontSize: 13, color: TEXT_DARK, bold: true });
  slide.addText([
    bullet("Legal review completion", 0, TEXT_DARK, 12),
    bullet('"Clearance to proceed"', 0, TEXT_DARK, 12),
  ], { x: 0.5, y: 3.5, w: 5.5, h: 1.4, valign: "top" });

  // Implication box
  slide.addShape(pptx.ShapeType.rect, { x: 6.8, y: 2.5, w: 6.2, h: 2.6, fill: { color: "DEEAF1" }, line: { color: BLUE_MID, width: 2 } });
  slide.addShape(pptx.ShapeType.rect, { x: 6.8, y: 2.5, w: 6.2, h: 0.42, fill: { color: BLUE_MID } });
  slide.addText("📌  Implication", { x: 7.0, y: 2.5, w: 5.8, h: 0.42, fontSize: 13, bold: true, color: WHITE, valign: "middle" });
  slide.addText("Execution depends on:", { x: 7.0, y: 3.0, w: 5.8, h: 0.45, fontSize: 13, color: TEXT_DARK, bold: true });
  slide.addText([
    bullet("Final legal confirmation", 0, TEXT_DARK, 12),
    bullet("Alignment on compliance readiness", 0, TEXT_DARK, 12),
  ], { x: 7.0, y: 3.5, w: 5.8, h: 1.4, valign: "top" });

  // Risk banner
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 5.4, w: 12.7, h: 0.7, fill: { color: RED } });
  slide.addText("⚠  Formal \"Clearance to Proceed\" confirmation is REQUIRED before execution commences", {
    x: 0.5, y: 5.4, w: 12.3, h: 0.7, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle",
  });
}

// ─────────────────────────────────────────────
// SLIDE 5 — WHAT "PROCEED" MEANS
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, 'Slide 5 — What "Proceed" Means');
  addFooter(slide);

  slide.addText("Executive approval is required to:", { x: 0.3, y: 0.9, w: 12.7, h: 0.45, fontSize: 15, bold: true, color: BLUE_DARK });

  const approvals = [
    { icon: "✦", label: "Proceed with termination of identified employees" },
    { icon: "✦", label: "Approve separation approach" },
    { icon: "✦", label: "Approve compensation schemes (legal-backed)" },
  ];
  approvals.forEach((a, i) => {
    slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 1.45 + i * 0.75, w: 12.7, h: 0.65, fill: { color: i % 2 === 0 ? BLUE_LIGHT : WHITE }, line: { color: "D0D0D0", width: 0.5 } });
    slide.addText(a.icon + "  " + a.label, { x: 0.6, y: 1.45 + i * 0.75, w: 12.3, h: 0.65, fontSize: 14, color: TEXT_DARK, valign: "middle" });
  });

  // Proceed = equation
  slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 3.8, w: 12.3, h: 2.85, fill: { color: BLUE_DARK } });
  slide.addText('"Proceed" means:', { x: 0.8, y: 3.9, w: 11.7, h: 0.45, fontSize: 15, bold: true, color: YELLOW });

  const items = [
    { label: "✅  Legal Clearance", x: 1.0, y: 4.45 },
    { label: "+", x: 5.4, y: 4.45, big: true },
    { label: "✅  EXCOM Approval", x: 6.3, y: 4.45 },
    { label: "→  Full Execution Begins", x: 3.5, y: 5.45 },
  ];
  items.forEach(it => {
    slide.addText(it.label, {
      x: it.x, y: it.y, w: it.big ? 0.8 : 4.5, h: 0.7,
      fontSize: it.big ? 28 : 16, bold: true,
      color: it.label.startsWith("→") ? YELLOW : WHITE,
      align: it.big ? "center" : "left", valign: "middle",
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 6 — EXECUTION PLAN
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 6 — Execution Plan (Next Phases)");
  addFooter(slide);

  const phases = [
    {
      num: "1", label: "Legal Formalities\n(Country Level)", date: "7 May → 30 May 2026",
      items: ["Local legal coordination", "Regulatory notifications", "Draft termination letters"],
      color: BLUE_DARK,
    },
    {
      num: "2", label: "Employee\nNotification", date: "1 Jun → 14 Jun 2026",
      items: ["Issue formal notifications", "Conduct 1:1 meetings", "Provide support resources"],
      color: BLUE_MID,
    },
    {
      num: "3", label: "Offboarding", date: "15 Jun → 28 Jun 2026",
      items: ["Knowledge transfer", "Equipment return", "Access revocation preparation"],
      color: "2E86C1",
    },
    {
      num: "4", label: "Final Settlement", date: "1 Jul → 14 Jul 2026",
      items: ["Process final payments", "Issue settlement agreements", "Legal sign-off"],
      color: ORANGE,
    },
    {
      num: "5", label: "System Closure", date: "1 Jul → 14 Jul 2026",
      items: ["Revoke system access", "Archive employee data", "Close HR records"],
      color: GREEN,
    },
  ];

  const bW = 2.45, bGap = 0.12, startX = 0.3, topY = 1.0, bH = 5.6;
  phases.forEach((p, i) => {
    const x = startX + i * (bW + bGap);
    slide.addShape(pptx.ShapeType.rect, { x, y: topY, w: bW, h: bH, fill: { color: p.color } });
    slide.addShape(pptx.ShapeType.ellipse, { x: x + bW / 2 - 0.28, y: topY + 0.12, w: 0.56, h: 0.56, fill: { color: YELLOW } });
    slide.addText(p.num, { x: x + bW / 2 - 0.28, y: topY + 0.12, w: 0.56, h: 0.56, fontSize: 14, bold: true, color: BLUE_DARK, align: "center", valign: "middle" });
    slide.addText("Phase " + p.num, { x, y: topY + 0.75, w: bW, h: 0.3, fontSize: 10, bold: true, color: YELLOW, align: "center" });
    slide.addText(p.label, { x: x + 0.1, y: topY + 1.1, w: bW - 0.2, h: 0.9, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.1, y: topY + 2.1, w: bW - 0.2, h: 0.42, fill: { color: "2A4A6A" } });
    slide.addText(p.date, { x: x + 0.1, y: topY + 2.1, w: bW - 0.2, h: 0.42, fontSize: 10, color: WHITE, bold: true, align: "center", valign: "middle" });
    const bulletsText = p.items.map(it => "• " + it).join("\n");
    slide.addText(bulletsText, { x: x + 0.1, y: topY + 2.65, w: bW - 0.2, h: 2.7, fontSize: 10, color: "E0E0E0", valign: "top", wrap: true });
  });
}

// ─────────────────────────────────────────────
// SLIDE 7 — ROLES & OWNERSHIP
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 7 — Roles & Ownership");
  addFooter(slide);

  const rows = [
    { area: "Legal Review & Compliance", owner: "Ahsan", icon: "⚖" },
    { area: "EXCOM Approval", owner: "Kumail", icon: "✔" },
    { area: "Employee Notification & Offboarding", owner: "Hiba", icon: "👥" },
    { area: "Final Settlements", owner: "Saif", icon: "💼" },
    { area: "System & Access Removal", owner: "Ridha", icon: "🔒" },
  ];

  // Table header
  slide.addShape(pptx.ShapeType.rect, { x: 1.5, y: 1.05, w: 7.5, h: 0.5, fill: { color: BLUE_DARK } });
  slide.addText("Area", { x: 1.7, y: 1.05, w: 5.0, h: 0.5, fontSize: 14, bold: true, color: WHITE, valign: "middle" });
  slide.addText("Owner", { x: 6.7, y: 1.05, w: 2.1, h: 0.5, fontSize: 14, bold: true, color: WHITE, valign: "middle", align: "center" });

  rows.forEach((r, i) => {
    const y = 1.65 + i * 0.85;
    const bg = i % 2 === 0 ? WHITE : BLUE_LIGHT;
    slide.addShape(pptx.ShapeType.rect, { x: 1.5, y, w: 7.5, h: 0.75, fill: { color: bg }, line: { color: "D0D0D0", width: 0.5 } });
    // icon cell
    slide.addShape(pptx.ShapeType.rect, { x: 1.5, y, w: 0.55, h: 0.75, fill: { color: BLUE_MID } });
    slide.addText(r.icon, { x: 1.5, y, w: 0.55, h: 0.75, fontSize: 16, align: "center", valign: "middle" });
    slide.addText(r.area, { x: 2.1, y, w: 4.5, h: 0.75, fontSize: 13, color: TEXT_DARK, valign: "middle" });
    // Owner badge
    slide.addShape(pptx.ShapeType.rect, { x: 6.7, y: y + 0.12, w: 2.1, h: 0.5, fill: { color: BLUE_DARK }, rectRadius: 0.1 });
    slide.addText(r.owner, { x: 6.7, y: y + 0.12, w: 2.1, h: 0.5, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
  });
}

// ─────────────────────────────────────────────
// SLIDE 8 — FINANCIAL IMPACT
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 8 — Financial Impact");
  addFooter(slide);

  // Big stat left
  slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.1, w: 5.6, h: 5.4, fill: { color: BLUE_DARK } });
  slide.addText("Monthly Savings", { x: 0.7, y: 1.3, w: 5.2, h: 0.5, fontSize: 15, bold: true, color: YELLOW, align: "center" });
  slide.addText("OMR", { x: 0.7, y: 2.0, w: 5.2, h: 0.5, fontSize: 18, color: BLUE_LIGHT, align: "center" });
  slide.addText("33,281", { x: 0.7, y: 2.45, w: 5.2, h: 1.1, fontSize: 48, bold: true, color: WHITE, align: "center" });
  slide.addText("per month", { x: 0.7, y: 3.6, w: 5.2, h: 0.4, fontSize: 13, color: BLUE_LIGHT, align: "center", italic: true });

  slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 4.2, w: 5.0, h: 0.03, fill: { color: YELLOW } });
  slide.addText("Starting from execution completion", { x: 0.7, y: 4.3, w: 5.2, h: 0.5, fontSize: 11, color: "AAAAAA", align: "center" });

  // Big stat right
  slide.addShape(pptx.ShapeType.rect, { x: 7.0, y: 1.1, w: 5.8, h: 5.4, fill: { color: BLUE_MID } });
  slide.addText("Aug–Dec 2026 Impact", { x: 7.2, y: 1.3, w: 5.4, h: 0.5, fontSize: 15, bold: true, color: WHITE, align: "center" });
  slide.addText("Total Projected Savings", { x: 7.2, y: 1.9, w: 5.4, h: 0.4, fontSize: 12, color: BLUE_LIGHT, align: "center" });
  slide.addText("OMR", { x: 7.2, y: 2.35, w: 5.4, h: 0.5, fontSize: 18, color: WHITE, align: "center" });
  slide.addText("166,405", { x: 7.2, y: 2.8, w: 5.4, h: 1.0, fontSize: 44, bold: true, color: YELLOW, align: "center" });
  slide.addText("(5 months)", { x: 7.2, y: 3.85, w: 5.4, h: 0.35, fontSize: 12, color: WHITE, align: "center", italic: true });

  slide.addShape(pptx.ShapeType.rect, { x: 7.2, y: 4.35, w: 5.0, h: 0.03, fill: { color: WHITE } });
  slide.addText("Strong contribution to 2026 cost optimization targets", { x: 7.2, y: 4.45, w: 5.4, h: 0.6, fontSize: 11, color: WHITE, align: "center", wrap: true });

  // Arrow connector (decorative)
  slide.addShape(pptx.ShapeType.rect, { x: 6.1, y: 3.1, w: 0.8, h: 0.55, fill: { color: YELLOW } });
  slide.addText("×5", { x: 6.1, y: 3.1, w: 0.8, h: 0.55, fontSize: 18, bold: true, color: BLUE_DARK, align: "center", valign: "middle" });
}

// ─────────────────────────────────────────────
// SLIDE 9 — RISKS & DEPENDENCIES
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 9 — Risks & Dependencies");
  addFooter(slide);

  // Risks
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.95, w: 6.0, h: 5.7, fill: { color: WHITE }, line: { color: RED, width: 2 } });
  slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.95, w: 6.0, h: 0.5, fill: { color: RED } });
  slide.addText("🔴  Key Risks", { x: 0.5, y: 0.95, w: 5.8, h: 0.5, fontSize: 14, bold: true, color: WHITE, valign: "middle" });

  const risks = [
    { text: "Lack of formal legal sign-off record", severity: "HIGH" },
    { text: "Country-specific compliance delays", severity: "MEDIUM" },
    { text: "Regulatory approval timeline overruns", severity: "MEDIUM" },
  ];
  risks.forEach((r, i) => {
    const y = 1.65 + i * 1.35;
    slide.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 5.6, h: 1.1, fill: { color: i === 0 ? "FDECEA" : "FFF3E0" }, line: { color: "E0E0E0", width: 0.5 } });
    slide.addShape(pptx.ShapeType.rect, { x: 4.8, y: y + 0.28, w: 1.1, h: 0.45, fill: { color: i === 0 ? RED : ORANGE } });
    slide.addText(r.severity, { x: 4.8, y: y + 0.28, w: 1.1, h: 0.45, fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle" });
    slide.addText(r.text, { x: 0.65, y: y + 0.18, w: 4.0, h: 0.75, fontSize: 12, color: TEXT_DARK, valign: "middle", wrap: true });
  });

  // Dependencies
  slide.addShape(pptx.ShapeType.rect, { x: 7.0, y: 0.95, w: 6.0, h: 5.7, fill: { color: WHITE }, line: { color: ORANGE, width: 2 } });
  slide.addShape(pptx.ShapeType.rect, { x: 7.0, y: 0.95, w: 6.0, h: 0.5, fill: { color: ORANGE } });
  slide.addText("🟡  Dependencies", { x: 7.2, y: 0.95, w: 5.8, h: 0.5, fontSize: 14, bold: true, color: WHITE, valign: "middle" });

  const deps = [
    "Legal final confirmation",
    "EXCOM decision",
    "Local legal execution readiness",
  ];
  deps.forEach((d, i) => {
    const y = 1.65 + i * 1.35;
    slide.addShape(pptx.ShapeType.rect, { x: 7.2, y, w: 5.6, h: 1.1, fill: { color: "FFF8E1" }, line: { color: "E0E0E0", width: 0.5 } });
    slide.addShape(pptx.ShapeType.ellipse, { x: 7.35, y: y + 0.3, w: 0.45, h: 0.45, fill: { color: ORANGE } });
    slide.addText(String(i + 1), { x: 7.35, y: y + 0.3, w: 0.45, h: 0.45, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    slide.addText(d, { x: 7.95, y: y + 0.18, w: 4.7, h: 0.75, fontSize: 13, color: TEXT_DARK, valign: "middle" });
  });
}

// ─────────────────────────────────────────────
// SLIDE 10 — IMMEDIATE NEXT STEPS
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, GREY_BG);
  addHeader(slide, "Slide 10 — Immediate Next Steps");
  addFooter(slide);

  const steps = [
    {
      num: "1", title: "Legal Alignment", color: BLUE_DARK,
      items: ["Obtain written confirmation: \"Legal review complete\"", "Confirm \"Cleared to proceed\" from Dentons"],
    },
    {
      num: "2", title: "Governance Closure", color: BLUE_MID,
      items: ["Confirm EXCOM approval status", "Document formal sign-off"],
    },
    {
      num: "3", title: "Execution Readiness", color: GREEN,
      items: ["Prepare legal documentation by country", "Prepare notification materials", "Prepare HR case handling procedures"],
    },
  ];

  steps.forEach((s, i) => {
    const x = 0.3 + i * 4.35;
    slide.addShape(pptx.ShapeType.rect, { x, y: 1.0, w: 4.1, h: 5.7, fill: { color: s.color } });
    slide.addShape(pptx.ShapeType.ellipse, { x: x + 1.75, y: 1.1, w: 0.6, h: 0.6, fill: { color: YELLOW } });
    slide.addText(s.num, { x: x + 1.75, y: 1.1, w: 0.6, h: 0.6, fontSize: 16, bold: true, color: BLUE_DARK, align: "center", valign: "middle" });
    slide.addText(s.title, { x: x + 0.1, y: 1.8, w: 3.9, h: 0.75, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.15, y: 2.65, w: 3.8, h: 0.03, fill: { color: YELLOW } });
    s.items.forEach((item, j) => {
      const iy = 2.8 + j * 1.0;
      slide.addShape(pptx.ShapeType.rect, { x: x + 0.15, y: iy, w: 3.8, h: 0.85, fill: { color: "1A3A5A" } });
      slide.addText("✅  " + item, { x: x + 0.25, y: iy + 0.05, w: 3.6, h: 0.75, fontSize: 11, color: WHITE, valign: "middle", wrap: true });
    });
  });
}

// ─────────────────────────────────────────────
// SLIDE 11 — KEY MESSAGE TO LEADERSHIP
// ─────────────────────────────────────────────
{
  const slide = pptx.addSlide();
  addSlideBg(slide, BLUE_DARK);

  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: "100%", fill: { color: YELLOW } });
  slide.addText("Key Message to Leadership", {
    x: 0.5, y: 0.3, w: 12.3, h: 0.6,
    fontSize: 22, bold: true, color: WHITE, align: "left",
  });
  slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.95, w: 12.3, h: 0.04, fill: { color: YELLOW } });

  const msgs = [
    { icon: "📅", text: "Project initiated Oct 2025 with Dentons" },
    { icon: "📋", text: "Structured into formal execution plan Apr–Jul 2026" },
    { icon: "🔄", text: "Currently at transition point into execution" },
    { icon: "⚖", text: "Final legal clearance confirmation is being aligned" },
  ];
  msgs.forEach((m, i) => {
    const y = 1.2 + i * 1.1;
    slide.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.3, h: 0.9, fill: { color: i % 2 === 0 ? "253B5A" : "1F3864" } });
    slide.addText(m.icon, { x: 0.7, y, w: 0.7, h: 0.9, fontSize: 20, align: "center", valign: "middle" });
    slide.addText(m.text, { x: 1.5, y, w: 11.1, h: 0.9, fontSize: 16, color: WHITE, valign: "middle" });
  });

  // Call to action
  slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.7, w: 12.3, h: 1.1, fill: { color: YELLOW } });
  slide.addText("Once confirmed → Immediate execution can proceed", {
    x: 0.7, y: 5.7, w: 11.9, h: 1.1,
    fontSize: 20, bold: true, color: BLUE_DARK, align: "center", valign: "middle",
  });

  addFooter(slide);
}

// ─────────────────────────────────────────────
// SAVE
// ─────────────────────────────────────────────
pptx.writeFile({ fileName: "Outstation_Redundancy_Program_May2026.pptx" })
  .then(() => console.log("✅  Saved: Outstation_Redundancy_Program_May2026.pptx"))
  .catch(err => { console.error("Error:", err); process.exit(1); });
