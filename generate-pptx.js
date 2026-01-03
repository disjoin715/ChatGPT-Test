const PptxGenJS = require("pptxgenjs");

const palette = {
  deepNavy: "0F2439",
  midNavy: "17395c",
  sky: "7fb4e0",
  gold: "e1b44c",
  softGray: "f4f6f8",
  text: "0c1b2a",
  muted: "4a5c70",
};

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";

const sharedText = {
  fontFace: "Inter",
  fontSize: 14,
  color: palette.text,
};

const scenarios = [
  {
    title: "Baseline Resourcing",
    pill: "4 FTE · BAU + Incremental Change",
    pillFill: "f2f6fb",
    pillColor: palette.midNavy,
    metrics: [
      { label: "Capacity", value: "~60%" },
      { label: "Change bandwidth", value: "2 tracks" },
      { label: "Risk posture", value: "Moderate" },
    ],
    bullets: [
      {
        title: "Maintain core operations",
        detail:
          "Focus on stability and regulatory reporting; limited innovation bandwidth.",
      },
      {
        title: "Sequential delivery",
        detail:
          "Two initiative tracks in sequence; elongates policy pilots and analytics upgrades.",
      },
      {
        title: "Deferred optimization",
        detail:
          "Automation and resilience improvements shift beyond the 18-month window.",
      },
    ],
  },
  {
    title: "More Resource (Preferred)",
    pill: "+3 FTE · Data, Ops, Delivery",
    pillFill: "fdf6e8",
    pillColor: "6f5316",
    metrics: [
      { label: "Capacity", value: "~90%" },
      { label: "Change bandwidth", value: "4 tracks" },
      { label: "Risk posture", value: "Lower" },
    ],
    bullets: [
      {
        title: "Parallel delivery",
        detail:
          "Run concurrent workstreams (policy pilots, data modernization, resiliency uplift) without sacrificing BAU.",
      },
      {
        title: "Faster regulatory readiness",
        detail:
          "Expedite compliance changes and scenario testing with embedded risk & controls support.",
      },
      {
        title: "Data & automation gains",
        detail:
          "Deliver prioritized automations, observability, and analytics that reduce manual effort and downtime.",
      },
    ],
  },
];

function addHeader(slide) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.2,
    y: 0.2,
    w: 9.6,
    h: 5.2,
    fill: "eef2f7",
    line: { color: "eef2f7" },
    shadow: { type: "outer", opacity: 15, blur: 8, offset: 1, angle: 90 },
  });

  slide.addText("18-Month Outlook · Resource Strategy", {
    ...sharedText,
    x: 0.6,
    y: 0.45,
    fontSize: 12,
    color: palette.muted,
    bold: true,
    charSpacing: 120,
  });

  slide.addText("Baseline vs. Accelerated Delivery", {
    ...sharedText,
    x: 0.6,
    y: 0.75,
    fontSize: 28,
    color: palette.deepNavy,
    bold: true,
  });

  slide.addText(
    "How resourcing choices shape delivery outcomes and policy readiness.",
    {
      ...sharedText,
      x: 0.6,
      y: 1.15,
      fontSize: 14,
      color: palette.muted,
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.6,
    y: 0.5,
    w: 2.3,
    h: 0.65,
    fill: palette.deepNavy,
    line: { color: palette.deepNavy },
    shadow: { type: "outer", opacity: 35, blur: 6, offset: 1.1, angle: 90 },
    rectRadius: 10,
  });

  slide.addShape(pptx.ShapeType.ellipse, {
    x: 7.75,
    y: 0.62,
    w: 0.25,
    h: 0.25,
    fill: palette.gold,
    line: { color: palette.gold },
  });

  slide.addText("Q4 FY24 \u2192 Q1 FY26", {
    ...sharedText,
    x: 8.05,
    y: 0.62,
    fontSize: 14,
    color: "FFFFFF",
    bold: true,
  });
}

function addMetrics(slide, x, y, metrics, columnWidth) {
  metrics.forEach((metric, idx) => {
    const metricX = x + idx * (columnWidth + 0.12);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: metricX,
      y,
      w: columnWidth,
      h: 0.9,
      fill: palette.softGray,
      line: { color: "e1e7ee" },
      rectRadius: 10,
    });

    slide.addText(metric.label.toUpperCase(), {
      ...sharedText,
      x: metricX + 0.15,
      y: y + 0.12,
      fontSize: 10,
      color: palette.muted,
      bold: true,
      charSpacing: 40,
    });

    slide.addText(metric.value, {
      ...sharedText,
      x: metricX + 0.15,
      y: y + 0.42,
      fontSize: 18,
      color: palette.deepNavy,
      bold: true,
    });
  });
}

function addBulletList(slide, x, startY, bullets) {
  let currentY = startY;
  bullets.forEach((item) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x,
      y: currentY + 0.05,
      w: 0.15,
      h: 0.15,
      fill: palette.gold,
      line: { color: palette.gold },
      shadow: { type: "outer", opacity: 25, blur: 4, offset: 0.08, angle: 90 },
    });

    slide.addText(item.title, {
      ...sharedText,
      x: x + 0.25,
      y: currentY,
      fontSize: 15,
      color: palette.deepNavy,
      bold: true,
    });

    slide.addText(item.detail, {
      ...sharedText,
      x: x + 0.25,
      y: currentY + 0.26,
      w: 3.8,
      fontSize: 12,
      color: palette.muted,
    });

    currentY += 0.7;
  });
}

function addScenarioCard(slide, x, scenario) {
  const y = 1.6;
  const cardWidth = 4.6;
  const cardHeight = 3.35;

  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: cardWidth,
    h: cardHeight,
    fill: "FFFFFF",
    line: { color: "dde5ef" },
    rectRadius: 12,
    shadow: { type: "outer", opacity: 18, blur: 7, offset: 0.8, angle: 90 },
  });

  slide.addText(scenario.title, {
    ...sharedText,
    x: x + 0.3,
    y: y + 0.25,
    fontSize: 18,
    color: palette.deepNavy,
    bold: true,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.3,
    y: y + 0.7,
    w: 3.5,
    h: 0.55,
    fill: scenario.pillFill,
    line: { color: scenario.pillFill },
    rectRadius: 20,
  });

  slide.addText(scenario.pill, {
    ...sharedText,
    x: x + 0.45,
    y: y + 0.77,
    fontSize: 12,
    color: scenario.pillColor,
    bold: true,
  });

  const metricsX = x + 0.3;
  const metricsY = y + 1.05;
  const metricWidth = 1.25;
  addMetrics(slide, metricsX, metricsY, scenario.metrics, metricWidth);

  addBulletList(slide, x + 0.35, y + 2.0, scenario.bullets);
}

function addAsk(slide) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 5.05,
    w: 8.9,
    h: 0.85,
    fill: {
      type: "solid",
      color: palette.deepNavy,
    },
    line: { color: palette.deepNavy },
    rectRadius: 12,
    shadow: { type: "outer", opacity: 20, blur: 6, offset: 0.6, angle: 90 },
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 5.18,
    w: 0.6,
    h: 0.6,
    fill: { color: "FFFFFF", transparency: 70 },
    line: { color: "FFFFFF", transparency: 70 },
    rectRadius: 12,
  });

  slide.addText("\u2605", {
    ...sharedText,
    x: 0.85,
    y: 5.28,
    fontSize: 24,
    color: "FFFFFF",
    bold: true,
  });

  slide.addText("Our ask: Approve +3 FTE for 18 months", {
    ...sharedText,
    x: 1.45,
    y: 5.12,
    fontSize: 16,
    color: "FFFFFF",
    bold: true,
  });

  slide.addText(
    "Enables four concurrent tracks, accelerates policy pilots, and reduces operational risk while safeguarding BAU.",
    {
      ...sharedText,
      x: 1.45,
      y: 5.36,
      w: 5.6,
      fontSize: 12,
      color: "d5e2f2",
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.4,
    y: 5.17,
    w: 1.8,
    h: 0.56,
    fill: palette.gold,
    line: { color: palette.gold },
    rectRadius: 10,
    shadow: { type: "outer", opacity: 35, blur: 6, offset: 1, angle: 90 },
  });

  slide.addText("Proceed with Preferred Plan", {
    ...sharedText,
    x: 7.53,
    y: 5.29,
    fontSize: 12,
    color: "1f1606",
    bold: true,
  });
}

function buildSlide() {
  const slide = pptx.addSlide();
  slide.background = { color: "ffffff" };

  addHeader(slide);
  addScenarioCard(slide, 0.6, scenarios[0]);
  addScenarioCard(slide, 5.0, scenarios[1]);
  addAsk(slide);
}

buildSlide();

pptx
  .writeFile({ fileName: "Resource_Strategy_Outlook.pptx" })
  .then(() => console.log("Presentation created: Resource_Strategy_Outlook.pptx"));
