const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

const SLIDE = { widthIn: 10, heightIn: 5.625 };
const DESIGN = {
  widthPx: 1280,
  heightPx: 720,
  shellMarginXPx: 28,
  shellMarginYPx: 24,
  paddingXPx: 30,
  paddingYPx: 26,
  columnGapPx: 18,
  verticalGapPx: 18,
  headerHeightPx: 118,
  askHeightPx: 96,
  cardPaddingPx: 22,
  metricGapPx: 10,
  listGapPx: 52,
};

const PX_PER_IN = DESIGN.widthPx / SLIDE.widthIn;

const palette = {
  deepNavy: "0F2439",
  midNavy: "17395c",
  sky: "7fb4e0",
  gold: "e1b44c",
  softGray: "f4f6f8",
  text: "0c1b2a",
  muted: "4a5c70",
};

const sharedText = {
  fontFace: "Segoe UI",
  fontSize: 12,
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

const clamp = (value, min, max) => {
  if (!Number.isFinite(value)) return min;
  if (typeof min === "number") value = Math.max(min, value);
  if (typeof max === "number") value = Math.min(max, value);
  return value;
};

const pxToIn = (px) => Number((clamp(px, 0, 5000) / PX_PER_IN).toFixed(4));
const pxToPt = (px) => Number(clamp((px * 72) / 96, 8, 64).toFixed(2));

const asBox = (xPx, yPx, wPx, hPx) => ({
  x: pxToIn(xPx),
  y: pxToIn(yPx),
  w: pxToIn(Math.max(wPx, 4)),
  h: pxToIn(Math.max(hPx, 4)),
});

const computeLayout = () => {
  const shell = {
    xPx: DESIGN.shellMarginXPx,
    yPx: DESIGN.shellMarginYPx,
    wPx: DESIGN.widthPx - DESIGN.shellMarginXPx * 2,
    hPx: DESIGN.heightPx - DESIGN.shellMarginYPx * 2,
    radius: 22,
  };

  const content = {
    xPx: shell.xPx + DESIGN.paddingXPx,
    yPx: shell.yPx + DESIGN.paddingYPx,
    wPx: shell.wPx - DESIGN.paddingXPx * 2,
    hPx: shell.hPx - DESIGN.paddingYPx * 2,
  };

  const cardsHeightPx =
    content.hPx - DESIGN.headerHeightPx - DESIGN.askHeightPx - DESIGN.verticalGapPx * 2;

  const leftWidthPx =
    (content.wPx - DESIGN.columnGapPx) * (1.1 / (1.1 + 0.9));
  const rightWidthPx =
    (content.wPx - DESIGN.columnGapPx) * (0.9 / (1.1 + 0.9));

  return {
    shell,
    content,
    header: {
      eyebrowBox: {
        ...asBox(content.xPx, content.yPx, content.wPx * 0.7, 24),
      },
      titleBox: {
        ...asBox(content.xPx, content.yPx + 26, content.wPx * 0.72, 44),
      },
      subtitleBox: {
        ...asBox(content.xPx, content.yPx + 70, content.wPx * 0.7, 28),
      },
      badge: {
        ...asBox(
          content.xPx + content.wPx - 220,
          content.yPx + 6,
          220,
          52
        ),
        radius: 12,
      },
    },
    cards: {
      yPx: content.yPx + DESIGN.headerHeightPx + DESIGN.verticalGapPx,
      heightPx: cardsHeightPx,
      leftWidthPx,
      rightWidthPx,
    },
    ask: {
      xPx: content.xPx,
      yPx:
        content.yPx +
        DESIGN.headerHeightPx +
        DESIGN.verticalGapPx +
        cardsHeightPx +
        DESIGN.verticalGapPx,
      wPx: content.wPx,
      hPx: DESIGN.askHeightPx,
    },
  };
};

const addShell = (slide, layout, pptx) => {
  slide.background = { color: palette.deepNavy };

  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(layout.shell.xPx, layout.shell.yPx, layout.shell.wPx, layout.shell.hPx),
    fill: "f6f9fe",
    line: { color: "dce3ed" },
    rectRadius: layout.shell.radius,
    shadow: { type: "outer", opacity: 25, blur: 9, offset: 0.2, angle: 90 },
  });
};

const addHeader = (slide, pptx, layout) => {
  slide.addText("18-Month Outlook · Resource Strategy", {
    ...sharedText,
    ...layout.header.eyebrowBox,
    fontSize: pxToPt(14),
    color: palette.muted,
    bold: true,
    charSpacing: 120,
  });

  slide.addText("Baseline vs. Accelerated Delivery", {
    ...sharedText,
    ...layout.header.titleBox,
    fontSize: pxToPt(32),
    color: palette.deepNavy,
    bold: true,
  });

  slide.addText(
    "How resourcing choices shape delivery outcomes and policy readiness.",
    {
      ...sharedText,
      ...layout.header.subtitleBox,
      fontSize: pxToPt(16),
      color: palette.muted,
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    ...layout.header.badge,
    fill: palette.deepNavy,
    line: { color: palette.deepNavy },
    rectRadius: 12,
    shadow: { type: "outer", opacity: 32, blur: 7, offset: 0.18, angle: 90 },
  });

  slide.addShape(pptx.ShapeType.ellipse, {
    x: layout.header.badge.x + pxToIn(12),
    y: layout.header.badge.y + pxToIn(15),
    w: pxToIn(12),
    h: pxToIn(12),
    fill: palette.gold,
    line: { color: palette.gold },
    shadow: { type: "outer", opacity: 30, blur: 6, offset: 0.05, angle: 90 },
  });

  slide.addText("Q4 FY24 → Q1 FY26", {
    ...sharedText,
    x: layout.header.badge.x + pxToIn(30),
    y: layout.header.badge.y + pxToIn(14),
    w: layout.header.badge.w - pxToIn(40),
    h: pxToIn(26),
    fontSize: pxToPt(14),
    color: "FFFFFF",
    bold: true,
  });
};

const addMetrics = (slide, pptx, xPx, yPx, widthPx, metrics) => {
  const metricGapPx = DESIGN.metricGapPx;
  const metricHeightPx = 96;
  const metricWidthPx =
    (widthPx - metricGapPx * 2) / 3;

  metrics.forEach((metric, idx) => {
    const metricX = xPx + idx * (metricWidthPx + metricGapPx);
    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(metricX, yPx, metricWidthPx, metricHeightPx),
      fill: palette.softGray,
      line: { color: "e1e7ee" },
      rectRadius: 12,
    });

    slide.addText(metric.label.toUpperCase(), {
      ...sharedText,
      ...asBox(metricX + 12, yPx + 12, metricWidthPx - 24, 18),
      fontSize: pxToPt(12),
      color: palette.muted,
      bold: true,
      charSpacing: 40,
    });

    slide.addText(metric.value, {
      ...sharedText,
      ...asBox(metricX + 12, yPx + 38, metricWidthPx - 24, 36),
      fontSize: pxToPt(20),
      color: palette.deepNavy,
      bold: true,
    });
  });
};

const addBulletList = (slide, pptx, xPx, startYPx, availableWidthPx, bullets) => {
  let currentY = startYPx;
  const lineHeightPx = 54;

  bullets.forEach((item) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: pxToIn(xPx),
      y: pxToIn(currentY + 6),
      w: pxToIn(10),
      h: pxToIn(10),
      fill: palette.gold,
      line: { color: palette.gold },
      shadow: { type: "outer", opacity: 26, blur: 5, offset: 0.06, angle: 90 },
    });

    slide.addText(item.title, {
      ...sharedText,
      ...asBox(xPx + 18, currentY, availableWidthPx - 18, 22),
      fontSize: pxToPt(15),
      color: palette.deepNavy,
      bold: true,
    });

    slide.addText(item.detail, {
      ...sharedText,
      ...asBox(xPx + 18, currentY + 18, availableWidthPx - 18, 40),
      fontSize: pxToPt(13),
      color: palette.muted,
    });

    currentY += lineHeightPx;
  });
};

const addScenarioCard = (slide, pptx, xPx, cardWidthPx, scenario, layout) => {
  const yPx = layout.cards.yPx;

  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(xPx, yPx, cardWidthPx, layout.cards.heightPx),
    fill: "FFFFFF",
    line: { color: "dde5ef" },
    rectRadius: 14,
    shadow: { type: "outer", opacity: 16, blur: 7, offset: 0.14, angle: 90 },
  });

  const paddingPx = DESIGN.cardPaddingPx;
  const textWidthPx = cardWidthPx - paddingPx * 2;

  slide.addText(scenario.title, {
    ...sharedText,
    ...asBox(xPx + paddingPx, yPx + paddingPx, textWidthPx, 26),
    fontSize: pxToPt(18),
    color: palette.deepNavy,
    bold: true,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(xPx + paddingPx, yPx + paddingPx + 30, 220, 34),
    fill: scenario.pillFill,
    line: { color: scenario.pillFill },
    rectRadius: 20,
  });

  slide.addText(scenario.pill, {
    ...sharedText,
    ...asBox(xPx + paddingPx + 12, yPx + paddingPx + 34, 200, 24),
    fontSize: pxToPt(12),
    color: scenario.pillColor,
    bold: true,
  });

  const metricsY = yPx + paddingPx + 68;
  addMetrics(slide, pptx, xPx + paddingPx, metricsY, textWidthPx, scenario.metrics);

  const listStartY = metricsY + 112;
  addBulletList(
    slide,
    pptx,
    xPx + paddingPx,
    listStartY,
    textWidthPx,
    scenario.bullets
  );
};

const addAsk = (slide, pptx, askLayout) => {
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx, askLayout.yPx, askLayout.wPx, askLayout.hPx),
    fill: palette.deepNavy,
    line: { color: palette.deepNavy },
    rectRadius: 14,
    shadow: { type: "outer", opacity: 22, blur: 7, offset: 0.16, angle: 90 },
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx + 16, askLayout.yPx + 14, 48, 48),
    fill: { color: "FFFFFF", transparency: 78 },
    line: { color: "FFFFFF", transparency: 78 },
    rectRadius: 12,
  });

  slide.addText("★", {
    ...sharedText,
    ...asBox(askLayout.xPx + 30, askLayout.yPx + 20, 18, 28),
    fontSize: pxToPt(24),
    color: "FFFFFF",
    bold: true,
  });

  slide.addText("Our ask: Approve +3 FTE for 18 months", {
    ...sharedText,
    ...asBox(askLayout.xPx + 76, askLayout.yPx + 12, askLayout.wPx * 0.6, 28),
    fontSize: pxToPt(16),
    color: "FFFFFF",
    bold: true,
  });

  slide.addText(
    "Enables four concurrent tracks, accelerates policy pilots, and reduces operational risk while safeguarding BAU.",
    {
      ...sharedText,
      ...asBox(askLayout.xPx + 76, askLayout.yPx + 36, askLayout.wPx * 0.6, 28),
      fontSize: pxToPt(13),
      color: "d5e2f2",
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx + askLayout.wPx - 210, askLayout.yPx + 16, 194, 46),
    fill: palette.gold,
    line: { color: palette.gold },
    rectRadius: 12,
    shadow: { type: "outer", opacity: 32, blur: 7, offset: 0.2, angle: 90 },
  });

  slide.addText("Proceed with Preferred Plan", {
    ...sharedText,
    ...asBox(askLayout.xPx + askLayout.wPx - 190, askLayout.yPx + 22, 160, 28),
    fontSize: pxToPt(12),
    color: "1f1606",
    bold: true,
  });
};

const buildSlide = (pptx) => {
  const slide = pptx.addSlide();
  const layout = computeLayout();

  addShell(slide, layout, pptx);
  addHeader(slide, pptx, layout);
  addScenarioCard(
    slide,
    pptx,
    layout.content.xPx,
    layout.cards.leftWidthPx,
    scenarios[0],
    layout
  );
  addScenarioCard(
    slide,
    pptx,
    layout.content.xPx + layout.cards.leftWidthPx + DESIGN.columnGapPx,
    layout.cards.rightWidthPx,
    scenarios[1],
    layout
  );
  addAsk(slide, pptx, layout.ask);
};

const createDeck = () => {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  buildSlide(pptx);
  return pptx;
};

async function buildSlides(outputPath = path.join(__dirname, "..", "dist", "deck.pptx")) {
  const pptx = createDeck();
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  const nodeBuffer = await pptx.write("nodebuffer");
  fs.writeFileSync(outputPath, nodeBuffer);
  return outputPath;
}

if (require.main === module) {
  buildSlides()
    .then((output) => console.log(`Presentation created: ${output}`))
    .catch((err) => {
      console.error("Failed to build slides:", err);
      process.exit(1);
    });
}

module.exports = { buildSlides, createDeck };
