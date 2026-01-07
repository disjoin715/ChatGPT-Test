const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const JSZip = require("jszip");

// Slide dimensions
const SLIDE = { widthIn: 10, heightIn: 5.625 };
const PX_PER_IN = 128; // 1280px / 10in

// Design constants (in pixels, matching HTML)
const DESIGN = {
  widthPx: 1280,
  heightPx: 720,
  shellMarginXPx: 28,
  shellMarginYPx: 24,
  paddingXPx: 44,
  paddingYPx: 40,
  columnGapPx: 20,
  verticalGapPx: 26,
  headerHeightPx: 100,
  askHeightPx: 76,
  cardPaddingPx: 24,
};

const palette = {
  deepNavy: "0F2439",
  midNavy: "17395c",
  sky: "7fb4e0",
  gold: "e1b44c",
  softGray: "f4f6f8",
  text: "0c1b2a",
  muted: "4a5c70",
  white: "FFFFFF",
  cardBorder: "dde5ef",
};

const scenarios = [
  {
    title: "Baseline Resourcing",
    pill: "4 FTE · BAU + Incremental Change",
    pillFill: "f2f6fb",
    pillColor: palette.midNavy,
    metrics: [
      { label: "CAPACITY", value: "~60%" },
      { label: "CHANGE BANDWIDTH", value: "2 tracks" },
      { label: "RISK EXPOSURE", value: "Moderate" },
    ],
    bullets: [
      {
        title: "Maintain core operations",
        detail: "Focus on stability and regulatory reporting; limited innovation bandwidth.",
      },
      {
        title: "Sequential delivery",
        detail: "Two initiative tracks in sequence; elongates policy pilots and analytics upgrades.",
      },
      {
        title: "Deferred optimization",
        detail: "Automation and resilience improvements shift beyond the 18-month window.",
      },
    ],
  },
  {
    title: "More Resource (Preferred)",
    pill: "+3 FTE · Data, Ops, Delivery",
    pillFill: "fdf6e8",
    pillColor: "6f5316",
    metrics: [
      { label: "CAPACITY", value: "~90%" },
      { label: "CHANGE BANDWIDTH", value: "4 tracks" },
      { label: "RISK EXPOSURE", value: "Lower" },
    ],
    bullets: [
      {
        title: "Parallel delivery",
        detail: "Run concurrent workstreams (policy pilots, data modernization, resiliency uplift) without sacrificing BAU.",
      },
      {
        title: "Faster regulatory readiness",
        detail: "Expedite compliance changes and scenario testing with embedded risk & controls support.",
      },
      {
        title: "Data & automation gains",
        detail: "Deliver prioritized automations, observability, and analytics that reduce manual effort and downtime.",
      },
    ],
  },
];

// Helper functions
const pxToIn = (px) => px / PX_PER_IN;
const pxToPt = (px) => (px * 72) / 96;

const asBox = (xPx, yPx, wPx, hPx) => ({
  x: pxToIn(xPx),
  y: pxToIn(yPx),
  w: pxToIn(Math.max(wPx, 1)),
  h: pxToIn(Math.max(hPx, 1)),
});

// Compute layout positions
const computeLayout = () => {
  const shell = {
    xPx: DESIGN.shellMarginXPx,
    yPx: DESIGN.shellMarginYPx,
    wPx: DESIGN.widthPx - DESIGN.shellMarginXPx * 2,
    hPx: DESIGN.heightPx - DESIGN.shellMarginYPx * 2,
  };

  const content = {
    xPx: shell.xPx + DESIGN.paddingXPx,
    yPx: shell.yPx + DESIGN.paddingYPx,
    wPx: shell.wPx - DESIGN.paddingXPx * 2,
    hPx: shell.hPx - DESIGN.paddingYPx * 2,
  };

  // Cards take remaining space after header and ask bar
  const cardsHeightPx = content.hPx - DESIGN.headerHeightPx - DESIGN.askHeightPx - DESIGN.verticalGapPx * 2;

  // Equal width cards
  const cardWidthPx = (content.wPx - DESIGN.columnGapPx) / 2;

  return {
    shell,
    content,
    header: {
      titleBox: asBox(content.xPx, content.yPx, content.wPx * 0.72, 42),
      subtitleBox: asBox(content.xPx, content.yPx + 48, content.wPx * 0.7, 24),
      badgeBox: asBox(content.xPx + content.wPx - 180, content.yPx, 180, 42),
    },
    cards: {
      yPx: content.yPx + DESIGN.headerHeightPx,
      heightPx: cardsHeightPx,
      widthPx: cardWidthPx,
      leftXPx: content.xPx,
      rightXPx: content.xPx + cardWidthPx + DESIGN.columnGapPx,
    },
    ask: {
      xPx: content.xPx,
      yPx: content.yPx + DESIGN.headerHeightPx + cardsHeightPx + DESIGN.verticalGapPx,
      wPx: content.wPx,
      hPx: DESIGN.askHeightPx,
    },
  };
};

// Add the slide background shell
const addShell = (slide, layout, pptx) => {
  slide.background = { color: palette.deepNavy };

  // Main rounded rectangle (slide background)
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(layout.shell.xPx, layout.shell.yPx, layout.shell.wPx, layout.shell.hPx),
    fill: { color: "f9fbff" },
    line: { color: "dce3ed", pt: 1 },
    rectRadius: 0.2,
  });
};

// Add header section
const addHeader = (slide, pptx, layout) => {
  // Title
  slide.addText("Baseline vs. Accelerated Delivery", {
    ...layout.header.titleBox,
    fontFace: "Segoe UI",
    fontSize: 34,
    color: palette.deepNavy,
    bold: true,
    valign: "top",
  });

  // Subtitle
  slide.addText("How resourcing choices shape delivery outcomes and policy readiness.", {
    ...layout.header.subtitleBox,
    fontFace: "Segoe UI",
    fontSize: 17,
    color: palette.muted,
    valign: "top",
  });

  // Badge background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...layout.header.badgeBox,
    fill: { color: palette.deepNavy },
    rectRadius: 0.15,
  });

  // Badge dot
  slide.addShape(pptx.ShapeType.ellipse, {
    x: layout.header.badgeBox.x + 0.1,
    y: layout.header.badgeBox.y + 0.12,
    w: 0.1,
    h: 0.1,
    fill: { color: palette.gold },
  });

  // Badge text
  slide.addText("Q4 FY24 → Q1 FY26", {
    x: layout.header.badgeBox.x + 0.25,
    y: layout.header.badgeBox.y,
    w: layout.header.badgeBox.w - 0.3,
    h: layout.header.badgeBox.h,
    fontFace: "Segoe UI",
    fontSize: 13,
    color: palette.white,
    bold: true,
    valign: "middle",
  });
};

// Add a scenario card
const addScenarioCard = (slide, pptx, xPx, yPx, widthPx, heightPx, scenario) => {
  const padding = DESIGN.cardPaddingPx;
  const innerWidth = widthPx - padding * 2;

  // Card background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(xPx, yPx, widthPx, heightPx),
    fill: { color: palette.white },
    line: { color: palette.cardBorder, pt: 1 },
    rectRadius: 0.15,
  });

  let currentY = yPx + padding;

  // Card title
  slide.addText(scenario.title, {
    ...asBox(xPx + padding, currentY, innerWidth, 24),
    fontFace: "Segoe UI",
    fontSize: 18,
    color: palette.deepNavy,
    bold: true,
    valign: "top",
  });
  currentY += 28;

  // Pill
  const pillWidth = 200;
  const pillHeight = 28;
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(xPx + padding, currentY, pillWidth, pillHeight),
    fill: { color: scenario.pillFill },
    rectRadius: 0.3,
  });

  slide.addText(scenario.pill, {
    ...asBox(xPx + padding + 10, currentY, pillWidth - 20, pillHeight),
    fontFace: "Segoe UI",
    fontSize: 11,
    color: scenario.pillColor,
    bold: true,
    valign: "middle",
  });
  currentY += pillHeight + 16;

  // Metrics row
  const metricGap = 12;
  const metricWidth = (innerWidth - metricGap * 2) / 3;
  const metricHeight = 80;

  scenario.metrics.forEach((metric, idx) => {
    const metricX = xPx + padding + idx * (metricWidth + metricGap);

    // Metric box background
    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(metricX, currentY, metricWidth, metricHeight),
      fill: { color: palette.softGray },
      line: { color: "e1e7ee", pt: 1 },
      rectRadius: 0.12,
    });

    // Label at top (smaller, uppercase)
    slide.addText(metric.label, {
      ...asBox(metricX + 8, currentY + 8, metricWidth - 16, 16),
      fontFace: "Segoe UI",
      fontSize: 9,
      color: palette.muted,
      bold: true,
      valign: "top",
    });

    // Value (large, centered)
    slide.addText(metric.value, {
      ...asBox(metricX + 8, currentY + 28, metricWidth - 16, 44),
      fontFace: "Segoe UI",
      fontSize: 22,
      color: palette.deepNavy,
      bold: true,
      valign: "middle",
    });
  });
  currentY += metricHeight + 18;

  // Bullet points
  const bulletSpacing = 54;

  scenario.bullets.forEach((item) => {
    // Bullet dot
    slide.addShape(pptx.ShapeType.ellipse, {
      ...asBox(xPx + padding, currentY + 5, 10, 10),
      fill: { color: palette.gold },
    });

    // Bullet title
    slide.addText(item.title, {
      ...asBox(xPx + padding + 18, currentY, innerWidth - 18, 18),
      fontFace: "Segoe UI",
      fontSize: 13,
      color: palette.deepNavy,
      bold: true,
      valign: "top",
    });

    // Bullet detail
    slide.addText(item.detail, {
      ...asBox(xPx + padding + 18, currentY + 18, innerWidth - 18, 34),
      fontFace: "Segoe UI",
      fontSize: 11,
      color: palette.muted,
      valign: "top",
      wrap: true,
    });

    currentY += bulletSpacing;
  });
};

// Add the bottom "ask" bar
const addAsk = (slide, pptx, askLayout) => {
  // Ask bar background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx, askLayout.yPx, askLayout.wPx, askLayout.hPx),
    fill: { color: palette.deepNavy },
    rectRadius: 0.15,
  });

  // Star icon container
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx + 16, askLayout.yPx + 14, 48, 48),
    fill: { color: "2a4a6b" },
    rectRadius: 0.12,
  });

  // Star character
  slide.addText("★", {
    ...asBox(askLayout.xPx + 16, askLayout.yPx + 14, 48, 48),
    fontFace: "Segoe UI",
    fontSize: 22,
    color: palette.white,
    align: "center",
    valign: "middle",
  });

  // Ask title
  slide.addText("Our ask: Approve +3 FTE for 18 months", {
    ...asBox(askLayout.xPx + 80, askLayout.yPx + 14, askLayout.wPx * 0.55, 22),
    fontFace: "Segoe UI",
    fontSize: 15,
    color: palette.white,
    bold: true,
    valign: "top",
  });

  // Ask subtitle
  slide.addText(
    "Enables four concurrent tracks, accelerates policy pilots, and reduces operational risk while safeguarding BAU.",
    {
      ...asBox(askLayout.xPx + 80, askLayout.yPx + 38, askLayout.wPx * 0.55, 28),
      fontFace: "Segoe UI",
      fontSize: 11,
      color: "d5e2f2",
      valign: "top",
      wrap: true,
    }
  );

  // CTA button
  const ctaWidth = 160;
  const ctaHeight = 40;
  const ctaX = askLayout.xPx + askLayout.wPx - ctaWidth - 20;
  const ctaY = askLayout.yPx + (askLayout.hPx - ctaHeight) / 2;

  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(ctaX, ctaY, ctaWidth, ctaHeight),
    fill: { color: palette.gold },
    rectRadius: 0.12,
  });

  slide.addText("Proceed with\nPreferred Plan", {
    ...asBox(ctaX, ctaY, ctaWidth, ctaHeight),
    fontFace: "Segoe UI",
    fontSize: 11,
    color: "1f1606",
    bold: true,
    align: "center",
    valign: "middle",
  });
};

// Build the main slide
const buildSlide = (pptx) => {
  const slide = pptx.addSlide();
  const layout = computeLayout();

  addShell(slide, layout, pptx);
  addHeader(slide, pptx, layout);

  // Left card
  addScenarioCard(
    slide,
    pptx,
    layout.cards.leftXPx,
    layout.cards.yPx,
    layout.cards.widthPx,
    layout.cards.heightPx,
    scenarios[0]
  );

  // Right card
  addScenarioCard(
    slide,
    pptx,
    layout.cards.rightXPx,
    layout.cards.yPx,
    layout.cards.widthPx,
    layout.cards.heightPx,
    scenarios[1]
  );

  addAsk(slide, pptx, layout.ask);
};

// Create the presentation deck
const createDeck = () => {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = "Baseline vs Accelerated Delivery";
  buildSlide(pptx);
  return pptx;
};

// Normalize group extents in PPTX XML (fixes rendering issues)
const normalizeGroupExtents = (xml, cx, cy) => {
  return xml
    .replace(/<a:ext\s+cx="0"\s+cy="0"\s*\/>/g, `<a:ext cx="${cx}" cy="${cy}"/>`)
    .replace(/<a:chExt\s+cx="0"\s+cy="0"\s*\/>/g, `<a:chExt cx="${cx}" cy="${cy}"/>`);
};

// Write the PPTX with normalized XML
const writeNormalizedPptx = async (pptx, outputPath) => {
  const nodeBuffer = await pptx.write("nodebuffer");
  const zip = await JSZip.loadAsync(nodeBuffer);

  const slideCx = Math.round(SLIDE.widthIn * 914400);
  const slideCy = Math.round(SLIDE.heightIn * 914400);

  const targets = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/") && name.endsWith(".xml")
  );

  for (const name of targets) {
    const originalXml = await zip.file(name).async("string");
    const normalizedXml = normalizeGroupExtents(originalXml, slideCx, slideCy);
    if (normalizedXml !== originalXml) {
      zip.file(name, normalizedXml);
    }
  }

  const normalizedBuffer = await zip.generateAsync({ type: "nodebuffer" });
  fs.writeFileSync(outputPath, normalizedBuffer);
};

// Main build function
async function buildSlides(outputPath = path.join(__dirname, "..", "dist", "deck.pptx")) {
  const pptx = createDeck();
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await writeNormalizedPptx(pptx, outputPath);
  return outputPath;
}

// Run if called directly
if (require.main === module) {
  buildSlides()
    .then((output) => console.log(`Presentation created: ${output}`))
    .catch((err) => {
      console.error("Failed to build slides:", err);
      process.exit(1);
    });
}

module.exports = { buildSlides, createDeck };
