const fs = require("fs");
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const JSZip = require("jszip");

const SLIDE = { widthIn: 10, heightIn: 5.625 };
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
  metricGapPx: 12,
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
  white: "FFFFFF",
  cardBg: "FFFFFF",
  shellBg: "f6f9fe",
  pillBg: "fdf6e8",
  pillColor: "6f5316",
};

const sharedText = {
  fontFace: "Segoe UI",
  fontSize: 12,
  color: palette.text,
};

// Slide content from index.html
const slidesData = [
  {
    header: {
      eyebrow: "Central Bank AI Strategy",
      title: "From Pilot to Production",
      subtitle: "Trusted, resilient AI for policy, supervision, and market operations.",
      badge: "Strategic Priorities 2024–2026",
    },
    leftCard: {
      title: "Priority Use Cases",
      type: "iconGrid",
      items: [
        { icon: "◎", title: "Macro-financial intelligence", detail: "Stress testing · systemic risk · liquidity signals" },
        { icon: "◆", title: "Supervisory early warning", detail: "Outlier detection · conduct monitoring · alerts" },
        { icon: "◈", title: "Operational excellence", detail: "Secure copilots · reporting · translation" },
      ],
      sparkline: [
        { value: "Policy", label: "Foresight" },
        { value: "Supervision", label: "Precision" },
        { value: "Operations", label: "Resilience" },
      ],
    },
    rightCard: {
      title: "What Enables Scale",
      type: "pills",
      pills: ["Trusted Data", "Model Risk", "Secure MLOps"],
      items: [
        { icon: "◍", title: "Golden sources", detail: "Lineage and access aligned with confidentiality." },
        { icon: "✓", title: "Governance by design", detail: "Validation, explainability, and human oversight." },
        { icon: "⬡", title: "Production resilience", detail: "Monitoring, drift response, and auditability." },
      ],
    },
    ask: {
      icon: "✦",
      title: "Executive Priority",
      text: "Institutionalise governance and invest in the data backbone to scale AI with confidence.",
      cta: "Mandate",
    },
  },
  {
    header: {
      eyebrow: "Lifecycle & Governance",
      title: "Lifecycle Discipline + Democratized Innovation",
      subtitle: "A repeatable model for trusted deployment and enterprise adoption.",
      badge: "Policy-Compliant AI",
    },
    leftCard: {
      title: "End-to-End Lifecycle",
      type: "journey",
      journey: [
        { step: "01", title: "Qualify", label: "Strategic fit" },
        { step: "02", title: "Validate", label: "Model risk" },
        { step: "03", title: "Deploy", label: "Secure release" },
        { step: "04", title: "Monitor", label: "Audit ready" },
      ],
      sparkline: [
        { value: "Governed", label: "By design" },
        { value: "Secure", label: "Every release" },
        { value: "Measured", label: "Outcomes" },
      ],
    },
    rightCard: {
      title: "Democratize Innovation",
      type: "iconGrid",
      items: [
        { icon: "◎", title: "Federated squads", detail: "Domain teams co-create with the AI center of excellence." },
        { icon: "◆", title: "Shared platforms", detail: "Reusable data products, model libraries, and APIs." },
        { icon: "◈", title: "Capability uplift", detail: "Executive briefings, analyst academies, certified training." },
      ],
      sparkline: [
        { value: "30%", label: "Faster delivery" },
        { value: "Enterprise", label: "Adoption" },
        { value: "Audit", label: "Confidence" },
      ],
    },
    ask: {
      icon: "◆",
      title: "Strategic Close",
      text: "Align governance, talent, and platforms to industrialise AI with public trust.",
      cta: "Commit",
    },
  },
];

const clamp = (value, min, max) => {
  if (!Number.isFinite(value)) return min;
  if (typeof min === "number") value = Math.max(min, value);
  if (typeof max === "number") value = Math.min(max, value);
  return value;
};

const pxToIn = (px) => Number((clamp(px, 0, 5000) / PX_PER_IN).toFixed(4));
const pxToPt = (px) => Number(clamp((px * 72) / 96, 6, 72).toFixed(2));
const pxRadiusToIn = (px) => Number((clamp(px, 0, 200) / PX_PER_IN).toFixed(4));
const opacityToDecimal = (percent) => clamp(percent, 0, 100) / 100;

const asBox = (xPx, yPx, wPx, hPx) => ({
  x: pxToIn(xPx),
  y: pxToIn(yPx),
  w: pxToIn(Math.max(wPx, 4)),
  h: pxToIn(Math.max(hPx, 4)),
});

const computeLayout = (split = false) => {
  const shell = {
    xPx: DESIGN.shellMarginXPx,
    yPx: DESIGN.shellMarginYPx,
    wPx: DESIGN.widthPx - DESIGN.shellMarginXPx * 2,
    hPx: DESIGN.heightPx - DESIGN.shellMarginYPx * 2,
    radius: 26,
  };

  const content = {
    xPx: shell.xPx + DESIGN.paddingXPx,
    yPx: shell.yPx + DESIGN.paddingYPx,
    wPx: shell.wPx - DESIGN.paddingXPx * 2,
    hPx: shell.hPx - DESIGN.paddingYPx * 2,
  };

  const cardsHeightPx =
    content.hPx - DESIGN.headerHeightPx - DESIGN.askHeightPx - DESIGN.verticalGapPx * 2;

  let leftWidthPx, rightWidthPx;
  if (split) {
    leftWidthPx = (content.wPx - DESIGN.columnGapPx) / 2;
    rightWidthPx = leftWidthPx;
  } else {
    leftWidthPx = (content.wPx - DESIGN.columnGapPx) * (1.1 / 2);
    rightWidthPx = (content.wPx - DESIGN.columnGapPx) * (0.9 / 2);
  }

  return {
    shell,
    content,
    header: {
      eyebrowBox: asBox(content.xPx, content.yPx, content.wPx * 0.72, 20),
      titleBox: asBox(content.xPx, content.yPx + 22, content.wPx * 0.72, 40),
      subtitleBox: asBox(content.xPx, content.yPx + 66, content.wPx * 0.72, 24),
      badge: {
        ...asBox(content.xPx + content.wPx - 260, content.yPx + 6, 250, 48),
        radius: 14,
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
      yPx: content.yPx + DESIGN.headerHeightPx + DESIGN.verticalGapPx + cardsHeightPx + DESIGN.verticalGapPx,
      wPx: content.wPx,
      hPx: DESIGN.askHeightPx,
    },
  };
};

const addShell = (slide, layout, pptx) => {
  slide.background = { color: palette.deepNavy };

  // Add gradient overlay effect via a subtle shape
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(layout.shell.xPx, layout.shell.yPx, layout.shell.wPx, layout.shell.hPx),
    fill: { color: "f9fbff", type: "solid" },
    line: { color: "dce3ed", width: 1 },
    rectRadius: pxRadiusToIn(layout.shell.radius),
    shadow: { type: "outer", opacity: opacityToDecimal(24), blur: 12, offset: 0.3, angle: 90 },
  });
};

const addHeader = (slide, pptx, layout, headerData) => {
  // Eyebrow
  slide.addText(headerData.eyebrow.toUpperCase(), {
    ...sharedText,
    ...layout.header.eyebrowBox,
    fontSize: pxToPt(12),
    color: palette.muted,
    bold: true,
    charSpacing: 3.2,
  });

  // Title
  slide.addText(headerData.title, {
    ...sharedText,
    ...layout.header.titleBox,
    fontSize: pxToPt(34),
    color: palette.deepNavy,
    bold: true,
  });

  // Subtitle
  slide.addText(headerData.subtitle, {
    ...sharedText,
    ...layout.header.subtitleBox,
    fontSize: pxToPt(17),
    color: palette.muted,
  });

  // Badge background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...layout.header.badge,
    fill: palette.deepNavy,
    line: { color: palette.deepNavy },
    rectRadius: pxRadiusToIn(14),
    shadow: { type: "outer", opacity: opacityToDecimal(25), blur: 7, offset: 0.15, angle: 90 },
  });

  // Badge dot
  slide.addShape(pptx.ShapeType.ellipse, {
    x: layout.header.badge.x + pxToIn(14),
    y: layout.header.badge.y + pxToIn(18),
    w: pxToIn(12),
    h: pxToIn(12),
    fill: palette.gold,
    line: { color: palette.gold },
    shadow: { type: "outer", opacity: opacityToDecimal(18), blur: 4, offset: 0.02, angle: 90 },
  });

  // Badge text
  slide.addText(headerData.badge, {
    ...sharedText,
    x: layout.header.badge.x + pxToIn(34),
    y: layout.header.badge.y + pxToIn(12),
    w: layout.header.badge.w - pxToIn(44),
    h: pxToIn(24),
    fontSize: pxToPt(14),
    color: palette.white,
    bold: true,
  });
};

const addSparkline = (slide, pptx, xPx, yPx, widthPx, sparklineData) => {
  const gapPx = DESIGN.metricGapPx;
  const itemWidthPx = (widthPx - gapPx * 2) / 3;
  const itemHeightPx = 58;

  sparklineData.forEach((item, idx) => {
    const itemX = xPx + idx * (itemWidthPx + gapPx);

    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(itemX, yPx, itemWidthPx, itemHeightPx),
      fill: palette.softGray,
      line: { color: "e1e7ee", width: 1 },
      rectRadius: pxRadiusToIn(16),
    });

    slide.addText(item.value, {
      ...sharedText,
      ...asBox(itemX + 8, yPx + 8, itemWidthPx - 16, 24),
      fontSize: pxToPt(20),
      color: palette.deepNavy,
      bold: true,
      align: "center",
    });

    slide.addText(item.label.toUpperCase(), {
      ...sharedText,
      ...asBox(itemX + 8, yPx + 34, itemWidthPx - 16, 18),
      fontSize: pxToPt(12),
      color: palette.muted,
      charSpacing: 0.6,
      align: "center",
    });
  });
};

const addIconGrid = (slide, pptx, xPx, yPx, widthPx, items) => {
  let currentY = yPx;
  const lineHeightPx = 68;
  const iconSize = 54;

  items.forEach((item) => {
    // Icon circle
    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(xPx, currentY, iconSize, iconSize),
      fill: { color: palette.midNavy, type: "solid" },
      line: { color: palette.midNavy },
      rectRadius: pxRadiusToIn(16),
      shadow: { type: "outer", opacity: opacityToDecimal(30), blur: 6, offset: 0.1, angle: 90 },
    });

    // Icon text
    slide.addText(item.icon, {
      ...sharedText,
      ...asBox(xPx, currentY + 12, iconSize, 30),
      fontSize: pxToPt(22),
      color: palette.white,
      align: "center",
    });

    // Title
    slide.addText(item.title, {
      ...sharedText,
      ...asBox(xPx + iconSize + 14, currentY + 4, widthPx - iconSize - 14, 22),
      fontSize: pxToPt(16),
      color: palette.deepNavy,
      bold: true,
    });

    // Detail
    slide.addText(item.detail, {
      ...sharedText,
      ...asBox(xPx + iconSize + 14, currentY + 26, widthPx - iconSize - 14, 38),
      fontSize: pxToPt(13),
      color: palette.muted,
    });

    currentY += lineHeightPx;
  });

  return currentY;
};

const addPillRow = (slide, pptx, xPx, yPx, widthPx, pills) => {
  const gapPx = 12;
  const pillWidthPx = (widthPx - gapPx * 2) / 3;
  const pillHeightPx = 32;

  pills.forEach((pill, idx) => {
    const pillX = xPx + idx * (pillWidthPx + gapPx);

    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(pillX, yPx, pillWidthPx, pillHeightPx),
      fill: palette.pillBg,
      line: { color: palette.pillBg },
      rectRadius: pxRadiusToIn(999),
    });

    slide.addText(pill.toUpperCase(), {
      ...sharedText,
      ...asBox(pillX + 8, yPx + 6, pillWidthPx - 16, 20),
      fontSize: pxToPt(12),
      color: palette.pillColor,
      bold: true,
      charSpacing: 0.4,
      align: "center",
    });
  });
};

const addJourney = (slide, pptx, xPx, yPx, widthPx, journey) => {
  const gapPx = 14;
  const stepWidthPx = (widthPx - gapPx * 3) / 4;
  const stepHeightPx = 90;

  journey.forEach((item, idx) => {
    const stepX = xPx + idx * (stepWidthPx + gapPx);

    // Step background
    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(stepX, yPx, stepWidthPx, stepHeightPx),
      fill: "f7f9fc",
      line: { color: "e1e7ee", width: 1 },
      rectRadius: pxRadiusToIn(16),
    });

    // Step icon
    slide.addShape(pptx.ShapeType.roundRect, {
      ...asBox(stepX + (stepWidthPx - 44) / 2, yPx + 10, 44, 44),
      fill: "fdf6e8",
      line: { color: "fdf6e8" },
      rectRadius: pxRadiusToIn(14),
    });

    slide.addText(item.step, {
      ...sharedText,
      ...asBox(stepX + (stepWidthPx - 44) / 2, yPx + 20, 44, 24),
      fontSize: pxToPt(14),
      color: palette.pillColor,
      bold: true,
      align: "center",
    });

    // Step title
    slide.addText(item.title, {
      ...sharedText,
      ...asBox(stepX + 4, yPx + 56, stepWidthPx - 8, 16),
      fontSize: pxToPt(13),
      color: palette.deepNavy,
      bold: true,
      align: "center",
    });

    // Step label
    slide.addText(item.label, {
      ...sharedText,
      ...asBox(stepX + 4, yPx + 72, stepWidthPx - 8, 14),
      fontSize: pxToPt(12),
      color: palette.muted,
      align: "center",
    });
  });
};

const addCard = (slide, pptx, xPx, yPx, widthPx, heightPx, cardData) => {
  // Card background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(xPx, yPx, widthPx, heightPx),
    fill: palette.cardBg,
    line: { color: "dde5ef", width: 1 },
    rectRadius: pxRadiusToIn(18),
    shadow: { type: "outer", opacity: opacityToDecimal(16), blur: 8, offset: 0.15, angle: 90 },
  });

  const paddingPx = DESIGN.cardPaddingPx;
  const innerWidthPx = widthPx - paddingPx * 2;
  let currentY = yPx + paddingPx;

  // Card title
  slide.addText(cardData.title.toUpperCase(), {
    ...sharedText,
    ...asBox(xPx + paddingPx, currentY, innerWidthPx, 22),
    fontSize: pxToPt(18),
    color: palette.deepNavy,
    bold: true,
    charSpacing: 0.3,
  });
  currentY += 28;

  // Handle different card types
  if (cardData.type === "pills" && cardData.pills) {
    addPillRow(slide, pptx, xPx + paddingPx, currentY, innerWidthPx, cardData.pills);
    currentY += 46;
  }

  if (cardData.type === "journey" && cardData.journey) {
    addJourney(slide, pptx, xPx + paddingPx, currentY, innerWidthPx, cardData.journey);
    currentY += 104;
  }

  if (cardData.items && (cardData.type === "iconGrid" || cardData.type === "pills")) {
    const endY = addIconGrid(slide, pptx, xPx + paddingPx, currentY, innerWidthPx, cardData.items);
    currentY = endY;
  }

  // Add sparkline at the bottom if present
  if (cardData.sparkline) {
    const sparklineY = yPx + heightPx - paddingPx - 66;
    addSparkline(slide, pptx, xPx + paddingPx, sparklineY, innerWidthPx, cardData.sparkline);
  }
};

const addAsk = (slide, pptx, askLayout, askData) => {
  // Ask background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx, askLayout.yPx, askLayout.wPx, askLayout.hPx),
    fill: { color: palette.deepNavy, type: "solid" },
    line: { color: palette.deepNavy },
    rectRadius: pxRadiusToIn(18),
    shadow: { type: "outer", opacity: opacityToDecimal(22), blur: 8, offset: 0.15, angle: 90 },
  });

  // Icon background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx + 18, askLayout.yPx + 14, 48, 48),
    fill: { color: palette.white, transparency: 88 },
    line: { color: palette.white, transparency: 88 },
    rectRadius: pxRadiusToIn(14),
  });

  // Icon
  slide.addText(askData.icon, {
    ...sharedText,
    ...asBox(askLayout.xPx + 18, askLayout.yPx + 20, 48, 36),
    fontSize: pxToPt(22),
    color: palette.white,
    align: "center",
  });

  // Title
  slide.addText(askData.title, {
    ...sharedText,
    ...asBox(askLayout.xPx + 80, askLayout.yPx + 12, askLayout.wPx * 0.6, 24),
    fontSize: pxToPt(18),
    color: palette.white,
    bold: true,
  });

  // Text
  slide.addText(askData.text, {
    ...sharedText,
    ...asBox(askLayout.xPx + 80, askLayout.yPx + 38, askLayout.wPx * 0.6, 24),
    fontSize: pxToPt(14),
    color: "d5e2f2",
  });

  // CTA button background
  slide.addShape(pptx.ShapeType.roundRect, {
    ...asBox(askLayout.xPx + askLayout.wPx - 130, askLayout.yPx + 18, 110, 40),
    fill: palette.gold,
    line: { color: palette.gold },
    rectRadius: pxRadiusToIn(12),
    shadow: { type: "outer", opacity: opacityToDecimal(35), blur: 6, offset: 0.15, angle: 90 },
  });

  // CTA text
  slide.addText(askData.cta.toUpperCase(), {
    ...sharedText,
    ...asBox(askLayout.xPx + askLayout.wPx - 125, askLayout.yPx + 24, 100, 28),
    fontSize: pxToPt(12),
    color: "1f1606",
    bold: true,
    charSpacing: 0.4,
    align: "center",
  });
};

const buildSlide = (pptx, slideData, isSplit = false) => {
  const slide = pptx.addSlide();
  const layout = computeLayout(isSplit);

  addShell(slide, layout, pptx);
  addHeader(slide, pptx, layout, slideData.header);

  // Left card
  addCard(
    slide,
    pptx,
    layout.content.xPx,
    layout.cards.yPx,
    layout.cards.leftWidthPx,
    layout.cards.heightPx,
    slideData.leftCard
  );

  // Right card
  addCard(
    slide,
    pptx,
    layout.content.xPx + layout.cards.leftWidthPx + DESIGN.columnGapPx,
    layout.cards.yPx,
    layout.cards.rightWidthPx,
    layout.cards.heightPx,
    slideData.rightCard
  );

  addAsk(slide, pptx, layout.ask, slideData.ask);
};

const createDeck = () => {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = "From Pilot to Production: Industrialising AI in a Central Bank";
  pptx.author = "Central Bank AI Strategy";

  // Build both slides from HTML content
  buildSlide(pptx, slidesData[0], false);
  buildSlide(pptx, slidesData[1], true);

  return pptx;
};

const normalizeGroupExtents = (xml, cx, cy) => {
  const replacements = [
    { pattern: /<a:ext\s+cx="0"\s+cy="0"\s*\/>/g, value: `<a:ext cx="${cx}" cy="${cy}"/>` },
    { pattern: /<a:chExt\s+cx="0"\s+cy="0"\s*\/>/g, value: `<a:chExt cx="${cx}" cy="${cy}"/>` },
  ];

  return replacements.reduce((current, { pattern, value }) => current.replace(pattern, value), xml);
};

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

async function buildSlides(
  outputPath = path.join(__dirname, "..", "artifacts", "slides", "latest.pptx")
) {
  const pptx = createDeck();
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await writeNormalizedPptx(pptx, outputPath);
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
