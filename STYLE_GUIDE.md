# PowerPoint Style Guide

> **For AI Coding Agents:** This document defines the visual design preferences for generated PowerPoint slides. Always reference this file when creating or modifying slides. Changes to this guide should be reflected in `build-slides.js`.

---

## Table of Contents

1. [Color Palette](#color-palette)
2. [Typography](#typography)
3. [Spacing & Layout](#spacing--layout)
4. [Component Styles](#component-styles)
5. [Shadow & Effects](#shadow--effects)
6. [Icons & Symbols](#icons--symbols)
7. [Content Guidelines](#content-guidelines)
8. [Customization](#customization)

---

## Color Palette

### Primary Colors

| Name | Hex Code | Usage |
|------|----------|-------|
| **Deep Navy** | `#0F2439` | Background, headers, CTA bar, badge |
| **Mid Navy** | `#17395C` | Icon backgrounds, secondary accents |
| **Accent Gold** | `#E1B44C` | CTA buttons, badge dots, bullet points |

### Neutral Colors

| Name | Hex Code | Usage |
|------|----------|-------|
| **Text** | `#0C1B2A` | Primary body text, card titles |
| **Muted** | `#4A5C70` | Subtitles, descriptions, labels |
| **Soft Gray** | `#F4F6F8` | Sparkline backgrounds, metric boxes |
| **Card Background** | `#FFFFFF` | Card fills |
| **Shell Background** | `#F6F9FE` | Main slide content area |
| **Border** | `#DDE5EF` | Card borders, dividers |

### Pill Colors

| Name | Hex Code | Usage |
|------|----------|-------|
| **Pill Background** | `#FDF6E8` | Tag/pill fill (warm cream) |
| **Pill Text** | `#6F5316` | Tag/pill text (dark gold) |

### Color Code for PptxGenJS

Remove the `#` prefix when using in JavaScript:

```javascript
const palette = {
  deepNavy: "0F2439",
  midNavy: "17395c",
  gold: "e1b44c",
  text: "0c1b2a",
  muted: "4a5c70",
  softGray: "f4f6f8",
  white: "FFFFFF",
  cardBg: "FFFFFF",
  shellBg: "f6f9fe",
  pillBg: "fdf6e8",
  pillColor: "6f5316",
};
```

---

## Typography

### Font Family

**Primary Font:** Segoe UI

> Chosen for Windows compatibility and consistent rendering in PowerPoint.

### Font Sizes (in CSS pixels, converted to points)

| Element | Size (px) | Size (pt) | Weight | Spacing |
|---------|-----------|-----------|--------|---------|
| Eyebrow | 12px | 9pt | Bold | 3.2 (wide) |
| Title (H1) | 34px | 25.5pt | Bold | Normal |
| Subtitle | 17px | 12.75pt | Normal | Normal |
| Badge Text | 14px | 10.5pt | Bold | Normal |
| Card Title | 18px | 13.5pt | Bold | 0.3 |
| Item Title | 16px | 12pt | Bold | Normal |
| Item Detail | 13px | 9.75pt | Normal | Normal |
| Sparkline Value | 20px | 15pt | Bold | Normal |
| Sparkline Label | 12px | 9pt | Normal | 0.6 |
| Pill Text | 12px | 9pt | Bold | 0.4 |
| CTA Title | 18px | 13.5pt | Bold | Normal |
| CTA Text | 14px | 10.5pt | Normal | Normal |
| CTA Button | 12px | 9pt | Bold | 0.4 |

### Text Colors by Context

| Context | Color |
|---------|-------|
| Main titles | Deep Navy |
| Subtitles/descriptions | Muted |
| Card titles | Deep Navy |
| On dark backgrounds | White |
| On CTA bar | White (title), Light blue `#D5E2F2` (description) |
| Pills | Pill Text (`#6F5316`) |

### Text Transformations

- **Eyebrow**: UPPERCASE
- **Card titles**: UPPERCASE
- **Pill text**: UPPERCASE
- **CTA button**: UPPERCASE
- **Sparkline labels**: UPPERCASE

---

## Spacing & Layout

### Slide Dimensions

| Property | Value |
|----------|-------|
| Design Width | 1280px |
| Design Height | 720px |
| Output Width | 10 inches |
| Output Height | 5.625 inches |
| Aspect Ratio | 16:9 |

### Conversion Formulas

```javascript
// Pixels to Inches
const pxToIn = (px) => px / 128;

// Pixels to Points (for font sizes)
const pxToPt = (px) => px * 72 / 96;
```

### Margins & Padding

| Property | Value (px) |
|----------|------------|
| Shell Margin X | 28px |
| Shell Margin Y | 24px |
| Content Padding X | 44px |
| Content Padding Y | 40px |
| Column Gap | 20px |
| Vertical Gap | 26px |
| Card Padding | 24px |
| Metric Gap | 12px |

### Region Heights

| Region | Value (px) |
|--------|------------|
| Header Height | 100px |
| CTA Bar Height | 76px |
| Cards | Calculated (remaining space) |

### Column Widths

| Layout | Left Card | Right Card |
|--------|-----------|------------|
| Default | 55% (1.1fr) | 45% (0.9fr) |
| Split | 50% (1fr) | 50% (1fr) |

---

## Component Styles

### Shell (Main Content Area)

```javascript
{
  fill: "f6f9fe",
  line: { color: "dce3ed", width: 1 },
  rectRadius: 26px → 0.203in,
  shadow: { opacity: 24%, blur: 12, offset: 0.3in }
}
```

### Cards

```javascript
{
  fill: "FFFFFF",
  line: { color: "dde5ef", width: 1 },
  rectRadius: 18px → 0.141in,
  shadow: { opacity: 16%, blur: 8, offset: 0.15in }
}
```

### Badge

```javascript
{
  fill: "0F2439" (deep navy),
  rectRadius: 14px → 0.109in,
  shadow: { opacity: 25%, blur: 7, offset: 0.15in }
}
// With gold dot indicator (12px circle)
```

### Icon Circles

```javascript
{
  size: 54px × 54px,
  fill: "17395c" (mid navy),
  rectRadius: 16px → 0.125in,
  shadow: { opacity: 30%, blur: 6, offset: 0.1in }
}
```

### Sparkline Boxes

```javascript
{
  height: 58px,
  fill: "f4f6f8" (soft gray),
  line: { color: "e1e7ee", width: 1 },
  rectRadius: 16px → 0.125in
}
```

### Pills

```javascript
{
  height: 32px,
  fill: "fdf6e8" (warm cream),
  rectRadius: 999px (fully rounded)
}
```

### Journey Steps

```javascript
{
  height: 90px,
  background: "f7f9fc",
  line: { color: "e1e7ee", width: 1 },
  rectRadius: 16px → 0.125in,
  // Inner icon box: 44px, fill "fdf6e8", radius 14px
}
```

### CTA Bar

```javascript
{
  height: 76px,
  fill: "0F2439" (deep navy),
  rectRadius: 18px → 0.141in,
  shadow: { opacity: 22%, blur: 8, offset: 0.15in }
}
```

### CTA Button

```javascript
{
  size: 110px × 40px,
  fill: "e1b44c" (gold),
  rectRadius: 12px → 0.094in,
  shadow: { opacity: 35%, blur: 6, offset: 0.15in }
}
```

---

## Shadow & Effects

### Standard Shadow Configuration

All shadows use:
- **Type**: Outer
- **Angle**: 90° (downward)

| Component | Opacity | Blur | Offset |
|-----------|---------|------|--------|
| Shell | 24% | 12 | 0.3in |
| Cards | 16% | 8 | 0.15in |
| Badge | 25% | 7 | 0.15in |
| Icon Circles | 30% | 6 | 0.1in |
| CTA Bar | 22% | 8 | 0.15in |
| CTA Button | 35% | 6 | 0.15in |
| Gold Elements | 18-35% | 4-6 | 0.02-0.15in |

### Gold Dot Glow Effect

The badge dot uses a subtle glow:
```javascript
shadow: { opacity: 18%, blur: 4, offset: 0.02in }
```

---

## Icons & Symbols

### Recommended Unicode Symbols

| Symbol | Unicode | Usage |
|--------|---------|-------|
| ◎ | U+25CE | Primary/first items |
| ◆ | U+25C6 | Secondary items |
| ◈ | U+25C8 | Tertiary items |
| ◍ | U+25CD | Data/sources |
| ✓ | U+2713 | Governance/validation |
| ⬡ | U+2B21 | Infrastructure/systems |
| ✦ | U+2726 | Featured/priority |
| ★ | U+2605 | Highlights |

### Icon Sizing

| Context | Font Size |
|---------|-----------|
| Icon in circle | 22px (16.5pt) |
| CTA icon | 22px (16.5pt) |
| Step numbers | 14px (10.5pt) |

---

## Content Guidelines

### Text Length Recommendations

| Element | Max Characters |
|---------|----------------|
| Eyebrow | 25 |
| Title | 50 |
| Subtitle | 80 |
| Badge | 30 |
| Card Title | 25 |
| Item Title | 30 |
| Item Detail | 60 |
| CTA Title | 25 |
| CTA Text | 100 |
| CTA Button | 15 |

### Structure Guidelines

1. **Each slide should have:**
   - Clear eyebrow category
   - Concise main title
   - Supporting subtitle
   - Two balanced content cards
   - Single call-to-action

2. **Cards should contain:**
   - 2-4 main content items
   - Optional sparkline (3 metrics max)
   - Optional pill row (3 items max)

3. **Journey steps:**
   - 3-5 steps optimal
   - Short titles (1-2 words)
   - Brief subtitles

---

## Customization

### Changing Colors

To modify the color palette:

1. Update the `palette` object in `build-slides.js`
2. Update this style guide to reflect changes
3. Rebuild slides with `npm run build:slides`

### Changing Fonts

To use a different font:

1. Update `sharedText.fontFace` in `build-slides.js`
2. Ensure the font is available on target systems
3. Test in PowerPoint for consistent rendering

### Changing Spacing

Modify the `DESIGN` object in `build-slides.js`:

```javascript
const DESIGN = {
  widthPx: 1280,
  heightPx: 720,
  shellMarginXPx: 28,    // Outer margin
  shellMarginYPx: 24,
  paddingXPx: 44,        // Inner padding
  paddingYPx: 40,
  columnGapPx: 20,       // Between cards
  verticalGapPx: 26,     // Between sections
  headerHeightPx: 100,   // Header region
  askHeightPx: 76,       // CTA region
  cardPaddingPx: 24,     // Inside cards
  metricGapPx: 12,       // Between metrics
};
```

### Adding New Component Types

1. Define the HTML structure in `index.html`
2. Add to the slide data structure
3. Create `add{ComponentName}()` function in `build-slides.js`
4. Document in this style guide

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-07 | Initial style guide |

---

## Quick Reference Card

```
┌─────────────────────────────────────────────────────────────┐
│  COLORS                                                      │
│  ───────                                                     │
│  Primary:   #0F2439 (navy)  #17395C (mid)  #E1B44C (gold)   │
│  Text:      #0C1B2A (dark)  #4A5C70 (muted)                 │
│  Background: #F6F9FE (shell) #FFFFFF (cards) #F4F6F8 (gray) │
│                                                              │
│  TYPOGRAPHY                                                  │
│  ──────────                                                  │
│  Font: Segoe UI                                              │
│  Title: 34px bold    Subtitle: 17px normal                  │
│  Card Title: 18px    Body: 13-16px                          │
│                                                              │
│  SPACING                                                     │
│  ───────                                                     │
│  Canvas: 1280×720px → 10×5.625in                            │
│  Margins: 28×24px   Padding: 44×40px                        │
│  Gaps: 20px (cols)  26px (rows)                             │
│                                                              │
│  CORNERS                                                     │
│  ───────                                                     │
│  Shell: 26px  Cards: 18px  Badges: 14px  Icons: 16px        │
└─────────────────────────────────────────────────────────────┘
```
