# HTML to PowerPoint Slide Generator

This project converts HTML-based slide content into fully editable PowerPoint presentations using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

> **For AI Coding Agents:** This README and the accompanying [STYLE_GUIDE.md](./STYLE_GUIDE.md) serve as your primary reference when generating or modifying PowerPoint slides. Always read both files before making changes.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [End-to-End Flow](#end-to-end-flow)
3. [Project Structure](#project-structure)
4. [HTML Content Format](#html-content-format)
5. [Slide Data Structure](#slide-data-structure)
6. [Component Reference](#component-reference)
7. [Publishing Workflow](#publishing-workflow)
8. [For AI Agents](#for-ai-agents)

---

## Quick Start

```bash
# Install dependencies
npm ci

# Generate the PowerPoint deck
npm run build:slides

# Output: artifacts/slides/latest.pptx
```

---

## End-to-End Flow

```
┌─────────────────┐      ┌──────────────────┐      ┌─────────────────┐
│   index.html    │ ──▶  │ build-slides.js  │ ──▶  │  latest.pptx    │
│ (Source Content)│      │ (Conversion)     │      │ (Output)        │
└─────────────────┘      └──────────────────┘      └─────────────────┘
                                │
                                ▼
                    ┌──────────────────────┐
                    │   STYLE_GUIDE.md     │
                    │ (Design Preferences) │
                    └──────────────────────┘
```

### Step 1: Define Content in HTML (`index.html`)

The `index.html` file serves as the **single source of truth** for slide content. It uses semantic HTML with specific CSS classes to define the structure:

- Each `<section class="slide">` represents one PowerPoint slide
- Content is organized into header, cards (left/right), and footer (ask/CTA)
- The HTML is human-readable and can be previewed in a browser

### Step 2: Convert to Data Structure

The build script extracts content from the HTML structure and maps it to a JavaScript data structure (`slidesData` array in `build-slides.js`). Each slide object contains:

```javascript
{
  header: { eyebrow, title, subtitle, badge },
  leftCard: { title, type, items, sparkline, ... },
  rightCard: { title, type, items, sparkline, ... },
  ask: { icon, title, text, cta }
}
```

### Step 3: Generate PowerPoint

The script uses PptxGenJS to create shapes, text boxes, and styling that match the HTML design. Key conversions:

- **Pixels → Inches**: `px / 128` (based on 1280px = 10 inches)
- **Pixels → Points**: `px * 72 / 96`
- **CSS Colors → Hex**: Remove `#` prefix (e.g., `#0f2439` → `0F2439`)

### Step 4: Publish via GitHub Actions

Push to `main` branch triggers the "Publish Slides" workflow, which:
1. Builds the PPTX
2. Commits it to the `artifacts` branch at `slides/latest.pptx`

---

## Project Structure

```
├── index.html              # Source HTML with slide content
├── STYLE_GUIDE.md          # Visual design preferences (READ THIS!)
├── scripts/
│   └── build-slides.js     # Main conversion script
├── artifacts/
│   └── slides/
│       └── latest.pptx     # Generated output (gitignored)
├── .github/
│   └── workflows/
│       ├── build-slides.yml    # CI build workflow
│       └── publish-slides.yml  # Publish to artifacts branch
└── package.json
```

---

## HTML Content Format

### Slide Structure

```html
<section class="slide">
  <!-- Header Region -->
  <header class="header">
    <div class="title-block">
      <div class="eyebrow">CATEGORY LABEL</div>
      <h1>Main Title</h1>
      <p class="subtitle">Supporting description text.</p>
    </div>
    <div class="badge">
      <span class="dot"></span>
      <span>Badge Text</span>
    </div>
  </header>

  <!-- Content Region (Two Cards) -->
  <section class="content">           <!-- or class="content split" for 50/50 -->
    <article class="card">...</article>  <!-- Left card -->
    <article class="card">...</article>  <!-- Right card -->
  </section>

  <!-- Footer Region (Call to Action) -->
  <section class="ask">
    <div class="icon">✦</div>
    <div class="text">
      <h3>CTA Title</h3>
      <p>CTA description text.</p>
    </div>
    <div class="cta">Button Text</div>
  </section>
</section>
```

### Card Content Types

#### Icon Grid (List with icons)
```html
<div class="icon-grid">
  <div class="icon-item">
    <div class="icon-circle">◎</div>
    <div>
      <h3>Item Title</h3>
      <p>Item description text.</p>
    </div>
  </div>
  <!-- More items... -->
</div>
```

#### Sparkline (Metric boxes)
```html
<div class="sparkline">
  <div class="spark">
    <div class="value">30%</div>
    <div class="label">Metric Label</div>
  </div>
  <!-- More metrics... -->
</div>
```

#### Pill Row (Tags/badges)
```html
<div class="pill-row">
  <div class="pill">Tag One</div>
  <div class="pill">Tag Two</div>
  <div class="pill">Tag Three</div>
</div>
```

#### Journey Steps (Process flow)
```html
<div class="journey">
  <div class="journey-step">
    <div class="step-icon">01</div>
    <h4>Step Title</h4>
    <span>Step subtitle</span>
  </div>
  <!-- More steps... -->
</div>
```

---

## Slide Data Structure

When modifying `build-slides.js`, use this data structure:

```javascript
const slidesData = [
  {
    header: {
      eyebrow: "CATEGORY",           // Uppercase label above title
      title: "Main Slide Title",     // Primary heading
      subtitle: "Description text",  // Supporting text
      badge: "Badge Label",          // Right-aligned badge
    },
    leftCard: {
      title: "Card Title",
      type: "iconGrid",              // "iconGrid" | "pills" | "journey"
      items: [                       // For iconGrid type
        { 
          icon: "◎",                 // Unicode symbol
          title: "Item Title", 
          detail: "Description" 
        },
      ],
      pills: ["Tag1", "Tag2"],       // For pills type (optional)
      journey: [                     // For journey type
        { step: "01", title: "Step", label: "Subtitle" }
      ],
      sparkline: [                   // Bottom metrics (optional)
        { value: "30%", label: "Label" }
      ],
    },
    rightCard: {
      // Same structure as leftCard
    },
    ask: {
      icon: "✦",                     // Unicode symbol
      title: "CTA Title",
      text: "CTA description",
      cta: "Button Text",
    },
  },
];
```

---

## Component Reference

| Component | HTML Class | Data Property | Description |
|-----------|------------|---------------|-------------|
| Header | `.header` | `header` | Top section with title, subtitle, badge |
| Eyebrow | `.eyebrow` | `header.eyebrow` | Small uppercase category label |
| Badge | `.badge` | `header.badge` | Right-aligned pill with dot indicator |
| Content | `.content` | - | Container for two cards |
| Card | `.card` | `leftCard`, `rightCard` | White rounded container |
| Icon Grid | `.icon-grid` | `items[]` | List with icon + title + detail |
| Sparkline | `.sparkline` | `sparkline[]` | Row of metric boxes |
| Pill Row | `.pill-row` | `pills[]` | Row of tag badges |
| Journey | `.journey` | `journey[]` | Numbered step sequence |
| Ask/CTA | `.ask` | `ask` | Bottom call-to-action bar |

---

## Publishing Workflow

### Automatic Publishing

The `publish-slides.yml` workflow triggers on:
- Push to `main` branch (when `scripts/**` changes)
- Manual `workflow_dispatch`

### Manual Publishing

```bash
# Via GitHub CLI (if you have permissions)
gh workflow run "Publish Slides (commit PPTX to artifacts branch)"

# Or push changes to main branch
git push origin main
```

### Output Location

Published PPTX is available at:
- Branch: `artifacts`
- Path: `slides/latest.pptx`
- URL: `https://github.com/{owner}/{repo}/blob/artifacts/slides/latest.pptx`

---

## For AI Agents

### Before Making Changes

1. **Read `STYLE_GUIDE.md`** - Contains color palette, typography, spacing, and layout preferences
2. **Review `index.html`** - Understand the current content structure
3. **Check `build-slides.js`** - See how components are implemented

### When Generating New Slides

1. **Follow the data structure** defined in [Slide Data Structure](#slide-data-structure)
2. **Use the color palette** from `STYLE_GUIDE.md`
3. **Maintain consistent spacing** using the `DESIGN` constants
4. **Test locally** with `npm run build:slides` before committing

### Common Tasks

| Task | Action |
|------|--------|
| Add new slide | Add entry to `slidesData` array |
| Change colors | Update `palette` object and `STYLE_GUIDE.md` |
| Modify layout | Update `DESIGN` constants |
| Add new component type | Create new `add{Component}()` function |
| Update content only | Modify `slidesData` entries or `index.html` |

### Key Files to Reference

| File | Purpose |
|------|---------|
| `STYLE_GUIDE.md` | Visual design preferences (colors, fonts, spacing) |
| `index.html` | Source content and HTML structure examples |
| `scripts/build-slides.js` | Conversion logic and component implementations |

### Technical Notes

- **Slide dimensions**: 1280×720 px design grid → 10×5.625 inch output
- **Font**: Segoe UI (Windows-safe, consistent sizing)
- **Shapes**: All use explicit dimensions to avoid zero-sized OOXML nodes
- **Post-processing**: Normalizes zero-dimension group extents (PptxGenJS quirk)
- **Output**: Fully editable vector-based PowerPoint (not images)

---

## Example: Current Slides

The current deck (`index.html`) contains two slides:

### Slide 1: "From Pilot to Production"
- **Theme**: Central Bank AI Strategy overview
- **Left Card**: Priority Use Cases (3 icon items + sparkline)
- **Right Card**: What Enables Scale (3 pills + 3 icon items)
- **CTA**: Executive Priority / Mandate

### Slide 2: "Lifecycle Discipline + Democratized Innovation"
- **Theme**: AI governance and adoption
- **Left Card**: End-to-End Lifecycle (4 journey steps + sparkline)
- **Right Card**: Democratize Innovation (3 icon items + sparkline)
- **CTA**: Strategic Close / Commit

---

## License

ISC
