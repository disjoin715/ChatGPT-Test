# PowerPoint Generator with MCP Support

This project provides **MCP (Model Context Protocol) tools** for generating professionally styled PowerPoint presentations, plus a traditional HTML-to-PPTX conversion pipeline using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

> **For AI Coding Agents:** This project includes an MCP server that exposes PowerPoint generation as callable tools. See [SKILLS.md](./SKILLS.md) for usage patterns and [STYLE_GUIDE.md](./STYLE_GUIDE.md) for visual design specifications.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [MCP Server Setup](#mcp-server-setup)
3. [End-to-End Flow](#end-to-end-flow)
4. [Project Structure](#project-structure)
5. [HTML Content Format](#html-content-format)
6. [Slide Data Structure](#slide-data-structure)
7. [Component Reference](#component-reference)
8. [Publishing Workflow](#publishing-workflow)
9. [For AI Agents](#for-ai-agents)

---

## Quick Start

```bash
# Install dependencies
npm install

# Generate the PowerPoint deck (from HTML)
npm run build:slides

# Output: artifacts/slides/latest.pptx

# Run MCP server
npm run mcp

# Test the slide builder
npm test
```

---

## MCP Server Setup

### For Cursor / Claude Desktop

Add to your MCP configuration (e.g., `~/.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "powerpoint-generator": {
      "command": "node",
      "args": ["/path/to/workspace/mcp-server.js"],
      "env": {}
    }
  }
}
```

### Available MCP Tools

| Tool | Description |
|------|-------------|
| `generate_slide` | Create a single PowerPoint slide with custom content |
| `generate_presentation` | Create a full presentation with multiple slides |
| `list_templates` | List available slide templates and content types |
| `get_color_palette` | Get the complete color palette from the style guide |
| `get_icons` | Get recommended Unicode icons/symbols |

### Available MCP Resources

| Resource URI | Description |
|--------------|-------------|
| `pptx://style-guide` | Complete STYLE_GUIDE.md documentation |
| `pptx://theme-reference` | MA_Theme.thmx formatting reference |

### Example Usage (MCP Tool Call)

```javascript
// generate_slide tool
{
  "header": {
    "eyebrow": "Strategy Overview",
    "title": "AI-Powered Operations",
    "subtitle": "Modernizing workflows with intelligent automation.",
    "badge": "2024 Initiative"
  },
  "leftCard": {
    "title": "Key Capabilities",
    "type": "iconGrid",
    "items": [
      { "icon": "◎", "title": "Automation", "detail": "Streamline processes." },
      { "icon": "◆", "title": "Analytics", "detail": "Data-driven insights." }
    ],
    "sparkline": [
      { "value": "40%", "label": "Efficiency" },
      { "value": "3x", "label": "Speed" },
      { "value": "99%", "label": "Accuracy" }
    ]
  },
  "rightCard": {
    "title": "Success Factors",
    "type": "pills",
    "pills": ["Data", "Governance", "Talent"],
    "items": [
      { "icon": "✓", "title": "Quality", "detail": "Clean data." }
    ]
  },
  "ask": {
    "icon": "✦",
    "title": "Next Steps",
    "text": "Approve funding for Phase 2.",
    "cta": "Approve"
  }
}
```

See [SKILLS.md](./SKILLS.md) for comprehensive usage documentation.

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
├── SKILLS.md               # MCP tool usage documentation (for AI agents)
├── MA_Theme.thmx           # PowerPoint theme file (color/font reference)
├── mcp-server.js           # MCP server exposing PowerPoint tools
├── lib/
│   └── slide-builder.js    # Core slide generation library
├── scripts/
│   └── build-slides.js     # HTML-to-PPTX conversion script
├── test-mcp.js             # Test script for slide builder
├── output/                 # Generated test files (gitignored)
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

### Recommended Approach: Use MCP Tools

The easiest way to generate slides is via the MCP server:

1. **Use the `generate_slide` tool** for single slides
2. **Use the `generate_presentation` tool** for full decks
3. **Use `list_templates`** to see available content patterns
4. **Use `get_color_palette`** and `get_icons`** for styling reference

See **[SKILLS.md](./SKILLS.md)** for comprehensive MCP tool documentation.

### Alternative: Direct Library Usage

```javascript
const { SlideBuilder } = require('./lib/slide-builder');

const builder = new SlideBuilder();
await builder.buildSingleSlide(slideData, './output/slide.pptx');
await builder.buildPresentation(slides, './output/deck.pptx', 'Title', 'Author');
```

### Key Files to Reference

| File | Purpose |
|------|---------|
| `SKILLS.md` | **START HERE** - MCP tool usage and examples |
| `STYLE_GUIDE.md` | Visual design preferences (colors, fonts, spacing) |
| `lib/slide-builder.js` | Core library for programmatic usage |
| `mcp-server.js` | MCP server implementation |
| `index.html` | HTML content structure examples |
| `scripts/build-slides.js` | HTML-to-PPTX conversion logic |

### Before Making Changes

1. **Read `SKILLS.md`** - Understand MCP tools and data structures
2. **Read `STYLE_GUIDE.md`** - Contains color palette, typography, spacing
3. **Review `MA_Theme.thmx` section** in STYLE_GUIDE.md - Theme-specific formatting

### Common Tasks

| Task | Recommended Method |
|------|-------------------|
| Generate single slide | MCP `generate_slide` tool |
| Generate full deck | MCP `generate_presentation` tool |
| Add new slide to existing | Modify `slidesData` in `build-slides.js` |
| Change colors | Update `palette` in `lib/slide-builder.js` and `STYLE_GUIDE.md` |
| Add new component type | Create new function in `lib/slide-builder.js` |

### Technical Notes

- **Slide dimensions**: 1280×720 px design grid → 10×5.625 inch output
- **Font**: Segoe UI (Windows-safe, consistent sizing)
- **Theme colors**: MA_Theme.thmx defines custom palette (purple-led, not blue)
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
