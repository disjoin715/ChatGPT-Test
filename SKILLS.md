# PowerPoint Generator Skills

> **For AI Coding Agents:** This document describes how to use the PowerPoint generation MCP tools and library to create professional slides following the MA_Theme style guide.

---

## Table of Contents

1. [Overview](#overview)
2. [MCP Server Setup](#mcp-server-setup)
3. [Available Tools](#available-tools)
4. [Slide Data Structure](#slide-data-structure)
5. [Content Types](#content-types)
6. [Usage Examples](#usage-examples)
7. [Best Practices](#best-practices)
8. [Programmatic API](#programmatic-api)

---

## Overview

This workspace provides an MCP (Model Context Protocol) server for generating PowerPoint presentations that follow the MA_Theme style guide. The generated slides feature:

- **Consistent styling**: Deep navy backgrounds, gold accents, Segoe UI typography
- **Professional layouts**: Header, two-column cards, call-to-action bar
- **Rich content types**: Icon grids, pills/tags, journey flows, sparkline metrics
- **Production-ready output**: 16:9 aspect ratio, normalized PPTX files

### When to Use

Use the PowerPoint generator when:
- Creating executive presentations or strategy decks
- Building consistent slide templates for reports
- Generating data-driven slides programmatically
- Producing branded materials following the style guide

---

## MCP Server Setup

### Configuration for Cursor/Claude

Add to your MCP settings (e.g., `~/.cursor/mcp.json` or Claude Desktop config):

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

### Starting the Server Manually

```bash
cd /workspace
npm install
node mcp-server.js
```

### Available Resources

The server exposes these resources:
- `pptx://style-guide` - Complete style guide documentation
- `pptx://theme-reference` - MA_Theme.thmx formatting reference

---

## Available Tools

### 1. `generate_slide`

Generate a single PowerPoint slide with custom content.

**Parameters:**
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `outputPath` | string | No | Output file path (default: `./output/slide.pptx`) |
| `header` | object | Yes | Header section content |
| `leftCard` | object | Yes | Left content card |
| `rightCard` | object | Yes | Right content card |
| `ask` | object | Yes | Call-to-action bar |
| `splitColumns` | boolean | No | Use 50/50 column split (default: 55/45) |

### 2. `generate_presentation`

Generate a complete presentation with multiple slides.

**Parameters:**
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `outputPath` | string | No | Output file path |
| `title` | string | No | Presentation title |
| `author` | string | No | Presentation author |
| `slides` | array | Yes | Array of slide definitions |

### 3. `list_templates`

List available slide templates with examples.

### 4. `get_color_palette`

Get the complete color palette from the style guide.

### 5. `get_icons`

Get recommended Unicode icons/symbols for slides.

---

## Slide Data Structure

### Complete Slide Schema

```javascript
{
  header: {
    eyebrow: "CATEGORY TEXT",        // Uppercase, max 25 chars
    title: "Main Slide Title",       // Max 50 chars
    subtitle: "Supporting description text here.", // Max 80 chars
    badge: "Badge Label"             // Max 30 chars
  },
  leftCard: {
    title: "Card Title",             // Uppercase in output
    type: "iconGrid" | "pills" | "journey",
    items: [...],                    // For iconGrid or pills
    pills: [...],                    // For pills type
    journey: [...],                  // For journey type
    sparkline: [...]                 // Optional metrics row
  },
  rightCard: {
    // Same structure as leftCard
  },
  ask: {
    icon: "✦",                       // Unicode symbol
    title: "CTA Title",              // Max 25 chars
    text: "Call-to-action description.", // Max 100 chars
    cta: "Button"                    // Max 15 chars
  },
  splitColumns: false                // Optional: 50/50 split
}
```

---

## Content Types

### Icon Grid (`type: "iconGrid"`)

Vertical list of items with icon, title, and detail text.

```javascript
{
  type: "iconGrid",
  items: [
    { 
      icon: "◎",                     // Unicode symbol
      title: "Feature Name",         // Max 30 chars
      detail: "Brief description of the feature." // Max 60 chars
    },
    // ... 2-4 items recommended
  ],
  sparkline: [                       // Optional
    { value: "95%", label: "Accuracy" },
    { value: "2.5x", label: "Faster" },
    { value: "100+", label: "Users" }
  ]
}
```

### Pills with Items (`type: "pills"`)

Row of pill/tag labels followed by icon items.

```javascript
{
  type: "pills",
  pills: ["Category A", "Category B", "Category C"], // Exactly 3
  items: [
    { icon: "◍", title: "Item One", detail: "Description here." },
    { icon: "✓", title: "Item Two", detail: "Another description." },
    { icon: "⬡", title: "Item Three", detail: "Final description." }
  ]
}
```

### Journey (`type: "journey"`)

4-step horizontal process flow.

```javascript
{
  type: "journey",
  journey: [
    { step: "01", title: "Discover", label: "Research phase" },
    { step: "02", title: "Design", label: "Solution design" },
    { step: "03", title: "Build", label: "Development" },
    { step: "04", title: "Launch", label: "Go-live" }
  ],
  sparkline: [
    { value: "4 weeks", label: "Duration" },
    { value: "Agile", label: "Method" },
    { value: "Iterative", label: "Approach" }
  ]
}
```

### Sparkline Metrics

Bottom metrics row (always 3 items).

```javascript
sparkline: [
  { value: "Policy", label: "Foresight" },
  { value: "30%", label: "Faster" },
  { value: "Enterprise", label: "Scale" }
]
```

---

## Usage Examples

### Example 1: Strategy Overview Slide

```javascript
// Using MCP tool: generate_slide
{
  "outputPath": "./output/strategy.pptx",
  "header": {
    "eyebrow": "Digital Transformation",
    "title": "AI-Powered Operations",
    "subtitle": "Modernizing workflows with intelligent automation.",
    "badge": "2024 Initiative"
  },
  "leftCard": {
    "title": "Key Capabilities",
    "type": "iconGrid",
    "items": [
      { "icon": "◎", "title": "Process Automation", "detail": "Streamline repetitive tasks with ML models." },
      { "icon": "◆", "title": "Predictive Analytics", "detail": "Forecast trends and identify risks early." },
      { "icon": "◈", "title": "Natural Language", "detail": "Extract insights from unstructured data." }
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
    "pills": ["Data Quality", "Governance", "Talent"],
    "items": [
      { "icon": "◍", "title": "Golden sources", "detail": "Curated datasets with full lineage." },
      { "icon": "✓", "title": "Model oversight", "detail": "Validation, monitoring, and audit trails." },
      { "icon": "⬡", "title": "Skills pipeline", "detail": "Training programs for technical teams." }
    ]
  },
  "ask": {
    "icon": "✦",
    "title": "Next Steps",
    "text": "Secure funding and establish cross-functional team by Q2.",
    "cta": "Approve"
  }
}
```

### Example 2: Process Journey Slide

```javascript
{
  "header": {
    "eyebrow": "Implementation Roadmap",
    "title": "From Pilot to Production",
    "subtitle": "A phased approach to enterprise AI deployment.",
    "badge": "Lifecycle Model"
  },
  "leftCard": {
    "title": "Delivery Phases",
    "type": "journey",
    "journey": [
      { "step": "01", "title": "Assess", "label": "Business case" },
      { "step": "02", "title": "Pilot", "label": "POC delivery" },
      { "step": "03", "title": "Scale", "label": "Production" },
      { "step": "04", "title": "Operate", "label": "BAU support" }
    ],
    "sparkline": [
      { "value": "12 weeks", "label": "Pilot" },
      { "value": "Agile", "label": "Delivery" },
      { "value": "CI/CD", "label": "Ops" }
    ]
  },
  "rightCard": {
    "title": "Governance Model",
    "type": "iconGrid",
    "items": [
      { "icon": "✓", "title": "Risk assessment", "detail": "Model risk framework aligned to policy." },
      { "icon": "◆", "title": "Ethics review", "detail": "Fairness, transparency, accountability." },
      { "icon": "◎", "title": "Audit readiness", "detail": "Documentation and explainability." }
    ],
    "sparkline": [
      { "value": "Tier 1", "label": "Risk class" },
      { "value": "Quarterly", "label": "Review" },
      { "value": "Full", "label": "Audit" }
    ]
  },
  "ask": {
    "icon": "◆",
    "title": "Governance Alignment",
    "text": "Ensure all pilots complete risk assessment before scaling.",
    "cta": "Commit"
  },
  "splitColumns": true
}
```

### Example 3: Multi-Slide Presentation

```javascript
// Using MCP tool: generate_presentation
{
  "outputPath": "./output/full-deck.pptx",
  "title": "Q4 Strategy Review",
  "author": "Strategy Team",
  "slides": [
    {
      "header": { /* slide 1 header */ },
      "leftCard": { /* slide 1 left */ },
      "rightCard": { /* slide 1 right */ },
      "ask": { /* slide 1 CTA */ }
    },
    {
      "header": { /* slide 2 header */ },
      "leftCard": { /* slide 2 left */ },
      "rightCard": { /* slide 2 right */ },
      "ask": { /* slide 2 CTA */ },
      "splitColumns": true
    }
    // ... more slides
  ]
}
```

---

## Best Practices

### Content Guidelines

1. **Keep text concise**
   - Eyebrow: Max 25 characters
   - Title: Max 50 characters
   - Subtitle: Max 80 characters
   - Item titles: Max 30 characters
   - Item details: Max 60 characters

2. **Use consistent icons**
   - Primary items: ◎ (U+25CE)
   - Secondary items: ◆ (U+25C6)
   - Tertiary items: ◈ (U+25C8)
   - Data/sources: ◍ (U+25CD)
   - Validation: ✓ (U+2713)
   - Systems: ⬡ (U+2B21)
   - Priority: ✦ (U+2726)

3. **Balance card content**
   - 2-4 items per icon grid
   - Exactly 3 pills in pill row
   - Exactly 4 steps in journey
   - Exactly 3 metrics in sparkline

4. **Choose appropriate layouts**
   - Default 55/45 split for uneven content
   - 50/50 split (`splitColumns: true`) for balanced content

### Styling Notes

- All styling is automatic based on STYLE_GUIDE.md
- Colors, fonts, shadows, and spacing are predefined
- Focus on content; formatting is handled by the builder

### Common Patterns

| Use Case | Left Card | Right Card |
|----------|-----------|------------|
| Feature overview | iconGrid | pills |
| Process flow | journey | iconGrid |
| Comparison | iconGrid | iconGrid |
| Categories | pills | pills |

---

## Programmatic API

### Using the SlideBuilder Library Directly

```javascript
const { SlideBuilder } = require('./lib/slide-builder');

const builder = new SlideBuilder();

// Build single slide
const slideData = {
  header: { /* ... */ },
  leftCard: { /* ... */ },
  rightCard: { /* ... */ },
  ask: { /* ... */ }
};

await builder.buildSingleSlide(slideData, './output/slide.pptx', false);

// Build presentation
const slides = [slideData1, slideData2, slideData3];
await builder.buildPresentation(
  slides, 
  './output/deck.pptx',
  'Presentation Title',
  'Author Name'
);

// Get color palette
const colors = builder.getColorPalette();
console.log(colors.deepNavy); // "0F2439"

// Get design specs
const specs = builder.getDesignSpecs();
console.log(specs.slide.widthIn); // 10
```

### Color Palette Reference

```javascript
const palette = {
  // Primary
  deepNavy: "0F2439",   // Background, headers, CTA
  midNavy: "17395c",    // Icon backgrounds
  gold: "e1b44c",       // Buttons, accents

  // Neutral
  text: "0c1b2a",       // Body text
  muted: "4a5c70",      // Subtitles
  softGray: "f4f6f8",   // Metric boxes
  cardBg: "FFFFFF",     // Card fills
  shellBg: "f6f9fe",    // Content area

  // Pills
  pillBg: "fdf6e8",     // Pill background
  pillColor: "6f5316",  // Pill text

  // Theme accents (from MA_Theme.thmx)
  accent1: "604878",    // Purple
  accent2: "D86B77",    // Coral/pink
  accent3: "8EC182",    // Light green
  accent4: "F9B268",    // Soft orange
  accent5: "1B587C",    // Dark teal
  accent6: "B26B02",    // Brown/amber
};
```

---

## Troubleshooting

### Common Issues

1. **"Module not found" error**
   ```bash
   npm install  # Ensure dependencies are installed
   ```

2. **Output file not created**
   - Check that output directory exists or is creatable
   - Verify write permissions

3. **Slide content truncated**
   - Reduce text length to fit within guidelines
   - Use fewer items in lists

4. **Fonts not rendering correctly**
   - Ensure Segoe UI is installed on the system
   - PowerPoint will substitute if font is missing

### Getting Help

1. Read the full STYLE_GUIDE.md for visual specifications
2. Use `list_templates` tool for content structure examples
3. Use `get_color_palette` for color values
4. Use `get_icons` for recommended symbols

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-07 | Initial skills documentation |
