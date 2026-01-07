#!/usr/bin/env node
/**
 * MCP Server for PowerPoint Generation
 * 
 * This server exposes tools for generating PowerPoint presentations
 * following the MA_Theme style guide formatting preferences.
 * 
 * Tools:
 * - generate_slide: Create a single slide with custom content
 * - generate_presentation: Create a full presentation with multiple slides
 * - list_templates: List available slide templates/layouts
 * - get_style_guide: Get the current style guide reference
 */

const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
} = require("@modelcontextprotocol/sdk/types.js");
const fs = require("fs");
const path = require("path");
const { SlideBuilder } = require("./lib/slide-builder");

// Initialize server
const server = new Server(
  {
    name: "powerpoint-generator",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
      resources: {},
    },
  }
);

// Tool definitions
const TOOLS = [
  {
    name: "generate_slide",
    description: `Generate a single PowerPoint slide following the MA_Theme style guide.

The slide follows a structured layout with:
- Header: eyebrow text, title, subtitle, and badge
- Two content cards (left and right) with various content types
- Call-to-action bar at the bottom

Content types supported:
- iconGrid: List of items with icon, title, and detail
- pills: Row of pill/tag elements followed by icon items
- journey: Step-by-step process flow (4 steps)

All styling (colors, fonts, shadows, spacing) follows the STYLE_GUIDE.md specifications.`,
    inputSchema: {
      type: "object",
      properties: {
        outputPath: {
          type: "string",
          description: "Output file path for the .pptx file (default: ./output/slide.pptx)",
        },
        header: {
          type: "object",
          description: "Header section content",
          properties: {
            eyebrow: { type: "string", description: "Category/eyebrow text (uppercase, max 25 chars)" },
            title: { type: "string", description: "Main slide title (max 50 chars)" },
            subtitle: { type: "string", description: "Supporting subtitle (max 80 chars)" },
            badge: { type: "string", description: "Badge text in top-right (max 30 chars)" },
          },
          required: ["eyebrow", "title", "subtitle", "badge"],
        },
        leftCard: {
          type: "object",
          description: "Left content card",
          properties: {
            title: { type: "string", description: "Card title (uppercase, max 25 chars)" },
            type: {
              type: "string",
              enum: ["iconGrid", "pills", "journey"],
              description: "Content layout type",
            },
            items: {
              type: "array",
              description: "Content items (for iconGrid or pills type)",
              items: {
                type: "object",
                properties: {
                  icon: { type: "string", description: "Unicode icon symbol" },
                  title: { type: "string", description: "Item title (max 30 chars)" },
                  detail: { type: "string", description: "Item detail text (max 60 chars)" },
                },
              },
            },
            pills: {
              type: "array",
              description: "Pill labels (for pills type, max 3)",
              items: { type: "string" },
            },
            journey: {
              type: "array",
              description: "Journey steps (for journey type, exactly 4)",
              items: {
                type: "object",
                properties: {
                  step: { type: "string", description: "Step number (e.g., '01')" },
                  title: { type: "string", description: "Step title (1-2 words)" },
                  label: { type: "string", description: "Step label/description" },
                },
              },
            },
            sparkline: {
              type: "array",
              description: "Bottom metrics row (max 3)",
              items: {
                type: "object",
                properties: {
                  value: { type: "string", description: "Metric value or label" },
                  label: { type: "string", description: "Metric description" },
                },
              },
            },
          },
          required: ["title", "type"],
        },
        rightCard: {
          type: "object",
          description: "Right content card (same structure as leftCard)",
          properties: {
            title: { type: "string" },
            type: { type: "string", enum: ["iconGrid", "pills", "journey"] },
            items: { type: "array" },
            pills: { type: "array" },
            journey: { type: "array" },
            sparkline: { type: "array" },
          },
          required: ["title", "type"],
        },
        ask: {
          type: "object",
          description: "Call-to-action bar at bottom",
          properties: {
            icon: { type: "string", description: "Unicode icon symbol" },
            title: { type: "string", description: "CTA title (max 25 chars)" },
            text: { type: "string", description: "CTA description (max 100 chars)" },
            cta: { type: "string", description: "Button text (max 15 chars)" },
          },
          required: ["icon", "title", "text", "cta"],
        },
        splitColumns: {
          type: "boolean",
          description: "Use equal 50/50 column split instead of default 55/45",
          default: false,
        },
      },
      required: ["header", "leftCard", "rightCard", "ask"],
    },
  },
  {
    name: "generate_presentation",
    description: `Generate a complete PowerPoint presentation with multiple slides.

Creates a .pptx file with multiple slides, all following the MA_Theme style guide.
Each slide in the array follows the same structure as generate_slide.

Use this for creating full decks with consistent styling.`,
    inputSchema: {
      type: "object",
      properties: {
        outputPath: {
          type: "string",
          description: "Output file path for the .pptx file (default: ./output/presentation.pptx)",
        },
        title: {
          type: "string",
          description: "Presentation title (shown in file properties)",
        },
        author: {
          type: "string",
          description: "Presentation author (shown in file properties)",
        },
        slides: {
          type: "array",
          description: "Array of slide definitions",
          items: {
            type: "object",
            properties: {
              header: { type: "object" },
              leftCard: { type: "object" },
              rightCard: { type: "object" },
              ask: { type: "object" },
              splitColumns: { type: "boolean" },
            },
            required: ["header", "leftCard", "rightCard", "ask"],
          },
        },
      },
      required: ["slides"],
    },
  },
  {
    name: "list_templates",
    description: "List available slide templates and content type options with examples.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
  {
    name: "get_color_palette",
    description: "Get the complete color palette from the style guide for use in slides.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
  {
    name: "get_icons",
    description: "Get recommended Unicode icons/symbols for use in slides.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
];

// Resource definitions
const RESOURCES = [
  {
    uri: "pptx://style-guide",
    name: "PowerPoint Style Guide",
    description: "Complete style guide for PowerPoint generation including colors, typography, spacing, and components",
    mimeType: "text/markdown",
  },
  {
    uri: "pptx://theme-reference",
    name: "MA_Theme Reference",
    description: "Reference documentation for MA_Theme.thmx custom formatting",
    mimeType: "text/markdown",
  },
];

// Handle list tools request
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: TOOLS };
});

// Handle list resources request
server.setRequestHandler(ListResourcesRequestSchema, async () => {
  return { resources: RESOURCES };
});

// Handle read resource request
server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
  const { uri } = request.params;

  if (uri === "pptx://style-guide" || uri === "pptx://theme-reference") {
    const styleGuidePath = path.join(__dirname, "STYLE_GUIDE.md");
    if (fs.existsSync(styleGuidePath)) {
      const content = fs.readFileSync(styleGuidePath, "utf-8");
      return {
        contents: [
          {
            uri,
            mimeType: "text/markdown",
            text: content,
          },
        ],
      };
    }
  }

  throw new Error(`Resource not found: ${uri}`);
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case "generate_slide": {
        const builder = new SlideBuilder();
        const slideData = {
          header: args.header,
          leftCard: args.leftCard,
          rightCard: args.rightCard,
          ask: args.ask,
        };
        
        const outputPath = args.outputPath || "./output/slide.pptx";
        const result = await builder.buildSingleSlide(slideData, outputPath, args.splitColumns);
        
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                message: `Slide generated successfully`,
                outputPath: result,
                styling: "MA_Theme style guide applied",
              }, null, 2),
            },
          ],
        };
      }

      case "generate_presentation": {
        const builder = new SlideBuilder();
        const outputPath = args.outputPath || "./output/presentation.pptx";
        const result = await builder.buildPresentation(
          args.slides,
          outputPath,
          args.title,
          args.author
        );
        
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                message: `Presentation generated with ${args.slides.length} slide(s)`,
                outputPath: result,
                styling: "MA_Theme style guide applied",
              }, null, 2),
            },
          ],
        };
      }

      case "list_templates": {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                templates: [
                  {
                    name: "Strategy Overview",
                    description: "High-level strategic priorities with icon grid and sparkline metrics",
                    leftCard: { type: "iconGrid", items: 3, sparkline: true },
                    rightCard: { type: "pills", pills: 3, items: 3 },
                  },
                  {
                    name: "Process Journey",
                    description: "Step-by-step process flow with supporting details",
                    leftCard: { type: "journey", steps: 4, sparkline: true },
                    rightCard: { type: "iconGrid", items: 3, sparkline: true },
                  },
                  {
                    name: "Feature Comparison",
                    description: "Two icon grids comparing features or options",
                    leftCard: { type: "iconGrid", items: 3, sparkline: true },
                    rightCard: { type: "iconGrid", items: 3, sparkline: true },
                  },
                ],
                contentTypes: {
                  iconGrid: {
                    description: "Vertical list of items with icon, title, and detail text",
                    maxItems: 4,
                    example: {
                      icon: "◎",
                      title: "Feature Name",
                      detail: "Brief description of the feature or capability.",
                    },
                  },
                  pills: {
                    description: "Row of 3 pill/tag labels followed by icon items",
                    maxPills: 3,
                    example: ["Category A", "Category B", "Category C"],
                  },
                  journey: {
                    description: "4-step horizontal process flow",
                    steps: 4,
                    example: {
                      step: "01",
                      title: "Discover",
                      label: "Research phase",
                    },
                  },
                  sparkline: {
                    description: "Bottom metrics row with 3 value/label pairs",
                    maxItems: 3,
                    example: {
                      value: "95%",
                      label: "Accuracy",
                    },
                  },
                },
              }, null, 2),
            },
          ],
        };
      }

      case "get_color_palette": {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                primary: {
                  deepNavy: { hex: "#0F2439", usage: "Background, headers, CTA bar, badge" },
                  midNavy: { hex: "#17395C", usage: "Icon backgrounds, secondary accents" },
                  gold: { hex: "#E1B44C", usage: "CTA buttons, badge dots, bullet points" },
                },
                neutral: {
                  text: { hex: "#0C1B2A", usage: "Primary body text, card titles" },
                  muted: { hex: "#4A5C70", usage: "Subtitles, descriptions, labels" },
                  softGray: { hex: "#F4F6F8", usage: "Sparkline backgrounds, metric boxes" },
                  cardBg: { hex: "#FFFFFF", usage: "Card fills" },
                  shellBg: { hex: "#F6F9FE", usage: "Main slide content area" },
                  border: { hex: "#DDE5EF", usage: "Card borders, dividers" },
                },
                pills: {
                  background: { hex: "#FDF6E8", usage: "Tag/pill fill (warm cream)" },
                  text: { hex: "#6F5316", usage: "Tag/pill text (dark gold)" },
                },
                themeColors: {
                  accent1: { hex: "#604878", usage: "Purple accent" },
                  accent2: { hex: "#D86B77", usage: "Coral/pink accent" },
                  accent3: { hex: "#8EC182", usage: "Light green accent" },
                  accent4: { hex: "#F9B268", usage: "Soft orange accent" },
                  accent5: { hex: "#1B587C", usage: "Dark teal accent" },
                  accent6: { hex: "#B26B02", usage: "Brown/amber accent" },
                  hyperlink: { hex: "#4EA5D8", usage: "Link color" },
                },
                brand: {
                  deepViolet: { hex: "#33018D", usage: "MA_Theme brand color for titles" },
                },
              }, null, 2),
            },
          ],
        };
      }

      case "get_icons": {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                recommended: [
                  { symbol: "◎", unicode: "U+25CE", usage: "Primary/first items" },
                  { symbol: "◆", unicode: "U+25C6", usage: "Secondary items" },
                  { symbol: "◈", unicode: "U+25C8", usage: "Tertiary items" },
                  { symbol: "◍", unicode: "U+25CD", usage: "Data/sources" },
                  { symbol: "✓", unicode: "U+2713", usage: "Governance/validation" },
                  { symbol: "⬡", unicode: "U+2B21", usage: "Infrastructure/systems" },
                  { symbol: "✦", unicode: "U+2726", usage: "Featured/priority" },
                  { symbol: "★", unicode: "U+2605", usage: "Highlights" },
                ],
                additional: [
                  { symbol: "●", unicode: "U+25CF", usage: "Bullet point" },
                  { symbol: "○", unicode: "U+25CB", usage: "Empty circle" },
                  { symbol: "▶", unicode: "U+25B6", usage: "Action/play" },
                  { symbol: "◀", unicode: "U+25C0", usage: "Back/previous" },
                  { symbol: "▲", unicode: "U+25B2", usage: "Up/increase" },
                  { symbol: "▼", unicode: "U+25BC", usage: "Down/decrease" },
                  { symbol: "⚡", unicode: "U+26A1", usage: "Speed/power" },
                  { symbol: "🔒", unicode: "U+1F512", usage: "Security" },
                  { symbol: "📊", unicode: "U+1F4CA", usage: "Analytics" },
                  { symbol: "⚙", unicode: "U+2699", usage: "Settings/config" },
                ],
              }, null, 2),
            },
          ],
        };
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            success: false,
            error: error.message,
          }, null, 2),
        },
      ],
      isError: true,
    };
  }
});

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("PowerPoint Generator MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});
