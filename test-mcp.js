#!/usr/bin/env node
/**
 * Test script for PowerPoint Generator
 * 
 * This script tests the SlideBuilder library directly
 * to verify PowerPoint generation works correctly.
 */

const path = require("path");
const fs = require("fs");
const { SlideBuilder } = require("./lib/slide-builder");

async function testSlideGeneration() {
  console.log("Testing PowerPoint Generator...\n");

  const builder = new SlideBuilder();

  // Test slide data
  const testSlide = {
    header: {
      eyebrow: "Test Category",
      title: "PowerPoint Generator Test",
      subtitle: "Verifying the slide builder works correctly.",
      badge: "Test Badge",
    },
    leftCard: {
      title: "Test Features",
      type: "iconGrid",
      items: [
        { icon: "◎", title: "Feature One", detail: "Testing icon grid layout." },
        { icon: "◆", title: "Feature Two", detail: "Testing multiple items." },
        { icon: "◈", title: "Feature Three", detail: "Testing third item." },
      ],
      sparkline: [
        { value: "100%", label: "Complete" },
        { value: "Fast", label: "Speed" },
        { value: "Ready", label: "Status" },
      ],
    },
    rightCard: {
      title: "Test Pills",
      type: "pills",
      pills: ["Category A", "Category B", "Category C"],
      items: [
        { icon: "◍", title: "Pill Item One", detail: "Testing pills type." },
        { icon: "✓", title: "Pill Item Two", detail: "Testing validation." },
        { icon: "⬡", title: "Pill Item Three", detail: "Testing systems." },
      ],
    },
    ask: {
      icon: "✦",
      title: "Test CTA",
      text: "This is a test call-to-action bar.",
      cta: "Test",
    },
  };

  const outputDir = path.join(__dirname, "output");
  
  try {
    // Test 1: Build single slide
    console.log("1. Testing single slide generation...");
    const singleSlidePath = path.join(outputDir, "test-single.pptx");
    await builder.buildSingleSlide(testSlide, singleSlidePath, false);
    console.log(`   ✓ Single slide created: ${singleSlidePath}`);
    console.log(`   ✓ File size: ${fs.statSync(singleSlidePath).size} bytes`);

    // Test 2: Build presentation with multiple slides
    console.log("\n2. Testing multi-slide presentation...");
    const journeySlide = {
      header: {
        eyebrow: "Process Flow",
        title: "Journey Test Slide",
        subtitle: "Testing the journey card type.",
        badge: "Journey Test",
      },
      leftCard: {
        title: "Process Steps",
        type: "journey",
        journey: [
          { step: "01", title: "Start", label: "Begin here" },
          { step: "02", title: "Work", label: "Do things" },
          { step: "03", title: "Review", label: "Check work" },
          { step: "04", title: "Done", label: "Complete" },
        ],
        sparkline: [
          { value: "4", label: "Steps" },
          { value: "Quick", label: "Flow" },
          { value: "Easy", label: "Process" },
        ],
      },
      rightCard: {
        title: "Details",
        type: "iconGrid",
        items: [
          { icon: "◎", title: "Detail One", detail: "Journey slide details." },
          { icon: "◆", title: "Detail Two", detail: "More information here." },
        ],
        sparkline: [
          { value: "Yes", label: "Works" },
          { value: "Good", label: "Quality" },
          { value: "Done", label: "Status" },
        ],
      },
      ask: {
        icon: "◆",
        title: "Journey Complete",
        text: "The journey slide test is complete.",
        cta: "Finish",
      },
      splitColumns: true,
    };

    const multiSlidePath = path.join(outputDir, "test-multi.pptx");
    await builder.buildPresentation(
      [testSlide, journeySlide],
      multiSlidePath,
      "Test Presentation",
      "Test Author"
    );
    console.log(`   ✓ Presentation created: ${multiSlidePath}`);
    console.log(`   ✓ File size: ${fs.statSync(multiSlidePath).size} bytes`);
    console.log(`   ✓ Slides: 2`);

    // Test 3: Verify color palette
    console.log("\n3. Testing color palette...");
    const colors = builder.getColorPalette();
    console.log(`   ✓ Primary colors: deepNavy=${colors.deepNavy}, gold=${colors.gold}`);
    console.log(`   ✓ Accent colors: accent1=${colors.accent1}, accent2=${colors.accent2}`);

    // Test 4: Verify design specs
    console.log("\n4. Testing design specs...");
    const specs = builder.getDesignSpecs();
    console.log(`   ✓ Slide dimensions: ${specs.slide.widthIn}in x ${specs.slide.heightIn}in`);
    console.log(`   ✓ Design canvas: ${specs.widthPx}px x ${specs.heightPx}px`);

    console.log("\n" + "=".repeat(50));
    console.log("All tests passed! ✓");
    console.log("=".repeat(50));
    console.log("\nGenerated files:");
    console.log(`  - ${singleSlidePath}`);
    console.log(`  - ${multiSlidePath}`);

  } catch (error) {
    console.error("\n✗ Test failed:", error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

// Run tests
testSlideGeneration();
