const path = require("path");
const { buildSlides } = require("./scripts/build-slides");

const outputPath = path.join(__dirname, "dist", "deck.pptx");

buildSlides(outputPath)
  .then((filePath) => console.log(`Presentation created: ${filePath}`))
  .catch((err) => {
    console.error("Failed to build slides:", err);
    process.exit(1);
  });
