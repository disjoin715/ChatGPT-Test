# Resource Strategy Slides

This project generates an editable PowerPoint version of the Resource Strategy Outlook slide using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

## Generate the slide deck

Run the build script to produce `dist/deck.pptx`:

```bash
npm install
npm run build:slides
```

## Notes
- The deck is emitted to `dist/deck.pptx` and ignored from version control; retrieve it from CI artifacts or generate locally.
- The slide is fully editable in PowerPoint (text, shapes, and layout are vector-based, not an embedded image).
