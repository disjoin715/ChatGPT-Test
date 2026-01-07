# Resource Strategy Slides

This project generates an editable PowerPoint version of the Resource Strategy Outlook slide using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

## Generate the slide deck

Run the build script to produce `artifacts/slides/latest.pptx`:

```bash
npm ci
npm run build:slides
```

## Publish the slide deck

Use the GitHub Action "Publish Slides (commit PPTX to artifacts branch)" via
`workflow_dispatch` to publish the latest deck.

The published output is committed to the `artifacts` branch at
`slides/latest.pptx`.

## Notes
- The deck is emitted to `artifacts/slides/latest.pptx` and ignored from version control; retrieve it from the artifacts branch or generate locally.
- The slide is fully editable in PowerPoint (text, shapes, and layout are vector-based, not an embedded image).
- Layout math uses a 1280×720 design grid mapped onto a 10×5.625 inch slide. Pixels are converted to inches with `px / 128`, and CSS-like font sizes are converted to points with `px * 72 / 96`. All text boxes and shapes receive explicit widths/heights to avoid zero-sized OOXML nodes.
- The deck uses the Windows-safe font family `Segoe UI` for consistent sizing in PowerPoint.
- The build step post-processes the generated PPTX to normalize zero-dimension group extents (a PptxGenJS quirk) so PowerPoint opens the file cleanly.
