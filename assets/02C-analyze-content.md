# Step: 02C-analyze-content

## Description
This stage focuses on extracting text, visual grids, or raw XML from an existing `.pptx` file without modifying it.

## Purpose
- Retrieve presentation content for analysis or downstream usage.
- Create visual overviews (thumbnails) to understand slide layouts.
- Inspect raw XML structures for debugging or deep analysis.

## Pre-stage Checkpoint
- The user must have provided a valid `.pptx` file path.
- Mode selection must be "Read".

### Version Control
- N/A (read-only operations).

## Workflow

### Process
1. **Text Extraction**: Use `python -m markitdown presentation.pptx` to extract plain text.
2. **Visual Overview**: Use `python scripts/thumbnail.py presentation.pptx` to generate a visual grid of slides (useful for template analysis).
3. **Raw XML Extraction**: Use `python scripts/office/unpack.py presentation.pptx unpacked/` to extract and pretty-print the underlying XML for inspection.
4. Execute only the steps relevant to the user's specific read request.

### Output Format
- Standard output containing text, or generated image/directory files.

## Post-stage Checkpoint

### Progress Tracking
- Check off `3. Analyze Content` in `.agents/skills-diary/pptx-processing/[<instance-name>]/checklist.md`.

### Version Control
- N/A.

### Human in the Loop (HITL)
- Present the extracted content or paths to the user and await further instructions.

### Auto pilot
- Proceed to Finalization and Verification.
