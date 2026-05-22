# Step: 02A-create-scratch

## Description
This stage handles the generation of entirely new `.pptx` files using `pptxgenjs`.

## Purpose
- Create visually appealing and modern presentations when no template is available.
- Use programmatic logic to construct slides, shapes, text, and charts.

## Pre-stage Checkpoint
- Mode selection must be "Create".
- Ensure `pptxgenjs` is installed (`npm install -g pptxgenjs`).

### Version Control
- Ensure the working directory is clean or tracked.

## Workflow

### Process
1. **Design Planning**: Review the visual principles in `[design-and-qa-guidelines.md](../references/design-and-qa-guidelines.md)` to select color palettes, typography, and layouts. Do NOT create plain black-and-white bullet slides.
2. **Script Generation**: Write a Node.js script that uses `pptxgenjs` to construct the presentation. Reference `[pptxgenjs-reference.md](../references/pptxgenjs-reference.md)` for precise API usage, including dimensions, shapes, images, charts, and common pitfalls (e.g., never use `#` with hex colors).
3. **Execution**: Run the generated script (`node generate.js`) to produce the `.pptx` file.

### Output Format
- A finalized `.pptx` file generated from scratch.

## Post-stage Checkpoint

### Progress Tracking
- Check off `3. Create from Scratch` in `.agents/skills-diary/pptx-processing/[<instance-name>]/checklist.md`.

### Version Control
- `git add <generate.js> <output.pptx>` and `git commit -m "Generated new presentation using pptxgenjs"`

### Human in the Loop (HITL)
- Inform the user that the presentation was created and present the file for review.

### Auto pilot
- Automatically proceed to Finalization for visual QA and image generation.
