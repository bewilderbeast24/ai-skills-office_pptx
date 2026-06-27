# Step: 02B-edit-presentation

## Description
This stage handles all modifications to existing `.pptx` files using a template-based unpack/edit/pack workflow.

## Purpose
- Safely modify text, layout, or structure in a `.pptx` file without causing corruption.
- Utilize parallel subagents for slide-by-slide XML editing.

## Pre-stage Checkpoint
- A template or target `.pptx` file must be identified.
- Mode selection must be "Edit".

### Version Control
- Ensure you have a backup of the original presentation or are operating in a version-controlled directory.

## Workflow

### Process
1. **Analyze Template**: Run `python scripts/thumbnail.py template.pptx` to review slide layouts and decide on mappings.
2. **Unpack**: Run `python scripts/office/unpack.py template.pptx unpacked/`.
3. **Structural Changes**: Modify `ppt/presentation.xml` (`<p:sldIdLst>`). Delete unwanted slides, duplicate needed ones (`python scripts/add_slide.py`), and reorder. Complete this BEFORE content edits.
4. **Edit Content**: Modify each `slide{N}.xml` file. Use subagents if available for parallel processing. Replace placeholders with final content. **Crucial**: Refer to `[editing-reference.md](../references/editing-reference.md)` for exact formatting rules (e.g., smart quotes, bullets, paragraphs).
5. **Clean**: Run `python scripts/clean.py unpacked/` to remove orphaned files.
6. **Pack**: Run `python scripts/office/pack.py unpacked/ output.pptx --original template.pptx`.

### Output Format
- A new `.pptx` file containing the edited presentation.

## Post-stage Checkpoint

### Progress Tracking
- Check off `3. Edit Presentation` in `.agents/skills-diary/pptx-processing/[<instance-name>]/checklist.md`.

### Version Control
- `git add <output.pptx>` and `git commit -m "Generated edited presentation"`

### Human in the Loop (HITL)
- Present the output file and ask the user if visual QA is satisfactory.

### Auto pilot
- Automatically proceed to the Finalization stage to perform visual QA using `scripts/render_slides.py`.
