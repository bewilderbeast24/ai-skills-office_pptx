---
name: office-pptx
description: "Use when a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file; editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments."
---

# `pptx-processing`

## Skill Overview

A multi-purpose skill designed to handle all aspects of `.pptx` file processing. This skill provides three distinct modes of operation:
1. **Read**: Extracting text, analyzing visual structure, or viewing raw XML content.
2. **Edit**: Using template-based workflows to modify, duplicate, delete, and re-pack existing presentations.
3. **Create**: Building modern, visually impressive presentations from scratch using `pptxgenjs`.

## Workflow Sequence

| Stage | Description | Workflow | Input | Output |
| :--- | :--- | :--- | :--- | :--- |
| **1. Mode Selection** | Present options to the user and capture choice. | [01-mode-selection.md](assets/01-mode-selection.md) | User Intent / `.pptx` file | Selected Mode |
| **2. Context Setup** | Load configurations/assets for the chosen mode. | Internal Processing | Selected Mode | Context |
| **3. Specialized Workflow** | Execute the logic specific to the selected mode. | Mode specific `assets/02*-...` file | Context | Results |
| **4. Outcome** | Finalize based on mode-specific criteria and verify output. | Finalization & QA | Results | Final Output |

## Pre-stage Checkpoint

- **Autopilot vs HITL**: By default, the skill operates with Human In The Loop (HITL) for major state transitions (e.g., verifying template design or QA checks). If `Autopilot mode` is explicitly requested, the agent will autonomously perform all steps, assuming its best-fit decisions align with the user's requirements.
- **Dependencies**: Ensure all required dependencies are installed (`markitdown`, `Pillow`, `python-pptx`, `pdf2image`, `numpy`, `pptxgenjs`). See [design-and-qa-guidelines.md](references/design-and-qa-guidelines.md) for details.

## Core Operation Flow

### Global Stages

1. **Initialization**: Initialize progress tracking via `.agents/skills-diary/pptx-processing/[<instance-name>]/checklist.md`.
2. **Mode Selection**: Use `assets/01-mode-selection.md` to determine whether the task requires *Reading*, *Editing*, or *Creating*. Save this mode selection to your internal state.

---

### Execution: Mode Read

- **Asset Path**: `assets/02C-analyze-content.md`
- Process the presentation using `markitdown`, `thumbnail.py`, or XML extraction as needed to gather the requested information.

---

### Execution: Mode Edit

- **Asset Path**: `assets/02B-edit-presentation.md`
- **References**: Strictly follow [editing-reference.md](references/editing-reference.md) for detailed XML modification strategies.
- Process: Analyze the template, unpack it into XML, make required modifications safely, clean orphaned files, and pack the result.

---

### Execution: Mode Create

- **Asset Path**: `assets/02A-create-scratch.md`
- **References**: Strictly follow [pptxgenjs-reference.md](references/pptxgenjs-reference.md) for code syntax and styling best practices.
- Process: Draft content, build the `pptxgenjs` script, execute it, and render the output.

---

### Finalization

- Perform QA on the resulting slides (for Edit and Create modes) according to [design-and-qa-guidelines.md](references/design-and-qa-guidelines.md). Convert slides to images and manually inspect for overflow, low contrast, or alignment issues.
- Do NOT skip the QA step under any circumstances.

## Handover & Confirmation

- Ensure all requirements from the user's initial prompt are satisfied.
- Provide a summary of the executed actions and the final output files.
- Display the generated directory/file tree if applicable.
- Confirm with the user if further iterations or fixes are required (unless in Autopilot).

## Additional Instructions

- All temporary files or unpack directories generated during execution must be properly cleaned up using the provided scripts (e.g., `clean.py`).

