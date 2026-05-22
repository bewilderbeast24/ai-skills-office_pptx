# Step: 01-mode-selection

## Description
This is the initial stage where the specific mode of operation for processing the `.pptx` file is determined.

## Purpose
- Identify whether the user's intent is to read, edit, or create a presentation.
- Route the workflow to the appropriate specialized asset file.

## Pre-stage Checkpoint
- Ensure the user has provided a clear intent or a target `.pptx` file.

### Version Control
- N/A for mode selection.

## Workflow

### Process
1. Analyze the user's request.
   - If the request is about extracting text, viewing XML, or getting a visual overview without making changes, select **Read**.
   - If the request involves modifying an existing `.pptx` file (adding slides, changing text, editing layout), select **Edit**.
   - If the request involves building a new presentation from scratch or using `pptxgenjs`, select **Create**.
2. If the user's intent is ambiguous, use an interactive prompt or ask a clarifying question.
3. Save the selected mode to your internal context.

### Output Format
- Target mode (Read, Edit, or Create) stored in agent's internal state.

## Post-stage Checkpoint

### Progress Tracking
- Check off `1. Mode Selection` in `.agents/skills-diary/pptx-processing/[<instance-name>]/checklist.md`.

### Version Control
- N/A.

### Human in the Loop (HITL)
- Wait for user clarification if the intent is ambiguous.

### Auto pilot
- Infer the best mode based on the user's prompt and proceed to the corresponding `02-*.md` step.
