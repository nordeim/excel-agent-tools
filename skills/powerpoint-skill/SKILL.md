---
name: powerpoint-skill
description: Create, edit, and validate PowerPoint presentations using the PowerPoint Agent Tools CLI suite. Provides systematic workflows for building presentations from scratch, modifying existing decks, merging presentations, and ensuring WCAG accessibility compliance. Use when asked to create a PowerPoint, build a presentation, edit slides, add charts/tables/images, merge decks, or validate presentation accessibility.
---

# PowerPoint Presentation Skill

Systematic creation and modification of PowerPoint (.pptx) presentations via 44 stateless CLI tools.

## Core Principles

1. **Clone-before-edit** — never modify source files; always clone first
2. **Probe-before-operate** — inspect structure before adding content
3. **JSON-first I/O** — all tools output structured JSON to stdout
4. **Version tracking** — capture `presentation_version` before/after mutations
5. **Index refresh** — re-query slide info after structural changes (add/remove shapes, z-order)

## Quick Start: Create a New Presentation

```bash
# 1. Create blank presentation
uv run tools/ppt_create_new.py --output work.pptx --json

# 2. Probe capabilities (layouts, theme, fonts)
uv run tools/ppt_capability_probe.py --file work.pptx --deep --json > probe.json

# 3. Add slides with layouts
uv run tools/ppt_add_slide.py --file work.pptx --layout "Title Slide" --json
uv run tools/ppt_add_slide.py --file work.pptx --layout "Title and Content" --json

# 4. Populate content (titles, text, shapes, images, charts, tables)
uv run tools/ppt_set_title.py --file work.pptx --slide 0 --title "My Presentation" --json
uv run tools/ppt_add_text_box.py --file work.pptx --slide 1 --text "Hello World" --position '{"left":"10%","top":"20%"}' --size '{"width":"80%","height":"30%"}' --json

# 5. Validate and check accessibility
uv run tools/ppt_validate_presentation.py --file work.pptx --policy standard --json
uv run tools/ppt_check_accessibility.py --file work.pptx --json
```

## Editing an Existing Presentation

```bash
# 1. CLONE first (mandatory safety)
uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx --json

# 2. Probe structure
uv run tools/ppt_get_info.py --file work.pptx --json

# 3. Make changes...

# 4. Validate
uv run tools/ppt_validate_presentation.py --file work.pptx --policy standard --json
```

## Position & Size Formats

**Position** (use percentage for responsive layouts):
- `{"left":"10%","top":"20%"}` — percentage (recommended)
- `{"left":1.5,"top":2.0}` — inches from top-left
- `{"anchor":"center"}` — named anchor points

**Size**:
- `{"width":"80%","height":"50%"}` — percentage
- `{"width":8.0,"height":4.5}` — inches
- `{"width":"50%","height":"auto"}` — auto aspect ratio

## Destructive Operations Require Approval Tokens

Slide deletion, shape removal, and presentation merging require `--approval-token`. Use the bundled script:

```bash
python scripts/generate_token.py --scope "slide:delete:0"
python scripts/generate_token.py --scope "shape:remove:0:3"
python scripts/generate_token.py --scope "merge:presentations:2"
```

## Key Patterns

- **Overlay**: add shape → refresh indices → send_to_back → refresh again
- **Chart update**: remove old chart (with token) → refresh indices → add new chart
- **Safe mutation**: capture version_before → mutate → capture version_after → verify changed

## Reference Files

- **Tool catalog** — `references/tool-catalog.md` — all 42 tools by category with arguments
- **Safety protocols** — `references/safety-protocols.md` — clone, tokens, versioning, index refresh
- **Workflow guide** — `references/workflow-guide.md` — step-by-step workflows for common tasks

## Troubleshooting (E2E-Validated)

| Symptom | Cause | Fix |
|---------|-------|-----|
| `jq: parse error` | Non-JSON on stdout | Always use `--json` flag; no `print()` in tools |
| `Shape index X out of range` | Indices shifted after structural change | Run `ppt_get_slide_info.py` to refresh |
| `Approval token required` (exit 4) | Missing token for destructive op | Generate with `scripts/generate_token.py` |
| `ppt_add_slide.py --title` fails | No `--title` arg on this tool | Use `ppt_set_title.py` separately |
| `ppt_add_shape.py --shape-type` fails | Arg is `--shape` not `--shape-type` | Use `--shape rectangle` |
| `ppt_set_footer.py --show-page-number` fails | Arg is `--show-number` | Use `--show-number` |
| PDF/Image export fails | LibreOffice not installed | Install `libreoffice-impress` (optional) |
| `ppt_remove_shape.py` silently succeeds | Tool now requires token | Add `--approval-token` argument |

## Exit Codes

| Code | Meaning | Recovery |
|------|---------|----------|
| 0 | Success | Proceed |
| 1 | Usage/General Error | Fix arguments |
| 2 | Validation Error | Fix input format |
| 3 | Transient/Timeout | Retry with backoff |
| 4 | Permission (token missing) | Generate approval token |
| 5 | Internal Error | Check logs, restore backup |
