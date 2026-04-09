# Tool Catalog — 42 PowerPoint Agent Tools

All tools follow this pattern: `uv run tools/ppt_<name>.py --file work.pptx [args] --json`

## Creation (4 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_create_new.py` | Blank presentation | `--output` |
| `ppt_create_from_template.py` | From .pptx template | `--template`, `--output` |
| `ppt_create_from_structure.py` | From JSON definition | `--structure`, `--output` |
| `ppt_clone_presentation.py` | **SAFETY**: Working copy | `--source`, `--output` |

## Slide Management (4 tools)

| Tool | Purpose | Key Args | Token? |
|------|---------|----------|--------|
| `ppt_add_slide.py` | Add slide | `--layout`, `--index` | No |
| `ppt_delete_slide.py` | Remove slide | `--slide` | **Yes** |
| `ppt_duplicate_slide.py` | Clone slide | `--slide` | No |
| `ppt_reorder_slides.py` | Move slide | `--from-index`, `--to-index` | No |

## Shapes (4 tools)

| Tool | Purpose | Key Args | Token? |
|------|---------|----------|--------|
| `ppt_add_shape.py` | Add shape | `--shape-type`, `--position`, `--size`, `--fill-color`, `--fill-opacity` | No |
| `ppt_remove_shape.py` | Delete shape | `--slide`, `--shape` | **Yes** |
| `ppt_format_shape.py` | Style shape | `--slide`, `--shape`, `--fill-color`, `--fill-opacity` | No |
| `ppt_set_z_order.py` | Layer control | `--slide`, `--shape`, `--action` (bring_to_front/send_to_back/bring_forward/send_backward) | No |

Shape types: `rectangle`, `rounded_rectangle`, `oval`, `triangle`, `diamond`, `pentagon`, `hexagon`, `arrow`, `line`, `chevron`, `callout`, `star`

## Text (4 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_add_text_box.py` | Add text | `--slide`, `--text`, `--position`, `--size` |
| `ppt_set_title.py` | Set title/subtitle | `--slide`, `--title`, `--subtitle` |
| `ppt_format_text.py` | Style text | `--slide`, `--shape`, `--font-name`, `--font-size`, `--bold`, `--color` |
| `ppt_replace_text.py` | Find/replace | `--find`, `--replace`, `--dry-run` |

## Content (5 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_add_bullet_list.py` | Bullet list | `--slide`, `--items` (comma-separated), `--position`, `--size` |
| `ppt_add_notes.py` | Speaker notes | `--slide`, `--text`, `--mode` (append/prepend/overwrite) |
| `ppt_add_connector.py` | Connect shapes | `--slide`, `--from-shape`, `--to-shape`, `--connector-type` (straight/elbow/curve) |
| `ppt_reposition_shape.py` | Move/resize shapes | `--slide`, `--shape`, `--position`, `--size` |
| `ppt_set_shape_text.py` | Update shape text | `--slide`, `--shape`, `--text` |
| `ppt_set_footer.py` | Footer config | `--text`, `--show-number`, `--show-date` |
| `ppt_set_background.py` | Background | `--slide` or `--all-slides`, `--color` or `--image` |

## Images (4 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_insert_image.py` | Insert image | `--slide`, `--image`, `--position`, `--size`, `--alt-text` |
| `ppt_replace_image.py` | Swap image | `--slide`, `--old-image`, `--new-image` |
| `ppt_crop_image.py` | Crop image | `--slide`, `--shape`, `--crop-box` |
| `ppt_set_image_properties.py` | Alt text | `--slide`, `--shape`, `--alt-text` |

## Charts (3 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_add_chart.py` | Add chart | `--slide`, `--chart-type`, `--data`, `--position`, `--size` |
| `ppt_update_chart_data.py` | Update chart | `--slide`, `--chart`, `--data` |
| `ppt_format_chart.py` | Style chart | `--slide`, `--chart`, `--title`, `--legend` |

Chart types: `column`, `column_stacked`, `bar`, `bar_stacked`, `line`, `line_markers`, `area`, `pie`, `doughnut`, `scatter`

## Tables (2 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_add_table.py` | Add table | `--slide`, `--rows`, `--cols`, `--position`, `--size`, `--data` |
| `ppt_format_table.py` | Style table | `--slide`, `--shape`, `--header-fill` |

## Layout (2 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_set_slide_layout.py` | Change layout | `--slide`, `--layout` |
| `ppt_set_title.py` | Set title | `--slide`, `--title`, `--subtitle` |

## Inspection (3 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_get_info.py` | Presentation metadata | `--file` |
| `ppt_get_slide_info.py` | Slide details (shapes, text) | `--file`, `--slide` |
| `ppt_capability_probe.py` | Deep probe (layouts, theme, fonts) | `--file`, `--deep`, `--timeout` |

## Export (2 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_export_pdf.py` | Export to PDF | `--file`, `--output` (requires LibreOffice) |
| `ppt_export_images.py` | Export slides as images | `--file`, `--output-dir`, `--format` (png/jpg) |

## Validation (5 tools)

| Tool | Purpose | Key Args |
|------|---------|----------|
| `ppt_validate_presentation.py` | Health check | `--file`, `--policy` (lenient/standard/strict) |
| `ppt_check_accessibility.py` | WCAG 2.1 audit | `--file` |
| `ppt_search_content.py` | Regex search | `--file`, `--query` |
| `ppt_json_adapter.py` | Validate JSON output | `--schema`, `--input` |
| `ppt_extract_notes.py` | Dump speaker notes | `--file` |

## Advanced (2 tools)

| Tool | Purpose | Key Args | Token? |
|------|---------|----------|--------|
| `ppt_merge_presentations.py` | Combine decks | `--sources` (JSON), `--output`, `--base-template` | **Yes** |
| `ppt_json_adapter.py` | Normalize output | `--schema`, `--input` | No |
