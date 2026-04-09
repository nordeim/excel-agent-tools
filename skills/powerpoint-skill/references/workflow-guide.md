# Workflow Guide

## Workflow 1: Create Presentation from Scratch

```bash
# Step 1: Create blank deck
uv run tools/ppt_create_new.py --output work.pptx --json

# Step 2: Probe capabilities
uv run tools/ppt_capability_probe.py --file work.pptx --deep --json > probe.json
# Read probe.json to get available layouts, theme colors, fonts

# Step 3: Add slides
uv run tools/ppt_add_slide.py --file work.pptx --layout "Title Slide" --json
uv run tools/ppt_add_slide.py --file work.pptx --layout "Title and Content" --json
uv run tools/ppt_add_slide.py --file work.pptx --layout "Blank" --json

# Step 4: Populate content
uv run tools/ppt_set_title.py --file work.pptx --slide 0 --title "Title" --subtitle "Subtitle" --json
uv run tools/ppt_add_text_box.py --file work.pptx --slide 1 --text "Content here" --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"50%"}' --json

# Step 5: Add visuals
uv run tools/ppt_add_shape.py --file work.pptx --slide 2 --shape-type rectangle --position '{"left":"5%","top":"10%"}' --size '{"width":"90%","height":"80%"}' --fill-color "#F0F0F0" --fill-opacity 0.5 --json
# REFRESH INDICES after structural change
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json

# Step 6: Speaker notes
uv run tools/ppt_add_notes.py --file work.pptx --slide 0 --text "Opening remarks" --mode overwrite --json

# Step 7: Validate
uv run tools/ppt_validate_presentation.py --file work.pptx --policy standard --json
uv run tools/ppt_check_accessibility.py --file work.pptx --json
```

## Workflow 2: Edit Existing Presentation

```bash
# Step 1: CLONE (mandatory)
uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx --json

# Step 2: Inspect
uv run tools/ppt_get_info.py --file work.pptx --json
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json

# Step 3: Make changes
uv run tools/ppt_replace_text.py --file work.pptx --find "Old" --replace "New" --json

# Step 4: Validate
uv run tools/ppt_validate_presentation.py --file work.pptx --policy standard --json
```

## Workflow 3: Delete Slide (Destructive)

```bash
# Step 1: Clone
uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx --json

# Step 2: Generate token
TOKEN=$(python scripts/generate_token.py --scope "slide:delete:2" --quiet)

# Step 3: Delete with token
uv run tools/ppt_delete_slide.py --file work.pptx --slide 2 --approval-token "$TOKEN" --json

# Step 4: Verify
uv run tools/ppt_get_info.py --file work.pptx --json
```

## Workflow 4: Add Overlay Background

```bash
# Step 1: Add overlay shape
uv run tools/ppt_add_shape.py --file work.pptx --slide 0 --shape-type rectangle --position '{"left":"0%","top":"0%"}' --size '{"width":"100%","height":"100%"}' --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# Step 2: REFRESH INDICES
SLIDE_INFO=$(uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json)
SHAPE_COUNT=$(echo "$SLIDE_INFO" | jq '.shape_count')
OVERLAY_INDEX=$((SHAPE_COUNT - 1))

# Step 3: Send to back
uv run tools/ppt_set_z_order.py --file work.pptx --slide 0 --shape "$OVERLAY_INDEX" --action send_to_back --json

# Step 4: REFRESH INDICES again
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json
```

## Workflow 5: Add Chart

```bash
# Step 1: Add chart
uv run tools/ppt_add_chart.py --file work.pptx --slide 1 --chart-type column --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"60%"}' --json

# Step 2: REFRESH INDICES
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 1 --json

# Step 3: Format chart
uv run tools/ppt_format_chart.py --file work.pptx --slide 1 --chart 0 --title "Revenue" --legend bottom --json
```

## Workflow 6: Merge Presentations

```bash
# Step 1: Generate token
TOKEN=$(python scripts/generate_token.py --scope "merge:presentations:2" --quiet)

# Step 2: Merge
uv run tools/ppt_merge_presentations.py \
  --sources '[{"file":"part1.pptx","slides":"all"},{"file":"part2.pptx","slides":[0,1,2]}]' \
  --output merged.pptx \
  --approval-token "$TOKEN" \
  --json
```

## Workflow 7: Accessibility Remediation

```bash
# Step 1: Run audit
uv run tools/ppt_check_accessibility.py --file work.pptx --json > audit.json

# Step 2: Fix missing alt text
uv run tools/ppt_set_image_properties.py --file work.pptx --slide 0 --shape 1 --alt-text "Description of image" --json

# Step 3: Re-validate
uv run tools/ppt_check_accessibility.py --file work.pptx --json
```

## Workflow 8: Export

```bash
# Export to PDF (requires LibreOffice)
uv run tools/ppt_export_pdf.py --file work.pptx --output final.pdf --json

# Export slides as images
uv run tools/ppt_export_images.py --file work.pptx --output-dir slides/ --format png --json

# Extract speaker notes
uv run tools/ppt_extract_notes.py --file work.pptx --json > notes.json
```
