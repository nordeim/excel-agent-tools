# Safety Protocols

## 1. Clone-Before-Edit (MANDATORY)

Never modify source files. Always create a working copy:

```bash
uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx --json
```

## 2. Probe-Before-Operate

Always inspect structure before modifying:

```bash
# Full probe (layouts, theme, fonts, capabilities)
uv run tools/ppt_capability_probe.py --file work.pptx --deep --json

# Quick info (slide count, version, dimensions)
uv run tools/ppt_get_info.py --file work.pptx --json

# Slide details (shapes, text, indices)
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json
```

## 3. Approval Tokens (Destructive Operations)

Required for: `ppt_delete_slide.py`, `ppt_remove_shape.py`, `ppt_merge_presentations.py`

```bash
# Generate token
python scripts/generate_token.py --scope "slide:delete:2"

# Use token
uv run tools/ppt_delete_slide.py --file work.pptx --slide 2 --approval-token "TOKEN_HERE" --json
```

Scope patterns:
- `slide:delete:<index>` — delete slide at index
- `shape:remove:<slide>:<shape>` — remove shape at index
- `merge:presentations:<count>` — merge N presentations

## 4. Version Tracking

Capture before/after every mutation to detect concurrent modifications:

```bash
BEFORE=$(uv run tools/ppt_get_info.py --file work.pptx --json | jq -r '.presentation_version')
# ... perform operations ...
AFTER=$(uv run tools/ppt_get_info.py --file work.pptx --json | jq -r '.presentation_version')
[ "$BEFORE" = "$AFTER" ] && echo "WARNING: Version unchanged — operation may have failed"
```

## 5. Index Refresh (CRITICAL)

Shape indices shift after structural operations. **ALWAYS** refresh after:

- `ppt_add_shape.py` — adds index at end
- `ppt_remove_shape.py` — shifts indices down
- `ppt_set_z_order.py` — reorders indices via XML
- `ppt_delete_slide.py` — invalidates all slide indices
- `ppt_merge_presentations.py` — completely restructures

```bash
# After any structural change:
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json
```

## 6. Output Hygiene

All tools suppress stderr. Always use `--json` flag. Parse output with `jq`:

```bash
SLIDE_INDEX=$(uv run tools/ppt_add_slide.py --file work.pptx --layout "Blank" --json | jq -r '.slide_index')
```

## 7. Recovery Protocol

If corruption detected:

```bash
# 1. Backup current state
cp work.pptx "work_$(date +%Y%m%d_%H%M%S).pptx"

# 2. Re-clone from source
uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx --json

# 3. Clear stale locks (>5 min old)
find . -name "*.lock" -mmin +5 -delete
```
