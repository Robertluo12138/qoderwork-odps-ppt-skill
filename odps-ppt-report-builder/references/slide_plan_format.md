# Slide Plan Format

`slide_plan.template.json` and `slide_plan.generated.json` use the same structure.

Top-level fields:

- `report_title`: display title for the deck
- `report_subtitle`: display subtitle for the deck
- `generated_at`: ISO timestamp from the extraction step
- `slides`: ordered array of slide objects

Supported slide types:

## `title`

Required fields:

- `type`
- `title`
- `subtitle`

## `schema`

Required fields:

- `type`
- `title`
- `tables`

Optional authoring fields:

- `intro_text`
- `speaker_notes`

## `table`

Required fields:

- `type`
- `title`
- `query`

Optional authoring fields:

- `takeaway_prompt`
- `takeaway_bullets`
- `speaker_notes`
- `max_rows`

## `chart`

Required fields:

- `type`
- `title`
- `query`

Optional authoring fields:

- `chart_type`
- `category_column`
- `value_columns`
- `max_points`
- `takeaway_prompt`
- `takeaway_bullets`
- `speaker_notes`

## `ai_bullets`

Required fields:

- `type`
- `title`
- `prompt`

Optional authoring fields:

- `subtitle`
- `bullets`
- `speaker_notes`
- `min_bullets`
- `max_bullets`

## Editing Rules

- Do not rename `query` values.
- Do not remove slides.
- Keep slide order unchanged.
- Fill only the narrative fields unless the user asks for structural edits.
- Keep bullets as plain strings.
- Prefer native PowerPoint objects such as text boxes, tables, charts, and shapes over pasted screenshots.
