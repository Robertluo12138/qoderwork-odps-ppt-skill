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

- `template`
- `intro_text`
- `speaker_notes`

## `table`

Required fields:

- `type`
- `title`
- `query`

Optional authoring fields:

- `template`
- `support_query`: optional secondary query for overview support strips
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

- `template`
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

- `template`
- `subtitle`
- `bullets`
- `speaker_notes`
- `min_bullets`
- `max_bullets`

## Suggested Template Names

- `cover`: management-report cover page
- `overview`: month-level opening business overview page
- `kpi_summary`: month-level KPI summary page
- `weekly_trend`: month trend page grouped by week
- `comparison`: month comparison / evidence page
- `time_slot`: time-slot structure page
- `price_band`: price-band structure page
- `time_price_matrix`: time-slot x price-band cross page
- `city_distribution`: city contribution page
- `aoi_distribution`: AOI structure page
- `new_old_mix`: new vs returning customer structure page
- `cluster_mix`: cluster detail page
- `conclusion`: monthly conclusion and next-action page
- `thanks`: clean closing / thank-you page
- `appendix`: appendix / daily auxiliary data page

## Editing Rules

- Do not rename `query` values.
- Do not remove slides.
- Keep slide order unchanged.
- Fill only the narrative fields unless the user asks for structural edits.
- Keep bullets as plain strings.
- Prefer native PowerPoint objects such as text boxes, tables, charts, and shapes over pasted screenshots.
