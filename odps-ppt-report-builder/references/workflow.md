# Workflow Notes

## Recommended Team Setup

- Keep the reusable skill package in a Git repository under `.qoder/skills/odps-ppt-report-builder/`
- When you install into Qoder Work, copy the skill folder itself into `~/.qoderwork/skills/`
- If your company also uses Qoder IDE, treat that as a separate install target
- Keep script dependencies outside the app folder in a shared per-user runtime
- Give each team or report line its own config file, for example:
  - `report_config.sales.yaml`
  - `report_config.marketing.yaml`
  - `report_config.risk.yaml`
- Only the config should change between report variants whenever possible

## Zero-Setup Mode

- This should be the default mode for operators and non-technical teammates.
- In this mode, the skill should rely only on:
  - Qoder Work
  - `odpscmd`
  - filled company private brief
- Do not require users to prepare Python, pip, git, or local package managers if the task can be completed without them.
- The most important file in this mode is `assets/company_private_brief.template.md`, which becomes `assets/company_private_brief.md` after the user fills it on the company machine.
- If a native `.pptx` cannot be rendered, the skill should still output a complete slide manuscript and a JSON slide plan.

## Weak-Model Guardrails

- On a weak model, always check environment first and report one result at a time.
- Do not mix macOS and Windows instructions in the same answer.
- If the user gives only a natural-language request like "generate Chagee March PPT", map it into:
  - business
  - date range
  - report type
- If year is omitted, default to the current year and say that assumption explicitly.
- If only one required field is missing, ask only one short question.
- If a required file like `company_private_brief.md` is missing, stop and give the exact next step rather than a general explanation.

## Scripted Mode

- Use scripted mode only when the current machine already has Python or the user explicitly wants a repeatable automation path.
- The YAML config and Python scripts are helpers for standardization, not the only way to use the skill.
- The runtime should default to:
  - macOS: `~/.qoderwork-runtime/python/odps-ppt-report-builder/`
  - Windows: `%USERPROFILE%\.qoderwork-runtime\python\odps-ppt-report-builder\`
- Do not install Python packages into the Qoder Work application folder.

## ODPS Notes

- The official MaxCompute documentation states that generic `odpscmd` output is not guaranteed to stay backward compatible across versions.
- To make shell parsing safer, this skill always sets `odps.sql.select.output.format` and expects fixed columns from the config.
- If your local client still prints extra logs into stdout, check whether `use_instance_tunnel=false` is configured in the local `odps_config.ini`.
- If schema lookup through `Information_Schema.columns` is blocked or delayed, add a custom `schema_sql` to the config and return the same columns expected by the script:
  - `table_schema`
  - `table_name`
  - `column_name`
  - `ordinal_position`
  - `data_type`
  - `is_partition_key`
  - `column_comment`

## Authoring Rules

- Table slides should stay close to the raw data and avoid interpretation drift.
- Narrative slides should be short and decision-oriented.
- When a chart is needed, prefer a native PowerPoint chart instead of a screenshot.
- When decoration is needed, prefer editable shapes and text boxes instead of flattening content into one image.
- If a slide prompt says "exactly 3 bullets", do not output 2 or 4.
- If a query returns no rows, write a bullet that says the result set is empty and avoid extrapolation.

## Suggested Delivery Pattern

1. Run the ODPS extraction bundle.
2. Read the Markdown brief.
3. Fill the generated slide plan.
4. Render the PPT.
5. Return the file path plus a concise summary.
