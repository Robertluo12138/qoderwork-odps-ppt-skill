---
name: odps-ppt-report-builder
description: Use this skill when the user wants to query ODPS or MaxCompute with fixed SQL, inspect source table schema, and generate a fixed-format PPT report or business review deck. This skill is for repeatable ODPS-to-PPT workflows that combine odpscmd, reusable SQL templates, and deterministic slide rendering for QoderWork, Qoder IDE, or Qoder CLI.
---

# ODPS PPT Report Builder

Use this skill for recurring reporting workflows such as weekly reviews, management updates, campaign summaries, business retrospectives, or standard operating reports that should be created from fixed ODPS queries and a fixed slide format.

## Preconditions

- The only hard requirement is `odpscmd` on the current machine.
- Do not assume Python, git, Node.js, or any other developer tooling exists.
- If Python is available, the bundled scripts can standardize the workflow, but they are optional.
- Prefer a team-owned private brief derived from `assets/company_private_brief.template.md`.
- Treat Qoder IDE and Qoder Work as separate installation targets. Do not assume a skill folder in one product is automatically visible in the other.

## Install Strategy

Use three separate layers. Do not mix them.

1. Skill files

- Install the skill folder into `~/.qoderwork/skills/odps-ppt-report-builder/`
- Do not install supporting packages into the Qoder Work application directory
- Qoder IDE and Qoder Work should each load their own skill directories

2. Shared user runtime

- If Python scripts are needed, install Python packages into a user-scoped shared runtime:
  - macOS: `~/.qoderwork-runtime/python/odps-ppt-report-builder/`
  - Windows: `%USERPROFILE%\.qoderwork-runtime\python\odps-ppt-report-builder\`
- This runtime is reusable across conversations and can also be reused by future Python-based skills
- Prefer this over `pip install --user` and prefer it over a venv inside the Qoder Work app folder

3. Machine-level tools

- Tools like `Python`, `odpscmd`, `Node.js`, `git`, `PowerPoint`, or `WPS` should be installed at the machine or user level, not inside the skill folder
- For future skills that need Node.js, install Node.js once for the whole machine or user profile
- Never install machine-level tools into the Qoder Work app bundle

## Setup Decision Tree

Follow this order exactly.

1. Confirm the product.

- If the user is in Qoder Work, use `~/.qoderwork/skills/`
- If the user is in Qoder IDE, do not assume the same folder applies

2. Confirm the operating system.

- macOS:
  - use `python3`
  - shared runtime under `~/.qoderwork-runtime/`
- Windows:
  - prefer `py -3`
  - shared runtime under `%USERPROFILE%\.qoderwork-runtime\`

3. Confirm which capabilities are already present.

- Check whether `odpscmd` exists
- Check whether Python exists
- Check whether the current machine can produce `.pptx`

4. Choose the execution path.

- If only `odpscmd` is available, use zero-setup mode and generate a slide manuscript if needed
- If `odpscmd` and Python are available, bootstrap the shared runtime and use the scripted path
- If `odpscmd` is missing, read the private brief for company-approved installation steps before proceeding

## When This Skill Triggers

Trigger this skill when the user asks for any of the following:

- Query one or more fixed ODPS tables and turn the result into a PPT
- Generate a recurring weekly or daily business deck from MaxCompute
- Inspect table schema before writing a management report
- Re-run a standard report with only date or business variables changed
- Produce a slide deck that must follow a fixed template, structure, tone, or narrative format

## Default Workflow: Zero-Setup Qoder Work Mode

This is the default mode for non-technical users. Prefer it unless the current machine clearly has Python available and the user wants the scripted path.

1. Confirm which private brief to use.

- If `assets/company_private_brief.md` exists, use it.
- Otherwise use `assets/company_private_brief.template.md` as the source of truth and ask the user to fill the private sections on the company machine.
- The private brief is where table names, SQL, schema rules, confidentiality constraints, and PPT writing requirements belong.

2. Avoid asking the user to install tooling.

- Use only:
  - `odpscmd`
  - built-in shell execution
  - Qoder Work reasoning and file editing
- If the machine has no Python, continue without the bundled scripts.

3. Read the private brief before touching ODPS.

- Extract these items from the private brief:
  - exact schema collection command or SQL
  - exact business queries
  - required variables such as date range, campaign name, owner, or business line
  - slide-by-slide PPT rules
  - words or conclusions that are forbidden

4. Run ODPS directly with the exact commands from the private brief.

- Save raw outputs into a task folder, for example:
  - `raw_schema.txt`
  - `raw_query_overview.txt`
  - `raw_query_detail.txt`
- Also create normalized reasoning files:
  - `report_context.md`
  - `slide_plan.generated.json`

5. Draft the report in the private brief's fixed format.

- Keep the slide order exactly as defined
- Keep bullet counts exactly as defined
- Make every claim traceable to the ODPS output
- If a query returns no rows, state that explicitly on the slide instead of guessing

6. Produce the final deliverable.

- Preferred:
  - a native `.pptx` if the current environment supports it
- Acceptable fallback:
  - a complete slide manuscript with slide titles, bullets, table content, speaker notes, and appendix content ready to paste into PowerPoint or WPS

7. Return the result clearly.

- Provide the file path
- Summarize the key findings in 3 to 5 lines
- Call out missing data, failed schema fetches, or blocked rendering steps

## Optional Workflow: Scripted Mode

Use this only when Python is available and the user wants a more repeatable pipeline.

1. Start from `assets/report_config.template.yaml`
2. Bootstrap the shared runtime

- macOS:
  ```bash
  python3 scripts/bootstrap_env.py
  ```
- Windows:
  ```powershell
  py -3 scripts\bootstrap_env.py
  ```

3. Build the report bundle

- macOS:
  ```bash
  ~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/bin/python scripts/build_report_bundle.py --config <config-path>
  ```
- Windows:
  ```powershell
  %USERPROFILE%\.qoderwork-runtime\python\odps-ppt-report-builder\venv\Scripts\python.exe scripts\build_report_bundle.py --config <config-path>
  ```

4. Generate a first-pass slide plan

- Preferred:
  - let Qoder Work read `report_context.md` and fill the plan
- Fallback for weaker models:
  - run `scripts/fill_slide_plan.py` to create a deterministic starter plan

5. Render the output

- Run `scripts/render_ppt.py` for `.pptx`
- Run `scripts/render_preview_html.py` if the user wants a browser preview
- Run `scripts/render_preview_png.py` if the user wants a shareable image preview without opening PowerPoint
- If native PPT rendering is blocked, return the generated plan and slide manuscript

## Guardrails

- Keep the deck factual. Do not infer business causes unless the data supports it.
- Prefer the schema query over guessing field meanings.
- If `Information_Schema.columns` is unavailable in the target environment, use a custom `schema_sql` in the config instead of parsing `DESC` output.
- In zero-setup mode, do not force non-technical users into a Python-based flow.
- Put confidential SQL, table names, field definitions, business rules, and writing requirements into the private brief instead of hardcoding them into the public skill body.
- Install reusable dependencies into the shared user runtime, not into the Qoder Work app directory and not into system site-packages unless the tool is intentionally machine-wide.
- Keep reusable logic in the config and scripts, not in ad hoc one-off prompts.
- Treat the config file as the source of truth for SQL, slide order, and formatting rules.

## References

- Read `references/workflow.md` for config conventions and ODPS caveats.
- Read `references/slide_plan_format.md` before editing `slide_plan.generated.json`.
- Read `references/setup_strategy_zh.md` for detailed setup guidance and cross-platform deployment decisions.
- Read `assets/company_private_brief.template.md` when setting up a new confidential reporting workflow.
- Use `assets/qoderwork_first_run_prompt.template.md` as the starter prompt for non-technical Qoder Work users.
