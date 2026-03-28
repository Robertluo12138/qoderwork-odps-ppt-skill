# Company Private Brief

Fill this file only on the company machine. This is the main place to put confidential ODPS details that should not live in the public skill body.

## 1. Workflow Mode

- preferred_mode: `qoderwork_only` or `scripted`
- target_product: `qoderwork`
- output_preference: `native_pptx` or `slide_manuscript_ok`
- report_owner:
- audience:
- default_language:

## 1.1 Natural-Language Mapping Rules

Use this section to help a weak Qoder Work model understand short user requests.

- default_report_type:
- default_date_rule_if_user_only_says_month:
- default_date_rule_if_user_only_says_last_week:
- default_date_rule_if_user_only_says_yesterday:
- if_year_is_missing_use_current_year: `yes` or `no`

### Business Aliases

- alias_1:
  - user_may_say:
  - normalize_to:
- alias_2:
  - user_may_say:
  - normalize_to:

### Example Requests

- example_1: `帮我生成霸王茶姬三月的 PPT`
  - interpret_as:
- example_2:
  - interpret_as:

## 2. ODPS Environment

- odpscmd_binary:
- default_project:
- endpoint:
- login_notes:
- machine_setup_policy:
  - qoderwork_skill_install_path:
  - shared_runtime_root_mac:
  - shared_runtime_root_windows:
  - machine_level_tools_allowed:
  - machine_level_tools_forbidden:
- python_policy:
  - minimum_version:
  - install_source:
  - whether_python_is_preinstalled:
- odpscmd_install_steps_mac:
  1.
  2.
- odpscmd_install_steps_windows:
  1.
  2.
- office_or_wps_notes:
  - preferred_renderer:
  - allowed_fallback:
- required_runtime_variables:
  - start_date:
  - end_date:
  - biz_line:
  - owner:

## 3. Schema Collection Rules

Paste the exact schema lookup SQL or command here. If multiple tables are involved, list them explicitly.

### Source Tables

- table_1:
  - project:
  - table_name:
  - business_meaning:
- table_2:
  - project:
  - table_name:
  - business_meaning:

### Schema SQL

```sql
-- Paste the exact schema SQL here.
-- Return fields that let Qoder understand the structure clearly.
```

### Schema Notes

- Which fields are partitions:
- Which fields are metrics:
- Which fields are dimensions:
- Which fields are sensitive and must not appear in the final PPT:

## 4. Fixed Data Queries

Define every recurring query separately.

### Query 1

- name:
- purpose:
- output_columns:
  - column_1:
  - column_2:
- how_to_read_this_query:

```sql
-- Paste the exact SQL here.
```

### Query 2

- name:
- purpose:
- output_columns:
  - column_1:
  - column_2:
- how_to_read_this_query:

```sql
-- Paste the exact SQL here.
```

## 5. Writing Rules For Qoder

- Tone:
- Style:
- Must include:
- Must avoid:
- Do not mention:
- Must highlight:
- Must not infer:
- Data validation rules:

## 6. PPT Outline

Define the exact slide order here.

### Slide 1

- title:
- slide_type: `title` / `bullets` / `table` / `chart` / `appendix`
- purpose:
- required_content:
- source_query:
- fixed_format_rules:
- exact_bullet_count:
- speaker_note_rules:

### Slide 2

- title:
- slide_type: `bullets` / `table` / `chart` / `appendix`
- purpose:
- required_content:
- source_query:
- fixed_format_rules:
- exact_bullet_count:
- speaker_note_rules:

## 7. Delivery Rules

- final_file_name_rule:
- save_folder_rule:
- appendix_required: `yes` or `no`
- watermark_or_label_rule:
- output_priority:
  1. native `.pptx`
  2. html preview
  3. slide manuscript
- review_checklist:
  - no fabricated conclusions
  - no sensitive fields exposed
  - slide order unchanged
  - all numbers trace back to ODPS output
