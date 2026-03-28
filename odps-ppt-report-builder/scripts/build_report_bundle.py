from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from common import (
    build_context,
    deep_render,
    default_schema_sql,
    dump_json,
    group_schema_rows,
    load_yaml,
    output_dir_from_config,
    parse_key_value_pairs,
    parse_tabular_stdout,
    rows_to_markdown,
    run_odps_sql,
    schema_columns,
    utc_timestamp,
    write_text,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a Qoder ODPS report bundle.")
    parser.add_argument("--config", required=True, help="Path to the YAML report config.")
    parser.add_argument(
        "--var",
        action="append",
        default=[],
        help="Template variable override in KEY=VALUE form. Can be provided multiple times.",
    )
    return parser.parse_args()


def build_slide_plan(config: dict[str, Any], output_dir: Path) -> dict[str, Any]:
    report = config.get("report", {})
    slides = []

    for slide in config.get("slides", []):
        slide_type = slide.get("type")
        if slide_type == "title":
            slides.append(
                {
                    "type": "title",
                    "title": slide.get("title", report.get("title", "")),
                    "subtitle": slide.get("subtitle", report.get("subtitle", "")),
                }
            )
            continue

        if slide_type == "schema":
            slides.append(
                {
                    "type": "schema",
                    "title": slide["title"],
                    "tables": slide.get("tables", []),
                    "intro_text": "",
                    "speaker_notes": "",
                }
            )
            continue

        if slide_type == "table":
            slides.append(
                {
                    "type": "table",
                    "title": slide["title"],
                    "query": slide["query"],
                    "max_rows": slide.get("max_rows", 12),
                    "takeaway_prompt": slide.get("takeaway_prompt", ""),
                    "takeaway_bullets": [],
                    "speaker_notes": "",
                }
            )
            continue

        if slide_type == "ai_bullets":
            slides.append(
                {
                    "type": "ai_bullets",
                    "title": slide["title"],
                    "subtitle": slide.get("subtitle", ""),
                    "prompt": slide["prompt"],
                    "min_bullets": slide.get("min_bullets", 3),
                    "max_bullets": slide.get("max_bullets", 5),
                    "bullets": [],
                    "speaker_notes": "",
                }
            )
            continue

        raise ValueError(f"Unsupported slide type: {slide_type}")

    return {
        "report_title": report.get("title", ""),
        "report_subtitle": report.get("subtitle", ""),
        "generated_at": utc_timestamp(),
        "output_dir": str(output_dir),
        "slides": slides,
    }


def build_context_markdown(payload: dict[str, Any], slide_plan_path: Path) -> str:
    metadata = payload["metadata"]
    lines = ["# Report Context", ""]
    lines.append("## Metadata")
    lines.append("")
    for key, value in metadata.items():
        if key == "output_files":
            continue
        lines.append(f"- {key}: {value}")
    lines.append("")

    lines.append("## Output Files")
    lines.append("")
    for key, value in metadata["output_files"].items():
        lines.append(f"- {key}: {value}")
    lines.append("")

    lines.append("## Source Tables")
    lines.append("")
    schema_by_table = payload["tables"]
    if not schema_by_table:
        lines.append("No schema rows returned.")
        lines.append("")
    else:
        for table_name, columns in schema_by_table.items():
            lines.append(f"### {table_name}")
            lines.append("")
            if not columns:
                lines.append("No columns returned.")
                lines.append("")
                continue
            lines.append(
                "| column_name | data_type | is_partition_key | column_comment |"
            )
            lines.append("| --- | --- | --- | --- |")
            for row in columns:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            row.get("column_name", ""),
                            row.get("data_type", ""),
                            row.get("is_partition_key", ""),
                            row.get("column_comment", "").replace("\n", " "),
                        ]
                    )
                    + " |"
                )
            lines.append("")

    lines.append("## Query Results")
    lines.append("")
    for query_name, query_result in payload["queries"].items():
        lines.append(rows_to_markdown(query_name, query_result["rows"], limit=10))

    lines.append("## Authoring Step")
    lines.append("")
    lines.append(
        f"Read `{slide_plan_path.name}` and save the completed version as `slide_plan.generated.json` in the same directory."
    )
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    args = parse_args()
    config_path = Path(args.config).resolve()
    base_config = load_yaml(config_path)
    overrides = parse_key_value_pairs(args.var)
    context = build_context(base_config, overrides)
    rendered_config = deep_render(base_config, context)

    output_dir = output_dir_from_config(config_path, rendered_config)
    output_dir.mkdir(parents=True, exist_ok=True)

    odps_config = rendered_config.get("odps", {})
    field_delim = odps_config.get("field_delim", "\t")
    schema_rows: list[dict[str, str]] = []
    schema_sql = rendered_config.get("schema_sql") or default_schema_sql(rendered_config)
    if schema_sql:
        schema_run = run_odps_sql(schema_sql, odps_config)
        if schema_run.returncode != 0:
            raise RuntimeError(
                "Schema query failed.\n"
                f"stdout:\n{schema_run.stdout}\n\nstderr:\n{schema_run.stderr}"
            )
        schema_rows = parse_tabular_stdout(schema_run.stdout, schema_columns(), field_delim)

    grouped_schema = group_schema_rows(schema_rows)

    query_results: dict[str, Any] = {}
    for query in rendered_config.get("queries", []):
        columns = query.get("columns")
        if not columns:
            raise ValueError(f"Query {query.get('name')} must define columns.")
        query_run = run_odps_sql(query["sql"], odps_config)
        if query_run.returncode != 0:
            raise RuntimeError(
                f"Query {query['name']} failed.\n"
                f"stdout:\n{query_run.stdout}\n\nstderr:\n{query_run.stderr}"
            )
        rows = parse_tabular_stdout(query_run.stdout, columns, field_delim)

        query_results[query["name"]] = {
            "title": query.get("title", query["name"]),
            "columns": columns,
            "sql": query["sql"],
            "rows": rows,
        }

    payload_path = output_dir / "report_payload.json"
    plan_path = output_dir / "slide_plan.template.json"
    context_path = output_dir / "report_context.md"

    payload = {
        "metadata": {
            "generated_at": utc_timestamp(),
            "config_path": str(config_path),
            "report_title": rendered_config.get("report", {}).get("title", ""),
            "report_subtitle": rendered_config.get("report", {}).get("subtitle", ""),
            "variables": context,
            "output_files": {
                "payload": str(payload_path),
                "context": str(context_path),
                "slide_plan_template": str(plan_path),
            },
        },
        "tables": grouped_schema,
        "queries": query_results,
    }

    slide_plan = build_slidePlan(rendered_config, output_dir)

    dump_json(payload_path, payload)
    dump_json(plan_path, slide_plan)
    write_text(context_path, build_context_markdown(payload, plan_path))

    print(f"payload={payload_path}")
    print(f"context={context_path}")
    print(f"plan_template={plan_path}")


if __name__ == "__main__":
    main()
