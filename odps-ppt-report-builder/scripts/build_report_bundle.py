from __future__ import annotations

import argparse
from datetime import date, datetime
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


def to_float(value: Any) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def parse_iso_date(value: Any) -> date | None:
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def detect_date_column(columns: list[str], rows: list[dict[str, Any]]) -> str:
    for column in columns:
        parsed = [parse_iso_date(row.get(column)) for row in rows if str(row.get(column, "")).strip()]
        if parsed and len(parsed) == len([row for row in rows if str(row.get(column, "")).strip()]):
            return column
    return columns[0] if columns else ""


def numeric_columns(columns: list[str], rows: list[dict[str, Any]], exclude: set[str] | None = None) -> list[str]:
    excluded = exclude or set()
    return [
        column
        for column in columns
        if column not in excluded and any(to_float(row.get(column)) is not None for row in rows)
    ]


def format_metric_value(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value))}"
    return f"{value:.2f}"


def derive_monthly_summary(query_name: str, item: dict[str, Any]) -> dict[str, Any] | None:
    rows = item.get("rows", [])
    columns = item.get("columns", [])
    if not rows or not columns:
        return None

    date_column = detect_date_column(columns, rows)
    metric_columns = numeric_columns(columns, rows, exclude={date_column})
    if not metric_columns:
        return None

    parsed_dates = [parse_iso_date(row.get(date_column)) for row in rows if parse_iso_date(row.get(date_column))]
    month_label = "本月累计"
    if parsed_dates:
        first_day = min(parsed_dates)
        month_label = f"{first_day.year}-{first_day.month:02d}月累计"

    summary_row: dict[str, str] = {"period": month_label}
    totals = {
        column: sum(to_float(row.get(column)) or 0.0 for row in rows)
        for column in metric_columns
    }
    for column, total in totals.items():
        summary_row[column] = format_metric_value(total)

    gmv = totals.get("gmv")
    order_cnt = totals.get("order_cnt")
    buyer_cnt = totals.get("buyer_cnt")
    if gmv is not None and buyer_cnt not in (None, 0):
        summary_row["arpu"] = f"{gmv / buyer_cnt:.2f}"
    if order_cnt is not None and buyer_cnt not in (None, 0):
        summary_row["freq"] = f"{order_cnt / buyer_cnt:.2f}"
    if gmv is not None and order_cnt not in (None, 0):
        summary_row["aov"] = f"{gmv / order_cnt:.2f}"

    summary_columns = ["period"] + metric_columns
    for derived in ["arpu", "freq", "aov"]:
        if derived in summary_row:
            summary_columns.append(derived)

    return {
        "title": f"{item.get('title', query_name)} Monthly Summary",
        "columns": summary_columns,
        "sql": f"derived::{query_name}__monthly_summary",
        "rows": [summary_row],
    }


def derive_weekly_trend(query_name: str, item: dict[str, Any]) -> dict[str, Any] | None:
    rows = item.get("rows", [])
    columns = item.get("columns", [])
    if not rows or not columns:
        return None

    date_column = detect_date_column(columns, rows)
    metric_columns = numeric_columns(columns, rows, exclude={date_column})
    if not metric_columns:
        return None

    dated_rows = []
    for row in rows:
        parsed = parse_iso_date(row.get(date_column))
        if parsed is None:
            continue
        dated_rows.append((parsed, row))
    if not dated_rows:
        return None

    dated_rows.sort(key=lambda item: item[0])
    month_start = date(dated_rows[0][0].year, dated_rows[0][0].month, 1)
    buckets: dict[int, dict[str, Any]] = {}
    for current_date, row in dated_rows:
        week_index = ((current_date - month_start).days // 7) + 1
        bucket = buckets.setdefault(
            week_index,
            {
                "week": f"第{week_index}周",
                "week_range": {"start": current_date, "end": current_date},
                **{column: 0.0 for column in metric_columns},
            },
        )
        bucket["week_range"]["start"] = min(bucket["week_range"]["start"], current_date)
        bucket["week_range"]["end"] = max(bucket["week_range"]["end"], current_date)
        for column in metric_columns:
            bucket[column] += to_float(row.get(column)) or 0.0

    derived_rows: list[dict[str, str]] = []
    for index in sorted(buckets):
        bucket = buckets[index]
        range_start = bucket["week_range"]["start"]
        range_end = bucket["week_range"]["end"]
        result_row: dict[str, str] = {
            "week": bucket["week"],
            "week_range": f"{range_start.month:02d}/{range_start.day:02d}-{range_end.month:02d}/{range_end.day:02d}",
        }
        for column in metric_columns:
            result_row[column] = format_metric_value(bucket[column])
        derived_rows.append(result_row)

    return {
        "title": f"{item.get('title', query_name)} Weekly Trend",
        "columns": ["week", *metric_columns, "week_range"],
        "sql": f"derived::{query_name}__weekly",
        "rows": derived_rows,
    }


def augment_monthly_queries(query_results: dict[str, Any]) -> dict[str, Any]:
    augmented = dict(query_results)
    for query_name, item in list(query_results.items()):
        monthly_summary = derive_monthly_summary(query_name, item)
        if monthly_summary:
            augmented[f"{query_name}__monthly_summary"] = monthly_summary
        weekly_trend = derive_weekly_trend(query_name, item)
        if weekly_trend:
            augmented[f"{query_name}__weekly"] = weekly_trend
    return augmented


def copy_optional_fields(slide: dict[str, Any], field_names: list[str]) -> dict[str, Any]:
    return {name: slide[name] for name in field_names if name in slide}


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
                    **copy_optional_fields(slide, ["template"]),
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
                    **copy_optional_fields(slide, ["template"]),
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
                    **copy_optional_fields(
                        slide,
                        ["template", "time_grain", "section_label", "display_mode", "support_query"],
                    ),
                }
            )
            continue

        if slide_type == "chart":
            slides.append(
                {
                    "type": "chart",
                    "title": slide["title"],
                    "query": slide["query"],
                    "chart_type": slide.get("chart_type", "line"),
                    "category_column": slide.get("category_column", ""),
                    "value_columns": slide.get("value_columns", []),
                    "max_points": slide.get("max_points", 12),
                    "takeaway_prompt": slide.get("takeaway_prompt", ""),
                    "takeaway_bullets": [],
                    "speaker_notes": "",
                    **copy_optional_fields(
                        slide,
                        ["template", "time_grain", "section_label", "display_mode"],
                    ),
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
                    **copy_optional_fields(slide, ["template", "section_label"]),
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

    query_results = augment_monthly_queries(query_results)

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

    slide_plan = build_slide_plan(rendered_config, output_dir)

    dump_json(payload_path, payload)
    dump_json(plan_path, slide_plan)
    write_text(context_path, build_context_markdown(payload, plan_path))

    print(f"payload={payload_path}")
    print(f"context={context_path}")
    print(f"plan_template={plan_path}")


if __name__ == "__main__":
    main()
