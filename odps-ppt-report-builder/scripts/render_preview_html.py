from __future__ import annotations

import argparse
import json
from html import escape
from pathlib import Path
from typing import Any

from common import deep_render, load_yaml


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render an HTML preview for the generated slide deck.")
    parser.add_argument("--config", required=True, help="Path to the YAML report config.")
    parser.add_argument("--payload", required=True, help="Path to report_payload.json.")
    parser.add_argument("--plan", required=True, help="Path to slide_plan.generated.json.")
    parser.add_argument("--output", required=True, help="Path to preview HTML file.")
    return parser.parse_args()


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def to_float(value: Any) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def render_table(headers: list[str], rows: list[dict[str, Any]], limit: int) -> str:
    shown = rows[:limit]
    head = "".join(f"<th>{escape(str(header))}</th>" for header in headers)
    body_rows = []
    for row in shown:
        body_rows.append(
            "<tr>" + "".join(f"<td>{escape(str(row.get(header, '')))}</td>" for header in headers) + "</tr>"
        )
    if not body_rows:
        body_rows.append(f"<tr><td colspan='{len(headers)}'>No rows returned.</td></tr>")
    return f"<table><thead><tr>{head}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"


def render_chart_preview(slide: dict[str, Any], payload: dict[str, Any]) -> str:
    query_data = payload.get("queries", {}).get(slide.get("query"), {})
    rows = query_data.get("rows", [])
    columns = query_data.get("columns", [])
    category_column = slide.get("category_column") or (columns[0] if columns else "")
    value_columns = slide.get("value_columns") or [
        column
        for column in columns
        if column != category_column
        and any(to_float(row.get(column)) is not None for row in rows)
    ][:3]

    if not rows or not category_column or not value_columns:
        return "<p class='notes'>Chart preview is not available for this slide yet.</p>"

    chart_rows = rows[: int(slide.get("max_points", 10))]
    first_metric = value_columns[0]
    values = [to_float(row.get(first_metric)) or 0.0 for row in chart_rows]
    max_value = max(values) if values else 0.0
    bars = []
    for row, value in zip(chart_rows, values):
        width = 0 if max_value <= 0 else value / max_value * 100
        bars.append(
            "<div class='chart-row'>"
            f"<span class='chart-label'>{escape(str(row.get(category_column, '')))}</span>"
            "<div class='chart-track'>"
            f"<div class='chart-bar' style='width:{width:.1f}%'></div>"
            "</div>"
            f"<span class='chart-value'>{escape(str(row.get(first_metric, '')))}</span>"
            "</div>"
        )
    return (
        f"<p class='chart-meta'>Chart: {escape(str(slide.get('chart_type', 'line')))} | "
        f"Category: {escape(str(category_column))} | Series: {escape(', '.join(value_columns))}</p>"
        + "".join(bars)
    )


def render_slide(slide: dict[str, Any], payload: dict[str, Any]) -> str:
    slide_type = slide.get("type")
    title = escape(str(slide.get("title", "")))
    subtitle = escape(str(slide.get("subtitle", "")))
    parts = [f"<section class='slide'><div class='accent'></div><h2>{title}</h2>"]
    if subtitle:
        parts.append(f"<p class='subtitle'>{subtitle}</p>")

    if slide_type == "title":
        parts.append(f"<div class='hero'><h1>{title}</h1><p>{subtitle}</p></div>")
    elif slide_type == "schema":
        intro = escape(str(slide.get("intro_text", "")))
        if intro:
            parts.append(f"<p>{intro}</p>")
        for table_name in slide.get("tables", []):
            rows = payload.get("tables", {}).get(table_name, [])
            headers = ["column_name", "data_type", "is_partition_key", "column_comment"]
            parts.append(f"<h3>{escape(str(table_name))}</h3>")
            parts.append(render_table(headers, rows, 8))
    elif slide_type == "table":
        bullets = slide.get("takeaway_bullets", [])
        if bullets:
            parts.append("<ul>" + "".join(f"<li>{escape(str(item))}</li>" for item in bullets) + "</ul>")
        query_data = payload.get("queries", {}).get(slide.get("query"), {})
        headers = query_data.get("columns", [])
        parts.append(render_table(headers, query_data.get("rows", []), int(slide.get("max_rows", 10))))
    elif slide_type == "chart":
        bullets = slide.get("takeaway_bullets", [])
        if bullets:
            parts.append("<ul>" + "".join(f"<li>{escape(str(item))}</li>" for item in bullets) + "</ul>")
        parts.append(render_chart_preview(slide, payload))
    elif slide_type == "ai_bullets":
        bullets = slide.get("bullets", [])
        parts.append("<ul>" + "".join(f"<li>{escape(str(item))}</li>" for item in bullets) + "</ul>")

    notes = escape(str(slide.get("speaker_notes", "")))
    if notes:
        parts.append(f"<p class='notes'>{notes}</p>")
    parts.append("</section>")
    return "".join(parts)


def main() -> None:
    args = parse_args()
    config = load_yaml(Path(args.config).resolve())
    payload = load_json(Path(args.payload).resolve())
    plan = load_json(Path(args.plan).resolve())
    context = payload.get("metadata", {}).get("variables", {})
    rendered_config = deep_render(config, context)
    theme = rendered_config.get("report", {}).get("theme", {})

    html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{escape(plan.get('report_title', 'Report Preview'))}</title>
  <style>
    :root {{
      --primary: #{theme.get('primary_color', 'FF6200')};
      --accent: #{theme.get('accent_color', 'FFA020')};
      --text: #{theme.get('text_color', '1F1F1F')};
      --bg: #{theme.get('background_color', 'FFFFFF')};
      --font: {theme.get('font_family', 'Arial')}, sans-serif;
    }}
    body {{
      margin: 0;
      padding: 24px;
      font-family: var(--font);
      color: var(--text);
      background: linear-gradient(180deg, #f5f7fb 0%, #eef2f8 100%);
    }}
    .deck {{
      display: grid;
      gap: 24px;
    }}
    .slide {{
      position: relative;
      background: var(--bg);
      border-radius: 20px;
      padding: 28px 28px 24px;
      box-shadow: 0 16px 40px rgba(255, 98, 0, 0.10);
      overflow: hidden;
    }}
    .accent {{
      position: absolute;
      left: 0;
      top: 0;
      width: 100%;
      height: 8px;
      background: linear-gradient(90deg, var(--primary), var(--accent));
    }}
    h1, h2, h3 {{
      margin: 0 0 12px;
    }}
    .subtitle {{
      color: #5b6472;
      margin-bottom: 16px;
    }}
    .hero {{
      min-height: 180px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      border-left: 8px solid var(--accent);
      padding-left: 20px;
      margin-top: 16px;
      background: linear-gradient(135deg, rgba(255, 98, 0, 0.05), rgba(255, 160, 32, 0.08));
      border-radius: 16px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
      font-size: 14px;
    }}
    th {{
      background: var(--primary);
      color: white;
      text-align: left;
      padding: 10px;
    }}
    td {{
      border-bottom: 1px solid #e7ebf2;
      padding: 10px;
      vertical-align: top;
    }}
    ul {{
      margin: 12px 0 18px;
      padding-left: 22px;
    }}
    li {{
      margin-bottom: 8px;
    }}
    .notes {{
      margin-top: 16px;
      padding-top: 12px;
      border-top: 1px dashed #d4dbe6;
      color: #667085;
      font-size: 13px;
    }}
    .chart-meta {{
      margin: 12px 0 16px;
      color: #5b6472;
    }}
    .chart-row {{
      display: grid;
      grid-template-columns: 180px 1fr 80px;
      gap: 12px;
      align-items: center;
      margin-bottom: 10px;
    }}
    .chart-label, .chart-value {{
      font-size: 14px;
    }}
    .chart-track {{
      height: 16px;
      background: #edf1f6;
      border-radius: 999px;
      overflow: hidden;
    }}
    .chart-bar {{
      height: 100%;
      background: linear-gradient(90deg, var(--primary), var(--accent));
    }}
  </style>
</head>
<body>
  <div class="deck">
    {''.join(render_slide(slide, payload) for slide in plan.get('slides', []))}
  </div>
</body>
</html>
"""

    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    print(f"preview={output_path}")


if __name__ == "__main__":
    main()
