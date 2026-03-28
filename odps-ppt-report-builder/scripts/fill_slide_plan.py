from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill a slide plan with deterministic starter content.")
    parser.add_argument("--payload", required=True, help="Path to report_payload.json.")
    parser.add_argument("--template", required=True, help="Path to slide_plan.template.json.")
    parser.add_argument("--output", required=True, help="Path to slide_plan.generated.json.")
    return parser.parse_args()


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def dump_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def to_float(value: Any) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def pick_numeric_columns(rows: list[dict[str, Any]], columns: list[str]) -> list[str]:
    if not rows:
        return []
    numeric_columns: list[str] = []
    for column in columns:
        if any(to_float(row.get(column)) is not None for row in rows):
            numeric_columns.append(column)
    return numeric_columns


def compare_first_last(rows: list[dict[str, Any]], column: str) -> str | None:
    if len(rows) < 2:
        return None
    first = to_float(rows[0].get(column))
    last = to_float(rows[-1].get(column))
    if first in (None, 0) or last is None:
        return None
    delta = last - first
    pct = delta / first * 100
    direction = "increased" if delta >= 0 else "decreased"
    return f"{column} {direction} from {rows[0].get(column)} to {rows[-1].get(column)} ({pct:.1f}%)."


def build_table_takeaways(query_name: str, query_data: dict[str, Any]) -> list[str]:
    rows = query_data.get("rows", [])
    columns = query_data.get("columns", [])
    if not rows:
        return [f"{query_name} returned no rows.", "Validate filters and source partitions before writing conclusions."]

    bullets = [f"{query_name} returned {len(rows)} rows; first row is {rows[0]}."]
    numeric_columns = pick_numeric_columns(rows, columns)
    if numeric_columns:
        trend_line = compare_first_last(rows, numeric_columns[0])
        if trend_line:
            bullets.append(trend_line)
        else:
            top_row = max(rows, key=lambda row: to_float(row.get(numeric_columns[0])) or float("-inf"))
            bullets.append(
                f"Highest {numeric_columns[0]} appears in row {top_row}."
            )
    else:
        bullets.append("No numeric columns were detected for trend analysis.")
    return bullets[:2]


def build_executive_summary(payload: dict[str, Any]) -> list[str]:
    query_map = payload.get("queries", {})
    populated = [name for name, item in query_map.items() if item.get("rows")]
    bullets: list[str] = []
    bullets.append(f"{len(populated)} query blocks returned data and are ready for review.")

    best_signal = None
    best_value = None
    for name, item in query_map.items():
        rows = item.get("rows", [])
        columns = item.get("columns", [])
        numeric_columns = pick_numeric_columns(rows, columns)
        if not rows or not numeric_columns:
            continue
        target_col = numeric_columns[0]
        local_best_row = max(rows, key=lambda row: to_float(row.get(target_col)) or float("-inf"))
        local_best_value = to_float(local_best_row.get(target_col))
        if local_best_value is None:
            continue
        if best_value is None or local_best_value > best_value:
            best_value = local_best_value
            best_signal = (name, target_col, local_best_row)

    if best_signal:
        name, column, row = best_signal
        bullets.append(f"Strongest visible signal is in {name}: {column} reached {row.get(column)}.")
    else:
        bullets.append("No strong numeric signal was detected automatically; manual review is still required.")

    bullets.append("Next step: validate the narrative against raw ODPS output before sharing externally.")
    return bullets[:3]


def main() -> None:
    args = parse_args()
    payload = load_json(Path(args.payload).resolve())
    template = load_json(Path(args.template).resolve())

    for slide in template.get("slides", []):
        slide_type = slide.get("type")
        if slide_type == "schema":
            slide["intro_text"] = "The source schema below is the basis for all table interpretation in this deck."
            slide["speaker_notes"] = "Review sensitive fields before final distribution."
        elif slide_type == "table":
            query_name = slide.get("query")
            query_data = payload.get("queries", {}).get(query_name, {})
            slide["takeaway_bullets"] = build_table_takeaways(query_name, query_data)
            slide["speaker_notes"] = "This slide was prefilled by a deterministic helper and should be reviewed by Qoder Work."
        elif slide_type == "ai_bullets":
            slide["bullets"] = build_executive_summary(payload)
            slide["speaker_notes"] = "Replace or refine this summary if stronger business context is available."

    dump_json(Path(args.output).resolve(), template)
    print(f"generated_plan={Path(args.output).resolve()}")


if __name__ == "__main__":
    main()
