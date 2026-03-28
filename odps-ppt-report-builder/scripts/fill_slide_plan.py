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
    direction = "上升" if delta >= 0 else "下降"
    return f"{column} 从 {rows[0].get(column)} 到 {rows[-1].get(column)}，整体{direction} {abs(pct):.1f}%。"


def build_table_takeaways(query_name: str, query_data: dict[str, Any]) -> list[str]:
    rows = query_data.get("rows", [])
    columns = query_data.get("columns", [])
    if not rows:
        return [f"{query_name} 这组查询没有返回数据。", "先检查筛选条件和分区，再决定这一页怎么写。"]

    bullets = [f"{query_name} 一共返回 {len(rows)} 行，首行结果是 {rows[0]}。"]
    numeric_columns = pick_numeric_columns(rows, columns)
    if numeric_columns:
        trend_line = compare_first_last(rows, numeric_columns[0])
        if trend_line:
            bullets.append(trend_line)
        else:
            top_row = max(rows, key=lambda row: to_float(row.get(numeric_columns[0])) or float("-inf"))
            bullets.append(
                f"{numeric_columns[0]} 最大的一行是 {top_row}。"
            )
    else:
        bullets.append("这组结果里没有明显可用于趋势判断的数值列。")
    return bullets[:2]


def build_chart_takeaways(slide: dict[str, Any], query_data: dict[str, Any]) -> list[str]:
    rows = query_data.get("rows", [])
    columns = query_data.get("columns", [])
    category_column = slide.get("category_column") or (columns[0] if columns else "")
    value_columns = slide.get("value_columns") or [
        column
        for column in columns
        if column != category_column
        and any(to_float(row.get(column)) is not None for row in rows)
    ][:3]

    if not rows:
        return ["这张图现在没有数据。", "先检查筛选条件和分区，再决定这一页怎么写。"]

    if not value_columns:
        return ["这张图还没有找到可用的数值列。", "先明确 value_columns，再生成图表。"]

    bullets = [f"这张图用 {category_column} 做横轴，展示 {', '.join(value_columns)}。"]

    first_metric = value_columns[0]
    numeric_rows = [row for row in rows if to_float(row.get(first_metric)) is not None]
    if numeric_rows:
        top_row = max(numeric_rows, key=lambda row: to_float(row.get(first_metric)) or float("-inf"))
        bullets.append(
            f"{first_metric} 最高的是 {top_row.get(category_column)}，值为 {top_row.get(first_metric)}。"
        )
    else:
        bullets.append("第一组图表序列里没有找到可用的数值。")

    return bullets[:2]


def build_executive_summary(payload: dict[str, Any]) -> list[str]:
    query_map = payload.get("queries", {})
    populated = [name for name, item in query_map.items() if item.get("rows")]
    bullets: list[str] = []
    bullets.append(f"这次一共有 {len(populated)} 组查询拿到了数据，可以继续写报告。")

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
        bullets.append(f"当前最明显的信号出现在 {name}：{column} 达到 {row.get(column)}。")
    else:
        bullets.append("自动检查没有识别出特别强的数值信号，还是需要人工再看一遍。")

    bullets.append("对外发送前，先拿原始 ODPS 结果再对一遍口径。")
    return bullets[:3]


def main() -> None:
    args = parse_args()
    payload = load_json(Path(args.payload).resolve())
    template = load_json(Path(args.template).resolve())

    for slide in template.get("slides", []):
        slide_type = slide.get("type")
        if slide_type == "schema":
            slide["intro_text"] = "下面这部分是本次报告依赖的源表结构，后面的解释都要以这里为准。"
            slide["speaker_notes"] = "正式发出去前，再看一遍有没有敏感字段。"
        elif slide_type == "table":
            query_name = slide.get("query")
            query_data = payload.get("queries", {}).get(query_name, {})
            slide["takeaway_bullets"] = build_table_takeaways(query_name, query_data)
            slide["speaker_notes"] = "这页是脚本先填的首版，后面最好再结合业务语境顺一遍。"
        elif slide_type == "chart":
            query_name = slide.get("query")
            query_data = payload.get("queries", {}).get(query_name, {})
            slide["takeaway_bullets"] = build_chart_takeaways(slide, query_data)
            slide["speaker_notes"] = "发出去前确认一下图表类型、横轴字段和数值列是不是都对。"
        elif slide_type == "ai_bullets":
            slide["bullets"] = build_executive_summary(payload)
            slide["speaker_notes"] = "如果你对业务更熟，这页建议再手动润一下。"

    dump_json(Path(args.output).resolve(), template)
    print(f"generated_plan={Path(args.output).resolve()}")


if __name__ == "__main__":
    main()
