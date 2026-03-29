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


def readable_metric_label(column: str) -> str:
    mapping = {
        "gmv": "GMV",
        "order_cnt": "订单量",
        "buyer_cnt": "买家数",
        "arpu": "ARPU",
        "freq": "频次",
        "aov": "笔单价",
        "week": "周次",
        "week_range": "周区间",
    }
    return mapping.get(column.strip().lower(), column.replace("_", " "))


def format_value(value: Any) -> str:
    numeric = to_float(value)
    if numeric is None:
        return str(value)
    if abs(numeric - round(numeric)) < 1e-6:
        return f"{int(round(numeric)):,}"
    return f"{numeric:,.2f}"


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


def build_table_takeaways(slide: dict[str, Any], query_name: str, query_data: dict[str, Any]) -> list[str]:
    rows = query_data.get("rows", [])
    columns = query_data.get("columns", [])
    template = str(slide.get("template", "")).strip().lower()
    if not rows:
        return [f"{query_name} 这组查询没有返回数据。", "先检查筛选条件和分区，再决定这一页怎么写。"]

    if template == "overview":
        row = rows[0]
        metrics = []
        for column in ["gmv", "order_cnt", "buyer_cnt"]:
            if column in row:
                metrics.append(f"{readable_metric_label(column)} {format_value(row.get(column))}")
        bullets = [
            "这一页已经切到月累计口径，用来作为整份月报的正式开场。",
            "，".join(metrics) + "。" if metrics else "优先在这一页讲清楚本月整体经营判断。",
        ]
        return bullets[:2]

    if template == "kpi_summary":
        visible_metrics = [readable_metric_label(column) for column in columns[1:7]]
        bullets = [
            "这一页只保留本月累计核心指标，避免再按天解释波动。",
            f"建议重点关注 {' / '.join(visible_metrics[:6])}。"
            if visible_metrics
            else "建议只保留管理层真正关心的核心指标。",
        ]
        return bullets[:2]

    if template == "appendix":
        bullets = [
            "附录页保留口径说明和日级辅助数据，方便复核但不进入主叙事。",
            f"{query_name} 当前展示 {min(len(rows), slide.get('max_rows', 12))} 行辅助数据。",
        ]
        return bullets[:2]

    if template in {"time_slot", "price_band", "time_price_matrix", "city_distribution", "aoi_distribution", "new_old_mix", "cluster_mix"}:
        numeric_cols = pick_numeric_columns(rows, columns)
        category_column = next((column for column in columns if column not in numeric_cols), columns[0] if columns else "")
        if numeric_cols and category_column:
            top_row = max(rows, key=lambda row: to_float(row.get(numeric_cols[0])) or float("-inf"))
            bullets = [
                f"这一页聚焦本月{slide.get('title', '').replace('分析', '').replace('页', '')}结构，不和其他主题混讲。",
                f"{top_row.get(category_column)} 在 {readable_metric_label(numeric_cols[0])} 上表现最强，适合作为页面主结论。",
            ]
            return bullets[:2]

    bullets = [f"{query_name} 一共返回 {len(rows)} 行，当前已切到月报页里使用。"]
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
    template = str(slide.get("template", "")).strip().lower()
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

    if template == "weekly_trend":
        first_metric = value_columns[0]
        numeric_rows = [row for row in rows if to_float(row.get(first_metric)) is not None]
        if numeric_rows:
            strongest_week = max(numeric_rows, key=lambda row: to_float(row.get(first_metric)) or float("-inf"))
            bullets = [
                "月报趋势页已经改成按周讲，不再用日级波动做主叙事。",
                f"{strongest_week.get(category_column)} 的 {readable_metric_label(first_metric)} 最高，值为 {format_value(strongest_week.get(first_metric))}。",
            ]
            return bullets[:2]

    if template == "comparison":
        first_metric = value_columns[0]
        numeric_rows = [row for row in rows if to_float(row.get(first_metric)) is not None]
        if len(numeric_rows) >= 2:
            first_row = numeric_rows[0]
            last_row = numeric_rows[-1]
            first_value = to_float(first_row.get(first_metric)) or 0.0
            last_value = to_float(last_row.get(first_metric)) or 0.0
            direction = "提升" if last_value >= first_value else "回落"
            bullets = [
                "这一页只负责支撑观点，对比关系不作为月报开场主叙事。",
                f"{readable_metric_label(first_metric)} 相比参考期整体{direction}，可用作本月判断的证据页。",
            ]
            return bullets[:2]

    if template in {"time_slot", "price_band", "city_distribution", "aoi_distribution", "new_old_mix"}:
        first_metric = value_columns[0]
        numeric_rows = [row for row in rows if to_float(row.get(first_metric)) is not None]
        if numeric_rows:
            top_row = max(numeric_rows, key=lambda row: to_float(row.get(first_metric)) or float("-inf"))
            bullets = [
                f"这一页只讲本月{slide.get('title', '').replace('分析', '').replace('结构', '').replace('分布', '')}结构，不和其他主题混讲。",
                f"{top_row.get(category_column)} 在 {readable_metric_label(first_metric)} 上表现最强，适合作为页面主结论。",
            ]
            return bullets[:2]

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
    bullets: list[str] = []
    overview_summary = query_map.get("overview_daily__monthly_summary") or query_map.get("overview__monthly_summary")
    weekly_trend = query_map.get("overview_daily__weekly") or query_map.get("overview__weekly")

    if overview_summary and overview_summary.get("rows"):
        row = overview_summary["rows"][0]
        bullets.append(
            f"本月整体经营已汇总到月累计口径，GMV {format_value(row.get('gmv'))}，订单量 {format_value(row.get('order_cnt'))}。"
        )
    else:
        bullets.append("本月整体经营口径已经切换成月累计视角，适合作为正式月报输出。")

    if weekly_trend and weekly_trend.get("rows"):
        first_metric = next(
            (column for column in weekly_trend.get("columns", []) if column not in {"week", "week_range"} and to_float(weekly_trend["rows"][0].get(column)) is not None),
            "",
        )
        if first_metric:
            top_row = max(
                weekly_trend["rows"],
                key=lambda row: to_float(row.get(first_metric)) or float("-inf"),
            )
            bullets.append(
                f"周趋势页会按周讲本月节奏，其中 {top_row.get('week')} 的 {readable_metric_label(first_metric)} 最强。"
            )

    structure_candidates = [
        ("time_slot", "时段"),
        ("price_band", "价格带"),
        ("city_distribution", "城市"),
        ("aoi_distribution", "AOI"),
        ("new_old_mix", "新老客"),
        ("cluster_mix", "cluster"),
    ]
    for query_name, label in structure_candidates:
        item = query_map.get(query_name)
        if not item or not item.get("rows"):
            continue
        rows = item["rows"]
        columns = item.get("columns", [])
        numeric_cols = pick_numeric_columns(rows, columns)
        category_column = next((column for column in columns if column not in numeric_cols), columns[0] if columns else "")
        if numeric_cols and category_column:
            top_row = max(rows, key=lambda row: to_float(row.get(numeric_cols[0])) or float("-inf"))
            bullets.append(f"{label}结构里，{top_row.get(category_column)} 是当前最值得优先关注的主力板块。")
            break

    bullets.append("下个月建议围绕主力结构继续拆解，并把日级数据退到附录或复核环节。")
    return bullets[:4]


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
            slide["takeaway_bullets"] = build_table_takeaways(slide, query_name, query_data)
            slide["speaker_notes"] = "这页默认按月报口径填充，发出去前再结合真实业务口径核一遍。"
        elif slide_type == "chart":
            query_name = slide.get("query")
            query_data = payload.get("queries", {}).get(query_name, {})
            slide["takeaway_bullets"] = build_chart_takeaways(slide, query_data)
            slide["speaker_notes"] = "趋势页默认按月报逻辑组织，确认横轴周次和参考关系是否符合业务口径。"
        elif slide_type == "ai_bullets":
            slide["bullets"] = build_executive_summary(payload)
            slide["speaker_notes"] = "结论和建议页是月报收束页，最好再结合业务背景顺一遍。"

    dump_json(Path(args.output).resolve(), template)
    print(f"generated_plan={Path(args.output).resolve()}")


if __name__ == "__main__":
    main()
