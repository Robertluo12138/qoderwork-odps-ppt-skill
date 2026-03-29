from __future__ import annotations

import argparse
import json
import math
import platform
import textwrap
from pathlib import Path
from typing import Any

from common import deep_render, load_yaml


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a PNG contact sheet preview for the slide deck.")
    parser.add_argument("--config", required=True, help="Path to the YAML report config.")
    parser.add_argument("--payload", required=True, help="Path to report_payload.json.")
    parser.add_argument("--plan", required=True, help="Path to slide_plan.generated.json.")
    parser.add_argument("--output", required=True, help="Path to the PNG output.")
    return parser.parse_args()


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def hex_to_rgb(color: str) -> tuple[int, int, int]:
    clean = color.strip().lstrip("#")
    return tuple(int(clean[index : index + 2], 16) for index in (0, 2, 4))


def to_float(value: Any) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def resolve_theme_font(raw_theme: dict[str, Any], base_key: str, default_value: str) -> str:
    system_name = platform.system().lower()
    suffix = "other"
    if system_name == "darwin":
        suffix = "mac"
    elif system_name == "windows":
        suffix = "windows"
    return str(raw_theme.get(f"{base_key}_{suffix}") or raw_theme.get(base_key) or default_value)


def default_cjk_preview_font() -> str:
    system_name = platform.system().lower()
    if system_name == "darwin":
        return "Hiragino Sans GB"
    if system_name == "windows":
        return "Microsoft YaHei"
    return "Noto Sans CJK SC"


def font_candidates(theme: dict[str, Any]) -> list[str]:
    family_candidates = {
        "Arial Unicode MS": [
            "/Library/Fonts/Arial Unicode.ttf",
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        ],
        "Hiragino Sans GB": [
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
        ],
        "STHeiti": [
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
        ],
        "Microsoft YaHei": [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/msyh.ttf",
        ],
        "DengXian": [
            "C:/Windows/Fonts/Deng.ttf",
        ],
        "SimHei": [
            "C:/Windows/Fonts/simhei.ttf",
        ],
        "SimSun": [
            "C:/Windows/Fonts/simsun.ttc",
        ],
        "Noto Sans CJK SC": [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ],
        "DejaVu Sans": [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ],
    }
    ordered_families = [
        theme.get("cjk_font_family", ""),
        theme.get("font_family", ""),
        "Arial Unicode MS",
        "Hiragino Sans GB",
        "STHeiti",
        "Microsoft YaHei",
        "DengXian",
        "SimHei",
        "SimSun",
        "Noto Sans CJK SC",
        "DejaVu Sans",
    ]
    seen: set[str] = set()
    candidates: list[str] = []
    for family in ordered_families:
        for path in family_candidates.get(str(family), []):
            if path not in seen:
                seen.add(path)
                candidates.append(path)
    candidates.extend(
        [
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Helvetica.ttc",
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/arial.ttf",
        ]
    )
    return candidates


def load_font(size: int, theme: dict[str, Any]):
    from PIL import ImageFont

    for candidate in font_candidates(theme):
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def wrap(text: str, width: int) -> list[str]:
    lines = textwrap.wrap(text, width=width)
    return lines or [text]


def infer_template_name(slide: dict[str, Any]) -> str:
    explicit = str(slide.get("template", "")).strip().lower()
    if explicit:
        return explicit
    title = str(slide.get("title", "")).lower()
    slide_type = str(slide.get("type", "")).lower()
    if slide_type == "title":
        return "cover"
    if "经营概览" in title or "总览" in title:
        return "overview"
    if "核心指标" in title or "指标摘要" in title or "摘要" in title:
        return "kpi_summary"
    if "趋势" in title:
        return "weekly_trend"
    if "对比" in title or "分析" in title:
        return "comparison"
    if "结论" in title or "总结" in title or "summary" in title:
        return "conclusion"
    if "致谢" in title or "感谢" in title or "thank" in title:
        return "thanks"
    if "附录" in title:
        return "appendix"
    return ""


def split_lead_and_points(slide: dict[str, Any]) -> tuple[str, list[str]]:
    subtitle = str(slide.get("subtitle", "")).strip()
    bullets = [str(item).strip() for item in slide.get("bullets", []) if str(item).strip()]
    takeaways = [
        str(item).strip()
        for item in slide.get("takeaway_bullets", [])
        if str(item).strip()
    ]
    points = bullets or takeaways
    if subtitle:
        return subtitle, points
    if points:
        return points[0], points[1:]
    return "", []


def format_metric_value(value: Any) -> str:
    numeric = to_float(value)
    if numeric is None:
        return str(value)
    if abs(numeric - round(numeric)) < 1e-6:
        return f"{int(round(numeric)):,}"
    return f"{numeric:,.1f}"


def format_delta_text(current: float | None, previous: float | None) -> str:
    if current is None or previous is None:
        return ""
    if abs(previous) > 1e-9:
        ratio = (current - previous) / abs(previous)
        sign = "+" if ratio >= 0 else ""
        return f"{sign}{ratio * 100:.1f}% vs 上期"
    delta = current - previous
    sign = "+" if delta >= 0 else ""
    return f"{sign}{delta:,.0f} vs 上期"


def readable_metric_label(column: str) -> str:
    mapping = {
        "gmv": "GMV",
        "order_cnt": "订单量",
        "buyer_cnt": "买家数",
        "pay_buyer_cnt": "支付买家数",
        "arpu": "ARPU",
        "freq": "频次",
        "aov": "笔单价",
        "uv": "访客数",
        "pv": "浏览量",
        "ctr": "点击率",
        "conv_rate": "转化率",
        "member_gmv": "会员GMV",
        "member_order_cnt": "会员订单量",
        "member_buyer_cnt": "会员买家数",
    }
    normalized = column.strip().lower()
    return mapping.get(normalized, column.replace("_", " ").title())


def compact_context_label(value: Any) -> str:
    text = str(value).strip()
    if len(text) == 10 and text[4] == "-" and text[7] == "-":
        return text[5:]
    return text


def extract_metric_cards(query_data: dict[str, Any], max_metrics: int = 4) -> tuple[str, list[dict[str, str]]]:
    columns = query_data.get("columns", [])
    rows = query_data.get("rows", [])
    if not columns or not rows:
        return "", []
    latest = rows[-1]
    previous = rows[-2] if len(rows) > 1 else None
    context_column = next(
        (
            column
            for column in columns
            if any(to_float(row.get(column)) is None for row in rows)
        ),
        columns[0],
    )
    context_label = str(latest.get(context_column, ""))
    metric_columns = [
        column
        for column in columns
        if column != context_column
        and any(to_float(row.get(column)) is not None for row in rows)
    ][:max_metrics]
    cards = []
    for column in metric_columns:
        current = to_float(latest.get(column))
        previous_value = to_float(previous.get(column)) if previous else None
        cards.append(
            {
                "label": readable_metric_label(column),
                "value": format_metric_value(latest.get(column, "")),
                "delta": format_delta_text(current, previous_value),
            }
        )
    return context_label, cards


def draw_card(draw, left: int, top: int, width: int, height: int, fill: tuple[int, int, int], outline: tuple[int, int, int], radius: int = 24):
    draw.rounded_rectangle((left, top, left + width, top + height), radius=radius, fill=fill, outline=outline)


def draw_metric_card(draw, x: int, y: int, width: int, height: int, metric: dict[str, str], font_body, font_small, font_title, theme: dict[str, Any], highlight: bool = False):
    fill = hex_to_rgb(theme["soft_fill_color"]) if highlight else hex_to_rgb(theme["card_bg_color"])
    outline = hex_to_rgb(theme["warm_fill_color"]) if highlight else hex_to_rgb(theme["line_color"])
    draw_card(draw, x, y, width, height, fill, outline, radius=28)
    draw.text((x + 18, y + 18), metric.get("label", ""), fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
    draw.text((x + 18, y + 52), metric.get("value", ""), fill=hex_to_rgb(theme["text_color"]), font=font_title)
    if metric.get("delta"):
        draw.text((x + 18, y + height - 34), metric["delta"], fill=hex_to_rgb(theme["primary_color"]), font=font_small)


def draw_table(draw, x: int, y: int, width: int, headers: list[str], rows: list[dict[str, Any]], font, small_font, theme: dict[str, Any]):
    from PIL import ImageColor

    row_height = 26
    visible_rows = rows[:6]
    columns = max(1, len(headers))
    col_width = width // columns
    draw.rectangle((x, y, x + width, y + row_height), fill=ImageColor.getrgb(f"#{theme['primary_color']}"))
    for index, header in enumerate(headers):
        left = x + index * col_width
        right = left + col_width
        draw.rectangle((left, y, right, y + row_height), outline=(220, 227, 238))
        draw.text((left + 8, y + 6), str(header), fill=(255, 255, 255), font=small_font)

    for row_idx, row in enumerate(visible_rows, start=1):
        top = y + row_idx * row_height
        fill = (255, 255, 255) if row_idx % 2 else (247, 249, 252)
        draw.rectangle((x, top, x + width, top + row_height), fill=fill, outline=(220, 227, 238))
        for col_idx, header in enumerate(headers):
            left = x + col_idx * col_width
            right = left + col_width
            draw.rectangle((left, top, right, top + row_height), outline=(220, 227, 238))
            draw.text((left + 8, top + 6), str(row.get(header, ""))[:22], fill=(31, 31, 31), font=small_font)


def draw_simple_chart(draw, x: int, y: int, width: int, rows: list[dict[str, Any]], category_column: str, value_column: str, font_body, font_small, theme: dict[str, Any]):
    chart_rows = rows[:6]
    values = [to_float(row.get(value_column)) or 0.0 for row in chart_rows]
    max_value = max(values) if values else 0.0
    current_y = y

    for row, value in zip(chart_rows, values):
        label = str(row.get(category_column, ""))[:18]
        draw.text((x, current_y), label, fill=(45, 53, 66), font=font_small)
        draw.rounded_rectangle((x + 220, current_y + 6, x + width, current_y + 28), radius=10, fill=(235, 240, 246))
        if max_value > 0:
            bar_width = int((width - 220) * value / max_value)
            draw.rounded_rectangle(
                (x + 220, current_y + 6, x + 220 + bar_width, current_y + 28),
                radius=10,
                fill=hex_to_rgb(theme["primary_color"]),
            )
        draw.text((x + width + 14, current_y + 2), str(row.get(value_column, ""))[:12], fill=(45, 53, 66), font=font_small)
        current_y += 54


def render_slide_image(slide: dict[str, Any], payload: dict[str, Any], theme: dict[str, Any]):
    from PIL import Image, ImageDraw

    width, height = 1600, 900
    image = Image.new("RGB", (width, height), color=hex_to_rgb(theme["background_color"]))
    draw = ImageDraw.Draw(image)
    font_title = load_font(36, theme)
    font_hero = load_font(54, theme)
    font_subtitle = load_font(22, theme)
    font_body = load_font(20, theme)
    font_small = load_font(16, theme)

    template = infer_template_name(slide)
    slide_type = slide.get("type")

    if template == "cover":
        logo_path = theme.get("_logo_resolved_path")
        if logo_path and Path(logo_path).exists():
            cover_logo = Image.open(logo_path).convert("RGBA")
            lh = 48
            lw = int(cover_logo.width * lh / cover_logo.height)
            cover_logo_resized = cover_logo.resize((lw, lh), Image.LANCZOS)
            image.paste(cover_logo_resized, (82, 6), cover_logo_resized)
        draw_card(draw, 70, 54, 168, 42, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=20)
        draw.text((92, 63), "品牌经营分析", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
        draw.rectangle((96, 162, 106, 760), fill=hex_to_rgb(theme["primary_color"]))
        title_lines = wrap(str(slide.get("title", "")), 12)
        start_y = 146
        for line in title_lines:
            draw.text((144, start_y), line, fill=hex_to_rgb(theme["text_color"]), font=font_hero)
            start_y += 82
        draw.text((144, 446), "淘宝闪购经营分析作战室", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
        draw.rectangle((144, 492, 754, 496), fill=hex_to_rgb(theme["line_color"]))
        subtitle = str(slide.get("subtitle", ""))
        if subtitle:
            draw.text((144, 528), subtitle, fill=hex_to_rgb(theme["muted_text_color"]), font=font_subtitle)
        draw.text((144, 650), "品牌经营分析 / 关键指标摘要 / 重点专题拆解", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
        draw.text((144, 790), "围绕本期经营表现、关键指标变化和重点结构问题展开。", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

        draw.rectangle((1088, 116, 1356, 792), fill=hex_to_rgb(theme["card_bg_color"]), outline=hex_to_rgb(theme["line_color"]))
        draw.rectangle((1328, 116, 1342, 792), fill=hex_to_rgb(theme["primary_color"]))
        draw.text((1114, 152), "MONTHLY REVIEW", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
        info_rows = [
            ("报告主题", str(slide.get("title", ""))),
            ("观察周期", subtitle or "-"),
            ("材料类型", "品牌经营分析报告"),
        ]
        row_y = 216
        for label, value in info_rows:
            draw.text((1114, row_y), label, fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
            for idx, line in enumerate(wrap(value, 11)[:2]):
                draw.text((1114, row_y + 30 + idx * 24), line, fill=hex_to_rgb(theme["text_color"]), font=font_small if label == "观察周期" else font_body)
            draw.rectangle((1114, row_y + 86, 1296, row_y + 88), fill=hex_to_rgb(theme["line_color"]))
            row_y += 128
        draw.rectangle((1114, 646, 1296, 734), fill=hex_to_rgb(theme["soft_fill_color"]), outline=hex_to_rgb(theme["soft_fill_color"]))
        draw.text((1140, 670), "摘要提示", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
        for idx, line in enumerate(wrap("先看品牌整体经营，再看结构与参考对比。", 10)[:2]):
            draw.text((1140, 700 + idx * 22), line, fill=hex_to_rgb(theme["text_color"]), font=font_small)
        return image

    if template not in {"overview", "comparison", "conclusion"}:
        chip_text = "经营分析"
        if template == "weekly_trend":
            chip_text = "本月趋势"
        elif template in {"time_slot", "price_band", "time_price_matrix", "city_distribution", "aoi_distribution", "new_old_mix", "cluster_mix"}:
            chip_text = "结构分析"
        elif template == "appendix":
            chip_text = "附录"
        # Draw hero banner for structure analysis and trend slides
        has_banner = template not in {"kpi_summary", "appendix", "thanks", ""}
        takeaway_texts = [str(item) for item in slide.get("takeaway_bullets", []) if str(item).strip()]
        if has_banner:
            draw.rectangle((0, 0, width, 82), fill=hex_to_rgb(theme["primary_color"]))
            banner_text = takeaway_texts[0] if takeaway_texts else str(slide.get("title", ""))
            for idx, line in enumerate(wrap(banner_text, 42)[:2]):
                draw.text((52, 16 + idx * 34), line, fill=(255, 255, 255), font=font_subtitle)
            draw_card(draw, 72, 98, 130, 38, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=18)
            draw.text((92, 108), chip_text, fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((72, 152), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_title)
        else:
            draw_card(draw, 72, 58, 130, 38, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=18)
            draw.text((92, 68), chip_text, fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((72, 118), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_title)

    if slide_type == "schema":
        intro = str(slide.get("intro_text", ""))
        if intro:
            draw.text((74, 168), intro, fill=hex_to_rgb(theme["muted_text_color"]), font=font_body)
        current_y = 220
        for table_name in slide.get("tables", []):
            draw.text((74, current_y), str(table_name), fill=hex_to_rgb(theme["primary_color"]), font=font_body)
            current_y += 38
            headers = ["column_name", "data_type", "is_partition_key", "column_comment"]
            rows = payload.get("tables", {}).get(table_name, [])
            draw_card(draw, 70, current_y - 6, 1460, 170, hex_to_rgb(theme["card_bg_color"]), hex_to_rgb(theme["line_color"]), radius=24)
            draw_table(draw, 82, current_y + 6, 1434, headers, rows, font_body, font_small, theme)
            current_y += 205
        return image

    if slide_type == "table":
        query_data = payload.get("queries", {}).get(slide.get("query"), {})
        headers = query_data.get("columns", [])
        rows = query_data.get("rows", [])
        context_label, metric_cards = extract_metric_cards(query_data, 6)
        takeaways = [str(item) for item in slide.get("takeaway_bullets", [])]

        if template == "overview":
            lead = takeaways[0] if takeaways else "本期整体经营表现继续走强，核心指标延续上升。"
            support = takeaways[1] if len(takeaways) > 1 else "经营盘面保持改善，末端走势仍在抬升。"
            hero_metric = metric_cards[0] if metric_cards else {"label": "核心指标", "value": "-", "delta": ""}
            support_metrics = metric_cards[1:4]
            latest_context = context_label or (rows[-1].get(headers[0], "") if rows and headers else "")
            support_query_name = str(slide.get("support_query", "")).strip()
            support_query_data = payload.get("queries", {}).get(support_query_name, {})
            support_headers = support_query_data.get("columns", [])
            support_rows = support_query_data.get("rows", [])

            draw_card(draw, 70, 54, 138, 40, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=18)
            draw.text((92, 63), "经营概览", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((72, 120), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_title)
            draw.text((1140, 108), f"观察窗口  {latest_context}", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            draw_card(draw, 72, 234, 786, 414, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["warm_fill_color"]), radius=36)
            draw_card(draw, 94, 258, 96, 30, hex_to_rgb(theme["background_color"]), hex_to_rgb(theme["background_color"]), radius=16)
            draw.text((110, 264), "本期判断", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            for idx, line in enumerate(wrap(lead, 22)[:3]):
                draw.text((102, 318 + idx * 36), line, fill=hex_to_rgb(theme["text_color"]), font=font_subtitle)
            draw.text((102, 470), hero_metric.get("label", ""), fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
            draw.text((102, 506), hero_metric.get("value", "-"), fill=hex_to_rgb(theme["text_color"]), font=load_font(58, theme))
            if hero_metric.get("delta"):
                draw.text((416, 526), hero_metric["delta"], fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((102, 590), support, fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            for index, metric in enumerate(support_metrics[:3]):
                top = 234 + index * 142
                draw_card(draw, 910, top, 542, 120, hex_to_rgb(theme["card_bg_color"]), hex_to_rgb(theme["line_color"]), radius=28)
                draw.text((942, top + 22), metric.get("label", ""), fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
                draw.text((942, top + 52), metric.get("value", "-"), fill=hex_to_rgb(theme["text_color"]), font=font_subtitle)
                draw.text((1190, top + 58), metric.get("delta", ""), fill=hex_to_rgb(theme["primary_color"]), font=font_small)

            draw_card(draw, 72, 688, 1386, 120, hex_to_rgb(theme["card_bg_color"]), hex_to_rgb(theme["line_color"]), radius=28)
            if support_rows and support_headers:
                draw.text((100, 714), "月内周趋势", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
                strip_rows = support_rows[-5:] if len(support_rows) >= 5 else support_rows
                strip_category = support_headers[0]
                strip_value = next(
                    (
                        column
                        for column in support_headers[1:]
                        if any(to_float(item.get(column)) is not None for item in support_rows)
                    ),
                    support_headers[1] if len(support_headers) > 1 else support_headers[0],
                )
            else:
                draw.text((100, 714), "近六日经营走势", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
                strip_rows = rows[-6:] if len(rows) >= 6 else rows
                strip_category = headers[0] if headers else ""
                strip_value = headers[1] if len(headers) > 1 else (headers[0] if headers else "")
            for index, row in enumerate(strip_rows[:6]):
                x = 310 + index * 186
                draw_card(draw, x, 706, 166, 72, hex_to_rgb(theme["background_color"]), hex_to_rgb(theme["line_color"]), radius=22)
                label = compact_context_label(row.get(strip_category, "")) if strip_category else "-"
                value = format_metric_value(row.get(strip_value, "")) if strip_value else "-"
                draw.text((x + 18, 724), label, fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
                draw.text((x + 18, 748), value, fill=hex_to_rgb(theme["text_color"]), font=font_body)
            return image

        if template == "kpi_summary":
            draw_card(draw, 72, 184, 1458, 64, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=24)
            lead = takeaways[0] if takeaways else "先看本期最值得关注的核心指标。"
            draw.text((98, 202), lead, fill=hex_to_rgb(theme["text_color"]), font=font_body)
            for index, metric in enumerate(metric_cards[:6]):
                row = index // 3
                col = index % 3
                draw_metric_card(draw, 72 + col * 486, 290 + row * 176, 438, 148, metric, font_body, font_small, font_title, theme, highlight=index == 0)
            if context_label:
                draw.text((96, 760), f"指标口径参考 | 最新数据点：{context_label}", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
            return image

        if takeaways:
            draw_card(draw, 72, 184, 1458, 64, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=24)
            draw.text((98, 202), takeaways[0], fill=hex_to_rgb(theme["text_color"]), font=font_body)
        draw_card(draw, 72, 286, 1458, 404, hex_to_rgb(theme["card_bg_color"]), hex_to_rgb(theme["line_color"]), radius=30)
        draw_table(draw, 90, 320, 1420, headers, rows, font_body, font_small, theme)
        return image

    if slide_type == "chart":
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
        takeaways = [str(item) for item in slide.get("takeaway_bullets", [])]

        if template == "comparison":
            lead = takeaways[0] if takeaways else "先看趋势主线，再把参考对比当作辅助说明。"
            support_points = takeaways[1:3]
            context_label, metric_cards = extract_metric_cards(query_data, 2)
            opinion_text = support_points[0] if support_points else "最新一期 GMV 延续上行，当前处于稳步抬升区间。"

            draw_card(draw, 70, 54, 138, 40, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=18)
            draw.text((92, 63), "分析对比", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((72, 120), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_title)
            draw.text((72, 174), lead, fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            for idx, line in enumerate(wrap(opinion_text, 28)[:2]):
                draw.text((72, 218 + idx * 34), line, fill=hex_to_rgb(theme["text_color"]), font=font_subtitle)

            draw_card(draw, 72, 314, 1010, 500, hex_to_rgb(theme["card_bg_color"]), hex_to_rgb(theme["line_color"]), radius=32)
            draw_card(draw, 96, 336, 108, 30, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=16)
            draw.text((114, 342), "品牌主视角", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((96, 390), "趋势图只负责支撑观点，参考关系只做轻量标注。", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            draw_card(draw, 1112, 360, 278, 360, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["warm_fill_color"]), radius=28)
            draw_card(draw, 1134, 384, 92, 28, hex_to_rgb(theme["background_color"]), hex_to_rgb(theme["background_color"]), radius=14)
            draw.text((1150, 389), "证据摘要", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            for idx, line in enumerate(wrap("大盘仅作辅助参考，不与品牌主视角并列。", 11)[:3]):
                draw.text((1138, 438 + idx * 28), line, fill=hex_to_rgb(theme["text_color"]), font=font_small)
            evidence_points = support_points[:2] or ["最新一周 GMV 延续上行。", "末端点位仍在抬升。"]
            for index, bullet in enumerate(evidence_points):
                y = 532 + index * 58
                draw.text((1140, y), f"{index + 1:02d}", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
                for line_idx, line in enumerate(wrap(bullet, 10)[:2]):
                    draw.text((1178, y - 2 + line_idx * 22), line, fill=hex_to_rgb(theme["text_color"]), font=font_small)
            if metric_cards:
                draw_metric_card(draw, 1138, 638, 186, 86, metric_cards[0], font_body, font_small, font_body, theme, highlight=True)
            if context_label:
                draw.text((1138, 742), f"参考窗口  {context_label}", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            if rows and category_column and value_columns:
                draw_simple_chart(draw, 112, 470, 800, rows, category_column, value_columns[0], font_body, font_small, theme)
            else:
                draw.text((132, 560), "Chart preview is not available for this slide.", fill=hex_to_rgb(theme["muted_text_color"]), font=font_body)
            return image

        if takeaways:
            draw_card(draw, 72, 184, 1458, 64, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=24)
            draw.text((98, 202), takeaways[0], fill=hex_to_rgb(theme["text_color"]), font=font_body)
        draw_card(draw, 72, 286, 1458, 404, hex_to_rgb(theme["card_bg_color"]), hex_to_rgb(theme["line_color"]), radius=30)
        if rows and category_column and value_columns:
            draw_simple_chart(draw, 110, 340, 1220, rows, category_column, value_columns[0], font_body, font_small, theme)
        else:
            draw.text((98, 360), "Chart preview is not available for this slide.", fill=hex_to_rgb(theme["muted_text_color"]), font=font_body)
        return image

    if slide_type == "ai_bullets":
        lead_text, points = split_lead_and_points(slide)
        if template == "thanks":
            draw_card(draw, 72, 54, 110, 36, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=18)
            draw.text((88, 63), "报告结束", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.rectangle((86, 158, 98, 510), fill=hex_to_rgb(theme["primary_color"]))
            draw.text((132, 168), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_hero)
            subtitle = lead_text or "欢迎继续沟通具体问题。"
            draw.text((134, 296), subtitle, fill=hex_to_rgb(theme["muted_text_color"]), font=font_subtitle)
            draw.text((134, 398), "如需继续展开时段、价格带或城市结构，可在此基础上继续拆解。", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            draw.rectangle((1020, 182, 1330, 510), fill=hex_to_rgb(theme["card_bg_color"]), outline=hex_to_rgb(theme["line_color"]))
            draw.rectangle((1020, 182, 1034, 510), fill=hex_to_rgb(theme["primary_color"]))
            detail_rows = [
                ("报告来源", "品牌经营分析作战室"),
                ("输出形式", "可编辑经营分析PPT"),
                ("下一步", subtitle),
            ]
            row_y = 214
            for label, value in detail_rows:
                draw.text((1060, row_y), label, fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)
                for idx, line in enumerate(wrap(value, 12)[:2]):
                    draw.text((1060, row_y + 28 + idx * 22), line, fill=hex_to_rgb(theme["text_color"]), font=font_small)
                if label != "下一步":
                    draw.rectangle((1060, row_y + 78, 1280, row_y + 80), fill=hex_to_rgb(theme["line_color"]))
                row_y += 96
            return image

        if template == "conclusion":
            draw_card(draw, 70, 54, 138, 40, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["soft_fill_color"]), radius=18)
            draw.text((92, 63), "本期结论", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            draw.text((72, 120), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_title)

            lead_statement = lead_text or (points[0] if points else "本期经营延续改善，后续重点看结构优化和延续性。")
            roadmap_points = points or [str(item) for item in slide.get("bullets", [])]
            if lead_text and roadmap_points and roadmap_points[0] == lead_text:
                roadmap_points = roadmap_points[1:]

            draw_card(draw, 72, 240, 642, 556, hex_to_rgb(theme["soft_fill_color"]), hex_to_rgb(theme["warm_fill_color"]), radius=34)
            draw.rectangle((96, 270, 110, 394), fill=hex_to_rgb(theme["primary_color"]))
            draw.text((134, 270), "总判断", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
            for idx, line in enumerate(wrap(lead_statement, 14)[:4]):
                draw.text((134, 328 + idx * 46), line, fill=hex_to_rgb(theme["text_color"]), font=font_subtitle)
            draw.text((134, 620), "这一页先给结论，再把后续关注点拆成三条，方便管理层直接抓重点。", fill=hex_to_rgb(theme["muted_text_color"]), font=font_small)

            draw.rectangle((902, 272, 906, 748), fill=hex_to_rgb(theme["warm_fill_color"]))
            for index, point in enumerate(roadmap_points[:3]):
                cy = 304 + index * 144
                draw.ellipse((872, cy, 920, cy + 48), fill=hex_to_rgb(theme["background_color"]), outline=hex_to_rgb(theme["primary_color"]), width=3)
                draw.text((884, cy + 12), f"{index + 1:02d}", fill=hex_to_rgb(theme["primary_color"]), font=font_small)
                for line_idx, line in enumerate(wrap(str(point), 18)[:3]):
                    draw.text((952, cy - 2 + line_idx * 32), line, fill=hex_to_rgb(theme["text_color"]), font=font_body if index == 0 else font_small)
            return image

        bullet_y = 220
        for bullet in slide.get("bullets", []):
            for line in wrap(str(bullet), 76):
                draw.text((100, bullet_y), f"- {line}", fill=hex_to_rgb(theme["text_color"]), font=font_body)
                bullet_y += 38
            bullet_y += 12
        return image

    return image


def main() -> None:
    from PIL import Image, ImageDraw

    args = parse_args()
    config = load_yaml(Path(args.config).resolve())
    payload = load_json(Path(args.payload).resolve())
    plan = load_json(Path(args.plan).resolve())
    context = payload.get("metadata", {}).get("variables", {})
    rendered_config = deep_render(config, context)
    raw_theme = rendered_config.get("report", {}).get("theme", {})
    theme = {
        "primary_color": raw_theme.get("primary_color", "FF6200"),
        "accent_color": raw_theme.get("accent_color", "FFA020"),
        "soft_fill_color": raw_theme.get("soft_fill_color", "FFF4EB"),
        "warm_fill_color": raw_theme.get("warm_fill_color", "FFDCC8"),
        "text_color": raw_theme.get("text_color", "1F1F1F"),
        "muted_text_color": raw_theme.get("muted_text_color", "666666"),
        "line_color": raw_theme.get("line_color", "EAEAEA"),
        "card_bg_color": raw_theme.get("card_bg_color", "FAFAFA"),
        "background_color": raw_theme.get("background_color", "FFFFFF"),
        "font_family": resolve_theme_font(raw_theme, "font_family", "Aptos"),
        "cjk_font_family": resolve_theme_font(
            raw_theme,
            "cjk_font_family",
            default_cjk_preview_font(),
        ),
    }

    # Resolve logo path for brand watermark
    logo_rel = raw_theme.get("logo_path", "taobao_flash_sale_logo.png")
    if logo_rel:
        config_dir = Path(args.config).resolve().parent
        logo_candidate = config_dir / logo_rel
        if not logo_candidate.exists():
            logo_candidate = Path(__file__).resolve().parent.parent / "assets" / logo_rel
        if logo_candidate.exists():
            theme["_logo_resolved_path"] = str(logo_candidate)

    logo_footer_img = None
    _logo_path = theme.get("_logo_resolved_path")
    if _logo_path and Path(_logo_path).exists():
        _footer_logo = Image.open(_logo_path).convert("RGBA")
        _fh = 28
        _fw = int(_footer_logo.width * _fh / _footer_logo.height)
        logo_footer_img = _footer_logo.resize((_fw, _fh), Image.LANCZOS)

    slide_images = []
    for slide_data in plan.get("slides", []):
        img = render_slide_image(slide_data, payload, theme)
        if logo_footer_img and infer_template_name(slide_data) != "cover":
            img.paste(logo_footer_img, (img.width - logo_footer_img.width - 72, img.height - 42), logo_footer_img)
        slide_images.append(img)
    if not slide_images:
        raise SystemExit("No slides found in plan.")

    cols = 2
    thumb_width, thumb_height = 640, 360
    gap = 28
    rows = math.ceil(len(slide_images) / cols)
    sheet_width = cols * thumb_width + (cols + 1) * gap
    sheet_height = rows * thumb_height + (rows + 1) * gap
    sheet = Image.new("RGB", (sheet_width, sheet_height), color=(240, 244, 249))
    sheet_draw = ImageDraw.Draw(sheet)
    label_font = load_font(22, theme)

    for index, slide_image in enumerate(slide_images):
        row = index // cols
        col = index % cols
        x = gap + col * (thumb_width + gap)
        y = gap + row * (thumb_height + gap)
        sheet.paste(slide_image.resize((thumb_width, thumb_height)), (x, y))
        sheet_draw.rectangle((x, y, x + thumb_width, y + thumb_height), outline=(200, 208, 220), width=2)
        sheet_draw.text((x + 16, y + 14), f"Slide {index + 1}", fill=(255, 255, 255), font=label_font)

    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output_path)
    print(f"preview_png={output_path}")


if __name__ == "__main__":
    main()
