from __future__ import annotations

import argparse
import json
import math
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


def load_font(size: int):
    from PIL import ImageFont

    candidates = [
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica.ttc",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def wrap(text: str, width: int) -> list[str]:
    lines = textwrap.wrap(text, width=width)
    return lines or [text]


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
    font_title = load_font(38)
    font_subtitle = load_font(24)
    font_body = load_font(22)
    font_small = load_font(18)

    draw.rectangle((0, 0, width, 18), fill=hex_to_rgb(theme["accent_color"]))
    draw.text((70, 60), str(slide.get("title", "")), fill=hex_to_rgb(theme["text_color"]), font=font_title)

    slide_type = slide.get("type")
    subtitle = slide.get("subtitle", "")
    if subtitle:
        draw.text((70, 122), str(subtitle), fill=(92, 102, 116), font=font_subtitle)

    if slide_type == "title":
        draw.rectangle((70, 220, 1520, 620), outline=hex_to_rgb(theme["accent_color"]), width=8)
        title_lines = wrap(str(slide.get("title", "")), 28)
        start_y = 280
        for line in title_lines:
            draw.text((110, start_y), line, fill=hex_to_rgb(theme["text_color"]), font=load_font(54))
            start_y += 74
        if subtitle:
            draw.text((110, start_y + 16), subtitle, fill=(92, 102, 116), font=font_subtitle)
        return image

    if slide_type == "schema":
        intro = str(slide.get("intro_text", ""))
        if intro:
            draw.text((70, 170), intro, fill=(60, 70, 85), font=font_body)
        current_y = 240
        for table_name in slide.get("tables", []):
            draw.text((70, current_y), str(table_name), fill=hex_to_rgb(theme["primary_color"]), font=font_body)
            current_y += 40
            headers = ["column_name", "data_type", "is_partition_key", "column_comment"]
            rows = payload.get("tables", {}).get(table_name, [])
            draw_table(draw, 70, current_y, 1460, headers, rows, font_body, font_small, theme)
            current_y += 240
        return image

    if slide_type == "table":
        bullet_y = 170
        for bullet in slide.get("takeaway_bullets", []):
            for line in wrap(str(bullet), 70):
                draw.text((100, bullet_y), f"- {line}", fill=(45, 53, 66), font=font_body)
                bullet_y += 34
        query = slide.get("query")
        query_data = payload.get("queries", {}).get(query, {})
        headers = query_data.get("columns", [])
        rows = query_data.get("rows", [])
        draw_table(draw, 70, 320, 1460, headers, rows, font_body, font_small, theme)
        return image

    if slide_type == "chart":
        bullet_y = 170
        for bullet in slide.get("takeaway_bullets", []):
            for line in wrap(str(bullet), 70):
                draw.text((100, bullet_y), f"- {line}", fill=(45, 53, 66), font=font_body)
                bullet_y += 34
        query = slide.get("query")
        query_data = payload.get("queries", {}).get(query, {})
        rows = query_data.get("rows", [])
        columns = query_data.get("columns", [])
        category_column = slide.get("category_column") or (columns[0] if columns else "")
        value_columns = slide.get("value_columns") or [
            column
            for column in columns
            if column != category_column
            and any(to_float(row.get(column)) is not None for row in rows)
        ][:3]
        meta = f"{slide.get('chart_type', 'line')} | {category_column} | {', '.join(value_columns)}"
        draw.text((70, 300), meta, fill=(92, 102, 116), font=font_small)
        if rows and category_column and value_columns:
            draw_simple_chart(draw, 70, 350, 1280, rows, category_column, value_columns[0], font_body, font_small, theme)
        else:
            draw.text((70, 380), "Chart preview is not available for this slide.", fill=(92, 102, 116), font=font_body)
        return image

    if slide_type == "ai_bullets":
        bullet_y = 190
        for bullet in slide.get("bullets", []):
            for line in wrap(str(bullet), 76):
                draw.text((100, bullet_y), f"- {line}", fill=(45, 53, 66), font=font_body)
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
    theme = rendered_config.get("report", {}).get("theme", {})
    theme = {
        "primary_color": theme.get("primary_color", "0F4C81"),
        "accent_color": theme.get("accent_color", "F28E2B"),
        "text_color": theme.get("text_color", "1F1F1F"),
        "background_color": theme.get("background_color", "FFFFFF"),
    }

    slide_images = [render_slide_image(slide, payload, theme) for slide in plan.get("slides", [])]
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
    label_font = load_font(22)

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
