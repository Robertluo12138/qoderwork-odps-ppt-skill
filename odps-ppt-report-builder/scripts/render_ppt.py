from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from common import deep_render, load_yaml, output_dir_from_config


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a PPTX from a report payload and slide plan.")
    parser.add_argument("--config", required=True, help="Path to the YAML report config.")
    parser.add_argument("--payload", required=True, help="Path to report_payload.json.")
    parser.add_argument("--plan", required=True, help="Path to slide_plan.generated.json.")
    parser.add_argument("--output", default="", help="Optional explicit output path.")
    return parser.parse_args()


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def hex_to_rgb(color: str) -> tuple[int, int, int]:
    clean = color.strip().lstrip("#")
    if len(clean) != 6:
        raise ValueError(f"Expected 6-digit hex color, got: {color}")
    return tuple(int(clean[index : index + 2], 16) for index in (0, 2, 4))


def ensure_dependencies():
    try:
        from pptx import Presentation  # noqa: F401
    except ImportError as exc:
        raise SystemExit(
            "python-pptx is not installed. Run scripts/bootstrap_env.py first."
        ) from exc


def apply_run_style(run, font_name: str, font_size: int, rgb: tuple[int, int, int], bold: bool = False):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*rgb)


def add_title_textbox(slide, title: str, font_name: str, text_rgb: tuple[int, int, int]):
    from pptx.util import Inches

    box = slide.shapes.add_textbox(Inches(0.6), Inches(0.45), Inches(12.1), Inches(0.8))
    paragraph = box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    apply_run_style(run, font_name, 24, text_rgb, bold=True)
    return box


def add_bullets_box(slide, bullets: list[str], left: float, top: float, width: float, height: float, font_name: str, text_rgb: tuple[int, int, int]):
    from pptx.util import Inches

    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True

    if not bullets:
        bullets = ["No narrative content was provided for this slide."]

    first = True
    for bullet in bullets:
        paragraph = frame.paragraphs[0] if first else frame.add_paragraph()
        paragraph.level = 0
        run = paragraph.add_run()
        run.text = bullet
        apply_run_style(run, font_name, 18, text_rgb, bold=False)
        first = False
    return box


def add_table_shape(slide, headers: list[str], rows: list[list[str]], left: float, top: float, width: float, height: float, theme: dict[str, Any]):
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    row_count = max(2, len(rows) + 1)
    col_count = max(1, len(headers))
    table_shape = slide.shapes.add_table(
        row_count,
        col_count,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    table = table_shape.table
    primary = RGBColor(*hex_to_rgb(theme["primary_color"]))
    text_rgb = RGBColor(*hex_to_rgb(theme["text_color"]))

    for idx, header in enumerate(headers):
        cell = table.cell(0, idx)
        cell.text = header
        paragraph = cell.text_frame.paragraphs[0]
        if paragraph.runs:
            run = paragraph.runs[0]
        else:
            run = paragraph.add_run()
            run.text = header
        run.font.bold = True
        run.font.name = theme["font_family"]
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(255, 255, 255)
        cell.fill.solid()
        cell.fill.fore_color.rgb = primary

    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = value
            paragraph = cell.text_frame.paragraphs[0]
            run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            run.font.name = theme["font_family"]
            run.font.size = Pt(10)
            run.font.color.rgb = text_rgb


def add_footer(slide, text: str, font_name: str, text_rgb: tuple[int, int, int]):
    from pptx.util import Inches

    if not text:
        return

    box = slide.shapes.add_textbox(Inches(0.6), Inches(7.0), Inches(12.1), Inches(0.3))
    paragraph = box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    apply_run_style(run, font_name, 10, text_rgb)


def normalize_rows(row_dicts: list[dict[str, Any]], headers: list[str]) -> list[list[str]]:
    return [[str(row.get(header, "")) for header in headers] for row in row_dicts]


def render_title_slide(prs, slide_data: dict[str, Any], theme: dict[str, Any]):
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["background_color"]))
    background.line.fill.background()

    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.0), Inches(0.22), Inches(4.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["accent_color"]))
    accent.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(1.1), Inches(1.2), Inches(10.8), Inches(1.8))
    paragraph = title_box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = slide_data.get("title", "")
    apply_run_style(run, theme["font_family"], 28, hex_to_rgb(theme["text_color"]), bold=True)

    subtitle_box = slide.shapes.add_textbox(Inches(1.1), Inches(3.0), Inches(10.8), Inches(0.8))
    subtitle_paragraph = subtitle_box.text_frame.paragraphs[0]
    subtitle_run = subtitle_paragraph.add_run()
    subtitle_run.text = slide_data.get("subtitle", "")
    apply_run_style(subtitle_run, theme["font_family"], 18, hex_to_rgb(theme["text_color"]))


def render_schema_slide(prs, slide_data: dict[str, Any], payload: dict[str, Any], theme: dict[str, Any]):
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.5), Inches(12.0), Inches(0.12)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["accent_color"]))
    accent.line.fill.background()

    add_title_textbox(slide, slide_data["title"], theme["font_family"], hex_to_rgb(theme["text_color"]))

    intro_text = slide_data.get("intro_text", "")
    if intro_text:
        add_bullets_box(
            slide,
            [intro_text],
            0.7,
            1.1,
            12.0,
            0.5,
            theme["font_family"],
            hex_to_rgb(theme["text_color"]),
        )

    tables = slide_data.get("tables", [])
    if not tables:
        add_bullets_box(
            slide,
            ["No schema tables were configured for this slide."],
            0.7,
            1.5,
            12.0,
            1.0,
            theme["font_family"],
            hex_to_rgb(theme["text_color"]),
        )
        return

    available_height = 5.2
    per_table_height = available_height / len(tables)
    top = 1.5
    for table_name in tables:
        schema_rows = payload["tables"].get(table_name, [])
        headers = ["column_name", "data_type", "is_partition_key", "column_comment"]
        normalized = normalize_rows(schema_rows, headers)[:8]
        if not normalized:
            normalized = [["No schema rows returned", "", "", ""]]
        add_table_shape(slide, headers, normalized, 0.7, top, 12.0, per_table_height - 0.15, theme)
        top += per_table_height

    add_footer(slide, slide_data.get("speaker_notes", ""), theme["font_family"], hex_to_rgb(theme["text_color"]))


def render_table_slide(prs, slide_data: dict[str, Any], payload: dict[str, Any], theme: dict[str, Any]):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.5), Inches(12.0), Inches(0.12)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["accent_color"]))
    accent.line.fill.background()

    add_title_textbox(slide, slide_data["title"], theme["font_family"], hex_to_rgb(theme["text_color"]))

    takeaways = slide_data.get("takeaway_bullets", [])
    add_bullets_box(
        slide,
        takeaways,
        0.7,
        1.1,
        12.0,
        1.0,
        theme["font_family"],
        hex_to_rgb(theme["text_color"]),
    )

    query_name = slide_data["query"]
    query_data = payload["queries"].get(query_name)
    if not query_data:
        add_bullets_box(
            slide,
            [f"Query not found in payload: {query_name}"],
            0.7,
            2.0,
            12.0,
            1.0,
            theme["font_family"],
            hex_to_rgb(theme["text_color"]),
        )
        return

    headers = query_data["columns"]
    max_rows = int(slide_data.get("max_rows", 12))
    rows = normalize_rows(query_data["rows"][:max_rows], headers)
    if not rows:
        rows = [["No rows returned"] + [""] * (len(headers) - 1)]
    add_table_shape(slide, headers, rows, 0.7, 2.2, 12.0, 4.4, theme)
    add_footer(slide, slide_data.get("speaker_notes", ""), theme["font_family"], hex_to_rgb(theme["text_color"]))


def render_bullet_slide(prs, slide_data: dict[str, Any], theme: dict[str, Any]):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.5), Inches(12.0), Inches(0.12)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["accent_color"]))
    accent.line.fill.background()

    add_title_textbox(slide, slide_data["title"], theme["font_family"], hex_to_rgb(theme["text_color"]))
    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        add_bullets_box(
            slide,
            [subtitle],
            0.7,
            1.1,
            12.0,
            0.4,
            theme["font_family"],
            hex_to_rgb(theme["text_color"]),
        )

    add_bullets_box(
        slide,
        slide_data.get("bullets", []),
        0.9,
        1.8,
        11.6,
        4.8,
        theme["font_family"],
        hex_to_rgb(theme["text_color"]),
    )
    add_footer(slide, slide_data.get("speaker_notes", ""), theme["font_family"], hex_to_rgb(theme["text_color"]))


def main() -> None:
    ensure_dependencies()
    from pptx import Presentation
    from pptx.util import Inches

    args = parse_args()
    config_path = Path(args.config).resolve()
    payload_path = Path(args.payload).resolve()
    plan_path = Path(args.plan).resolve()

    base_config = load_yaml(config_path)
    payload = load_json(payload_path)
    context = payload["metadata"].get("variables", {})
    rendered_config = deep_render(base_config, context)
    plan = load_json(plan_path)

    theme = rendered_config.get("report", {}).get("theme", {})
    theme = {
        "primary_color": theme.get("primary_color", "0F4C81"),
        "accent_color": theme.get("accent_color", "F28E2B"),
        "text_color": theme.get("text_color", "1F1F1F"),
        "background_color": theme.get("background_color", "FFFFFF"),
        "font_family": theme.get("font_family", "Arial"),
    }

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for slide_data in plan.get("slides", []):
        slide_type = slide_data.get("type")
        if slide_type == "title":
            render_title_slide(prs, slide_data, theme)
        elif slide_type == "schema":
            render_schema_slide(prs, slide_data, payload, theme)
        elif slide_type == "table":
            render_table_slide(prs, slide_data, payload, theme)
        elif slide_type == "ai_bullets":
            render_bullet_slide(prs, slide_data, theme)
        else:
            raise ValueError(f"Unsupported slide type: {slide_type}")

    if args.output:
        output_path = Path(args.output).resolve()
    else:
        report = rendered_config.get("report", {})
        default_output = report.get("output_filename", "report.pptx")
        output_path = output_dir_from_config(config_path, rendered_config) / default_output

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    print(f"ppt={output_path}")


if __name__ == "__main__":
    main()
