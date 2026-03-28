from __future__ import annotations

import argparse
import json
import platform
from pathlib import Path
from typing import Any

from common import deep_render, load_yaml, output_dir_from_config


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render an editable PPTX from a report payload and slide plan."
    )
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


def to_float(value: Any) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except ValueError:
        return None


def ensure_dependencies() -> None:
    try:
        from pptx import Presentation  # noqa: F401
    except ImportError as exc:
        raise SystemExit(
            "python-pptx is not installed. Run scripts/bootstrap_env.py first."
        ) from exc


def resolve_font_family(raw_theme: dict[str, Any]) -> str:
    system_name = platform.system().lower()
    if system_name == "darwin":
        return str(
            raw_theme.get("font_family_mac")
            or raw_theme.get("font_family")
            or "Aptos"
        )
    if system_name == "windows":
        return str(
            raw_theme.get("font_family_windows")
            or raw_theme.get("font_family")
            or "Aptos"
        )
    return str(
        raw_theme.get("font_family_other")
        or raw_theme.get("font_family")
        or "Aptos"
    )


def resolve_cjk_font_family(raw_theme: dict[str, Any]) -> str:
    system_name = platform.system().lower()
    if system_name == "darwin":
        return str(
            raw_theme.get("cjk_font_family_mac")
            or raw_theme.get("cjk_font_family")
            or "Hiragino Sans GB"
        )
    if system_name == "windows":
        return str(
            raw_theme.get("cjk_font_family_windows")
            or raw_theme.get("cjk_font_family")
            or "Microsoft YaHei"
        )
    return str(
        raw_theme.get("cjk_font_family_other")
        or raw_theme.get("cjk_font_family")
        or "Noto Sans CJK SC"
    )


def build_theme(raw_theme: dict[str, Any]) -> dict[str, Any]:
    chart_palette = raw_theme.get("chart_palette") or [
        "0F4C81",
        "F28E2B",
        "59A14F",
        "E15759",
        "76B7B2",
    ]
    return {
        "primary_color": raw_theme.get("primary_color", "0F4C81"),
        "accent_color": raw_theme.get("accent_color", "F28E2B"),
        "text_color": raw_theme.get("text_color", "1F1F1F"),
        "background_color": raw_theme.get("background_color", "FFFFFF"),
        "font_family": resolve_font_family(raw_theme),
        "cjk_font_family": resolve_cjk_font_family(raw_theme),
        "language": str(raw_theme.get("language", "zh-CN")),
        "title_font_size": int(raw_theme.get("title_font_size", 24)),
        "hero_title_font_size": int(raw_theme.get("hero_title_font_size", 28)),
        "subtitle_font_size": int(raw_theme.get("subtitle_font_size", 18)),
        "body_font_size": int(raw_theme.get("body_font_size", 16)),
        "table_header_font_size": int(raw_theme.get("table_header_font_size", 11)),
        "table_body_font_size": int(raw_theme.get("table_body_font_size", 10)),
        "footer_font_size": int(raw_theme.get("footer_font_size", 10)),
        "chart_palette": [str(color).lstrip("#") for color in chart_palette],
    }


def contains_cjk(text: str) -> bool:
    for char in text:
        codepoint = ord(char)
        if (
            0x3400 <= codepoint <= 0x4DBF
            or 0x4E00 <= codepoint <= 0x9FFF
            or 0xF900 <= codepoint <= 0xFAFF
            or 0x3040 <= codepoint <= 0x30FF
            or 0xAC00 <= codepoint <= 0xD7AF
        ):
            return True
    return False


def set_run_typefaces(
    run,
    latin_font_name: str,
    cjk_font_name: str,
    language: str,
) -> None:
    from pptx.oxml.ns import qn
    from pptx.oxml.xmlchemy import OxmlElement

    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", language)

    def upsert_font(tag: str, typeface: str) -> None:
        child = rPr.find(qn(tag))
        if child is None:
            child = OxmlElement(tag)
            rPr.insert_element_before(
                child,
                "a:hlinkClick",
                "a:hlinkMouseOver",
                "a:rtl",
                "a:extLst",
            )
        child.set("typeface", typeface)

    upsert_font("a:latin", latin_font_name)
    upsert_font("a:ea", cjk_font_name or latin_font_name)
    upsert_font("a:cs", cjk_font_name or latin_font_name)


def apply_run_style(
    run,
    font_name: str,
    font_size: int,
    rgb: tuple[int, int, int],
    bold: bool = False,
    cjk_font_name: str = "",
    language: str = "zh-CN",
):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    run_text = getattr(run, "text", "") or ""
    preferred_font = cjk_font_name if cjk_font_name and contains_cjk(run_text) else font_name
    run.font.name = preferred_font
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*rgb)
    set_run_typefaces(run, font_name, cjk_font_name or preferred_font, language)


def add_textbox(slide, left: float, top: float, width: float, height: float):
    from pptx.util import Inches

    return slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )


def add_title_textbox(slide, title: str, theme: dict[str, Any]):
    box = add_textbox(slide, 0.6, 0.45, 12.1, 0.8)
    paragraph = box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    apply_run_style(
        run,
        theme["font_family"],
        theme["title_font_size"],
        hex_to_rgb(theme["text_color"]),
        bold=True,
        cjk_font_name=theme["cjk_font_family"],
        language=theme["language"],
    )
    return box


def add_bullets_box(
    slide,
    bullets: list[str],
    left: float,
    top: float,
    width: float,
    height: float,
    theme: dict[str, Any],
    font_size: int | None = None,
):
    box = add_textbox(slide, left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True

    if not bullets:
        bullets = ["本页还没有填内容。"]

    first = True
    for bullet in bullets:
        paragraph = frame.paragraphs[0] if first else frame.add_paragraph()
        paragraph.level = 0
        run = paragraph.add_run()
        run.text = bullet
        apply_run_style(
            run,
            theme["font_family"],
            font_size or theme["body_font_size"],
            hex_to_rgb(theme["text_color"]),
            bold=False,
            cjk_font_name=theme["cjk_font_family"],
            language=theme["language"],
        )
        first = False
    return box


def add_table_shape(
    slide,
    headers: list[str],
    rows: list[list[str]],
    left: float,
    top: float,
    width: float,
    height: float,
    theme: dict[str, Any],
):
    from pptx.dml.color import RGBColor
    from pptx.util import Inches

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
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        if not paragraph.runs:
            run.text = header
        apply_run_style(
            run,
            theme["font_family"],
            theme["table_header_font_size"],
            (255, 255, 255),
            bold=True,
            cjk_font_name=theme["cjk_font_family"],
            language=theme["language"],
        )
        cell.fill.solid()
        cell.fill.fore_color.rgb = primary

    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = value
            paragraph = cell.text_frame.paragraphs[0]
            run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            apply_run_style(
                run,
                theme["font_family"],
                theme["table_body_font_size"],
                hex_to_rgb(theme["text_color"]),
                cjk_font_name=theme["cjk_font_family"],
                language=theme["language"],
            )

    return table_shape


def add_footer(slide, text: str, theme: dict[str, Any]):
    if not text:
        return

    box = add_textbox(slide, 0.6, 7.0, 12.1, 0.3)
    paragraph = box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    apply_run_style(
        run,
        theme["font_family"],
        theme["footer_font_size"],
        hex_to_rgb(theme["text_color"]),
        cjk_font_name=theme["cjk_font_family"],
        language=theme["language"],
    )


def normalize_rows(
    row_dicts: list[dict[str, Any]], headers: list[str]
) -> list[list[str]]:
    return [[str(row.get(header, "")) for header in headers] for row in row_dicts]


def add_slide_background(slide, prs, theme: dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(
        *hex_to_rgb(theme["background_color"])
    )
    background.line.fill.background()


def add_top_accent(slide, theme: dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.5), Inches(12.0), Inches(0.12)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["accent_color"]))
    accent.line.fill.background()


def resolve_chart_columns(
    query_data: dict[str, Any],
    category_column: str | None,
    value_columns: list[str] | None,
) -> tuple[str | None, list[str]]:
    columns = query_data.get("columns", [])
    rows = query_data.get("rows", [])
    if not columns or not rows:
        return None, []

    resolved_category = category_column or columns[0]
    resolved_values = list(value_columns or [])

    if not resolved_values:
        resolved_values = [
            column
            for column in columns
            if column != resolved_category
            and any(to_float(row.get(column)) is not None for row in rows)
        ][:3]

    return resolved_category, resolved_values


def build_chart_series(
    query_data: dict[str, Any],
    category_column: str,
    value_columns: list[str],
    max_points: int,
) -> tuple[list[str], list[tuple[str, list[float]]]]:
    rows = query_data.get("rows", [])[:max_points]
    categories: list[str] = []
    series_map = {column: [] for column in value_columns}

    for row in rows:
        categories.append(str(row.get(category_column, "")))
        for column in value_columns:
            value = to_float(row.get(column))
            series_map[column].append(0.0 if value is None else value)

    series = [(column, values) for column, values in series_map.items()]
    return categories, series


def chart_type_from_name(name: str):
    from pptx.enum.chart import XL_CHART_TYPE

    mapping = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "line_markers": XL_CHART_TYPE.LINE_MARKERS,
        "area": XL_CHART_TYPE.AREA,
        "pie": XL_CHART_TYPE.PIE,
    }
    return mapping.get(str(name).lower(), XL_CHART_TYPE.LINE_MARKERS)


def style_chart(chart, theme: dict[str, Any], chart_type_name: str) -> None:
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.util import Pt

    chart.has_title = False
    chart.has_legend = len(chart.series) > 1 and chart_type_name != "pie"
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        try:
            chart.legend.font.name = theme["cjk_font_family"] or theme["font_family"]
            chart.legend.font.size = Pt(theme["table_body_font_size"])
        except AttributeError:
            pass

    plot = chart.plots[0]
    plot.has_data_labels = False

    palette = theme["chart_palette"]
    for index, series in enumerate(chart.series):
        color = RGBColor(*hex_to_rgb(palette[index % len(palette)]))
        try:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color
        except AttributeError:
            pass
        try:
            series.format.line.color.rgb = color
        except AttributeError:
            pass

    try:
        category_axis = chart.category_axis
        category_axis.tick_labels.font.name = theme["cjk_font_family"] or theme["font_family"]
        category_axis.tick_labels.font.size = Pt(theme["table_body_font_size"])
    except AttributeError:
        pass

    try:
        value_axis = chart.value_axis
        value_axis.has_major_gridlines = True
        value_axis.tick_labels.font.name = theme["cjk_font_family"] or theme["font_family"]
        value_axis.tick_labels.font.size = Pt(theme["table_body_font_size"])
    except AttributeError:
        pass


def add_chart_shape(
    slide,
    query_data: dict[str, Any],
    slide_data: dict[str, Any],
    theme: dict[str, Any],
):
    from pptx.chart.data import ChartData
    from pptx.util import Inches

    category_column, value_columns = resolve_chart_columns(
        query_data,
        slide_data.get("category_column"),
        slide_data.get("value_columns"),
    )
    if not category_column or not value_columns:
        return None

    max_points = int(slide_data.get("max_points", 12))
    categories, series_data = build_chart_series(
        query_data, category_column, value_columns, max_points
    )
    if not categories or not series_data:
        return None

    chart_data = ChartData()
    chart_data.categories = categories
    for series_name, values in series_data:
        chart_data.add_series(series_name, values)

    chart_type_name = str(slide_data.get("chart_type", "line")).lower()
    chart_shape = slide.shapes.add_chart(
        chart_type_from_name(chart_type_name),
        Inches(0.7),
        Inches(2.2),
        Inches(12.0),
        Inches(4.25),
        chart_data,
    )
    chart = chart_shape.chart
    style_chart(chart, theme, chart_type_name)
    return chart_shape


def render_title_slide(prs, slide_data: dict[str, Any], theme: dict[str, Any]):
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, prs, theme)

    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.0), Inches(0.22), Inches(4.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["accent_color"]))
    accent.line.fill.background()

    title_box = add_textbox(slide, 1.1, 1.2, 10.8, 1.8)
    paragraph = title_box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = slide_data.get("title", "")
    apply_run_style(
        run,
        theme["font_family"],
        theme["hero_title_font_size"],
        hex_to_rgb(theme["text_color"]),
        bold=True,
        cjk_font_name=theme["cjk_font_family"],
        language=theme["language"],
    )

    subtitle_box = add_textbox(slide, 1.1, 3.0, 10.8, 0.8)
    subtitle_paragraph = subtitle_box.text_frame.paragraphs[0]
    subtitle_run = subtitle_paragraph.add_run()
    subtitle_run.text = slide_data.get("subtitle", "")
    apply_run_style(
        subtitle_run,
        theme["font_family"],
        theme["subtitle_font_size"],
        hex_to_rgb(theme["text_color"]),
        cjk_font_name=theme["cjk_font_family"],
        language=theme["language"],
    )


def render_schema_slide(
    prs, slide_data: dict[str, Any], payload: dict[str, Any], theme: dict[str, Any]
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, prs, theme)
    add_top_accent(slide, theme)

    add_title_textbox(slide, slide_data["title"], theme)

    intro_text = slide_data.get("intro_text", "")
    if intro_text:
        add_bullets_box(slide, [intro_text], 0.7, 1.1, 12.0, 0.5, theme)

    tables = slide_data.get("tables", [])
    if not tables:
        add_bullets_box(
            slide,
            ["这一页没有配置 schema 表。"],
            0.7,
            1.5,
            12.0,
            1.0,
            theme,
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
            normalized = [["没有查到 schema 结果", "", "", ""]]
        add_table_shape(
            slide,
            headers,
            normalized,
            0.7,
            top,
            12.0,
            per_table_height - 0.15,
            theme,
        )
        top += per_table_height

    add_footer(slide, slide_data.get("speaker_notes", ""), theme)


def render_table_slide(
    prs, slide_data: dict[str, Any], payload: dict[str, Any], theme: dict[str, Any]
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, prs, theme)
    add_top_accent(slide, theme)

    add_title_textbox(slide, slide_data["title"], theme)

    takeaways = slide_data.get("takeaway_bullets", [])
    add_bullets_box(slide, takeaways, 0.7, 1.1, 12.0, 1.0, theme)

    query_name = slide_data["query"]
    query_data = payload["queries"].get(query_name)
    if not query_data:
        add_bullets_box(
            slide,
            [f"没有在 payload 里找到查询结果：{query_name}"],
            0.7,
            2.0,
            12.0,
            1.0,
            theme,
        )
        return

    headers = query_data["columns"]
    max_rows = int(slide_data.get("max_rows", 12))
    rows = normalize_rows(query_data["rows"][:max_rows], headers)
    if not rows:
        rows = [["没有返回数据"] + [""] * (len(headers) - 1)]
    add_table_shape(slide, headers, rows, 0.7, 2.2, 12.0, 4.4, theme)
    add_footer(slide, slide_data.get("speaker_notes", ""), theme)


def render_chart_slide(
    prs, slide_data: dict[str, Any], payload: dict[str, Any], theme: dict[str, Any]
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, prs, theme)
    add_top_accent(slide, theme)

    add_title_textbox(slide, slide_data["title"], theme)

    takeaways = slide_data.get("takeaway_bullets", [])
    add_bullets_box(slide, takeaways, 0.7, 1.1, 12.0, 1.0, theme)

    query_name = slide_data["query"]
    query_data = payload["queries"].get(query_name)
    if not query_data:
        add_bullets_box(
            slide,
            [f"没有在 payload 里找到查询结果：{query_name}"],
            0.7,
            2.0,
            12.0,
            1.0,
            theme,
        )
        return

    chart_shape = add_chart_shape(slide, query_data, slide_data, theme)
    if chart_shape is None:
        add_bullets_box(
            slide,
            ["这一页没有足够的维度列或数值列，暂时没法生成原生图表。"],
            0.7,
            2.0,
            12.0,
            1.0,
            theme,
        )
        return

    add_footer(slide, slide_data.get("speaker_notes", ""), theme)


def render_bullet_slide(prs, slide_data: dict[str, Any], theme: dict[str, Any]):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_background(slide, prs, theme)
    add_top_accent(slide, theme)

    add_title_textbox(slide, slide_data["title"], theme)
    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        add_bullets_box(
            slide,
            [subtitle],
            0.7,
            1.1,
            12.0,
            0.4,
            theme,
            font_size=theme["subtitle_font_size"],
        )

    add_bullets_box(slide, slide_data.get("bullets", []), 0.9, 1.8, 11.6, 4.8, theme)
    add_footer(slide, slide_data.get("speaker_notes", ""), theme)


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

    theme = build_theme(rendered_config.get("report", {}).get("theme", {}))

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
        elif slide_type == "chart":
            render_chart_slide(prs, slide_data, payload, theme)
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
