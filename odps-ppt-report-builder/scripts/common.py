from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import yaml


PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")


def skill_root() -> Path:
    return Path(__file__).resolve().parents[1]


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def load_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"Config must be a mapping: {path}")
    return data


def dump_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def parse_key_value_pairs(items: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in items:
        if "=" not in item:
            raise ValueError(f"Expected KEY=VALUE, got: {item}")
        key, value = item.split("=", 1)
        result[key.strip()] = value
    return result


def render_string(template: str, context: dict[str, Any]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        return str(context.get(key, match.group(0)))

    return PLACEHOLDER_PATTERN.sub(replace, template)


def deep_render(value: Any, context: dict[str, Any]) -> Any:
    if isinstance(value, str):
        return render_string(value, context)
    if isinstance(value, list):
        return [deep_render(item, context) for item in value]
    if isinstance(value, dict):
        return {key: deep_render(item, context) for key, item in value.items()}
    return value


def resolve_path(path_text: str | None, base: Path) -> Path | None:
    if not path_text:
        return None
    raw = Path(path_text)
    return raw if raw.is_absolute() else (base / raw).resolve()


def output_dir_from_config(config_path: Path, config: dict[str, Any]) -> Path:
    del config_path
    report = config.get("report", {})
    output_dir = report.get("output_dir", "output")
    raw = Path(output_dir)
    if raw.is_absolute():
        return raw
    return (Path.cwd() / raw).resolve()


def build_context(config: dict[str, Any], overrides: dict[str, str]) -> dict[str, Any]:
    context: dict[str, Any] = {}
    context.update(config.get("vars", {}))
    context.update(overrides)
    return context


def resolve_odps_binary(explicit: str | None = None) -> str:
    candidates: list[str] = []
    env_bin = os.environ.get("ODPSCMD_BIN")
    if env_bin:
        candidates.append(env_bin)
    if explicit:
        candidates.append(explicit)
    candidates.extend(["odpscmd", "odpscmd.exe", "odpscmd.cmd", "odpscmd.bat"])

    for candidate in candidates:
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise FileNotFoundError(
        "Could not find odpscmd. Set odps.bin in the config or ODPSCMD_BIN in the environment."
    )


def ensure_sql_statements(items: list[str], sql: str) -> str:
    statements = [item.strip().rstrip(";") for item in items if item and item.strip()]
    statements.append(sql.strip().rstrip(";"))
    return ";\n".join(statements) + ";\n"


def run_odps_sql(sql: str, odps_config: dict[str, Any]) -> subprocess.CompletedProcess[str]:
    odps_bin = resolve_odps_binary(odps_config.get("bin"))
    pre_sql = odps_config.get("pre_sql", [])
    full_sql = ensure_sql_statements(pre_sql, sql)

    with tempfile.NamedTemporaryFile("w", suffix=".sql", delete=False, encoding="utf-8") as handle:
        handle.write(full_sql)
        temp_path = Path(handle.name)

    command: list[str] = [odps_bin]
    if odps_config.get("config_file"):
        command.append(f"--config={odps_config['config_file']}")
    if odps_config.get("project"):
        command.append(f"--project={odps_config['project']}")
    if odps_config.get("endpoint"):
        command.append(f"--endpoint={odps_config['endpoint']}")
    command.extend(["-f", str(temp_path)])

    try:
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    finally:
        temp_path.unlink(missing_ok=True)


def parse_tabular_stdout(stdout: str, columns: list[str], field_delim: str) -> list[dict[str, str]]:
    expected = len(columns)
    rows: list[dict[str, str]] = []

    for raw_line in stdout.splitlines():
        line = raw_line.rstrip("\r")
        if not line.strip():
            continue
        parts = line.split(field_delim)
        if len(parts) != expected:
            continue
        rows.append({name: value for name, value in zip(columns, parts)})

    return rows


def schema_columns() -> list[str]:
    return [
        "table_schema",
        "table_name",
        "column_name",
        "ordinal_position",
        "data_type",
        "is_partition_key",
        "column_comment",
    ]


def default_schema_sql(config: dict[str, Any]) -> str:
    odps_config = config.get("odps", {})
    default_project = odps_config.get("project", "")
    tables = config.get("tables", [])
    table_names = [entry["name"] for entry in tables if entry.get("name")]
    if not table_names:
        return ""

    quoted_names = ", ".join(f"'{name}'" for name in table_names)
    project_clause = f" and table_schema = '{default_project}'" if default_project else ""

    return f"""
select
  table_schema,
  table_name,
  column_name,
  cast(ordinal_position as string) as ordinal_position,
  data_type,
  case when is_partition_key then 'true' else 'false' end as is_partition_key,
  nvl(column_comment, '') as column_comment
from Information_Schema.columns
where table_name in ({quoted_names}){project_clause}
order by table_name, cast(ordinal_position as bigint)
""".strip()


def rows_to_markdown(title: str, rows: list[dict[str, Any]], limit: int = 8) -> str:
    if not rows:
        return f"## {title}\n\nNo rows returned.\n"

    sample = rows[:limit]
    headers = list(sample[0].keys())
    lines = [f"## {title}", ""]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for row in sample:
        values = [str(row.get(header, "")).replace("\n", " ") for header in headers]
        lines.append("| " + " | ".join(values) + " |")
    if len(rows) > limit:
        lines.append("")
        lines.append(f"Only the first {limit} rows are shown here.")
    lines.append("")
    return "\n".join(lines)


def group_schema_rows(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        grouped.setdefault(row["table_name"], []).append(row)
    return grouped


def default_runtime_root(skill_name: str) -> Path:
    return (Path.home() / ".qoderwork-runtime" / "python" / skill_name).resolve()
