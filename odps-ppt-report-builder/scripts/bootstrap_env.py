from __future__ import annotations

import argparse
import subprocess
import sys
import venv
from pathlib import Path

from common import default_runtime_root


SKILL_NAME = "odps-ppt-report-builder"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create or update the shared Python runtime for this Qoder Work skill."
    )
    parser.add_argument(
        "--runtime-root",
        default="",
        help="Optional override for the shared runtime root.",
    )
    parser.add_argument(
        "--local-skill-env",
        action="store_true",
        help="Use a local .venv inside the skill folder instead of the shared runtime.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    skill_root = Path(__file__).resolve().parents[1]
    requirements = skill_root / "requirements.txt"
    if args.local_skill_env:
        runtime_root = skill_root
        venv_dir = skill_root / ".venv"
    else:
        runtime_root = (
            Path(args.runtime_root).expanduser().resolve()
            if args.runtime_root
            else default_runtime_root(SKILL_NAME)
        )
        venv_dir = runtime_root / "venv"

    if not venv_dir.exists():
        venv_dir.parent.mkdir(parents=True, exist_ok=True)
        builder = venv.EnvBuilder(with_pip=True)
        builder.create(venv_dir)

    if sys.platform.startswith("win"):
        python_bin = venv_dir / "Scripts" / "python.exe"
        pip_bin = venv_dir / "Scripts" / "pip.exe"
    else:
        python_bin = venv_dir / "bin" / "python"
        pip_bin = venv_dir / "bin" / "pip"

    subprocess.run([str(python_bin), "-m", "pip", "install", "--upgrade", "pip"], check=False)
    subprocess.run([str(pip_bin), "install", "-r", str(requirements)], check=True)

    print(f"runtime_root={runtime_root}")
    print(f"python={python_bin}")
    print(f"pip={pip_bin}")


if __name__ == "__main__":
    main()
