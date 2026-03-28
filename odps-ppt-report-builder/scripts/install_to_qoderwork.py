from __future__ import annotations

import argparse
import shutil
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Install this skill into ~/.qoderwork/skills.")
    parser.add_argument(
        "--target-root",
        default="",
        help="Optional override for the QoderWork skills root. Defaults to ~/.qoderwork/skills.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source_dir = Path(__file__).resolve().parents[1]
    target_root = (
        Path(args.target_root).expanduser().resolve()
        if args.target_root
        else (Path.home() / ".qoderwork" / "skills").resolve()
    )
    target_root.mkdir(parents=True, exist_ok=True)
    target_dir = target_root / source_dir.name

    def ignore(_folder: str, names: list[str]) -> set[str]:
        ignored = {".venv", "output", "__pycache__"}
        return {name for name in names if name in ignored or name.endswith(".pyc")}

    shutil.copytree(source_dir, target_dir, dirs_exist_ok=True, ignore=ignore)
    print(f"installed={target_dir}")


if __name__ == "__main__":
    main()
