#!/usr/bin/env python3
"""Sync the shared spec into the package: spec/{tools,errors.json} -> _specdata/.

The Python package ships its own copy of the tool/error contracts so that
``docxengine._spec`` works from an installed wheel without the repo checkout.
Run this script after any change under ``spec/`` and commit the result:

    python scripts/sync_spec.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parents[1]
SPEC_DIR = REPO_DIR / "spec"
DEST_DIR = REPO_DIR / "src" / "docxengine" / "_specdata"


def main() -> int:
    tools_src = SPEC_DIR / "tools"
    errors_src = SPEC_DIR / "errors.json"
    if not tools_src.is_dir() or not errors_src.is_file():
        print(f"spec sources missing under {SPEC_DIR}", file=sys.stderr)
        return 1
    tools_dest = DEST_DIR / "tools"
    if tools_dest.exists():
        shutil.rmtree(tools_dest)
    tools_dest.mkdir(parents=True)
    copied = ["errors.json"]
    shutil.copy2(errors_src, DEST_DIR / "errors.json")
    for path in sorted(tools_src.glob("*.json")):
        shutil.copy2(path, tools_dest / path.name)
        copied.append(f"tools/{path.name}")
    print(f"synced {len(copied)} spec files into {DEST_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
