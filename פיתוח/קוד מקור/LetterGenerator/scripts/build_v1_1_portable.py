"""Build LetterGenerator_V1.1_Portable (sets LG_RELEASE_DIR and runs build_release)."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RELEASE_NAME = "LetterGenerator_V1.1_Portable"


def main() -> int:
    env = {**os.environ, "LG_RELEASE_DIR": RELEASE_NAME, "PYTHONUTF8": "1"}
    return subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "build_release.py")],
        cwd=ROOT,
        env=env,
    ).returncode


if __name__ == "__main__":
    raise SystemExit(main())
