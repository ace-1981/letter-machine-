"""Application entry point."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def main() -> int:
    if len(sys.argv) > 1:
        from src.cli import run_cli

        return run_cli(sys.argv[1:])
    from src.ui.main_window import run_app

    return run_app()


if __name__ == "__main__":
    raise SystemExit(main())
