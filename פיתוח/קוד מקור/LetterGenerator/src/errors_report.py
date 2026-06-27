"""Write batch error report as CSV."""

from __future__ import annotations

import csv
from pathlib import Path


def write_errors_report(path: Path, errors: list[dict]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["excel_row", "error"])
        writer.writeheader()
        writer.writerows(errors)
    return path
