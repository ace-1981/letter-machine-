"""V1.2 Word pooling benchmark — compare batch PDF times after optimization."""

from __future__ import annotations

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

WORKSPACE = Path(__file__).resolve().parent.parent
DEV_ROOT = WORKSPACE / "LetterGenerator"
CONFIG = WORKSPACE / "LetterGenerator_V1.1_Portable" / "templates" / "תחשיב זכויות אישי.json"
OUT = WORKSPACE / "cursor" / "perf_v1_2_pooling"
REPORT = WORKSPACE / "cursor" / "דוח-Performance-V1.2-WordPooling.md"

BASELINE = {
    10: {"total_s": 92.3, "avg_s": 9.23},
    50: {"total_s": 500.6, "avg_s": 10.01},
    100: {"total_s": 1083.0, "avg_s": 10.83},
}
BASELINE_700_MIN = 126

sys.path.insert(0, str(DEV_ROOT))

from src.letter_generator import generate_letters  # noqa: E402
from src.pdf_converter import _WORD_POST_SLEEP_S, WordComBatchSession  # noqa: E402

_DISPATCH_COUNT = 0
_ORIG_DISPATCH = None


def _patch_dispatch_counter() -> None:
    global _ORIG_DISPATCH
    import win32com.client

    _ORIG_DISPATCH = win32com.client.Dispatch

    def counted(prog_id, *args, **kwargs):
        global _DISPATCH_COUNT
        if str(prog_id).lower() in ("word.application", "kwps.application"):
            _DISPATCH_COUNT += 1
        return _ORIG_DISPATCH(prog_id, *args, **kwargs)

    win32com.client.Dispatch = counted  # type: ignore[assignment]


def _winword_count() -> int:
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/NH"],
            capture_output=True,
            text=True,
            timeout=15,
        )
        lines = [ln for ln in out.stdout.splitlines() if "WINWORD.EXE" in ln.upper()]
        return len(lines)
    except Exception:
        return -1


def _make_excel(n: int, path: Path) -> Path:
    rows = []
    for i in range(n):
        cols = {chr(65 + j): "" for j in range(20)}
        cols.update({
            "C": 40001 + i, "E": "כהן", "F": "ישראל",
            "H": 10, "I": 1000, "J": 100, "L": 500, "M": 1600,
            "O": 200, "P": 1400, "S": f"12-345-{678900+i:06d}", "R": "", "T": "",
        })
        rows.append(cols)
    path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_excel(path, index=False)
    return path


def _run_batch(n: int) -> dict:
    global _DISPATCH_COUNT
    _DISPATCH_COUNT = 0
    excel = _make_excel(n, OUT / f"bench_{n}.xlsx")
    out_dir = OUT / f"run_{n}"
    if out_dir.exists():
        import shutil
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    before_word = _winword_count()
    t0 = time.perf_counter()
    result = generate_letters(excel, CONFIG, out_dir, output_format="pdf", pdf_preferred="word")
    total = time.perf_counter() - t0
    after_word = _winword_count()

    return {
        "n": n,
        "total_s": total,
        "avg_s": total / n,
        "success": result.success,
        "errors": len(result.errors),
        "word_dispatch_count": _DISPATCH_COUNT,
        "winword_before": before_word,
        "winword_after": after_word,
    }


def write_report(rows: list[dict]) -> None:
    lines = [
        "# דוח Performance V1.2 — Word Pooling",
        "",
        f"**תאריך:** {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"**שינוי:** `WordComBatchSession` — Word נפתח פעם אחת ל-batch",
        f"**sleep אחרי המרה (single-shot):** `{_WORD_POST_SLEEP_S}` שניות (env: `LETTER_GEN_WORD_SLEEP`)",
        f"**sleep ב-batch:** `0` (ללא המתנה בין מכתבים)",
        "",
        "---",
        "",
        "## השוואה ל-V1.1 (לפני pooling)",
        "",
        "| שורות | לפני — סה\"כ | אחרי — סה\"כ | שיפור | לפני — ממוצע | אחרי — ממוצע | שיפור |",
        "|-------|-------------|--------------|-------|-------------|--------------|-------|",
    ]
    for r in rows:
        b = BASELINE[r["n"]]
        imp_t = (1 - r["total_s"] / b["total_s"]) * 100
        imp_a = (1 - r["avg_s"] / b["avg_s"]) * 100
        lines.append(
            f"| {r['n']} | {b['total_s']:.1f}s | {r['total_s']:.1f}s | {imp_t:+.0f}% | "
            f"{b['avg_s']:.2f}s | {r['avg_s']:.2f}s | {imp_a:+.0f}% |"
        )

    avg_100 = next(r["avg_s"] for r in rows if r["n"] == 100)
    est_700 = avg_100 * 700 / 60
    imp_700 = (1 - est_700 / BASELINE_700_MIN) * 100
    lines.extend([
        "",
        f"**הערכת 700 מכתבים:** לפני ~{BASELINE_700_MIN} דקות → אחרי ~{est_700:.0f} דקות ({imp_700:+.0f}%)",
        "",
        "---",
        "",
        "## Word COM",
        "",
        "| שורות | Dispatch('Word.Application') בפועל | WINWORD אחרי סיום |",
        "|-------|-----------------------------------|-------------------|",
    ])
    for r in rows:
        lines.append(
            f"| {r['n']} | {r['word_dispatch_count']} | {r['winword_after']} |"
        )

    lines.extend([
        "",
        "**מצופה:** Dispatch = **1** לכל batch; WINWORD אחרי סיום = **0** (או זהה ללפני).",
        "",
        "---",
        "",
        "## sleep",
        "",
        f"- נבדק עם `LETTER_GEN_WORD_SLEEP={_WORD_POST_SLEEP_S}`",
        "- batch: ללא sleep בין המרות",
        "- אם בעתיד יש בעיות יציבות ב-preview בלבד: להגדיר `LETTER_GEN_WORD_SLEEP=0.1`",
        "",
        "---",
        "",
        "## קבצים",
        "",
        f"- קוד: `LetterGenerator/src/pdf_converter.py` — `WordComBatchSession`",
        f"- batch: `LetterGenerator/src/letter_generator.py` — `create_batch` + `finally close()`",
        f"- preview: ללא שינוי — `PdfConverterFactory.create()` (Word לכל מכתב)",
        "",
    ])
    REPORT.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    _patch_dispatch_counter()
    OUT.mkdir(parents=True, exist_ok=True)
    results = []
    for n in (10, 50, 100):
        print(f"Benchmark {n} rows...")
        r = _run_batch(n)
        results.append(r)
        print(
            f"  {r['total_s']:.1f}s avg={r['avg_s']:.2f}s "
            f"dispatch={r['word_dispatch_count']} winword_after={r['winword_after']}"
        )
    write_report(results)
    print(f"Report: {REPORT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
