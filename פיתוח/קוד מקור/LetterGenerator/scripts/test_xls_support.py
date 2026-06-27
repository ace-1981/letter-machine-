"""Verify .xls reading support via xlrd."""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from src.excel_reader import describe_excel_support, read_excel


def main() -> None:
    print(describe_excel_support())
    samples = ROOT / "samples"
    xlsx_path = samples / "sample_data.xlsx"

    if xlsx_path.exists():
        df = read_excel(xlsx_path)
        print(f"xlsx OK: {len(df)} rows, {len(df.columns)} columns")

    xls_path = samples / "sample_data.xls"
    if not xls_path.exists():
        print("No sample .xls file present. Place a real Excel 97-2003 .xls in samples/ to test.")
        return

    try:
        df_xls = read_excel(xls_path)
        print(f"xls OK: {len(df_xls)} rows, {len(df_xls.columns)} columns")
    except ImportError as exc:
        print(f"xls requires xlrd: {exc}")
    except Exception as exc:
        print(f"xls read failed: {exc}")


if __name__ == "__main__":
    main()
