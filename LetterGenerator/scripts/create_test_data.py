"""Create batch Excel for testing (15 rows, 2 intentional errors)."""

from pathlib import Path

import pandas as pd

from scripts.create_sample_excel import create_sample_excel


def _base_row(member_code: int, last: str, first: str, bank: str, p_value) -> dict:
    return {
        "A": "",
        "B": "",
        "C": member_code,
        "D": "",
        "E": last,
        "F": first,
        "G": "",
        "H": 10 + (member_code % 5),
        "I": 5000,
        "J": 50000,
        "K": "",
        "L": 1000 if member_code % 3 == 0 else 0,
        "M": 56000,
        "N": "",
        "O": 500 if member_code % 2 == 0 else 0,
        "P": p_value,
        "Q": "",
        "R": "" if member_code % 4 else "חוב בנייה",
        "S": bank,
        "T": "נקלט.ת" if member_code % 5 == 0 else "פעיל",
    }


def create_batch_excel(output_path: Path) -> Path:
    families = ["כהן", "לוי", "מזרחי", "אברהם", "דוד", "שלום", "ברק", "רוזן"]
    first_names = ["ישראל", "דנה", "יוסף", "מירי", "אבי", "נועה", "עמית", "הילה"]

    rows = []
    for i in range(15):
        code = 10001 + i
        last = families[i % len(families)]
        first = first_names[i % len(first_names)]
        bank = f"12-345-{678900 + i:06d}"
        p_value = 55500 + i * 100

        if i == 7:
            bank = ""
        if i == 11:
            p_value = "לא-מספר"

        rows.append(_base_row(code, last, first, bank, p_value))

    df = pd.DataFrame(rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


def create_xls_from_xlsx(xlsx_path: Path, xls_path: Path) -> Path:
    try:
        import xlwt
    except ImportError as exc:
        raise ImportError("Creating .xls test file requires xlwt: pip install xlwt") from exc

    df = pd.read_excel(xlsx_path, engine="openpyxl")
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet1")

    for col_idx, column in enumerate(df.columns):
        sheet.write(0, col_idx, str(column))

    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            if pd.isna(value):
                sheet.write(row_idx, col_idx, "")
            elif isinstance(value, (int, float)):
                sheet.write(row_idx, col_idx, value)
            else:
                sheet.write(row_idx, col_idx, str(value))

    xls_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(str(xls_path))
    return xls_path


def create_missing_column_excel(output_path: Path) -> Path:
    base = Path(__file__).resolve().parent.parent / "samples" / "sample_data.xlsx"
    if not base.exists():
        create_sample_excel(base)
    df = pd.read_excel(base, engine="openpyxl")
    df = df.iloc[:, :10]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path
