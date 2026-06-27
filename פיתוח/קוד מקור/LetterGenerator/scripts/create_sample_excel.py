"""Create sample Excel file for POC (xlsx)."""

from pathlib import Path

import pandas as pd


def create_sample_excel(output_path: Path) -> Path:
    # Columns A-T to match letter-based mapping in JSON
    data = {
        "A": [""],
        "B": [""],
        "C": [12345],           # קוד חבר
        "D": [""],
        "E": ["כהן"],           # שם משפחה
        "F": ["ישראל"],         # שם פרטי
        "G": [""],
        "H": [15],              # ותק
        "I": [5000],            # סכום קבוע
        "J": [75000],           # סכום לפי ותק
        "K": [""],
        "L": [2000],            # תוספת נפטרים (triggers death section)
        "M": [82000],           # סהכ
        "N": [""],
        "O": [1500],            # חוב מענק (triggers work grant section)
        "P": [80500],           # סהכ אחרי קיזוז
        "Q": [""],
        "R": ["חוב בנייה"],     # triggers building debt section
        "S": ["12-345-678901"], # חשבון בנק
        "T": ["נקלט.ת"],        # triggers new member section
    }
    df = pd.DataFrame(data)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


if __name__ == "__main__":
    base = Path(__file__).resolve().parent.parent / "samples"
    path = create_sample_excel(base / "sample_data.xlsx")
    print(f"Created: {path}")
