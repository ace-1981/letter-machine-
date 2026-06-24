# שמירה — סיכום Portable Build

**תאריך:** 24/06/2026

## תיקיית הפצה (מוכנה)

`C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_release`

```
LetterGenerator_release/
├── app.exe              (332.6 MB)
├── templates/
│   ├── תחשיב זכויות אישי.docx
│   └── תחשיב זכויות אישי.json
├── output/
└── README.txt
```

## מה נשמר בקוד (LetterGenerator/)

| קובץ | תפקיד |
|------|--------|
| `src/app_paths.py` | נתיבים יחסיים ל-app.exe |
| `src/startup_check.py` | בדיקת templates בהפעלה |
| `src/cli.py` | CLI: check / preview / info |
| `src/pdf_converter.py` | המרת PDF (Word COM) |
| `scripts/build_release.py` | build portable (debug → release) |
| `scripts/run_final_build_report.py` | בדיקות + דוח סופי |
| `scripts/verify_portable_design.py` | אימות עיצוב portable |
| `release/README.txt` | הוראות למשתמש |

## דוחות

- `cursor/דוח-Build-Portable.md` — דוח build סופי
- `cursor/דוח-ביניים-Build.md` — דוח ביניים

## הפעלה

1. העתק את כל `LetterGenerator_release` למיקום כלשהו.
2. דאבל-קליק על `app.exe`.
3. Word מותקן נדרש להמרת PDF.

## build מחדש (אם צריך)

```powershell
cd "C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator"
python scripts/build_release.py
python scripts/run_final_build_report.py
```

## הערות

- תבניות חיצוניות — עריכת DOCX/JSON בלי rebuild.
- `dist/app_debug.exe` — גרסת debug עם console (לא בתיקיית ההפצה).
- אין git repository בפרויקט.
