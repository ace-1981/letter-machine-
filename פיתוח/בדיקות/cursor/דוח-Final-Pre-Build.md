# דוח Final Pre-Build

**תאריך:** 2026-06-24T22:03:51
**מוכן ל-app.exe:** כן

## סיכום בדיקות

| בדיקה | סטטוס | פירוט |
|--------|--------|--------|
| JSON במצב מקורי | PASS | template_name=תחשיב זכויות אישי |
| גיבוי docx/json | PASS | `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\templates\backup_pre_build\תחשיב זכויות אישי.json` |
| Preview רגיל | PASS | `12345 כהן ישראל תחשיב זכויות אישי.pdf` |
| Batch 20 שורות | PASS | 18/20 הצלחות, 2 שגיאות |
| errors_report.csv | PASS | `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\pre_build_batch\errors_report.csv` |
| חתימה דיגיטלית | PASS | interactive=True |
| JSON משפיע | PASS | אושר בבדיקה נפרדת |
| PDF תקין | PASS | 1 עמודים |
| GUI | PASS | preview + batch buttons |
| אין טקסטי מכתב ב-JSON | PASS | מותר: template_name, ערך תנאי מ-Excel, anchor_text טכני |
| PDF converter | PASS | WordComPdfConverter |

## גיבויים

- JSON: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\templates\backup_pre_build\תחשיב זכויות אישי_20260624_220351.json`
- DOCX: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\templates\backup_pre_build\תחשיב זכויות אישי_20260624_220351.docx`

## Preview

- PDF: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\pre_build_output\12345 כהן ישראל תחשיב זכויות אישי.pdf`

## Batch

- Excel: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\pre_build_batch_20.xlsx`
- Output: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\pre_build_batch`
- שורות שגיאה: [9, 13]
