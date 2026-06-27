# דוח Build סופי — Portable App

## 1. תיקיית הפצה

**נוצרה:** כן
**נתיב:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_release`

## 2. קבצים בתיקייה

| קובץ/תיקייה | קיים |
|-------------|------|
| `app.exe` | כן |
| `templates/תחשיב זכויות אישי.json` | כן |
| `templates/תחשיב זכויות אישי.docx` | כן |
| `output` | כן |
| `README.txt` | כן |

## 3. app_debug.exe + smoke test

**עבר:** כן

## 4. app.exe — דאבל-קליק

**נפתח בהצלחה:** כן

## 5. שינוי JSON אחרי build (template_name → שם PDF)

**משפיע:** כן

## 6. שינוי DOCX אחרי build (טקסט → תוכן PDF)

**משפיע:** כן

## 7. העתקה למיקום חדש

**עובד:** כן

## 8. נתיב עברי

**תיקייה:** `בדיקת מחולל מכתבים`
**עובד:** כן

## 9. נתיבים קשיחים ל-C:\Users\dfusb

**לא נמצאו ב-src:** כן

## 10. Word לא זמין

**הודעת שגיאה ברורה (לא קריסה):** כן

הודעה שהוצגה בבדיקה:
`No PDF converter available. ... Microsoft Word is not available for COM automation.`
(exit code 1, ללא קריסה)

## 11. גדלים

- `app.exe`: 332.6 MB (348,738,659 bytes)
- תיקיית `LetterGenerator_release` כולה: 332.7 MB (348,877,270 bytes)

## הפעלה

1. העתיקו את כל `LetterGenerator_release` למיקום כלשהו (כולל נתיב עברי).
2. דאבל-קליק על `app.exe`.
3. בחרו Excel ותיקיית יעד (`output/` כברירת מחדל).

## קבצים חובה ליד app.exe

- `app.exe`
- `templates/תחשיב זכויות אישי.json`
- `templates/תחשיב זכויות אישי.docx`
- `output/`
- `README.txt`

## Microsoft Word

כן — Word מותקן נדרש להמרת DOCX ל-PDF.

## פירוט בדיקות

### Release folder structure
- folder_exists: True
- path: C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_release
- present: {'app.exe': 348738659, 'templates/תחשיב זכויות אישי.json': 1655, 'templates/תחשיב זכויות אישי.docx': 29234, 'output': 'dir', 'README.txt': 2123}
- missing: []
- ok: True

### No hardcoded C:\Users\dfusb paths in runtime src
- hits: []
- ok: True

### app_debug.exe smoke test
- ok: True
- pdf: 12345 כהן ישראל תחשיב זכויות אישי.pdf

### app.exe double-click launch
- ok: True
- pid_started: True
- note: GUI נפתח ונשאר פעיל — נסגר אוטומטית לאחר הבדיקה

### Runtime JSON (template_name)
- ok: True
- pdf: 12345 כהן ישראל תחשיב זכויות אישי FINAL_JSON_TEST.pdf

### Runtime DOCX (text change)
- ok: True
- pdf: 12345 כהן ישראל תחשיב זכויות אישי.pdf

### Relocated folder (ASCII temp path)
- location: C:\Users\dfusb\AppData\Local\Temp\lg_reloc_ascii_ha7ve9du\LetterGenerator_release
- startup_ok: True
- json_runtime_ok: True
- ok: True
- stdout: OK  root=C:\Users\dfusb\AppData\Local\Temp\lg_reloc_ascii_ha7ve9du\LetterGenerator_release
    templates=C:\Users\dfusb\AppData\Local\Temp\lg_reloc_ascii_ha7ve9du\LetterGenerator_release\templates

### Hebrew path (בדיקת מחולל מכתבים)
- location: C:\Users\dfusb\AppData\Local\Temp\בדיקת מחולל מכתבים\LetterGenerator_release
- startup_ok: True
- preview_ok: True
- ok: True
- stdout: OK  root=C:\Users\dfusb\AppData\Local\Temp\בדיקת מחולל מכתבים\LetterGenerator_release
    templates=C:\Users\dfusb\AppData\Local\Temp\בדיקת מחולל מכתבים\LetterGenerator_release\templates

### Word unavailable — clear error, no crash
- exit_code: 1
- method: python app.py (same runtime code as app.exe)
- message_snippet: Validation failed: No PDF converter available. ... Microsoft Word is not available for COM automation.
- ok: True
