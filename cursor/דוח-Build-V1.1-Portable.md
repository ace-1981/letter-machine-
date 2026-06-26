# דוח Build V1.1 — Portable

**תאריך:** 2026-06-26 01:38
**תיקייה:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.1_Portable`
**ZIP:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.1_Portable.zip`

## סיכום

בדיקות 1–12: **12/12 עברו**

| # | בדיקה | תוצאה |
|---|-------|--------|
| 1 | פתיחת app.exe בדאבל-קליק | עבר |
| 2 | מצב PDF — PDF בלבד | עבר |
| 3 | מצב DOCX — Word בלבד | עבר |
| 4 | Preview בשני המצבים | עבר |
| 5 | L=0 מסתיר שורת תוספת פטירה | עבר |
| 6 | O=0 מסתיר מענק עידוד וסה״כ אחרי קיזוז | עבר |
| 7 | ללא הערות — ללא כותרת הערות והבהרות | עבר |
| 8 | שדה חתימה אינטראקטיבי | עבר |
| 9 | JSON runtime test אחרי build | עבר |
| 10 | DOCX runtime test אחרי build | עבר |
| 11 | העתקת תיקייה למיקום חדש | עבר |
| 12 | בדיקה מנתיב עברי | עבר |

## בדיקות נוספות

- **מבנה תיקיית Portable:** עבר
- **ללא נתיבים קשיחים ב-src:** עבר
- **app_debug.exe smoke:** עבר

## מבנה Portable

```
LetterGenerator_V1.1_Portable/
├── app.exe
├── templates/
│   ├── תחשיב זכויות אישי.docx
│   └── תחשיב זכויות אישי.json
├── output/
└── README.txt
```

## עקרונות שנשמרו

- אין Installer
- DOCX ו-JSON חיצוניים ב-templates/
- שינוי תבנית אחרי build ללא rebuild
- ללא נתיבים קשיחים ב-src

## גדלים

- `app.exe`: 342.5 MB (359,121,252 bytes)
- תיקיית Portable: 342.5 MB
- `LetterGenerator_V1.1_Portable.zip`: 340.3 MB

## פירוט

### מבנה תיקיית Portable
- missing: []
- path: C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.1_Portable

### ללא נתיבים קשיחים ב-src
- hits: []

### פתיחת app.exe בדאבל-קליק

### מצב PDF — PDF בלבד
- files: ['12345 כהן ישראל תחשיב זכויות אישי.pdf']

### מצב DOCX — Word בלבד
- files: ['12345 כהן ישראל תחשיב זכויות אישי.docx']

### Preview בשני המצבים
- pdf: 12345 כהן ישראל תחשיב זכויות אישי.pdf
- docx: 12345 כהן ישראל תחשיב זכויות אישי.docx

### L=0 מסתיר שורת תוספת פטירה

### O=0 מסתיר מענק עידוד וסה״כ אחרי קיזוז

### ללא הערות — ללא כותרת הערות והבהרות

### שדה חתימה אינטראקטיבי
- info: {'found': True, 'field_name': 'MemberSignature', 'field_type': '/Sig', 'is_signature_field': True, 'rect': [36.0, 473.3761901855469, 261.0, 515.376220703125], 'rect_pdf': [36.0, 326.54376220703125, 261.0, 368.5437927246094], 'annot_flags': 4, 'locked': False, 'interactive': True, 'message': 'Valid interactive /Sig field'}

### JSON runtime test אחרי build
- pdf: 12345 כהן ישראל תחשיב זכויות אישי V11_PORTABLE_JSON.pdf

### DOCX runtime test אחרי build
- pdf: 12345 כהן ישראל תחשיב זכויות אישי.pdf

### העתקת תיקייה למיקום חדש
- location: C:\Users\dfusb\AppData\Local\Temp\lg_reloc_v11_k6flv0i1\LetterGenerator_V1.1_Portable
- stdout: OK  root=C:\Users\dfusb\AppData\Local\Temp\lg_reloc_v11_k6flv0i1\LetterGenerator_V1.1_Portable
    templates=C:\Users\dfusb\AppData\Local\Temp\lg_reloc_v11_k6flv0i1\LetterGenerator_V1.1_Portable\templates

### בדיקה מנתיב עברי
- location: C:\Users\dfusb\AppData\Local\Temp\בדיקת מחולל V1.1\LetterGenerator_V1.1_Portable
- stdout: OK  root=C:\Users\dfusb\AppData\Local\Temp\בדיקת מחולל V1.1\LetterGenerator_V1.1_Portable
    templates=C:\Users\dfusb\AppData\Local\Temp\בדיקת מחולל V1.1\LetterGenerator_V1.1_Portable\templates

### app_debug.exe smoke
