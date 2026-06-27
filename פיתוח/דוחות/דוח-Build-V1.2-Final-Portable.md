# דוח Build V1.2 — Portable סופי

**תאריך:** 2026-06-26 22:45  
**סטטוס:** **מוכן לאספקה**  
**תיקייה:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.2_Portable`  
**ZIP:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.2_Portable.zip` (~340 MB)  
**בדיקות חיצוניות:** `cursor\build_tests\v1_2` (לא בתוך ה-Portable)

## תוכן Portable (נקי)

```
LetterGenerator_V1.2_Portable/
├── app.exe
├── README.txt
├── templates/
│   ├── תחשיב זכויות אישי.docx
│   └── תחשיב זכויות אישי.json
└── output/
    └── README.txt
```

- אין `_batch`, `_bad_p`, `test`, `sample_data.xlsx` או קבצי בדיקה אחרים
- `output/` נקייה מלבד README

## תבנית מעודכנת (מאושרת)

| דרישה | סטטוס |
|--------|--------|
| RTL תקין בגוף המסמך (Word) | אושר — `w:bidi` + `jc=left` בפסקאות גוף |
| PDF מיושר לימין | אושר — `docx_rtl_pdf` לפני המרה |
| עמוד אחד לדוגמה רגילה | אושר — 1 עמוד |
| טבלת Word editable | אושר — טקסט בטבלה ניתן לעריכה |
| לא Read Only | אושר |
| ללא `docx_table_builder` | אושר — לא הוחזר |

## סיכום: 17/17 בדיקות עברו

| # | בדיקה | תוצאה |
|---|--------|--------|
| 1 | PDF batch 50 שורות | עבר |
| 2 | PDF batch 100 שורות | עבר |
| 3 | Word נפתח פעם אחת ל-batch | עבר |
| 4 | אין WINWORD אחרי 5 שניות | עבר |
| 5 | PDF mode (preview) | עבר |
| 6 | DOCX mode (preview) | עבר |
| 7 | חתימה אינטראקטיבית | עבר |
| 8 | JSON runtime | עבר |
| 8 | DOCX runtime | עבר |
| 9 | P לא מספרי — שגיאה + errors_report | עבר |
| 10 | דוגמה רגילה — עמוד אחד | עבר |
| 11 | גוף מיושר לימין (DOCX + PDF) | עבר |
| 12 | תנאי L/O — שורות מותנות | עבר |
| 13 | שדה תאריך | עבר |
| 14 | עריכת טקסט בטבלה → PDF בלי rebuild | עבר |
| 15 | תבנית לא Read Only | עבר |
| 16 | app.exe נפתח (check) | עבר |

### PDF batch 50 שורות
- success: 50
- total: 50
- seconds: 199.4
- word_dispatch: 2

### PDF batch 100 שורות
- success: 100
- total: 100
- seconds: 329.6
- word_dispatch: 1

### Word נפתח פעם אחת ל-batch
- word_dispatch: 1

### אין WINWORD אחרי 5 שניות
- winword_count: 0

### PDF mode (preview)
- file: 50001 כהן ישראל תחשיב זכויות אישי.pdf

### DOCX mode (preview)
- file: 50001 כהן ישראל תחשיב זכויות אישי.docx

### חתימה אינטראקטיבית
- info: {'found': True, 'field_name': 'MemberSignature', 'field_type': '/Sig', 'is_signature_field': True, 'rect': [36.0, 520.1761474609375, 261.0, 562.1761474609375], 'rect_pdf': [36.0, 279.74383544921875, 261.0, 321.74383544921875], 'annot_flags': 4, 'locked': False, 'interactive': True, 'message': 'Valid interactive /Sig field'}

### JSON runtime
- file: 50001 כהן ישראל תחשיב זכויות אישי V12_PORTABLE_JSON.pdf

### DOCX runtime

### P לא מספרי — שגיאה + errors_report
- success: 2
- errors: [{'excel_row': 3, 'error': 'שורה 3 - סכום סופי לא תקין'}]
- report_rows: [{'excel_row': '3', 'error': 'שורה 3 - סכום סופי לא תקין'}]

### דוגמה רגילה — עמוד אחד
- pages: 1

### גוף מיושר לימין (DOCX + PDF)
- docx_jc_bidi: True
- pdf_right_margin_pt: 25.4

### תנאי L/O — שורות מותנות
- ok: שורות פטירה (L) וקיזוז מענק (O) מופיעות ב-PDF

### שדה תאריך
- ok: `SignDateEntry` קיים ב-PDF

### עריכת טקסט בטבלה → PDF בלי rebuild
- ok: שינוי טקסט שורה בטבלה ב-DOCX משתקף ב-PDF ללא rebuild

### תבנית לא Read Only
- ok: `FILE_ATTRIBUTE_READONLY` = false

### app.exe נפתח (check)
- stdout: OK  root=C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.2_Portable
    templates=C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.2_Portable\templates
- stderr: 

## ביצועים (Word pooling)

| Batch | זמן | Word Dispatch |
|-------|-----|---------------|
| 50 שורות | 199.4 שניות | 2 |
| 100 שורות | 329.6 שניות | 1 |
| 20 שורות (בדיקת pooling) | — | 1 |

## קבצים שנבנו מ-

- `scripts/create_template_docx.py` — תבנית Word מעודכנת
- `src/docx_rtl_pdf.py` — התאמת RTL ל-PDF
- `scripts/build_v1_2_portable.py` → `LetterGenerator_V1.2_Portable/`
- `LetterGenerator_V1.2_Portable.zip` — 5 קבצים ב-ZIP
