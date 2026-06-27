# דוח Build V1.2 — Portable

**תאריך:** 2026-06-26 12:09
**תיקייה:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.2_Portable`
**ZIP:** `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator_V1.2_Portable.zip`
**בדיקות חיצוניות:** `C:\Users\dfusb\Documents\מכונת מכתבים\cursor\build_tests\v1_2`

## סיכום: 10/10 בדיקות עברו

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

### PDF batch 50 שורות
- success: 50
- total: 50
- seconds: 150.0
- word_dispatch: 2

### PDF batch 100 שורות
- success: 100
- total: 100
- seconds: 275.1
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
- info: {'found': True, 'field_name': 'MemberSignature', 'field_type': '/Sig', 'is_signature_field': True, 'rect': [36.0, 605.7362060546875, 261.0, 647.7362060546875], 'rect_pdf': [36.0, 194.18377685546875, 261.0, 236.18377685546875], 'annot_flags': 4, 'locked': False, 'interactive': True, 'message': 'Valid interactive /Sig field'}

### JSON runtime
- file: 50001 כהן ישראל תחשיב זכויות אישי V12_PORTABLE_JSON.pdf

### DOCX runtime

### P לא מספרי — שגיאה + errors_report
- success: 2
- errors: [{'excel_row': 3, 'error': 'שורה 3 - סכום סופי לא תקין'}]
- report_rows: [{'excel_row': '3', 'error': 'שורה 3 - סכום סופי לא תקין'}]
