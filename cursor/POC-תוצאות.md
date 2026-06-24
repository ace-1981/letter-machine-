# POC — תוצאות

## קבצים שנוצרו

| קובץ | נתיב |
|------|------|
| Excel לדוגמה | `LetterGenerator/samples/sample_data.xlsx` |
| DOCX זמני | `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\poc_output\12345 כהן ישראל תחשיב זכויות אישי.docx` |
| PDF סופי | `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\poc_output\12345 כהן ישראל תחשיב זכויות אישי.pdf` |
| JSON טכני | `LetterGenerator/templates/תחשיב זכויות אישי.json` |

## שם קובץ פלט

`12345 כהן ישראל תחשיב זכויות אישי.pdf`

פורמט: קוד חבר + שם משפחה + שם פרטי + שם המכתב (לפי האפיון).

## ממיר PDF

WordComPdfConverter

## אימות שדה חתימה (אוטומטי)

```json
{
  "found": true,
  "field_name": "MemberSignature",
  "field_type": "/Sig",
  "is_signature_field": true,
  "message": "Valid /Sig field detected"
}
```

שדה החתימה נוצר עם **pyHanko** כשדה AcroForm מסוג `/Sig` (לא מלבן רגיל).

### בדיקה ידנית ב-Adobe Reader / Foxit

1. פתח את `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\poc_output\12345 כהן ישראל תחשיב זכויות אישי.pdf`
2. ב-Adobe Reader: **Fill & Sign** (מילוי וחתימה) → אמור להופיע שדה חתימה בשם `MemberSignature`
3. ב-Foxit: **Fill & Sign** → Sign Document → בחר את שדה החתימה
4. חתום (חתימה דיגיטלית או ציור) — השדה אמור להיות אינטראקטיבי

## תלויות

ראה `LetterGenerator/requirements.txt`

## שלב הבא

לאחר אישור POC — בניית ממשק PySide6 לפי האפיון.
