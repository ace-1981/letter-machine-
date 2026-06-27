# דוח בדיקות לפני GUI

**תוצאה כוללת:** עבר

## 1. שדה חתימה

- סטטוס אוטומטי: **PASS**
- PDF: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\poc_output\12345 כהן ישראל תחשיב זכויות אישי.pdf`

```json
{
  "found": true,
  "field_name": "MemberSignature",
  "field_type": "/Sig",
  "is_signature_field": true,
  "rect": [
    135.66000366210938,
    432.7015686035156,
    315.6600036621094,
    474.7015686035156
  ],
  "rect_pdf": [
    135.66000366210938,
    367.21843139648433,
    315.6600036621094,
    409.21843139648433
  ],
  "message": "Valid /Sig field detected"
}
```

**בדיקה ידנית:** פתח את ה-PDF ב-Adobe Reader או Foxit → Fill & Sign → חתום ב-MemberSignature


## 2. תמיכת xls

- סטטוס: **PASS**
- Supported formats: .xlsx (openpyxl), .xls (xlrd). For legacy .xls, xlrd 2.x reads Excel 97-2003 only (not .xlsb).
- קובץ: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\sample_data.xls`
- שורות: 1, עמודות: 20

## 3. validate שלילי

### עמודה חסרה
- תוצאה: עבר
  - Required Excel column L is missing (file has 10 columns).
  - Required Excel column M is missing (file has 10 columns).
  - Required Excel column O is missing (file has 10 columns).
  - Required Excel column P is missing (file has 10 columns).
  - Required Excel column S is missing (file has 10 columns).

### batch תקין structurally
- תוצאה: עבר

### שורה עם חשבון בנק חסר
- תוצאה: עבר
  - שורה 9 - חשבון בנק חסר

### שורה עם סכום לא תקין
- תוצאה: עבר
  - שורה 13 - סכום סופי לא תקין

## 4. batch (15 שורות)

- סטטוס: **PASS**
- סה"כ: 15, הצליחו: 13, שגיאות: 2
- PDFs שנוצרו: 13
- errors_report.csv: כן

- שורה 9: שורה 9 - חשבון בנק חסר
- שורה 13: שורה 13 - סכום סופי לא תקין