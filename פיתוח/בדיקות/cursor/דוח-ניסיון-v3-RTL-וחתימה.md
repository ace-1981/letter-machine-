# ניסיון v3 — Word-native RTL + חתימה

## שינוי עיקרי

התבנית נבנית עכשיו **ישירות ב-Microsoft Word** (COM) עם:
- `ReadingOrder = RTL`
- `Table.Direction = RTL`
- גופנים/שוליים/מסגרת כמו בדוגמה

שדה החתימה ממוקם לפי **עוגן טקסט** `חתימה (שדה לחתימה דיגיטלית)` ב-PDF.

## פלט

- עמודים: **1**
- PDF: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\fidelity_output\12345 כהן ישראל תחשיב זכויות אישי.pdf`
- תצוגה: `preview_page1.png`
- סימון שדה חתימה: `C:\Users\dfusb\Documents\מכונת מכתבים\LetterGenerator\samples\fidelity_output\preview_signature_marked.png`

## אימות חתימה

```json
{
  "found": true,
  "field_name": "MemberSignature",
  "field_type": "/Sig",
  "is_signature_field": true,
  "rect": null,
  "message": "Valid /Sig field detected"
}
```

## בדיקה ידנית

1. פתח את ה-PDF
2. Fill & Sign / חתימה
3. ודא שניתן לחתום בשדה `MemberSignature` בתוך המלבן התכלת
