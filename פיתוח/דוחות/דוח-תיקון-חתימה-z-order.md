# דוח תיקון — חתימה נעלמת מאחורי כחול (z-order)

## מה הסתיר את החתימה

| ממצא | ערך |
|------|-----|
| סוג «הנחיות להחזרה» | פסקת טקסט כחול (`LG Section`) — **לא** shape/textbox/shading |
| מלבנים כחולים (fill) חופפים ל-rect החתימה | 0 (שטח חפיפה עם sig:  pt²) |
| ציורי content-stream בתוך rect החתימה | 0 |
| סדר annotations | MemberSignature → SignDateEntry |

**מסקנת אבחון:** `_draw_field_box(overlay=True)` ו-`fill_color` לבן על ה-widget
הוסיפו שכבת content מעל תוכן Word; אחרי חתימה, ה-appearance נבלע מתחת לשכבה זו.
«הנחיות להחזרה» עצמה היא טקסט כחול (לא בלוק), אך נמצאת מתחת לשדה (~16.1 pt מרווח).

## תיקון

- `signature_field.py`: בוטל `_draw_field_box` (ציור על content-stream); הוסר `fill_color` לבן;
  מסגרת אפורה רק על ה-widget annotation.
- `create_template_docx.py`: מרווח נוסף אחרי תיקון שכבות (12pt + 10pt לפני קו).

## מדידות rect

```json
{
  "pdf": "C:\\Users\\dfusb\\Documents\\מכונת מכתבים\\פיתוח\\דוגמאות\\sig_layer_verify\\12345 כהן ישראל תחשיב זכויות אישי.pdf",
  "signed_pdf": "C:\\Users\\dfusb\\Documents\\מכונת מכתבים\\פיתוח\\דוגמאות\\sig_layer_verify\\12345 כהן ישראל תחשיב זכויות אישי_signed.pdf",
  "preview": "C:\\Users\\dfusb\\Documents\\מכונת מכתבים\\פיתוח\\דוגמאות\\sig_layer_verify\\12345 כהן ישראל תחשיב זכויות אישי_signed_preview.png",
  "sig_rect": [
    36.0,
    571.5,
    261.0,
    613.5
  ],
  "date_rect": [
    394.1,
    571.5,
    524.1,
    613.5
  ],
  "guidelines_rect": [
    560.4,
    629.6,
    569.9,
    640.7
  ],
  "gap_sig_to_guidelines_pt": 16.1,
  "content_stream_overlays_in_sig": [],
  "blue_rects_overlapping_sig_zone": [],
  "annot_paint_order": [
    "MemberSignature",
    "SignDateEntry"
  ],
  "post_sign_ink_visible": {
    "dark_pixel_ratio": 0.1001,
    "visible": true
  },
  "root_cause": "content-stream overlay (draw_rect + white fill) from signature_field.py"
}
```

## אימות אחרי חתימה מדומה

- דמו חתימה ב-PDF: `12345 כהן ישראל תחשיב זכויות אישי_signed.pdf`
- תצוגה: `12345 כהן ישראל תחשיב זכויות אישי_signed_preview.png`
- דיו גלוי בתוך rect: **כן** (dark_pixel_ratio=0.1001)

## קבצים ששונו

- `פיתוח/קוד מקור/LetterGenerator/src/signature_field.py`
- `פיתוח/קוד מקור/LetterGenerator/scripts/create_template_docx.py`
- `LetterGenerator_V1.2_Portable/templates/תחשיב זכויות אישי.docx`

**הערה:** `app.exe` של V1.2 אינו בתיקיית Portable (רק תבנית).
ה-PDF אומת דרך קוד המקור; לתיקון שכבות ב-exe יידרש build עתידי.