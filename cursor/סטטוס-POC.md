# עדכון כיוון — POC (24.06.2026)

הדוח המקורי אושר עקרונית. הפיתוח שונה לפי ההנחיות:

## מה בוצע

- [x] POC מקצה לקצה (ללא GUI)
- [x] Excel → שורה אחת → DOCX → PDF → שדה חתימה `/Sig` → שם קובץ C+E+F+מכתב
- [x] JSON טכני בלבד (ללא `texts`)
- [x] `validate` לפני הפקה
- [x] `pdf_converter.py` מופרד — Word COM + LibreOffice
- [x] `errors_report.csv` במודול batch
- [x] תמיכת xls דרך xlrd (ראה `scripts/test_xls_support.py`)

## תוצאות POC

ראה [`POC-תוצאות.md`](POC-תוצאות.md)

## מה הבא

- [ ] אישור ידני של שדה החתימה ב-Adobe/Foxit
- [ ] החלפת תבנית DOCX בנוסח הסופי שלך
- [ ] Excel אמיתי במקום sample
- [ ] בניית PySide6 UI

## הרצת POC

```powershell
cd "LetterGenerator"
pip install -r requirements.txt
python poc_run.py
```
