# מחולל מכתבים — Letter Generator

אפליקציית Windows ליצירת PDF מכתבים מ-Excel ותבנית Word (עברית RTL).

## מבנה

- `LetterGenerator/` — קוד מקור
- `LetterGenerator_release/` — תיקיית הפצה portable (לאחר build)
- `cursor/` — דוחות ותיעוד

## הפעלה (פיתוח)

```powershell
cd LetterGenerator
pip install -r requirements.txt
python app.py
```

## Build portable

```powershell
cd LetterGenerator
python scripts/build_release.py
```

נוצרת תיקייה `LetterGenerator_release/` עם `app.exe` ותבניות חיצוניות ב-`templates/`.

**דרישות:** Microsoft Word מותקן (המרת DOCX→PDF).

## תבניות

עריכת טקסט/עיצוב: `templates/*.docx`  
מיפוי, תנאים, שם קובץ: `templates/*.json`  

שינוי תבנית אינו דורש build מחדש של `app.exe`.
