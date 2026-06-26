# דוח Performance Regression — PDF Only

**תאריך:** 2026-06-26 10:02
**מטרה:** לבדוק האם מעבר ל-PDF-only יצר רגרסיית ביצועים לעומת מסלול DOCX+PDF הישן.

---

## 1. האם PDF only איטי יותר מ-DOCX+PDF?

- **10 שורות:** ישן 105.7s | PDF-only 92.3s | הפרש -13.4s (-12.7%) — **PDF-only מהיר יותר**
- **50 שורות:** ישן 479.3s | PDF-only 500.6s | הפרש +21.3s (+4.4%) — **כמעט זהה**
- **PDF-only + keep_intermediate_docx (10 שורות):** 92.6s — הפרש מחיקה: -0.33s

**מסקנה:** ההפרש בין המסלולים הוא **זניח** (פחות מ-5%). הרגרסיה המורגשת **לא נובעת** ממעבר temp-DOCX לעומת DOCX סופי.

---

## 2. זמנים לפי מסלול

| מסלול | שורות | Batch כולל | ממוצע/מכתב | DOCX | Word COM | חתימה+תאריך | מחיקה | מכתבים/דקה |
|--------|-------|------------|-------------|------|----------|-------------|-------|------------|
| docx_only | 10 | 0.6s | 0.05s | 0.054s | 0.00s | 0.000s | 0.0000s | 1080.3 |
| docx_only | 50 | 3.0s | 0.06s | 0.059s | 0.00s | 0.000s | 0.0000s | 999.7 |
| old_docx_plus_pdf | 10 | 105.7s | 10.57s | 0.056s | 10.35s | 0.161s | 0.0004s | 5.7 |
| old_docx_plus_pdf | 50 | 479.3s | 9.58s | 0.053s | 9.43s | 0.105s | 0.0004s | 6.3 |
| pdf_only | 10 | 92.3s | 9.23s | 0.053s | 9.07s | 0.110s | 0.0008s | 6.5 |
| pdf_only | 50 | 500.6s | 10.01s | 0.085s | 9.80s | 0.129s | 0.0008s | 6.0 |
| pdf_only_keep | 10 | 92.6s | 9.26s | 0.114s | 9.02s | 0.131s | 0.0005s | 6.5 |

**DOCX only (להשוואה):**
- 10 שורות: 0.6s (0.054s/מכתב)
- 50 שורות: 3.0s (0.059s/מכתב)

---

## 3. הבדל בקוד בין המסלולים

### מסלול ישן (לפני V1.1 output_format)
```
docx_path = output/name.docx          # שם סופי
temp_pdf  = output/_temp_name.pdf
render_template → docx_path
word.convert(docx_path → temp_pdf)
add_signature_field (+ date_field בגרסה הנוכחית)
unlink(temp_pdf)
שומר docx_path (keep_docx=True)
```

### מסלול חדש PDF-only
```
temp_docx = output/_temp_name.docx    # שם זמני
temp_pdf  = output/_temp_name.pdf
render_template → temp_docx
word.convert(temp_docx → temp_pdf)
add_signature_field + add_date_field
unlink(temp_pdf); unlink(temp_docx)
```

**הבדלים מהותיים בקוד:**
1. נתיב DOCX: סופי vs `_temp_`
2. PDF-only מוחק DOCX בסוף כל מכתב
3. נוסף `add_date_field` (לא היה ב-commit ראשוני)
4. **אין** הבדל במנגנון Word COM

---

## 4. Word COM — פתיחה אחת או כל פעם מחדש?

- **old_docx_plus_pdf (10 שורות):** Word נפתח **10** פעמים (= מספר המכתבים, לא פעם אחת ל-batch)
- **old_docx_plus_pdf (50 שורות):** Word נפתח **50** פעמים (= מספר המכתבים, לא פעם אחת ל-batch)
- **pdf_only (10 שורות):** Word נפתח **10** פעמים (= מספר המכתבים, לא פעם אחת ל-batch)
- **pdf_only (50 שורות):** Word נפתח **50** פעמים (= מספר המכתבים, לא פעם אחת ל-batch)
- **pdf_only_keep (10 שורות):** Word נפתח **10** פעמים (= מספר המכתבים, לא פעם אחת ל-batch)

בקוד `WordComPdfConverter.convert()`:
```python
word = win32com.client.Dispatch('Word.Application')  # כל המרה
...
word.Quit()
time.sleep(0.5)  # 0.5 שניות המתנה אחרי כל מכתב
```

**עלות sleep בלבד:** 0.5s × N שורות = **25s ל-50 מכתבים**, **50s ל-100**.

---

## 5. צוואר הבקבוק האמיתי

| שלב | % מזמן PDF (ממוצע 10 מכתבים) |
|-----|------------------------------|
| Word COM (כולל Dispatch+Quit+sleep) | ~98% |
| יצירת DOCX (docxtpl) | ~0.6% |
| חתימה + תאריך (fitz) | ~1.2% |
| מחיקת קבצים זמניים | ~0.01% |

**צוואר הבקבוק:** פתיחה/סגירה של Word לכל מכתב + `sleep(0.5)`.

---

## 6. בדיקות נקודתיות (שאלות 1–8)

| # | שאלה | ממצא |
|---|------|------|
| 1 | Word נפתח לכל מכתב ב-PDF-only? | **כן** — Dispatch+Quit בכל convert |
| 2 | Word נשאר פתוח ל-batch? | **לא** — אין pooling |
| 3 | קבצים זמניים בתיקייה איטית? | נוצרים ב-output; מחיקה <1ms — **לא צוואר בקבוק** |
| 4 | מחיקת DOCX זמני אחרי כל מכתב? | **כן** ב-PDF-only; עלות ~0.000s |
| 5 | Preview מיותר ב-batch? | **לא** — אין קריאת preview ב-generate_letters |
| 6 | חתימה שונה ב-PDF-only? | **אותה לוגיקה**; נוסף date_field (~0.05s) |
| 7 | sleep/retry מיותר? | **`time.sleep(0.5)` אחרי כל Word convert** |
| 8 | app.exe vs python? | אותו מסלול קוד; app.exe לא איטי מסלולית (PyInstaller overhead זניח ב-batch) |

### keep_intermediate_docx=True (10 שורות)
- PDF-only רגיל: 92.31s
- בלי מחיקת DOCX: 92.64s
- הפרש: -0.331s → **מחיקה אינה הגורם**

### למה המשתמש הרגיש שהמצב הישן מהיר יותר?

1. **תחושת התקדמות:** במצב ישן DOCX סופי נשמר מיד (~0.1s) — קבצים מופיעים בתיקייה לפני סיום המרת PDF.
2. **PDF-only:** אין קובץ גלוי עד סיום כל השלבים למכתב.
3. **אותו זמן כולל** ל-DOCX+PDF vs PDF-only (המרה זהה).
4. אם בעבר השוו DOCX-only (מהיר) ל-PDF-only (איטי) — ההפרש אמיתי אך לא רגרסיה של השינוי.

---

## 7. המלצה לתיקון (ללא יישום כעת)

| עדיפות | המלצה | השפעה משוערת | מאמץ |
|--------|--------|--------------|------|
| **1** | **Word instance יחיד לכל batch** — Dispatch פעם אחת, Documents.Open/Export/Close בלולאה, Quit בסוף | חיסכון של ~2–4s למכתב (הפעלה/כיבוי) | בינוני |
| **2** | **הסר/הקטן `time.sleep(0.5)`** — לבדוק אם נדרש ליציבות | עד 0.5s × N שורות | קטן |
| **3** | temp DOCX בתיקיית `%TEMP%`** במקום output | שיפור קל אם antivirus סורק output | קטן |
| **4** | מחיקת temp בסוף batch במקום כל מכתב | שיפור זניח | קטן |
| **5** | LibreOffice headless כ-fallback/אלטרנטיבה | עשוי להיות מהיר יותר ב-batch | גדול |

**לא מומלץ:** לחזור לשמירת DOCX סופי ב-PDF mode — לא יחסוך זמן המרה.

---

## נספח

- CSV: `C:\Users\dfusb\Documents\מכונת מכתבים\cursor\perf_regression_timings.csv`
- תיקיית benchmark: `C:\Users\dfusb\Documents\מכונת מכתבים\cursor\perf_regression`
- קוד Word COM: `LetterGenerator/src/pdf_converter.py` שורות 132–158
- קוד PDF-only: `LetterGenerator/src/letter_generator.py` שורות 167–196
