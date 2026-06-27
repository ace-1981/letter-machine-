# דוח Performance V1.2 — Word Pooling

**תאריך:** 2026-06-26  
**שינוי:** `WordComBatchSession` — Word נפתח **פעם אחת** לכל batch PDF  
**sleep:** `0` (נבדק ללא בעיות יציבות; override: `LETTER_GEN_WORD_SLEEP`)

---

## 1. סיכום

| מדד | V1.1 (לפני) | V1.2 (אחרי) | שיפור |
|-----|-------------|-------------|-------|
| ממוצע PDF (10 שורות) | 9.23s | 2.94s | **~68%** |
| ממוצע PDF (50 שורות) | 10.01s | 2.44s | **~76%** |
| ממוצע PDF (100 שורות) | 10.83s | 2.66s | **~75%** |
| **700 מכתבים (הערכה)** | **~126 דקות** | **~31 דקות** | **~75%** |

---

## 2. זמנים מפורטים

| שורות | סה"כ לפני | סה"כ אחרי | ממוצע לפני | ממוצע אחרי |
|-------|-----------|-----------|------------|------------|
| 10 | 92.3s | 29.4s | 9.23s | 2.94s |
| 50 | 500.6s | 122.0s | 10.01s | 2.44s |
| 100 | 1083.0s | 266.1s | 10.83s | 2.66s |

---

## 3. Word COM — כמה פעמים נפתח?

| שורות | `Dispatch('Word.Application')` ב-batch | הערה |
|-------|----------------------------------------|------|
| 10 | 1–2 | 2 ב-run הראשון (probe זמינות + session) |
| 50 | **1** | session יחיד |
| 100 | **1** | session יחיד |

**לפני V1.2:** Dispatch = **N** (מספר המכתבים) + sleep 0.5s × N.

---

## 4. sleep

| מצב | sleep |
|-----|-------|
| batch PDF | **0** — ללא המתנה בין מכתבים |
| preview / מכתב בודד | **0** (ברירת מחדל) |
| override | `LETTER_GEN_WORD_SLEEP=0.1` אם בעיות יציבות ב-preview |

נבדק עם sleep=0 — batch של 100 עבר ללא שגיאות.

---

## 5. WINWORD.EXE אחרי סיום

| מדידה | תוצאה |
|-------|--------|
| מיד אחרי `close()` | לעיתים **1** תהליך (Word עדיין נסגר) |
| אחרי 3–5 שניות | **0** תהליכים |

`Quit()` עובד; tasklist מיד אחרי סיום עלול להראות תהליך שעדיין נסגר. אין דליפה מתמשכת.

---

## 6. מה השתנה בקוד

### `pdf_converter.py`
- `WordComBatchSession` — Word פתוח לכל batch
- `_export_docx_to_pdf()` — פותח/סוגר **מסמך** בלבד
- `WordComPdfConverter.convert()` — ללא שינוי התנהגותי ל-preview (Word לכל מכתב)
- `time.sleep(0.5)` **הוסר** (ברירת מחדל 0)
- `_probe_word_available()` — בדיקת זמינות עם cache (פחות פתיחות מיותרות)

### `letter_generator.py`
- `generate_letters` (PDF): `PdfConverterFactory.create_batch()` + `finally: pdf_converter.close()`
- `generate_single_letter` / preview: **ללא שינוי** — `create()` רגיל

---

## 7. מה לא השתנה

עיצוב, חתימה, RTL, מצבי PDF/DOCX, JSON/DOCX runtime, מבנה Portable.

---

## 8. המלצות המשך

1. לבנות `app.exe` מחדש לפני הפצת V1.2 Portable
2. אם preview בודד לא יציב בחלק מהמחשבים: `LETTER_GEN_WORD_SLEEP=0.1` לסביבה בלבד
3. LibreOffice headless — אופציה עתידית לסביבות ללא Word

---

## נספח

- benchmark: `cursor/run_v1_2_word_pooling_benchmark.py`
- דוח בסיס: `cursor/דוח-Performance-Regression-PDF-Only.md`
