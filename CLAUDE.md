# דשבורד שאלוני מילואים — הוראות לעבודה עם Claude

## מה הפרויקט

דשבורד Streamlit חד-קובץ (`app.py`) לניתוח שאלוני רווחה של משפחות
משרתי מילואים. עברית RTL. מותאם לצילומי מסך לשקופיות Google Slides
(פונטים גדולים, full-width, מקרא צבעוני).

## הרצה מקומית

לחיצה כפולה על `run.bat` (Windows). הסקריפט סוגר תהליך ישן בפורט 8501,
מתקין תלויות אם צריך, ופותח את הדפדפן אוטומטית.

או: `python -m streamlit run app.py`.

## פריסה

Streamlit Community Cloud, פרטי (allowlist של Gmail אחד).
דחיפה ל-`main` ב-GitHub → המערכת מתעדכנת אוטומטית תוך ~30 שניות.

## עריכת קוד עם עברית — אזהרה חשובה

**לא** להשתמש ב-Edit tool ל-blocks ארוכים שמכילים עברית — נכשל
על regex של chars בטווח unicode עברי (0590-05FF).

**במקום**: כותבים patch script ב-Python (`_patch_*.py`) שמשתמש
ב-`str.replace` עם anchors ASCII כעוגן, ומריצים `python _patch_xxx.py`.
דוגמה אחרונה: `_patch_v4.py` (אפשר למחוק אחרי השימוש; ה-`.gitignore`
מתעלם ממנו ממילא).

## מבנה `app.py`

1. **CONFIGURATION** — צבעי מותג, `COL_PATTERNS`, `LABELS`, `RATING_COLS`,
   `MULTI_CHOICE_COLS`, `CATEGORICAL_COLS`, `FREE_TEXT_COLS`,
   `HEBREW_STOPWORDS`, `COLOR_SEQ`, `COLOR_CAT`
2. **CSS** — Heebo font, RTL, `.metric-card`, `.section-header`,
   `.quote-box`, `.insight-box`
3. **DATA LOADING** — `_process_df`, `load_data_from_bytes`, `load_default`
4. **FILTERS** — `apply_filters`
5. **HELPERS** — `pct_ge4`, `parse_multi`, `base_layout`, `_truncate`,
   `_extract_ngrams`
6. **CHART RENDERERS** — `render_rating_chart`, `render_multichoice_chart`,
   `render_categorical_chart`, `render_freetext_summary`
7. **HERO RENDERERS** — `render_rating_hero`, `render_mixed_hero`,
   `render_multichoice_hero` (גרפים מתכללים בראש כל עמוד)
8. **SIDEBAR** — `render_sidebar`: file_uploader → ניווט → פילטרים
9. **PAGES** — `render_*_page` לכל עמוד בסיידבר
10. **MAIN** — בוחר uploaded vs default ומפנה לעמוד הנכון

## קונבנציות גרפים

- כל גרף קורא ל-`fig.update_layout(**base_layout())` קודם, ואז overrides.
- צבעים קטגוריאליים: `COLOR_CAT` (12 צבעים מובחנים).
- צבעי מותג: `BRAND_BLUE="#1D2E6F"`, `BRAND_RED="#E15049"`.
- גרפים אופקיים עם הרבה פריטים: ציר Y בלי תוויות, מקרא בצד.
- גובה דינמי: `max(540, min(820, 50 * n_bars + 200))`.
- כל emoji בתחילת `section-header` (📊 📋 📈 📌 💬).

## עדכון קובץ נתונים

- **מקומית**: החלף את `data.xlsx` בתיקייה. רענן את הדף.
- **בענן**:
  1. push חדש של `data.xlsx` ל-GitHub (קבוע, גם משתמשים אחרים יראו)
  2. העלאה דרך הסיידבר באתר (זמני, רק לסשן הנוכחי שלך)

## הערות חיוניות

- אל תוסיף תלויות ל-`requirements.txt` בלי טעם טוב — Streamlit Cloud
  בונה image מחדש בכל deploy.
- אל תשבור התאמה בין מחרוזות עברית ב-`COL_PATTERNS` למה שמופיע באקסל —
  ההתאמה היא substring match. כשהאקסל משתנה, בדוק שכותרות העמודות
  עדיין מכילות את ה-patterns.
- `RATING_COLS` הן עמודות שצריך להמיר ל-numeric. אם מוסיפים שאלת דירוג
  חדשה, יש להוסיף אותה גם ל-`RATING_COLS` וגם ל-`COL_PATTERNS`/`LABELS`.
- שינויים שמעורבים בשם קובץ הברירה — לעדכן את `FILE_PATH` בתחילת `app.py`.
