# CHP מחיר גורף — Web App

## הרצה מקומית

```bash
pip install -r requirements.txt
playwright install chromium
python app.py
```
פתח: http://localhost:5000

---

## העלאה ל-Render (חינמי)

### שלב 1 — GitHub
1. צור repo חדש ב-GitHub
2. העלה את כל הקבצים

### שלב 2 — Render
1. היכנס ל-https://render.com
2. New → Web Service → חבר GitHub repo
3. הגדרות:
   - **Build Command:** `pip install -r requirements.txt && playwright install chromium && playwright install-deps chromium`
   - **Start Command:** `gunicorn app:app --workers 1 --timeout 3600 --bind 0.0.0.0:$PORT`
   - **Plan:** Free (מספיק לשימוש אישי)
4. Deploy!

### ⚠️ חשוב על Render Free Plan:
- השרת "ישן" אחרי 15 דקות ללא שימוש (הפעלה ראשונה איטית)
- לסריקות ארוכות (282 מוצרים × 3 ערים ≈ 17 דקות) — שמור החלון פתוח
- אין persistent storage — הקבצים נמחקים בין deployments (לא בעיה, הורדה מיידית)

---

## מבנה הפרויקט

```
chp-app/
├── app.py              # Flask server + scraper logic
├── templates/
│   └── index.html      # UI
├── requirements.txt
├── Procfile
└── render.yaml
```

---

## הוספת עיר

ב-UI לחץ "+ הוסף עיר" ומלא:
- **שם עיר** — בעברית כמו שמופיע ב-URL של CHP
- **קוד 1** — בדרך כלל 9000
- **קוד 2** — ייחודי לכל עיר

כדי למצוא קודים: חפש עיר ב-CHP ובדוק URL:
`https://chp.co.il/[שם]/[קוד1]/[קוד2]/7290.../0`
