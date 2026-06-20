"""
dump_chp.py
-----------
מריץ Playwright, נכנס ל-CHP עם ברקוד אמיתי, ושומר את ה-HTML המלא לקובץ.
תריץ:  python dump_chp.py
התוצאה תישמר ב: chp_dump.html  ו- chp_tables.txt
"""

import asyncio
from playwright.async_api import async_playwright

# ─── שנה לפי הצורך ───
BARCODE = "8593868004713"          # בירה סטלה ארטואה
CITY    = "טירת כרמל"
CODE1   = "9000"
CODE2   = "2100"
# ──────────────────────

async def main():
    from urllib.parse import quote
    url = f"https://chp.co.il/{quote(CITY)}/{CODE1}/{CODE2}/{BARCODE}/0"
    print(f"Fetching: {url}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)   # headless=False כדי לראות מה קורה
        page = await browser.new_page(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
        )
        await page.goto(url, wait_until="networkidle", timeout=20000)
        await asyncio.sleep(2)   # תן לדף להיטען לגמרי

        # שמור HTML מלא
        html = await page.content()
        with open("chp_dump.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"✓ HTML saved to chp_dump.html ({len(html)} chars)")

        # חלץ טבלאות בלבד לקובץ נפרד (קל יותר לקריאה)
        tables_info = await page.evaluate("""() => {
            const tables = Array.from(document.querySelectorAll('table'));
            return tables.map((t, i) => {
                const headers = Array.from(t.querySelectorAll('th')).map(th => th.textContent.trim());
                const rows = Array.from(t.querySelectorAll('tr')).slice(0, 6).map(row => {
                    return Array.from(row.querySelectorAll('td,th')).map(cell => ({
                        text: cell.textContent.trim(),
                        html: cell.innerHTML.trim().substring(0, 200)
                    }));
                });
                // Get surrounding heading
                let heading = '';
                let el = t.parentElement;
                while (el && el !== document.body) {
                    const h = el.querySelector('h1,h2,h3,h4,strong,b,p');
                    if (h) { heading = h.textContent.trim().substring(0, 100); break; }
                    el = el.parentElement;
                }
                return { tableIndex: i, headers, nearestHeading: heading, rows };
            });
        }""")

        with open("chp_tables.txt", "w", encoding="utf-8") as f:
            for t in tables_info:
                f.write(f"\n{'='*60}\n")
                f.write(f"TABLE {t['tableIndex']}\n")
                f.write(f"Nearest heading: {t['nearestHeading']}\n")
                f.write(f"Headers: {t['headers']}\n")
                f.write("Rows (first 5):\n")
                for ri, row in enumerate(t['rows']):
                    f.write(f"  Row {ri}:\n")
                    for ci, cell in enumerate(row):
                        f.write(f"    Cell {ci}: text={repr(cell['text'])} | html={repr(cell['html'])}\n")

        print(f"✓ Tables saved to chp_tables.txt ({len(tables_info)} tables found)")

        await browser.close()
        print("\nשלח לי את תוכן קובץ chp_tables.txt")

asyncio.run(main())
