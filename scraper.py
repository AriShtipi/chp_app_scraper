"""
סריקת chp.co.il עבור כל מוצר/עיר, באמצעות Playwright.

תיקון מרכזי בקובץ הזה: parsePrice() בתוך extract_prices() עכשיו מעדיף
מספר שצמוד לסימן ₪ בתא, במקום לתפוס את המספר הראשון שמופיע בטקסט של התא.
זה היה ככל הנראה מקור הבאג של "מחיר 1.00" — אם בתא יש גם טקסט של כמות/יחידות
("1 יח'" וכו') לפני המחיר עצמו, הביטוי הרגולרי הישן היה תופס את ה-1 הזה
במקום את המחיר האמיתי. אם הבעיה עדיין חוזרת אחרי התיקון, הכי טוב יהיה
לשמור דוגמה של תא בעייתי (טקסט גולמי) כדי לדייק את ה-regex עוד יותר.
"""

import asyncio


def build_url(barcode, city):
    from urllib.parse import quote
    return f"https://chp.co.il/{quote(city['name'])}/{city['code1']}/{city['code2']}/{barcode}/0"


async def extract_prices(page):
    try:
        result = await page.evaluate("""() => {
            const out = { local: [], online: [] };

            function parsePrice(text) {
                if (!text) return null;
                const raw = text.replace(/,/g, '');
                // מעדיף מספר שצמוד לסימן ₪ (המחיר האמיתי) ולא כל מספר ראשון בתא
                // (לדוגמה תא שמכיל גם "1 יח'" וגם "12.90 ₪" - לא רוצים לתפוס את ה-1)
                let m = raw.match(/(\\d+(?:\\.\\d+)?)\\s*\\u20aa/);
                if (!m) m = raw.match(/\\u20aa\\s*(\\d+(?:\\.\\d+)?)/);
                if (!m) m = raw.match(/(\\d+(?:\\.\\d+)?)/);  // נפילה חזרה אם אין סימן ₪ כלל
                return m ? parseFloat(m[1]) : null;
            }

            const allH4 = Array.from(document.querySelectorAll('h4'));
            let onlineHeading = null;
            for (const h of allH4) {
                if (h.textContent.includes('תוצאות מחנויות באינטרנט')) {
                    onlineHeading = h;
                    break;
                }
            }

            function parseTable(table, isOnline) {
                const rows = Array.from(table.querySelectorAll('tr'));
                if (!rows.length) return;

                const headers = Array.from(rows[0].querySelectorAll('th')).map(th => th.textContent.trim());
                if (!headers.includes('מחיר')) return;

                const col = {};
                headers.forEach((h, i) => col[h] = i);

                for (let ri = 1; ri < rows.length; ri++) {
                    const cells = Array.from(rows[ri].querySelectorAll('td'));
                    if (cells.length === 1 || !cells.length) continue;

                    const getCell = name => (col[name] !== undefined && col[name] < cells.length)
                        ? cells[col[name]] : null;

                    const network = getCell('רשת')?.textContent?.trim() || '';
                    const store   = getCell('שם החנות')?.textContent?.trim() || '';

                    const priceCell = getCell('מחיר');
                    const price = priceCell ? parsePrice(priceCell.textContent.trim()) : null;

                    let sale = null;
                    const saleCell = getCell('מבצע');
                    if (saleCell) {
                        const saleTxt = saleCell.textContent.replace(/,/g, '').trim();
                        let m = saleTxt.match(/(\\d+(?:\\.\\d+)?)\\s*\\*/);
                        if (!m) m = saleTxt.match(/(\\d+(?:\\.\\d+)?)\\s*\\u20aa/);
                        if (m) sale = parseFloat(m[1]);
                    }

                    const effective = (sale !== null) ? sale : price;
                    if (effective && effective > 0 && effective < 10000) {
                        const entry = { network, store, price, sale, effective };
                        if (isOnline) out.online.push(entry);
                        else out.local.push(entry);
                    }
                }
            }

            const allTables = Array.from(document.querySelectorAll('table'));
            allTables.forEach(table => {
                let isOnline = false;
                if (onlineHeading) {
                    const pos = onlineHeading.compareDocumentPosition(table);
                    isOnline = !!(pos & 4);
                }
                const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent.trim());
                if (headers.includes('אתר אינטרנט')) isOnline = true;

                if (isOnline || (!isOnline && headers.includes("מחיר"))) {
                    parseTable(table, isOnline);
                }
            });

            out.local.sort((a, b) => a.effective - b.effective);
            out.online.sort((a, b) => a.effective - b.effective);
            return out;
        }""")
        return result
    except Exception:
        return {"local": [], "online": []}


async def scrape_one(context, semaphore, prod, city, job, delay):
    async with semaphore:
        url = build_url(prod["barcode"], city)
        page = await context.new_page()
        try:
            for attempt in range(3):
                try:
                    await page.goto(url, wait_until="networkidle", timeout=18000)
                    await asyncio.sleep(0.4)
                    data = await extract_prices(page)
                    prod["chp"][city["name"]] = data
                    break
                except Exception:
                    if attempt < 2:
                        await asyncio.sleep(2)
                    else:
                        prod["chp"][city["name"]] = {"local": [], "online": []}
        finally:
            await page.close()

        job["progress"] += 1
        job["log"] = f"{prod['name'][:35]} / {city['name']}"
        await asyncio.sleep(delay)


async def scrape_job(job, products, cities, delay, batch_size, on_finish):
    """
    מריץ את הסריקה לכל המוצרים/ערים, ומעדכן את job (dict, מועבר by-reference)
    תוך כדי ריצה. בסיום קורא ל-on_finish(products, cities) לשמירת התוצאה.
    """
    from playwright.async_api import async_playwright

    job["total"] = len(products) * len(cities)
    job["progress"] = 0
    job["status"] = "running"

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
                viewport={"width": 1280, "height": 800}
            )

            semaphore = asyncio.Semaphore(batch_size)
            tasks = [
                scrape_one(context, semaphore, prod, city, job, delay)
                for city in cities
                for prod in products
            ]
            await asyncio.gather(*tasks)
            await browser.close()

        on_finish(products, cities)
        job["status"] = "done"

    except Exception as e:
        job["status"] = "error"
        job["error"] = str(e)


def run_scrape_thread(job, products, cities, delay, batch_size, on_finish):
    asyncio.run(scrape_job(job, products, cities, delay, batch_size, on_finish))