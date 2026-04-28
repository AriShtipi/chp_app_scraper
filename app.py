import asyncio
import os
import re
import threading
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

jobs = {}

# ── Config (env vars as defaults) ────────────────────────────────────────────
VAT         = float(os.environ.get("VAT",          "1.18"))
DELAY_SEC   = float(os.environ.get("DELAY_SEC",   "1.0"))
BATCH_SIZE  = int  (os.environ.get("BATCH_SIZE",  "10"))
TOP_OFFLINE = int  (os.environ.get("TOP_OFFLINE", "2"))
TOP_ONLINE  = int  (os.environ.get("TOP_ONLINE",  "1"))

DEFAULT_CITIES = [
    {"name": "טירת כרמל", "code1": "9000", "code2": "2100"},
    {"name": "אריאל",      "code1": "9000", "code2": "3600"},
    {"name": "ביתר עילית", "code1": "9000", "code2": "3400"},
]


# ── Helpers ───────────────────────────────────────────────────────────────────
def is_online_store(entry: dict) -> bool:
    combined = ((entry.get("network") or "") + " " + (entry.get("store") or "")).lower()
    return "אונליין" in combined or "online" in combined


def parse_price(text: str):
    """Extract numeric price – prefer decimal match to avoid grabbing row numbers."""
    text = text.strip()
    m = re.search(r"\d+[.,]\d+", text)
    if m:
        return float(m.group().replace(",", "."))
    m = re.search(r"\d+", text)
    if m:
        return float(m.group())
    return None


# ── Excel parser ──────────────────────────────────────────────────────────────
def read_supplier_file(path: str) -> list:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    products = []
    current_category = ""

    for row in ws.iter_rows(values_only=True):
        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        d = row[3] if len(row) > 3 else None
        f = row[5] if len(row) > 5 else None

        if b and isinstance(b, (int, float)) and len(str(int(b))) >= 12:
            barcode = str(int(b))
            name = str(c or "").strip()
            price_list = float(d) if d and isinstance(d, (int, float)) else 0.0
            price_buy  = float(f) if f and isinstance(f, (int, float)) else price_list
            products.append({
                "category":          current_category,
                "barcode":           barcode,
                "name":              name,
                "price_list_ex_vat": round(price_list, 4),
                "price_buy_ex_vat":  round(price_buy, 4),
                "price_buy_inc_vat": round(price_buy * VAT, 2),
                "chp":               {},
            })
        elif b is None or b == 0:
            a = row[0] if len(row) > 0 else None
            cat_text = str(a or c or "").strip()
            if cat_text and len(cat_text) < 60 and not cat_text.startswith("יח'") and cat_text not in ("0", ""):
                current_category = cat_text

    return products


# ── CHP scraper ───────────────────────────────────────────────────────────────
async def extract_prices(page) -> list:
    results = []
    try:
        tables = await page.query_selector_all("table")
        for table in tables:
            headers = await table.query_selector_all("th")
            header_texts = [t.strip() for t in [await h.inner_text() for h in headers]]
            if "מחיר" not in header_texts:
                continue
            col = {t: i for i, t in enumerate(header_texts)}
            rows = await table.query_selector_all("tr")
            for row in rows[1:]:
                cells = await row.query_selector_all("td")
                if not cells:
                    continue

                def sc(idx):
                    return cells[idx] if idx < len(cells) else None

                network = ""
                store   = ""
                if col.get("רשת") is not None and sc(col["רשת"]):
                    network = (await sc(col["רשת"]).inner_text()).strip()
                if col.get("שם החנות") is not None and sc(col["שם החנות"]):
                    store = (await sc(col["שם החנות"]).inner_text()).strip()

                price, sale = None, None

                if col.get("מחיר") is not None:
                    el = sc(col["מחיר"])
                    if el:
                        price = parse_price(await el.inner_text())

                if col.get("מבצע") is not None:
                    el = sc(col["מבצע"])
                    if el:
                        txt = await el.inner_text()
                        m = re.search(r"\*\s*(\d+[.,]\d+)", txt) or re.search(r"\*\s*(\d+)", txt)
                        if m:
                            sale = float(m.group(1).replace(",", "."))

                effective = sale if sale else price
                if effective and effective > 0:
                    results.append({
                        "network":   network,
                        "store":     store,
                        "price":     price,
                        "sale":      sale,
                        "effective": effective,
                    })
    except Exception:
        pass
    return results


async def scrape_one(semaphore, context, prod, city, job, delay):
    """Scrape one (product × city) pair under a semaphore."""
    async with semaphore:
        from urllib.parse import quote
        url = (f"https://chp.co.il/{quote(city['name'])}/"
               f"{city['code1']}/{city['code2']}/{prod['barcode']}/0")

        page = await context.new_page()
        try:
            prices = []
            for attempt in range(3):
                try:
                    await page.goto(url, wait_until="networkidle", timeout=15000)
                    await asyncio.sleep(0.4)
                    prices = await extract_prices(page)
                    break
                except Exception:
                    if attempt < 2:
                        await asyncio.sleep(2)
        finally:
            await page.close()

        online  = sorted([p for p in prices if     is_online_store(p)], key=lambda x: x["effective"])
        offline = sorted([p for p in prices if not is_online_store(p)], key=lambda x: x["effective"])

        prod["chp"][city["name"]] = {
            "online":  online [:TOP_ONLINE],
            "offline": offline[:TOP_OFFLINE],
        }
        job["progress"] += 1
        job["log"] = f"{prod['name'][:35]} / {city['name']}"
        await asyncio.sleep(delay)


async def scrape_job(job_id: str, products: list, cities: list,
                     batch_size: int, delay: float):
    from playwright.async_api import async_playwright

    job = jobs[job_id]
    job["total"]    = len(products) * len(cities)
    job["progress"] = 0
    job["status"]   = "running"

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            context = await browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
            )

            # Init chp dict for every product
            for prod in products:
                prod["chp"] = {}

            semaphore = asyncio.Semaphore(batch_size)
            tasks = [
                scrape_one(semaphore, context, prod, city, job, delay)
                for prod in products
                for city in cities
            ]
            await asyncio.gather(*tasks)
            await browser.close()

        result_path = OUTPUT_DIR / f"{job_id}.xlsx"
        write_results(products, cities, str(result_path))
        job["status"]      = "done"
        job["result_path"] = str(result_path)

    except Exception as e:
        job["status"] = "error"
        job["error"]  = str(e)


def run_scrape_thread(job_id, products, cities, batch_size, delay):
    asyncio.run(scrape_job(job_id, products, cities, batch_size, delay))


# ── Excel writer ──────────────────────────────────────────────────────────────
def write_results(products: list, cities: list, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "תוצאות CHP"
    ws.sheet_view.rightToLeft = True

    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    BLUE        = "1a73e8"
    WHITE       = "FFFFFF"
    OFFLINE_CLR = ["DCE8FB", "D4F0E8", "FEF3D0", "F3E8FF", "FFE8E8"]
    ONLINE_CLR  = ["B3D1F7", "A8E6D8", "FDE8A0", "E5D4FF", "FFD4D4"]
    CAT_BG      = "E8F0FE"

    def cell(r, c, val="", bold=False, bg=None, fg="000000",
             align="right", fmt=None, wrap=False):
        cl = ws.cell(row=r, column=c, value=val)
        cl.font = Font(name="Arial", bold=bold, color=fg, size=10)
        if bg:
            cl.fill = PatternFill("solid", fgColor=bg)
        cl.alignment = Alignment(
            horizontal=align, vertical="center",
            wrap_text=wrap, readingOrder=2
        )
        cl.border = border
        if fmt:
            cl.number_format = fmt
        return cl

    FIXED = 5
    vat_pct = int(round((VAT - 1) * 100))

    # Columns per city: TOP_OFFLINE*2 + 1(offline cheapest) + TOP_ONLINE*2 + 1(online cheapest)
    COLS_OFFLINE = TOP_OFFLINE * 2 + 1
    COLS_ONLINE  = TOP_ONLINE  * 2 + 1
    COLS_CITY    = COLS_OFFLINE + COLS_ONLINE

    # Map city → (start_col, offline_start, online_start)
    city_layout = {}
    col = FIXED + 1
    for city in cities:
        city_layout[city["name"]] = {
            "start":          col,
            "offline_start":  col,
            "online_start":   col + COLS_OFFLINE,
        }
        col += COLS_CITY
    TOTAL_COLS = col - 1

    # ── Row 1: Title ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS)
    c = ws.cell(row=1, column=1, value="📊  השוואת מחירי CHP — תוצאות סריקה")
    c.font      = Font(name="Arial", bold=True, size=13, color=WHITE)
    c.fill      = PatternFill("solid", fgColor=BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    ws.row_dimensions[1].height = 26

    # ── Row 2: Fixed headers + city merged ──
    fixed_headers = [
        "קטגוריה",
        "שם פריט",
        "ברקוד",
        f"מחיר קנייה\n(ללא מע\"מ)",
        f"מחיר קנייה\n(+מע\"מ {vat_pct}%)",
    ]
    for ci, h in enumerate(fixed_headers, 1):
        cell(2, ci, h, bold=True, bg=BLUE, fg=WHITE, align="center", wrap=True)

    for i, city in enumerate(cities):
        layout = city_layout[city["name"]]
        ws.merge_cells(
            start_row=2, start_column=layout["start"],
            end_row=2,   end_column=layout["start"] + COLS_CITY - 1
        )
        c = ws.cell(row=2, column=layout["start"], value=city["name"])
        c.font      = Font(name="Arial", bold=True, size=11, color="222222")
        c.fill      = PatternFill("solid", fgColor=OFFLINE_CLR[i % len(OFFLINE_CLR)])
        c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        c.border    = border
    ws.row_dimensions[2].height = 22

    # ── Row 3: Offline / Online sub-headers ──
    for ci in range(1, FIXED + 1):
        cell(3, ci, "", bg="F0F4FF")

    for i, city in enumerate(cities):
        layout   = city_layout[city["name"]]
        off_bg   = OFFLINE_CLR[i % len(OFFLINE_CLR)]
        on_bg    = ONLINE_CLR [i % len(ONLINE_CLR)]

        ws.merge_cells(
            start_row=3, start_column=layout["offline_start"],
            end_row=3,   end_column=layout["offline_start"] + COLS_OFFLINE - 1
        )
        c = ws.cell(row=3, column=layout["offline_start"], value="🏪 לא אונליין")
        c.font      = Font(name="Arial", bold=True, size=10, color="222222")
        c.fill      = PatternFill("solid", fgColor=off_bg)
        c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        c.border    = border

        ws.merge_cells(
            start_row=3, start_column=layout["online_start"],
            end_row=3,   end_column=layout["online_start"] + COLS_ONLINE - 1
        )
        c = ws.cell(row=3, column=layout["online_start"], value="🌐 אונליין")
        c.font      = Font(name="Arial", bold=True, size=10, color="222222")
        c.fill      = PatternFill("solid", fgColor=on_bg)
        c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        c.border    = border
    ws.row_dimensions[3].height = 20

    # ── Row 4: Per-column headers ──
    for ci in range(1, FIXED + 1):
        cell(4, ci, "", bg="F0F4FF")

    for i, city in enumerate(cities):
        layout  = city_layout[city["name"]]
        off_bg  = OFFLINE_CLR[i % len(OFFLINE_CLR)]
        on_bg   = ONLINE_CLR [i % len(ONLINE_CLR)]

        c = layout["offline_start"]
        for n in range(TOP_OFFLINE):
            cell(4, c,   f"#{n+1} שם עסק", bold=True, bg=off_bg, align="center")
            cell(4, c+1, f"#{n+1} מחיר",   bold=True, bg=off_bg, align="center")
            c += 2
        cell(4, c, "הכי זול\nלא אונליין", bold=True, bg=off_bg, align="center", wrap=True)

        c = layout["online_start"]
        for n in range(TOP_ONLINE):
            cell(4, c,   f"#{n+1} שם (אונליין)", bold=True, bg=on_bg, align="center")
            cell(4, c+1, f"#{n+1} מחיר",         bold=True, bg=on_bg, align="center")
            c += 2
        cell(4, c, "הכי זול\nאונליין", bold=True, bg=on_bg, align="center", wrap=True)

    ws.row_dimensions[4].height = 22

    # ── Data rows ──
    row_num  = 5
    last_cat = ""
    for prod in products:
        if prod["category"] != last_cat and prod["category"]:
            ws.merge_cells(
                start_row=row_num, start_column=1,
                end_row=row_num,   end_column=TOTAL_COLS
            )
            c = ws.cell(row=row_num, column=1, value=f"  {prod['category']}")
            c.font      = Font(name="Arial", bold=True, size=10, color="1a3a6b")
            c.fill      = PatternFill("solid", fgColor=CAT_BG)
            c.alignment = Alignment(horizontal="right", vertical="center", readingOrder=2)
            c.border    = border
            ws.row_dimensions[row_num].height = 18
            row_num += 1
            last_cat = prod["category"]

        cell(row_num, 1, prod["category"])
        cell(row_num, 2, prod["name"])
        cell(row_num, 3, prod["barcode"])
        cell(row_num, 4, prod["price_buy_ex_vat"],  fmt="#,##0.00")
        cell(row_num, 5, prod["price_buy_inc_vat"],  fmt="#,##0.00", bold=True)

        for i, city in enumerate(cities):
            layout     = city_layout[city["name"]]
            chp_data   = prod["chp"].get(city["name"], {"offline": [], "online": []})
            offline_list = chp_data.get("offline", [])
            online_list  = chp_data.get("online",  [])
            off_bg = OFFLINE_CLR[i % len(OFFLINE_CLR)]
            on_bg  = ONLINE_CLR [i % len(ONLINE_CLR)]

            # Offline columns
            c = layout["offline_start"]
            min_offline = None
            for n in range(TOP_OFFLINE):
                if n < len(offline_list):
                    e     = offline_list[n]
                    label = e.get("store") or e.get("network") or "—"
                    pval  = e["effective"]
                    if min_offline is None or pval < min_offline:
                        min_offline = pval
                    cell(row_num, c,   label, bg=off_bg)
                    cell(row_num, c+1, pval,  bg=off_bg, fmt="#,##0.00")
                else:
                    cell(row_num, c,   "—", bg=off_bg)
                    cell(row_num, c+1, "—", bg=off_bg)
                c += 2

            # Offline cheapest summary
            if min_offline is not None:
                inc = prod["price_buy_inc_vat"]
                pct = ((inc - min_offline) / inc * 100) if inc else None
                cl  = cell(row_num, c, min_offline, fmt="#,##0.00", bold=True)
                if pct is not None:
                    if pct >= 25:
                        cl.fill = PatternFill("solid", fgColor="D4EDDA")
                        cl.font = Font(name="Arial", bold=True, color="155724", size=10)
                    elif pct >= 10:
                        cl.fill = PatternFill("solid", fgColor="FFF3CD")
                        cl.font = Font(name="Arial", bold=True, color="856404", size=10)
                    else:
                        cl.fill = PatternFill("solid", fgColor="F8D7DA")
                        cl.font = Font(name="Arial", bold=True, color="721c24", size=10)
            else:
                cell(row_num, c, "לא נמצא", fg="999999", bg=off_bg)

            # Online columns
            c = layout["online_start"]
            min_online = None
            for n in range(TOP_ONLINE):
                if n < len(online_list):
                    e     = online_list[n]
                    label = e.get("store") or e.get("network") or "—"
                    pval  = e["effective"]
                    if min_online is None or pval < min_online:
                        min_online = pval
                    cell(row_num, c,   label, bg=on_bg)
                    cell(row_num, c+1, pval,  bg=on_bg, fmt="#,##0.00")
                else:
                    cell(row_num, c,   "—", bg=on_bg)
                    cell(row_num, c+1, "—", bg=on_bg)
                c += 2

            # Online cheapest summary
            if min_online is not None:
                cell(row_num, c, min_online, fmt="#,##0.00", bold=True, bg=on_bg)
            else:
                cell(row_num, c, "לא נמצא", fg="999999", bg=on_bg)

        ws.row_dimensions[row_num].height = 18
        row_num += 1

    # ── Column widths ──
    widths = [14, 36, 16, 14, 14]
    for _ in cities:
        for _ in range(TOP_OFFLINE):
            widths += [22, 9]
        widths.append(10)
        for _ in range(TOP_ONLINE):
            widths += [22, 9]
        widths.append(10)
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A5"
    wb.save(output_path)


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html",
                           default_batch_size=BATCH_SIZE,
                           default_delay=DELAY_SEC)


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "לא נשלח קובץ"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "קובץ חייב להיות xlsx"}), 400

    job_id    = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{job_id}.xlsx"
    f.save(str(save_path))

    try:
        products = read_supplier_file(str(save_path))
    except Exception as e:
        return jsonify({"error": f"שגיאה בקריאת הקובץ: {e}"}), 400

    jobs[job_id] = {
        "status":        "ready",
        "progress":      0,
        "total":         len(products) * len(DEFAULT_CITIES),
        "log":           "",
        "result_path":   None,
        "error":         None,
        "product_count": len(products),
        "upload_path":   str(save_path),
    }
    return jsonify({"job_id": job_id, "product_count": len(products)})


@app.route("/api/start/<job_id>", methods=["POST"])
def start_job(job_id):
    if job_id not in jobs:
        return jsonify({"error": "job not found"}), 404

    data        = request.json or {}
    cities      = data.get("cities", DEFAULT_CITIES)
    batch_size  = int  (data.get("batch_size",  BATCH_SIZE))
    delay       = float(data.get("delay_sec",   DELAY_SEC))

    upload_path = jobs[job_id]["upload_path"]
    products    = read_supplier_file(upload_path)

    jobs[job_id]["total"]  = len(products) * len(cities)
    jobs[job_id]["status"] = "running"

    t = threading.Thread(
        target=run_scrape_thread,
        args=(job_id, products, cities, batch_size, delay),
        daemon=True,
    )
    t.start()
    return jsonify({"ok": True})


@app.route("/api/status/<job_id>")
def job_status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "not found"}), 404
    j = jobs[job_id]
    return jsonify({
        "status":   j["status"],
        "progress": j["progress"],
        "total":    j["total"],
        "log":      j["log"],
        "error":    j.get("error"),
        "pct":      int(j["progress"] / j["total"] * 100) if j["total"] > 0 else 0,
    })


@app.route("/api/download/<job_id>")
def download(job_id):
    if job_id not in jobs:
        return jsonify({"error": "not found"}), 404
    path = jobs[job_id].get("result_path")
    if not path or not Path(path).exists():
        return jsonify({"error": "קובץ לא מוכן"}), 400
    return send_file(path, as_attachment=True, download_name="chp_results.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
