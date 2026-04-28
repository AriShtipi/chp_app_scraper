import asyncio
import json
import os
import re
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Job store: job_id -> {status, progress, total, log, result_path, error}
jobs = {}

VAT = 1.18
DELAY_SEC = 1.0
TOP_N = 2

DEFAULT_CITIES = [
    {"name": "טירת כרמל", "code1": "9000", "code2": "2100"},
    {"name": "אריאל",      "code1": "9000", "code2": "3600"},
    {"name": "ביתר עילית", "code1": "9000", "code2": "3400"},
]


# ──────────────────────────────────────────
# Excel parser
# ──────────────────────────────────────────
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
            price_buy = float(f) if f and isinstance(f, (int, float)) else price_list
            products.append({
                "category": current_category,
                "barcode": barcode,
                "name": name,
                "price_list_ex_vat": round(price_list, 4),
                "price_buy_ex_vat": round(price_buy, 4),
                "price_buy_inc_vat": round(price_buy * VAT, 2),
                "chp": {}
            })
        elif b is None or b == 0:
            a = row[0] if len(row) > 0 else None
            cat_text = str(a or c or "").strip()
            if cat_text and len(cat_text) < 60 and not cat_text.startswith("יח'") and cat_text not in ("0", ""):
                current_category = cat_text

    return products


# ──────────────────────────────────────────
# CHP scraper (Playwright)
# ──────────────────────────────────────────
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

                network = (await sc(col.get("רשת", -1)).inner_text()).strip() if col.get("רשת") is not None and sc(col["רשת"]) else ""
                store   = (await sc(col.get("שם החנות", -1)).inner_text()).strip() if col.get("שם החנות") is not None and sc(col["שם החנות"]) else ""
                price, sale = None, None

                if col.get("מחיר") is not None:
                    el = sc(col["מחיר"])
                    if el:
                        m = re.search(r"[\d.]+", await el.inner_text())
                        if m: price = float(m.group())

                if col.get("מבצע") is not None:
                    el = sc(col["מבצע"])
                    if el:
                        m = re.search(r"\*\s*([\d.]+)", await el.inner_text())
                        if m: sale = float(m.group(1))

                effective = sale if sale else price
                if effective and effective > 0:
                    results.append({"network": network, "store": store,
                                    "price": price, "sale": sale, "effective": effective})
    except Exception:
        pass
    return results


async def scrape_job(job_id: str, products: list, cities: list):
    from playwright.async_api import async_playwright

    job = jobs[job_id]
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
            page = await context.new_page()

            for pi, prod in enumerate(products):
                prod["chp"] = {}
                for city in cities:
                    from urllib.parse import quote
                    url = f"https://chp.co.il/{quote(city['name'])}/{city['code1']}/{city['code2']}/{prod['barcode']}/0"

                    for attempt in range(3):
                        try:
                            await page.goto(url, wait_until="networkidle", timeout=15000)
                            await asyncio.sleep(0.4)
                            prices = await extract_prices(page)
                            break
                        except Exception:
                            if attempt < 2:
                                await asyncio.sleep(2)
                            else:
                                prices = []

                    prices.sort(key=lambda x: x["effective"])
                    prod["chp"][city["name"]] = prices[:TOP_N]
                    job["progress"] += 1
                    job["log"] = f"{prod['name'][:35]} / {city['name']}"
                    await asyncio.sleep(DELAY_SEC)

            await browser.close()

        # Write result Excel
        result_path = OUTPUT_DIR / f"{job_id}.xlsx"
        write_results(products, cities, str(result_path))
        job["status"] = "done"
        job["result_path"] = str(result_path)

    except Exception as e:
        job["status"] = "error"
        job["error"] = str(e)


def run_scrape_thread(job_id, products, cities):
    asyncio.run(scrape_job(job_id, products, cities))


# ──────────────────────────────────────────
# Excel writer
# ──────────────────────────────────────────
def write_results(products: list, cities: list, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "תוצאות CHP"
    ws.sheet_view.rightToLeft = True

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    BLUE = "1a73e8"
    WHITE = "FFFFFF"
    CITY_COLORS = ["DCE8FB", "D4F0E8", "FEF3D0", "F3E8FF", "FFE8E8"]
    CAT_BG = "E8F0FE"

    def cell(r, c, val="", bold=False, bg=None, fg="000000", align="right", fmt=None, wrap=False):
        cl = ws.cell(row=r, column=c, value=val)
        cl.font = Font(name="Arial", bold=bold, color=fg, size=10)
        if bg:
            cl.fill = PatternFill("solid", fgColor=bg)
        cl.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap, readingOrder=2)
        cl.border = border
        if fmt:
            cl.number_format = fmt
        return cl

    FIXED = 5
    city_starts = {}
    col = FIXED + 1
    for city in cities:
        city_starts[city["name"]] = col
        col += TOP_N * 2 + 1
    TOTAL_COLS = col - 1

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS)
    c = ws.cell(row=1, column=1, value="📊  השוואת מחירי CHP — תוצאות סריקה")
    c.font = Font(name="Arial", bold=True, size=13, color=WHITE)
    c.fill = PatternFill("solid", fgColor=BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    ws.row_dimensions[1].height = 26

    # Row 2: fixed headers + city merged headers
    for ci, h in enumerate(["קטגוריה", "שם פריט", "ברקוד", "מחיר קנייה\n(ללא מע\"מ)", f"מחיר קנייה\n(+מע\"מ {int((VAT-1)*100)}%)"], 1):
        cell(2, ci, h, bold=True, bg=BLUE, fg=WHITE, align="center", wrap=True)

    for i, city in enumerate(cities):
        start = city_starts[city["name"]]
        span = TOP_N * 2 + 1
        ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=start + span - 1)
        c = ws.cell(row=2, column=start, value=city["name"])
        c.font = Font(name="Arial", bold=True, size=11, color="222222")
        c.fill = PatternFill("solid", fgColor=CITY_COLORS[i % len(CITY_COLORS)])
        c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        c.border = border
    ws.row_dimensions[2].height = 22

    # Row 3: sub-headers
    for ci in range(1, FIXED + 1):
        cell(3, ci, "", bg="F0F4FF")
    col = FIXED + 1
    for i, city in enumerate(cities):
        cc = CITY_COLORS[i % len(CITY_COLORS)]
        for n in range(TOP_N):
            cell(3, col,   f"#{n+1} שם עסק", bold=True, bg=cc, align="center")
            cell(3, col+1, f"#{n+1} מחיר",   bold=True, bg=cc, align="center")
            col += 2
        cell(3, col, "הכי זול", bold=True, bg=cc, align="center")
        col += 1
    ws.row_dimensions[3].height = 20

    # Data rows
    row_num = 4
    last_cat = ""
    for prod in products:
        if prod["category"] != last_cat and prod["category"]:
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=TOTAL_COLS)
            c = ws.cell(row=row_num, column=1, value=f"  {prod['category']}")
            c.font = Font(name="Arial", bold=True, size=10, color="1a3a6b")
            c.fill = PatternFill("solid", fgColor=CAT_BG)
            c.alignment = Alignment(horizontal="right", vertical="center", readingOrder=2)
            c.border = border
            ws.row_dimensions[row_num].height = 18
            row_num += 1
            last_cat = prod["category"]

        cell(row_num, 1, prod["category"])
        cell(row_num, 2, prod["name"])
        cell(row_num, 3, prod["barcode"])
        cell(row_num, 4, prod["price_buy_ex_vat"], fmt="#,##0.00")
        cell(row_num, 5, prod["price_buy_inc_vat"], fmt="#,##0.00", bold=True)

        col = FIXED + 1
        for i, city in enumerate(cities):
            city_prices = prod["chp"].get(city["name"], [])
            min_price = None
            for n in range(TOP_N):
                if n < len(city_prices):
                    e = city_prices[n]
                    label = e.get("store") or e.get("network") or "—"
                    pval = e["effective"]
                    if min_price is None or pval < min_price:
                        min_price = pval
                    cell(row_num, col,   label, bg="FAFAFA")
                    cell(row_num, col+1, pval,  bg="FAFAFA", fmt="#,##0.00")
                else:
                    cell(row_num, col,   "—", bg="FAFAFA")
                    cell(row_num, col+1, "—", bg="FAFAFA")
                col += 2

            if min_price:
                inc = prod["price_buy_inc_vat"]
                pct = ((inc - min_price) / inc * 100) if inc else None
                cl = cell(row_num, col, min_price, fmt="#,##0.00", bold=True)
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
                cell(row_num, col, "לא נמצא", fg="999999")
            col += 1

        ws.row_dimensions[row_num].height = 18
        row_num += 1

    # Column widths
    widths = [14, 36, 16, 14, 14]
    for _ in cities:
        for _ in range(TOP_N):
            widths += [22, 9]
        widths.append(10)
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A4"
    wb.save(output_path)


# ──────────────────────────────────────────
# Routes
# ──────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "לא נשלח קובץ"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "קובץ חייב להיות xlsx"}), 400

    job_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{job_id}.xlsx"
    f.save(str(save_path))

    try:
        products = read_supplier_file(str(save_path))
    except Exception as e:
        return jsonify({"error": f"שגיאה בקריאת הקובץ: {e}"}), 400

    jobs[job_id] = {
        "status": "ready",
        "progress": 0,
        "total": len(products) * len(DEFAULT_CITIES),
        "log": "",
        "result_path": None,
        "error": None,
        "product_count": len(products),
        "upload_path": str(save_path),
    }
    return jsonify({"job_id": job_id, "product_count": len(products)})


@app.route("/api/start/<job_id>", methods=["POST"])
def start_job(job_id):
    if job_id not in jobs:
        return jsonify({"error": "job not found"}), 404

    data = request.json or {}
    cities = data.get("cities", DEFAULT_CITIES)
    upload_path = jobs[job_id]["upload_path"]

    products = read_supplier_file(upload_path)
    jobs[job_id]["total"] = len(products) * len(cities)
    jobs[job_id]["status"] = "running"

    t = threading.Thread(target=run_scrape_thread, args=(job_id, products, cities), daemon=True)
    t.start()
    return jsonify({"ok": True})


@app.route("/api/status/<job_id>")
def job_status(job_id):
    if job_id not in jobs:
        return jsonify({"error": "not found"}), 404
    j = jobs[job_id]
    return jsonify({
        "status": j["status"],
        "progress": j["progress"],
        "total": j["total"],
        "log": j["log"],
        "error": j.get("error"),
        "pct": int(j["progress"] / j["total"] * 100) if j["total"] > 0 else 0
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
