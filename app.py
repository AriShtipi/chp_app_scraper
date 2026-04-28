import asyncio
import json
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

# ── env-controlled settings ──
VAT         = float(os.environ.get("VAT", 1.18))
DELAY_SEC   = float(os.environ.get("DELAY_SEC", 1.0))
BATCH_SIZE  = int(os.environ.get("BATCH_SIZE", 10))
TOP_N_LOCAL  = int(os.environ.get("TOP_N_LOCAL", 2))   # חנויות פיזיות
TOP_N_ONLINE = int(os.environ.get("TOP_N_ONLINE", 1))  # אונליין

DEFAULT_CITIES = [
    {"name": "טירת כרמל", "code1": "9000", "code2": "2100"},
    {"name": "אריאל",      "code1": "9000", "code2": "3600"},
    {"name": "ביתר עילית", "code1": "9000", "code2": "3400"},
]

# ────────────────────────────────────────
# Excel parser
# ────────────────────────────────────────
def read_supplier_file(path: str) -> list:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    products = []
    current_category = ""
    header_found = False

    for row in ws.iter_rows(values_only=True):
        # Detect header row
        if not header_found:
            row_text = " ".join(str(v or "") for v in row)
            if "ברקוד" in row_text:
                header_found = True
            elif any(v and isinstance(v, str) and len(v.strip()) > 3
                     and not str(v).startswith("יח") for v in row):
                # category candidate before header
                cat = next((str(v).strip() for v in row
                            if v and isinstance(v, str) and len(str(v).strip()) > 3), "")
                if cat:
                    current_category = cat
            continue

        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        d = row[3] if len(row) > 3 else None
        f = row[5] if len(row) > 5 else None

        # Barcode row
        if b and isinstance(b, (int, float)) and len(str(int(b))) >= 12:
            barcode = str(int(b))
            name = str(c or "").strip().strip('"\'')
            price_list = round(float(d), 4) if isinstance(d, (int, float)) else 0.0
            price_buy  = round(float(f), 4) if isinstance(f, (int, float)) else price_list
            products.append({
                "category": current_category,
                "barcode": barcode,
                "name": name,
                "price_list_ex_vat": price_list,
                "price_buy_ex_vat":  price_buy,
                "price_buy_inc_vat": round(price_buy * VAT, 2),
                "chp": {}
            })
        else:
            # Category row — look for a non-empty text cell
            for v in row:
                if v and isinstance(v, str):
                    t = v.strip().strip('"\'')
                    if t and len(t) > 2 and not t.startswith("יח"):
                        current_category = t
                        break

    return products


# ────────────────────────────────────────
# CHP scraper
# ────────────────────────────────────────
def build_url(barcode, city):
    from urllib.parse import quote
    return f"https://chp.co.il/{quote(city['name'])}/{city['code1']}/{city['code2']}/{barcode}/0"


async def extract_prices(page):
    """
    Extracts local and online prices from a CHP product page.

    Page structure (confirmed from real HTML dump):
    - Product info table (no <th>) — skip
    - Optional: <h4>לא נמצאו תוצאות ב[עיר]</h4>  — means no local results
    - Optional: local results <table class="results-table"> with columns:
        רשת | שם החנות | כתובת החנות | מבצע | מחיר
    - <h4>תוצאות מחנויות באינטרנט ...
    - Online results <table> with columns:
        רשת | שם החנות | אתר אינטרנט | מבצע | מחיר
    - Between each store: a separator row with single cell "אתר: https://..."  — skip it
    - Sale badge text format: "8.33 *"  (number THEN asterisk)
    - Price cell: plain text "11.90"
    """
    try:
        result = await page.evaluate("""() => {
            const out = { local: [], online: [] };

            // Split page HTML into local section and online section by the h4 heading
            const allH4 = Array.from(document.querySelectorAll('h4'));
            let onlineHeading = null;
            for (const h of allH4) {
                if (h.textContent.includes('תוצאות מחנויות באינטרנט')) {
                    onlineHeading = h;
                    break;
                }
            }

            // Check if local results exist
            let localHasNoResults = false;
            for (const h of allH4) {
                if (h.textContent.includes('לא נמצאו תוצאות')) {
                    localHasNoResults = true;
                    break;
                }
            }

            function parseTable(table, isOnline) {
                const rows = Array.from(table.querySelectorAll('tr'));
                if (!rows.length) return;

                // Build column index from header row
                const headers = Array.from(rows[0].querySelectorAll('th')).map(th => th.textContent.trim());
                if (!headers.includes('מחיר')) return; // not a price table

                const col = {};
                headers.forEach((h, i) => col[h] = i);

                for (let ri = 1; ri < rows.length; ri++) {
                    const cells = Array.from(rows[ri].querySelectorAll('td'));

                    // Skip separator rows (single cell starting with 'אתר:')
                    if (cells.length === 1) continue;
                    if (!cells.length) continue;

                    const getCell = name => (col[name] !== undefined && col[name] < cells.length)
                        ? cells[col[name]] : null;

                    const network = getCell('רשת')?.textContent?.trim() || '';
                    const store   = getCell('שם החנות')?.textContent?.trim() || '';

                    // Price: plain text in cell, format "12.90"
                    let price = null;
                    const priceCell = getCell('מחיר');
                    if (priceCell) {
                        const raw = priceCell.textContent.trim().replace(/,/g, '');
                        const m = raw.match(/(\d+\.\d+)/);
                        if (m) price = parseFloat(m[1]);
                    }

                    // Sale: button text format "8.33 *" (number then asterisk)
                    // Also handles "מבצע" text (no number = member discount, skip number extraction)
                    let sale = null;
                    const saleCell = getCell('מבצע');
                    if (saleCell) {
                        const saleTxt = saleCell.textContent.replace(/,/g, '').trim();
                        // Match pattern: number followed by space and asterisk
                        const m = saleTxt.match(/(\d+\.\d+)\s*\*/);
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

            // Find all results-table tables and classify them
            const allTables = Array.from(document.querySelectorAll('table'));
            allTables.forEach(table => {
                // Determine if this table is before or after the online heading
                let isOnline = false;
                if (onlineHeading) {
                    // compareDocumentPosition: 4 means table comes AFTER onlineHeading
                    const pos = onlineHeading.compareDocumentPosition(table);
                    isOnline = !!(pos & 4);
                }

                // Also detect by column: if it has 'אתר אינטרנט' column → online
                const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent.trim());
                if (headers.includes('אתר אינטרנט')) isOnline = true;

                if (isOnline || (!localHasNoResults && !isOnline)) {
                    parseTable(table, isOnline);
                }
            });

            out.local.sort((a, b) => a.effective - b.effective);
            out.online.sort((a, b) => a.effective - b.effective);
            return out;
        }""")
        return result
    except Exception as e:
        return {"local": [], "online": []}


async def scrape_one(context, semaphore, prod, city, job, delay):
    """Scrape a single product/city in its own tab. Semaphore limits concurrency."""
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


async def scrape_job(job_id, products, cities, delay, batch_size):
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

            # Semaphore = max concurrent tabs at any moment
            semaphore = asyncio.Semaphore(batch_size)

            # All tasks: every product × every city run truly in parallel
            tasks = [
                scrape_one(context, semaphore, prod, city, job, delay)
                for city in cities
                for prod in products
            ]

            await asyncio.gather(*tasks)
            await browser.close()

        result_path = OUTPUT_DIR / f"{job_id}.xlsx"
        write_results(products, cities, str(result_path))
        job["status"] = "done"
        job["result_path"] = str(result_path)

    except Exception as e:
        job["status"] = "error"
        job["error"] = str(e)


def run_scrape_thread(job_id, products, cities, delay, batch_size):
    asyncio.run(scrape_job(job_id, products, cities, delay, batch_size))


# ────────────────────────────────────────
# Excel writer
# ────────────────────────────────────────
def write_results(products, cities, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "תוצאות CHP"
    ws.sheet_view.rightToLeft = True

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    BLUE = "1a73e8"
    WHITE = "FFFFFF"
    CITY_COLORS  = ["DCE8FB", "D4F0E8", "FEF3D0", "F3E8FF", "FFE8E8"]
    LOCAL_HDR    = "4A90D9"
    ONLINE_HDR   = "27AE60"
    CAT_BG       = "E8F0FE"

    def c(r, col, val="", bold=False, bg=None, fg="000000",
          align="right", fmt=None, wrap=False, size=10):
        cl = ws.cell(row=r, column=col, value=val)
        cl.font = Font(name="Arial", bold=bold, color=fg, size=size)
        if bg:
            cl.fill = PatternFill("solid", fgColor=bg)
        cl.alignment = Alignment(horizontal=align, vertical="center",
                                  wrap_text=wrap, readingOrder=2)
        cl.border = border
        if fmt:
            cl.number_format = fmt
        return cl

    FIXED = 5  # category, name, barcode, price ex-vat, price inc-vat

    # columns per city: TOP_N_LOCAL*(store+price) + local_min + TOP_N_ONLINE*(store+price) + online_min
    cols_per_city = TOP_N_LOCAL * 2 + 1 + TOP_N_ONLINE * 2 + 1
    city_starts = {}
    col_idx = FIXED + 1
    for city in cities:
        city_starts[city["name"]] = col_idx
        col_idx += cols_per_city
    TOTAL_COLS = col_idx - 1

    # ── Row 1: title ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS)
    cl = ws.cell(row=1, column=1, value="📊  השוואת מחירי CHP — תוצאות סריקה")
    cl.font = Font(name="Arial", bold=True, size=13, color=WHITE)
    cl.fill = PatternFill("solid", fgColor=BLUE)
    cl.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    ws.row_dimensions[1].height = 26

    # ── Row 2: fixed headers + city merged headers ──
    fixed_headers = [
        "קטגוריה", "שם פריט", "ברקוד",
        'מחיר קנייה\n(ללא מע"מ)',
        f'מחיר קנייה\n(+מע"מ {int((VAT-1)*100)}%)'
    ]
    for ci, h in enumerate(fixed_headers, 1):
        c(2, ci, h, bold=True, bg=BLUE, fg=WHITE, align="center", wrap=True)

    for i, city in enumerate(cities):
        start = city_starts[city["name"]]
        ws.merge_cells(start_row=2, start_column=start,
                        end_row=2, end_column=start + cols_per_city - 1)
        cl = ws.cell(row=2, column=start, value=city["name"])
        cl.font = Font(name="Arial", bold=True, size=11, color="222222")
        cl.fill = PatternFill("solid", fgColor=CITY_COLORS[i % len(CITY_COLORS)])
        cl.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        cl.border = border
    ws.row_dimensions[2].height = 22

    # ── Row 3: sub-headers (local | online) ──
    for ci in range(1, FIXED + 1):
        c(3, ci, "", bg="F0F4FF")

    for i, city in enumerate(cities):
        col_idx = city_starts[city["name"]]
        # Local section header
        local_span = TOP_N_LOCAL * 2 + 1
        ws.merge_cells(start_row=3, start_column=col_idx,
                        end_row=3, end_column=col_idx + local_span - 1)
        cl = ws.cell(row=3, column=col_idx, value="🏪 חנויות פיזיות")
        cl.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        cl.fill = PatternFill("solid", fgColor=LOCAL_HDR)
        cl.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        cl.border = border
        col_idx += local_span

        # Online section header
        online_span = TOP_N_ONLINE * 2 + 1
        ws.merge_cells(start_row=3, start_column=col_idx,
                        end_row=3, end_column=col_idx + online_span - 1)
        cl = ws.cell(row=3, column=col_idx, value="🌐 אונליין")
        cl.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        cl.fill = PatternFill("solid", fgColor=ONLINE_HDR)
        cl.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        cl.border = border
        col_idx += online_span
    ws.row_dimensions[3].height = 20

    # ── Row 4: field sub-headers ──
    for ci in range(1, FIXED + 1):
        c(4, ci, "", bg="F0F4FF")

    for i, city in enumerate(cities):
        col_idx = city_starts[city["name"]]
        for n in range(TOP_N_LOCAL):
            c(4, col_idx,   f"#{n+1} שם עסק",  bold=True, bg="EBF3FB", align="center")
            c(4, col_idx+1, f"#{n+1} מחיר",    bold=True, bg="EBF3FB", align="center")
            col_idx += 2
        c(4, col_idx, "הכי זול", bold=True, bg="EBF3FB", align="center")
        col_idx += 1
        for n in range(TOP_N_ONLINE):
            c(4, col_idx,   f"#{n+1} אתר",     bold=True, bg="E8F8F1", align="center")
            c(4, col_idx+1, f"#{n+1} מחיר",    bold=True, bg="E8F8F1", align="center")
            col_idx += 2
        c(4, col_idx, "הכי זול", bold=True, bg="E8F8F1", align="center")
        col_idx += 1
    ws.row_dimensions[4].height = 20

    # ── Data rows ──
    row_num = 5
    last_cat = ""

    def color_min(cl_cell, effective, inc_vat):
        if not effective or not inc_vat:
            return
        pct = (inc_vat - effective) / inc_vat * 100
        if pct >= 25:
            cl_cell.fill = PatternFill("solid", fgColor="D4EDDA")
            cl_cell.font = Font(name="Arial", bold=True, color="155724", size=10)
        elif pct >= 10:
            cl_cell.fill = PatternFill("solid", fgColor="FFF3CD")
            cl_cell.font = Font(name="Arial", bold=True, color="856404", size=10)
        else:
            cl_cell.fill = PatternFill("solid", fgColor="F8D7DA")
            cl_cell.font = Font(name="Arial", bold=True, color="721c24", size=10)

    for prod in products:
        if prod["category"] != last_cat and prod["category"]:
            ws.merge_cells(start_row=row_num, start_column=1,
                            end_row=row_num, end_column=TOTAL_COLS)
            cl = ws.cell(row=row_num, column=1, value=f"  {prod['category']}")
            cl.font = Font(name="Arial", bold=True, size=10, color="1a3a6b")
            cl.fill = PatternFill("solid", fgColor=CAT_BG)
            cl.alignment = Alignment(horizontal="right", vertical="center", readingOrder=2)
            cl.border = border
            ws.row_dimensions[row_num].height = 18
            row_num += 1
            last_cat = prod["category"]

        c(row_num, 1, prod["category"])
        c(row_num, 2, prod["name"])
        c(row_num, 3, prod["barcode"])
        c(row_num, 4, prod["price_buy_ex_vat"],  fmt='0.00')
        c(row_num, 5, prod["price_buy_inc_vat"],  fmt='0.00', bold=True)

        inc_vat = prod["price_buy_inc_vat"]
        col_idx = FIXED + 1

        for city in cities:
            city_data = prod["chp"].get(city["name"], {"local": [], "online": []})
            local_prices  = city_data.get("local",  [])[:TOP_N_LOCAL]
            online_prices = city_data.get("online", [])[:TOP_N_ONLINE]

            # Local columns
            local_min = None
            for n in range(TOP_N_LOCAL):
                if n < len(local_prices):
                    e = local_prices[n]
                    label = e.get("store") or e.get("network") or "—"
                    pval  = e["effective"]
                    if local_min is None or pval < local_min:
                        local_min = pval
                    c(row_num, col_idx,   label, bg="FAFAFA")
                    c(row_num, col_idx+1, pval,  bg="FAFAFA", fmt='0.00')
                else:
                    c(row_num, col_idx,   "—", bg="FAFAFA")
                    c(row_num, col_idx+1, "—", bg="FAFAFA")
                col_idx += 2

            min_cell = c(row_num, col_idx,
                         local_min if local_min else "לא נמצא",
                         fmt='0.00' if local_min else None, bold=bool(local_min))
            if local_min:
                color_min(min_cell, local_min, inc_vat)
            col_idx += 1

            # Online columns
            online_min = None
            for n in range(TOP_N_ONLINE):
                if n < len(online_prices):
                    e = online_prices[n]
                    label = e.get("store") or e.get("network") or "—"
                    pval  = e["effective"]
                    if online_min is None or pval < online_min:
                        online_min = pval
                    c(row_num, col_idx,   label, bg="F0FFF8")
                    c(row_num, col_idx+1, pval,  bg="F0FFF8", fmt='0.00')
                else:
                    c(row_num, col_idx,   "—", bg="F0FFF8")
                    c(row_num, col_idx+1, "—", bg="F0FFF8")
                col_idx += 2

            min_cell = c(row_num, col_idx,
                         online_min if online_min else "לא נמצא",
                         fmt='0.00' if online_min else None, bold=bool(online_min))
            if online_min:
                color_min(min_cell, online_min, inc_vat)
            col_idx += 1

        ws.row_dimensions[row_num].height = 18
        row_num += 1

    # Column widths
    widths = [14, 36, 16, 12, 12]
    for _ in cities:
        for _ in range(TOP_N_LOCAL):
            widths += [22, 9]
        widths.append(10)
        for _ in range(TOP_N_ONLINE):
            widths += [22, 9]
        widths.append(10)
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A5"
    wb.save(output_path)


# ────────────────────────────────────────
# Routes
# ────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html",
                           vat=int((VAT-1)*100),
                           delay=DELAY_SEC,
                           batch_size=BATCH_SIZE)


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "לא נשלח קובץ"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "קובץ חייב להיות xlsx"}), 400

    job_id = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{job_id}.xlsx"
    f.save(str(save_path))

    try:
        products = read_supplier_file(str(save_path))
    except Exception as e:
        save_path.unlink(missing_ok=True)
        return jsonify({"error": f"שגיאה בקריאת הקובץ: {e}"}), 400

    jobs[job_id] = {
        "status": "ready", "progress": 0, "total": 0,
        "log": "", "result_path": None, "error": None,
        "product_count": len(products),
        "upload_path": str(save_path),
    }
    return jsonify({"job_id": job_id, "product_count": len(products)})


@app.route("/api/start/<job_id>", methods=["POST"])
def start_job(job_id):
    if job_id not in jobs:
        return jsonify({"error": "job not found"}), 404

    data = request.json or {}
    cities    = data.get("cities", DEFAULT_CITIES)
    delay     = float(data.get("delay", DELAY_SEC))
    batch_sz  = int(data.get("batch_size", BATCH_SIZE))

    upload_path = jobs[job_id]["upload_path"]
    products = read_supplier_file(upload_path)
    jobs[job_id]["total"] = len(products) * len(cities)
    jobs[job_id]["status"] = "running"

    t = threading.Thread(
        target=run_scrape_thread,
        args=(job_id, products, cities, delay, batch_sz),
        daemon=True
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
        "pct":      int(j["progress"] / j["total"] * 100) if j["total"] > 0 else 0
    })


@app.route("/api/download/<job_id>")
def download(job_id):
    if job_id not in jobs:
        return jsonify({"error": "not found"}), 404
    path = jobs[job_id].get("result_path")
    if not path or not Path(path).exists():
        return jsonify({"error": "קובץ לא מוכן"}), 400

    response = send_file(path, as_attachment=True, download_name="chp_results.xlsx")

    # Cleanup after sending
    def cleanup():
        import time; time.sleep(5)
        try: Path(path).unlink(missing_ok=True)
        except: pass
        try: Path(jobs[job_id]["upload_path"]).unlink(missing_ok=True)
        except: pass
        jobs.pop(job_id, None)

    threading.Thread(target=cleanup, daemon=True).start()
    return response


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
