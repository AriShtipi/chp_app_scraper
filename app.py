import os
import threading
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from config import VAT, DELAY_SEC, BATCH_SIZE, DEFAULT_CITIES
from excel_parser import read_supplier_file
from excel_writer import write_results
from scraper import run_scrape_thread

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

jobs = {}


def _make_on_finish(job_id):
    """בונה את ה-callback שנקרא בסיום הסריקה כדי לשמור את קובץ התוצאה."""
    def on_finish(products, cities):
        result_path = OUTPUT_DIR / f"{job_id}.xlsx"
        write_results(products, cities, str(result_path))
        jobs[job_id]["result_path"] = str(result_path)
    return on_finish


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

    job_id    = str(uuid.uuid4())
    save_path = UPLOAD_DIR / f"{job_id}.xlsx"
    f.save(str(save_path))

    try:
        products = read_supplier_file(str(save_path))
    except Exception as e:
        save_path.unlink(missing_ok=True)
        return jsonify({"error": f"שגיאה בקריאת הקובץ: {e}"}), 400

    if not products:
        save_path.unlink(missing_ok=True)
        return jsonify({"error": "לא נמצאו מוצרים בקובץ — בדוק שהפורמט תקין"}), 400

    jobs[job_id] = {
        "status":        "ready",
        "progress":      0,
        "total":         0,
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

    job = jobs[job_id]
    if job["status"] == "running":
        return jsonify({"error": "הסריקה כבר רצה"}), 400

    data     = request.json or {}
    cities   = data.get("cities", DEFAULT_CITIES)
    delay    = float(data.get("delay", DELAY_SEC))
    batch_sz = int(data.get("batch_size", BATCH_SIZE))

    products = read_supplier_file(job["upload_path"])
    job["total"]  = len(products) * len(cities)
    job["status"] = "running"

    t = threading.Thread(
        target=run_scrape_thread,
        args=(job, products, cities, delay, batch_sz, _make_on_finish(job_id)),
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


@app.route("/api/cancel/<job_id>", methods=["POST"])
def cancel_job(job_id):
    if job_id not in jobs:
        return jsonify({"error": "not found"}), 404
    j = jobs[job_id]
    if j["status"] == "running":
        j["status"] = "cancelled"
    return jsonify({"ok": True, "status": j["status"]})


@app.route("/api/download/<job_id>")
def download(job_id):
    if job_id not in jobs:
        return jsonify({"error": "not found"}), 404
    path = jobs[job_id].get("result_path")
    if not path or not Path(path).exists():
        return jsonify({"error": "קובץ לא מוכן"}), 400

    response = send_file(path, as_attachment=True, download_name="chp_results.xlsx")

    def cleanup():
        import time
        time.sleep(5)
        try:
            Path(path).unlink(missing_ok=True)
        except Exception:
            pass
        try:
            Path(jobs[job_id]["upload_path"]).unlink(missing_ok=True)
        except Exception:
            pass
        jobs.pop(job_id, None)

    threading.Thread(target=cleanup, daemon=True).start()
    return response


@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "active_jobs": len(jobs)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)