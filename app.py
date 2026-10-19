"""
JKDM HS Code Lookup Tool - Web App (Flask)
Run:  python app.py      Open: http://localhost:5000
"""
import os, re, uuid, threading
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
import auth
import hs_core as core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
app.secret_key = auth.secret_key()
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  SESSION_COOKIE_SECURE=os.environ.get("HS_HTTPS") == "1",
                  PERMANENT_SESSION_LIFETIME=8 * 3600)


@app.before_request
def require_login():
    if request.endpoint in ("login", "static") or session.get("user"):
        return
    if request.path.startswith("/api/"):
        return jsonify({"error": "Not logged in."}), 401
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        ip = request.remote_addr
        if auth.locked(ip):
            error = "Too many attempts. Wait 5 minutes."
        elif auth.verify(request.form.get("username"), request.form.get("password")):
            auth.record(ip, True)
            session.clear(); session.permanent = True
            session["user"] = request.form["username"].strip().lower()
            return redirect(url_for("index"))
        else:
            auth.record(ip, False); error = "Wrong user ID or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login"))
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

JOBS = {}
DBJOB = {"running": False, "done": 0, "total": 0, "log": [], "result": None}
MODES = ("auto", "offline", "live")


@app.route("/")
def index():
    return render_template("index.html", tariff_types=core.TARIFF_TYPES)


@app.route("/api/lookup", methods=["POST"])
def api_lookup():
    d = request.get_json() or {}
    digits, note = core.normalize_code(d.get("hs_code", ""))
    tariff = d.get("tariff_type", "PDK")
    mode = d.get("mode", "auto") if d.get("mode") in MODES else "auto"
    if len(digits) < 6:
        return jsonify({"error": "Please enter at least 6 digits."}), 400
    res = core.lookup(digits, tariff, mode)
    if res["status"] == "NETWORK ERROR":
        return jsonify({"error": "Network error (and no offline data): " + res["error"]}), 502
    if res["status"] == "NOT IN DB":
        return jsonify({"error": res["error"]}), 404
    if not res["rows"]:
        return jsonify({"error": "No HS codes found for '%s'." % digits}), 404
    return jsonify({"query": digits, "count": len(res["rows"]), "results": res["rows"],
                    "final": res["rows"][-1], "source": res["source"], "note": note})


# ---------------- bulk ----------------
@app.route("/api/bulk/start", methods=["POST"])
def bulk_start():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded."}), 400
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Upload an Excel (.xlsx or .xls) file."}), 400
    tariff = request.form.get("tariff_type", "PDK")
    mode = request.form.get("mode", "auto")
    mode = mode if mode in MODES else "auto"
    job_id = str(uuid.uuid4())
    path = os.path.join(UPLOAD_FOLDER, job_id + "_input.xlsx")
    f.save(path)
    JOBS[job_id] = {"status": "running", "progress": 0, "total": 0, "rows": [], "error": None, "tariff": tariff}
    threading.Thread(target=run_bulk, args=(job_id, path, tariff, mode, request.form.get("hs_col", "")), daemon=True).start()
    return jsonify({"job_id": job_id})


def run_bulk(job_id, path, tariff, mode, hs_col):
    job = JOBS[job_id]
    def on_row(i, n, rows):
        job.update(progress=i, total=n, rows=list(rows))
    try:
        out_path = os.path.join(UPLOAD_FOLDER, job_id + "_result.xlsx")
        core.process_excel(path, out_path, tariff, mode, hs_col, on_row)
        job.update(status="done", out_path=out_path)
    except Exception as e:
        job.update(status="error", error=str(e))


@app.route("/api/bulk/status/<job_id>")
def bulk_status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    rows = job["rows"]
    cnt = lambda s: sum(1 for r in rows if r.get("STATUS") == s)
    first = lambda r: str(list(r.values())[0])
    recent = [{"input": first(r), "final": r.get("FINAL_HS_CODE", ""), "status": r.get("STATUS", ""),
               "desc": r.get("DESCRIPTION", "")} for r in rows[-6:]]
    return jsonify({"status": job["status"], "progress": job["progress"], "total": job["total"],
                    "ok": cnt("OK"), "not_found": cnt("NOT FOUND"), "error_count": cnt("ERROR"),
                    "skipped": cnt("SKIPPED"), "recent": recent, "err_msg": job.get("error")})


@app.route("/api/bulk/download/<job_id>")
def bulk_download(job_id):
    job = JOBS.get(job_id)
    if not job or job.get("status") != "done":
        return jsonify({"error": "Result not ready"}), 404
    name = "HS_Results_%s_%s.xlsx" % (job["tariff"], datetime.now().strftime("%Y%m%d_%H%M%S"))
    return send_file(job["out_path"], as_attachment=True, download_name=name)


# ---------------- offline database ----------------
@app.route("/api/db/status")
def db_status():
    return jsonify({"tables": core.db_stats(), "job": {k: DBJOB[k] for k in ("running", "done", "total", "result")},
                    "log": DBJOB["log"][-6:]})


@app.route("/api/db/build", methods=["POST"])
def db_build():
    if DBJOB["running"]:
        return jsonify({"error": "A build is already running."}), 409
    tariff = (request.get_json() or {}).get("tariff_type", "PDK")
    DBJOB.update(running=True, done=0, total=0, log=[], result=None)

    def job():
        try:
            r = core.build_database(tariff, 4, log=lambda m: DBJOB["log"].append(m),
                                    progress=lambda d, t: DBJOB.update(done=d, total=t))
            DBJOB["result"] = "complete" if r["ok"] else "incomplete - click again to retry failed parts"
        except Exception as e:
            DBJOB["result"] = "error: %s" % e
        DBJOB["running"] = False
    threading.Thread(target=job, daemon=True).start()
    return jsonify({"started": True})


if __name__ == "__main__":
    if not auth.load_users():
        raise SystemExit("No users yet. Create one first:  python manage_users.py add yourname")
    print("\n  JKDM HS Code Lookup  ->  http://localhost:5000\n")
    app.run(debug=False, host=os.environ.get("HS_HOST", "127.0.0.1"), port=5000)
