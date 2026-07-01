# -*- coding: utf-8 -*-
"""
MN STEEL DOOR — AKVA Export Docs (Pro)
نظام توثيق تصدير احترافي: يولّد عدّة وثائق تصدير PDF ويرتّبها تلقائياً بمسارات منظّمة.
يعمل محلياً على localhost — توليد PDF مباشر (بدون Print Center).
"""
import os, json, sqlite3, math, shutil, datetime, re, secrets, functools
from flask import (Flask, request, session, redirect, url_for, render_template,
                   g, jsonify, abort, Response, flash)
from werkzeug.security import generate_password_hash, check_password_hash
try:
    import requests
except Exception:
    requests = None

# ---------------------------------------------------------------- paths
APP_DIR = os.path.expanduser("~/Desktop/door_ops")            # بيانات البرنامج
OUTPUT_BASE = os.path.expanduser("~/Desktop/TLT_DOOR_ORGANIZED")  # مخرجات الوثائق المنظّمة
DB_PATH = os.path.join(APP_DIR, "akva.db")
BACKUP_DIR = os.path.join(APP_DIR, "backups")
for d in (APP_DIR, OUTPUT_BASE, BACKUP_DIR):
    os.makedirs(d, exist_ok=True)

# ---------------------------------------------------------------- constants
HS_CODE = "7308300000"
ORIGIN = "TURKEY"
DEFAULT_PASSWORD = "MN2026"
CURRENCIES = ["USD", "EUR", "EGP", "IQD", "JOD", "SAR", "AED", "LYD", "TRY"]
INCOTERMS = ["CFR", "EXW", "FOB", "CIF", "DAP", "DDP"]

DOC_TYPES = {
    "commercial": {"label": "Commercial Invoice", "ar": "فاتورة تجارية", "folder": "Commercial_Invoice"},
    "proforma":   {"label": "Proforma Invoice",   "ar": "بروفورما",      "folder": "Proforma_Invoice"},
    "packing":    {"label": "Packing List",       "ar": "قائمة تعبئة",   "folder": "Packing_List"},
    "shipping":   {"label": "Shipping Instructions", "ar": "تعليمات شحن", "folder": "Shipping_Instructions"},
    "origin":     {"label": "Certificate of Origin", "ar": "شهادة منشأ",  "folder": "Certificate_of_Origin"},
}

MODEL_SEED = [("DÜZ LAMİNOKS", "100", "215", "15"), ("UV GENİŞ AÇILI", "100", "215", "15")]
for _c in [46,101,106,113,115,118,119,120,144,150,194,207,210,244,279,314,328,356,360,365,670]:
    MODEL_SEED.append((f"Code {_c}", "100", "215", "15"))

app = Flask(__name__)

@app.template_filter("fromjson")
def _fromjson(s):
    try:
        return json.loads(s or "[]")
    except Exception:
        return []

# ---------------------------------------------------------------- db
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH); g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def _close(e):
    db = g.pop("db", None)
    if db: db.close()

def setting(k, d=None):
    r = get_db().execute("SELECT value FROM settings WHERE key=?", (k,)).fetchone()
    if not r: return d
    try: return json.loads(r["value"])
    except Exception: return r["value"]

def set_setting(k, v):
    v = json.dumps(v) if not isinstance(v, str) else v
    get_db().execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, v))
    get_db().commit()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.executescript("""
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS companies(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, short_name TEXT DEFAULT '',
        address TEXT, tax_no TEXT DEFAULT '', phone TEXT, email TEXT, bank_name TEXT, branch TEXT, branch_code TEXT, iban TEXT, swift TEXT);
    CREATE TABLE IF NOT EXISTS customers(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, contact TEXT,
        address TEXT, phone TEXT, tax_no TEXT, country TEXT, currency TEXT DEFAULT 'USD', acid TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS models(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT,
        width TEXT, height TEXT, beam TEXT, default_price REAL DEFAULT 0, last_price REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS shipments(id INTEGER PRIMARY KEY AUTOINCREMENT, ref_no TEXT UNIQUE,
        created_at TEXT, company_id INTEGER, customer_id INTEGER, incoterm TEXT, port_loading TEXT,
        port_discharge TEXT, acid_no TEXT, payment_terms TEXT, currency TEXT, fx_rate REAL,
        subtotal REAL, total REAL, qty_total INTEGER, gross_weight REAL, container_count INTEGER,
        notes TEXT, items_json TEXT, containers_json TEXT, docs_json TEXT, files_json TEXT);
    """)
    db.commit(); cur = db.cursor()
    # migrations لقواعد البيانات الموجودة (إضافة أعمدة جديدة)
    for tbl, col, ddl in [
        ("companies", "short_name", "ALTER TABLE companies ADD COLUMN short_name TEXT DEFAULT ''"),
        ("companies", "tax_no", "ALTER TABLE companies ADD COLUMN tax_no TEXT DEFAULT ''"),
        ("customers", "acid", "ALTER TABLE customers ADD COLUMN acid TEXT DEFAULT ''")]:
        if col not in [r[1] for r in cur.execute(f"PRAGMA table_info({tbl})").fetchall()]:
            cur.execute(ddl)
    db.commit()
    def _s(k, v):
        if not cur.execute("SELECT 1 FROM settings WHERE key=?", (k,)).fetchone():
            cur.execute("INSERT INTO settings(key,value) VALUES(?,?)", (k, json.dumps(v)))
    if not cur.execute("SELECT 1 FROM settings WHERE key='password_hash'").fetchone():
        cur.execute("INSERT INTO settings(key,value) VALUES('password_hash',?)", (generate_password_hash(DEFAULT_PASSWORD),))
        cur.execute("INSERT INTO settings(key,value) VALUES('must_change','true')")
    if not cur.execute("SELECT 1 FROM settings WHERE key='secret_key'").fetchone():
        cur.execute("INSERT INTO settings(key,value) VALUES('secret_key',?)", (json.dumps(secrets.token_hex(24)),))
    _s("ref_seq", 0)
    _s("default_payment", "CASH AGAINST DOCUMENTS 100%")

    seed = {}
    sp = os.path.join(os.path.dirname(__file__), "seed_data.json")
    if os.path.exists(sp):
        try: seed = json.load(open(sp, encoding="utf-8"))
        except Exception: seed = {}
    def _ins_company(c):
        cur.execute("""INSERT INTO companies(name,short_name,address,tax_no,phone,email,bank_name,branch,branch_code,iban,swift)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (c.get("name",""), c.get("short",""), c.get("address",""),
            c.get("tax_no",""), c.get("phone",""), c.get("email",""), c.get("bank_name",""), c.get("branch",""),
            c.get("branch_code",""), c.get("iban",""), c.get("swift","")))
    def _ins_customer(c):
        cur.execute("""INSERT INTO customers(name,contact,address,phone,tax_no,country,currency,email,acid)
            VALUES(?,?,?,?,?,?,?,?,?)""", (c.get("name",""), c.get("contact",""), c.get("address",""),
            c.get("phone",""), c.get("tax_no",""), c.get("country",""), c.get("currency","USD"),
            c.get("email",""), c.get("acid","")))

    if not cur.execute("SELECT 1 FROM companies").fetchone():
        for c in seed.get("companies", [{"name": "SAIROGULLARI DIS TICARET LTD STI"}]):
            _ins_company(c)
    else:
        # تعبئة بيانات الشركات الكاملة (لقواعد البيانات التي تحتوي أسماء مبدئية فقط)
        rows = cur.execute("SELECT * FROM companies ORDER BY id").fetchall()
        cols = [d[0] for d in cur.description]
        for i, sc in enumerate(seed.get("companies", [])):
            if i < len(rows):
                row = dict(zip(cols, rows[i])); cid = row["id"]
                if not (row.get("short_name") or "").strip():
                    cur.execute("UPDATE companies SET short_name=? WHERE id=?", (sc.get("short",""), cid))
                if not (row.get("address") or "").strip():   # شركة مبدئية فارغة → عبّئها بالكامل
                    cur.execute("""UPDATE companies SET name=?,address=?,tax_no=?,phone=?,email=?,bank_name=?,
                        branch=?,branch_code=?,iban=?,swift=? WHERE id=?""", (sc.get("name",""), sc.get("address",""),
                        sc.get("tax_no",""), sc.get("phone",""), sc.get("email",""), sc.get("bank_name",""),
                        sc.get("branch",""), sc.get("branch_code",""), sc.get("iban",""), sc.get("swift",""), cid))
            else:
                _ins_company(sc)

    if not cur.execute("SELECT 1 FROM customers").fetchone():
        for c in seed.get("customers", []):
            _ins_customer(c)
    else:
        # إعادة تسمية الزبائن المحدّثين (تفادي التكرار على قواعد البيانات القديمة)
        rename_map = {"MODERN STYLE FOR IMPORT & EXPORT": "سعد زكي — Modern Style (مصر)",
                      "ياسر مرسي": "ياسر مرسي — مصر", "إيمان السيد": "إيمان السيد — مصر"}
        names_now = {r[0] for r in cur.execute("SELECT name FROM customers").fetchall()}
        for old, new in rename_map.items():
            if old in names_now and new not in names_now:
                cur.execute("UPDATE customers SET name=? WHERE name=?", (new, old))
        # إضافة أي زبون جديد + تعبئة الحقول الفارغة للزبائن الموجودين
        existing = {r[0]: r[1] for r in cur.execute("SELECT name,id FROM customers").fetchall()}
        for c in seed.get("customers", []):
            if c["name"] not in existing:
                _ins_customer(c)
            elif c.get("acid") or c.get("address"):
                cur.execute("""UPDATE customers SET contact=COALESCE(NULLIF(contact,''),?),
                    address=COALESCE(NULLIF(address,''),?), phone=COALESCE(NULLIF(phone,''),?),
                    tax_no=COALESCE(NULLIF(tax_no,''),?), acid=COALESCE(NULLIF(acid,''),?) WHERE name=?""",
                    (c.get("contact",""), c.get("address",""), c.get("phone",""), c.get("tax_no",""),
                     c.get("acid",""), c["name"]))

    if not cur.execute("SELECT 1 FROM models").fetchone():
        for name, w, h, b in MODEL_SEED:
            cur.execute("INSERT INTO models(name,width,height,beam,default_price,last_price) VALUES(?,?,?,?,0,0)", (name, w, h, b))
    db.commit(); db.close()

with app.app_context():
    init_db()
    with sqlite3.connect(DB_PATH) as _c:
        _c.row_factory = sqlite3.Row
        app.secret_key = json.loads(_c.execute("SELECT value FROM settings WHERE key='secret_key'").fetchone()["value"])

# ---------------------------------------------------------------- auth
def login_required(f):
    @functools.wraps(f)
    def w(*a, **k):
        if not session.get("auth"): return redirect(url_for("login"))
        if setting("must_change") in (True, "true") and request.endpoint not in ("change_password", "logout", "static"):
            return redirect(url_for("change_password"))
        return f(*a, **k)
    return w

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if check_password_hash(setting("password_hash") or "", request.form.get("password", "")):
            session["auth"] = True; return redirect(url_for("index"))
        flash("كلمة السر غير صحيحة")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("login"))

@app.route("/change-password", methods=["GET", "POST"])
def change_password():
    if not session.get("auth"): return redirect(url_for("login"))
    if request.method == "POST":
        new, conf = request.form.get("new", ""), request.form.get("confirm", "")
        if len(new) < 6: flash("كلمة السر 6 أحرف على الأقل")
        elif new != conf: flash("غير متطابقتين")
        else:
            set_setting("password_hash", generate_password_hash(new)); set_setting("must_change", "false")
            flash("تم تغيير كلمة السر"); return redirect(url_for("index"))
    return render_template("change_password.html")

# ---------------------------------------------------------------- helpers
def next_ref():
    seq = int(setting("ref_seq", 0)) + 1; set_setting("ref_seq", seq)
    return f"AKVA{datetime.date.today().year}-{seq:04d}"

def fx_rate(to):
    if not to or to == "USD" or requests is None: return 1.0
    try:
        return float(requests.get(f"https://api.frankfurter.app/latest?from=USD&to={to}", timeout=6).json()["rates"][to])
    except Exception:
        return None

def safe(s): return re.sub(r"[^\w؀-ۿ .&-]", "_", s or "").strip()[:60] or "x"

def initials(co):
    src = (co["short_name"] or co["name"] or "").strip()
    words = [w for w in re.split(r"\s+", src) if w]
    if not words: return "—"
    if len(words) == 1: return words[0][:2].upper()
    return (words[0][0] + words[1][0]).upper()

# ---------------------------------------------------------------- screens
@app.route("/")
@login_required
def index():
    db = get_db()
    return render_template("index.html",
        customers=db.execute("SELECT * FROM customers ORDER BY id").fetchall(),
        companies=db.execute("SELECT * FROM companies ORDER BY id").fetchall(),
        models=db.execute("SELECT * FROM models ORDER BY id").fetchall(),
        models_json=json.dumps([dict(m) for m in db.execute("SELECT * FROM models ORDER BY id").fetchall()], ensure_ascii=False),
        currencies=CURRENCIES, incoterms=INCOTERMS, doc_types=DOC_TYPES,
        default_payment=setting("default_payment", ""), hs=HS_CODE, origin=ORIGIN,
        out_base="~/Desktop/TLT_DOOR_ORGANIZED")

@app.route("/api/customer", methods=["POST"])
@login_required
def add_customer():
    d = request.json; db = get_db()
    cur = db.execute("INSERT INTO customers(name,contact,address,phone,tax_no,country,currency,acid) VALUES(?,?,?,?,?,?,?,?)",
        (d.get("name"), d.get("contact"), d.get("address"), d.get("phone"), d.get("tax_no"), d.get("country"), d.get("currency","USD"), d.get("acid","")))
    db.commit()
    return jsonify({"id": cur.lastrowid})

@app.route("/api/rate")
@login_required
def api_rate():
    to = request.args.get("to", "USD"); return jsonify({"rate": fx_rate(to), "to": to})

# ---------------------------------------------------------------- generate
@app.route("/generate", methods=["POST"])
@login_required
def generate():
    db = get_db(); f = request.form
    items = [it for it in json.loads(f.get("items_json", "[]")) if float(it.get("qty", 0)) > 0]
    if not items:
        flash("أضف باباً واحداً على الأقل"); return redirect(url_for("index"))
    docs = request.form.getlist("docs")
    if not docs:
        flash("اختر وثيقة واحدة على الأقل"); return redirect(url_for("index"))
    containers = json.loads(f.get("containers_json", "[]"))

    subtotal = 0.0; qty_total = 0
    for it in items:
        it["qty"] = float(it["qty"]); it["unit_price"] = float(it.get("unit_price", 0))
        it["line_total"] = round(it["qty"] * it["unit_price"], 2)
        subtotal += it["line_total"]; qty_total += it["qty"]
        if it.get("model_id"):
            db.execute("UPDATE models SET last_price=? WHERE id=?", (it["unit_price"], int(it["model_id"])))
    gross = sum(float(c.get("weight") or 0) for c in containers)
    currency = f.get("currency", "USD"); rate = fx_rate(currency) or 0
    ref = next_ref()
    cust = db.execute("SELECT * FROM customers WHERE id=?", (int(f["customer_id"]),)).fetchone()

    cur = db.execute("""INSERT INTO shipments(ref_no,created_at,company_id,customer_id,incoterm,port_loading,
        port_discharge,acid_no,payment_terms,currency,fx_rate,subtotal,total,qty_total,gross_weight,
        container_count,notes,items_json,containers_json,docs_json,files_json)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
        ref, f.get("date") or datetime.date.today().isoformat(), int(f["company_id"]), int(f["customer_id"]),
        f.get("incoterm","CFR"), f.get("port_loading",""), f.get("port_discharge",""), f.get("acid_no",""),
        f.get("payment_terms",""), currency, rate, round(subtotal,2), round(subtotal,2), int(qty_total),
        gross, len(containers), f.get("notes",""), json.dumps(items, ensure_ascii=False),
        json.dumps(containers, ensure_ascii=False), json.dumps(docs), "[]"))
    db.commit(); sid = cur.lastrowid

    files, errors = _render_and_save(sid, docs, cust)
    db.execute("UPDATE shipments SET files_json=? WHERE id=?", (json.dumps(files, ensure_ascii=False), sid))
    db.commit()
    for e in errors[:3]:
        flash("⚠️ " + e)
    if not any(f.get("xlsx_rel") or f.get("pdf_rel") for f in files):
        flash("لم تُولّد ملفات. شغّل داخل مجلد export_docs:  pip install -r requirements.txt")
    return redirect(url_for("result", sid=sid))

def _ctx(sid):
    db = get_db()
    s = db.execute("SELECT * FROM shipments WHERE id=?", (sid,)).fetchone()
    if not s: return None
    company = db.execute("SELECT * FROM companies WHERE id=?", (s["company_id"],)).fetchone()
    return dict(s=s, company=company,
        cust=db.execute("SELECT * FROM customers WHERE id=?", (s["customer_id"],)).fetchone(),
        items=json.loads(s["items_json"]), containers=json.loads(s["containers_json"]),
        hs=HS_CODE, origin=ORIGIN, doc_types=DOC_TYPES, initials=initials(company))

def _render_doc(sid, doctype):
    ctx = _ctx(sid)
    if not ctx: return None
    ctx["doctype"] = doctype; ctx["doc"] = DOC_TYPES[doctype]
    return render_template(f"docs/{doctype}.html", **ctx)

def _render_and_save(sid, docs, cust):
    ctx = _ctx(sid); s = ctx["s"]
    ym = (s["created_at"] or "")[:7] or datetime.date.today().strftime("%Y-%m")
    saved, errors = [], []
    try:
        import xlsx_docs
        have_xlsx = True
    except Exception as e:
        have_xlsx = False; errors.append(f"Excel غير متاح: {e} — ثبّت openpyxl")
    try:
        from weasyprint import HTML
        have_pdf = True
    except Exception as e:
        have_pdf = False; errors.append(f"PDF غير متاح: {e}")
    for dt in docs:
        if dt not in DOC_TYPES: continue
        folder = os.path.join(OUTPUT_BASE, safe(cust["name"]), ym, DOC_TYPES[dt]["folder"])
        os.makedirs(folder, exist_ok=True)
        stem = f"{s['created_at']}_{DOC_TYPES[dt]['folder']}_{safe(cust['name'])}_{s['ref_no']}"
        xlsx_rel = pdf_rel = None
        if have_xlsx:   # Excel — الصيغة الأساسية المطلوبة
            try:
                p = os.path.join(folder, stem + ".xlsx")
                dctx = dict(ctx); dctx["doc"] = DOC_TYPES[dt]
                xlsx_docs.build(dt, dctx).save(p)
                xlsx_rel = p.replace(os.path.expanduser("~"), "~")
            except Exception as e:
                errors.append(f"{DOC_TYPES[dt]['label']} (xlsx): {e}")
        if have_pdf:    # PDF — نسخة إضافية للطباعة/التوقيع
            try:
                p = os.path.join(folder, stem + ".pdf")
                HTML(string=_render_doc(sid, dt), base_url=request.url_root).write_pdf(p)
                pdf_rel = p.replace(os.path.expanduser("~"), "~")
            except Exception as e:
                errors.append(f"{DOC_TYPES[dt]['label']} (pdf): {e}")
        saved.append({"type": dt, "label": DOC_TYPES[dt]["label"], "xlsx_rel": xlsx_rel, "pdf_rel": pdf_rel})
    return saved, errors

@app.route("/result/<int:sid>")
@login_required
def result(sid):
    s = get_db().execute("SELECT * FROM shipments WHERE id=?", (sid,)).fetchone()
    if not s: abort(404)
    files = json.loads(s["files_json"] or "[]")
    return render_template("result.html", s=s, files=files, doc_types=DOC_TYPES,
                           out_base="~/Desktop/TLT_DOOR_ORGANIZED")

@app.route("/shipment/<int:sid>/doc/<doctype>")
@login_required
def view_doc(sid, doctype):
    if doctype not in DOC_TYPES: abort(404)
    html = _render_doc(sid, doctype)
    if html is None: abort(404)
    return html

@app.route("/shipment/<int:sid>/doc/<doctype>/pdf")
@login_required
def doc_pdf(sid, doctype):
    from weasyprint import HTML
    if doctype not in DOC_TYPES: abort(404)
    html = _render_doc(sid, doctype)
    if html is None: abort(404)
    pdf = HTML(string=html, base_url=request.url_root).write_pdf()
    s = get_db().execute("SELECT ref_no FROM shipments WHERE id=?", (sid,)).fetchone()
    return Response(pdf, mimetype="application/pdf",
                   headers={"Content-Disposition": f"inline; filename={s['ref_no']}_{doctype}.pdf"})

@app.route("/shipment/<int:sid>/doc/<doctype>/xlsx")
@login_required
def doc_xlsx(sid, doctype):
    import xlsx_docs
    if doctype not in DOC_TYPES: abort(404)
    ctx = _ctx(sid)
    if not ctx: abort(404)
    ctx["doc"] = DOC_TYPES[doctype]
    data = xlsx_docs.to_bytes(xlsx_docs.build(doctype, ctx))
    s = get_db().execute("SELECT ref_no FROM shipments WHERE id=?", (sid,)).fetchone()
    return Response(data, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   headers={"Content-Disposition": f"attachment; filename={s['ref_no']}_{doctype}.xlsx"})

# ---------------------------------------------------------------- history / settings
@app.route("/history")
@login_required
def history():
    q = request.args.get("q", "").strip(); db = get_db()
    sql = "SELECT s.*, c.name cust_name FROM shipments s JOIN customers c ON c.id=s.customer_id"
    p = []
    if q:
        sql += " WHERE s.ref_no LIKE ? OR c.name LIKE ? OR s.created_at LIKE ?"; p = [f"%{q}%"]*3
    sql += " ORDER BY s.id DESC"
    return render_template("history.html", rows=db.execute(sql, p).fetchall(), q=q, doc_types=DOC_TYPES)

@app.route("/shipment/<int:sid>/duplicate")
@login_required
def duplicate(sid):
    db = get_db(); s = db.execute("SELECT * FROM shipments WHERE id=?", (sid,)).fetchone()
    if not s: abort(404)
    dup = dict(s); dup["items"] = json.loads(s["items_json"]); dup["containers"] = json.loads(s["containers_json"])
    dup["docs"] = json.loads(s["docs_json"])
    return render_template("index.html",
        customers=db.execute("SELECT * FROM customers ORDER BY id").fetchall(),
        companies=db.execute("SELECT * FROM companies ORDER BY id").fetchall(),
        models=db.execute("SELECT * FROM models ORDER BY id").fetchall(),
        models_json=json.dumps([dict(m) for m in db.execute("SELECT * FROM models ORDER BY id").fetchall()], ensure_ascii=False),
        currencies=CURRENCIES, incoterms=INCOTERMS, doc_types=DOC_TYPES,
        default_payment=s["payment_terms"], hs=HS_CODE, origin=ORIGIN,
        out_base="~/Desktop/TLT_DOOR_ORGANIZED", dup=dup,
        dup_json=json.dumps({"items": dup["items"], "containers": dup["containers"], "docs": dup["docs"]}, ensure_ascii=False))

@app.route("/shipment/<int:sid>/delete", methods=["POST"])
@login_required
def delete(sid):
    shutil.copy2(DB_PATH, os.path.join(BACKUP_DIR, f"akva_{datetime.datetime.now():%Y%m%d_%H%M%S}.db"))
    get_db().execute("DELETE FROM shipments WHERE id=?", (sid,)); get_db().commit()
    flash("تم الحذف (مع نسخة احتياطية)"); return redirect(url_for("history"))

@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings_page():
    db = get_db()
    if request.method == "POST":
        sec = request.form.get("section")
        if sec == "company":
            db.execute("""UPDATE companies SET name=?,address=?,phone=?,email=?,bank_name=?,branch=?,
                branch_code=?,iban=?,swift=? WHERE id=?""", (request.form["name"], request.form["address"],
                request.form["phone"], request.form["email"], request.form["bank_name"], request.form["branch"],
                request.form["branch_code"], request.form["iban"], request.form["swift"], int(request.form["id"])))
            db.commit(); flash("تم حفظ الشركة")
        elif sec == "models":
            for m in db.execute("SELECT id FROM models").fetchall():
                v = request.form.get(f"price_{m['id']}")
                if v is not None: db.execute("UPDATE models SET default_price=? WHERE id=?", (float(v or 0), m["id"]))
            db.commit(); flash("تم حفظ الأسعار")
        elif sec == "payment":
            set_setting("default_payment", request.form.get("default_payment", "")); flash("تم الحفظ")
        return redirect(url_for("settings_page"))
    return render_template("settings.html",
        companies=db.execute("SELECT * FROM companies ORDER BY id").fetchall(),
        models=db.execute("SELECT * FROM models ORDER BY id").fetchall(),
        default_payment=setting("default_payment", ""))

if __name__ == "__main__":
    import sys
    port = 8000
    if len(sys.argv) > 1 and sys.argv[1].isdigit(): port = int(sys.argv[1])
    elif os.environ.get("PORT", "").isdigit(): port = int(os.environ["PORT"])
    print("\n  MN STEEL DOOR — AKVA Export Docs (Pro)")
    print(f"  افتح المتصفح على:  http://localhost:{port}")
    print(f"  مخرجات الوثائق في:  {OUTPUT_BASE}\n")
    app.run(host="127.0.0.1", port=port, debug=False)
