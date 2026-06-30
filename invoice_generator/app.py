# -*- coding: utf-8 -*-
"""
MN STEEL DOOR — Invoice Generator
مولّد فواتير احترافي لأعمال تصدير الأبواب الصلب — يعمل محلياً على localhost.
"""
import os, json, sqlite3, math, shutil, datetime, re, secrets, functools
from flask import (Flask, request, session, redirect, url_for, render_template,
                   g, jsonify, abort, send_file, Response, flash)
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import requests
except Exception:
    requests = None

# ---------------------------------------------------------------- paths
DATA_DIR = os.path.expanduser("~/Desktop/door_ops")
DB_PATH = os.path.join(DATA_DIR, "invoice.db")
INVOICES_DIR = os.path.join(DATA_DIR, "invoices")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
for d in (DATA_DIR, INVOICES_DIR, BACKUP_DIR):
    os.makedirs(d, exist_ok=True)

# ---------------------------------------------------------------- constants
HS_CODE = "7308300000"
ORIGIN = "TURKEY"
FRAME_CM = 15
DEFAULT_PASSWORD = "MN2026"
CURRENCIES = ["USD", "EGP", "IQD", "JOD", "SAR", "AED", "LYD", "TRY"]

MODEL_SEED = [
    ("DÜZ LAMİNOKS", "DÜZ LAMİNOKS", "100×215×15"),
    ("UV GENİŞ AÇILI", "UV GENİŞ AÇILI", "100×215×15"),
]
for _c in [46,101,106,113,115,118,119,120,144,150,194,207,210,244,279,314,328,356,360,365,670]:
    MODEL_SEED.append((f"Code {_c}", f"Code {_c}", "100×215×15"))

DEFAULT_COMMISSION = {
    "iraq_per_m2": 3.0,
    "tiers": [   # عمولة العادي متدرّجة حسب عرض الباب (سم)
        {"max_w": 90, "amt": 5.0},
        {"max_w": 100, "amt": 7.0},
        {"max_w": 9999, "amt": 10.0},
    ],
}
DEFAULT_MARGIN = {"red_pct": 10.0, "yellow_pct": 25.0}   # عتبات حارس الربح
DEFAULT_CONTAINER_CAP = 220   # عدد الأبواب القياسي لكل كونتينر (قابل للتعديل)

app = Flask(__name__)

# ---------------------------------------------------------------- db
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys=ON")
    return g.db

@app.teardown_appcontext
def close_db(exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()

def setting(key, default=None):
    row = get_db().execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    if not row:
        return default
    try:
        return json.loads(row["value"])
    except Exception:
        return row["value"]

def set_setting(key, value):
    v = json.dumps(value) if not isinstance(value, str) else value
    get_db().execute(
        "INSERT INTO settings(key,value) VALUES(?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, v))
    get_db().commit()

def init_db():
    db = sqlite3.connect(DB_PATH)
    db.executescript("""
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS companies(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, address TEXT, phone TEXT,
        email TEXT, bank_name TEXT, branch TEXT, branch_code TEXT, iban TEXT, swift TEXT);
    CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, contact TEXT, address TEXT,
        phone TEXT, tax_no TEXT, country TEXT, currency TEXT DEFAULT 'USD', email TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS models(
        id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT, name TEXT, size TEXT,
        default_price REAL DEFAULT 0, cost_price REAL DEFAULT 0, last_price REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS customer_prices(
        customer_id INTEGER, model_id INTEGER, price REAL,
        PRIMARY KEY(customer_id, model_id));
    CREATE TABLE IF NOT EXISTS invoices(
        id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT UNIQUE, inv_type TEXT,
        created_at TEXT, company_id INTEGER, customer_id INTEGER, incoterm TEXT,
        port_loading TEXT, port_discharge TEXT, acid_no TEXT,
        advance_pct REAL, balance_pct REAL, delivery_weeks INTEGER,
        currency TEXT, fx_rate REAL, subtotal REAL, total REAL,
        advance_amt REAL, balance_amt REAL, commission REAL, container_count INTEGER,
        notes TEXT, items_json TEXT);
    """)
    db.commit()
    cur = db.cursor()

    # seed settings
    def _seed_setting(k, v):
        if not cur.execute("SELECT 1 FROM settings WHERE key=?", (k,)).fetchone():
            cur.execute("INSERT INTO settings(key,value) VALUES(?,?)", (k, json.dumps(v)))

    if not cur.execute("SELECT 1 FROM settings WHERE key='password_hash'").fetchone():
        cur.execute("INSERT INTO settings(key,value) VALUES('password_hash',?)",
                    (generate_password_hash(DEFAULT_PASSWORD),))
        cur.execute("INSERT INTO settings(key,value) VALUES('must_change','true')")
    if not cur.execute("SELECT 1 FROM settings WHERE key='secret_key'").fetchone():
        cur.execute("INSERT INTO settings(key,value) VALUES('secret_key',?)",
                    (json.dumps(secrets.token_hex(24)),))
    _seed_setting("invoice_seq", 0)
    _seed_setting("commission", DEFAULT_COMMISSION)
    _seed_setting("margin", DEFAULT_MARGIN)
    _seed_setting("container_cap", DEFAULT_CONTAINER_CAP)

    # seed companies
    if not cur.execute("SELECT 1 FROM companies").fetchone():
        cur.execute("""INSERT INTO companies(name,address,phone,email,bank_name,branch,branch_code,iban,swift)
            VALUES(?,?,?,?,?,?,?,?,?)""", (
            "SAIROGULLARI DIS TICARET LTD STI",
            "OYMAAGAC MAH. 5062 CAD, KAYSERI / TURKEY",
            "+90 541 924 98 70", "talat@sairogullari.com",
            "Türkiye İş Bankası A.Ş.", "Kayseri O.S.B.", "4000",
            "TR26 0006 4000 0025 3060 0556 24", "ISBKTRIS"))
        # company 2 & 3 placeholders (fill later)
        cur.execute("INSERT INTO companies(name,address,phone,email,bank_name,branch,branch_code,iban,swift)"
                    " VALUES('COMPANY 2 — (املأ بياناتها)','','','','','','','','')")
        cur.execute("INSERT INTO companies(name,address,phone,email,bank_name,branch,branch_code,iban,swift)"
                    " VALUES('COMPANY 3 — (املأ بياناتها)','','','','','','','','')")

    # seed customers (from seed_customers.json — مولّد من جهات اتصالك الفعلية)
    if not cur.execute("SELECT 1 FROM customers").fetchone():
        seed_path = os.path.join(os.path.dirname(__file__), "seed_customers.json")
        seed_list = []
        if os.path.exists(seed_path):
            try:
                seed_list = json.load(open(seed_path, encoding="utf-8"))
            except Exception:
                seed_list = []
        if not seed_list:
            seed_list = [{"name": "MODERN STYLE FOR IMPORT & EXPORT", "contact": "SAAD ZAKY FAHIM",
                          "address": "99 Mostafa El-Nahhas St, Nasr City, 8th District, Cairo, Egypt",
                          "phone": "+20 121 116 4530", "tax_no": "741496100",
                          "country": "Egypt", "currency": "EGP", "email": ""}]
        for c in seed_list:
            cur.execute("""INSERT INTO customers(name,contact,address,phone,tax_no,country,currency,email)
                VALUES(?,?,?,?,?,?,?,?)""", (
                c.get("name", ""), c.get("contact", ""), c.get("address", ""),
                c.get("phone", ""), c.get("tax_no", ""), c.get("country", ""),
                c.get("currency", "USD"), c.get("email", "")))

    # seed models
    if not cur.execute("SELECT 1 FROM models").fetchone():
        for code, name, size in MODEL_SEED:
            cur.execute("INSERT INTO models(code,name,size,default_price,cost_price,last_price)"
                        " VALUES(?,?,?,0,0,0)", (code, name, size))
    db.commit()
    db.close()

with app.app_context():
    init_db()
    with sqlite3.connect(DB_PATH) as _c:
        _c.row_factory = sqlite3.Row
        app.secret_key = json.loads(_c.execute(
            "SELECT value FROM settings WHERE key='secret_key'").fetchone()["value"])

# ---------------------------------------------------------------- auth
def login_required(f):
    @functools.wraps(f)
    def w(*a, **k):
        if not session.get("auth"):
            return redirect(url_for("login"))
        if setting("must_change") in (True, "true") and request.endpoint not in ("change_password", "logout", "static"):
            return redirect(url_for("change_password"))
        return f(*a, **k)
    return w

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        pw = request.form.get("password", "")
        ph = setting("password_hash")
        if ph and check_password_hash(ph, pw):
            session["auth"] = True
            return redirect(url_for("index"))
        flash("كلمة السر غير صحيحة")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/change-password", methods=["GET", "POST"])
def change_password():
    if not session.get("auth"):
        return redirect(url_for("login"))
    if request.method == "POST":
        new = request.form.get("new", "")
        conf = request.form.get("confirm", "")
        if len(new) < 6:
            flash("كلمة السر يجب أن تكون 6 أحرف على الأقل")
        elif new != conf:
            flash("كلمتا السر غير متطابقتين")
        else:
            set_setting("password_hash", generate_password_hash(new))
            set_setting("must_change", "false")
            flash("تم تغيير كلمة السر بنجاح")
            return redirect(url_for("index"))
    return render_template("change_password.html")

# ---------------------------------------------------------------- helpers
def model_width(size):
    m = re.match(r"\s*(\d+)", size or "")
    return int(m.group(1)) if m else 100

def model_area(size):
    nums = re.findall(r"\d+", size or "")
    if len(nums) >= 2:
        return (int(nums[0]) / 100.0) * (int(nums[1]) / 100.0)
    return 2.15

def next_invoice_no():
    seq = int(setting("invoice_seq", 0)) + 1
    set_setting("invoice_seq", seq)
    return f"SGL{datetime.date.today().year}{seq:07d}"

def commission_for(items, country):
    cfg = setting("commission", DEFAULT_COMMISSION)
    total = 0.0
    iraq = (country or "").strip().lower() in ("iraq", "العراق", "عراق")
    for it in items:
        qty = float(it.get("qty", 0))
        size = it.get("size", "")
        if iraq:
            total += cfg["iraq_per_m2"] * model_area(size) * qty
        else:
            w = model_width(size)
            amt = cfg["tiers"][-1]["amt"]
            for t in cfg["tiers"]:
                if w <= t["max_w"]:
                    amt = t["amt"]; break
            total += amt * qty
    return round(total, 2)

def fx_rate(to_cur):
    if not to_cur or to_cur == "USD" or requests is None:
        return 1.0
    try:
        r = requests.get(f"https://api.frankfurter.app/latest?from=USD&to={to_cur}", timeout=6)
        return float(r.json()["rates"][to_cur])
    except Exception:
        return None

def safe_name(s):
    return re.sub(r"[^\w؀-ۿ .-]", "_", s or "").strip()[:60] or "customer"

# ---------------------------------------------------------------- main screens
@app.route("/")
@login_required
def index():
    db = get_db()
    customers = db.execute("SELECT * FROM customers ORDER BY name").fetchall()
    companies = db.execute("SELECT * FROM companies ORDER BY id").fetchall()
    models = db.execute("SELECT * FROM models ORDER BY id").fetchall()
    models_json = json.dumps([dict(m) for m in models], ensure_ascii=False)
    return render_template("index.html", customers=customers, companies=companies,
                           models=models, models_json=models_json, dup_items_json="null",
                           hs=HS_CODE, origin=ORIGIN,
                           container_cap=setting("container_cap", DEFAULT_CONTAINER_CAP),
                           margin=setting("margin", DEFAULT_MARGIN),
                           currencies=CURRENCIES, dup=None)

@app.route("/api/customer", methods=["POST"])
@login_required
def add_customer():
    d = request.json
    db = get_db()
    cur = db.execute("""INSERT INTO customers(name,contact,address,phone,tax_no,country,currency,email)
        VALUES(?,?,?,?,?,?,?,?)""", (d.get("name"), d.get("contact"), d.get("address"),
        d.get("phone"), d.get("tax_no"), d.get("country"), d.get("currency", "USD"), d.get("email", "")))
    db.commit()
    return jsonify({"id": cur.lastrowid})

@app.route("/api/customer-prices/<int:cid>")
@login_required
def customer_prices(cid):
    rows = get_db().execute("SELECT model_id, price FROM customer_prices WHERE customer_id=?",
                            (cid,)).fetchall()
    cust = get_db().execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    return jsonify({"prices": {r["model_id"]: r["price"] for r in rows},
                    "country": cust["country"] if cust else "",
                    "currency": cust["currency"] if cust else "USD"})

@app.route("/api/rate")
@login_required
def api_rate():
    to = request.args.get("to", "USD")
    r = fx_rate(to)
    return jsonify({"rate": r, "to": to})

# ---------------------------------------------------------------- create invoice
@app.route("/invoice/create", methods=["POST"])
@login_required
def create_invoice():
    db = get_db()
    f = request.form
    items = json.loads(f.get("items_json", "[]"))
    items = [it for it in items if float(it.get("qty", 0)) > 0]
    if not items:
        flash("أضف بنداً واحداً على الأقل"); return redirect(url_for("index"))

    customer_id = int(f["customer_id"]); company_id = int(f["company_id"])
    cust = db.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()

    subtotal = 0.0
    for it in items:
        it["qty"] = float(it["qty"]); it["unit_price"] = float(it["unit_price"])
        it["line_total"] = round(it["qty"] * it["unit_price"], 2)
        subtotal += it["line_total"]
        # price memory (global + per customer)
        mid = int(it["model_id"])
        db.execute("UPDATE models SET last_price=? WHERE id=?", (it["unit_price"], mid))
        db.execute("INSERT INTO customer_prices(customer_id,model_id,price) VALUES(?,?,?) "
                   "ON CONFLICT(customer_id,model_id) DO UPDATE SET price=excluded.price",
                   (customer_id, mid, it["unit_price"]))

    subtotal = round(subtotal, 2)
    advance_pct = float(f.get("advance_pct", 30)); balance_pct = float(f.get("balance_pct", 70))
    total = subtotal
    advance_amt = round(total * advance_pct / 100.0, 2)
    balance_amt = round(total - advance_amt, 2)
    commission = commission_for(items, cust["country"] if cust else "")
    cap = int(setting("container_cap", DEFAULT_CONTAINER_CAP))
    qty_total = sum(it["qty"] for it in items)
    container_count = math.ceil(qty_total / cap) if cap else 0
    currency = f.get("currency", "USD")
    rate = fx_rate(currency) or 0

    inv_no = next_invoice_no()
    cur = db.execute("""INSERT INTO invoices(invoice_no,inv_type,created_at,company_id,customer_id,
        incoterm,port_loading,port_discharge,acid_no,advance_pct,balance_pct,delivery_weeks,
        currency,fx_rate,subtotal,total,advance_amt,balance_amt,commission,container_count,notes,items_json)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
        inv_no, f.get("inv_type", "Proforma"), datetime.date.today().isoformat(),
        company_id, customer_id, f.get("incoterm", "FOB"), f.get("port_loading", ""),
        f.get("port_discharge", ""), f.get("acid_no", ""), advance_pct, balance_pct,
        int(f.get("delivery_weeks", 6)), currency, rate, subtotal, total,
        advance_amt, balance_amt, commission, container_count, f.get("notes", ""),
        json.dumps(items, ensure_ascii=False)))
    db.commit()
    inv_id = cur.lastrowid

    # archive both PDF copies
    try:
        _archive_pdfs(inv_id)
    except Exception as e:
        flash(f"تم حفظ الفاتورة لكن تعذّر توليد PDF: {e}")
    return redirect(url_for("view_invoice", inv_id=inv_id, copy=f.get("inv_type", "Proforma").lower()))

def _build_context(inv_id):
    db = get_db()
    inv = db.execute("SELECT * FROM invoices WHERE id=?", (inv_id,)).fetchone()
    if not inv:
        return None
    company = db.execute("SELECT * FROM companies WHERE id=?", (inv["company_id"],)).fetchone()
    cust = db.execute("SELECT * FROM customers WHERE id=?", (inv["customer_id"],)).fetchone()
    items = json.loads(inv["items_json"])
    return dict(inv=inv, company=company, cust=cust, items=items,
                hs=HS_CODE, origin=ORIGIN, frame=FRAME_CM)

def render_invoice(inv_id, copy_type):
    ctx = _build_context(inv_id)
    if not ctx:
        return None
    ctx["copy_type"] = "Commercial" if copy_type.lower().startswith("comm") else "Proforma"
    return render_template("invoice.html", **ctx)

def _archive_pdfs(inv_id):
    from weasyprint import HTML
    ctx = _build_context(inv_id)
    inv = ctx["inv"]; cust = ctx["cust"]
    folder = os.path.join(INVOICES_DIR, safe_name(cust["name"]))
    os.makedirs(folder, exist_ok=True)
    for copy_type in ("Proforma", "Commercial"):
        html = render_invoice(inv_id, copy_type)
        fname = f"{inv['created_at']}_{copy_type}_{safe_name(cust['name'])}_{inv['invoice_no']}.pdf"
        HTML(string=html, base_url=request.url_root).write_pdf(os.path.join(folder, fname))

# ---------------------------------------------------------------- view / pdf
@app.route("/invoice/<int:inv_id>")
@login_required
def view_invoice(inv_id):
    copy_type = request.args.get("copy", "proforma")
    html = render_invoice(inv_id, copy_type)
    if html is None:
        abort(404)
    return html

@app.route("/invoice/<int:inv_id>/pdf")
@login_required
def invoice_pdf(inv_id):
    from weasyprint import HTML
    copy_type = request.args.get("copy", "proforma")
    html = render_invoice(inv_id, copy_type)
    if html is None:
        abort(404)
    pdf = HTML(string=html, base_url=request.url_root).write_pdf()
    inv = get_db().execute("SELECT invoice_no FROM invoices WHERE id=?", (inv_id,)).fetchone()
    return Response(pdf, mimetype="application/pdf", headers={
        "Content-Disposition": f"inline; filename={inv['invoice_no']}_{copy_type}.pdf"})

# ---------------------------------------------------------------- history
@app.route("/history")
@login_required
def history():
    q = request.args.get("q", "").strip()
    db = get_db()
    sql = """SELECT i.*, c.name AS cust_name FROM invoices i
             JOIN customers c ON c.id=i.customer_id"""
    params = []
    if q:
        sql += " WHERE i.invoice_no LIKE ? OR c.name LIKE ? OR i.created_at LIKE ?"
        params = [f"%{q}%"] * 3
    sql += " ORDER BY i.id DESC"
    rows = db.execute(sql, params).fetchall()
    return render_template("history.html", rows=rows, q=q)

@app.route("/invoice/<int:inv_id>/duplicate")
@login_required
def duplicate_invoice(inv_id):
    db = get_db()
    inv = db.execute("SELECT * FROM invoices WHERE id=?", (inv_id,)).fetchone()
    if not inv:
        abort(404)
    customers = db.execute("SELECT * FROM customers ORDER BY name").fetchall()
    companies = db.execute("SELECT * FROM companies ORDER BY id").fetchall()
    models = db.execute("SELECT * FROM models ORDER BY id").fetchall()
    dup = dict(inv); dup["items"] = json.loads(inv["items_json"])
    models_json = json.dumps([dict(m) for m in models], ensure_ascii=False)
    return render_template("index.html", customers=customers, companies=companies,
                           models=models, models_json=models_json,
                           dup_items_json=json.dumps(dup["items"], ensure_ascii=False),
                           hs=HS_CODE, origin=ORIGIN,
                           container_cap=setting("container_cap", DEFAULT_CONTAINER_CAP),
                           margin=setting("margin", DEFAULT_MARGIN),
                           currencies=CURRENCIES, dup=dup)

@app.route("/invoice/<int:inv_id>/delete", methods=["POST"])
@login_required
def delete_invoice(inv_id):
    # backup before delete
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(DB_PATH, os.path.join(BACKUP_DIR, f"invoice_{ts}.db"))
    get_db().execute("DELETE FROM invoices WHERE id=?", (inv_id,))
    get_db().commit()
    flash("تم حذف الفاتورة (مع أخذ نسخة احتياطية)")
    return redirect(url_for("history"))

# ---------------------------------------------------------------- settings
@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings_page():
    db = get_db()
    if request.method == "POST":
        sec = request.form.get("section")
        if sec == "company":
            cid = int(request.form["id"])
            db.execute("""UPDATE companies SET name=?,address=?,phone=?,email=?,bank_name=?,
                branch=?,branch_code=?,iban=?,swift=? WHERE id=?""", (
                request.form["name"], request.form["address"], request.form["phone"],
                request.form["email"], request.form["bank_name"], request.form["branch"],
                request.form["branch_code"], request.form["iban"], request.form["swift"], cid))
            db.commit(); flash("تم حفظ بيانات الشركة")
        elif sec == "models":
            for m in db.execute("SELECT id FROM models").fetchall():
                dp = request.form.get(f"default_{m['id']}")
                cp = request.form.get(f"cost_{m['id']}")
                if dp is not None:
                    db.execute("UPDATE models SET default_price=?, cost_price=? WHERE id=?",
                               (float(dp or 0), float(cp or 0), m["id"]))
            db.commit(); flash("تم حفظ أسعار الموديلات")
        elif sec == "general":
            set_setting("container_cap", int(request.form.get("container_cap", 220)))
            set_setting("margin", {"red_pct": float(request.form.get("red_pct", 10)),
                                   "yellow_pct": float(request.form.get("yellow_pct", 25))})
            flash("تم حفظ الإعدادات العامة")
        return redirect(url_for("settings_page"))
    companies = db.execute("SELECT * FROM companies ORDER BY id").fetchall()
    models = db.execute("SELECT * FROM models ORDER BY id").fetchall()
    return render_template("settings.html", companies=companies, models=models,
                           container_cap=setting("container_cap", DEFAULT_CONTAINER_CAP),
                           margin=setting("margin", DEFAULT_MARGIN))

if __name__ == "__main__":
    print("\n  MN STEEL DOOR — Invoice Generator")
    print("  افتح المتصفح على:  http://localhost:5000")
    print(f"  بيانات العمل في:  {DATA_DIR}\n")
    app.run(host="127.0.0.1", port=5000, debug=False)
