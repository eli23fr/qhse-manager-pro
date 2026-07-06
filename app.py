
# ============================================================
# QHSE Manager Pro - Chemical Register
# Version: v1.0 STABLE
# Copyright © 2026 Kodjotse Eli ADIGBLI. Tous droits réservés.
# ============================================================

from flask import Flask, request, redirect, url_for, session, send_file, render_template_string, flash, get_flashed_messages
import os
import sqlite3
from pathlib import Path
from datetime import datetime, date
import webbrowser
import threading
from markupsafe import escape
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_NAME = "QHSE Manager Pro - Chemical Register"
VERSION = "v2.3 ALIAS MODIFIABLES / SUPPRIMABLES"
COPYRIGHT = "Copyright © 2026 Kodjotse Eli ADIGBLI. Tous droits réservés."

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "qhse_chemical_register.db"
EXPORT_DIR = BASE_DIR / "exports"
UPLOAD_DIR = BASE_DIR / "uploads"
EXPORT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "qhse_manager_pro_change_this_key")
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax"
)
if os.environ.get("CLOUD_MODE", "false").lower() == "true":
    app.config["SESSION_COOKIE_SECURE"] = True

@app.after_request
def add_security_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    return response

CSS = """
<style>
:root{--main:#1f3347;--light:#eef3f7;--border:#cfd8e3;--danger:#b00020;--ok:#087a2c;--warn:#b36b00}
body{font-family:Arial,sans-serif;margin:0;background:#f6f8fb;color:#222}
header{background:var(--main);color:white;padding:14px 24px;display:flex;justify-content:space-between;align-items:center}
header h2{margin:0;font-size:21px}header small{opacity:.9}
nav{background:#e7edf3;padding:10px 24px;border-bottom:1px solid var(--border)}
nav a{margin-right:14px;color:var(--main);font-weight:bold;text-decoration:none}
main{padding:24px}.footer{margin-top:28px;color:#666;font-size:12px}
table{border-collapse:collapse;width:100%;background:white;margin-top:16px}
th,td{border:1px solid var(--border);padding:8px;text-align:left;vertical-align:top}
th{background:var(--main);color:white}
input,select,textarea{width:100%;padding:8px;margin:4px 0 10px 0;box-sizing:border-box;border:1px solid var(--border);border-radius:4px}
button,.btn{background:var(--main);color:white;padding:9px 14px;border:0;border-radius:4px;text-decoration:none;display:inline-block;cursor:pointer}
.box{background:white;padding:16px;border:1px solid var(--border);border-radius:8px;max-width:1200px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.row{display:grid;grid-template-columns:1fr 1fr;gap:16px}.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
.card{background:white;border:1px solid var(--border);border-radius:8px;padding:16px}.value{font-size:28px;font-weight:bold;margin-top:8px}
.danger{color:var(--danger);font-weight:bold}.ok{color:var(--ok);font-weight:bold}.warn{color:var(--warn);font-weight:bold}
.btn-danger{background:var(--danger)}.btn-warn{background:var(--warn)}.actions{white-space:nowrap}.muted{color:#666;font-size:12px}.mini{font-size:12px;padding:5px 8px}
.flash{background:#fff2cc;padding:10px;border:1px solid #e0c36c;margin-bottom:12px;border-radius:6px}
.help{background:#eef7ff;border:1px solid #bed7ee;padding:12px;border-radius:8px;margin-bottom:14px}
.badge{display:inline-block;padding:3px 8px;border-radius:12px;background:#eef3f7;font-size:12px}
@media(max-width:900px){.grid,.row,.row3{grid-template-columns:1fr}nav a{display:inline-block;margin-bottom:8px}}
</style>
"""

# ---------- DB helpers ----------
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def ensure_column(conn, table, col, definition):
    cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if col not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {definition}")

def init_db():
    conn = db()
    cur = conn.cursor()

    cur.execute("""CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'super_admin',
        subcontractor_id INTEGER,
        project_id INTEGER,
        status TEXT DEFAULT 'Actif'
    )""")
    for col, definition in [
        ("subcontractor_id","INTEGER"),("project_id","INTEGER"),("status","TEXT DEFAULT 'Actif'")
    ]:
        ensure_column(conn, "users", col, definition)

    cur.execute("""CREATE TABLE IF NOT EXISTS projects(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        location TEXT,
        status TEXT DEFAULT 'Actif'
    )""")
    for col, definition in [
        ("client","TEXT"),("owner","TEXT"),("main_contractor","TEXT"),("project_manager","TEXT"),
        ("qhse_manager","TEXT"),("coordinator","TEXT"),("country","TEXT"),("start_date","TEXT"),
        ("end_date","TEXT"),("description","TEXT")
    ]:
        ensure_column(conn, "projects", col, definition)

    cur.execute("""CREATE TABLE IF NOT EXISTS subcontractors(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        contact TEXT,
        phone TEXT,
        email TEXT,
        status TEXT DEFAULT 'Actif'
    )""")
    for col, definition in [
        ("address","TEXT"),("lot","TEXT"),("work_zone","TEXT"),("employees_count","INTEGER DEFAULT 0"),
        ("start_date","TEXT"),("end_date","TEXT"),("hse_plan","TEXT DEFAULT 'NON'"),
        ("insurance","TEXT DEFAULT 'NON'"),("ppsps","TEXT DEFAULT 'NON'"),("observations","TEXT")
    ]:
        ensure_column(conn, "subcontractors", col, definition)

    cur.execute("""CREATE TABLE IF NOT EXISTS products(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        root_name TEXT UNIQUE NOT NULL,
        manufacturer TEXT,
        family TEXT,
        physical_state TEXT,
        conditionnement TEXT,
        utilisation TEXT,
        fds TEXT DEFAULT 'NON',
        pictogrammes TEXT,
        incompatibilites TEXT,
        danger_class TEXT,
        epi TEXT,
        observations TEXT
    )""")
    for col, definition in [
        ("commercial_name","TEXT"),("reference","TEXT"),("h_statements","TEXT"),("p_statements","TEXT"),
        ("first_aid","TEXT"),("fire_measures","TEXT"),("spill_measures","TEXT"),("storage_rules","TEXT"),
        ("fds_date","TEXT"),("fds_version","TEXT"),("fds_expiry","TEXT"),
        ("stock_min","REAL DEFAULT 0"),("stock_max","REAL DEFAULT 0")
    ]:
        ensure_column(conn, "products", col, definition)

    cur.execute("""CREATE TABLE IF NOT EXISTS aliases(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        alias_name TEXT UNIQUE NOT NULL,
        product_id INTEGER NOT NULL
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS entries(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER,
        entry_date TEXT NOT NULL,
        month TEXT NOT NULL,
        subcontractor_id INTEGER NOT NULL,
        declared_product TEXT NOT NULL,
        product_id INTEGER,
        qty_in REAL DEFAULT 0,
        qty_used REAL DEFAULT 0,
        qty_stock REAL DEFAULT 0,
        unit TEXT DEFAULT 'u',
        storage_location TEXT,
        fds_available TEXT DEFAULT 'NON',
        observations TEXT
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS audit_log(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action_date TEXT,
        username TEXT,
        action TEXT,
        details TEXT
    )""")

    cur.execute("UPDATE users SET role='super_admin' WHERE role='admin'")
    if cur.execute("SELECT COUNT(*) c FROM users").fetchone()["c"] == 0:
        admin_username = os.environ.get("ADMIN_USERNAME", "admin")
        admin_password = os.environ.get("ADMIN_PASSWORD")
        cloud_mode = os.environ.get("CLOUD_MODE", "false").lower() == "true"
        if not admin_password:
            if cloud_mode:
                raise RuntimeError("ADMIN_PASSWORD doit être défini dans Render.")
            admin_password = "admin123"
        cur.execute("INSERT INTO users(username,password_hash,role,status) VALUES(?,?,?,?)",
                    (admin_username, generate_password_hash(admin_password), "super_admin", "Actif"))

    cur.execute("INSERT OR IGNORE INTO projects(name,location,country,main_contractor,status) VALUES(?,?,?,?,?)",
                ("CHU Kara", "Kara", "Togo", "Ellipse Projects", "Actif"))
    cur.execute("INSERT OR IGNORE INTO projects(name,location,country,main_contractor,status) VALUES(?,?,?,?,?)",
                ("CHU Campus Lomé", "Lomé", "Togo", "Ellipse Projects", "Actif"))

    for st in ["SICONE","CENTRO","BETEIR","BATICO","MGP"]:
        cur.execute("INSERT OR IGNORE INTO subcontractors(name,status) VALUES(?,?)", (st, "Actif"))

    sample_products = [
        ("AQUADERE","Résine / primaire","Liquide","Bidon / seau","Résine d’accrochage pour mortier et béton","NON","","","Résine / primaire","Gants, lunettes",0,0),
        ("SOPRALENE","Étanchéité","Solide","Rouleau membrane","Étanchéité des toitures et ouvrages","NON","","Éloigner sources de chaleur","Produit bitumineux","Gants, chaussures de sécurité",5,250),
        ("EQUERRE DE KENF","Accessoire","Solide","Carton / pièce","Accessoire de fixation et assemblage","NON","","","Accessoire","Gants",0,0),
        ("PANTEX VELOUR","Peinture","Liquide","Seau plastique 15 à 25 kg","Peinture de finition intérieure murs et plafonds","NON","SGH07 possible","Oxydants forts","Peinture","Gants, lunettes",5,300),
        ("PRIMAIRE D ACCROCHE","Primaire","Liquide","Seau 5 à 20 kg","Préparation des supports avant revêtement","NON","SGH07 possible","Oxydants forts","Primaire","Gants, lunettes",0,200),
        ("SIGMA COVER","Peinture","Liquide","Seau peinture","Peinture de finition et protection","NON","SGH07 possible","Oxydants forts","Peinture","Gants, lunettes",0,100),
        ("FLINTKOTE IKOPRIM","Étanchéité","Liquide","Bidon / seau","Primaire bitumineux pour travaux d’étanchéité","NON","SGH02/SGH07 possible","Sources d’ignition, oxydants","Bitume / primaire","Gants, lunettes, ventilation",0,80),
        ("FLINTKOTE BE6","Étanchéité","Liquide","Seau / bidon","Revêtement bitumineux de protection et d’étanchéité","NON","SGH02/SGH07 possible","Sources d’ignition, oxydants","Bitume","Gants, lunettes",0,80),
        ("ULTRABONDE CONDUCTIVITE","Colle","Pâte","Seau 16 kg","Colle conductrice pour revêtements de sol techniques","NON","SGH07 possible","Oxydants forts","Colle","Gants, lunettes",0,80),
        ("ULTRABONDE EVOLUTION","Colle","Pâte","Seau 14 kg","Colle pour revêtements de sols et murs","NON","SGH07 possible","Oxydants forts","Colle","Gants, lunettes",0,150),
        ("MC TECHNIFLOW 57GH","Mortier / Coulis","Poudre","Sac 25 kg","Coulis de scellement et calage haute performance","NON","SGH07 possible","Acides forts","Coulis cimentaire","Gants, lunettes, masque poussières",0,100),
        ("AFRICA ENDUIT","Enduit","Poudre","Sac 25 kg","Enduit de lissage et de finition des murs","NON","SGH07 possible","Acides forts","Enduit","Gants, masque poussières",0,150),
        ("POT DE PLATRE SAFEBOARD","Plâtre","Poudre","Sac / seau","Traitement des joints et finition des plaques de plâtre","NON","SGH07 possible","Humidité excessive","Plâtre","Gants, masque poussières",0,50),
        ("DILUAN DJESSI","Solvant","Liquide","Bidon","Dilution des peintures et nettoyage du matériel","NON","SGH02/SGH07 possible","Sources d’ignition, oxydants","Diluant","Gants, lunettes, ventilation",0,100),
        ("LAC PRO","Vernis","Liquide","Bidon / seau","Vernis et protection des surfaces peintes","NON","SGH02/SGH07 possible","Sources d’ignition, oxydants","Vernis","Gants, lunettes",0,50),
        ("PEINTURE DJESSI","Peinture","Liquide","Seau peinture","Peinture intérieure et extérieure pour bâtiment","NON","SGH07 possible","Oxydants forts","Peinture","Gants, lunettes",0,150),
        ("SIKA MORTIER","Mortier","Poudre","Sac 25 kg","Réparation, scellement et reprofilage du béton","NON","SGH05/SGH07 possible","Acides forts","Mortier cimentaire","Gants, lunettes, masque poussières",0,80),
        ("SIKALASTIC","Étanchéité","Liquide","Seau ou kit","Membrane liquide d’étanchéité","NON","SGH07 possible","Oxydants forts","Étanchéité liquide","Gants, lunettes",0,100),
        ("ALSAN FLASHING","Étanchéité","Liquide","Seau ou kit","Résine liquide pour relevés et détails d’étanchéité","NON","SGH02/SGH07 possible","Sources d’ignition, oxydants","Résine d’étanchéité","Gants, lunettes, ventilation",0,60),
        ("ENDUIT EVOLUTION","Enduit","Poudre","Sac 14 ou 30 kg","Enduit de rebouchage et de lissage","NON","SGH07 possible","Acides forts","Enduit","Gants, masque poussières",0,150),
        ("IMPRIM LUX COLORIS","Primaire","Liquide","Seau","Sous-couche et primaire avant peinture","NON","SGH07 possible","Oxydants forts","Primaire peinture","Gants, lunettes",0,150),
        ("GARNYTEX","Revêtement","Liquide","Seau","Revêtement décoratif et protection des surfaces","NON","SGH07 possible","Oxydants forts","Revêtement décoratif","Gants, lunettes",0,100),
        ("COLLE SPM ACRYLIQUE","Colle","Pâte","Seau / cartouche","Collage de revêtements et matériaux de finition","NON","SGH07 possible","Oxydants forts","Colle acrylique","Gants, lunettes",0,100),
        ("PANTIPRIM","Primaire","Liquide","Seau plastique","Couche d’impression avant peinture ou enduit","NON","SGH07 possible","Oxydants forts","Primaire","Gants, lunettes",0,100),
        ("COLORIS SENIOR","Peinture","Liquide","Seau peinture","Peinture de finition intérieure/extérieure","NON","SGH07 possible","Oxydants forts","Peinture","Gants, lunettes",0,100),
        ("BELLE & REBELLE","Peinture","Liquide","Seau peinture","Peinture décorative de finition","NON","SGH07 possible","Oxydants forts","Peinture décorative","Gants, lunettes",0,80),
        ("SIKA CEM HYDROFUGE LIQUIDE","Adjuvant","Liquide","Bidon","Adjuvant hydrofuge pour mortiers et bétons","NON","SGH07 possible","Oxydants forts","Adjuvant hydrofuge","Gants, lunettes",0,50),
        ("SOLUTION ENDUIT CIMENT","Enduit","Poudre","Sac 25 kg","Enduit ciment pour maçonnerie et façades","NON","SGH05/SGH07 possible","Acides forts","Enduit cimentaire","Gants, lunettes, masque poussières",0,200),
    ]

    for p in sample_products:
        cur.execute("""INSERT OR IGNORE INTO products(root_name,family,physical_state,conditionnement,utilisation,fds,pictogrammes,incompatibilites,danger_class,epi,stock_min,stock_max)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""", p)

    conn.commit()
    conn.close()

def log_action(action, details=""):
    try:
        conn = db()
        conn.execute("INSERT INTO audit_log(action_date,username,action,details) VALUES(?,?,?,?)",
                     (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), session.get("username","system"), action, details))
        conn.commit()
        conn.close()
    except Exception:
        pass

# ---------- Auth / roles ----------
def role(): return session.get("role")
def is_super_admin(): return role() == "super_admin"
def is_project_admin(): return role() == "project_admin"
def is_subcontractor(): return role() == "subcontractor"
def is_reader(): return role() == "reader"
def current_project_id(): return session.get("project_id")
def current_subcontractor_id(): return session.get("subcontractor_id")
def can_access_admin_features(): return is_super_admin() or is_project_admin()
def logged(): return session.get("user_id") is not None

def can_manage_entry(entry_row):
    if is_super_admin():
        return True
    if is_project_admin():
        return str(entry_row["project_id"]) == str(current_project_id())
    if is_subcontractor():
        return str(entry_row["subcontractor_id"]) == str(current_subcontractor_id())
    return False


def can_manage_project(project_row):
    if is_super_admin():
        return True
    if is_project_admin():
        return str(project_row["id"]) == str(current_project_id())
    return False

def can_delete_project(project_row):
    return is_super_admin()

def can_manage_subcontractor(st_row):
    if is_super_admin():
        return True
    if is_project_admin():
        return True
    if is_subcontractor():
        return str(st_row["id"]) == str(current_subcontractor_id())
    return False

def can_delete_subcontractor(st_row):
    return is_super_admin() or is_project_admin()

def can_manage_user(user_row):
    if is_super_admin():
        return True
    if is_project_admin():
        return str(user_row["project_id"]) == str(current_project_id()) and user_row["role"] in ["subcontractor", "reader"]
    return False

def can_delete_user(user_row):
    if str(user_row["id"]) == str(session.get("user_id")):
        return False
    return can_manage_user(user_row)

def yes_no_options(value):
    value = value or "NON"
    return "".join([f"<option {'selected' if value==v else ''}>{v}</option>" for v in ["OUI","NON"]])

def status_options(value):
    value = value or "Actif"
    return "".join([f"<option {'selected' if value==v else ''}>{v}</option>" for v in ["Actif","Inactif","Clôturé"]])

def safe_int(value, default=0):
    try:
        return int(value or default)
    except Exception:
        return default

def safe_float(value, default=0):
    try:
        return float(value or default)
    except Exception:
        return default


def where_scope(prefix="e"):
    if is_super_admin():
        return "", []
    if is_project_admin() or is_reader():
        return f" WHERE {prefix}.project_id=? ", [current_project_id()]
    if is_subcontractor():
        return f" WHERE {prefix}.subcontractor_id=? ", [current_subcontractor_id()]
    return " WHERE 1=0 ", []

def scoped_and_clause():
    where, params = where_scope("e")
    return (where + (" AND " if where else " WHERE "), params)

# ---------- Product intelligence ----------
def normalize(text):
    text = (text or "").upper()
    for ch in ["-", "_", ".", ",", "/", "\\", "(", ")", "[", "]", "&"]:
        text = text.replace(ch, " ")
    remove = {"KG","KGS","L","LTR","LITRE","LITRES","SEAU","BIDON","SAC","POT","16","14","20","25","30"}
    return " ".join([w for w in text.split() if w not in remove])

def find_product_id(name):
    n = normalize(name)
    if not n:
        return None
    conn = db()
    row = conn.execute("SELECT product_id FROM aliases WHERE alias_name=?", (n,)).fetchone()
    if row:
        conn.close()
        return row["product_id"]

    best_id, best_score = None, 0
    nw = set(n.split())
    for p in conn.execute("SELECT id,root_name FROM products").fetchall():
        r = normalize(p["root_name"])
        rw = set(r.split())
        if n == r:
            conn.close()
            return p["id"]
        if r in n or n in r:
            score = 90
        else:
            score = int(100 * len(nw.intersection(rw)) / max(len(rw), 1))
        if score > best_score:
            best_score, best_id = score, p["id"]
    conn.close()
    return best_id if best_score >= 60 else None

def days_to(date_text):
    try:
        d = datetime.strptime(date_text, "%Y-%m-%d").date()
        return (d - date.today()).days
    except Exception:
        return None

# ---------- UI ----------
def layout(title, body):
    nav = ""
    if logged():
        admin_links = ""
        if can_access_admin_features():
            admin_links = f"""
            <a href="{url_for('projects')}">Projets</a>
            <a href="{url_for('users')}">Utilisateurs</a>
            <a href="{url_for('subcontractors')}">Sous-traitants</a>
            <a href="{url_for('products')}">Base produits</a>
            <a href="{url_for('alerts')}">Alertes</a>
            <a href="{url_for('aliases')}">Alias/Doublons</a>
            <a href="{url_for('consolidation')}">Consolidation</a>
            <a href="{url_for('export_excel')}">Export Excel</a>
            <a href="{url_for('audit_trail')}">Journal</a>
            """
        if is_reader():
            admin_links = f"""<a href="{url_for('consolidation')}">Consolidation</a><a href="{url_for('alerts')}">Alertes</a>"""
        nav = f"""
        <nav>
            <a href="{url_for('dashboard')}">Tableau de bord</a>
            {admin_links}
            <a href="{url_for('entries')}">Saisie</a>
            <a href="{url_for('about')}">À propos</a>
            <a href="{url_for('logout')}">Déconnexion</a>
        </nav>"""
    flashes = "".join([f"<div class='flash'>{m}</div>" for m in get_flashed_messages()])
    userlabel = session.get("username","")
    if role(): userlabel += f" | {role()}"
    return render_template_string(f"""<!doctype html>
    <html lang="fr"><head><meta charset="utf-8"><title>{title}</title>{CSS}</head>
    <body>
    <header><div><h2>{APP_NAME}</h2><small>{VERSION} — {COPYRIGHT}</small></div><small>{userlabel}</small></header>
    {nav}<main>{flashes}{body}<div class="footer">© 2026 Kodjotse Eli ADIGBLI – QHSE Manager Pro</div></main>
    </body></html>""")

# ---------- Data queries ----------
def consolidation_rows():
    conn = db()
    where, params = where_scope("e")
    rows = conn.execute(f"""SELECT COALESCE(p.root_name,'PRODUIT NON RECONNU: '||e.declared_product) product_name,
    p.family,p.conditionnement,p.utilisation,p.fds,p.fds_expiry,p.pictogrammes,p.incompatibilites,
    p.danger_class,p.epi,p.stock_min,p.stock_max,e.unit,
    SUM(e.qty_in) total_in,SUM(e.qty_used) total_used,SUM(e.qty_stock) total_stock,
    COUNT(DISTINCT e.subcontractor_id) nb_st,
    GROUP_CONCAT(DISTINCT s.name) sts,
    GROUP_CONCAT(DISTINCT pr.name) projects,
    GROUP_CONCAT(DISTINCT e.storage_location) locations
    FROM entries e 
    LEFT JOIN products p ON p.id=e.product_id 
    JOIN subcontractors s ON s.id=e.subcontractor_id 
    LEFT JOIN projects pr ON pr.id=e.project_id
    {where}
    GROUP BY product_name,e.unit 
    ORDER BY product_name""", params).fetchall()
    conn.close()
    return rows

def get_alert_rows():
    alerts = []
    for r in consolidation_rows():
        product = r["product_name"]
        if product.startswith("PRODUIT NON RECONNU"):
            alerts.append(("CRITIQUE","Produit non reconnu",product,"Créer un produit racine ou un alias."))
        if (r["fds"] or "NON") == "NON":
            alerts.append(("CRITIQUE","FDS manquante",product,"Demander la FDS et compléter la base produit."))
        d = days_to(r["fds_expiry"])
        if d is not None and d < 0:
            alerts.append(("CRITIQUE","FDS expirée",product,f"FDS expirée depuis {-d} jour(s)."))
        elif d is not None and d <= 30:
            alerts.append(("ATTENTION","FDS proche expiration",product,f"FDS expire dans {d} jour(s)."))
        stock = r["total_stock"] or 0
        if (r["stock_min"] or 0) > 0 and stock < r["stock_min"]:
            alerts.append(("ATTENTION","Stock critique",product,f"Stock {stock} inférieur au minimum {r['stock_min']}."))
        if (r["stock_max"] or 0) > 0 and stock > r["stock_max"]:
            alerts.append(("CRITIQUE","Stock dépassé",product,f"Stock {stock} supérieur au maximum {r['stock_max']}."))
    return alerts

# ---------- Routes ----------
@app.route("/")
def home():
    return redirect(url_for("dashboard") if logged() else url_for("login"))

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        conn = db()
        user = conn.execute("SELECT * FROM users WHERE username=? AND status='Actif'", (request.form["username"],)).fetchone()
        conn.close()
        if user and check_password_hash(user["password_hash"], request.form["password"]):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            session["subcontractor_id"] = user["subcontractor_id"]
            session["project_id"] = user["project_id"]
            log_action("Connexion", "Connexion utilisateur")
            return redirect(url_for("dashboard"))
        flash("Identifiants incorrects ou compte inactif.")
    return layout("Connexion", """<h2>Connexion</h2><div class="box"><form method="post" autocomplete="off">
        <label>Utilisateur</label><input name="username" autocomplete="off" required autofocus placeholder="Saisir votre identifiant">
        <label>Mot de passe</label><input name="password" type="password" autocomplete="new-password" required placeholder="Saisir votre mot de passe">
        <button>Se connecter</button></form></div>""")

@app.route("/logout")
def logout():
    log_action("Déconnexion", "Déconnexion utilisateur")
    session.clear()
    response = redirect(url_for("login"))
    response.delete_cookie(app.config.get("SESSION_COOKIE_NAME", "session"))
    return response

@app.route("/dashboard")
def dashboard():
    if not logged(): return redirect(url_for("login"))
    conn = db()
    where, params = where_scope("e")
    and_clause, and_params = scoped_and_clause()

    if is_super_admin():
        total_projects = conn.execute("SELECT COUNT(*) c FROM projects").fetchone()["c"]
        total_st = conn.execute("SELECT COUNT(*) c FROM subcontractors").fetchone()["c"]
        total_products = conn.execute("SELECT COUNT(*) c FROM products").fetchone()["c"]
        help_text = "Vue Super Admin : tous les projets, tous les sous-traitants, consolidation générale."
    elif is_project_admin():
        total_projects = 1
        total_st = conn.execute("SELECT COUNT(DISTINCT subcontractor_id) c FROM entries WHERE project_id=?", (current_project_id(),)).fetchone()["c"]
        total_products = conn.execute("SELECT COUNT(DISTINCT product_id) c FROM entries WHERE project_id=? AND product_id IS NOT NULL", (current_project_id(),)).fetchone()["c"]
        help_text = "Vue Admin Projet : tu gères uniquement le projet qui t'est attribué."
    elif is_reader():
        total_projects = 1
        total_st = conn.execute("SELECT COUNT(DISTINCT subcontractor_id) c FROM entries WHERE project_id=?", (current_project_id(),)).fetchone()["c"]
        total_products = conn.execute("SELECT COUNT(DISTINCT product_id) c FROM entries WHERE project_id=? AND product_id IS NOT NULL", (current_project_id(),)).fetchone()["c"]
        help_text = "Vue Lecteur : consultation limitée aux rapports autorisés."
    else:
        total_projects = 1
        total_st = 1
        total_products = conn.execute("SELECT COUNT(DISTINCT declared_product) c FROM entries WHERE subcontractor_id=?", (current_subcontractor_id(),)).fetchone()["c"]
        help_text = "Vue Sous-traitant : tu saisis uniquement tes propres produits chimiques."

    total_stock = conn.execute(f"SELECT COALESCE(SUM(qty_stock),0) s FROM entries e {where}", params).fetchone()["s"]
    total_in = conn.execute(f"SELECT COALESCE(SUM(qty_in),0) s FROM entries e {where}", params).fetchone()["s"]
    total_used = conn.execute(f"SELECT COALESCE(SUM(qty_used),0) s FROM entries e {where}", params).fetchone()["s"]
    unknown = conn.execute(f"SELECT COUNT(*) c FROM entries e {and_clause} e.product_id IS NULL", and_params).fetchone()["c"]
    no_fds = conn.execute(f"SELECT COUNT(*) c FROM entries e {and_clause} e.fds_available='NON'", and_params).fetchone()["c"]

    alert_rows = get_alert_rows()
    alerts_count = len(alert_rows)
    critical_alerts = sum(1 for a in alert_rows if a[0] == "CRITIQUE")
    warning_alerts = sum(1 for a in alert_rows if a[0] == "ATTENTION")

    top_products = conn.execute(f"""SELECT COALESCE(p.root_name,e.declared_product) product, SUM(e.qty_stock) stock
                          FROM entries e LEFT JOIN products p ON p.id=e.product_id
                          {where}
                          GROUP BY product ORDER BY stock DESC LIMIT 10""", params).fetchall()

    top_used = conn.execute(f"""SELECT COALESCE(p.root_name,e.declared_product) product, SUM(e.qty_used) used_qty
                          FROM entries e LEFT JOIN products p ON p.id=e.product_id
                          {where}
                          GROUP BY product ORDER BY used_qty DESC LIMIT 10""", params).fetchall()

    by_subcontractor = conn.execute(f"""SELECT s.name st, SUM(e.qty_stock) stock, COUNT(DISTINCT COALESCE(e.product_id, e.declared_product)) products_count
                          FROM entries e JOIN subcontractors s ON s.id=e.subcontractor_id
                          {where}
                          GROUP BY s.name ORDER BY stock DESC LIMIT 10""", params).fetchall()

    by_family = conn.execute(f"""SELECT COALESCE(p.family,'Non classé') family, COUNT(DISTINCT COALESCE(e.product_id, e.declared_product)) count_products, SUM(e.qty_stock) stock
                          FROM entries e LEFT JOIN products p ON p.id=e.product_id
                          {where}
                          GROUP BY family ORDER BY stock DESC LIMIT 10""", params).fetchall()

    by_month = conn.execute(f"""SELECT e.month, SUM(e.qty_in) qty_in, SUM(e.qty_used) qty_used, SUM(e.qty_stock) stock
                          FROM entries e
                          {where}
                          GROUP BY e.month ORDER BY e.month DESC LIMIT 6""", params).fetchall()

    conn.close()

    top_products_rows = "".join([f"<tr><td>{r['product']}</td><td>{r['stock']}</td></tr>" for r in top_products]) or "<tr><td colspan='2'>Aucune donnée</td></tr>"
    top_used_rows = "".join([f"<tr><td>{r['product']}</td><td>{r['used_qty']}</td></tr>" for r in top_used]) or "<tr><td colspan='2'>Aucune donnée</td></tr>"
    st_rows = "".join([f"<tr><td>{r['st']}</td><td>{r['products_count']}</td><td>{r['stock']}</td></tr>" for r in by_subcontractor]) or "<tr><td colspan='3'>Aucune donnée</td></tr>"
    family_rows = "".join([f"<tr><td>{r['family']}</td><td>{r['count_products']}</td><td>{r['stock']}</td></tr>" for r in by_family]) or "<tr><td colspan='3'>Aucune donnée</td></tr>"
    month_rows = "".join([f"<tr><td>{r['month']}</td><td>{r['qty_in']}</td><td>{r['qty_used']}</td><td>{r['stock']}</td></tr>" for r in by_month]) or "<tr><td colspan='4'>Aucune donnée</td></tr>"

    return layout("Tableau de bord", f"""
    <h2>Tableau de bord</h2><div class="help">{help_text}</div>
    <div class="grid">
        <div class="card">Projets<div class="value">{total_projects}</div></div>
        <div class="card">Sous-traitants<div class="value">{total_st}</div></div>
        <div class="card">Produits<div class="value">{total_products}</div></div>
        <div class="card">Stock total<div class="value">{total_stock}</div></div>
    </div><br>
    <div class="grid">
        <div class="card">Entrées totales<div class="value ok">{total_in}</div></div>
        <div class="card">Utilisées totales<div class="value warn">{total_used}</div></div>
        <div class="card">Produits non reconnus<div class="value danger">{unknown}</div></div>
        <div class="card">Saisies sans FDS<div class="value danger">{no_fds}</div></div>
    </div><br>
    <div class="grid">
        <div class="card">Alertes QHSE<div class="value warn">{alerts_count}</div></div>
        <div class="card">Alertes critiques<div class="value danger">{critical_alerts}</div></div>
        <div class="card">Alertes attention<div class="value warn">{warning_alerts}</div></div>
        <div class="card">Version<div class="value ok">2.1</div></div>
    </div>

    <h3>Analyse rapide</h3>
    <div class="row">
        <div>
            <h3>Top 10 produits stockés</h3>
            <table><tr><th>Produit</th><th>Stock</th></tr>{top_products_rows}</table>
        </div>
        <div>
            <h3>Top 10 produits utilisés</h3>
            <table><tr><th>Produit</th><th>Quantité utilisée</th></tr>{top_used_rows}</table>
        </div>
    </div>

    <div class="row">
        <div>
            <h3>Stock par sous-traitant</h3>
            <table><tr><th>Sous-traitant</th><th>Nb produits</th><th>Stock</th></tr>{st_rows}</table>
        </div>
        <div>
            <h3>Répartition par famille</h3>
            <table><tr><th>Famille</th><th>Nb produits</th><th>Stock</th></tr>{family_rows}</table>
        </div>
    </div>

    <h3>Évolution mensuelle</h3>
    <table><tr><th>Mois</th><th>Entrées</th><th>Utilisées</th><th>Stock</th></tr>{month_rows}</table>
    """)

@app.route("/projects", methods=["GET","POST"])
def projects():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        if is_project_admin():
            flash("Un Admin Projet peut modifier son projet depuis la liste, mais ne peut pas créer un nouveau projet.")
            conn.close(); return redirect(url_for("projects"))
        name = request.form["name"].strip()
        conn.execute("""INSERT OR IGNORE INTO projects(name,location,country,client,owner,main_contractor,project_manager,qhse_manager,coordinator,start_date,end_date,status,description)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (name, request.form.get("location",""), request.form.get("country",""), request.form.get("client",""),
                      request.form.get("owner",""), request.form.get("main_contractor",""), request.form.get("project_manager",""),
                      request.form.get("qhse_manager",""), request.form.get("coordinator",""), request.form.get("start_date",""),
                      request.form.get("end_date",""), request.form.get("status","Actif"), request.form.get("description","")))
        conn.commit()
        log_action("Création projet", name)
        flash("Projet enregistré.")
    if is_super_admin():
        rows = conn.execute("SELECT * FROM projects ORDER BY name").fetchall()
    else:
        rows = conn.execute("SELECT * FROM projects WHERE id=?", (current_project_id(),)).fetchall()
    conn.close()
    trs = ""
    for r in rows:
        edit_btn = f"<a class='btn mini' href='{url_for('edit_project', project_id=r['id'])}'>Modifier</a>" if can_manage_project(r) else ""
        del_btn = f"<form method='post' action='{url_for('delete_project', project_id=r['id'])}' style='display:inline' onsubmit='return confirm(&quot;Supprimer ce projet ? Cette action est définitive.&quot;)'><button class='mini btn-danger'>Supprimer</button></form>" if can_delete_project(r) else ""
        trs += f"<tr><td>{r['name']}</td><td>{r['client'] or ''}</td><td>{r['main_contractor'] or ''}</td><td>{r['qhse_manager'] or ''}</td><td>{r['location'] or ''}</td><td>{r['country'] or ''}</td><td>{r['status']}</td><td class='actions'>{edit_btn} {del_btn}</td></tr>"
    add_form = "" if is_project_admin() else f"""<div class="box"><form method="post">
    <div class="row3"><div><label>Nom du projet</label><input name="name" required placeholder="Ex : CHU Kara"></div><div><label>Ville</label><input name="location" placeholder="Ex : Kara"></div><div><label>Pays</label><input name="country" value="Togo" placeholder="Ex : Togo"></div></div>
    <div class="row3"><div><label>Client</label><input name="client" placeholder="Nom du client"></div><div><label>Maître d'ouvrage</label><input name="owner" placeholder="Institution / propriétaire"></div><div><label>Entreprise principale</label><input name="main_contractor" value="Ellipse Projects" placeholder="Entreprise principale"></div></div>
    <div class="row3"><div><label>Chef de projet</label><input name="project_manager" placeholder="Nom du chef de projet"></div><div><label>Responsable QHSE</label><input name="qhse_manager" placeholder="Nom du responsable QHSE"></div><div><label>Coordinateur</label><input name="coordinator" placeholder="Nom du coordinateur"></div></div>
    <div class="row3"><div><label>Date début</label><input type="date" name="start_date"></div><div><label>Date fin</label><input type="date" name="end_date"></div><div><label>Statut</label><select name="status"><option>Actif</option><option>Clôturé</option><option>Inactif</option></select></div></div>
    <label>Description</label><textarea name="description" placeholder="Résumé du projet, lots concernés, observations importantes"></textarea><button>Ajouter</button></form></div>"""
    return layout("Projets", f"""<h2>Projets</h2><div class="help">Chaque ligne peut être modifiée ou supprimée selon ton niveau de droit. Le Super Admin gère tous les projets ; l'Admin Projet ne voit que son projet.</div>{add_form}
    <table><tr><th>Projet</th><th>Client</th><th>Entreprise principale</th><th>Resp. QHSE</th><th>Ville</th><th>Pays</th><th>Statut</th><th>Actions</th></tr>{trs}</table>""")

@app.route("/projects/edit/<int:project_id>", methods=["GET","POST"])
def edit_project(project_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db(); r = conn.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
    if not r: conn.close(); flash("Projet introuvable."); return redirect(url_for("projects"))
    if not can_manage_project(r): conn.close(); flash("Accès refusé."); return redirect(url_for("projects"))
    if request.method == "POST":
        conn.execute("""UPDATE projects SET name=?,location=?,country=?,client=?,owner=?,main_contractor=?,project_manager=?,qhse_manager=?,coordinator=?,start_date=?,end_date=?,status=?,description=? WHERE id=?""",
                     (request.form["name"].strip(), request.form.get("location",""), request.form.get("country",""), request.form.get("client",""), request.form.get("owner",""), request.form.get("main_contractor",""), request.form.get("project_manager",""), request.form.get("qhse_manager",""), request.form.get("coordinator",""), request.form.get("start_date",""), request.form.get("end_date",""), request.form.get("status","Actif"), request.form.get("description",""), project_id))
        conn.commit(); conn.close(); log_action("Modification projet", request.form["name"]); flash("Projet modifié."); return redirect(url_for("projects"))
    status = status_options(r['status'])
    body=f"""<h2>Modifier le projet</h2><div class="box"><form method="post">
    <div class="row3"><div><label>Nom du projet</label><input name="name" value="{escape(r['name'])}" required placeholder="Ex : CHU Kara"></div><div><label>Ville</label><input name="location" value="{escape(r['location'] or '')}" placeholder="Ex : Kara"></div><div><label>Pays</label><input name="country" value="{escape(r['country'] or '')}" placeholder="Ex : Togo"></div></div>
    <div class="row3"><div><label>Client</label><input name="client" value="{escape(r['client'] or '')}" placeholder="Nom du client"></div><div><label>Maître d'ouvrage</label><input name="owner" value="{escape(r['owner'] or '')}" placeholder="Institution / propriétaire"></div><div><label>Entreprise principale</label><input name="main_contractor" value="{escape(r['main_contractor'] or '')}" placeholder="Entreprise principale"></div></div>
    <div class="row3"><div><label>Chef de projet</label><input name="project_manager" value="{escape(r['project_manager'] or '')}" placeholder="Nom du chef de projet"></div><div><label>Responsable QHSE</label><input name="qhse_manager" value="{escape(r['qhse_manager'] or '')}" placeholder="Nom du responsable QHSE"></div><div><label>Coordinateur</label><input name="coordinator" value="{escape(r['coordinator'] or '')}" placeholder="Nom du coordinateur"></div></div>
    <div class="row3"><div><label>Date début</label><input type="date" name="start_date" value="{r['start_date'] or ''}"></div><div><label>Date fin</label><input type="date" name="end_date" value="{r['end_date'] or ''}"></div><div><label>Statut</label><select name="status">{status}</select></div></div>
    <label>Description</label><textarea name="description" placeholder="Résumé du projet, lots concernés, observations importantes">{escape(r['description'] or '')}</textarea>
    <button>Enregistrer</button> <a class="btn" href="{url_for('projects')}">Annuler</a></form></div>"""
    conn.close(); return layout("Modifier projet", body)

@app.route("/projects/delete/<int:project_id>", methods=["POST"])
def delete_project(project_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db(); r = conn.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
    if not r or not can_delete_project(r): conn.close(); flash("Suppression non autorisée."); return redirect(url_for("projects"))
    if conn.execute("SELECT COUNT(*) c FROM entries WHERE project_id=?", (project_id,)).fetchone()["c"] > 0:
        conn.close(); flash("Impossible de supprimer : ce projet contient déjà des saisies. Passe-le plutôt en Inactif/Clôturé."); return redirect(url_for("projects"))
    conn.execute("DELETE FROM projects WHERE id=?", (project_id,)); conn.commit(); conn.close(); log_action("Suppression projet", r['name']); flash("Projet supprimé."); return redirect(url_for("projects"))

@app.route("/users", methods=["GET","POST"])
def users():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()
        r = request.form["role"]
        project_id = request.form.get("project_id") or None
        subcontractor_id = request.form.get("subcontractor_id") or None
        if not username or not password:
            flash("Nom d'utilisateur et mot de passe obligatoires."); conn.close(); return redirect(url_for("users"))
        if is_project_admin():
            if r not in ["subcontractor","reader"]:
                flash("Un Admin Projet ne peut créer que des comptes Sous-traitant ou Lecteur."); conn.close(); return redirect(url_for("users"))
            project_id = current_project_id()
        if r == "super_admin":
            project_id = None; subcontractor_id = None
        if r in ["project_admin","reader"]:
            subcontractor_id = None
        if r == "project_admin" and not project_id:
            flash("Choisis un projet pour l'Admin Projet."); conn.close(); return redirect(url_for("users"))
        if r == "subcontractor" and not subcontractor_id:
            flash("Choisis un sous-traitant associé."); conn.close(); return redirect(url_for("users"))
        try:
            conn.execute("INSERT INTO users(username,password_hash,role,subcontractor_id,project_id,status) VALUES(?,?,?,?,?,?)",
                         (username, generate_password_hash(password), r, subcontractor_id, project_id, request.form.get("status","Actif")))
            conn.commit(); log_action("Création utilisateur", username); flash("Utilisateur créé.")
        except sqlite3.IntegrityError:
            flash("Ce nom d'utilisateur existe déjà.")
    if is_super_admin():
        projects_rows = conn.execute("SELECT * FROM projects ORDER BY name").fetchall()
        users_rows = conn.execute("""SELECT u.*, s.name st_name, p.name project_name FROM users u 
                                     LEFT JOIN subcontractors s ON s.id=u.subcontractor_id
                                     LEFT JOIN projects p ON p.id=u.project_id ORDER BY u.username""").fetchall()
    else:
        projects_rows = conn.execute("SELECT * FROM projects WHERE id=?", (current_project_id(),)).fetchall()
        users_rows = conn.execute("""SELECT u.*, s.name st_name, p.name project_name FROM users u 
                                     LEFT JOIN subcontractors s ON s.id=u.subcontractor_id
                                     LEFT JOIN projects p ON p.id=u.project_id
                                     WHERE u.project_id=? ORDER BY u.username""", (current_project_id(),)).fetchall()
    sts = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall(); conn.close()
    role_options = "<option value='subcontractor'>Sous-traitant</option><option value='reader'>Lecteur</option>"
    if is_super_admin():
        role_options = "<option value='subcontractor'>Sous-traitant</option><option value='project_admin'>Admin Projet</option><option value='reader'>Lecteur</option><option value='super_admin'>Super Admin</option>"
    project_opts = "".join([f"<option value='{p['id']}'>{p['name']}</option>" for p in projects_rows])
    st_opts = "".join([f"<option value='{st['id']}'>{st['name']}</option>" for st in sts])
    trs = ""
    for u in users_rows:
        edit_btn = f"<a class='btn mini' href='{url_for('edit_user', user_id=u['id'])}'>Modifier</a>" if can_manage_user(u) else ""
        del_btn = f"<form method='post' action='{url_for('delete_user', user_id=u['id'])}' style='display:inline' onsubmit='return confirm(&quot;Supprimer cet utilisateur ?&quot;)'><button class='mini btn-danger'>Supprimer</button></form>" if can_delete_user(u) else ""
        trs += f"<tr><td>{u['username']}</td><td>{u['role']}</td><td>{u['project_name'] or 'Tous / Aucun'}</td><td>{u['st_name'] or '-'}</td><td>{u['status']}</td><td class='actions'>{edit_btn} {del_btn}</td></tr>"
    return layout("Utilisateurs", f"""<h2>Utilisateurs et droits</h2>
    <div class="help"><strong>Correction intégrée :</strong> la création d'utilisateur vérifie maintenant le rôle, le projet et le sous-traitant. Les boutons Modifier/Supprimer apparaissent selon les droits.</div>
    <div class="box"><form method="post">
    <div class="row"><div><label>Nom d'utilisateur</label><input name="username" required placeholder="Ex : sicone_user ou admin_chu_kara"></div><div><label>Mot de passe provisoire</label><input name="password" type="password" required placeholder="Mot de passe temporaire à communiquer à l'utilisateur"></div></div>
    <div class="row"><div><label>Rôle</label><select name="role">{role_options}</select></div><div><label>Projet autorisé</label><select name="project_id"><option value="">Tous / Aucun</option>{project_opts}</select></div></div>
    <div class="row"><div><label>Sous-traitant associé</label><select name="subcontractor_id"><option value="">-- aucun --</option>{st_opts}</select></div><div><label>Statut</label><select name="status"><option>Actif</option><option>Inactif</option></select></div></div>
    <button>Créer l'utilisateur</button></form></div>
    <table><tr><th>Utilisateur</th><th>Rôle</th><th>Projet</th><th>Sous-traitant associé</th><th>Statut</th><th>Actions</th></tr>{trs}</table>""")

@app.route("/users/edit/<int:user_id>", methods=["GET","POST"])
def edit_user(user_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db(); u = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    if not u or not can_manage_user(u): conn.close(); flash("Utilisateur introuvable ou accès refusé."); return redirect(url_for("users"))
    if request.method == "POST":
        r = request.form["role"]; project_id = request.form.get("project_id") or None; subcontractor_id = request.form.get("subcontractor_id") or None
        if is_project_admin():
            if r not in ["subcontractor","reader"]: flash("Rôle non autorisé."); conn.close(); return redirect(url_for("users"))
            project_id = current_project_id()
        if r == "super_admin": project_id = None; subcontractor_id = None
        if r in ["project_admin","reader"]: subcontractor_id = None
        if r == "project_admin" and not project_id: flash("Projet obligatoire pour Admin Projet."); conn.close(); return redirect(url_for("users"))
        if r == "subcontractor" and not subcontractor_id: flash("Sous-traitant obligatoire."); conn.close(); return redirect(url_for("users"))
        if request.form.get("password","").strip():
            conn.execute("UPDATE users SET username=?, password_hash=?, role=?, subcontractor_id=?, project_id=?, status=? WHERE id=?", (request.form["username"].strip(), generate_password_hash(request.form["password"].strip()), r, subcontractor_id, project_id, request.form.get("status","Actif"), user_id))
        else:
            conn.execute("UPDATE users SET username=?, role=?, subcontractor_id=?, project_id=?, status=? WHERE id=?", (request.form["username"].strip(), r, subcontractor_id, project_id, request.form.get("status","Actif"), user_id))
        conn.commit(); conn.close(); log_action("Modification utilisateur", request.form["username"]); flash("Utilisateur modifié."); return redirect(url_for("users"))
    projects_rows = conn.execute("SELECT * FROM projects ORDER BY name").fetchall() if is_super_admin() else conn.execute("SELECT * FROM projects WHERE id=?", (current_project_id(),)).fetchall()
    sts = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall(); conn.close()
    roles = ["subcontractor","reader"] if is_project_admin() else ["subcontractor","project_admin","reader","super_admin"]
    role_options = "".join([f"<option value='{r}' {'selected' if u['role']==r else ''}>{r}</option>" for r in roles])
    project_opts = "".join([f"<option value='{p['id']}' {'selected' if str(u['project_id'])==str(p['id']) else ''}>{p['name']}</option>" for p in projects_rows])
    st_opts = "".join([f"<option value='{st['id']}' {'selected' if str(u['subcontractor_id'])==str(st['id']) else ''}>{st['name']}</option>" for st in sts])
    status = "".join([f"<option {'selected' if u['status']==v else ''}>{v}</option>" for v in ["Actif","Inactif"]])
    return layout("Modifier utilisateur", f"""<h2>Modifier utilisateur</h2><div class="box"><form method="post">
    <div class="row"><div><label>Nom d'utilisateur</label><input name="username" value="{escape(u['username'])}" required placeholder="Identifiant de connexion"></div><div><label>Nouveau mot de passe</label><input name="password" type="password" placeholder="Laisser vide pour conserver l'ancien"></div></div>
    <div class="row"><div><label>Rôle</label><select name="role">{role_options}</select></div><div><label>Projet autorisé</label><select name="project_id"><option value="">Tous / Aucun</option>{project_opts}</select></div></div>
    <div class="row"><div><label>Sous-traitant associé</label><select name="subcontractor_id"><option value="">-- aucun --</option>{st_opts}</select></div><div><label>Statut</label><select name="status">{status}</select></div></div>
    <button>Enregistrer</button> <a class="btn" href="{url_for('users')}">Annuler</a></form></div>""")

@app.route("/users/delete/<int:user_id>", methods=["POST"])
def delete_user(user_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db(); u = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    if not u or not can_delete_user(u): conn.close(); flash("Suppression non autorisée."); return redirect(url_for("users"))
    conn.execute("DELETE FROM users WHERE id=?", (user_id,)); conn.commit(); conn.close(); log_action("Suppression utilisateur", u['username']); flash("Utilisateur supprimé."); return redirect(url_for("users"))

@app.route("/subcontractors", methods=["GET","POST"])
def subcontractors():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        name = request.form["name"].strip().upper()
        conn.execute("""INSERT OR IGNORE INTO subcontractors(name,contact,phone,email,address,lot,work_zone,employees_count,start_date,end_date,hse_plan,insurance,ppsps,status,observations)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (name, request.form.get("contact",""), request.form.get("phone",""), request.form.get("email",""),
                      request.form.get("address",""), request.form.get("lot",""), request.form.get("work_zone",""), safe_int(request.form.get("employees_count")),
                      request.form.get("start_date",""), request.form.get("end_date",""), request.form.get("hse_plan","NON"), request.form.get("insurance","NON"),
                      request.form.get("ppsps","NON"), request.form.get("status","Actif"), request.form.get("observations","")))
        conn.commit(); log_action("Création sous-traitant", name); flash("Sous-traitant enregistré.")
    rows = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall(); conn.close()
    trs = ""
    for r in rows:
        edit_btn = f"<a class='btn mini' href='{url_for('edit_subcontractor', subcontractor_id=r['id'])}'>Modifier</a>" if can_manage_subcontractor(r) else ""
        del_btn = f"<form method='post' action='{url_for('delete_subcontractor', subcontractor_id=r['id'])}' style='display:inline' onsubmit='return confirm(&quot;Supprimer ce sous-traitant ?&quot;)'><button class='mini btn-danger'>Supprimer</button></form>" if can_delete_subcontractor(r) else ""
        trs += f"<tr><td>{r['name']}</td><td>{r['contact'] or ''}</td><td>{r['phone'] or ''}</td><td>{r['email'] or ''}</td><td>{r['lot'] or ''}</td><td>{r['work_zone'] or ''}</td><td>{r['employees_count'] or 0}</td><td>{r['hse_plan']}</td><td>{r['insurance']}</td><td>{r['ppsps']}</td><td>{r['status']}</td><td class='actions'>{edit_btn} {del_btn}</td></tr>"
    return layout("Sous-traitants", f"""<h2>Sous-traitants</h2>
    <div class="help"><strong>PPSPS</strong> : Plan Particulier de Sécurité et de Protection de la Santé. <strong>Assurance</strong> : preuve de couverture de l'entreprise. <strong>Plan HSE</strong> : organisation santé, sécurité et environnement du sous-traitant.</div>
    <div class="box"><form method="post">
    <div class="row3"><div><label>Nom</label><input name="name" required placeholder="Ex : SICONE"></div><div><label>Contact</label><input name="contact" placeholder="Nom du responsable / point focal"></div><div><label>Téléphone</label><input name="phone" placeholder="Ex : +228 xx xx xx xx"></div></div>
    <div class="row3"><div><label>Email</label><input name="email" type="email" placeholder="exemple@societe.com"></div><div><label>Lot</label><input name="lot" placeholder="Ex : électricité, plomberie, peinture"></div><div><label>Zone d'intervention</label><input name="work_zone" placeholder="Ex : bâtiment A, magasin, toiture"></div></div>
    <label>Adresse</label><input name="address" placeholder="Adresse complète du sous-traitant">
    <div class="row3"><div><label>Nombre employés</label><input type="number" name="employees_count" value="0" placeholder="Effectif présent sur site"></div><div><label>Date arrivée</label><input type="date" name="start_date"></div><div><label>Date départ</label><input type="date" name="end_date"></div></div>
    <div class="row3"><div><label>Plan HSE</label><select name="hse_plan"><option>OUI</option><option selected>NON</option></select><div class="muted">Document décrivant l'organisation HSE.</div></div><div><label>Assurance</label><select name="insurance"><option>OUI</option><option selected>NON</option></select><div class="muted">Attestation d'assurance valide.</div></div><div><label>PPSPS</label><select name="ppsps"><option>OUI</option><option selected>NON</option></select><div class="muted">Plan sécurité spécifique aux travaux.</div></div></div>
    <label>Statut</label><select name="status"><option>Actif</option><option>Inactif</option></select><label>Observations</label><textarea name="observations" placeholder="Documents manquants, remarques HSE, restrictions d'accès..."></textarea>
    <button>Ajouter</button></form></div>
    <table><tr><th>Nom</th><th>Contact</th><th>Téléphone</th><th>Email</th><th>Lot</th><th>Zone</th><th>Employés</th><th>Plan HSE</th><th>Assurance</th><th>PPSPS</th><th>Statut</th><th>Actions</th></tr>{trs}</table>""")

@app.route("/subcontractors/edit/<int:subcontractor_id>", methods=["GET","POST"])
def edit_subcontractor(subcontractor_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn=db(); r=conn.execute("SELECT * FROM subcontractors WHERE id=?", (subcontractor_id,)).fetchone()
    if not r or not can_manage_subcontractor(r): conn.close(); flash("Sous-traitant introuvable ou accès refusé."); return redirect(url_for("subcontractors"))
    if request.method=="POST":
        name=request.form["name"].strip().upper()
        conn.execute("""UPDATE subcontractors SET name=?,contact=?,phone=?,email=?,address=?,lot=?,work_zone=?,employees_count=?,start_date=?,end_date=?,hse_plan=?,insurance=?,ppsps=?,status=?,observations=? WHERE id=?""", (name,request.form.get("contact",""),request.form.get("phone",""),request.form.get("email",""),request.form.get("address",""),request.form.get("lot",""),request.form.get("work_zone",""),safe_int(request.form.get("employees_count")),request.form.get("start_date",""),request.form.get("end_date",""),request.form.get("hse_plan","NON"),request.form.get("insurance","NON"),request.form.get("ppsps","NON"),request.form.get("status","Actif"),request.form.get("observations",""),subcontractor_id))
        conn.commit(); conn.close(); log_action("Modification sous-traitant", name); flash("Sous-traitant modifié."); return redirect(url_for("subcontractors"))
    body=f"""<h2>Modifier sous-traitant</h2><div class="box"><form method="post">
    <div class="row3"><div><label>Nom</label><input name="name" value="{escape(r['name'])}" required placeholder="Ex : SICONE"></div><div><label>Contact</label><input name="contact" value="{escape(r['contact'] or '')}" placeholder="Nom du responsable / point focal"></div><div><label>Téléphone</label><input name="phone" value="{escape(r['phone'] or '')}" placeholder="Ex : +228 xx xx xx xx"></div></div>
    <div class="row3"><div><label>Email</label><input name="email" type="email" value="{escape(r['email'] or '')}" placeholder="exemple@societe.com"></div><div><label>Lot</label><input name="lot" value="{escape(r['lot'] or '')}" placeholder="Ex : électricité, plomberie, peinture"></div><div><label>Zone d'intervention</label><input name="work_zone" value="{escape(r['work_zone'] or '')}" placeholder="Ex : bâtiment A, magasin, toiture"></div></div>
    <label>Adresse</label><input name="address" value="{escape(r['address'] or '')}" placeholder="Adresse complète du sous-traitant">
    <div class="row3"><div><label>Nombre employés</label><input type="number" name="employees_count" value="{r['employees_count'] or 0}" placeholder="Effectif présent sur site"></div><div><label>Date arrivée</label><input type="date" name="start_date" value="{r['start_date'] or ''}"></div><div><label>Date départ</label><input type="date" name="end_date" value="{r['end_date'] or ''}"></div></div>
    <div class="row3"><div><label>Plan HSE</label><select name="hse_plan">{yes_no_options(r['hse_plan'])}</select></div><div><label>Assurance</label><select name="insurance">{yes_no_options(r['insurance'])}</select></div><div><label>PPSPS</label><select name="ppsps">{yes_no_options(r['ppsps'])}</select></div></div>
    <label>Statut</label><select name="status">{status_options(r['status'])}</select><label>Observations</label><textarea name="observations" placeholder="Documents manquants, remarques HSE, restrictions d'accès...">{escape(r['observations'] or '')}</textarea>
    <button>Enregistrer</button> <a class="btn" href="{url_for('subcontractors')}">Annuler</a></form></div>"""
    conn.close(); return layout("Modifier sous-traitant", body)

@app.route("/subcontractors/delete/<int:subcontractor_id>", methods=["POST"])
def delete_subcontractor(subcontractor_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn=db(); r=conn.execute("SELECT * FROM subcontractors WHERE id=?", (subcontractor_id,)).fetchone()
    if not r or not can_delete_subcontractor(r): conn.close(); flash("Suppression non autorisée."); return redirect(url_for("subcontractors"))
    if conn.execute("SELECT COUNT(*) c FROM entries WHERE subcontractor_id=?", (subcontractor_id,)).fetchone()["c"]>0:
        conn.close(); flash("Impossible de supprimer : ce sous-traitant contient déjà des saisies. Passe-le plutôt en Inactif."); return redirect(url_for("subcontractors"))
    conn.execute("DELETE FROM subcontractors WHERE id=?", (subcontractor_id,)); conn.commit(); conn.close(); log_action("Suppression sous-traitant", r['name']); flash("Sous-traitant supprimé."); return redirect(url_for("subcontractors"))

def product_form_html(action_label="Ajouter", r=None):
    r = r or {}
    def v(k, default=""):
        return escape(str(r[k] if k in r.keys() and r[k] is not None else default)) if hasattr(r, 'keys') else default
    fds_val = r['fds'] if hasattr(r,'keys') and 'fds' in r.keys() else 'NON'
    pictos = ["SGH01 Explosif","SGH02 Inflammable","SGH03 Comburant","SGH04 Gaz sous pression","SGH05 Corrosif","SGH06 Toxicité aiguë","SGH07 Irritant / nocif","SGH08 Danger grave pour la santé","SGH09 Danger environnement","SGH02 + SGH07","SGH05 + SGH07"]
    current_picto = r['pictogrammes'] if hasattr(r,'keys') and 'pictogrammes' in r.keys() else ''
    picto_opts = "<option value=''>-- Sélectionner --</option>" + "".join([f"<option {'selected' if current_picto==x else ''}>{x}</option>" for x in pictos])
    return f"""<div class="box"><form method="post">
    <h3>Identification</h3>
    <div class="row3"><div><label>Produit racine</label><input name="root_name" value="{v('root_name')}" required placeholder="Ex : SOPRALENE"></div><div><label>Nom commercial</label><input name="commercial_name" value="{v('commercial_name')}" placeholder="Nom inscrit sur le contenant"></div><div><label>Fabricant</label><input name="manufacturer" value="{v('manufacturer')}" placeholder="Fabricant / fournisseur"></div></div>
    <div class="row3"><div><label>Référence</label><input name="reference" value="{v('reference')}" placeholder="Référence produit, code fournisseur ou CAS"></div><div><label>Famille</label><input name="family" value="{v('family')}" placeholder="Peinture, colle, solvant, étanchéité..."></div><div><label>État physique</label><input name="physical_state" value="{v('physical_state')}" placeholder="Liquide, solide, pâte, poudre, gaz..."></div></div>
    <label>Conditionnement</label><input name="conditionnement" value="{v('conditionnement')}" placeholder="Ex : seau 20 kg, bidon 5 L, sac 25 kg"><label>Utilisation</label><input name="utilisation" value="{v('utilisation')}" placeholder="Ex : peinture de finition, étanchéité toiture...">
    <h3>FDS et sécurité</h3>
    <div class="row3"><div><label>FDS</label><select name="fds">{yes_no_options(fds_val)}</select></div><div><label>Date FDS</label><input type="date" name="fds_date" value="{v('fds_date')}" placeholder="Date d'émission de la FDS"></div><div><label>Expiration FDS</label><input type="date" name="fds_expiry" value="{v('fds_expiry')}" placeholder="Date de fin de validité ou prochaine révision"></div></div>
    <div class="row3"><div><label>Version FDS</label><input name="fds_version" value="{v('fds_version')}" placeholder="Ex : Version 1.0 / Révision 2026"></div><div><label>Pictogrammes</label><select name="pictogrammes">{picto_opts}</select></div><div><label>Classe de danger</label><input name="danger_class" value="{v('danger_class')}" placeholder="Ex : inflammable, corrosif, toxique..."></div></div>
    <label>Mentions H</label><input name="h_statements" value="{v('h_statements')}" placeholder="Ex : H225 – Liquide et vapeurs très inflammables">
    <label>Mentions P</label><input name="p_statements" value="{v('p_statements')}" placeholder="Ex : P280 – Porter des gants de protection">
    <label>Incompatibilités</label><input name="incompatibilites" value="{v('incompatibilites')}" placeholder="Produits à éviter : acides, bases, oxydants...">
    <label>EPI requis</label><input name="epi" value="{v('epi')}" placeholder="Gants, lunettes, masque, combinaison...">
    <label>Premiers secours</label><textarea name="first_aid" placeholder="Mesures à appliquer en cas d'exposition, inhalation, contact peau/yeux...">{v('first_aid')}</textarea>
    <label>Mesures incendie</label><textarea name="fire_measures" placeholder="Moyens d'extinction appropriés/interdits, risques particuliers...">{v('fire_measures')}</textarea>
    <label>Déversement accidentel</label><textarea name="spill_measures" placeholder="Procédure de confinement, récupération, nettoyage et élimination...">{v('spill_measures')}</textarea>
    <label>Règles de stockage</label><textarea name="storage_rules" placeholder="Conditions de stockage, ventilation, séparation des incompatibles...">{v('storage_rules')}</textarea>
    <h3>Seuils de stock</h3><div class="row"><div><label>Stock minimum</label><input type="number" step="0.01" name="stock_min" value="{v('stock_min','0')}" placeholder="Quantité minimale avant alerte"></div><div><label>Stock maximum</label><input type="number" step="0.01" name="stock_max" value="{v('stock_max','0')}" placeholder="Quantité maximale autorisée"></div></div>
    <label>Observations</label><textarea name="observations" placeholder="Remarques complémentaires, restrictions, suivi audit...">{v('observations')}</textarea><button>{action_label}</button></form></div>"""

def save_product_from_form(conn, product_id=None):
    data = (request.form["root_name"].strip().upper(), request.form.get("commercial_name",""), request.form.get("manufacturer",""), request.form.get("reference",""),
         request.form.get("family",""), request.form.get("physical_state",""), request.form.get("conditionnement",""), request.form.get("utilisation",""),
         request.form.get("fds","NON"), request.form.get("fds_date",""), request.form.get("fds_version",""), request.form.get("fds_expiry",""),
         request.form.get("pictogrammes",""), request.form.get("danger_class",""), request.form.get("h_statements",""), request.form.get("p_statements",""),
         request.form.get("incompatibilites",""), request.form.get("epi",""), request.form.get("first_aid",""), request.form.get("fire_measures",""),
         request.form.get("spill_measures",""), request.form.get("storage_rules",""), request.form.get("stock_min") or 0, request.form.get("stock_max") or 0,
         request.form.get("observations",""))
    if product_id:
        conn.execute("""UPDATE products SET root_name=?,commercial_name=?,manufacturer=?,reference=?,family=?,physical_state=?,conditionnement=?,utilisation=?,fds=?,fds_date=?,fds_version=?,fds_expiry=?,pictogrammes=?,danger_class=?,h_statements=?,p_statements=?,incompatibilites=?,epi=?,first_aid=?,fire_measures=?,spill_measures=?,storage_rules=?,stock_min=?,stock_max=?,observations=? WHERE id=?""", data + (product_id,))
    else:
        conn.execute("""INSERT OR IGNORE INTO products(root_name,commercial_name,manufacturer,reference,family,physical_state,conditionnement,utilisation,fds,fds_date,fds_version,fds_expiry,pictogrammes,danger_class,h_statements,p_statements,incompatibilites,epi,first_aid,fire_measures,spill_measures,storage_rules,stock_min,stock_max,observations)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", data)

@app.route("/products", methods=["GET","POST"])
def products():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        save_product_from_form(conn)
        conn.commit(); log_action("Création produit", request.form["root_name"]); flash("Produit enregistré.")
    rows = conn.execute("SELECT * FROM products ORDER BY root_name").fetchall(); conn.close()
    trs = ""
    for r in rows:
        edit_btn = f"<a class='btn mini' href='{url_for('edit_product', product_id=r['id'])}'>Modifier</a>"
        del_btn = f"<form method='post' action='{url_for('delete_product', product_id=r['id'])}' style='display:inline' onsubmit='return confirm(&quot;Supprimer ce produit ?&quot;)'><button class='mini btn-danger'>Supprimer</button></form>"
        trs += f"<tr><td>{r['root_name']}</td><td>{r['family'] or ''}</td><td>{r['conditionnement'] or ''}</td><td>{r['utilisation'] or ''}</td><td>{r['fds']}</td><td>{r['fds_expiry'] or ''}</td><td>{r['pictogrammes'] or ''}</td><td>{r['incompatibilites'] or ''}</td><td>{r['stock_min'] or 0}</td><td>{r['stock_max'] or 0}</td><td class='actions'>{edit_btn} {del_btn}</td></tr>"
    return layout("Base produits", f"""<h2>Base produits chimiques</h2><div class="help">Tous les champs de saisie contiennent désormais des indications. Les produits peuvent être modifiés ou supprimés selon les droits d'administration.</div>{product_form_html()}
    <table><tr><th>Produit</th><th>Famille</th><th>Conditionnement</th><th>Utilisation</th><th>FDS</th><th>Expiration FDS</th><th>Pictogrammes</th><th>Incompatibilités</th><th>Stock min</th><th>Stock max</th><th>Actions</th></tr>{trs}</table>""")

@app.route("/products/edit/<int:product_id>", methods=["GET","POST"])
def edit_product(product_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn=db(); r=conn.execute("SELECT * FROM products WHERE id=?", (product_id,)).fetchone()
    if not r: conn.close(); flash("Produit introuvable."); return redirect(url_for("products"))
    if request.method=="POST":
        save_product_from_form(conn, product_id); conn.commit(); conn.close(); log_action("Modification produit", request.form["root_name"]); flash("Produit modifié."); return redirect(url_for("products"))
    conn.close(); return layout("Modifier produit", f"<h2>Modifier produit chimique</h2>{product_form_html('Enregistrer les modifications', r)}<p><a class='btn' href='{url_for('products')}'>Retour</a></p>")

@app.route("/products/delete/<int:product_id>", methods=["POST"])
def delete_product(product_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn=db(); r=conn.execute("SELECT * FROM products WHERE id=?", (product_id,)).fetchone()
    if not r: conn.close(); flash("Produit introuvable."); return redirect(url_for("products"))
    if conn.execute("SELECT COUNT(*) c FROM entries WHERE product_id=?", (product_id,)).fetchone()["c"]>0:
        conn.close(); flash("Impossible de supprimer : ce produit est utilisé dans des saisies. Modifie-le ou crée un alias."); return redirect(url_for("products"))
    conn.execute("DELETE FROM aliases WHERE product_id=?", (product_id,)); conn.execute("DELETE FROM products WHERE id=?", (product_id,)); conn.commit(); conn.close(); log_action("Suppression produit", r['root_name']); flash("Produit supprimé."); return redirect(url_for("products"))

@app.route("/aliases", methods=["GET","POST"])
def aliases():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        alias_name = normalize(request.form.get("alias_name", ""))
        product_id = request.form.get("product_id")
        if not alias_name or not product_id:
            flash("Alias non enregistré : le nom de l’alias et le produit racine sont obligatoires.")
        else:
            try:
                conn.execute("INSERT INTO aliases(alias_name,product_id) VALUES(?,?)", (alias_name, product_id))
                conn.commit(); log_action("Création alias", alias_name); flash("Alias enregistré.")
            except sqlite3.IntegrityError:
                flash("Cet alias existe déjà. Utilise le bouton Modifier pour le corriger ou le rattacher à un autre produit racine.")
    products = conn.execute("SELECT * FROM products ORDER BY root_name").fetchall()
    aliases_rows = conn.execute("SELECT a.id,a.alias_name,a.product_id,p.root_name FROM aliases a JOIN products p ON p.id=a.product_id ORDER BY a.alias_name").fetchall(); conn.close()
    opts = "".join([f"<option value='{p['id']}'>{p['root_name']}</option>" for p in products])
    trs = ""
    for a in aliases_rows:
        trs += f"""<tr>
            <td>{escape(a['alias_name'])}</td>
            <td>{escape(a['root_name'])}</td>
            <td class='actions'>
                <a class='btn mini' href='{url_for('edit_alias', alias_id=a['id'])}'>Modifier</a>
                <form method='post' action='{url_for('delete_alias', alias_id=a['id'])}' style='display:inline' onsubmit='return confirm(&quot;Supprimer cet alias/doublon ? Cette action retirera uniquement l’alias, pas le produit racine.&quot;)'>
                    <button type='submit' class='mini btn-danger'>Supprimer</button>
                </form>
            </td>
        </tr>"""
    return layout("Alias", f"""<h2>Alias / Doublons</h2>
    <div class="help"><b>Alias / doublon :</b> permet de rattacher une appellation erronée, commerciale ou abrégée à un produit racine. En cas d’erreur de saisie, utiliser <b>Modifier</b>. Si l’alias n’est plus utile, utiliser <b>Supprimer</b>.</div>
    <div class="box"><form method="post"><label>Nom saisi / alias</label><input name="alias_name" required placeholder="Ex : PANTEX VELOUR 20KG, Pantax velour, PANTEX VEL"><label>Produit racine</label><select name="product_id">{opts}</select><button>Créer l'alias</button></form></div>
    <table><tr><th>Alias / doublon saisi</th><th>Produit racine rattaché</th><th>Modifier / Supprimer</th></tr>{trs}</table>""")

@app.route("/aliases/edit/<int:alias_id>", methods=["GET","POST"])
def edit_alias(alias_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn=db(); a=conn.execute("SELECT * FROM aliases WHERE id=?", (alias_id,)).fetchone()
    if not a: conn.close(); flash("Alias introuvable."); return redirect(url_for("aliases"))
    if request.method=="POST":
        alias_name = normalize(request.form.get("alias_name", ""))
        product_id = request.form.get("product_id")
        if not alias_name or not product_id:
            flash("Modification impossible : le nom de l’alias et le produit racine sont obligatoires.")
        else:
            try:
                conn.execute("UPDATE aliases SET alias_name=?, product_id=? WHERE id=?", (alias_name, product_id, alias_id))
                conn.commit(); conn.close(); log_action("Modification alias", alias_name); flash("Alias modifié."); return redirect(url_for("aliases"))
            except sqlite3.IntegrityError:
                flash("Modification impossible : un autre alias porte déjà ce nom.")
    products=conn.execute("SELECT * FROM products ORDER BY root_name").fetchall(); conn.close()
    opts="".join([f"<option value='{p['id']}' {'selected' if str(a['product_id'])==str(p['id']) else ''}>{p['root_name']}</option>" for p in products])
    return layout("Modifier alias", f"""<h2>Modifier alias / doublon</h2><div class="help">Corrige ici une erreur de saisie ou rattache l’alias à un autre produit racine.</div><div class="box"><form method="post"><label>Nom saisi / alias</label><input name="alias_name" value="{escape(a['alias_name'])}" required placeholder="Ex : PANTEX VELOUR 20KG"><label>Produit racine</label><select name="product_id">{opts}</select><button>Enregistrer</button> <a class="btn" href="{url_for('aliases')}">Annuler</a></form></div>""")

@app.route("/aliases/delete/<int:alias_id>", methods=["POST"])
def delete_alias(alias_id):
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn=db(); a=conn.execute("SELECT * FROM aliases WHERE id=?", (alias_id,)).fetchone()
    if not a: conn.close(); flash("Alias introuvable."); return redirect(url_for("aliases"))
    conn.execute("DELETE FROM aliases WHERE id=?", (alias_id,)); conn.commit(); conn.close(); log_action("Suppression alias", a['alias_name']); flash("Alias supprimé. Le produit racine reste conservé."); return redirect(url_for("aliases"))

@app.route("/entries", methods=["GET","POST"])
def entries():
    if not logged() or is_reader(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        qin = float(request.form.get("qty_in") or 0)
        qused = float(request.form.get("qty_used") or 0)
        qstock = float(request.form.get("qty_stock") or (qin - qused))
        selected_product_id = request.form.get("product_id")
        declared_new = request.form.get("declared_product_new","").strip()

        if selected_product_id:
            product_id = int(selected_product_id)
            p = conn.execute("SELECT root_name,fds FROM products WHERE id=?", (product_id,)).fetchone()
            declared_product = p["root_name"] if p else declared_new
            default_fds = p["fds"] if p else "NON"
        else:
            declared_product = declared_new
            product_id = find_product_id(declared_product)
            p = conn.execute("SELECT fds FROM products WHERE id=?", (product_id,)).fetchone() if product_id else None
            default_fds = p["fds"] if p else "NON"
        fds_available = request.form.get("fds_available") or default_fds

        if can_access_admin_features():
            subcontractor_id = request.form["subcontractor_id"]
            project_id = request.form["project_id"]
        else:
            subcontractor_id = current_subcontractor_id()
            project_id = current_project_id() or 1
        if is_project_admin() and str(project_id) != str(current_project_id()):
            flash("Tu ne peux saisir que pour ton projet autorisé.")
            return redirect(url_for("entries"))
        if not declared_product:
            flash("Choisis un produit validé ou saisis un nouveau produit déclaré.")
            return redirect(url_for("entries"))

        conn.execute("""INSERT INTO entries(project_id,entry_date,month,subcontractor_id,declared_product,product_id,qty_in,qty_used,qty_stock,unit,storage_location,fds_available,observations)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (project_id, request.form["entry_date"], request.form["month"], subcontractor_id, declared_product, product_id,
                      qin, qused, qstock, request.form.get("unit","u"), request.form.get("storage_location",""), fds_available, request.form.get("observations","")))
        conn.commit()
        log_action("Saisie produit", declared_product)
        flash("Saisie enregistrée.")

    products = conn.execute("SELECT * FROM products ORDER BY root_name").fetchall()
    product_opts = "".join([f"<option value='{p['id']}'>{p['root_name']} ({p['family'] or 'Famille non définie'})</option>" for p in products])

    if can_access_admin_features():
        projects = conn.execute("SELECT * FROM projects WHERE status='Actif' ORDER BY name").fetchall() if is_super_admin() else conn.execute("SELECT * FROM projects WHERE id=?", (current_project_id(),)).fetchall()
        project_opts = "".join([f"<option value='{p['id']}'>{p['name']}</option>" for p in projects])
        sts = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall()
        st_opts = "".join([f"<option value='{s['id']}'>{s['name']}</option>" for s in sts])
        admin_fields = f"""<div class="row"><div><label>Projet</label><select name="project_id">{project_opts}</select></div><div><label>Sous-traitant</label><select name="subcontractor_id">{st_opts}</select></div></div>"""
        where, params = where_scope("e")
        rows = conn.execute(f"""SELECT e.*,s.name st,p.root_name,pr.name project_name FROM entries e 
                               JOIN subcontractors s ON s.id=e.subcontractor_id 
                               LEFT JOIN products p ON p.id=e.product_id
                               LEFT JOIN projects pr ON pr.id=e.project_id
                               {where} ORDER BY e.entry_date DESC,e.id DESC LIMIT 300""", params).fetchall()
    else:
        admin_fields = "<div class='help'>Tu saisis uniquement pour ton entreprise et ton projet.</div>"
        rows = conn.execute("""SELECT e.*,s.name st,p.root_name,pr.name project_name FROM entries e 
                               JOIN subcontractors s ON s.id=e.subcontractor_id 
                               LEFT JOIN products p ON p.id=e.product_id
                               LEFT JOIN projects pr ON pr.id=e.project_id
                               WHERE e.subcontractor_id=?
                               ORDER BY e.entry_date DESC,e.id DESC LIMIT 300""", (current_subcontractor_id(),)).fetchall()
    conn.close()

    trs = "".join([f"<tr><td>{r['entry_date']}</td><td>{r['month']}</td><td>{r['project_name'] or ''}</td><td>{r['st']}</td><td>{r['declared_product']}</td><td>{r['root_name'] or '<span class=danger>Non reconnu</span>'}</td><td>{r['qty_in']}</td><td>{r['qty_used']}</td><td>{r['qty_stock']}</td><td>{r['unit']}</td><td>{r['fds_available']}</td><td>{r['storage_location'] or ''}</td><td><a class='btn' href='{url_for('edit_entry', entry_id=r['id'])}'>Modifier</a> <form method='post' action='{url_for('delete_entry', entry_id=r['id'])}' style='display:inline' onsubmit='return confirm(&quot;Supprimer cette saisie ?&quot;)'><button>Supprimer</button></form></td></tr>" for r in rows])
    return layout("Saisie", f"""<h2>Saisie mensuelle</h2><div class="help">Choisis de préférence un produit validé. Le stock est calculé automatiquement si tu laisses le champ stock vide.</div>
    <div class="box"><form method="post">{admin_fields}
    <div class="row"><div><label>Date</label><input type="date" name="entry_date" required></div><div><label>Mois</label><input name="month" placeholder="2026-06" required></div></div>
    <label>Produit validé</label><select name="product_id"><option value="">-- Nouveau produit / non listé --</option>{product_opts}</select>
    <label>Nouveau produit déclaré</label><input name="declared_product_new" placeholder="À remplir seulement si le produit n'est pas dans la liste">
    <div class="row"><div><label>Quantité entrée</label><input id="qty_in" type="number" step="0.01" name="qty_in" value="0" placeholder="Quantité reçue"></div><div><label>Quantité utilisée</label><input id="qty_used" type="number" step="0.01" name="qty_used" value="0" placeholder="Quantité consommée"></div></div>
    <div class="row"><div><label>Stock restant</label><input id="qty_stock" type="number" step="0.01" name="qty_stock" placeholder="Calcul automatique : entrée - utilisée"></div><div><label>Unité</label><select name="unit"><option>u</option><option>kg</option><option>L</option><option>seau</option><option>sac</option><option>rouleau</option><option>bidon</option></select></div></div>
    <label>Lieu de stockage</label><input name="storage_location" placeholder="Ex : magasin peinture, local étanchéité"><label>FDS disponible</label><select name="fds_available"><option>OUI</option><option selected>NON</option></select>
    <label>Observations</label><textarea name="observations"></textarea><button>Enregistrer</button></form></div>
    <table><tr><th>Date</th><th>Mois</th><th>Projet</th><th>Sous-traitant</th><th>Produit déclaré</th><th>Produit racine</th><th>Entrée</th><th>Utilisée</th><th>Stock</th><th>Unité</th><th>FDS</th><th>Stockage</th><th>Actions</th></tr>{trs}</table>
    <script>
    const qi=document.getElementById('qty_in'); const qu=document.getElementById('qty_used'); const qs=document.getElementById('qty_stock');
    function calcStock(){{ if(!qi||!qu||!qs) return; const a=parseFloat(qi.value)||0; const b=parseFloat(qu.value)||0; qs.value=(a-b).toFixed(2); }}
    if(qi&&qu&&qs){{ qi.addEventListener('input', calcStock); qu.addEventListener('input', calcStock); }}
    </script>""")


@app.route("/entries/edit/<int:entry_id>", methods=["GET","POST"])
def edit_entry(entry_id):
    if not logged() or is_reader():
        return redirect(url_for("dashboard"))
    conn = db()
    entry = conn.execute("SELECT * FROM entries WHERE id=?", (entry_id,)).fetchone()
    if not entry:
        conn.close(); flash("Saisie introuvable."); return redirect(url_for("entries"))
    if not can_manage_entry(entry):
        conn.close(); flash("Accès refusé."); return redirect(url_for("entries"))
    if request.method == "POST":
        qin = safe_float(request.form.get("qty_in"))
        qused = safe_float(request.form.get("qty_used"))
        qstock = safe_float(request.form.get("qty_stock"), qin - qused)
        conn.execute("""UPDATE entries SET entry_date=?, month=?, qty_in=?, qty_used=?, qty_stock=?, unit=?, storage_location=?, fds_available=?, observations=? WHERE id=?""",
                     (request.form["entry_date"], request.form["month"], qin, qused, qstock, request.form.get("unit","u"), request.form.get("storage_location",""), request.form.get("fds_available","NON"), request.form.get("observations",""), entry_id))
        conn.commit(); conn.close()
        log_action("Modification saisie", f"Saisie ID {entry_id}")
        flash("Saisie modifiée.")
        return redirect(url_for("entries"))
    conn.close()
    unit_options = "".join([f"<option {'selected' if entry['unit']==u else ''}>{u}</option>" for u in ['u','kg','L','seau','sac','rouleau','bidon']])
    fds_options = f"<option {'selected' if entry['fds_available']=='OUI' else ''}>OUI</option><option {'selected' if entry['fds_available']=='NON' else ''}>NON</option>"
    return layout("Modifier une saisie", f"""<h2>Modifier une saisie</h2>
    <div class="box"><form method="post">
    <div class="row"><div><label>Date</label><input type="date" name="entry_date" value="{entry['entry_date']}" required></div>
    <div><label>Mois</label><input name="month" value="{entry['month']}" required placeholder="Ex : 2026-06"></div></div>
    <div class="row"><div><label>Quantité entrée</label><input id="edit_qty_in" type="number" step="0.01" name="qty_in" value="{entry['qty_in']}" placeholder="Quantité reçue"></div>
    <div><label>Quantité utilisée</label><input id="edit_qty_used" type="number" step="0.01" name="qty_used" value="{entry['qty_used']}" placeholder="Quantité consommée"></div></div>
    <div class="row"><div><label>Stock restant</label><input id="edit_qty_stock" type="number" step="0.01" name="qty_stock" value="{entry['qty_stock']}" placeholder="Calcul automatique"></div>
    <div><label>Unité</label><select name="unit">{unit_options}</select></div></div>
    <label>Lieu de stockage</label><input name="storage_location" value="{entry['storage_location'] or ''}" placeholder="Ex : magasin peinture, local étanchéité">
    <label>FDS disponible</label><select name="fds_available">{fds_options}</select>
    <label>Observations</label><textarea name="observations" placeholder="Commentaires, anomalies, remarques">{entry['observations'] or ''}</textarea>
    <button>Enregistrer les modifications</button> <a class="btn" href="{url_for('entries')}">Annuler</a>
    </form></div>
    <script>
    const qi=document.getElementById('edit_qty_in'); const qu=document.getElementById('edit_qty_used'); const qs=document.getElementById('edit_qty_stock');
    function calcStock(){{ const a=parseFloat(qi.value)||0; const b=parseFloat(qu.value)||0; qs.value=(a-b).toFixed(2); }}
    qi.addEventListener('input', calcStock); qu.addEventListener('input', calcStock);
    </script>""")

@app.route("/entries/delete/<int:entry_id>", methods=["POST"])
def delete_entry(entry_id):
    if not logged() or is_reader():
        return redirect(url_for("dashboard"))
    conn = db()
    entry = conn.execute("SELECT * FROM entries WHERE id=?", (entry_id,)).fetchone()
    if not entry:
        conn.close(); flash("Saisie introuvable."); return redirect(url_for("entries"))
    if not can_manage_entry(entry):
        conn.close(); flash("Accès refusé."); return redirect(url_for("entries"))
    conn.execute("DELETE FROM entries WHERE id=?", (entry_id,)); conn.commit(); conn.close()
    log_action("Suppression saisie", f"Saisie ID {entry_id}")
    flash("Saisie supprimée.")
    return redirect(url_for("entries"))

@app.route("/alerts")
def alerts():
    if not logged() or is_subcontractor(): return redirect(url_for("dashboard"))
    trs = ""
    for level, typ, product, action in get_alert_rows():
        cls = "danger" if level == "CRITIQUE" else "warn"
        trs += f"<tr><td class='{cls}'>{level}</td><td>{typ}</td><td>{product}</td><td>{action}</td></tr>"
    if not trs:
        trs = "<tr><td colspan='4' class='ok'>Aucune alerte majeure détectée.</td></tr>"
    return layout("Alertes QHSE", f"""<h2>Alertes QHSE</h2><div class="help">Alertes automatiques : FDS, produits non reconnus et seuils de stock.</div>
    <table><tr><th>Niveau</th><th>Type</th><th>Produit</th><th>Action recommandée</th></tr>{trs}</table>""")

@app.route("/consolidation")
def consolidation():
    if not logged() or is_subcontractor(): return redirect(url_for("dashboard"))
    trs = "".join([f"<tr><td>{r['projects'] or ''}</td><td>{r['product_name']}</td><td>{r['family'] or ''}</td><td>{r['conditionnement'] or ''}</td><td>{r['utilisation'] or ''}</td><td>{r['fds'] or ''}</td><td>{r['fds_expiry'] or ''}</td><td>{r['pictogrammes'] or ''}</td><td>{r['incompatibilites'] or ''}</td><td>{r['total_in']}</td><td>{r['total_used']}</td><td>{r['total_stock']}</td><td>{r['unit']}</td><td>{r['nb_st']}</td><td>{r['sts']}</td><td>{r['locations'] or ''}</td></tr>" for r in consolidation_rows()])
    export_button = f"<p><a class='btn' href='{url_for('export_excel')}'>Exporter en Excel</a></p>" if can_access_admin_features() else ""
    return layout("Consolidation", f"""<h2>Consolidation générale</h2>{export_button}
    <table><tr><th>Projet</th><th>Produit</th><th>Famille</th><th>Conditionnement</th><th>Utilisation</th><th>FDS</th><th>Exp. FDS</th><th>Pictogrammes</th><th>Incompatibilités</th><th>Entrée</th><th>Utilisée</th><th>Stock</th><th>Unité</th><th>Nb ST</th><th>Sous-traitants</th><th>Lieux stockage</th></tr>{trs}</table>""")

@app.route("/export_excel", methods=["GET","POST"])
def export_excel():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "GET" and not request.args.get("generate"):
        if is_super_admin():
            projects = conn.execute("SELECT * FROM projects ORDER BY name").fetchall()
        else:
            projects = conn.execute("SELECT * FROM projects WHERE id=?", (current_project_id(),)).fetchall()
        sts = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall(); conn.close()
        project_opts = "<option value=''>Tous les projets autorisés</option>" + "".join([f"<option value='{p['id']}'>{p['name']}</option>" for p in projects])
        st_opts = "<option value=''>Tous les sous-traitants</option>" + "".join([f"<option value='{st['id']}'>{st['name']}</option>" for st in sts])
        return layout("Rapport Excel", f"""<h2>Générer rapport Excel</h2>
        <div class="help">Le rapport Excel est réservé au Super Admin et aux Admin Projet. Sélectionne une période, puis filtre éventuellement par projet ou sous-traitant.</div>
        <div class="box"><form method="get">
        <input type="hidden" name="generate" value="1">
        <div class="row"><div><label>Date début</label><input type="date" name="start_date" required placeholder="Début de la période"></div><div><label>Date fin</label><input type="date" name="end_date" required placeholder="Fin de la période"></div></div>
        <div class="row"><div><label>Projet</label><select name="project_id">{project_opts}</select></div><div><label>Sous-traitant</label><select name="subcontractor_id">{st_opts}</select></div></div>
        <button>Générer le rapport Excel</button></form></div>""")
    start_date = request.values.get("start_date", "")
    end_date = request.values.get("end_date", "")
    project_filter = request.values.get("project_id") or None
    st_filter = request.values.get("subcontractor_id") or None
    clauses=[]; params=[]
    if start_date: clauses.append("e.entry_date>=?"); params.append(start_date)
    if end_date: clauses.append("e.entry_date<=?"); params.append(end_date)
    if project_filter: clauses.append("e.project_id=?"); params.append(project_filter)
    if st_filter: clauses.append("e.subcontractor_id=?"); params.append(st_filter)
    if is_project_admin(): clauses.append("e.project_id=?"); params.append(current_project_id())
    where = " WHERE " + " AND ".join(clauses) if clauses else ""

    detail_rows = conn.execute(f"""SELECT e.*, COALESCE(p.root_name,e.declared_product) product_name, p.family,p.conditionnement,p.utilisation,p.fds,p.fds_expiry,p.pictogrammes,p.incompatibilites,p.danger_class,p.epi,s.name st_name,pr.name project_name
        FROM entries e
        LEFT JOIN products p ON p.id=e.product_id
        JOIN subcontractors s ON s.id=e.subcontractor_id
        LEFT JOIN projects pr ON pr.id=e.project_id
        {where}
        ORDER BY e.entry_date, pr.name, s.name, product_name""", params).fetchall()
    summary_rows = conn.execute(f"""SELECT COALESCE(pr.name,'') project_name, COALESCE(p.root_name,e.declared_product) product_name, COALESCE(p.family,'') family, e.unit,
        SUM(e.qty_in) total_in, SUM(e.qty_used) total_used, SUM(e.qty_stock) total_stock,
        COUNT(DISTINCT e.subcontractor_id) nb_st, GROUP_CONCAT(DISTINCT s.name) sts, GROUP_CONCAT(DISTINCT e.storage_location) locations,
        p.fds, p.fds_expiry, p.pictogrammes, p.incompatibilites, p.danger_class, p.epi, p.stock_min, p.stock_max
        FROM entries e
        LEFT JOIN products p ON p.id=e.product_id
        JOIN subcontractors s ON s.id=e.subcontractor_id
        LEFT JOIN projects pr ON pr.id=e.project_id
        {where}
        GROUP BY project_name, product_name, e.unit
        ORDER BY project_name, product_name""", params).fetchall()
    conn.close()

    wb = Workbook(); ws = wb.active; ws.title = "Synthèse période"
    ws.append(["Période", f"{start_date or 'début'} au {end_date or 'fin'}"])
    headers = ["Projet","Produit","Famille","FDS","Expiration FDS","Pictogrammes","Incompatibilités","Classe danger","EPI","Entrée totale","Utilisée totale","Stock total","Unité","Stock min","Stock max","Nb sous-traitants","Sous-traitants","Lieux stockage"]
    ws.append(headers)
    for r in summary_rows:
        ws.append([r["project_name"],r["product_name"],r["family"],r["fds"],r["fds_expiry"],r["pictogrammes"],r["incompatibilites"],r["danger_class"],r["epi"],r["total_in"],r["total_used"],r["total_stock"],r["unit"],r["stock_min"],r["stock_max"],r["nb_st"],r["sts"],r["locations"]])
    ws_detail = wb.create_sheet("Détails saisies")
    ws_detail.append(["Date","Mois","Projet","Sous-traitant","Produit déclaré","Produit racine","Famille","Entrée","Utilisée","Stock","Unité","FDS saisie","Lieu stockage","Observations"])
    for r in detail_rows:
        ws_detail.append([r["entry_date"],r["month"],r["project_name"],r["st_name"],r["declared_product"],r["product_name"],r["family"],r["qty_in"],r["qty_used"],r["qty_stock"],r["unit"],r["fds_available"],r["storage_location"],r["observations"]])
    ws_alert = wb.create_sheet("Alertes")
    ws_alert.append(["Niveau","Type","Produit","Action recommandée"])
    for a in get_alert_rows(): ws_alert.append(list(a))
    ws_about = wb.create_sheet("À propos")
    for row in [["Logiciel",APP_NAME],["Version",VERSION],["Auteur","Kodjotse Eli ADIGBLI"],["Copyright",COPYRIGHT],["Période",f"{start_date or 'début'} au {end_date or 'fin'}"],["Généré le",datetime.now().strftime("%d/%m/%Y %H:%M")]]: ws_about.append(row)
    fill = PatternFill("solid", fgColor="1F3347")
    for sheet in wb.worksheets:
        if sheet.max_row:
            header_row = 2 if sheet.title == "Synthèse période" else 1
            for cell in sheet[header_row]:
                cell.font = Font(bold=True, color="FFFFFF"); cell.fill = fill; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in sheet.columns:
                sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max(len(str(c.value or "")) for c in col)+2, 55)
    filename = EXPORT_DIR / f"Rapport_Produits_Chimiques_{start_date or 'debut'}_{end_date or 'fin'}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(filename); log_action("Export Excel période", filename.name)
    return send_file(filename, as_attachment=True)

@app.route("/audit_trail")
def audit_trail():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    rows = conn.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT 300").fetchall()
    conn.close()
    trs = "".join([f"<tr><td>{r['action_date']}</td><td>{r['username']}</td><td>{r['action']}</td><td>{r['details']}</td></tr>" for r in rows])
    return layout("Journal des actions", f"""<h2>Journal des actions</h2><table><tr><th>Date</th><th>Utilisateur</th><th>Action</th><th>Détails</th></tr>{trs}</table>""")


@app.route("/health")
def health():
    return {"status": "ok", "app": APP_NAME, "version": VERSION}

@app.route("/about")
def about():
    if not logged(): return redirect(url_for("login"))
    return layout("À propos", f"""<h2>À propos</h2><div class="box">
    <p><strong>{APP_NAME}</strong></p><p>Version : <strong>{VERSION}</strong></p>
    <p>Développé par : <strong>Kodjotse Eli ADIGBLI</strong></p><p>{COPYRIGHT}</p></div>""")

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000")

# Cloud initialization
try:
    init_db()
except Exception as cloud_init_error:
    print("Database initialization warning:", cloud_init_error)

if __name__ == "__main__":
    init_db()
    cloud_mode = os.environ.get("CLOUD_MODE", "false").lower() == "true"
    port = int(os.environ.get("PORT", "5000"))
    host = "0.0.0.0" if cloud_mode else "127.0.0.1"
    if not cloud_mode:
        threading.Timer(1.2, open_browser).start()
    app.run(host=host, port=port, debug=False)
