
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
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_NAME = "QHSE Manager Pro - Chemical Register"
VERSION = "v2.0 CLOUD READY"
COPYRIGHT = "Copyright © 2026 Kodjotse Eli ADIGBLI. Tous droits réservés."

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "qhse_chemical_register.db"
EXPORT_DIR = BASE_DIR / "exports"
UPLOAD_DIR = BASE_DIR / "uploads"
EXPORT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "qhse_manager_pro_change_this_key")

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
        cur.execute("INSERT INTO users(username,password_hash,role,status) VALUES(?,?,?,?)",
                    ("admin", generate_password_hash("admin123"), "super_admin", "Actif"))

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
    return layout("Connexion", """<h2>Connexion</h2><div class="box"><form method="post">
        <label>Utilisateur</label><input name="username" value="admin" required>
        <label>Mot de passe</label><input name="password" type="password" value="admin123" required>
        <button>Se connecter</button></form></div>""")

@app.route("/logout")
def logout():
    log_action("Déconnexion", "Déconnexion utilisateur")
    session.clear()
    return redirect(url_for("login"))

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
        <div class="card">Version<div class="value ok">2.0</div></div>
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
    if not logged() or not is_super_admin(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        conn.execute("""INSERT OR IGNORE INTO projects(name,location,country,client,owner,main_contractor,project_manager,qhse_manager,coordinator,start_date,end_date,status,description)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (request.form["name"], request.form.get("location",""), request.form.get("country",""), request.form.get("client",""),
                      request.form.get("owner",""), request.form.get("main_contractor",""), request.form.get("project_manager",""),
                      request.form.get("qhse_manager",""), request.form.get("coordinator",""), request.form.get("start_date",""),
                      request.form.get("end_date",""), request.form.get("status","Actif"), request.form.get("description","")))
        conn.commit()
        log_action("Création projet", request.form["name"])
    rows = conn.execute("SELECT * FROM projects ORDER BY name").fetchall()
    conn.close()
    trs = "".join([f"<tr><td>{r['name']}</td><td>{r['client'] or ''}</td><td>{r['main_contractor'] or ''}</td><td>{r['qhse_manager'] or ''}</td><td>{r['location']}</td><td>{r['country'] or ''}</td><td>{r['status']}</td></tr>" for r in rows])
    return layout("Projets", f"""<h2>Projets</h2><div class="box"><form method="post">
    <div class="row3"><div><label>Nom du projet</label><input name="name" required></div><div><label>Ville</label><input name="location"></div><div><label>Pays</label><input name="country" value="Togo"></div></div>
    <div class="row3"><div><label>Client</label><input name="client"></div><div><label>Maître d'ouvrage</label><input name="owner"></div><div><label>Entreprise principale</label><input name="main_contractor" value="Ellipse Projects"></div></div>
    <div class="row3"><div><label>Chef de projet</label><input name="project_manager"></div><div><label>Responsable QHSE</label><input name="qhse_manager"></div><div><label>Coordinateur</label><input name="coordinator"></div></div>
    <div class="row3"><div><label>Date début</label><input type="date" name="start_date"></div><div><label>Date fin</label><input type="date" name="end_date"></div><div><label>Statut</label><select name="status"><option>Actif</option><option>Clôturé</option></select></div></div>
    <label>Description</label><textarea name="description"></textarea><button>Ajouter</button></form></div>
    <table><tr><th>Projet</th><th>Client</th><th>Entreprise principale</th><th>Resp. QHSE</th><th>Ville</th><th>Pays</th><th>Statut</th></tr>{trs}</table>""")

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
        if is_project_admin():
            if r not in ["subcontractor","reader"]:
                flash("Un Admin Projet ne peut créer que des comptes Sous-traitant ou Lecteur.")
                return redirect(url_for("users"))
            project_id = current_project_id()
        if r == "super_admin":
            project_id = None; subcontractor_id = None
        if r in ["project_admin","reader"]:
            subcontractor_id = None
        if r == "subcontractor" and not subcontractor_id:
            flash("Choisis un sous-traitant associé.")
            return redirect(url_for("users"))
        try:
            conn.execute("INSERT INTO users(username,password_hash,role,subcontractor_id,project_id,status) VALUES(?,?,?,?,?,?)",
                         (username, generate_password_hash(password), r, subcontractor_id, project_id, request.form.get("status","Actif")))
            conn.commit()
            log_action("Création utilisateur", username)
            flash("Utilisateur créé.")
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
    sts = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall()
    conn.close()

    role_options = "<option value='subcontractor'>Sous-traitant</option><option value='reader'>Lecteur</option>"
    if is_super_admin():
        role_options = "<option value='subcontractor'>Sous-traitant</option><option value='project_admin'>Admin Projet</option><option value='reader'>Lecteur</option><option value='super_admin'>Super Admin</option>"
    project_opts = "".join([f"<option value='{p['id']}'>{p['name']}</option>" for p in projects_rows])
    st_opts = "".join([f"<option value='{s['id']}'>{s['name']}</option>" for s in sts])
    trs = "".join([f"<tr><td>{u['username']}</td><td>{u['role']}</td><td>{u['project_name'] or 'Tous'}</td><td>{u['st_name'] or '-'}</td><td>{u['status']}</td></tr>" for u in users_rows])
    return layout("Utilisateurs", f"""<h2>Utilisateurs et droits</h2>
    <div class="help"><strong>Super Admin</strong> : tous les projets. <strong>Admin Projet</strong> : son projet. <strong>Sous-traitant</strong> : ses saisies. <strong>Lecteur</strong> : consultation.</div>
    <div class="box"><form method="post">
    <div class="row"><div><label>Nom d'utilisateur</label><input name="username" required></div><div><label>Mot de passe provisoire</label><input name="password" required></div></div>
    <div class="row"><div><label>Rôle</label><select name="role">{role_options}</select></div><div><label>Projet autorisé</label><select name="project_id"><option value="">Tous / Aucun</option>{project_opts}</select></div></div>
    <div class="row"><div><label>Sous-traitant associé</label><select name="subcontractor_id"><option value="">-- aucun --</option>{st_opts}</select></div><div><label>Statut</label><select name="status"><option>Actif</option><option>Inactif</option></select></div></div>
    <button>Créer l'utilisateur</button></form></div>
    <table><tr><th>Utilisateur</th><th>Rôle</th><th>Projet</th><th>Sous-traitant associé</th><th>Statut</th></tr>{trs}</table>""")

@app.route("/subcontractors", methods=["GET","POST"])
def subcontractors():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        conn.execute("""INSERT OR IGNORE INTO subcontractors(name,contact,phone,email,address,lot,work_zone,employees_count,start_date,end_date,hse_plan,insurance,ppsps,status,observations)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (request.form["name"].upper(), request.form.get("contact",""), request.form.get("phone",""), request.form.get("email",""),
                      request.form.get("address",""), request.form.get("lot",""), request.form.get("work_zone",""), request.form.get("employees_count") or 0,
                      request.form.get("start_date",""), request.form.get("end_date",""), request.form.get("hse_plan","NON"), request.form.get("insurance","NON"),
                      request.form.get("ppsps","NON"), request.form.get("status","Actif"), request.form.get("observations","")))
        conn.commit()
        log_action("Création sous-traitant", request.form["name"])
    rows = conn.execute("SELECT * FROM subcontractors ORDER BY name").fetchall()
    conn.close()
    trs = "".join([f"<tr><td>{r['name']}</td><td>{r['contact']}</td><td>{r['phone']}</td><td>{r['email']}</td><td>{r['lot'] or ''}</td><td>{r['work_zone'] or ''}</td><td>{r['employees_count'] or 0}</td><td>{r['hse_plan']}</td><td>{r['insurance']}</td><td>{r['ppsps']}</td><td>{r['status']}</td></tr>" for r in rows])
    return layout("Sous-traitants", f"""<h2>Sous-traitants</h2><div class="box"><form method="post">
    <div class="row3"><div><label>Nom</label><input name="name" required></div><div><label>Contact</label><input name="contact"></div><div><label>Téléphone</label><input name="phone"></div></div>
    <div class="row3"><div><label>Email</label><input name="email"></div><div><label>Lot</label><input name="lot"></div><div><label>Zone d'intervention</label><input name="work_zone"></div></div>
    <label>Adresse</label><input name="address">
    <div class="row3"><div><label>Nombre employés</label><input type="number" name="employees_count" value="0"></div><div><label>Date arrivée</label><input type="date" name="start_date"></div><div><label>Date départ</label><input type="date" name="end_date"></div></div>
    <div class="row3"><div><label>Plan HSE</label><select name="hse_plan"><option>OUI</option><option selected>NON</option></select></div><div><label>Assurance</label><select name="insurance"><option>OUI</option><option selected>NON</option></select></div><div><label>PPSPS</label><select name="ppsps"><option>OUI</option><option selected>NON</option></select></div></div>
    <label>Statut</label><select name="status"><option>Actif</option><option>Inactif</option></select><label>Observations</label><textarea name="observations"></textarea>
    <button>Ajouter</button></form></div>
    <table><tr><th>Nom</th><th>Contact</th><th>Téléphone</th><th>Email</th><th>Lot</th><th>Zone</th><th>Employés</th><th>Plan HSE</th><th>Assurance</th><th>PPSPS</th><th>Statut</th></tr>{trs}</table>""")

@app.route("/products", methods=["GET","POST"])
def products():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        conn.execute("""INSERT OR IGNORE INTO products(root_name,commercial_name,manufacturer,reference,family,physical_state,conditionnement,utilisation,fds,fds_date,fds_version,fds_expiry,pictogrammes,danger_class,h_statements,p_statements,incompatibilites,epi,first_aid,fire_measures,spill_measures,storage_rules,stock_min,stock_max,observations)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (request.form["root_name"].upper(), request.form.get("commercial_name",""), request.form.get("manufacturer",""), request.form.get("reference",""),
         request.form.get("family",""), request.form.get("physical_state",""), request.form.get("conditionnement",""), request.form.get("utilisation",""),
         request.form.get("fds","NON"), request.form.get("fds_date",""), request.form.get("fds_version",""), request.form.get("fds_expiry",""),
         request.form.get("pictogrammes",""), request.form.get("danger_class",""), request.form.get("h_statements",""), request.form.get("p_statements",""),
         request.form.get("incompatibilites",""), request.form.get("epi",""), request.form.get("first_aid",""), request.form.get("fire_measures",""),
         request.form.get("spill_measures",""), request.form.get("storage_rules",""), request.form.get("stock_min") or 0, request.form.get("stock_max") or 0,
         request.form.get("observations","")))
        conn.commit()
        log_action("Création produit", request.form["root_name"])
    rows = conn.execute("SELECT * FROM products ORDER BY root_name").fetchall()
    conn.close()
    trs = "".join([f"<tr><td>{r['root_name']}</td><td>{r['family'] or ''}</td><td>{r['conditionnement'] or ''}</td><td>{r['utilisation'] or ''}</td><td>{r['fds']}</td><td>{r['fds_expiry'] or ''}</td><td>{r['pictogrammes'] or ''}</td><td>{r['incompatibilites'] or ''}</td><td>{r['stock_min'] or 0}</td><td>{r['stock_max'] or 0}</td></tr>" for r in rows])
    return layout("Base produits", f"""<h2>Base produits chimiques</h2><div class="box"><form method="post">
    <h3>Identification</h3>
    <div class="row3"><div><label>Produit racine</label><input name="root_name" required></div><div><label>Nom commercial</label><input name="commercial_name"></div><div><label>Fabricant</label><input name="manufacturer"></div></div>
    <div class="row3"><div><label>Référence</label><input name="reference"></div><div><label>Famille</label><input name="family"></div><div><label>État physique</label><input name="physical_state"></div></div>
    <label>Conditionnement</label><input name="conditionnement"><label>Utilisation</label><input name="utilisation">
    <h3>FDS et sécurité</h3>
    <div class="row3"><div><label>FDS</label><select name="fds"><option>OUI</option><option selected>NON</option></select></div><div><label>Date FDS</label><input type="date" name="fds_date"></div><div><label>Expiration FDS</label><input type="date" name="fds_expiry"></div></div>
    <div class="row3"><div><label>Version FDS</label><input name="fds_version"></div><div><label>Pictogrammes</label><input name="pictogrammes"></div><div><label>Classe de danger</label><input name="danger_class"></div></div>
    <label>Mentions H</label><input name="h_statements"><label>Mentions P</label><input name="p_statements"><label>Incompatibilités</label><input name="incompatibilites"><label>EPI requis</label><input name="epi">
    <label>Premiers secours</label><textarea name="first_aid"></textarea><label>Mesures incendie</label><textarea name="fire_measures"></textarea><label>Déversement accidentel</label><textarea name="spill_measures"></textarea><label>Règles de stockage</label><textarea name="storage_rules"></textarea>
    <h3>Seuils de stock</h3><div class="row"><div><label>Stock minimum</label><input type="number" step="0.01" name="stock_min" value="0"></div><div><label>Stock maximum</label><input type="number" step="0.01" name="stock_max" value="0"></div></div>
    <label>Observations</label><textarea name="observations"></textarea><button>Ajouter</button></form></div>
    <table><tr><th>Produit</th><th>Famille</th><th>Conditionnement</th><th>Utilisation</th><th>FDS</th><th>Expiration FDS</th><th>Pictogrammes</th><th>Incompatibilités</th><th>Stock min</th><th>Stock max</th></tr>{trs}</table>""")

@app.route("/aliases", methods=["GET","POST"])
def aliases():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    conn = db()
    if request.method == "POST":
        conn.execute("INSERT OR IGNORE INTO aliases(alias_name,product_id) VALUES(?,?)", (normalize(request.form["alias_name"]), request.form["product_id"]))
        conn.commit()
        log_action("Création alias", request.form["alias_name"])
    products = conn.execute("SELECT * FROM products ORDER BY root_name").fetchall()
    aliases = conn.execute("SELECT a.alias_name,p.root_name FROM aliases a JOIN products p ON p.id=a.product_id ORDER BY a.alias_name").fetchall()
    conn.close()
    opts = "".join([f"<option value='{p['id']}'>{p['root_name']}</option>" for p in products])
    trs = "".join([f"<tr><td>{a['alias_name']}</td><td>{a['root_name']}</td></tr>" for a in aliases])
    return layout("Alias", f"""<h2>Alias / Doublons</h2><div class="help">Exemple : rattacher “PANTEX VELOUR 20KG” à “PANTEX VELOUR”.</div>
    <div class="box"><form method="post"><label>Nom saisi / alias</label><input name="alias_name" required><label>Produit racine</label><select name="product_id">{opts}</select><button>Créer l'alias</button></form></div>
    <table><tr><th>Alias</th><th>Produit racine</th></tr>{trs}</table>""")

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

    trs = "".join([f"<tr><td>{r['entry_date']}</td><td>{r['month']}</td><td>{r['project_name'] or ''}</td><td>{r['st']}</td><td>{r['declared_product']}</td><td>{r['root_name'] or '<span class=danger>Non reconnu</span>'}</td><td>{r['qty_in']}</td><td>{r['qty_used']}</td><td>{r['qty_stock']}</td><td>{r['unit']}</td><td>{r['fds_available']}</td><td>{r['storage_location'] or ''}</td></tr>" for r in rows])
    return layout("Saisie", f"""<h2>Saisie mensuelle</h2><div class="help">Choisis de préférence un produit validé. Le stock est calculé automatiquement si tu laisses le champ stock vide.</div>
    <div class="box"><form method="post">{admin_fields}
    <div class="row"><div><label>Date</label><input type="date" name="entry_date" required></div><div><label>Mois</label><input name="month" placeholder="2026-06" required></div></div>
    <label>Produit validé</label><select name="product_id"><option value="">-- Nouveau produit / non listé --</option>{product_opts}</select>
    <label>Nouveau produit déclaré</label><input name="declared_product_new" placeholder="À remplir seulement si le produit n'est pas dans la liste">
    <div class="row"><div><label>Quantité entrée</label><input type="number" step="0.01" name="qty_in" value="0"></div><div><label>Quantité utilisée</label><input type="number" step="0.01" name="qty_used" value="0"></div></div>
    <div class="row"><div><label>Stock restant</label><input type="number" step="0.01" name="qty_stock" placeholder="Vide = entrée - utilisée"></div><div><label>Unité</label><select name="unit"><option>u</option><option>kg</option><option>L</option><option>seau</option><option>sac</option><option>rouleau</option><option>bidon</option></select></div></div>
    <label>Lieu de stockage</label><input name="storage_location"><label>FDS disponible</label><select name="fds_available"><option>OUI</option><option selected>NON</option></select>
    <label>Observations</label><textarea name="observations"></textarea><button>Enregistrer</button></form></div>
    <table><tr><th>Date</th><th>Mois</th><th>Projet</th><th>Sous-traitant</th><th>Produit déclaré</th><th>Produit racine</th><th>Entrée</th><th>Utilisée</th><th>Stock</th><th>Unité</th><th>FDS</th><th>Stockage</th></tr>{trs}</table>""")

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

@app.route("/export_excel")
def export_excel():
    if not logged() or not can_access_admin_features(): return redirect(url_for("dashboard"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Rapport Général"
    headers = ["Projet","Produit","Famille","Conditionnement","Utilisation","FDS","Expiration FDS","Pictogrammes","Incompatibilités","Classe danger","EPI","Entrée totale","Utilisée totale","Stock total","Unité","Stock min","Stock max","Nb sous-traitants","Sous-traitants","Lieux stockage"]
    ws.append(headers)
    for r in consolidation_rows():
        ws.append([r["projects"],r["product_name"],r["family"],r["conditionnement"],r["utilisation"],r["fds"],r["fds_expiry"],r["pictogrammes"],r["incompatibilites"],r["danger_class"],r["epi"],r["total_in"],r["total_used"],r["total_stock"],r["unit"],r["stock_min"],r["stock_max"],r["nb_st"],r["sts"],r["locations"]])
    ws_alert = wb.create_sheet("Alertes")
    ws_alert.append(["Niveau","Type","Produit","Action recommandée"])
    for a in get_alert_rows():
        ws_alert.append(list(a))
    ws_about = wb.create_sheet("À propos")
    for row in [["Logiciel",APP_NAME],["Version",VERSION],["Auteur","Kodjotse Eli ADIGBLI"],["Copyright",COPYRIGHT],["Généré le",datetime.now().strftime("%d/%m/%Y %H:%M")]]:
        ws_about.append(row)
    fill = PatternFill("solid", fgColor="1F3347")
    for sheet in wb.worksheets:
        if sheet.max_row:
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in sheet.columns:
                sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max(len(str(c.value or "")) for c in col)+2, 55)
    filename = EXPORT_DIR / f"Rapport_Produits_Chimiques_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(filename)
    log_action("Export Excel", filename.name)
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
