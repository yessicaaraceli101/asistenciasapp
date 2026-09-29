# app.py
from flask import Flask, render_template, request, redirect, url_for, send_file, session, flash, jsonify, g
import sqlite3
import io
import os
import requests
import csv, re, unicodedata
import json
from docx import Document
from fpdf import FPDF
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, date, timedelta
from unicodedata import normalize
from io import StringIO, BytesIO

# Variables de entorno desde un archivo .env (opcional: pip install python-dotenv)
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass

# Config via ENV con defaults
VERIFY_TOKEN = os.getenv("WHATSAPP_VERIFY_TOKEN", "mi_token_secreto")
DEFAULT_WHATSAPP_SUCURSAL = int(os.getenv("WHATSAPP_DEFAULT_SUCURSAL", "1"))

# --- Ejecutar migración para asegurar columnas (incluye sucursal_id, stock y kardex en crear_db.py) ---
from crear_db import main as migrate_db

app = Flask(__name__)
app.secret_key = os.getenv("APP_SECRET_KEY", "tu_clave_secreta_super_segura")
app.permanent_session_lifetime = timedelta(days=30)

# Asegurar estructura de BD al arrancar
migrate_db()

# ---- Config de subida de imágenes ----
HERRAMIENTAS_UPLOAD_FOLDER = os.path.join('static', 'img', 'herramientas')
EQUIPOS_UPLOAD_FOLDER      = os.path.join('static', 'img', 'equipos')
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
os.makedirs(HERRAMIENTAS_UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EQUIPOS_UPLOAD_FOLDER, exist_ok=True)

# Avatares de usuario
USER_UPLOAD_FOLDER = os.path.join('static', 'img', 'users')
os.makedirs(USER_UPLOAD_FOLDER, exist_ok=True)

# Estados válidos de un ticket (asistencia)
ESTADOS_VALIDOS = {'pendiente', 'en_progreso', 'resuelto', 'cancelado'}

# -------------------- Utilidades BD / Sucursal --------------------

def _try_geocode_server(q: str):
    """
    Geocodifica usando Nominatim si no vinieron lat/lng desde el formulario.
    Devuelve (lat, lng) o (None, None) si falla.
    """
    q = (q or "").strip()
    if not q:
        return None, None
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"format": "json", "q": f"{q} Paraguay"},
            headers={"User-Agent": "spynet-local/1.0"}
        )
        r.raise_for_status()
        data = r.json()
        if data:
            return float(data[0]["lat"]), float(data[0]["lon"])
    except Exception:
        pass
    return None, None


def _sucursal_channels(sucursal_id: int):
    """
    Devuelve (PHONE_NUMBER_ID, DESTINO) según la sucursal.
    Los valores se leen de variables de entorno (.env).
    """
    if sucursal_id == 1:  # Valenzuela
        return (
            os.getenv("WA_PHONE_ID_VALENZUELA"),   # el phone_number_id de Meta
            os.getenv("STAFF_WA_VALENZUELA")      # el número destino (ej: +595983399215)
        )
    elif sucursal_id == 2:  # Sapucai
        return (
            os.getenv("WA_PHONE_ID_SAPUCAI"),
            os.getenv("STAFF_WA_SAPUCAI")
        )
    else:
        return (None, None)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_db():
    db_path = os.path.join(app.root_path, "asistencias.db")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def col(row, key, default=None):
    try:
        if hasattr(row, "keys") and key in row.keys() and row[key] is not None:
            return row[key]
    except Exception:
        pass
    return default

def table_columns(conn, table):
    cur = conn.execute(f"PRAGMA table_info({table})")
    return {r[1] for r in cur.fetchall()}

def update_user_fields(conn, user_id, fields: dict):
    cols = table_columns(conn, "usuarios")
    data = {k: v for k, v in fields.items() if k in cols}
    if not data:
        return
    sets = ", ".join([f"{k}=?" for k in data.keys()])
    sql = f"UPDATE usuarios SET {sets}, updated_at=CURRENT_TIMESTAMP WHERE id=?"
    params = list(data.values()) + [user_id]
    conn.execute(sql, params)
    conn.commit()

def insert_row(conn, table, data: dict):
    cols = table_columns(conn, table)
    filt = {k: v for k, v in data.items() if k in cols}
    if not filt:
        return None
    placeholders = ", ".join(["?"] * len(filt))
    sql = f"INSERT INTO {table} ({', '.join(filt.keys())}) VALUES ({placeholders})"
    conn.execute(sql, list(filt.values()))
    conn.commit()
    try:
        rid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        return rid
    except Exception:
        return None

def get_current_sucursal():
    sid = session.get("sucursal_id", 1)
    sname = session.get("sucursal_nombre", "Sapucai")
    return int(sid), sname

def current_sucursal_id():
    return int(session.get("sucursal_id", 1))

def set_sucursal(sid: int):
    sid = int(sid)
    session["sucursal_id"] = sid
    try:
        conn = get_db()
        row = conn.execute("SELECT nombre FROM sucursales WHERE id=?", (sid,)).fetchone()
        conn.close()
        session["sucursal_nombre"] = row["nombre"] if row else ("Sapucai" if sid == 1 else f"Sucursal {sid}")
    except Exception:
        session["sucursal_nombre"] = "Sapucai"


# === NUEVO: asegurar columnas de clientes que usa el sistema ===
# insert_row() descarta en silencio las columnas que no existen; si faltaban
# 'cedula' o 'pppoe', esos datos se perdían al crear/editar clientes.
def _ensure_cliente_cols():
    try:
        conn = get_db()
        cols = table_columns(conn, "clientes")
        if cols:  # la tabla existe
            for c in ("cedula", "pppoe", "direccion"):
                if c not in cols:
                    conn.execute(f"ALTER TABLE clientes ADD COLUMN {c} TEXT")
            for c in ("lat", "lng"):
                if c not in cols:
                    conn.execute(f"ALTER TABLE clientes ADD COLUMN {c} REAL")
            # Instalación: equipo, contrato y vivienda (para retirar equipos si cancela)
            for c, tipo in (("onu_serie", "TEXT"), ("contrato_tipo", "TEXT"), ("contrato_meses", "INTEGER"),
                            ("vivienda", "TEXT"), ("dueno_nombre", "TEXT"), ("dueno_ci", "TEXT"),
                            ("dueno_telefono", "TEXT")):
                if c not in cols:
                    conn.execute(f"ALTER TABLE clientes ADD COLUMN {c} {tipo}")
            conn.commit()
        conn.close()
    except Exception as e:
        print("Aviso: no se pudieron verificar columnas de clientes:", e)

_ensure_cliente_cols()


# === Configuración: columnas de preferencias, técnicos y datos de la empresa ===
def _ensure_config():
    try:
        conn = get_db()
        ucols = table_columns(conn, "usuarios")
        if ucols:
            if "notifs" not in ucols:
                conn.execute("ALTER TABLE usuarios ADD COLUMN notifs INTEGER DEFAULT 1")
            if "notif_sonido" not in ucols:
                conn.execute("ALTER TABLE usuarios ADD COLUMN notif_sonido INTEGER DEFAULT 1")
            if "sucursal_pred" not in ucols:
                conn.execute("ALTER TABLE usuarios ADD COLUMN sucursal_pred INTEGER")
            if "activo" not in ucols:
                conn.execute("ALTER TABLE usuarios ADD COLUMN activo INTEGER DEFAULT 1")
        tcols = table_columns(conn, "tecnicos")
        if tcols:
            if "telefono" not in tcols:
                conn.execute("ALTER TABLE tecnicos ADD COLUMN telefono TEXT")
            if "activo" not in tcols:
                conn.execute("ALTER TABLE tecnicos ADD COLUMN activo INTEGER DEFAULT 1")
            if "sucursal_id" not in tcols:
                conn.execute("ALTER TABLE tecnicos ADD COLUMN sucursal_id INTEGER")
        # Tickets: dispositivos conectados al WiFi y repetidores
        acols = table_columns(conn, "asistencias")
        if acols:
            for c, tipo in (("dispositivos", "TEXT"), ("repetidor", "INTEGER"), ("repetidor_cant", "INTEGER")):
                if c not in acols:
                    conn.execute(f"ALTER TABLE asistencias ADD COLUMN {c} {tipo}")
        conn.execute("CREATE TABLE IF NOT EXISTS config (clave TEXT PRIMARY KEY, valor TEXT)")
        conn.commit()
        conn.close()
    except Exception as e:
        print("Aviso: no se pudo preparar la configuración:", e)

_ensure_config()

EMPRESA_CAMPOS = ["nombre_comercial", "razon_social", "ruc", "telefono", "email", "direccion", "pie_ot"]

def get_empresa(conn=None):
    """Datos de la empresa guardados en Configuración (para la OT)."""
    propia = conn is None
    conn = conn or get_db()
    try:
        filas = conn.execute("SELECT clave, valor FROM config WHERE clave LIKE 'empresa.%'").fetchall()
        datos = {r["clave"].split(".", 1)[1]: r["valor"] for r in filas}
    except sqlite3.Error:
        datos = {}
    finally:
        if propia:
            conn.close()
    return {k: (datos.get(k) or "") for k in EMPRESA_CAMPOS}


# === BLOQUE GEOLOCALIZACIÓN: Helpers ===
def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _get_tecnico_by_token(conn, token: str):
    if not token:
        return None
    return conn.execute(
        "SELECT id, sucursal_id FROM tecnicos WHERE tracking_token=?",
        (token,)
    ).fetchone()

def _insert_tecnico_pos(conn, tecnico_id: int, lat: float, lng: float, ts: str, sucursal_id: int):
    conn.execute(
        "INSERT INTO tecnico_pos (tecnico_id, lat, lng, ts, sucursal_id) VALUES (?,?,?,?,?)",
        (tecnico_id, float(lat), float(lng), ts, int(sucursal_id))
    )
    conn.execute(
        "UPDATE tecnicos SET lat=?, lng=?, pos_updated_at=?, sucursal_id=? WHERE id=?",
        (float(lat), float(lng), ts, int(sucursal_id), tecnico_id)
    )


@app.before_request
def inject_sucursal():
    g.sucursal_id, g.sucursal_nombre = get_current_sucursal()

@app.context_processor
def inject_centros():
    sid = current_sucursal_id()
    nombres = {1: "Sapucai", 2: "Valenzuela"}
    centros = {
        1: (-25.6789, -56.9497),     # 📍 Sapucai
        2: (-25.5949816, -56.873283) # 📍 Valenzuela (Google Maps)
    }
    return {
        "sucursal_id": sid,
        "sucursal_nombre": nombres.get(sid, ""),
        "sucursal_center": centros.get(sid, (-25.50, -57.10))
    }

@app.context_processor
def inject_google_maps():
    """Clave de Google Maps para los selectores de ubicación (vacía = mapa alternativo)."""
    return {"google_maps_key": os.getenv("GOOGLE_MAPS_API_KEY", "").strip(),
            # Dirección de la API de notificaciones (Node). En el VPS se define en .env
            "notif_api": os.getenv("NOTIF_API_URL", "http://localhost:4000").strip().rstrip("/")}

@app.context_processor
def inject_nav_fechas():
    """Fechas que usan los submenús de la barra lateral (agenda, estadísticas...)."""
    hoy = date.today()
    lunes = hoy - timedelta(days=hoy.weekday())
    mes_ant = hoy.replace(day=1) - timedelta(days=1)
    return {"nav": {
        "hoy": hoy.isoformat(),
        "manana": (hoy + timedelta(days=1)).isoformat(),
        "lunes": lunes.isoformat(),
        "domingo": (lunes + timedelta(days=6)).isoformat(),
        "mes": hoy.strftime("%Y-%m"),
        "mes_ant": mes_ant.strftime("%Y-%m"),
    }}

# -------------------- Importación CSV tolerante --------------------
def _norm_key(s: str) -> str:
    s = (s or "").strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_")

_HEADER_MAP = {
    "id": "external_id",
    "external_id": "external_id",
    "nombre": "nombre",
    "referencia": "referencia",
    "barrio": "barrio",
    "telefono": "telefono",
    "telefono_1": "telefono",
    "tel": "telefono",
    "tefono": "telefono",
    "celular": "telefono",
    "situacion": "situacion",
    "estado": "situacion",
    "exonerado": "exonerado",
    "exento": "exonerado",
    "tipo_valor": "tipo_valor",
    "tipo": "tipo",
    "valor": "valor",
    "plan": "valor",
    "vencimiento": "vencimiento",
    "fecha_vencimiento": "vencimiento",
    # --- NUEVO: cédula, PPPoE y dirección ---
    "cedula": "cedula",
    "ci": "cedula",
    "c_i": "cedula",
    "documento": "cedula",
    "nro_cedula": "cedula",
    "numero_cedula": "cedula",
    "pppoe": "pppoe",
    "usuario_pppoe": "pppoe",
    "user_pppoe": "pppoe",
    "direccion": "direccion",
    "domicilio": "direccion",
}

def _parse_bool(v):
    s = (str(v or "")).strip().lower()
    return 1 if s in ("1","si","sí","true","verdadero","x","s","y","yes") else 0

def _parse_date_to_iso(v):
    s = (str(v or "")).strip()
    if not s: return None
    m = re.match(r"^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$", s)
    if m:
        d, mth, y = m.groups()
        y = "20"+y if len(y)==2 else y
        try:
            return date(int(y), int(mth), int(d)).isoformat()
        except Exception:
            return s
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return s
    return s

def _only_digits(s):
    return re.sub(r"\D+", "", str(s or ""))

def _try_decode(file_storage):
    data = file_storage.read()
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return data.decode(enc), enc
        except Exception:
            continue
    return data.decode("latin-1", errors="ignore"), "latin-1"

def _guess_delimiter(text):
    try:
        dialect = csv.Sniffer().sniff(text[:1024], delimiters=[",",";","|","\t"])
        return dialect.delimiter
    except Exception:
        head = text.splitlines()[0] if text.splitlines() else ""
        return ";" if head.count(";") > head.count(",") else ","

def _split_tipo_valor(raw_tipo_o_tv, raw_valor):
    tipo = (raw_tipo_o_tv or "").strip().lower()
    if not tipo:
        tipo = "cliente"
    valor = (raw_valor or "").strip()
    if not valor and raw_tipo_o_tv:
        m = re.search(r"(\d[\d\.\,]*)", str(raw_tipo_o_tv))
        if m:
            valor = m.group(1).strip()
    return tipo, (valor or None)

# ===========================
#  AUTH: Registro / Forgot
# ===========================
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        usuario = request.form.get("usuario","").strip()
        contrasena = request.form.get("contrasena","")
        email = request.form.get("email","").strip()
        nombre = request.form.get("nombre","").strip()

        if not usuario or not contrasena:
            flash("Usuario y contraseña son obligatorios", "danger")
            return render_template("register.html")

        db = get_db()
        cols = table_columns(db, "usuarios")
        # Chequeo de duplicado tolerante
        if "sucursal_id" in cols:
            exists = db.execute("SELECT 1 FROM usuarios WHERE usuario=? AND sucursal_id=?",
                                (usuario, current_sucursal_id())).fetchone()
        else:
            exists = db.execute("SELECT 1 FROM usuarios WHERE usuario=?",
                                (usuario,)).fetchone()
        if exists:
            db.close()
            flash("Ese usuario ya existe" + (" en la sucursal" if "sucursal_id" in cols else ""), "danger")
            return render_template("register.html")

        if "password_hash" in cols:
            pwd_hash = generate_password_hash(contrasena)
            insert_row(db, "usuarios", {
                "usuario": usuario,
                "password_hash": pwd_hash,
                "email": email or None,
                "nombre": nombre or usuario,
                "rol": "operador",
                "foto_url": None,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "sucursal_id": current_sucursal_id(),
            })
        else:
            insert_row(db, "usuarios", {
                "usuario": usuario, "contrasena": contrasena,
                "sucursal_id": current_sucursal_id()
            })
        db.close()
        flash("Usuario registrado. Ya podés iniciar sesión.", "success")
        return redirect(url_for("login"))

    return render_template("register.html")

@app.route("/forgot_password", methods=["GET","POST"])
def forgot_password():
    if request.method == "POST":
        usuario = request.form.get("usuario","").strip()
        nueva = request.form.get("nueva","")
        confirmar = request.form.get("confirmar","")
        if not usuario or not nueva or not confirmar:
            flash("Completá todos los campos","danger")
            return render_template("forgot_password.html")
        if nueva != confirmar:
            flash("Las contraseñas no coinciden","danger")
            return render_template("forgot_password.html")

        db = get_db()
        cols = table_columns(db, "usuarios")
        # Select tolerante a ausencia de sucursal_id
        if "sucursal_id" in cols:
            user = db.execute("SELECT * FROM usuarios WHERE usuario=? AND sucursal_id=?",
                              (usuario, current_sucursal_id())).fetchone()
        else:
            user = db.execute("SELECT * FROM usuarios WHERE usuario=?",
                              (usuario,)).fetchone()

        if not user:
            db.close()
            flash("El usuario no existe" + (" en esta sucursal" if "sucursal_id" in cols else ""), "danger")
            return render_template("forgot_password.html")

        if "password_hash" in cols:
            pwd_hash = generate_password_hash(nueva)
            update_user_fields(db, col(user, "id"), {"password_hash": pwd_hash, "contrasena": None})
        else:
            update_user_fields(db, col(user, "id"), {"contrasena": nueva})

        flash("Contraseña actualizada. Iniciá sesión.","success")
        return redirect(url_for("login"))

    return render_template("forgot_password.html")

# ===========================
#  LOGIN (ÚNICO) + alias /login
# ===========================
@app.route("/", methods=["GET", "POST"], endpoint="login")
def login_view():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        contrasena = request.form.get("contrasena", "")

        db = get_db()
        # Traemos todas las filas con ese usuario (por si hay varias sucursales)
        users = db.execute("SELECT * FROM usuarios WHERE usuario = ?", (usuario,)).fetchall()
        db.close()

        winner = None
        desactivado = False
        for user in users:
            if col(user, "activo", 1) == 0:
                desactivado = True
                continue
            pwd_hash = col(user, "password_hash")
            plano    = col(user, "contrasena")
            ok = False
            if pwd_hash:
                ok = check_password_hash(pwd_hash, contrasena)
            elif plano is not None:
                ok = (contrasena == plano)
            if ok:
                winner = user
                break

        if winner:
            session.clear()
            session["usuario_id"] = col(winner, "id")
            session["usuario"]    = col(winner, "usuario")
            session["nombre"]     = col(winner, "nombre", col(winner, "usuario"))
            session["email"]      = col(winner, "email")
            session["rol"]        = col(winner, "rol", "operador")

            foto_rel = col(winner, "foto_url")
            session["foto_url"] = url_for('static', filename=foto_rel) if foto_rel else None

            sid = col(winner, "sucursal_id", 1)
            set_sucursal(sid)

            # Preferencias del usuario
            session["notifs"]       = 0 if col(winner, "notifs", 1) == 0 else 1
            session["notif_sonido"] = 0 if col(winner, "notif_sonido", 1) == 0 else 1

            session.permanent = bool(request.form.get("recordarme"))

            # Sucursal predeterminada: entrar directo sin elegir
            pred = col(winner, "sucursal_pred")
            if pred:
                conn = get_db()
                existe = conn.execute("SELECT 1 FROM sucursales WHERE id=?", (pred,)).fetchone()
                conn.close()
                if existe:
                    set_sucursal(pred)
                    return redirect(url_for("menu"))
            return redirect(url_for("seleccionar_sucursal"))
        elif desactivado and not any(col(u, "activo", 1) != 0 for u in users):
            flash("Tu usuario está desactivado. Consulta con un administrador.", "warning")
        else:
            flash("Usuario o contraseña incorrectos", "danger")

    return render_template("login.html")

@app.route("/login", methods=["GET", "POST"], endpoint="login_alias")
def login_alias():
    return login_view()

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ===========================
#  Sucursales (listar/seleccionar/crear/editar/eliminar)
# ===========================
@app.route("/sucursales")
def sucursales_list():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    conn = get_db()
    rows = conn.execute("SELECT id, nombre, direccion, telefono FROM sucursales ORDER BY id").fetchall()
    conn.close()
    return render_template("sucursales.html", sucursales=rows,
                           actual=g.sucursal_nombre, actual_id=g.sucursal_id)

@app.route("/sucursales/set/<int:sid>")
def sucursales_set(sid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    conn = get_db()
    row = conn.execute("SELECT id, nombre FROM sucursales WHERE id=?", (sid,)).fetchone()
    conn.close()
    if not row:
        flash("Sucursal no encontrada", "warning")
        return redirect(url_for("sucursales_list"))

    session["sucursal_id"] = row["id"]
    session["sucursal_nombre"] = row["nombre"]
    flash(f"Sucursal actual: {row['nombre']}", "success")
    return redirect(url_for("menu"))

@app.route("/sucursales/nueva", methods=["POST"])
def sucursales_nueva():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    nombre = (request.form.get("nombre") or "").strip()
    direccion = (request.form.get("direccion") or "").strip()
    telefono = (request.form.get("telefono") or "").strip()
    if not nombre:
        flash("El nombre es obligatorio", "danger")
        return redirect(url_for("sucursales_list"))

    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO sucursales (nombre, direccion, telefono, created_at) VALUES (?, ?, ?, datetime('now'))",
            (nombre, direccion or None, telefono or None)
        )
        conn.commit()
        flash("Sucursal creada", "success")
    except sqlite3.IntegrityError:
        flash("Ya existe una sucursal con ese nombre", "warning")
    finally:
        conn.close()

    return redirect(url_for("sucursales_list"))

@app.route("/sucursales/<int:sid>/editar", methods=["POST"])
def sucursales_editar(sid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    nombre    = (request.form.get("nombre") or "").strip()
    direccion = (request.form.get("direccion") or "").strip() or None
    telefono  = (request.form.get("telefono") or "").strip() or None

    if not nombre:
        flash("El nombre de la sucursal es obligatorio.", "error")
        return redirect(url_for("sucursales_list"))

    conn = get_db()
    try:
        conn.execute(
            "UPDATE sucursales SET nombre=?, direccion=?, telefono=? WHERE id=?",
            (nombre, direccion, telefono, sid)
        )
        conn.commit()
        # Si se editó la sucursal en uso, actualizar el nombre visible
        if current_sucursal_id() == sid:
            session["sucursal_nombre"] = nombre
        flash(f'Sucursal "{nombre}" actualizada.', "success")
    except sqlite3.IntegrityError:
        flash("Ya existe una sucursal con ese nombre.", "warning")
    finally:
        conn.close()

    return redirect(url_for("sucursales_list"))

@app.route("/sucursales/<int:sid>/eliminar", methods=["POST"])
def sucursales_eliminar(sid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    # No permitir borrar la sucursal que se está usando
    if current_sucursal_id() == sid:
        flash("No puedes eliminar la sucursal en uso. Cambia de sucursal primero.", "warning")
        return redirect(url_for("sucursales_list"))

    conn = get_db()
    row = conn.execute("SELECT nombre FROM sucursales WHERE id=?", (sid,)).fetchone()
    if not row:
        conn.close()
        flash("La sucursal ya no existe.", "warning")
        return redirect(url_for("sucursales_list"))

    # Evitar dejar tickets, clientes o equipos huérfanos
    en_uso = 0
    for tabla in ("asistencias", "clientes", "equipos"):
        try:
            if "sucursal_id" in table_columns(conn, tabla):
                en_uso += conn.execute(f"SELECT COUNT(*) FROM {tabla} WHERE sucursal_id=?", (sid,)).fetchone()[0]
        except sqlite3.OperationalError:
            pass
    if en_uso:
        conn.close()
        flash(f'No se puede eliminar "{row["nombre"]}": tiene {en_uso} registros asociados (tickets, clientes o equipos).', "error")
        return redirect(url_for("sucursales_list"))

    try:
        conn.execute("DELETE FROM sucursales WHERE id=?", (sid,))
        conn.commit()
        flash(f'Sucursal "{row["nombre"]}" eliminada.', "success")
    except sqlite3.Error:
        conn.rollback()
        flash("No se pudo eliminar la sucursal.", "error")
    finally:
        conn.close()

    return redirect(url_for("sucursales_list"))

@app.route("/seleccionar_sucursal", endpoint="seleccionar_sucursal")
def seleccionar_sucursal():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    conn = get_db()
    sucursales = conn.execute(
        "SELECT id, nombre, IFNULL(direccion,'') AS direccion, IFNULL(telefono,'') AS telefono "
        "FROM sucursales ORDER BY id"
    ).fetchall()
    conn.close()

    if len(sucursales) == 1:
        return redirect(url_for("sucursales_set", sid=sucursales[0]["id"]))

    return render_template("seleccionar_sucursal.html",
                           sucursales=sucursales,
                           actual=g.sucursal_nombre)

# ===========================
#  Perfil de usuario
# ===========================
@app.route("/perfil", methods=["GET", "POST"])
def perfil():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    user = db.execute("SELECT * FROM usuarios WHERE id=?", (session["usuario_id"],)).fetchone()
    if not user:
        db.close()
        flash("Usuario no encontrado.", "danger")
        return redirect(url_for("menu"))

    if request.method == "POST":
        nombre   = (request.form.get("nombre") or "").strip()
        email    = (request.form.get("email") or "").strip()
        telefono = (request.form.get("telefono") or "").strip()
        area     = (request.form.get("area") or "").strip()
        turno    = (request.form.get("turno") or "").strip()
        # Las preferencias (notificaciones, sonido, sucursal) se guardan en Configuración → Mis preferencias

        if not nombre:
            db.close()
            flash("El nombre es obligatorio.", "warning")
            return redirect(url_for("perfil"))

        rol = user["rol"]
        if session.get("rol") == "admin":
            rol = (request.form.get("rol") or rol).strip()
            if rol != "admin" and user["rol"] == "admin":
                otros = db.execute("SELECT COUNT(*) FROM usuarios WHERE rol='admin' AND IFNULL(activo,1)=1 AND id<>?",
                                   (user["id"],)).fetchone()[0]
                if otros == 0:
                    rol = "admin"
                    flash("Sigues como administrador: debe quedar al menos uno en el sistema.", "warning")

        update_user_fields(db, user["id"], {
            "nombre": nombre,
            "email": email or None,
            "telefono": telefono or None,
            "area": area or None,
            "turno": turno or None,
            "rol": rol
        })

        session["nombre"] = nombre or session.get("usuario")
        session["email"]  = email or None
        session["rol"]    = rol

        db.close()
        flash("Perfil actualizado.", "success")
        return redirect(url_for("perfil"))

    ctx = dict(user=dict(user))
    db.close()
    return render_template("perfil.html", **ctx)

@app.route("/perfil/password", methods=["POST"])
def perfil_password():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    actual  = request.form.get("actual", "")
    nueva   = request.form.get("nueva", "")
    repetir = request.form.get("repetir", "")

    if not nueva or nueva != repetir:
        flash("Las contraseñas no coinciden.", "warning")
        return redirect(url_for("perfil"))
    if len(nueva) < 6:
        flash("La contraseña nueva debe tener al menos 6 caracteres.", "warning")
        return redirect(url_for("perfil"))

    db = get_db()
    user = db.execute("SELECT * FROM usuarios WHERE id=?", (session["usuario_id"],)).fetchone()
    if not user:
        db.close(); flash("Usuario no encontrado.", "danger"); return redirect(url_for("perfil"))

    ok = False
    if user["password_hash"]:
        ok = check_password_hash(user["password_hash"], actual)
    elif user["contrasena"] is not None:
        ok = (actual == user["contrasena"])

    if not ok:
        db.close()
        flash("La contraseña actual es incorrecta.", "danger")
        return redirect(url_for("perfil"))

    cols = table_columns(db, "usuarios")
    if "password_hash" in cols:
        ph = generate_password_hash(nueva)
        update_user_fields(db, user["id"], {"password_hash": ph, "contrasena": None})
    else:
        update_user_fields(db, user["id"], {"contrasena": nueva})

    db.close()
    flash("Contraseña actualizada.", "success")
    return redirect(url_for("perfil"))

@app.route("/perfil/avatar", methods=["POST"])
def perfil_avatar():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    file = request.files.get("avatar")
    if not file or file.filename == "":
        flash("Seleccioná una imagen.", "warning")
        return redirect(url_for("perfil"))

    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in {"png","jpg","jpeg","gif"}:
        flash("Formato no permitido. Usa JPG/PNG/GIF.", "danger")
        return redirect(url_for("perfil"))

    base = secure_filename(f"user_{session['usuario_id']}_{int(datetime.now().timestamp())}.{ext}")
    save_path = os.path.join(USER_UPLOAD_FOLDER, base)
    file.save(save_path)

    rel = os.path.join("img", "users", base).replace("\\", "/")

    db = get_db()
    update_user_fields(db, session["usuario_id"], {"foto_url": rel})
    db.close()

    session["foto_url"] = url_for("static", filename=rel)
    flash("Foto actualizada.", "success")
    return redirect(url_for("perfil"))

# ===========================
#  Páginas principales
# ===========================
@app.route("/menu")
def menu():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    sid = current_sucursal_id()
    hoy = date.today().isoformat()
    acols = table_columns(db, "asistencias")
    ccols = table_columns(db, "clientes")
    ecols = table_columns(db, "equipos")

    def uno(sql, params=()):
        try:
            v = db.execute(sql, params).fetchone()[0]
            return v or 0
        except sqlite3.Error:
            return 0

    est = "lower(IFNULL(estado,'pendiente'))" if "estado" in acols else "'pendiente'"
    fecha_trabajo = "COALESCE(programada_en, fecha)" if "programada_en" in acols else "fecha"

    stats = {
        "pendientes": uno(f"SELECT COUNT(*) FROM asistencias WHERE sucursal_id=? AND {est}='pendiente'", (sid,)),
        "en_progreso": uno(f"SELECT COUNT(*) FROM asistencias WHERE sucursal_id=? AND {est}='en_progreso'", (sid,)),
        "inst_hoy": uno(f"""SELECT COUNT(*) FROM asistencias
                            WHERE sucursal_id=? AND lower(IFNULL(tipo,'')) LIKE 'instal%'
                              AND date({fecha_trabajo})=? AND {est}<>'cancelado'""", (sid, hoy)),
        "inst_hoy_hechas": uno(f"""SELECT COUNT(*) FROM asistencias
                            WHERE sucursal_id=? AND lower(IFNULL(tipo,'')) LIKE 'instal%'
                              AND date({fecha_trabajo})=? AND {est}='resuelto'""", (sid, hoy)),
        "equipos_stock": uno("SELECT COALESCE(SUM(stock),0) FROM equipos WHERE sucursal_id=?", (sid,)) if "stock" in ecols else 0,
        "equipos_tipos": uno("SELECT COUNT(*) FROM equipos WHERE sucursal_id=?", (sid,)) if ecols else 0,
        "clientes_total": uno("SELECT COUNT(*) FROM clientes WHERE sucursal_id=?", (sid,)) if ccols else 0,
        "clientes_activos": 0,
    }
    if ccols:
        sit = "lower(IFNULL(situacion,''))" if "situacion" in ccols else "''"
        act = "IFNULL(activo,1)" if "activo" in ccols else "1"
        stats["clientes_activos"] = uno(f"""
            SELECT COUNT(*) FROM clientes
             WHERE sucursal_id=? AND {act}=1
               AND NOT ({sit} LIKE '%cancel%' OR {sit} LIKE '%baja%' OR {sit} LIKE '%mora%' OR {sit} LIKE '%suspend%')
        """, (sid,))

    # Actividad reciente: últimos tickets de la sucursal
    try:
        recientes = [dict(r) for r in db.execute("""
            SELECT id, cliente, tipo, estado, fecha, tecnico
              FROM asistencias WHERE sucursal_id=?
             ORDER BY datetime(fecha) DESC LIMIT 6
        """, (sid,)).fetchall()]
    except sqlite3.Error:
        recientes = []
    db.close()

    return render_template("menu.html", stats=stats, recientes=recientes)

@app.route("/nuevo_ticket", methods=["GET", "POST"])
def nuevo_ticket():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    # ----- POST: guardar ticket -----
    if request.method == "POST":
        cliente        = (request.form.get("cliente") or "").strip()
        cliente_id     = request.form.get("cliente_id") or None
        direccion      = (request.form.get("direccion") or "").strip()
        tipo           = request.form.get("tipo") or "Soporte"
        prioridad      = request.form.get("prioridad") or "Media"
        tecnico_nombre = (request.form.get("tecnico") or "").strip()
        tecnico_id     = request.form.get("tecnico_id") or None
        problema       = (request.form.get("problema") or "").strip()
        cedula         = (request.form.get("cedula") or "").strip()
        pppoe          = (request.form.get("pppoe") or "").strip()
        canal          = (request.form.get("canal") or "web").strip()
        estado         = (request.form.get("estado") or "pendiente").strip()
        programada_local = (request.form.get("programada_local") or "").strip()
        programada_en  = programada_local.replace("T", " ") if programada_local else None

        if not cliente:
            flash("Selecciona un cliente.", "warning")
            return redirect(url_for("nuevo_ticket"))

        # Coordenadas del formulario
        lat = request.form.get("lat", type=float)
        lng = request.form.get("lng", type=float)

        # Si no llegaron coordenadas, intentamos geocodificar en el servidor
        if (lat is None or lng is None) and direccion:
            glat, glng = _try_geocode_server(direccion)
            if glat is not None and glng is not None:
                lat, lng = glat, glng

        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Dispositivos conectados al WiFi (TV, cámaras…) y repetidor
        dispositivos = [d.strip() for d in request.form.getlist("dispositivos") if d.strip()]
        otro = (request.form.get("dispositivo_otro") or "").strip()
        if "Otro" in dispositivos:
            dispositivos.remove("Otro")
            if otro:
                dispositivos.append(otro)
        rep = request.form.get("repetidor")
        repetidor = 1 if rep == "si" else (0 if rep == "no" else None)
        repetidor_cant = (request.form.get("repetidor_cant", type=int) or 1) if repetidor == 1 else None

        db = get_db()
        insert_row(db, "asistencias", {
            "dispositivos": ", ".join(dispositivos) or None,
            "repetidor": repetidor,
            "repetidor_cant": repetidor_cant,
            "cliente": cliente,
            "direccion": direccion,
            "tipo": tipo,
            "prioridad": prioridad,
            "tecnico": tecnico_nombre,
            "problema": problema,
            "fecha": fecha,
            "pppoe": pppoe,
            "cliente_id": cliente_id,
            "cedula": cedula,
            "programada_en": programada_en,
            "estado": estado,
            "canal": canal,
            "tecnico_id": tecnico_id,
            "sucursal_id": current_sucursal_id(),
            "lat": lat,
            "lng": lng,
        })
        db.close()

        flash("Ticket registrado.", "success")
        return redirect(url_for("tickets"))

    # ----- GET: mostrar formulario -----
    db = get_db()

    ccols = table_columns(db, "clientes")
    candidates = ["id","nombre","apellido","direccion","cedula","pppoe","telefono",
                  "barrio","referencia","tipo_valor","valor","plan","tipo","lat","lng"]
    select_cols = [c for c in candidates if c in ccols]
    if "id" not in select_cols: select_cols.insert(0, "id")
    if "nombre" not in select_cols: select_cols.insert(1, "nombre")

    sql = f"""
        SELECT {', '.join(select_cols)}
          FROM clientes
         WHERE activo=1 AND sucursal_id=?
      ORDER BY nombre COLLATE NOCASE ASC
    """
    clientes_rows = db.execute(sql, (current_sucursal_id(),)).fetchall()

    def _txt(v):
        return "" if v is None else str(v).strip()

    clientes = []
    for r in clientes_rows:
        rd = dict(r)
        # Plan: valor -> plan -> número dentro de tipo_valor
        valor = _txt(rd.get("valor")) or _txt(rd.get("plan"))
        if not valor:
            m = re.search(r"(\d[\d\.\,]*)", _txt(rd.get("tipo_valor")))
            valor = m.group(1) if m else ""
        clientes.append({
            "id": rd.get("id"),
            "nombre": _txt(rd.get("nombre")),
            "apellido": _txt(rd.get("apellido")),
            "direccion": _txt(rd.get("direccion")),
            "cedula": _txt(rd.get("cedula")),
            "pppoe": _txt(rd.get("pppoe")),
            "telefono": _txt(rd.get("telefono")),
            "barrio": _txt(rd.get("barrio")),
            "referencia": _txt(rd.get("referencia")),
            "tipo_valor": _txt(rd.get("tipo_valor")),
            "valor": valor,
            "plan": _txt(rd.get("plan")),
            "tipo": _txt(rd.get("tipo")) or "cliente",
            "lat": _txt(rd.get("lat")),
            "lng": _txt(rd.get("lng")),
        })

    tecnicos = db.execute("""
        SELECT id, nombre
          FROM tecnicos
         WHERE activo=1 AND sucursal_id=?
      ORDER BY nombre COLLATE NOCASE ASC
    """, (current_sucursal_id(),)).fetchall()
    db.close()

    return render_template("nuevo_ticket.html", clientes=clientes, tecnicos=tecnicos)

@app.route("/tickets")
def tickets():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    acols = table_columns(db, "asistencias")
    ccols = table_columns(db, "clientes")

    # Si existen las columnas necesarias, traemos el barrio desde clientes
    if "barrio" in ccols:
        if "cliente_id" in acols:
            # JOIN por cliente_id + fallback por nombre (case-insensitive)
            sql = """
                SELECT a.*,
                       COALESCE(c.barrio, c2.barrio) AS c_barrio
                  FROM asistencias a
             LEFT JOIN clientes c
                    ON c.id = a.cliente_id
                   AND c.sucursal_id = a.sucursal_id
             LEFT JOIN clientes c2
                    ON a.cliente_id IS NULL
                   AND lower(c2.nombre) = lower(a.cliente)
                   AND c2.sucursal_id = a.sucursal_id
                 WHERE a.sucursal_id = ?
              ORDER BY datetime(a.fecha) DESC
            """
            rows = db.execute(sql, (current_sucursal_id(),)).fetchall()
        else:
            # Sin cliente_id, solo fallback por nombre
            sql = """
                SELECT a.*,
                       c.barrio AS c_barrio
                  FROM asistencias a
             LEFT JOIN clientes c
                    ON lower(c.nombre) = lower(a.cliente)
                   AND c.sucursal_id = a.sucursal_id
                 WHERE a.sucursal_id = ?
              ORDER BY datetime(a.fecha) DESC
            """
            rows = db.execute(sql, (current_sucursal_id(),)).fetchall()
    else:
        # Sin columna barrio en clientes: select original
        rows = db.execute("""
            SELECT a.*
              FROM asistencias a
             WHERE a.sucursal_id = ?
          ORDER BY datetime(a.fecha) DESC
        """, (current_sucursal_id(),)).fetchall()

    data = [dict(r) for r in rows]
    db.close()
    return render_template("tickets.html", tickets=data)

# ===========================
#  Descargas (PDF/WORD)
# ===========================
@app.route("/descargar/pdf")
def descargar_pdf():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    rows = db.execute("""
        SELECT * FROM asistencias
        WHERE sucursal_id=?
        ORDER BY datetime(fecha) DESC
    """, (current_sucursal_id(),)).fetchall()
    tickets = [dict(row) for row in rows]
    db.close()

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, txt="Tickets de Asistencia", ln=True, align="C")

    for t in tickets:
        pdf.ln(4)
        pdf.set_font("Arial", size=11)
        pdf.multi_cell(0, 6, txt=(
            f"Cliente: {t.get('cliente','')}\n"
            f"Dirección: {t.get('direccion','')}\n"
            f"Técnico: {t.get('tecnico','')}\n"
            f"Tipo: {t.get('tipo','')}\n"
            f"Prioridad: {t.get('prioridad','')}\n"
            f"PPPoE: {t.get('pppoe','')}\n"
            f"Problema: {t.get('problema') or 'N/A'}\n"
            f"Fecha: {t.get('fecha','')}"
        ))

    pdf_bytes = pdf.output(dest='S').encode('latin-1', errors='replace')
    output = io.BytesIO(pdf_bytes)
    output.seek(0)
    return send_file(output, download_name="asistencias.pdf", as_attachment=True, mimetype="application/pdf")

@app.route("/descargar/word")
def descargar_word():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    rows = db.execute("""
        SELECT * FROM asistencias
        WHERE sucursal_id=?
        ORDER BY datetime(fecha) DESC
    """, (current_sucursal_id(),)).fetchall()
    tickets = [dict(row) for row in rows]
    db.close()

    doc = Document()
    doc.add_heading("Tickets de Asistencia", 0)

    for t in tickets:
        doc.add_paragraph(f"Cliente: {t.get('cliente','')}")
        doc.add_paragraph(f"Dirección: {t.get('direccion','')}")
        doc.add_paragraph(f"Técnico: {t.get('tecnico','')}")
        doc.add_paragraph(f"Tipo: {t.get('tipo','')}")
        doc.add_paragraph(f"Prioridad: {t.get('prioridad','')}")
        doc.add_paragraph(f"PPPoE: {t.get('pppoe','')}")
        doc.add_paragraph(f"Problema: {t.get('problema') or 'N/A'}")
        doc.add_paragraph(f"Fecha: {t.get('fecha','')}")
        doc.add_paragraph("")

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return send_file(
        output,
        download_name="asistencias.docx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ===========================
#  Mapa + Export PDF
# ===========================
@app.route("/mapa")
def mapa():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    db = get_db()
    tecnicos = db.execute(
        "SELECT id, nombre FROM tecnicos WHERE activo=1 AND sucursal_id=? ORDER BY nombre COLLATE NOCASE ASC",
        (current_sucursal_id(),)
    ).fetchall()
    db.close()
    return render_template("mapa.html", tecnicos=tecnicos)

# Helper para PDF del mapa
def _query_tickets_for_report(conn, desde, hasta, estado=None):
    params = [current_sucursal_id(), desde, hasta]
    sql = """
        SELECT cliente, direccion, tipo, prioridad, estado, tecnico, programada_en, lat, lng, fecha
        FROM asistencias
        WHERE sucursal_id=? AND date(fecha) BETWEEN ? AND ?
    """
    if estado:
        sql += " AND lower(coalesce(estado,'')) = lower(?)"
        params.append(estado)
    sql += " ORDER BY datetime(coalesce(programada_en, fecha)) ASC"
    rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]

@app.route("/mapa/export/pdf", methods=["GET"])
def mapa_export_pdf():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    desde = request.args.get("desde") or date.today().isoformat()
    hasta = request.args.get("hasta") or date.today().isoformat()
    estado = request.args.get("estado") or None

    conn = get_db()
    rows = _query_tickets_for_report(conn, desde, hasta, estado)
    conn.close()

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    # Logo opcional
    try:
        logo_path = os.path.join(app.static_folder, "img", "logo.png")
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=12, y=10, w=20)
    except Exception:
        pass

    # Encabezado
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, txt="Reporte de Tickets", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    sub = f"Rango: {desde} a {hasta}"
    if estado: sub += f" - Estado: {estado}"
    pdf.cell(0, 6, txt=sub, ln=1, align="C")
    pdf.ln(2)

    # Tabla
    pdf.set_font("Arial", "B", 9)
    headers = ["Fecha", "Cliente", "Dirección", "Tipo", "Prior.", "Estado", "Técnico"]
    widths  = [25, 35, 50, 22, 15, 22, 30]
    for h, w in zip(headers, widths):
        pdf.cell(w, 8, h, 1, 0, "C")
    pdf.ln(8)

    pdf.set_font("Arial", "", 9)
    for r in rows:
        cells = [
            (r.get("fecha") or "")[:16],
            (r.get("cliente") or "")[:22],
            (r.get("direccion") or "")[:34],
            (r.get("tipo") or "")[:12],
            (r.get("prioridad") or "")[:8],
            (r.get("estado") or "")[:12],
            (r.get("tecnico") or "")[:18],
        ]
        for txt, w in zip(cells, widths):
            pdf.cell(w, 7, txt, 1, 0)
        pdf.ln(7)

    out = pdf.output(dest="S").encode("latin-1", errors="replace")
    buf = io.BytesIO(out); buf.seek(0)
    fname = f"tickets_{desde}_a_{hasta}.pdf"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")

@app.route("/instalaciones")
def instalaciones():
    """Instalaciones = tickets cuyo tipo empieza con 'Instal'."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    vista = request.args.get("vista", "pendientes")  # pendientes | hoy | semana | realizadas | todas
    hoy = date.today()
    lunes = hoy - timedelta(days=hoy.weekday())
    domingo = lunes + timedelta(days=6)
    mes_ini = hoy.replace(day=1)

    db = get_db()
    sid = current_sucursal_id()
    acols = table_columns(db, "asistencias")
    ccols = table_columns(db, "clientes")

    est = "lower(IFNULL(a.estado,'pendiente'))" if "estado" in acols else "'pendiente'"
    fvis = "COALESCE(a.programada_en, a.fecha)" if "programada_en" in acols else "a.fecha"
    prog = "a.programada_en" if "programada_en" in acols else "NULL"
    base_where = "a.sucursal_id=? AND lower(IFNULL(a.tipo,'')) LIKE 'instal%'"

    filtros = {
        "pendientes": f"AND {est} IN ('pendiente','en_progreso')",
        "hoy":        f"AND date({fvis})=date('{hoy.isoformat()}') AND {est}<>'cancelado'",
        "semana":     f"AND date({fvis}) BETWEEN '{lunes.isoformat()}' AND '{domingo.isoformat()}' AND {est}<>'cancelado'",
        "realizadas": f"AND {est}='resuelto'",
        "todas":      "",
    }
    if vista not in filtros:
        vista = "pendientes"

    # Barrio / dirección del cliente (por cliente_id o por nombre)
    if "cliente_id" in acols:
        join_c = """LEFT JOIN clientes c ON c.id = a.cliente_id
                    LEFT JOIN clientes c2 ON a.cliente_id IS NULL AND lower(c2.nombre)=lower(a.cliente) AND c2.sucursal_id=a.sucursal_id"""
        barrio = "COALESCE(c.barrio, c2.barrio)" if "barrio" in ccols else "NULL"
        tel = "COALESCE(c.telefono, c2.telefono)" if "telefono" in ccols else "NULL"
    else:
        join_c = "LEFT JOIN clientes c ON lower(c.nombre)=lower(a.cliente) AND c.sucursal_id=a.sucursal_id"
        barrio = "c.barrio" if "barrio" in ccols else "NULL"
        tel = "c.telefono" if "telefono" in ccols else "NULL"

    # Pendientes: primero las programadas más próximas; el resto, lo más reciente primero
    orden = (f"CASE WHEN {prog} IS NULL THEN 1 ELSE 0 END, datetime({prog}) ASC, datetime(a.fecha) ASC"
             if vista in ("pendientes", "hoy", "semana") else f"datetime({fvis}) DESC")

    rows = db.execute(f"""
        SELECT a.*, {barrio} AS c_barrio, {tel} AS c_telefono, {prog} AS visita
          FROM asistencias a
          {join_c}
         WHERE {base_where} {filtros[vista]}
         ORDER BY {orden}
         LIMIT 300
    """, (sid,)).fetchall()
    instalaciones_rows = [dict(r) for r in rows]

    def contar(extra):
        return db.execute(f"SELECT COUNT(*) FROM asistencias a WHERE {base_where} {extra}", (sid,)).fetchone()[0]

    cuentas = {k: contar(v) for k, v in filtros.items()}
    kpis = {
        "pendientes": cuentas["pendientes"],
        "hoy": cuentas["hoy"],
        "semana": cuentas["semana"],
        "mes_realizadas": contar(f"AND {est}='resuelto' AND date(a.fecha) BETWEEN '{mes_ini.isoformat()}' AND '{hoy.isoformat()}'"),
        "sin_programar": contar(f"AND {est} IN ('pendiente','en_progreso') AND {prog} IS NULL"),
        "sin_tecnico": contar(f"AND {est} IN ('pendiente','en_progreso') AND trim(IFNULL(a.tecnico,''))=''"),
    }

    try:
        tecnicos = db.execute(
            "SELECT id, nombre FROM tecnicos WHERE activo=1 AND sucursal_id=? ORDER BY nombre COLLATE NOCASE",
            (sid,)
        ).fetchall()
    except sqlite3.OperationalError:
        tecnicos = []
    db.close()

    return render_template("instalaciones.html",
                           instalaciones=instalaciones_rows, vista=vista, cuentas=cuentas,
                           kpis=kpis, tecnicos=tecnicos, hoy=hoy.isoformat(),
                           semana_txt=f"{lunes:%d/%m} al {domingo:%d/%m}",
                           tiene_programacion=("programada_en" in acols),
                           tiene_tecnico_id=("tecnico_id" in acols))


@app.route("/instalaciones/nueva", methods=["GET", "POST"])
def instalaciones_nueva():
    """Registrar una instalación: con un cliente nuevo o uno existente."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    sid = current_sucursal_id()

    if request.method == "POST":
        f = request.form
        modo = f.get("modo", "nuevo")
        db = get_db()
        ccols = table_columns(db, "clientes")

        lat = f.get("lat", type=float)
        lng = f.get("lng", type=float)

        servicio = _servicio_form_data(f)
        if servicio["vivienda"] == "alquilada" and not (servicio["dueno_nombre"] and servicio["dueno_telefono"]):
            db.close()
            flash("Si la casa es alquilada, carga el nombre y el teléfono del dueño.", "warning")
            return redirect(url_for("instalaciones_nueva"))
        if servicio["contrato_tipo"] == "temporal" and not servicio["contrato_meses"]:
            db.close()
            flash("Indica por cuántos meses es el servicio.", "warning")
            return redirect(url_for("instalaciones_nueva"))

        if modo == "existente":
            cid = f.get("cliente_id", type=int)
            row = db.execute("SELECT * FROM clientes WHERE id=? AND sucursal_id=?", (cid, sid)).fetchone() if cid else None
            if not row:
                db.close()
                flash("Selecciona un cliente.", "warning")
                return redirect(url_for("instalaciones_nueva"))
            cli = dict(row)
            direccion = (f.get("direccion_existente") or cli.get("direccion") or "").strip()
            # Guardar la ubicación en el cliente si todavía no la tenía
            if lat is not None and lng is not None and "lat" in ccols and not cli.get("lat"):
                db.execute("UPDATE clientes SET lat=?, lng=? WHERE id=?", (lat, lng, cid))
            if direccion and "direccion" in ccols and not cli.get("direccion"):
                db.execute("UPDATE clientes SET direccion=? WHERE id=?", (direccion, cid))
            # Equipo, contrato y vivienda: se actualizan con lo cargado en esta instalación
            cambios = {k: v for k, v in servicio.items() if k in ccols and (v is not None or k.startswith("dueno_"))}
            if servicio["vivienda"] is None:
                cambios = {k: v for k, v in cambios.items() if not k.startswith("dueno_")}
            if servicio["contrato_tipo"] and "contrato_meses" in ccols:
                cambios["contrato_meses"] = servicio["contrato_meses"]   # indefinido: sin meses
            if cambios:
                sets = ", ".join(f"{k}=?" for k in cambios)
                db.execute(f"UPDATE clientes SET {sets} WHERE id=?", list(cambios.values()) + [cid])
                cli.update(cambios)
        else:
            nombre = (f.get("nombre") or "").strip()
            telefono = (f.get("telefono") or "").strip()
            cedula = (f.get("cedula") or "").strip()
            if not nombre or not telefono:
                db.close()
                flash("Para un cliente nuevo, el nombre y el teléfono son obligatorios.", "warning")
                return redirect(url_for("instalaciones_nueva"))

            existente = None
            if cedula and "cedula" in ccols:
                existente = db.execute("SELECT * FROM clientes WHERE cedula=? AND sucursal_id=?", (cedula, sid)).fetchone()
            direccion = (f.get("direccion") or "").strip()

            if existente:
                cli = dict(existente)
                flash(f'Ya existía un cliente con la CI {cedula} ({cli["nombre"]}); la instalación se registró a su nombre.', "info")
            else:
                valor = (f.get("valor") or "").strip() or None
                data = {
                    "nombre": nombre,
                    "cedula": cedula or None,
                    "telefono": telefono,
                    "barrio": (f.get("barrio") or "").strip() or None,
                    "direccion": direccion or None,
                    "referencia": (f.get("referencia") or "").strip() or None,
                    "pppoe": (f.get("pppoe") or "").strip() or None,
                    "vencimiento": (f.get("vencimiento") or "").strip() or None,
                    "situacion": "pendiente de instalación",
                    "activo": 0,
                    "exonerado": 0,
                    "sucursal_id": sid,
                    "lat": lat, "lng": lng,
                    "tipo": "cliente",
                    "valor": valor,
                    **servicio,
                }
                if "tipo_valor" in ccols and ("tipo" not in ccols or "valor" not in ccols):
                    data["tipo_valor"] = f"cliente {valor}" if valor else "cliente"
                nuevo_id = insert_row(db, "clientes", data)
                cli = dict(db.execute("SELECT * FROM clientes WHERE id=?", (nuevo_id,)).fetchone())

        # Si no se marcó en el mapa, intentar ubicar la dirección
        if (lat is None or lng is None):
            lat, lng = cli.get("lat"), cli.get("lng")
        if (lat is None or lng is None) and direccion:
            lat, lng = _try_geocode_server(" ".join(x for x in [direccion, cli.get("barrio") or ""] if x))

        # Descripción del trabajo: equipos a llevar + observaciones
        equipos = f.getlist("equipos")
        metros = (f.get("metros") or "").strip()
        partes = ["Instalación nueva."]
        if equipos:
            partes.append("Equipos: " + ", ".join(equipos) + ".")
        if metros:
            partes.append(f"Cable de acometida: {metros} m.")
        if f.get("observaciones", "").strip():
            partes.append(f.get("observaciones").strip())

        # Técnico
        tecnico_id = f.get("tecnico_id", type=int)
        tecnico_nombre = None
        if tecnico_id:
            t = db.execute("SELECT nombre FROM tecnicos WHERE id=? AND sucursal_id=?", (tecnico_id, sid)).fetchone()
            tecnico_nombre = t["nombre"] if t else None
            if not t:
                tecnico_id = None

        prog = (f.get("programada_local") or "").strip()
        nombre_completo = (f"{cli.get('nombre') or ''} {cli.get('apellido') or ''}").strip()

        tid = insert_row(db, "asistencias", {
            "cliente": nombre_completo,
            "cliente_id": cli.get("id"),
            "cedula": cli.get("cedula"),
            "pppoe": cli.get("pppoe"),
            "direccion": direccion,
            "tipo": "Instalación",
            "prioridad": None,
            "tecnico": tecnico_nombre,
            "tecnico_id": tecnico_id,
            "problema": " ".join(partes),
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "programada_en": prog.replace("T", " ") if prog else None,
            "estado": "pendiente",
            "canal": "web",
            "sucursal_id": sid,
            "lat": lat, "lng": lng,
        })
        db.commit()
        db.close()

        flash(f"Instalación registrada para {nombre_completo}.", "success")
        if f.get("despues") == "ot" and tid:
            return redirect(url_for("tickets_ot", tid=tid))
        return redirect(url_for("instalaciones"))

    # ----- GET -----
    db = get_db()
    ccols = table_columns(db, "clientes")
    campos = [c for c in ("id", "nombre", "apellido", "cedula", "telefono", "barrio", "direccion", "referencia", "lat", "lng")
              + SERVICIO_CAMPOS if c in ccols]
    clientes = [dict(r) for r in db.execute(
        f"SELECT {', '.join(campos)} FROM clientes WHERE sucursal_id=? ORDER BY nombre COLLATE NOCASE",
        (sid,)
    ).fetchall()]
    try:
        tecnicos = db.execute("SELECT id, nombre FROM tecnicos WHERE activo=1 AND sucursal_id=? ORDER BY nombre COLLATE NOCASE",
                              (sid,)).fetchall()
    except sqlite3.OperationalError:
        tecnicos = []
    db.close()

    centro = {1: (-25.6789, -56.9497), 2: (-25.5949816, -56.873283)}.get(sid, (-25.5949816, -56.873283))
    return render_template("instalacion_nueva.html", clientes=clientes, tecnicos=tecnicos,
                           cliente_pre=request.args.get("cliente_id", type=int), centro=centro)

# ===========================
#  Equipos (incluye stock y movimientos)
# ===========================
def ajustar_stock(conn, equipo_id, cantidad, tipo_mov, tecnico=None, motivo=""):
    # 1. Validar stock suficiente (solo lectura)
    row = conn.execute(
        "SELECT stock FROM equipos WHERE id=? AND sucursal_id=?",
        (equipo_id, current_sucursal_id())
    ).fetchone()
    if not row:
        return False, "Equipo no encontrado en esta sucursal"

    stock_actual = row["stock"] or 0
    nuevo = stock_actual + int(cantidad)
    if nuevo < 0:
        return False, f"Stock insuficiente (actual: {stock_actual})"

    # 2. Determinar tipo de movimiento
    tipo = "ingreso" if cantidad > 0 else "egreso"
    if tipo_mov in ("ingreso", "egreso"):
        tipo = tipo_mov

    # 3. Solo insertar movimiento (el trigger actualizará el stock)
    conn.execute("""
        INSERT INTO movimientos_equipos (equipo_id, tipo, cantidad, tecnico, motivo, sucursal_id)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (equipo_id, tipo, abs(int(cantidad)), tecnico, motivo or "", current_sucursal_id()))
    conn.commit()
    return True, nuevo

@app.route("/equipos")
def equipos():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    conn = get_db()
    sid = current_sucursal_id()

    equipos = conn.execute("""
        SELECT * FROM equipos
        WHERE sucursal_id=?
        ORDER BY nombre
    """, (sid,)).fetchall()

    herramientas = conn.execute("""
        SELECT * FROM herramientas
        WHERE sucursal_id=?
        ORDER BY nombre
    """, (sid,)).fetchall()

    tecnicos = conn.execute("""
        SELECT id, nombre
        FROM tecnicos
        WHERE activo=1
          AND sucursal_id=?
        ORDER BY nombre COLLATE NOCASE ASC
    """, (sid,)).fetchall()

    # ===========================
    # KPI DE EQUIPOS
    # ===========================

    total_equipos = conn.execute("""
        SELECT COUNT(*)
        FROM equipos
        WHERE sucursal_id=?
    """, (sid,)).fetchone()[0]

    hoy_str = date.today().isoformat()

    en_uso = conn.execute("""
        SELECT COALESCE(SUM(cantidad),0)
        FROM uso_items
        WHERE item_type='equipo'
          AND sucursal_id=?
          AND date(fecha)=?
    """, (sid, hoy_str)).fetchone()[0] or 0

    # Total de unidades disponibles en stock
    disponibles = conn.execute("""
        SELECT COALESCE(SUM(stock),0)
        FROM equipos
        WHERE sucursal_id=?
    """, (sid,)).fetchone()[0] or 0

    # ===========================
    # HISTORIAL
    # ===========================

    desde = request.args.get("desde") or hoy_str
    hasta = request.args.get("hasta") or hoy_str
    tecnico_id = request.args.get("tecnico_id") or None

    where = " AND date(u.fecha) BETWEEN ? AND ? "
    params = [sid, desde, hasta]

    if tecnico_id:
        tecnico = conn.execute("""
            SELECT nombre
            FROM tecnicos
            WHERE id=?
              AND sucursal_id=?
        """, (tecnico_id, sid)).fetchone()

        if tecnico:
            where += " AND u.tecnico=? "
            params.append(tecnico["nombre"])

    historial = conn.execute(f"""
        SELECT
            u.id,
            u.item_type,
            CASE
                WHEN u.item_type='herramienta' THEN h.nombre
                WHEN u.item_type='equipo' THEN e.nombre
                ELSE 'Desconocido'
            END AS nombre_item,
            u.tecnico,
            IFNULL(u.cantidad,1) AS cantidad,
            u.fecha,
            u.servicio
        FROM uso_items u
        LEFT JOIN herramientas h
            ON u.item_type='herramienta'
           AND u.item_id=h.id
        LEFT JOIN equipos e
            ON u.item_type='equipo'
           AND u.item_id=e.id
        WHERE u.sucursal_id=? {where}
        ORDER BY datetime(u.fecha) DESC
        LIMIT 500
    """, params).fetchall()

    conn.close()

    return render_template(
        "equipos.html",
        equipos=equipos,
        herramientas=herramientas,
        historial=historial,
        total=total_equipos,
        en_uso=int(en_uso),
        disponibles=int(disponibles),
        tecnicos=tecnicos,
        filtros={
            "desde": desde,
            "hasta": hasta,
            "tecnico_id": tecnico_id
        }
    )

@app.route('/registrar_equipo', methods=['POST'])
def registrar_equipo():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    nombre = request.form['nombre']
    tipo = request.form['tipo']
    descripcion = request.form.get('descripcion', '')
    stock = request.form.get('stock', type=int) or 0   # <-- lee el stock del formulario

    conn = get_db()
    conn.execute("""
        INSERT INTO equipos (nombre, tipo, descripcion, sucursal_id, stock)
        VALUES (?, ?, ?, ?, ?)
    """, (nombre, tipo, descripcion, current_sucursal_id(), stock))
    conn.commit()
    conn.close()

    flash("Equipo registrado correctamente.", "success")
    return redirect(url_for('equipos'))

# ---- Editar equipo ----
@app.route("/equipos/<int:eid>/editar", methods=["POST"])
def equipos_editar(eid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    nombre = (request.form.get("nombre") or "").strip()
    tipo   = (request.form.get("tipo") or "").strip()
    desc   = (request.form.get("descripcion") or "").strip()

    conn = get_db()
    conn.execute("""
        UPDATE equipos
           SET nombre=?, tipo=?, descripcion=?
         WHERE id=? AND sucursal_id=?
    """, (nombre, tipo, desc, eid, current_sucursal_id()))
    conn.commit(); conn.close()

    flash("Equipo actualizado correctamente.", "success")
    return redirect(url_for("equipos"))

# ---- Eliminar equipo ----
@app.route("/equipos/<int:eid>/eliminar", methods=["POST"])
def equipos_eliminar(eid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    conn = get_db()
    conn.execute("DELETE FROM equipos WHERE id=? AND sucursal_id=?", (eid, current_sucursal_id()))
    conn.commit(); conn.close()

    flash("Equipo eliminado.", "success")
    return redirect(url_for("equipos"))

@app.route('/registrar_uso_item', methods=['POST'])
def registrar_uso_item():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    item_id = request.form.get('item_id', type=int)
    tecnico = (request.form.get('tecnico') or session.get("usuario") or "").strip()
    servicio = (request.form.get('servicio') or "Uso de equipo").strip()
    cantidad = request.form.get('cantidad', type=int) or 1
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not item_id or cantidad <= 0:
        flash("Datos inválidos.", "danger")
        return redirect(url_for("equipos"))

    conn = get_db()
    sid = current_sucursal_id()

    # ¿Es equipo?
    equipo = conn.execute("""
        SELECT id
        FROM equipos
        WHERE id=? AND sucursal_id=?
    """, (item_id, sid)).fetchone()

    if equipo:
        # Descontar stock utilizando la función existente
        ok, res = ajustar_stock(conn, item_id, -cantidad, "egreso", tecnico, servicio)
        if not ok:
            conn.close()
            flash(res, "warning")
            return redirect(url_for("equipos"))
        item_type = "equipo"
    else:
        # ¿Es herramienta?
        herramienta = conn.execute("""
            SELECT id
            FROM herramientas
            WHERE id=? AND sucursal_id=?
        """, (item_id, sid)).fetchone()

        if not herramienta:
            conn.close()
            flash("Elemento no encontrado.", "danger")
            return redirect(url_for("equipos"))
        item_type = "herramienta"

    conn.execute("""
        INSERT INTO uso_items
        (item_type,item_id,tecnico,fecha,servicio,cantidad,sucursal_id)
        VALUES (?,?,?,?,?,?,?)
    """, (item_type, item_id, tecnico, fecha, servicio, cantidad, sid))

    conn.commit()
    conn.close()

    flash("Uso registrado correctamente.", "success")
    return redirect(url_for("equipos"))

@app.route("/equipos/ingresar_stock", methods=["POST"])
def equipos_ingresar_stock():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    equipo_id = request.form.get("equipo_id", type=int)
    cantidad  = request.form.get("cantidad", type=int)
    motivo    = (request.form.get("motivo") or "").strip()
    tecnico   = session.get("usuario")

    if not equipo_id or not cantidad or cantidad <= 0:
        flash("Completá equipo y una cantidad positiva.", "warning")
        return redirect(url_for("equipos"))

    conn = get_db()
    ok, res = ajustar_stock(conn, equipo_id, cantidad, "ingreso", tecnico, motivo or "Ingreso manual")
    conn.close()

    if ok:
        flash(f"Ingreso registrado. Stock nuevo: {res}", "success")
    else:
        flash(res, "danger")
    return redirect(url_for("equipos"))

# ---- Exportar / importar equipos (pasar el inventario entre la PC y el servidor) ----
EQUIPOS_CAMPOS_EXPORT = ("nombre", "tipo", "descripcion", "stock", "imagen")

@app.route("/equipos/exportar")
def equipos_exportar():
    """Descarga los equipos de la sucursal actual en un archivo .json."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    conn = get_db()
    campos = [c for c in EQUIPOS_CAMPOS_EXPORT if c in table_columns(conn, "equipos")]
    filas = conn.execute(
        f"SELECT {', '.join(campos)} FROM equipos WHERE sucursal_id=? ORDER BY nombre",
        (current_sucursal_id(),)
    ).fetchall()
    conn.close()

    datos = {"sucursal": g.sucursal_nombre, "equipos": [dict(f) for f in filas]}
    buf = BytesIO(json.dumps(datos, ensure_ascii=False, indent=2).encode("utf-8"))
    nombre = secure_filename(f"equipos_{g.sucursal_nombre}_{date.today().isoformat()}.json")
    return send_file(buf, as_attachment=True, download_name=nombre, mimetype="application/json")

@app.route("/equipos/importar", methods=["POST"])
def equipos_importar():
    """Agrega a la sucursal actual los equipos de un archivo exportado. No duplica por nombre."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    if session.get("rol") != "admin":
        flash("Solo un administrador puede importar equipos.", "warning")
        return redirect(url_for("equipos"))

    archivo = request.files.get("archivo")
    if not archivo or archivo.filename == "":
        flash("Seleccioná el archivo exportado (.json).", "warning")
        return redirect(url_for("equipos"))

    try:
        datos = json.loads(archivo.read().decode("utf-8-sig"))
        lista = datos.get("equipos", []) if isinstance(datos, dict) else datos
        if not isinstance(lista, list):
            raise ValueError
    except (ValueError, UnicodeDecodeError):
        flash("El archivo no es un export de equipos válido.", "danger")
        return redirect(url_for("equipos"))

    conn = get_db()
    sid = current_sucursal_id()
    cols = table_columns(conn, "equipos")
    agregados, salteados = 0, 0
    for e in lista:
        if not isinstance(e, dict) or not (e.get("nombre") or "").strip():
            continue
        nombre = e["nombre"].strip()
        existe = conn.execute(
            "SELECT 1 FROM equipos WHERE lower(nombre)=lower(?) AND sucursal_id=?", (nombre, sid)
        ).fetchone()
        if existe:
            salteados += 1
            continue
        fila = {k: e.get(k) for k in EQUIPOS_CAMPOS_EXPORT if k in cols}
        fila["nombre"] = nombre
        fila["stock"] = int(fila.get("stock") or 0) if "stock" in cols else None
        fila["sucursal_id"] = sid
        fila = {k: v for k, v in fila.items() if k in cols}
        conn.execute(f"INSERT INTO equipos ({', '.join(fila)}) VALUES ({', '.join('?' * len(fila))})",
                     list(fila.values()))
        agregados += 1
    conn.commit()
    conn.close()

    msg = f"Importación lista: {agregados} equipo(s) agregado(s) en {g.sucursal_nombre}."
    if salteados:
        msg += f" {salteados} ya existían y se saltearon."
    flash(msg, "success")
    return redirect(url_for("equipos"))

# ---- Subir imagen para HERRAMIENTA ----
@app.route('/subir_imagen_herramienta', methods=['POST'])
def subir_imagen_herramienta():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    if 'imagen' not in request.files:
        flash('No se ha seleccionado ningún archivo', 'danger')
        return redirect(url_for('equipos'))

    file = request.files['imagen']
    if file.filename == '':
        flash('No se ha seleccionado ningún archivo', 'danger')
        return redirect(url_for('equipos'))

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(HERRAMIENTAS_UPLOAD_FOLDER, filename)
        file.save(filepath)

        herramienta_id = request.form.get('herramienta_id')
        if herramienta_id:
            conn = get_db()
            cur = conn.execute('UPDATE herramientas SET imagen = ? WHERE id = ? AND sucursal_id=?',
                               (filename, herramienta_id, current_sucursal_id()))
            conn.commit()
            conn.close()
            if cur.rowcount:
                flash('Imagen subida correctamente', 'success')
            else:
                flash('Herramienta no encontrada en esta sucursal', 'danger')
        else:
            flash('No se indicó la herramienta', 'danger')
    else:
        flash('Formato de archivo no permitido', 'danger')

    return redirect(url_for('equipos'))

# ---- Subir imagen para EQUIPO ----
@app.route('/subir_imagen_equipo', methods=['POST'])
def subir_imagen_equipo():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    if 'imagen' not in request.files:
        flash('No se ha seleccionado ningún archivo', 'danger')
        return redirect(url_for('equipos'))

    file = request.files['imagen']
    if file.filename == '':
        flash('No se ha seleccionado ningún archivo', 'danger')
        return redirect(url_for('equipos'))

    if file and allowed_file(file.filename):
        # Nombre único por equipo y momento: evita que una foto pise a otra con el mismo nombre
        ext = file.filename.rsplit('.', 1)[1].lower()
        filename = f"equipo_{request.form.get('equipo_id')}_{int(datetime.now().timestamp())}.{ext}"
        filepath = os.path.join(EQUIPOS_UPLOAD_FOLDER, filename)
        file.save(filepath)

        equipo_id = request.form.get('equipo_id')
        if equipo_id:
            conn = get_db()
            cur = conn.execute('UPDATE equipos SET imagen = ? WHERE id = ? AND sucursal_id=?',
                               (filename, equipo_id, current_sucursal_id()))
            conn.commit()
            conn.close()
            if cur.rowcount:
                flash('Imagen de equipo subida correctamente', 'success')
            else:
                flash('Equipo no encontrado en esta sucursal', 'danger')
        else:
            flash('No se indicó el equipo', 'danger')
    else:
        flash('Formato de archivo no permitido', 'danger')

    return redirect(url_for('equipos'))

@app.route("/subir_foto_instalacion", methods=["POST"])
def subir_foto_instalacion():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    foto = request.files.get("foto")
    descripcion = request.form.get("descripcion", "")
    tecnico = session.get("usuario", "Desconocido")
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not foto or foto.filename == "":
        flash("No se seleccionó ninguna foto", "danger")
        return redirect(url_for("equipos"))

    carpeta_destino = HERRAMIENTAS_UPLOAD_FOLDER
    os.makedirs(carpeta_destino, exist_ok=True)

    nombre_archivo = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{secure_filename(foto.filename)}"
    ruta_guardado = os.path.join(carpeta_destino, nombre_archivo)
    foto.save(ruta_guardado)

    conn = get_db()
    conn.execute("""
        INSERT INTO fotos_asistencia (asistencia_id, tecnico, ruta_foto, descripcion, fecha, sucursal_id)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (None, tecnico, ruta_guardado, descripcion, fecha_actual, current_sucursal_id()))
    conn.commit()
    conn.close()

    flash("Foto subida correctamente.", "success")
    return redirect(url_for("equipos"))

# ---------- Reportes de uso_items (CSV/PDF) ----------
def _query_uso(conn, where_sql, params):
    sql = f"""
        SELECT u.item_type,
               CASE 
                 WHEN u.item_type='herramienta' THEN h.nombre
                 WHEN u.item_type='equipo'      THEN e.nombre
                 ELSE 'Desconocido'
               END AS nombre_item,
               u.tecnico,
               IFNULL(u.cantidad, 1) AS cantidad,
               IFNULL(u.servicio, '') AS servicio,
               u.fecha
        FROM uso_items u
        LEFT JOIN herramientas h ON u.item_type='herramienta' AND u.item_id=h.id
        LEFT JOIN equipos e      ON u.item_type='equipo'      AND u.item_id=e.id
        WHERE u.sucursal_id=? {where_sql}
        ORDER BY datetime(u.fecha) ASC
    """
    rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]

def _query_uso_hoy(conn):
    hoy = date.today().isoformat()
    rows = _query_uso(conn, "AND date(u.fecha)=?", [current_sucursal_id(), hoy])
    return hoy, rows

def _query_uso_rango(conn, desde, hasta, tecnico_id=None):
    where = "AND date(u.fecha) BETWEEN ? AND ?"
    params = [current_sucursal_id(), desde, hasta]
    if tecnico_id:
        tr = conn.execute("SELECT nombre FROM tecnicos WHERE id=? AND sucursal_id=?", (tecnico_id, current_sucursal_id())).fetchone()
        if tr and tr["nombre"]:
            where += " AND u.tecnico=?"
            params.append(tr["nombre"])
    rows = _query_uso(conn, where, params)
    return rows

def _uso_pdf(titulo, rows, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, txt=titulo, ln=True, align="C")

    pdf.set_font("Arial", "B", 10)
    headers = ["Fecha", "Tipo", "Ítem", "Técnico", "Cant.", "Servicio"]
    widths  = [32, 22, 52, 38, 14, 32]
    for h, w in zip(headers, widths): pdf.cell(w, 8, h, 1, 0, "C")
    pdf.ln(8)

    pdf.set_font("Arial", "", 10)
    for r in rows:
        cells = [
            r["fecha"] or "", r["item_type"] or "", (r["nombre_item"] or "")[:40],
            (r["tecnico"] or "")[:24], str(r["cantidad"]), (r["servicio"] or "")[:22],
        ]
        for txt, w in zip(cells, widths): pdf.cell(w, 8, txt, 1, 0)
        pdf.ln(8)

    out = pdf.output(dest="S").encode("latin-1", errors="replace")
    buf = io.BytesIO(out); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename, mimetype="application/pdf")

def _uso_csv(rows, filename):
    si = StringIO()
    writer = csv.writer(si)
    writer.writerow(["Fecha", "Tipo", "Ítem", "Técnico", "Cantidad", "Servicio"])
    for r in rows:
        writer.writerow([r["fecha"], r["item_type"], r["nombre_item"], r["tecnico"], r["cantidad"], r["servicio"]])
    buf = io.BytesIO(si.getvalue().encode("utf-8-sig"))
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename, mimetype="text/csv")

@app.route("/equipos/descargar/hoy.csv")
def equipos_descargar_hoy_csv():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    conn = get_db()
    hoy, rows = _query_uso_hoy(conn)
    conn.close()
    return _uso_csv(rows, f"uso_items_{hoy}.csv")

@app.route("/equipos/descargar/hoy.pdf")
def equipos_descargar_hoy_pdf():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    conn = get_db()
    hoy, rows = _query_uso_hoy(conn)
    conn.close()
    return _uso_pdf(f"Usos registrados - {hoy}", rows, f"uso_items_{hoy}.pdf")

@app.route("/equipos/descargar/rango.csv")
def equipos_descargar_rango_csv():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    desde = request.args.get("desde") or date.today().isoformat()
    hasta = request.args.get("hasta") or date.today().isoformat()
    tecnico_id = request.args.get("tecnico_id") or None

    conn = get_db()
    rows = _query_uso_rango(conn, desde, hasta, tecnico_id)
    conn.close()
    return _uso_csv(rows, f"uso_items_{desde}_a_{hasta}.csv")

@app.route("/equipos/descargar/rango.pdf")
def equipos_descargar_rango_pdf():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    desde = request.args.get("desde") or date.today().isoformat()
    hasta = request.args.get("hasta") or date.today().isoformat()
    tecnico_id = request.args.get("tecnico_id") or None

    conn = get_db()
    rows = _query_uso_rango(conn, desde, hasta, tecnico_id)
    conn.close()
    return _uso_pdf(f"Usos registrados - {desde} a {hasta}", rows, f"uso_items_{desde}_a_{hasta}.pdf")

# ---------- Kardex (movimientos de equipos) ----------
def _query_movs(conn, desde, hasta, equipo_id=None):
    sql = """
      SELECT m.fecha, e.nombre AS equipo, m.tipo, m.cantidad, IFNULL(m.tecnico,'') AS tecnico, IFNULL(m.motivo,'') AS motivo
      FROM movimientos_equipos m
      JOIN equipos e ON e.id = m.equipo_id
      WHERE m.sucursal_id=? AND date(m.fecha) BETWEEN ? AND ?
    """
    params = [current_sucursal_id(), desde, hasta]
    if equipo_id:
        sql += " AND m.equipo_id=?"
        params.append(equipo_id)
    sql += " ORDER BY datetime(m.fecha) ASC"
    return [dict(r) for r in conn.execute(sql, params).fetchall()]

@app.route("/equipos/movimientos.csv")
def equipos_movs_csv():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    desde = request.args.get("desde") or date.today().isoformat()
    hasta = request.args.get("hasta") or date.today().isoformat()
    equipo_id = request.args.get("equipo_id", type=int)

    conn = get_db(); rows = _query_movs(conn, desde, hasta, equipo_id); conn.close()
    si = StringIO(); w = csv.writer(si)
    w.writerow(["Fecha","Equipo","Tipo","Cantidad","Técnico","Motivo"])
    for r in rows: w.writerow([r["fecha"], r["equipo"], r["tipo"], r["cantidad"], r["tecnico"], r["motivo"]])
    buf = BytesIO(si.getvalue().encode("utf-8-sig")); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"movimientos_{desde}_a_{hasta}.csv", mimetype="text/csv")

@app.route("/equipos/movimientos.pdf")
def equipos_movs_pdf():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    desde = request.args.get("desde") or date.today().isoformat()
    hasta = request.args.get("hasta") or date.today().isoformat()
    equipo_id = request.args.get("equipo_id", type=int)

    conn = get_db(); rows = _query_movs(conn, desde, hasta, equipo_id); conn.close()
    pdf = FPDF(); pdf.add_page(); pdf.set_font("Arial","B",14)
    pdf.cell(0,10, f"Movimientos de equipos ({desde} a {hasta})", ln=1, align="C")
    pdf.set_font("Arial","B",10)
    headers = ["Fecha","Equipo","Tipo","Cant.","Técnico","Motivo"]; widths=[32,45,20,14,35,44]
    for h,w in zip(headers,widths): pdf.cell(w,8,h,1,0,"C")
    pdf.ln(8); pdf.set_font("Arial","",10)
    for r in rows:
        cells=[r["fecha"] or "", r["equipo"] or "", r["tipo"] or "", str(r["cantidad"]), (r["tecnico"] or "")[:20], (r["motivo"] or "")[:32]]
        for txt,w in zip(cells,widths): pdf.cell(w,7,txt,1,0)
        pdf.ln(7)
    out = pdf.output(dest="S").encode("latin-1","replace"); buf = BytesIO(out); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"movimientos_{desde}_a_{hasta}.pdf", mimetype="application/pdf")

# ===========================
#  Clientes
# ===========================
@app.route("/clientes")
def clientes():
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))

    q        = request.args.get("q","").strip()
    situ     = request.args.get("situacion","").strip()
    exo      = request.args.get("exonerado","")
    barrio_f = request.args.get("barrio","").strip()
    page     = request.args.get("page", 1, type=int)
    per_page = 25

    db = get_db()
    ccols = table_columns(db, "clientes")

    where = " WHERE sucursal_id=?"
    p = [current_sucursal_id()]
    if q:
        like = f"%{q}%"
        campos = ["nombre", "telefono", "referencia"] + [c for c in ("cedula", "pppoe") if c in ccols]
        where += " AND (" + " OR ".join(f"IFNULL({c},'') LIKE ?" for c in campos) + ")"
        p += [like] * len(campos)
    if situ:
        where += " AND IFNULL(situacion,'') LIKE ?"
        p += [f"%{situ}%"]
    if exo in ("0","1"):
        where += " AND exonerado=?"
        p += [int(exo)]
    if barrio_f:
        where += " AND IFNULL(barrio,'') LIKE ?"
        p += [f"%{barrio_f}%"]

    # total de registros que matchean el filtro (no el total absoluto)
    total = db.execute(f"SELECT COUNT(*) FROM clientes{where}", p).fetchone()[0]
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = max(1, min(page, total_pages))
    offset = (page - 1) * per_page

    sql = f"SELECT * FROM clientes{where} ORDER BY nombre COLLATE NOCASE ASC LIMIT ? OFFSET ?"
    rows = db.execute(sql, p + [per_page, offset]).fetchall()
    db.close()

    clientes_norm = []
    for r in rows:
        rd = dict(r)
        tipo = rd.get("tipo")
        valor = rd.get("valor")
        if not valor:
            tv = rd.get("tipo_valor") or ""
            m = re.search(r"(\d[\d\.\,]*)", tv)
            if m:
                valor = m.group(1)
        if not tipo:
            tipo = "cliente"
        rd["tipo"] = tipo
        rd["valor"] = valor
        clientes_norm.append(rd)

    return render_template("clientes.html",
                           clientes=clientes_norm,
                           q=q, situacion=situ, exonerado=exo, barrio=barrio_f,
                           total=total, page=page, total_pages=total_pages, per_page=per_page)

SERVICIO_CAMPOS = ("onu_serie", "contrato_tipo", "contrato_meses", "vivienda",
                   "dueno_nombre", "dueno_ci", "dueno_telefono")

def _servicio_form_data(form):
    """Equipo, contrato y vivienda. Si la casa es propia, no se guardan datos de dueño."""
    tipo = (form.get("contrato_tipo") or "").strip() or None
    meses = form.get("contrato_meses", type=int) if tipo == "temporal" else None
    vivienda = (form.get("vivienda") or "").strip() or None
    alquilada = vivienda == "alquilada"
    return {
        "onu_serie": (form.get("onu_serie") or "").strip().upper() or None,
        "contrato_tipo": tipo,
        "contrato_meses": meses,
        "vivienda": vivienda,
        "dueno_nombre": ((form.get("dueno_nombre") or "").strip() or None) if alquilada else None,
        "dueno_ci": ((form.get("dueno_ci") or "").strip() or None) if alquilada else None,
        "dueno_telefono": ((form.get("dueno_telefono") or "").strip() or None) if alquilada else None,
    }

def _cliente_form_data(cols):
    """Lee el formulario de cliente y devuelve solo las columnas que existen."""
    external_id = request.form.get("external_id","").strip()
    nombre      = request.form.get("nombre","").strip()
    referencia  = request.form.get("referencia","").strip()
    barrio      = request.form.get("barrio","").strip()
    telefono    = request.form.get("telefono","").strip()
    cedula      = request.form.get("cedula","").strip()
    pppoe       = request.form.get("pppoe","").strip()
    direccion   = request.form.get("direccion","").strip()
    situacion   = request.form.get("situacion","").strip()
    exonerado   = 1 if request.form.get("exonerado") == "on" else 0
    tipo        = request.form.get("tipo","").strip() or "cliente"
    valor       = request.form.get("valor","").strip() or None
    vencimiento = request.form.get("vencimiento","").strip()
    activo      = 1 if any(x in situacion.lower() for x in ["activo","al día","al dia","en servicio","ok"]) else 0

    data = {
        "external_id": external_id or None,
        "nombre": nombre,
        "referencia": referencia or None,
        "barrio": barrio or None,
        "telefono": telefono or None,
        "cedula": cedula or None,
        "pppoe": pppoe or None,
        "situacion": situacion or None,
        "exonerado": exonerado,
        "vencimiento": vencimiento or None,
        "activo": activo,
    }
    # La dirección solo se toca si el formulario la trae (evita borrarla)
    if "direccion" in request.form:
        data["direccion"] = direccion or None
    if "tipo" in cols: data["tipo"] = tipo
    if "valor" in cols: data["valor"] = valor
    if "tipo_valor" in cols and ("tipo" not in cols or "valor" not in cols):
        data["tipo_valor"] = f"{tipo} {valor}" if valor else tipo
    if "vivienda" in request.form:
        data.update(_servicio_form_data(request.form))

    return {k: v for k, v in data.items() if k in cols}

@app.route("/clientes/nuevo", methods=["GET","POST"])
def clientes_nuevo():
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))
    if request.method == "POST":
        db = get_db()
        cols = table_columns(db, "clientes")
        data = _cliente_form_data(cols)
        if not data.get("nombre"):
            db.close()
            flash("El nombre del cliente es obligatorio.", "warning")
            return redirect(url_for("clientes_nuevo"))
        data["sucursal_id"] = current_sucursal_id()
        insert_row(db, "clientes", data)
        db.close()
        flash("Cliente creado.", "success")
        return redirect(url_for("clientes"))
    return render_template("cliente_form.html", mode="new", cliente=None)

@app.route("/clientes/<int:cid>")
def clientes_detalle(cid):
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))

    db = get_db()
    c = db.execute("SELECT * FROM clientes WHERE id=? AND sucursal_id=?",
                   (cid, current_sucursal_id())).fetchone()
    if not c:
        db.close()
        flash("Cliente no encontrado.", "warning")
        return redirect(url_for("clientes"))

    try:
        tiene_apellido = ("apellido" in c.keys()) and bool(c["apellido"])
    except Exception:
        tiene_apellido = False
    full_name = f"{c['nombre']} {c['apellido']}".strip() if tiene_apellido else c["nombre"]

    acols = table_columns(db, "asistencias")
    if "cliente_id" in acols:
        tickets = db.execute("""
            SELECT * FROM asistencias
            WHERE (cliente_id = ? OR (cliente_id IS NULL AND cliente = ?)) AND sucursal_id=?
            ORDER BY datetime(fecha) DESC
            LIMIT 10
        """, (cid, full_name, current_sucursal_id())).fetchall()
    else:
        tickets = db.execute("""
            SELECT * FROM asistencias
            WHERE cliente = ? AND sucursal_id=?
            ORDER BY datetime(fecha) DESC
            LIMIT 10
        """, (full_name, current_sucursal_id())).fetchall()
    db.close()
    return render_template("cliente_detalle.html", c=c, tickets=tickets)

@app.route("/clientes/<int:cid>/editar", methods=["GET","POST"])
def clientes_editar(cid):
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))
    db = get_db()
    c = db.execute("SELECT * FROM clientes WHERE id=? AND sucursal_id=?",
                   (cid, current_sucursal_id())).fetchone()
    if not c:
        db.close(); flash("Cliente no encontrado.", "warning"); return redirect(url_for("clientes"))

    if request.method == "POST":
        cols = table_columns(db, "clientes")
        data = _cliente_form_data(cols)
        if not data.get("nombre"):
            db.close()
            flash("El nombre del cliente es obligatorio.", "warning")
            return redirect(url_for("clientes_editar", cid=cid))

        sets = ", ".join([f"{k}=?" for k in data.keys()])
        db.execute(f"UPDATE clientes SET {sets} WHERE id=? AND sucursal_id=?",
                   list(data.values())+[cid, current_sucursal_id()])
        db.commit(); db.close()
        flash("Cliente actualizado.", "success")
        return redirect(url_for("clientes"))
    db.close()
    return render_template("cliente_form.html", mode="edit", cliente=c)

@app.route("/clientes/<int:cid>/toggle", methods=["POST"])
def clientes_toggle(cid):
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))
    db = get_db()
    cur = db.execute("SELECT activo FROM clientes WHERE id=? AND sucursal_id=?",
                     (cid, current_sucursal_id())).fetchone()
    if cur:
        nuevo = 0 if cur["activo"]==1 else 1
        db.execute("UPDATE clientes SET activo=? WHERE id=? AND sucursal_id=?",
                   (nuevo, cid, current_sucursal_id()))
        db.commit()
        flash("Estado actualizado.", "success")
    db.close()
    return redirect(url_for("clientes"))

@app.route("/clientes/<int:cid>/eliminar", methods=["POST"])
def clientes_eliminar(cid):
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))
    db = get_db()
    db.execute("DELETE FROM clientes WHERE id=? AND sucursal_id=?",
               (cid, current_sucursal_id()))
    db.commit(); db.close()
    flash("Cliente eliminado.", "success")
    return redirect(url_for("clientes"))

# ===========================
#  Agenda + Acciones de tickets (POST)
# ===========================
@app.route("/agenda")
def agenda():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    # Filtros
    dia        = (request.args.get("dia") or "").strip()
    estado_f   = (request.args.get("estado") or "").strip()
    tecnico_f  = (request.args.get("tecnico_id") or "").strip()

    # Rango (opcional)
    desde_r    = (request.args.get("desde") or "").strip()
    hasta_r    = (request.args.get("hasta") or "").strip()
    range_mode = bool(desde_r and hasta_r)

    db = get_db()

    # Técnicos (para los selects)
    try:
        tecnicos = db.execute(
            "SELECT id, nombre FROM tecnicos WHERE activo=1 AND sucursal_id=? ORDER BY nombre",
            (current_sucursal_id(),)
        ).fetchall()
    except sqlite3.OperationalError:
        tecnicos = []

    acols = table_columns(db, "asistencias")
    ccols = table_columns(db, "clientes")

    if "programada_en" not in acols:
        db.close()
        flash("Tu base de datos no tiene la columna 'programada_en' en asistencias. Ejecutá crear_db.py para actualizar.", "warning")
        return render_template("agenda.html",
                               eventos=[],
                               dia=dia or date.today().isoformat(),
                               prev_dia=(datetime.strptime(dia or date.today().isoformat(), "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d"),
                               next_dia=(datetime.strptime(dia or date.today().isoformat(), "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d"),
                               estado_f=estado_f, tecnico_f=tecnico_f,
                               tecnicos=tecnicos,
                               range_mode=False, desde="", hasta="")

    # Joins opcionales
    if "cliente_id" in acols:
        join_c = " LEFT JOIN clientes c ON a.cliente_id = c.id "
        select_c = " , c.nombre AS c_nombre "
        select_c += (" , c.apellido AS c_apellido " if "apellido" in ccols else " , NULL AS c_apellido ")
    else:
        join_c = " "
        select_c = " , NULL AS c_nombre, NULL AS c_apellido "

    join_t = " LEFT JOIN tecnicos t ON a.tecnico_id = t.id " if "tecnico_id" in acols else " "
    select_t = " , t.nombre AS tecnico_nombre " if "tecnico_id" in acols else " , NULL AS tecnico_nombre "

    # --- MODO RANGO ---
    if range_mode:
        sql = f"""
          SELECT a.* {select_c} {select_t}
            FROM asistencias a
            {join_c}
            {join_t}
           WHERE a.sucursal_id=?
             AND date(COALESCE(a.programada_en, a.fecha)) BETWEEN ? AND ?
        """
        params = [current_sucursal_id(), desde_r, hasta_r]

    # --- MODO DÍA ---
    else:
        dia_use = dia or datetime.now().strftime("%Y-%m-%d")
        sql = f"""
          SELECT a.* {select_c} {select_t}
            FROM asistencias a
            {join_c}
            {join_t}
           WHERE a.sucursal_id=?
             AND (
                   (a.programada_en IS NOT NULL AND date(a.programada_en) = ?)
                OR (a.programada_en IS NULL AND date(a.fecha) = ? AND IFNULL(a.estado,'pendiente') = 'pendiente')
                 )
        """
        params = [current_sucursal_id(), dia_use, dia_use]

    # Filtros extra
    if estado_f and "estado" in acols:
        sql += " AND a.estado = ?"
        params.append(estado_f)
    if tecnico_f and "tecnico_id" in acols:
        sql += " AND a.tecnico_id = ?"
        params.append(tecnico_f)

    sql += " ORDER BY date(COALESCE(a.programada_en, a.fecha)) ASC, time(COALESCE(a.programada_en, a.fecha)) ASC"
    eventos = db.execute(sql, params).fetchall()
    db.close()

    # Navegación día a día si NO es rango
    if not range_mode:
        d = datetime.strptime(dia or datetime.now().strftime("%Y-%m-%d"), "%Y-%m-%d")
        prev_dia = (d - timedelta(days=1)).strftime("%Y-%m-%d")
        next_dia = (d + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        prev_dia = next_dia = ""

    return render_template(
        "agenda.html",
        eventos=eventos,
        dia=dia or datetime.now().strftime("%Y-%m-%d"),
        prev_dia=prev_dia,
        next_dia=next_dia,
        estado_f=estado_f,
        tecnico_f=tecnico_f,
        tecnicos=tecnicos,
        range_mode=range_mode,
        desde=desde_r, hasta=hasta_r
    )

@app.route("/tickets/<int:tid>/programar", methods=["POST"])
def tickets_programar(tid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    prog = (request.form.get("programada_local") or "").strip()
    programada_en = prog.replace("T", " ") if prog else None

    db = get_db()
    acols = table_columns(db, "asistencias")
    if "programada_en" not in acols:
        db.close()
        flash("No existe la columna 'programada_en' en asistencias. Actualizá la BD.", "warning")
        return redirect(request.referrer or url_for("agenda"))

    db.execute("UPDATE asistencias SET programada_en=? WHERE id=? AND sucursal_id=?",
               (programada_en, tid, current_sucursal_id()))
    db.commit()
    db.close()

    flash("Cita reprogramada.", "success")
    return redirect(request.referrer or url_for("agenda"))

@app.route("/tickets/<int:tid>/estado", methods=["POST"])
def tickets_cambiar_estado(tid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    nuevo = (request.form.get("estado") or "").strip()
    db = get_db()
    acols = table_columns(db, "asistencias")
    if "estado" not in acols:
        db.close()
        flash("No existe la columna 'estado' en asistencias. Actualizá la BD.", "warning")
        return redirect(request.referrer or url_for("agenda"))

    if nuevo not in ESTADOS_VALIDOS:
        db.close()
        flash("Estado inválido.", "warning")
        return redirect(request.referrer or url_for("agenda"))

    db.execute("UPDATE asistencias SET estado=? WHERE id=? AND sucursal_id=?",
               (nuevo, tid, current_sucursal_id()))

    # Instalación realizada: el cliente que esperaba la instalación pasa a activo
    if nuevo == "resuelto" and "cliente_id" in acols:
        t = db.execute("SELECT tipo, cliente_id FROM asistencias WHERE id=?", (tid,)).fetchone()
        if t and t["cliente_id"] and str(t["tipo"] or "").lower().startswith("instal"):
            ccols = table_columns(db, "clientes")
            if "situacion" in ccols:
                db.execute("""UPDATE clientes SET activo=1, situacion='activo'
                               WHERE id=? AND lower(IFNULL(situacion,'')) LIKE 'pendiente%'""",
                           (t["cliente_id"],))
    db.commit()
    db.close()

    flash("Estado actualizado.", "success")
    return redirect(request.referrer or url_for("agenda"))

# ---------- Editar ticket (modal de la lista de tickets) ----------
@app.route("/tickets/<int:tid>/editar", methods=["POST"])
def tickets_editar(tid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    tipo      = (request.form.get("tipo") or "").strip()
    prioridad = (request.form.get("prioridad") or "").strip() or None
    tecnico   = (request.form.get("tecnico") or "").strip() or None
    estado    = (request.form.get("estado") or "pendiente").strip().lower()

    if not tipo:
        flash("El tipo del ticket es obligatorio.", "error")
        return redirect(request.referrer or url_for("tickets"))
    if estado not in ESTADOS_VALIDOS:
        estado = "pendiente"

    db = get_db()
    sid = current_sucursal_id()
    acols = table_columns(db, "asistencias")

    data = {"tipo": tipo, "prioridad": prioridad, "tecnico": tecnico, "estado": estado}
    # Mantener tecnico_id sincronizado para que la Agenda muestre el técnico correcto
    if "tecnico_id" in acols:
        t = db.execute(
            "SELECT id FROM tecnicos WHERE lower(nombre)=lower(?) AND sucursal_id=?",
            (tecnico or "", sid)
        ).fetchone()
        data["tecnico_id"] = t["id"] if t else None
    data = {k: v for k, v in data.items() if k in acols}

    sets = ", ".join(f"{k}=?" for k in data)
    cur = db.execute(f"UPDATE asistencias SET {sets} WHERE id=? AND sucursal_id=?",
                     list(data.values()) + [tid, sid])
    db.commit()
    db.close()

    if cur.rowcount:
        flash("Ticket actualizado.", "success")
    else:
        flash("Ticket no encontrado en esta sucursal.", "warning")
    return redirect(request.referrer or url_for("tickets"))

@app.route("/tickets/<int:tid>/asignar", methods=["POST"])
def tickets_asignar(tid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    tecnico_id = request.form.get("tecnico_id") or None

    db = get_db()
    acols = table_columns(db, "asistencias")
    if "tecnico_id" not in acols:
        db.close()
        flash("No existe la columna 'tecnico_id' en asistencias. Actualizá la BD.", "warning")
        return redirect(request.referrer or url_for("agenda"))

    # Mantener también el nombre del técnico (lo usa la lista de tickets)
    nombre = None
    if tecnico_id:
        t = db.execute("SELECT nombre FROM tecnicos WHERE id=? AND sucursal_id=?",
                       (tecnico_id, current_sucursal_id())).fetchone()
        nombre = t["nombre"] if t else None

    if "tecnico" in acols:
        db.execute("UPDATE asistencias SET tecnico_id=?, tecnico=? WHERE id=? AND sucursal_id=?",
                   (tecnico_id, nombre, tid, current_sucursal_id()))
    else:
        db.execute("UPDATE asistencias SET tecnico_id=? WHERE id=? AND sucursal_id=?",
                   (tecnico_id, tid, current_sucursal_id()))
    db.commit()
    db.close()

    flash("Técnico asignado.", "success")
    return redirect(request.referrer or url_for("agenda"))

# ---------- Eliminar tickets ----------
@app.route("/tickets/<int:tid>/eliminar", methods=["POST"])
def tickets_eliminar(tid):
    """Elimina definitivamente un ticket SOLO si su estado es 'resuelto' en la sucursal actual."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    acols = table_columns(db, "asistencias")
    if "estado" not in acols:
        db.close()
        flash("No existe la columna 'estado' en asistencias. Actualizá la BD.", "warning")
        return redirect(request.referrer or url_for("tickets"))

    row = db.execute("SELECT estado FROM asistencias WHERE id=? AND sucursal_id=?",
                     (tid, current_sucursal_id())).fetchone()
    if not row:
        db.close()
        flash("Asistencia no encontrada.", "warning")
        return redirect(request.referrer or url_for("tickets"))

    if (row["estado"] or "").lower() != "resuelto":
        db.close()
        flash("Solo se pueden eliminar asistencias RESUELTAS.", "warning")
        return redirect(request.referrer or url_for("tickets"))

    db.execute("DELETE FROM asistencias WHERE id=? AND sucursal_id=?", (tid, current_sucursal_id()))
    db.commit(); db.close()
    flash("Asistencia eliminada.", "success")
    return redirect(request.referrer or url_for("tickets"))

@app.route("/tickets/<int:tid>/borrar", methods=["POST"])
def tickets_borrar(tid):
    """Elimina un ticket en cualquier estado (solo administradores). Se usa desde Estadísticas."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    volver = request.referrer or url_for("estadisticas")
    if session.get("rol") != "admin":
        flash("Solo un administrador puede eliminar trabajos.", "warning")
        return redirect(volver)
    db = get_db()
    t = db.execute("SELECT cliente, tipo FROM asistencias WHERE id=? AND sucursal_id=?",
                   (tid, current_sucursal_id())).fetchone()
    if not t:
        db.close()
        flash("El trabajo ya no existe.", "warning")
        return redirect(volver)
    db.execute("DELETE FROM asistencias WHERE id=? AND sucursal_id=?", (tid, current_sucursal_id()))
    db.commit()
    db.close()
    flash(f"Se eliminó {(t['tipo'] or 'el trabajo').lower()} de {t['cliente'] or 'cliente sin nombre'}.", "success")
    return redirect(volver)

@app.route("/tickets/eliminar_resueltos", methods=["POST"])
def tickets_eliminar_resueltos():
    """Elimina todos los tickets en estado 'resuelto' de la sucursal actual."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    acols = table_columns(db, "asistencias")
    if "estado" not in acols:
        db.close()
        flash("No existe la columna 'estado' en asistencias. Actualizá la BD.", "warning")
        return redirect(request.referrer or url_for("tickets"))

    cur = db.execute("DELETE FROM asistencias WHERE sucursal_id=? AND lower(IFNULL(estado,''))='resuelto'",
                     (current_sucursal_id(),))
    borrados = cur.rowcount if cur.rowcount is not None else 0
    db.commit(); db.close()

    flash(f"Eliminadas {borrados} asistencias resueltas.", "success")
    return redirect(request.referrer or url_for("tickets"))

# ===========================
#  Importar clientes (CSV)
# ===========================
@app.route("/clientes/importar", methods=["POST"])
def clientes_importar():
    if "usuario_id" not in session and "usuario" not in session:
        return redirect(url_for("login"))

    f = request.files.get("csvfile")
    if not f or f.filename == "":
        flash("Seleccioná un archivo CSV.", "warning")
        return redirect(url_for("clientes"))

    text, enc = _try_decode(f)
    delim = _guess_delimiter(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    if not reader.fieldnames:
        flash("No pude leer encabezados del CSV.", "danger")
        return redirect(url_for("clientes"))

    keymap, desconocidas = {}, []
    for h in reader.fieldnames:
        nk = _norm_key(h)
        tk = _HEADER_MAP.get(nk)
        if tk: keymap[h] = tk
        else: desconocidas.append(h)
    if not keymap:
        flash("No reconocí columnas del CSV. Revisá los encabezados.", "danger")
        return redirect(url_for("clientes"))
    if desconocidas:
        flash(f"Aviso: columnas ignoradas: {', '.join(desconocidas)}", "warning")

    db = get_db()
    cols = table_columns(db, "clientes")
    ins = upd = 0
    sid = current_sucursal_id()
    csv_trae_plan = bool({"tipo", "valor", "tipo_valor"} & set(keymap.values()))

    for raw in reader:
        if not any((str(v or "").strip() for v in raw.values())):
            continue

        norm = { keymap[k]: (raw.get(k) or "").strip() for k in raw if k in keymap }

        ext_id     = norm.get("external_id") or None
        nombre     = norm.get("nombre") or ""
        referencia = norm.get("referencia") or None
        barrio     = norm.get("barrio") or None
        telefono   = norm.get("telefono") or None
        if telefono:
            digits = _only_digits(telefono)
            telefono = digits if len(digits) >= 6 else telefono
        situacion  = norm.get("situacion") or ""
        exonerado  = _parse_bool(norm.get("exonerado"))
        tv         = norm.get("tipo_valor")
        tipo_in    = norm.get("tipo")
        valor_in   = norm.get("valor")
        tipo_final, valor_final = _split_tipo_valor(tipo_in or tv, valor_in)
        venc       = _parse_date_to_iso(norm.get("vencimiento"))

        activo = 0 if situacion.lower() in ("inactivo","baja","suspendido","cancelado") else 1
        if not (ext_id or nombre or telefono):
            continue

        data = {
            "external_id": ext_id,
            "nombre": nombre,
            "referencia": referencia,
            "barrio": barrio,
            "telefono": telefono,
            "situacion": situacion or None,
            "exonerado": exonerado,
            "vencimiento": venc,
            "activo": activo,
            "sucursal_id": sid
        }
        if "tipo" in cols:   data["tipo"] = tipo_final
        if "valor" in cols:  data["valor"] = valor_final
        if "tipo_valor" in cols and ("tipo" not in cols or "valor" not in cols):
            data["tipo_valor"] = tv or f"{tipo_final} {valor_final or ''}".strip()

        # NUEVO: cédula, PPPoE y dirección (solo si vienen en el CSV, para no pisar datos)
        for extra in ("cedula", "pppoe", "direccion"):
            if extra in cols and norm.get(extra):
                data[extra] = norm.get(extra)

        data = {k: v for k, v in data.items() if k in cols}

        row = None
        if ext_id:
            row = db.execute("SELECT id FROM clientes WHERE external_id=? AND sucursal_id=?",
                             (ext_id, sid)).fetchone()
        if not row and telefono:
            row = db.execute("SELECT id FROM clientes WHERE telefono=? AND sucursal_id=?",
                             (telefono, sid)).fetchone()

        if row:
            # Al actualizar, no borrar el plan si el CSV no trae esas columnas
            if not csv_trae_plan:
                for k in ("tipo", "valor", "tipo_valor"):
                    data.pop(k, None)
            sets = ", ".join([f"{k}=?" for k in data.keys()])
            db.execute(f"UPDATE clientes SET {sets} WHERE id=?", list(data.values())+[row["id"]])
            upd += 1
        else:
            qs = ", ".join(["?"]*len(data))
            db.execute(f"INSERT INTO clientes ({', '.join(data.keys())}) VALUES ({qs})", list(data.values()))
            ins += 1

    db.commit(); db.close()
    flash(f"Importación OK. Insertados {ins}, actualizados {upd}. (codificación {enc}, separador '{delim}')", "success")
    return redirect(url_for("clientes"))

# --- API datos para el mapa ---
@app.route("/api/mapa_datos", endpoint="api_mapa_datos")
def api_mapa_datos():
    if "usuario" not in session and "usuario_id" not in session:
        return jsonify({"error": "no_auth"}), 401

    db = get_db()
    sid = current_sucursal_id()

    # Tickets con coordenadas
    tickets = db.execute("""
        SELECT a.id, a.cliente, a.direccion, a.tipo, a.prioridad, a.estado,
               a.programada_en, a.lat, a.lng,
               COALESCE(t.nombre, a.tecnico) AS tecnico
        FROM asistencias a
        LEFT JOIN tecnicos t ON a.tecnico_id = t.id
        WHERE a.sucursal_id=?
          AND a.lat IS NOT NULL AND a.lng IS NOT NULL
        ORDER BY datetime(COALESCE(a.programada_en, a.fecha)) DESC
        LIMIT 500
    """, (sid,)).fetchall()

    # Última posición por técnico (por sucursal)
    pos = db.execute("""
        SELECT tp.tecnico_id, tp.lat, tp.lng, tp.ts, te.nombre
        FROM tecnico_pos tp
        JOIN (
            SELECT tecnico_id, MAX(ts) AS mts
            FROM tecnico_pos
            WHERE sucursal_id = ?
            GROUP BY tecnico_id
        ) x ON x.tecnico_id = tp.tecnico_id AND x.mts = tp.ts
        LEFT JOIN tecnicos te ON te.id = tp.tecnico_id
        WHERE tp.sucursal_id = ?
    """, (sid, sid)).fetchall()
    db.close()

    return jsonify({
        "tickets": [dict(r) for r in tickets],
        "tecnicos": [
            {"id": r["tecnico_id"], "nombre": r["nombre"],
             "lat": r["lat"], "lng": r["lng"], "ts": r["ts"]}
            for r in pos
        ]
    })

# --- Trayectoria de un técnico (por fecha) ---
@app.route("/api/tecnico_trayectoria/<int:tid>", endpoint="api_tecnico_trayectoria")
def api_tecnico_trayectoria(tid):
    if "usuario" not in session and "usuario_id" not in session:
        return jsonify({"error": "no_auth"}), 401

    desde = request.args.get("desde") or (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    hasta = request.args.get("hasta") or datetime.now().strftime("%Y-%m-%d")
    sid   = current_sucursal_id()

    db = get_db()

    # 1) Normal: por sucursal y rango
    rows = db.execute("""
        SELECT lat, lng, ts
          FROM tecnico_pos
         WHERE sucursal_id = ?
           AND tecnico_id   = ?
           AND date(ts) BETWEEN ? AND ?
         ORDER BY ts ASC
    """, (sid, tid, desde, hasta)).fetchall()

    # 2) Fallback: ignorar sucursal y ampliar a 30 días
    if not rows:
        d2 = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
        h2 = datetime.now().strftime("%Y-%m-%d")
        rows = db.execute("""
            SELECT lat, lng, ts
              FROM tecnico_pos
             WHERE tecnico_id = ?
               AND date(ts) BETWEEN ? AND ?
             ORDER BY ts ASC
        """, (tid, d2, h2)).fetchall()

    # 3) Fallback 2: tecnico_tracks (por si se guardó allí)
    if not rows:
        try:
            rows = db.execute("""
                SELECT lat, lng, ts
                  FROM tecnico_tracks
                 WHERE tecnico_id = ?
                 ORDER BY ts ASC
                 LIMIT 500
            """, (tid,)).fetchall()
        except sqlite3.OperationalError:
            rows = []

    db.close()
    return jsonify([dict(r) for r in rows])

# --- Ping GPS (móvil) ---
@app.route("/gps", methods=["GET", "POST"], endpoint="gps_ping")
def gps_ping():
    tecnico_id = request.values.get("tecnico_id", type=int)
    lat = request.values.get("lat", type=float)
    lng = request.values.get("lng", type=float)
    sucursal_id = request.values.get("sucursal_id", type=int)

    if not tecnico_id or lat is None or lng is None:
        return "Faltan parametros (tecnico_id, lat, lng)", 400

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    db = get_db()

    # Si no vino sucursal, inferimos desde el técnico o usamos la sucursal actual
    if not sucursal_id:
        row = db.execute("SELECT sucursal_id FROM tecnicos WHERE id=?", (tecnico_id,)).fetchone()
        sucursal_id = row["sucursal_id"] if row and row["sucursal_id"] else current_sucursal_id()

    db.execute(
        "INSERT INTO tecnico_pos (tecnico_id, lat, lng, ts, sucursal_id) VALUES (?,?,?,?,?)",
        (tecnico_id, float(lat), float(lng), ts, int(sucursal_id))
    )

    cols = table_columns(db, "tecnicos")
    if {"lat", "lng", "pos_updated_at", "sucursal_id"} <= cols:
        db.execute(
            "UPDATE tecnicos SET lat=?, lng=?, pos_updated_at=?, sucursal_id=? WHERE id=?",
            (float(lat), float(lng), ts, int(sucursal_id), tecnico_id)
        )
    db.commit()
    db.close()
    return "ok"

# --- Configuración (preferencias simples) ---
@app.route("/configuracion", methods=["GET", "POST"])
def configuracion():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    tab = request.args.get("tab") or request.form.get("tab") or "tecnicos"
    uid = session.get("usuario_id")

    if request.method == "POST":
        db = get_db()
        seccion = request.form.get("seccion")
        try:
            if seccion == "preferencias":
                notifs = 1 if request.form.get("notifs") == "on" else 0
                sonido = 1 if request.form.get("notif_sonido") == "on" and notifs else 0
                pred = request.form.get("sucursal_pred", type=int)
                update_user_fields(db, uid, {"notifs": notifs, "notif_sonido": sonido, "sucursal_pred": pred})
                session["notifs"], session["notif_sonido"] = notifs, sonido
                flash("Preferencias guardadas.", "success")
                tab = "preferencias"
            elif seccion == "empresa":
                for k in EMPRESA_CAMPOS:
                    db.execute("INSERT INTO config (clave, valor) VALUES (?, ?) "
                               "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor",
                               (f"empresa.{k}", (request.form.get(k) or "").strip()))
                db.commit()
                flash("Datos de la empresa guardados. Ya aparecen en las órdenes de trabajo.", "success")
                tab = "empresa"
        finally:
            db.close()
        return redirect(url_for("configuracion", tab=tab))

    db = get_db()
    user = db.execute("SELECT * FROM usuarios WHERE id=?", (uid,)).fetchone()
    user = dict(user) if user else {}
    sucursales = db.execute("SELECT id, nombre FROM sucursales ORDER BY id").fetchall()
    try:
        # Solo los de la sucursal actual (y los que quedaron sin sucursal, para poder asignarlos)
        tecnicos = [dict(r) for r in db.execute("""
            SELECT t.*, s.nombre AS sucursal_nombre
              FROM tecnicos t LEFT JOIN sucursales s ON s.id = t.sucursal_id
             WHERE IFNULL(t.sucursal_id, 0) IN (?, 0)
             ORDER BY IFNULL(t.activo,1) DESC, t.nombre COLLATE NOCASE
        """, (current_sucursal_id(),)).fetchall()]
    except sqlite3.OperationalError:
        tecnicos = []
    empresa = get_empresa(db)
    db.close()

    # Tickets abiertos de cada técnico (para avisar antes de eliminar)
    abiertos = {}
    try:
        db2 = get_db()
        for r in db2.execute("""
            SELECT tecnico_id, COUNT(*) AS n FROM asistencias
             WHERE tecnico_id IS NOT NULL AND lower(IFNULL(estado,'pendiente')) IN ('pendiente','en_progreso')
             GROUP BY tecnico_id"""):
            abiertos[r["tecnico_id"]] = r["n"]
        db2.close()
    except sqlite3.Error:
        pass

    for t in tecnicos:
        t["wa"] = _tel_whatsapp(t.get("telefono") or "")
        t["abiertos"] = abiertos.get(t["id"], 0)

    return render_template("configuracion.html", tab=tab, user=user, sucursales=sucursales,
                           tecnicos=tecnicos, empresa=empresa)

@app.route("/configuracion/tecnicos/guardar", methods=["POST"])
def config_tecnico_guardar():
    """Crea un técnico nuevo o edita uno existente (si viene id)."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    tid = request.form.get("id", type=int)
    nombre = (request.form.get("nombre") or "").strip()
    telefono = (request.form.get("telefono") or "").strip()
    # Cada sucursal tiene sus propios técnicos: por defecto, la sucursal actual
    sucursal_id = request.form.get("sucursal_id", type=int) or current_sucursal_id()

    db = get_db()

    if tid:
        viejo = db.execute("SELECT nombre FROM tecnicos WHERE id=?", (tid,)).fetchone()
        otro = db.execute("SELECT 1 FROM tecnicos WHERE lower(nombre)=lower(?) AND sucursal_id=? AND id<>?",
                          (nombre, sucursal_id, tid)).fetchone()
        if otro:
            db.close()
            flash(f"Ya existe un técnico llamado {nombre} en esa sucursal.", "warning")
            return redirect(url_for("configuracion", tab="tecnicos"))
        db.execute("UPDATE tecnicos SET nombre=?, telefono=?, sucursal_id=? WHERE id=?",
                   (nombre, telefono or None, sucursal_id, tid))
        # Si cambió el nombre, actualizarlo en sus tickets para no perder el historial
        if viejo and viejo["nombre"] != nombre and "tecnico" in table_columns(db, "asistencias"):
            db.execute("UPDATE asistencias SET tecnico=? WHERE tecnico=?", (nombre, viejo["nombre"]))
        flash(f"Técnico {nombre} actualizado.", "success")
    else:
        existe = db.execute("SELECT 1 FROM tecnicos WHERE lower(nombre)=lower(?) AND sucursal_id=?",
                            (nombre, sucursal_id)).fetchone()
        if existe:
            db.close()
            flash(f"Ya existe un técnico llamado {nombre} en esa sucursal.", "warning")
            return redirect(url_for("configuracion", tab="tecnicos"))
        insert_row(db, "tecnicos", {"nombre": nombre, "telefono": telefono or None,
                                    "sucursal_id": sucursal_id, "activo": 1})
        flash(f"Técnico {nombre} agregado.", "success")
    db.commit()
    db.close()
    return redirect(url_for("configuracion", tab="tecnicos"))

@app.route("/configuracion/tecnicos/<int:tid>/estado", methods=["POST"])
def config_tecnico_estado(tid):
    """Activa o desactiva un técnico (no se borra, para conservar su historial)."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    db = get_db()
    t = db.execute("SELECT nombre, IFNULL(activo,1) AS activo FROM tecnicos WHERE id=?", (tid,)).fetchone()
    if t:
        nuevo = 0 if t["activo"] == 1 else 1
        db.execute("UPDATE tecnicos SET activo=? WHERE id=?", (nuevo, tid))
        db.commit()
        flash(f"{t['nombre']} {'activado' if nuevo else 'desactivado'}.", "success")
    db.close()
    return redirect(url_for("configuracion", tab="tecnicos"))


@app.route("/configuracion/tecnicos/<int:tid>/eliminar", methods=["POST"])
def config_tecnico_eliminar(tid):
    """Elimina un técnico. Sus tickets conservan el nombre, pero quedan sin técnico asignado."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    db = get_db()
    t = db.execute("SELECT nombre FROM tecnicos WHERE id=?", (tid,)).fetchone()
    if not t:
        db.close()
        flash("El técnico ya no existe.", "warning")
        return redirect(url_for("configuracion", tab="tecnicos"))
    try:
        acols = table_columns(db, "asistencias")
        if "tecnico_id" in acols:
            # Tickets abiertos: quedan sin técnico para reasignarlos
            if "tecnico" in acols:
                db.execute("""UPDATE asistencias SET tecnico=NULL
                               WHERE tecnico_id=? AND lower(IFNULL(estado,'pendiente')) IN ('pendiente','en_progreso')""", (tid,))
            # Todos: se quita la referencia; los cerrados conservan el nombre como historial
            db.execute("UPDATE asistencias SET tecnico_id=NULL WHERE tecnico_id=?", (tid,))
        db.execute("DELETE FROM tecnicos WHERE id=?", (tid,))
        db.commit()
        flash(f"Técnico {t['nombre']} eliminado.", "success")
    except sqlite3.Error:
        db.rollback()
        flash(f"No se pudo eliminar a {t['nombre']}. Puedes desactivarlo en su lugar.", "error")
    finally:
        db.close()
    return redirect(url_for("configuracion", tab="tecnicos"))


@app.route("/usuarios")
def usuarios():
    """Lista de usuarios del sistema (todos la ven; solo admin la modifica)."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    db = get_db()
    filas = [dict(r) for r in db.execute("""
        SELECT u.*, s.nombre AS sucursal_nombre
          FROM usuarios u LEFT JOIN sucursales s ON s.id = u.sucursal_id
         ORDER BY IFNULL(u.activo,1) DESC, u.nombre COLLATE NOCASE, u.usuario COLLATE NOCASE
    """).fetchall()]
    sucursales = db.execute("SELECT id, nombre FROM sucursales ORDER BY id").fetchall()
    db.close()
    for u in filas:
        foto = u.get("foto_url")
        u["foto"] = url_for("static", filename=foto) if foto else None
    return render_template("usuarios.html", usuarios=filas, sucursales=sucursales,
                           es_admin=_puede_gestionar_usuarios(), sin_admin=not _hay_admin())

def _hay_admin():
    try:
        db = get_db()
        n = db.execute("SELECT COUNT(*) FROM usuarios WHERE rol='admin' AND IFNULL(activo,1)=1").fetchone()[0]
        db.close()
        return n > 0
    except sqlite3.Error:
        return True

def _puede_gestionar_usuarios():
    return session.get("rol") == "admin" or not _hay_admin()

def _solo_admin():
    if not _puede_gestionar_usuarios():
        flash("Solo un administrador puede gestionar usuarios.", "warning")
        return redirect(url_for("usuarios"))
    return None

@app.route("/configuracion/usuarios/guardar", methods=["POST"])
def config_usuario_guardar():
    """Crea un usuario o edita uno existente (si viene id)."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    no = _solo_admin()
    if no: return no

    f = request.form
    uid = f.get("id", type=int)
    usuario = (f.get("usuario") or "").strip()
    nombre = (f.get("nombre") or "").strip()
    clave = f.get("contrasena") or ""
    rol = f.get("rol") if f.get("rol") in ("admin", "operador", "tecnico") else "operador"
    sucursal_id = f.get("sucursal_id", type=int) or current_sucursal_id()

    if not usuario or not nombre:
        flash("El nombre y el usuario son obligatorios.", "warning")
        return redirect(url_for("usuarios"))
    if (not uid and len(clave) < 6) or (uid and clave and len(clave) < 6):
        flash("La contraseña debe tener al menos 6 caracteres.", "warning")
        return redirect(url_for("usuarios"))

    db = get_db()
    cols = table_columns(db, "usuarios")
    datos = {
        "usuario": usuario, "nombre": nombre,
        "email": (f.get("email") or "").strip() or None,
        "telefono": (f.get("telefono") or "").strip() or None,
        "area": (f.get("area") or "").strip() or None,
        "turno": (f.get("turno") or "").strip() or None,
        "rol": rol, "sucursal_id": sucursal_id,
    }
    if clave:
        if "password_hash" in cols:
            datos["password_hash"] = generate_password_hash(clave)
            datos["contrasena"] = None
        else:
            datos["contrasena"] = clave

    # No dejar el sistema sin administradores
    if uid and rol != "admin":
        era_admin = db.execute("SELECT rol FROM usuarios WHERE id=?", (uid,)).fetchone()
        otros = db.execute("SELECT COUNT(*) FROM usuarios WHERE rol='admin' AND IFNULL(activo,1)=1 AND id<>?", (uid,)).fetchone()[0]
        if era_admin and era_admin["rol"] == "admin" and otros == 0:
            db.close()
            flash("Debe quedar al menos un administrador activo.", "warning")
            return redirect(url_for("usuarios"))

    datos = {k: v for k, v in datos.items() if k in cols}
    try:
        if uid:
            sets = ", ".join(f"{k}=?" for k in datos)
            extra = ", updated_at=CURRENT_TIMESTAMP" if "updated_at" in cols else ""
            db.execute(f"UPDATE usuarios SET {sets}{extra} WHERE id=?", list(datos.values()) + [uid])
            if uid == session.get("usuario_id"):
                session["nombre"], session["rol"], session["email"] = nombre, rol, datos.get("email")
            flash(f"Usuario {usuario} actualizado." + (" Se cambió su contraseña." if clave else ""), "success")
        else:
            ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for k, v in (("created_at", ahora), ("updated_at", ahora), ("activo", 1)):
                if k in cols: datos[k] = v
            insert_row(db, "usuarios", datos)
            flash(f"Usuario {usuario} creado. Ya puede iniciar sesión.", "success")
        db.commit()
    except sqlite3.IntegrityError:
        db.rollback()
        flash(f"Ya existe el usuario «{usuario}» en esa sucursal.", "warning")
    finally:
        db.close()
    return redirect(url_for("usuarios"))

@app.route("/configuracion/usuarios/<int:uid>/estado", methods=["POST"])
def config_usuario_estado(uid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    no = _solo_admin()
    if no: return no
    if uid == session.get("usuario_id"):
        flash("No puedes desactivar tu propio usuario.", "warning")
        return redirect(url_for("usuarios"))
    db = get_db()
    u = db.execute("SELECT usuario, rol, IFNULL(activo,1) AS activo FROM usuarios WHERE id=?", (uid,)).fetchone()
    if u:
        nuevo = 0 if u["activo"] == 1 else 1
        if nuevo == 0 and u["rol"] == "admin":
            otros = db.execute("SELECT COUNT(*) FROM usuarios WHERE rol='admin' AND IFNULL(activo,1)=1 AND id<>?", (uid,)).fetchone()[0]
            if otros == 0:
                db.close()
                flash("Debe quedar al menos un administrador activo.", "warning")
                return redirect(url_for("usuarios"))
        db.execute("UPDATE usuarios SET activo=? WHERE id=?", (nuevo, uid))
        db.commit()
        flash(f"Usuario {u['usuario']} {'activado' if nuevo else 'desactivado: ya no puede iniciar sesión'}.", "success")
    db.close()
    return redirect(url_for("usuarios"))

@app.route("/configuracion/usuarios/<int:uid>/eliminar", methods=["POST"])
def config_usuario_eliminar(uid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))
    no = _solo_admin()
    if no: return no
    if uid == session.get("usuario_id"):
        flash("No puedes eliminar tu propio usuario.", "warning")
        return redirect(url_for("usuarios"))
    db = get_db()
    u = db.execute("SELECT usuario, rol FROM usuarios WHERE id=?", (uid,)).fetchone()
    if u:
        if u["rol"] == "admin" and db.execute(
                "SELECT COUNT(*) FROM usuarios WHERE rol='admin' AND IFNULL(activo,1)=1 AND id<>?", (uid,)).fetchone()[0] == 0:
            db.close()
            flash("Debe quedar al menos un administrador activo.", "warning")
            return redirect(url_for("usuarios"))
        db.execute("DELETE FROM usuarios WHERE id=?", (uid,))
        db.commit()
        flash(f"Usuario {u['usuario']} eliminado.", "success")
    db.close()
    return redirect(url_for("usuarios"))

# ===========================
#  Notificaciones (helper)
# ===========================
def push_notificacion(titulo: str, cuerpo: str = "", sucursal_id: int | None = None):
    """
    Inserta una notificación y devuelve el id.
    Usa la sucursal actual si no se especifica.
    """
    sid = sucursal_id or current_sucursal_id()
    db = get_db()
    cur = db.execute("""
        INSERT INTO notificaciones (sucursal_id, titulo, cuerpo, creado_en)
        VALUES (?, ?, ?, datetime('now'))
    """, (sid, titulo.strip() or "Notificación", cuerpo or ""))
    db.commit()
    nid = cur.lastrowid
    db.close()
    return nid

# ===========================
#  API: Notificaciones (poll)
# ===========================
@app.route("/api/notificaciones")
def api_notificaciones():
    if "usuario" not in session and "usuario_id" not in session:
        return jsonify([])

    since_id = request.args.get("since_id", type=int)
    sid = current_sucursal_id()

    db = get_db()
    if since_id:
        rows = db.execute("""
            SELECT id, titulo, cuerpo, creado_en
              FROM notificaciones
             WHERE sucursal_id=? AND id > ?
             ORDER BY id ASC
             LIMIT 100
        """, (sid, since_id)).fetchall()
    else:
        rows = db.execute("""
            SELECT id, titulo, cuerpo, creado_en
              FROM notificaciones
             WHERE sucursal_id=?
             ORDER BY id DESC
             LIMIT 20
        """, (sid,)).fetchall()
        rows = rows[::-1]  # las más viejas primero para el primer render

    db.close()
    return jsonify([dict(r) for r in rows])

# Prueba manual: http://127.0.0.1:5000/api/notificaciones/test?msg=Hola
@app.route("/api/notificaciones/test")
def api_notif_test():
    if "usuario" not in session and "usuario_id" not in session:
        return jsonify({"ok": False, "error": "no_auth"}), 401
    msg = request.args.get("msg") or "Mensaje de prueba"
    nid = push_notificacion("Nuevo mensaje", msg)
    return jsonify({"ok": True, "id": nid})

# === BLOQUE GEOLOCALIZACIÓN: API token ===
@app.route("/api/location/push", methods=["POST"])
def api_location_push():
    if not request.is_json:
        return jsonify({"ok": False, "error": "json_required"}), 400

    data = request.get_json()
    token = (data.get("token") or "").strip()
    lat = data.get("lat")
    lon = data.get("lon")

    if not token or lat is None or lon is None:
        return jsonify({"ok": False, "error": "missing_fields"}), 400

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    trow = _get_tecnico_by_token(conn, token)
    if not trow:
        conn.close()
        return jsonify({"ok": False, "error": "invalid_token"}), 401

    _insert_tecnico_pos(conn, trow["id"], lat, lon, ts, trow["sucursal_id"])
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "ts": ts})

@app.route("/api/location/live", methods=["GET"])
def api_location_live():
    sid = current_sucursal_id()
    conn = get_db()
    rows = conn.execute("""
        SELECT te.id, te.nombre, te.lat, te.lng AS lon, te.pos_updated_at AS ts
          FROM tecnicos te
         WHERE te.sucursal_id=? AND te.lat IS NOT NULL
         ORDER BY te.nombre
    """, (sid,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

# === BLOQUE GEOLOCALIZACIÓN: Vehículos ===
@app.route("/api/vehiculos/posiciones", methods=["GET"])
def api_vehiculos_posiciones():
    sid = current_sucursal_id()
    conn = get_db()
    rows = conn.execute("""
        SELECT v.id, v.placa, v.lat, v.lon, v.last_ts AS ts, t.nombre AS tecnico
          FROM vehiculos v
     LEFT JOIN tecnicos t ON t.id = v.tecnico_id
         WHERE v.sucursal_id=? ORDER BY v.placa
    """, (sid,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/gps-webhook", methods=["POST"])
def api_gps_webhook():
    data = request.get_json()
    if not data:
        return jsonify({"ok": False, "error": "json_required"}), 400

    items = data if isinstance(data, list) else [data]
    conn = get_db()
    for p in items:
        imei = p.get("deviceId") or p.get("imei")
        lat = p.get("lat")
        lon = p.get("lon")
        ts = p.get("timestamp") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if not imei or lat is None or lon is None:
            continue
        v = conn.execute("SELECT id FROM vehiculos WHERE imei_gps=?", (imei,)).fetchone()
        if not v:
            conn.execute("INSERT INTO vehiculos (placa, imei_gps, lat, lon, last_ts, sucursal_id) VALUES (?,?,?,?,?,?)",
                         (f"IMEI-{str(imei)[-6:]}", imei, lat, lon, ts, current_sucursal_id()))
        else:
            conn.execute("UPDATE vehiculos SET lat=?, lon=?, last_ts=? WHERE id=?", (lat, lon, ts, v["id"]))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

# === BLOQUE GEOLOCALIZACIÓN: CORS ===
@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return resp


# ===========================
#  Orden de trabajo (OT) imprimible por ticket
# ===========================
def _tel_whatsapp(tel):
    """Convierte un teléfono paraguayo (0981...) al formato de WhatsApp (595981...)."""
    d = _only_digits(tel)
    if not d:
        return ""
    if d.startswith("595"):
        return d
    if d.startswith("0"):
        d = d[1:]
    return "595" + d

def _ot_datos(db, t, sid):
    """Arma los datos de la OT de un ticket: cliente, técnico, mapa y mensaje de WhatsApp."""
    from urllib.parse import quote

    # Cliente: por cliente_id o, si no hay, por nombre
    c = None
    if t.get("cliente_id"):
        c = db.execute("SELECT * FROM clientes WHERE id=?", (t["cliente_id"],)).fetchone()
    if not c and t.get("cliente"):
        c = db.execute("SELECT * FROM clientes WHERE lower(nombre)=lower(?) AND sucursal_id=?",
                       (t["cliente"], sid)).fetchone()
    c = dict(c) if c else {}

    # Plan del cliente: valor -> plan -> número dentro de tipo_valor
    plan = c.get("valor") or c.get("plan")
    if not plan:
        m = re.search(r"(\d[\d\.\,]*)", str(c.get("tipo_valor") or ""))
        plan = m.group(1) if m else None
    c["plan_mostrar"] = plan

    # Técnico: por tecnico_id o por nombre
    tec = None
    try:
        if t.get("tecnico_id"):
            tec = db.execute("SELECT * FROM tecnicos WHERE id=?", (t["tecnico_id"],)).fetchone()
        if not tec and t.get("tecnico"):
            tec = db.execute("SELECT * FROM tecnicos WHERE lower(nombre)=lower(?) AND sucursal_id=?",
                             (t["tecnico"], sid)).fetchone()
    except sqlite3.OperationalError:
        tec = None
    tec = dict(tec) if tec else {}

    ot_num = f"OT-{int(t['id']):06d}"
    lat = t.get("lat") or c.get("lat")
    lng = t.get("lng") or c.get("lng")
    maps_url = f"https://www.google.com/maps?q={lat},{lng}" if lat and lng else None
    ot_url = url_for("tickets_ot", tid=t["id"], _external=True)
    direccion = t.get("direccion") or c.get("direccion") or ""

    # Mensaje de WhatsApp con lo esencial (el técnico puede no tener usuario en el sistema)
    lineas = [
        f"*Orden de trabajo {ot_num}*",
        f"Tipo: {t.get('tipo') or '-'}  |  Señal: {t.get('prioridad') or '-'}",
    ]
    if t.get("programada_en"):
        lineas.append(f"Visita: {str(t['programada_en'])[:16]}")
    lineas += [
        "",
        f"*Cliente:* {t.get('cliente') or '-'}",
        f"CI: {t.get('cedula') or c.get('cedula') or '-'}  |  PPPoE: {t.get('pppoe') or c.get('pppoe') or '-'}",
        f"Tel: {c.get('telefono') or '-'}",
        f"Barrio: {c.get('barrio') or '-'}",
    ]
    if direccion:
        lineas.append(f"Dirección: {direccion}")
    if c.get("referencia"):
        lineas.append(f"Referencia: {c['referencia']}")
    if maps_url:
        lineas.append(f"Ubicación: {maps_url}")
    lineas += ["", f"*Problema:* {t.get('problema') or '-'}", "", f"OT completa: {ot_url}"]
    texto = quote("\n".join(lineas))

    tel_tec = _tel_whatsapp(tec.get("telefono") or tec.get("celular") or "")
    wa_url = f"https://wa.me/{tel_tec}?text={texto}" if tel_tec else f"https://wa.me/?text={texto}"

    return dict(c=c, tec=tec, ot_num=ot_num, maps_url=maps_url, ot_url=ot_url,
                wa_url=wa_url, tec_tiene_tel=bool(tel_tec), direccion=direccion)

@app.route("/tickets/<int:tid>/ot")
def tickets_ot(tid):
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    db = get_db()
    sid = current_sucursal_id()
    row = db.execute("SELECT * FROM asistencias WHERE id=? AND sucursal_id=?", (tid, sid)).fetchone()
    if not row:
        db.close()
        flash("Ticket no encontrado en esta sucursal.", "warning")
        return redirect(url_for("ordenes_trabajo"))
    t = dict(row)
    datos = _ot_datos(db, t, sid)

    suc = db.execute("SELECT * FROM sucursales WHERE id=?", (sid,)).fetchone()
    suc = dict(suc) if suc else {"nombre": g.sucursal_nombre}
    db.close()

    return render_template(
        "ordendetrabajo.html",
        t=t, suc=suc, empresa=get_empresa(), **datos,
        emitida=datetime.now().strftime("%d/%m/%Y %H:%M"),
        emitida_por=session.get("nombre") or session.get("usuario"),
    )

@app.route("/ordenes")
def ordenes_trabajo():
    """Lista de órdenes de trabajo (una por ticket) con filtros por estado, técnico y fecha."""
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    vista      = request.args.get("vista", "abiertas")   # abiertas | hoy | todas
    tecnico_f  = (request.args.get("tecnico") or "").strip()

    db = get_db()
    sid = current_sucursal_id()
    acols = table_columns(db, "asistencias")

    sql = "SELECT * FROM asistencias WHERE sucursal_id=?"
    params = [sid]
    if vista == "abiertas" and "estado" in acols:
        sql += " AND lower(IFNULL(estado,'pendiente')) IN ('pendiente','en_progreso')"
    elif vista == "hoy":
        col_fecha = "COALESCE(programada_en, fecha)" if "programada_en" in acols else "fecha"
        sql += f" AND date({col_fecha}) = date('now','localtime')"
    if tecnico_f:
        sql += " AND lower(IFNULL(tecnico,'')) = lower(?)"
        params.append(tecnico_f)
    orden = "datetime(COALESCE(programada_en, fecha))" if "programada_en" in acols else "datetime(fecha)"
    sql += f" ORDER BY {orden} DESC LIMIT 300"
    rows = [dict(r) for r in db.execute(sql, params).fetchall()]

    ordenes = []
    for t in rows:
        d = _ot_datos(db, t, sid)
        ordenes.append(dict(t=t, **d))

    try:
        tecnicos = db.execute(
            "SELECT nombre FROM tecnicos WHERE activo=1 AND sucursal_id=? ORDER BY nombre COLLATE NOCASE",
            (sid,)
        ).fetchall()
        tecnicos = [r["nombre"] for r in tecnicos]
    except sqlite3.OperationalError:
        tecnicos = []

    # Contadores para las pestañas
    def contar(extra):
        return db.execute(f"SELECT COUNT(*) FROM asistencias WHERE sucursal_id=? {extra}", (sid,)).fetchone()[0]
    col_fecha = "COALESCE(programada_en, fecha)" if "programada_en" in acols else "fecha"
    cuentas = {
        "abiertas": contar("AND lower(IFNULL(estado,'pendiente')) IN ('pendiente','en_progreso')") if "estado" in acols else 0,
        "hoy": contar(f"AND date({col_fecha}) = date('now','localtime')"),
        "todas": contar(""),
    }
    db.close()

    return render_template("ordenes_trabajo.html", ordenes=ordenes, vista=vista,
                           tecnico_f=tecnico_f, tecnicos=tecnicos, cuentas=cuentas)


# ===========================
#  Estadísticas
# ===========================
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

@app.route("/estadisticas")
def estadisticas():
    if "usuario" not in session and "usuario_id" not in session:
        return redirect(url_for("login"))

    import calendar
    hoy = date.today()

    # Mes elegido (?mes=2026-09); por defecto el actual
    try:
        y, m = map(int, (request.args.get("mes") or "").split("-"))
        date(y, m, 1)
    except Exception:
        y, m = hoy.year, hoy.month
    mes_ini = date(y, m, 1)
    mes_fin = date(y, m, calendar.monthrange(y, m)[1])
    prev_mes = (mes_ini - timedelta(days=1)).strftime("%Y-%m")
    next_mes = (mes_fin + timedelta(days=1)).strftime("%Y-%m")
    es_mes_actual = (y, m) == (hoy.year, hoy.month)

    # Semana actual (lunes a viernes) y último fin de semana (sábado y domingo)
    wd = hoy.weekday()  # lunes=0 ... domingo=6
    lunes = hoy - timedelta(days=wd)
    viernes = lunes + timedelta(days=4)
    sabado = hoy - timedelta(days=wd - 5) if wd >= 5 else hoy - timedelta(days=wd + 2)
    domingo = sabado + timedelta(days=1)

    db = get_db()
    sid = current_sucursal_id()
    acols = table_columns(db, "asistencias")
    ES_INST = "lower(IFNULL(tipo,'')) LIKE 'instal%'"

    def contar(desde, hasta):
        r = db.execute(f"""
            SELECT COUNT(*) AS total,
                   COALESCE(SUM(CASE WHEN {ES_INST} THEN 1 ELSE 0 END), 0) AS inst
              FROM asistencias
             WHERE sucursal_id=? AND date(fecha) BETWEEN ? AND ?
        """, (sid, desde.isoformat(), hasta.isoformat())).fetchone()
        total, inst = r["total"] or 0, r["inst"] or 0
        return {"total": total, "inst": inst, "asist": total - inst}

    periodos = [
        {"clave": "hoy", "titulo": "Hoy", "rango": hoy.strftime("%d/%m"), **contar(hoy, hoy)},
        {"clave": "semana", "titulo": "Lunes a viernes", "rango": f"{lunes:%d/%m} al {viernes:%d/%m}", **contar(lunes, viernes)},
        {"clave": "finde", "titulo": "Fin de semana" if wd >= 5 else "Último fin de semana",
         "rango": f"{sabado:%d/%m} y {domingo:%d/%m}", **contar(sabado, domingo)},
        {"clave": "mes", "titulo": "Este mes" if es_mes_actual else "Mes elegido",
         "rango": f"{MESES_ES[m-1].capitalize()} {y}", **contar(mes_ini, mes_fin)},
    ]

    # Serie diaria del mes + división días hábiles / fines de semana
    filas = db.execute(f"""
        SELECT date(fecha) AS dia, COUNT(*) AS total,
               COALESCE(SUM(CASE WHEN {ES_INST} THEN 1 ELSE 0 END), 0) AS inst
          FROM asistencias
         WHERE sucursal_id=? AND date(fecha) BETWEEN ? AND ?
         GROUP BY date(fecha)
    """, (sid, mes_ini.isoformat(), mes_fin.isoformat())).fetchall()
    por_dia = {r["dia"]: (r["total"] - r["inst"], r["inst"]) for r in filas}
    dias, serie_asist, serie_inst = [], [], []
    habiles = {"asist": 0, "inst": 0}
    finde = {"asist": 0, "inst": 0}
    d = mes_ini
    while d <= mes_fin:
        a, i = por_dia.get(d.isoformat(), (0, 0))
        dias.append(d.day); serie_asist.append(a); serie_inst.append(i)
        destino = finde if d.weekday() >= 5 else habiles
        destino["asist"] += a; destino["inst"] += i
        d += timedelta(days=1)

    # Tickets del mes por estado
    estados = {"pendiente": 0, "en_progreso": 0, "resuelto": 0, "cancelado": 0}
    if "estado" in acols:
        for r in db.execute("""
            SELECT lower(IFNULL(estado,'pendiente')) AS e, COUNT(*) AS n
              FROM asistencias
             WHERE sucursal_id=? AND date(fecha) BETWEEN ? AND ?
             GROUP BY lower(IFNULL(estado,'pendiente'))
        """, (sid, mes_ini.isoformat(), mes_fin.isoformat())).fetchall():
            estados[r["e"] if r["e"] in estados else "pendiente"] += r["n"]

    # Por técnico en el mes
    por_tecnico = db.execute(f"""
        SELECT COALESCE(NULLIF(trim(tecnico),''), 'Sin asignar') AS tecnico,
               COUNT(*) AS total,
               COALESCE(SUM(CASE WHEN {ES_INST} THEN 1 ELSE 0 END), 0) AS inst,
               COALESCE(SUM(CASE WHEN lower(IFNULL(estado,''))='resuelto' THEN 1 ELSE 0 END), 0) AS resueltos
          FROM asistencias
         WHERE sucursal_id=? AND date(fecha) BETWEEN ? AND ?
         GROUP BY 1 ORDER BY total DESC
    """, (sid, mes_ini.isoformat(), mes_fin.isoformat())).fetchall()

    # Detalle por técnico: cada trabajo con fecha y hora, cliente y ubicación
    ccols_det = table_columns(db, "clientes")
    join_c = "LEFT JOIN clientes c ON c.id = a.cliente_id" if "cliente_id" in acols else "LEFT JOIN clientes c ON 0"
    barrio_c = "c.barrio" if "barrio" in ccols_det else "NULL"
    refer_c = "c.referencia" if "referencia" in ccols_det else "NULL"
    dir_c = "c.direccion" if "direccion" in ccols_det else "NULL"
    prog = "a.programada_en" if "programada_en" in acols else "NULL"
    detalle = {}
    for r in db.execute(f"""
        SELECT a.id, a.cliente, a.tipo, a.estado, a.fecha, {prog} AS programada_en,
               COALESCE(NULLIF(trim(a.direccion),''), {dir_c}) AS direccion, a.lat, a.lng,
               {barrio_c} AS barrio, {refer_c} AS referencia,
               COALESCE(NULLIF(trim(a.tecnico),''), 'Sin asignar') AS tecnico
          FROM asistencias a {join_c}
         WHERE a.sucursal_id=? AND date(a.fecha) BETWEEN ? AND ?
         ORDER BY datetime(COALESCE({prog}, a.fecha)) DESC
    """, (sid, mes_ini.isoformat(), mes_fin.isoformat())).fetchall():
        d = dict(r)
        d["cuando"] = str(d["programada_en"] or d["fecha"] or "")[:16]
        d["ubicacion"] = " · ".join(x for x in [d.get("barrio"), d.get("direccion") or d.get("referencia")] if x)
        d["mapa"] = f"https://www.google.com/maps?q={d['lat']},{d['lng']}" if d.get("lat") and d.get("lng") else None
        detalle.setdefault(d["tecnico"], []).append(d)

    # Clientes: activos, cancelados y otros (según "situacion" y "activo")
    ccols = table_columns(db, "clientes")
    clientes = {"total": 0, "activos": 0, "cancelados": 0, "otros": 0}
    if ccols:
        sit = "lower(IFNULL(situacion,''))" if "situacion" in ccols else "''"
        act = "IFNULL(activo,1)" if "activo" in ccols else "1"
        r = db.execute(f"""
            SELECT COUNT(*) AS total,
                   SUM(CASE WHEN {sit} LIKE '%cancel%' OR {sit} LIKE '%baja%' THEN 1 ELSE 0 END) AS cancelados,
                   SUM(CASE WHEN NOT ({sit} LIKE '%cancel%' OR {sit} LIKE '%baja%')
                             AND {act}=1
                             AND NOT ({sit} LIKE '%mora%' OR {sit} LIKE '%suspend%') THEN 1 ELSE 0 END) AS activos
              FROM clientes WHERE sucursal_id=?
        """, (sid,)).fetchone()
        clientes["total"] = r["total"] or 0
        clientes["cancelados"] = r["cancelados"] or 0
        clientes["activos"] = r["activos"] or 0
        clientes["otros"] = clientes["total"] - clientes["activos"] - clientes["cancelados"]
    db.close()

    return render_template(
        "estadisticas.html",
        periodos=periodos, dias=dias, serie_asist=serie_asist, serie_inst=serie_inst,
        habiles=habiles, finde=finde, estados=estados, por_tecnico=por_tecnico, detalle=detalle,
        clientes=clientes, mes_valor=f"{y}-{m:02d}", mes_nombre=f"{MESES_ES[m-1]} {y}",
        prev_mes=prev_mes, next_mes=next_mes, es_mes_actual=es_mes_actual,
        hay_siguiente=(mes_fin < hoy),
    )


# ===========================
#  Ubicación pegada desde Google Maps (enlaces cortos maps.app.goo.gl)
# ===========================
_RE_COORDS = [
    re.compile(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)"),                                   # lugar exacto
    re.compile(r"[?&](?:q|query|ll|center|destination)=(-?\d+\.\d+)(?:,|%2C)\s*(-?\d+\.\d+)"),
    re.compile(r"@(-?\d+\.\d+),(-?\d+\.\d+)"),                                       # centro de la vista
]

def _coords_de_texto(txt):
    for rx in _RE_COORDS:
        m = rx.search(txt or "")
        if m:
            lat, lng = float(m.group(1)), float(m.group(2))
            if -90 <= lat <= 90 and -180 <= lng <= 180:
                return lat, lng
    return None

@app.route("/api/ubicacion_desde_enlace")
def api_ubicacion_desde_enlace():
    """Abre un enlace corto de Google Maps (el de 'Compartir') y devuelve sus coordenadas."""
    if "usuario" not in session and "usuario_id" not in session:
        return jsonify({"error": "no_auth"}), 401
    url = (request.args.get("url") or "").strip()
    from urllib.parse import urlparse, unquote
    host = (urlparse(url).hostname or "").lower()
    permitidos = ("maps.app.goo.gl", "goo.gl", "g.co", "google.com", "www.google.com", "maps.google.com")
    if not url.startswith(("http://", "https://")) or not any(host == h or host.endswith("." + h) for h in permitidos):
        return jsonify({"error": "enlace_no_valido"}), 400
    try:
        r = requests.get(url, allow_redirects=True, timeout=8,
                         headers={"User-Agent": "Mozilla/5.0 (spynet)", "Accept-Language": "es"})
        coords = _coords_de_texto(unquote(r.url)) or _coords_de_texto(r.text[:200000])
    except Exception:
        coords = None
    if not coords:
        return jsonify({"error": "sin_coordenadas"}), 404
    return jsonify({"lat": coords[0], "lng": coords[1]})

# ===========================
#  Main
# ===========================
if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "1") == "1")