import os
import sqlite3
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), "asistencias.db")

# Hash opcional para usuarios
try:
    from werkzeug.security import generate_password_hash
    CAN_HASH = True
except Exception:
    print("⚠️  werkzeug no disponible: se omite migración a password_hash.")
    CAN_HASH = False


# ------------------ Helpers ------------------
def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def has_column(cur, table, column):
    cur.execute(f"PRAGMA table_info({table})")
    return any(r[1] == column for r in cur.fetchall())

def ensure_table(cur, create_sql):
    cur.execute(create_sql)

def add_column_constant_default_if_missing(cur, table, name, type_sql, default_constant=None):
    if not has_column(cur, table, name):
        if default_constant is None:
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {name} {type_sql}")
        else:
            if isinstance(default_constant, str):
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {name} {type_sql} DEFAULT '{default_constant}'")
            else:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {name} {type_sql} DEFAULT {default_constant}")

def add_timestamp_column_and_backfill_now(cur, table, name):
    if not has_column(cur, table, name):
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {name} TEXT")
    cur.execute(f"UPDATE {table} SET {name} = COALESCE({name}, datetime('now'))")

def ensure_index(cur, name, table, cols):
    cur.execute(f"CREATE INDEX IF NOT EXISTS {name} ON {table} ({cols})")

def ensure_unique_index(cur, name, table, cols):
    cur.execute(f"CREATE UNIQUE INDEX IF NOT EXISTS {name} ON {table} ({cols})")

def add_sucursal_fk(cur, table, default_value=1):
    add_column_constant_default_if_missing(cur, table, "sucursal_id", "INTEGER", default_constant=default_value)
    cur.execute(f"UPDATE {table} SET sucursal_id = COALESCE(sucursal_id, ?) ", (default_value,))
    ensure_index(cur, f"idx_{table}_sucursal", table, "sucursal_id")


# ------------------ Seeds / Migraciones ------------------
def seed_admin(cur):
    cur.execute("SELECT COUNT(*) FROM usuarios")
    count = cur.fetchone()[0]
    if count == 0:
        if has_column(cur, "usuarios", "password_hash") and CAN_HASH:
            pwd = generate_password_hash("fibra123")
            cols, vals = [], []
            seed = {
                "usuario": "admin",
                "password_hash": pwd,
                "email": None,
                "nombre": "Administrador",
                "rol": "admin",
                "foto_url": None,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "sucursal_id": 1,
            }
            for k, v in seed.items():
                if has_column(cur, "usuarios", k):
                    cols.append(k); vals.append(v)
            sql = f"INSERT INTO usuarios ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})"
            cur.execute(sql, vals)
        else:
            cur.execute("INSERT INTO usuarios (usuario, contrasena, sucursal_id) VALUES (?, ?, ?)",
                        ("admin", "fibra123", 1))
        print("✅ Usuario admin creado (usuario: admin / contraseña: fibra123).")
    else:
        print("ℹ️ Ya existe al menos un usuario en la tabla usuarios.")

def migrate_passwords_to_hash(cur):
    if not CAN_HASH or not has_column(cur, "usuarios", "password_hash"):
        return 0
    cur.execute("""
        SELECT id, contrasena FROM usuarios
        WHERE (password_hash IS NULL OR TRIM(password_hash) = '')
          AND contrasena IS NOT NULL AND TRIM(contrasena) <> ''
    """)
    rows = cur.fetchall()
    migrated = 0
    for r in rows:
        new_hash = generate_password_hash(r["contrasena"])
        cur.execute(
            "UPDATE usuarios SET password_hash=?, contrasena='', updated_at=datetime('now') WHERE id=?",
            (new_hash, r["id"])
        )
        migrated += 1
    return migrated

def seed_tecnicos_default(cur):
    """Técnicos iniciales. Solo se usa con una base nueva (ver main)."""
    tecnicos = [
        "Cristhian Alcaraz",
        "Néstor Villalba",
        "Josué Ortiz",
        "Víctor Bobadilla",
        "Robert Leckie",
        "Carlos Balbuena",
    ]
    for sucursal_id in (1, 2):
        for nombre in tecnicos:
            cur.execute("SELECT id FROM tecnicos WHERE nombre=? AND sucursal_id=?", (nombre, sucursal_id))
            if not cur.fetchone():
                cur.execute(
                    "INSERT INTO tecnicos (nombre, activo, sucursal_id) VALUES (?, 1, ?)",
                    (nombre, sucursal_id)
                )
        print(f"✅ Técnicos base cargados en sucursal {sucursal_id}.")

def seed_demo_items(cur, sucursal_id: int = 1):
    cur.execute("SELECT COUNT(*) FROM equipos WHERE sucursal_id=?", (sucursal_id,))
    if cur.fetchone()[0] == 0:
        equipos = [
            ("Router 5G", "Red", "Router de alta velocidad 5G", 5),
            ("Router 2.4", "Red", "Router banda 2.4 GHz", 8),
            ("ONU", "Red", "Unidad de red óptica", 10),
            ("Puntero por cantidad", "Herramienta", "Puntero para señal por cantidad", 20),
            ("Drop por metro", "Cableado", "Cable drop vendido por metro", 100),
            ("Otro", "Varios", "Ítem genérico", 0),
        ]
        cur.executemany(
            "INSERT INTO equipos (nombre, tipo, descripcion, stock, sucursal_id) VALUES (?, ?, ?, ?, ?)",
            [(n, t, d, s, sucursal_id) for (n, t, d, s) in equipos]
        )
        print(f"✅ Insertados {len(equipos)} equipos de prueba en sucursal {sucursal_id}.")
    else:
        cur.execute("UPDATE equipos SET stock = COALESCE(stock, 0) WHERE sucursal_id=?", (sucursal_id,))
        print(f"ℹ️ Equipos ya cargados (sucursal {sucursal_id}).")

    cur.execute("SELECT COUNT(*) FROM herramientas WHERE sucursal_id=?", (sucursal_id,))
    if cur.fetchone()[0] == 0:
        herramientas = [
            ("Crimpadora", "Herramienta de red", "crimpadora.jpg"),
            ("Tester", "Medidor de señal", "tester.jpg"),
            ("Taladro", "Herramienta eléctrica", "taladro.jpg"),
        ]
        cur.executemany(
            "INSERT INTO herramientas (nombre, tipo, imagen, sucursal_id) VALUES (?, ?, ?, ?)",
            [(n, t, img, sucursal_id) for (n, t, img) in herramientas]
        )
        print(f"✅ Insertadas {len(herramientas)} herramientas de prueba en sucursal {sucursal_id}.")
    else:
        print(f"ℹ️ Herramientas ya cargadas (sucursal {sucursal_id}).")

def seed_demo_map(cur, sucursal_id: int = 1):
    cur.execute("SELECT id FROM tecnicos WHERE sucursal_id=? LIMIT 1", (sucursal_id,))
    row_tec = cur.fetchone()
    if not row_tec:
        cur.execute("INSERT INTO tecnicos (nombre, activo, sucursal_id) VALUES ('Juan', 1, ?)", (sucursal_id,))
        tec_id = cur.lastrowid
        print(f"✅ Técnico demo creado (Juan, id={tec_id}, sucursal {sucursal_id}).")
    else:
        tec_id = row_tec["id"]

    cur.execute("SELECT COUNT(*) AS c FROM tecnico_pos WHERE sucursal_id=?", (sucursal_id,))
    if cur.fetchone()[0] == 0:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        points = [(-25.286, -57.645), (-25.290, -57.640), (-25.295, -57.635)]
        for la, ln in points:
            cur.execute("""
                INSERT INTO tecnico_pos (tecnico_id, lat, lng, ts, sucursal_id)
                VALUES (?, ?, ?, ?, ?)
            """, (tec_id, la, ln, now, sucursal_id))
        print(f"✅ Posiciones demo agregadas a tecnico_pos (sucursal {sucursal_id}).")
    else:
        print(f"ℹ️ Ya existen posiciones en tecnico_pos (sucursal {sucursal_id}).")

    cur.execute("SELECT id FROM clientes WHERE nombre='Cliente Demo' AND sucursal_id=?", (sucursal_id,))
    if not cur.fetchone():
        cur.execute("""
            INSERT INTO clientes (nombre, telefono, situacion, activo, sucursal_id)
            VALUES ('Cliente Demo','0981111111','activo',1,?)
        """, (sucursal_id,))
        print(f"✅ Cliente Demo creado (sucursal {sucursal_id}).")
    else:
        print(f"ℹ️ Cliente Demo ya existe (sucursal {sucursal_id}).")

    cur.execute("SELECT COUNT(*) FROM asistencias WHERE sucursal_id=? AND lat IS NOT NULL AND lng IS NOT NULL",
                (sucursal_id,))
    if cur.fetchone()[0] == 0:
        cur.execute("""
            INSERT INTO asistencias (cliente, direccion, tipo, prioridad, tecnico, problema, fecha, pppoe, lat, lng, estado, canal, sucursal_id)
            VALUES ('Cliente Demo', 'Centro Asunción', 'Soporte', 'Media', 'Juan', 'Problema de prueba',
                    datetime('now'), 'cliente@spynet.com', -25.286, -57.645, 'pendiente', 'web', ?)
        """, (sucursal_id,))
        print(f"✅ Ticket demo con coordenadas creado (sucursal {sucursal_id}).")
    else:
        print(f"ℹ️ Ya existen asistencias con coordenadas (sucursal {sucursal_id}).")

def seed_demo_vehiculos(cur, sucursal_id: int = 1):
    cur.execute("SELECT id FROM vehiculos WHERE sucursal_id=? LIMIT 1", (sucursal_id,))
    row = cur.fetchone()
    if not row:
        cur.execute("""
            INSERT INTO vehiculos (placa, imei_gps, lat, lon, last_ts, sucursal_id)
            VALUES ('ABC-123', '359770053123456', -25.287, -57.642, datetime('now'), ?)
        """, (sucursal_id,))
        veh_id = cur.lastrowid
        pts = [(-25.287, -57.642), (-25.290, -57.640), (-25.293, -57.637)]
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for la, lo in pts:
            cur.execute("""
                INSERT INTO vehiculo_posicion (vehiculo_id, lat, lon, ts, sucursal_id)
                VALUES (?, ?, ?, ?, ?)
            """, (veh_id, la, lo, now, sucursal_id))
        print(f"✅ Vehículo demo creado con posiciones (sucursal {sucursal_id}).")
    else:
        print(f"ℹ️ Ya existe al menos un vehículo en sucursal {sucursal_id}.")


# -------- Usuarios únicos por sucursal (migración) --------
def ensure_unique_user_per_branch(cur):
    row = cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='usuarios'").fetchone()
    sql_def = row[0] if row else ""
    if "usuario TEXT UNIQUE" in sql_def:
        cur.execute("""
        CREATE TABLE usuarios_v2 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL,
            contrasena TEXT,
            password_hash TEXT,
            email TEXT,
            nombre TEXT,
            rol TEXT DEFAULT 'operador',
            foto_url TEXT,
            telefono TEXT,
            area TEXT,
            turno TEXT,
            dark_mode INTEGER DEFAULT 0,
            notifs INTEGER DEFAULT 1,
            sucursal_id INTEGER,
            created_at TEXT,
            updated_at TEXT
        )""")
        cur.execute("""
            INSERT INTO usuarios_v2
            (id, usuario, contrasena, password_hash, email, nombre, rol, foto_url, telefono, area, turno, dark_mode, notifs, sucursal_id, created_at, updated_at)
            SELECT id, usuario, contrasena, password_hash, email, nombre, rol, foto_url, telefono, area, turno, dark_mode, notifs, sucursal_id, created_at, updated_at
            FROM usuarios
        """)
        cur.execute("DROP TABLE usuarios")
        cur.execute("ALTER TABLE usuarios_v2 RENAME TO usuarios")

    ensure_unique_index(cur, "idx_usuarios_user_suc", "usuarios", "usuario, sucursal_id")


# -------- Triggers para stock (kardex) --------
def ensure_stock_triggers(cur):
    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_mov_equipos_ai
    AFTER INSERT ON movimientos_equipos
    BEGIN
        UPDATE equipos
           SET stock = COALESCE(stock,0) + CASE WHEN NEW.tipo='ingreso' THEN NEW.cantidad ELSE -NEW.cantidad END
         WHERE id = NEW.equipo_id
           AND (sucursal_id = NEW.sucursal_id OR NEW.sucursal_id IS NULL);
    END;""")

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_mov_equipos_ad
    AFTER DELETE ON movimientos_equipos
    BEGIN
        UPDATE equipos
           SET stock = COALESCE(stock,0) - CASE WHEN OLD.tipo='ingreso' THEN OLD.cantidad ELSE -OLD.cantidad END
         WHERE id = OLD.equipo_id
           AND (sucursal_id = OLD.sucursal_id OR OLD.sucursal_id IS NULL);
    END;""")

    cur.execute("""
    CREATE TRIGGER IF NOT EXISTS trg_mov_equipos_au
    AFTER UPDATE ON movimientos_equipos
    BEGIN
        -- revertimos el viejo
        UPDATE equipos
           SET stock = COALESCE(stock,0) - CASE WHEN OLD.tipo='ingreso' THEN OLD.cantidad ELSE -OLD.cantidad END
         WHERE id = OLD.equipo_id
           AND (sucursal_id = OLD.sucursal_id OR OLD.sucursal_id IS NULL);

        -- aplicamos el nuevo
        UPDATE equipos
           SET stock = COALESCE(stock,0) + CASE WHEN NEW.tipo='ingreso' THEN NEW.cantidad ELSE -NEW.cantidad END
         WHERE id = NEW.equipo_id
           AND (sucursal_id = NEW.sucursal_id OR NEW.sucursal_id IS NULL);
    END;""")

# ------------------ Vehículos / GPS ------------------
def ensure_vehiculos_schema(cur):
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS vehiculos (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      placa      TEXT,
      imei_gps   TEXT,
      lat        REAL,
      lon        REAL,
      last_ts    TEXT,
      tecnico_id INTEGER,
      sucursal_id INTEGER
    )""")
    add_sucursal_fk(cur, "vehiculos")
    ensure_unique_index(cur, "ux_vehiculo_imei", "vehiculos", "imei_gps")
    ensure_index(cur, "idx_vehiculo_tecnico", "vehiculos", "tecnico_id")
    ensure_index(cur, "idx_vehiculo_last_ts", "vehiculos", "last_ts")

    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS vehiculo_posicion (
      id          INTEGER PRIMARY KEY AUTOINCREMENT,
      vehiculo_id INTEGER NOT NULL,
      lat         REAL NOT NULL,
      lon         REAL NOT NULL,
      ts          TEXT NOT NULL,
      velocidad   REAL,
      rumbo       REAL,
      sucursal_id INTEGER
    )""")
    add_sucursal_fk(cur, "vehiculo_posicion")
    ensure_index(cur, "idx_veh_pos_veh_ts", "vehiculo_posicion", "vehiculo_id, ts")


# ------------------ Main migration ------------------
def main():
    print("DB usada por migración:", os.path.abspath(DB_PATH))
    conn = connect()
    cur = conn.cursor()

    # ¿Base nueva? Solo en ese caso se cargan los datos de ejemplo (técnicos,
    # equipos, cliente y ticket demo, vehículo demo).
    # En una base que ya se usa NO se vuelven a insertar: así no reaparecen
    # los técnicos o clientes que se eliminaron desde el sistema.
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='sucursales'")
    base_nueva = cur.fetchone() is None

    # ---------- Sucursales ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS sucursales (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nombre TEXT NOT NULL UNIQUE,
      direccion TEXT,
      telefono TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    )""")

    cur.execute("SELECT COUNT(*) FROM sucursales")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO sucursales (nombre) VALUES (?)", [("Sapucai",), ("Valenzuela",)])
        print("✅ Sucursales creadas: Sapucai (id=1), Valenzuela (id=2)")

    # ---------- Asistencias ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS asistencias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente TEXT,
        direccion TEXT,
        tipo TEXT,
        prioridad TEXT,
        tecnico TEXT,
        problema TEXT,
        fecha TEXT,
        pppoe TEXT
    )""")
    add_column_constant_default_if_missing(cur, "asistencias", "estado", "TEXT", default_constant="pendiente")
    add_column_constant_default_if_missing(cur, "asistencias", "lat", "REAL")
    add_column_constant_default_if_missing(cur, "asistencias", "lng", "REAL")
    add_column_constant_default_if_missing(cur, "asistencias", "cliente_id", "INTEGER")
    add_column_constant_default_if_missing(cur, "asistencias", "cedula", "TEXT")
    add_column_constant_default_if_missing(cur, "asistencias", "programada_en", "TEXT")
    add_column_constant_default_if_missing(cur, "asistencias", "tecnico_id", "INTEGER")
    add_column_constant_default_if_missing(cur, "asistencias", "canal", "TEXT", default_constant="web")
    add_sucursal_fk(cur, "asistencias")
    cur.execute("UPDATE asistencias SET estado='pendiente' WHERE estado IS NULL OR TRIM(estado)=''")
    cur.execute("UPDATE asistencias SET canal='web'      WHERE canal  IS NULL OR TRIM(canal)  =''")

    # ---------- Equipos / Herramientas / Uso ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS equipos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        tipo TEXT NOT NULL,
        descripcion TEXT
    )""")
    add_column_constant_default_if_missing(cur, "equipos", "imagen", "TEXT")
    add_column_constant_default_if_missing(cur, "equipos", "stock", "INTEGER", default_constant=0)
    add_sucursal_fk(cur, "equipos")
    ensure_index(cur, "idx_equipos_stock", "equipos", "stock")

    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS herramientas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        tipo TEXT,
        imagen TEXT
    )""")
    add_sucursal_fk(cur, "herramientas")

    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS uso_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_type TEXT NOT NULL,
        item_id INTEGER NOT NULL,
        tecnico TEXT NOT NULL,
        fecha TEXT NOT NULL,
        servicio TEXT
    )""")
    add_column_constant_default_if_missing(cur, "uso_items", "cantidad", "INTEGER", default_constant=1)
    add_sucursal_fk(cur, "uso_items")
    ensure_index(cur, "idx_uso_items_fecha", "uso_items", "fecha")

    # ---------- Movimientos (kardex) ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS movimientos_equipos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        equipo_id INTEGER NOT NULL,
        tipo TEXT CHECK(tipo IN ('ingreso','egreso')) NOT NULL,
        cantidad INTEGER NOT NULL,
        tecnico TEXT,
        motivo TEXT,
        fecha TEXT NOT NULL DEFAULT (datetime('now')),
        sucursal_id INTEGER
    )""")
    add_sucursal_fk(cur, "movimientos_equipos")
    ensure_index(cur, "idx_movs_fecha", "movimientos_equipos", "fecha")
    ensure_index(cur, "idx_movs_equipo", "movimientos_equipos", "equipo_id")
    ensure_index(cur, "idx_movs_sucursal", "movimientos_equipos", "sucursal_id")
    ensure_stock_triggers(cur)

    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS fotos_asistencia (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        asistencia_id INTEGER,
        tecnico TEXT,
        ruta_foto TEXT NOT NULL,
        descripcion TEXT,
        fecha TEXT
    )""")
    add_sucursal_fk(cur, "fotos_asistencia")

    # ---------- Usuarios ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT NOT NULL,
        contrasena TEXT
    )""")
    add_column_constant_default_if_missing(cur, "usuarios", "password_hash", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "email", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "nombre", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "rol", "TEXT", default_constant="operador")
    add_column_constant_default_if_missing(cur, "usuarios", "foto_url", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "telefono", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "area", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "turno", "TEXT")
    add_column_constant_default_if_missing(cur, "usuarios", "dark_mode", "INTEGER", default_constant=0)
    add_column_constant_default_if_missing(cur, "usuarios", "notifs", "INTEGER", default_constant=1)
    add_sucursal_fk(cur, "usuarios")
    add_timestamp_column_and_backfill_now(cur, "usuarios", "created_at")
    add_timestamp_column_and_backfill_now(cur, "usuarios", "updated_at")

    ensure_unique_user_per_branch(cur)

    seed_admin(cur)
    migrated = migrate_passwords_to_hash(cur)
    if migrated:
        print(f"🔐 Migradas {migrated} contraseñas a password_hash.")

    # ---------- Clientes ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS clientes (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      external_id  TEXT,
      nombre       TEXT,
      referencia   TEXT,
      barrio       TEXT,
      telefono     TEXT,
      situacion    TEXT,
      exonerado    INTEGER DEFAULT 0,
      tipo         TEXT,
      valor        TEXT,
      tipo_valor   TEXT,
      vencimiento  TEXT,
      cedula       TEXT,
      pppoe        TEXT,
      activo       INTEGER DEFAULT 1,
      lat          REAL,
      lng          REAL
    )""")
    add_column_constant_default_if_missing(cur, "clientes", "external_id", "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "nombre",      "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "referencia",  "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "barrio",      "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "telefono",    "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "situacion",   "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "exonerado",   "INTEGER", default_constant=0)
    add_column_constant_default_if_missing(cur, "clientes", "tipo",        "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "valor",       "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "tipo_valor",  "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "vencimiento", "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "cedula",      "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "pppoe",       "TEXT")
    add_column_constant_default_if_missing(cur, "clientes", "activo",      "INTEGER", default_constant=1)
    add_column_constant_default_if_missing(cur, "clientes", "lat", "REAL")
    add_column_constant_default_if_missing(cur, "clientes", "lng", "REAL")
    add_sucursal_fk(cur, "clientes")
    ensure_index(cur, "idx_clientes_latlng", "clientes", "lat, lng")

    # ---------- Técnicos ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS tecnicos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nombre TEXT,
      telefono TEXT,
      activo INTEGER DEFAULT 1
    )""")
    add_column_constant_default_if_missing(cur, "tecnicos", "telefono_whatsapp", "TEXT")
    add_column_constant_default_if_missing(cur, "tecnicos", "movil", "TEXT")
    add_column_constant_default_if_missing(cur, "tecnicos", "tracking_token", "TEXT")
    add_column_constant_default_if_missing(cur, "tecnicos", "lat", "REAL")
    add_column_constant_default_if_missing(cur, "tecnicos", "lng", "REAL")
    add_column_constant_default_if_missing(cur, "tecnicos", "pos_updated_at", "TEXT")
    add_sucursal_fk(cur, "tecnicos")
    ensure_index(cur, "idx_tecnicos_sucursal", "tecnicos", "sucursal_id")
    ensure_unique_index(cur, "ux_tecnico_nombre_suc", "tecnicos", "nombre, sucursal_id")
    ensure_index(cur, "idx_tecnico_token", "tecnicos", "tracking_token")

    # ---------- Tracking de técnicos (histórico) ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS tecnico_tracks (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      tecnico_id INTEGER NOT NULL,
      movil      TEXT,
      lat        REAL NOT NULL,
      lng        REAL NOT NULL,
      accuracy   REAL,
      battery    REAL,
      source     TEXT,
      ts         TEXT NOT NULL
    )""")
    add_sucursal_fk(cur, "tecnico_tracks")
    ensure_index(cur, "idx_tracks_tecnico_ts", "tecnico_tracks", "tecnico_id, ts")

    # ---------- Posiciones puntuales (para el mapa) ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS tecnico_pos (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      tecnico_id INTEGER NOT NULL,
      lat        REAL NOT NULL,
      lng        REAL NOT NULL,
      ts         TEXT NOT NULL
    )""")
    add_sucursal_fk(cur, "tecnico_pos")
    ensure_index(cur, "idx_tecnico_pos_tecnico_ts", "tecnico_pos", "tecnico_id, ts")

    # ---------- ticket_fotos (opcional) ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS ticket_fotos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      ticket_id INTEGER NOT NULL,
      archivo TEXT NOT NULL,
      created_at TEXT
    )""")
    add_timestamp_column_and_backfill_now(cur, "ticket_fotos", "created_at")
    add_sucursal_fk(cur, "ticket_fotos")

    # ---------- NOTIFICACIONES ----------
    ensure_table(cur, """
    CREATE TABLE IF NOT EXISTS notificaciones (
      id          INTEGER PRIMARY KEY AUTOINCREMENT,
      sucursal_id INTEGER NOT NULL,
      titulo      TEXT NOT NULL,
      cuerpo      TEXT,
      creado_en   TEXT
    )""")
    add_sucursal_fk(cur, "notificaciones")
    add_timestamp_column_and_backfill_now(cur, "notificaciones", "creado_en")
    ensure_index(cur, "idx_notif_sucursal_id", "notificaciones", "sucursal_id")
    ensure_index(cur, "idx_notif_creado_en",   "notificaciones", "creado_en")

    # ---------- Vehículos / GPS ----------
    ensure_vehiculos_schema(cur)

    # ---------- Índices útiles ----------
    ensure_index(cur, "idx_asistencias_fecha",   "asistencias", "fecha")
    ensure_index(cur, "idx_asistencias_estado",  "asistencias", "estado")
    ensure_index(cur, "idx_asistencias_prog",    "asistencias", "programada_en")
    ensure_index(cur, "idx_asistencias_tecnico", "asistencias", "tecnico_id")

    ensure_index(cur, "idx_clientes_external", "clientes", "external_id")
    ensure_index(cur, "idx_clientes_tel",      "clientes", "telefono")
    ensure_index(cur, "idx_clientes_cedula",   "clientes", "cedula")
    ensure_index(cur, "idx_clientes_pppoe",    "clientes", "pppoe")

    # ---------- Datos de ejemplo: SOLO con una base nueva ----------
    if base_nueva:
        seed_tecnicos_default(cur)
        seed_demo_items(cur, 1)
        seed_demo_map(cur, 1)
        seed_demo_vehiculos(cur, 1)
    else:
        print("ℹ️ Base existente: no se cargan datos de ejemplo "
              "(se respetan los técnicos y clientes que eliminaste).")

    conn.commit()
    conn.close()
    print("\n✅ Tablas, columnas, triggers e índices verificados/creados. BD lista.")


if __name__ == "__main__":
    main()