import sqlite3

conn = sqlite3.connect("asistencias.db")
cursor = conn.cursor()

# ==========================
# TABLA ASISTENCIAS
# ==========================
cursor.execute("""
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
)
""")

# ==========================
# TABLA EQUIPOS
# ==========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS equipos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    tipo TEXT NOT NULL,
    descripcion TEXT,
    stock INTEGER NOT NULL DEFAULT 0,
    imagen TEXT,
    sucursal_id INTEGER DEFAULT 1
)
""")

# ==========================
# TABLA HERRAMIENTAS
# ==========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS herramientas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    tipo TEXT,
    imagen TEXT,
    stock INTEGER NOT NULL DEFAULT 0,
    sucursal_id INTEGER DEFAULT 1
)
""")

# ==========================
# TABLA USO ITEMS
# ==========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS uso_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_type TEXT NOT NULL,
    item_id INTEGER NOT NULL,
    tecnico TEXT NOT NULL,
    fecha TEXT NOT NULL,
    servicio TEXT,
    cantidad INTEGER DEFAULT 1,
    sucursal_id INTEGER DEFAULT 1
)
""")

# ==========================
# TABLA FOTOS
# ==========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS fotos_asistencia (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asistencia_id INTEGER,
    tecnico TEXT,
    ruta_foto TEXT NOT NULL,
    descripcion TEXT,
    fecha TEXT
)
""")

# ==================================================
# MIGRACIÓN AUTOMÁTICA DE EQUIPOS
# ==================================================

cursor.execute("PRAGMA table_info(equipos)")
columnas = [c[1] for c in cursor.fetchall()]

if "stock" not in columnas:
    cursor.execute("ALTER TABLE equipos ADD COLUMN stock INTEGER NOT NULL DEFAULT 0")
    print("✅ Columna stock agregada a equipos.")

if "imagen" not in columnas:
    cursor.execute("ALTER TABLE equipos ADD COLUMN imagen TEXT")
    print("✅ Columna imagen agregada a equipos.")

if "sucursal_id" not in columnas:
    cursor.execute("ALTER TABLE equipos ADD COLUMN sucursal_id INTEGER DEFAULT 1")
    print("✅ Columna sucursal_id agregada a equipos.")

# ==================================================
# MIGRACIÓN AUTOMÁTICA DE HERRAMIENTAS
# ==================================================

cursor.execute("PRAGMA table_info(herramientas)")
columnas = [c[1] for c in cursor.fetchall()]

if "stock" not in columnas:
    cursor.execute("ALTER TABLE herramientas ADD COLUMN stock INTEGER NOT NULL DEFAULT 0")
    print("✅ Columna stock agregada a herramientas.")

if "sucursal_id" not in columnas:
    cursor.execute("ALTER TABLE herramientas ADD COLUMN sucursal_id INTEGER DEFAULT 1")
    print("✅ Columna sucursal_id agregada a herramientas.")

# ==================================================
# MIGRACIÓN AUTOMÁTICA DE USO_ITEMS
# ==================================================

cursor.execute("PRAGMA table_info(uso_items)")
columnas = [c[1] for c in cursor.fetchall()]

if "cantidad" not in columnas:
    cursor.execute("ALTER TABLE uso_items ADD COLUMN cantidad INTEGER DEFAULT 1")
    print("✅ Columna cantidad agregada.")

if "sucursal_id" not in columnas:
    cursor.execute("ALTER TABLE uso_items ADD COLUMN sucursal_id INTEGER DEFAULT 1")
    print("✅ Columna sucursal_id agregada.")

# ==================================================
# INSERTAR EQUIPOS
# ==================================================

cursor.execute("SELECT COUNT(*) FROM equipos")

if cursor.fetchone()[0] == 0:

    equipos = [

        ("Router 5G", "Red", "Router de alta velocidad 5G", 10),
        ("ONU", "Red", "Unidad de red óptica", 15),
        ("Puntero por cantidad", "Herramienta", "Puntero para señal por cantidad", 20),
        ("Drop por metro", "Cableado", "Cable drop vendido por metro", 500)

    ]

    cursor.executemany("""
        INSERT INTO equipos
        (nombre,tipo,descripcion,stock)
        VALUES (?,?,?,?)
    """, equipos)

    print(f"✅ Insertados {len(equipos)} equipos.")

else:
    print("ℹ️ Equipos ya cargados.")

# ==================================================
# INSERTAR HERRAMIENTAS
# ==================================================

cursor.execute("SELECT COUNT(*) FROM herramientas")

if cursor.fetchone()[0] == 0:

    herramientas = [

        ("Crimpadora", "Herramienta de red", "crimpadora.jpg", 5),
        ("Tester", "Medidor de señal", "tester.jpg", 3),
        ("Taladro", "Herramienta eléctrica", "taladro.jpg", 2)

    ]

    cursor.executemany("""
        INSERT INTO herramientas
        (nombre,tipo,imagen,stock)
        VALUES (?,?,?,?)
    """, herramientas)

    print(f"✅ Insertadas {len(herramientas)} herramientas.")

else:
    print("ℹ️ Herramientas ya cargadas.")

conn.commit()
conn.close()

print("\n✅ Base de datos actualizada correctamente.")