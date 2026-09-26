import sqlite3

conn = sqlite3.connect("asistencias.db")
cursor = conn.cursor()

try:
    cursor.execute("ALTER TABLE herramientas ADD COLUMN tipo TEXT")
    print("✅ Columna 'tipo' añadida correctamente.")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("ℹ️ La columna 'tipo' ya existe en la tabla.")
    else:
        print("❌ Error:", e)

conn.commit()
conn.close()