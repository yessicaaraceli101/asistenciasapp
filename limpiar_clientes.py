import sqlite3, os
db = os.path.join(os.path.dirname(__file__), "asistencias.db")
con = sqlite3.connect(db)
con.execute("DELETE FROM clientes WHERE IFNULL(nombre,'')='' AND IFNULL(telefono,'')='' AND IFNULL(external_id,'')=''")
con.commit(); con.close()
print("Limpieza OK")