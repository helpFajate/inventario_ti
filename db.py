import os
from dotenv import load_dotenv
import datetime as _dt

# Fallback: intenta pyodbc y si no, usa pypyodbc con el mismo alias
try:
    import pyodbc
except ImportError:
    import pypyodbc as pyodbc

load_dotenv()

def get_connection():
    return pyodbc.connect(
        f"DRIVER={{{os.getenv('SQL_DRIVER', 'ODBC Driver 17 for SQL Server')}}};"
        f"SERVER={os.getenv('SQL_SERVER')};"
        f"DATABASE={os.getenv('SQL_DATABASE', 'Reportes')};"
        f"Trusted_Connection={os.getenv('SQL_TRUSTED','yes')};"
    )

# ---------- SP helpers ----------

def sp_equipo_upsert(tag, modelo=None, serial=None, ubicacion=None, persona_asignada=None, cargo=None):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        DECLARE @EquipoId INT;
        EXEC ti.sp_Equipo_Upsert
            @Tag=?, @Modelo=?, @Serial=?, @Ubicacion=?,
            @PersonaAsignadaNombre=?, @Cargo=?,           -- << nuevo parámetro
            @EquipoId=@EquipoId OUTPUT;
        SELECT @EquipoId AS EquipoId;
    """, (tag, modelo, serial, ubicacion, persona_asignada, cargo))
    row = cur.fetchone()
    eid = row[0] if row else None
    conn.commit(); cur.close(); conn.close()
    return eid


def sp_equipo_agregar_cambio(tag, tipo, desc=None, fecha=None, persona_rel=None, registrado_por='TI'):
    """
    Agrega un cambio al historial.
    - fecha: puede venir como '', 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM', o 'YYYY-MM-DDTHH:MM'
             Aquí la convertimos a datetime o la dejamos en None.
    """
    # --- Normalizar fecha ---
    if isinstance(fecha, str):
        fecha = fecha.strip()
        if not fecha:
            fecha = None
        else:
            # Aceptar 'YYYY-MM-DDTHH:MM' o 'YYYY-MM-DD HH:MM[:SS]'
            txt = fecha.replace('T', ' ')
            parsed = None
            for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
                try:
                    parsed = _dt.datetime.strptime(txt, fmt)
                    break
                except ValueError:
                    pass
            fecha = parsed  # si no parsea, queda None

    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        EXEC ti.sp_Equipo_AgregarCambio
            @Tag=?, @TipoCambio=?, @Descripcion=?, @FechaCambio=?, @PersonaRelacionada=?, @RegistradoPor=?;
    """, (tag, tipo, desc, fecha, persona_rel, registrado_por))
    conn.commit()
    cur.close()
    conn.close()


def sp_historial_por_persona(nombre):
    conn = get_connection(); cur = conn.cursor()
    cur.execute("EXEC ti.sp_Historial_PorPersona @Nombre=?", (nombre,))
    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close(); conn.close()
    return rows

# ---------- Consultas directas sobre la vista ----------

def query_dispositivos(filtro_tag=None, filtro_persona=None):
    conn = get_connection()
    cur = conn.cursor()

    sql = """
    SELECT
        e.EquipoId,
        e.Tag AS Activo,               -- << alias aquí
        e.Modelo,
        e.Serial,
        e.Ubicacion,
        pa.Nombre AS PersonaAsignada
    FROM ti.Equipo e
    LEFT JOIN ti.Persona pa ON pa.PersonaId = e.PersonaAsignadaId
    WHERE e.Tag IS NOT NULL AND LTRIM(RTRIM(e.Tag)) <> ''
    """
    params = []

    if filtro_tag:
        sql += " AND e.Tag LIKE ?"
        params.append(f"%{filtro_tag}%")
    if filtro_persona:
        sql += " AND pa.Nombre LIKE ?"
        params.append(f"%{filtro_persona}%")

    sql += " ORDER BY e.Tag"

    cur.execute(sql, tuple(params))
    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close(); conn.close()
    return rows


def historial_por_equipo(tag):
    """
    Devuelve historial por Activo (Tag).
    - Trim a ambos lados por si hay espacios.
    - Ordena con NULLS al final para que no rompa si no hay cambios.
    """
    safe_tag = (tag or "").strip()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM ti.v_EquipoHistorial
        WHERE LTRIM(RTRIM(Tag)) = ?
        ORDER BY
            CASE WHEN FechaCambio IS NULL THEN 1 ELSE 0 END,
            FechaCambio DESC,
            CambioId DESC
    """, (safe_tag,))

    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]

    # --- debug al terminal ---
    print(f"[historial_por_equipo] tag={safe_tag!r} filas={len(rows)}")
    if rows:
        print("  ejemplo:", {k: rows[0].get(k) for k in ("Tag","CambioId","TipoCambio","FechaCambio","RegistradoPor")})

    cur.close(); conn.close()
    return rows
