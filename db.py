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

def query_dispositivos(filtro_tag=None, filtro_persona=None, solo_activos=True):
    """
    Devuelve filas con claves en minúscula para que coincidan con el template:
      tag, modelo, serial, ubicacion, personaasignada, estado, fechabaja
    """
    conn = get_connection()
    cur = conn.cursor()

    sql = """
    SELECT
        e.EquipoId                          AS equipoid,
        LTRIM(RTRIM(e.Tag))                 AS tag,
        e.Modelo                            AS modelo,
        e.Serial                            AS serial,
        e.Ubicacion                         AS ubicacion,
        ISNULL(e.Estado,'ACTIVO')           AS estado,
        e.FechaBaja                         AS fechabaja,
        pa.Nombre                           AS personaasignada
    FROM ti.Equipo e
    LEFT JOIN ti.Persona pa ON pa.PersonaId = e.PersonaAsignadaId
    WHERE e.Tag IS NOT NULL AND LTRIM(RTRIM(e.Tag)) <> ''
    """
    params = []
    if solo_activos:
        sql += " AND (e.Estado IS NULL OR e.Estado <> 'BAJA')"
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
    raw = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close(); conn.close()

    # Normaliza las claves a formato esperado por Jinja (camel-case)
    norm_map = {
        'tag': 'Tag',
        'equipoid': 'EquipoId',
        'modelo': 'Modelo',
        'serial': 'Serial',
        'ubicacion': 'Ubicacion',
        'personaasignadaid': 'PersonaAsignadaId',
        'personaasignada': 'PersonaAsignada',
        'cambioid': 'CambioId',
        'tipocambio': 'TipoCambio',
        'descripcion': 'Descripcion',
        'fechacambio': 'FechaCambio',
        'registradopor': 'RegistradoPor',
        'fecharegistro': 'FechaRegistro',
        'personacambioid': 'PersonaCambioId',
        'personacambio': 'PersonaCambio',
    }

    rows = []
    for r in raw:
        norm = {}
        for k, v in r.items():
            key = norm_map.get(k.lower(), k)
            norm[key] = v
        rows.append(norm)

    # Debug para consola
    print(f"[historial_por_equipo] tag={safe_tag!r}, filas={len(rows)}")
    if rows:
        print(" ejemplo:", {k: rows[0].get(k) for k in ("Tag","CambioId","TipoCambio","FechaCambio","RegistradoPor")})

    return rows

def sp_equipo_reasignar(tag, nueva_persona, cargo=None, fecha=None, registrado_por='TI', descripcion=None):
    # normaliza fecha
    import datetime as _dt
    parsed = None
    if fecha:
        txt = fecha.replace('T', ' ')
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
            try:
                parsed = _dt.datetime.strptime(txt, fmt); break
            except ValueError:
                pass

    conn = get_connection(); cur = conn.cursor()
    cur.execute("""
        EXEC ti.sp_Equipo_Reasignar
            @Tag=?, @NuevaPersona=?, @Cargo=?, @FechaCambio=?, @RegistradoPor=?, @Descripcion=?;
    """, (tag, nueva_persona, cargo, parsed, registrado_por, descripcion))
    conn.commit(); cur.close(); conn.close()


def sp_equipo_dar_baja(tag, motivo=None, fecha_baja=None, registrado_por='TI'):
    import datetime as _dt
    parsed = None
    if fecha_baja:
        txt = fecha_baja.replace('T', ' ')
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
            try:
                parsed = _dt.datetime.strptime(txt, fmt); break
            except ValueError:
                pass

    conn = get_connection(); cur = conn.cursor()
    cur.execute("""
        EXEC ti.sp_Equipo_DarBaja
            @Tag=?, @Motivo=?, @FechaBaja=?, @RegistradoPor=?;
    """, (tag, motivo, parsed, registrado_por))
    conn.commit(); cur.close(); conn.close()

def obtener_equipo_por_tag(tag):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT
            e.EquipoId,
            LTRIM(RTRIM(e.Tag)) AS tag,
            e.Modelo           AS modelo,
            e.Marca            AS marca,
            e.Serial           AS serial,
            e.Ubicacion        AS ubicacion,
            e.TipoEquipo       AS tipoequipo,        -- <-- tipo de equipo (Todo en uno, PC...)
            pa.Nombre          AS personaasignada,   -- <-- nombre de la persona asignada
            e.Cargador         AS cargador,
            e.Maletin          AS maletin,
            e.Mouse            AS mouse,
            e.Teclado          AS teclado,
            ISNULL(e.Impresora, 0) AS impresora,     -- opcional: checkbox impresora
            ISNULL(e.Lector, 0)   AS lector,         -- opcional: checkbox lector
            e.Observaciones    AS observaciones
        FROM ti.Equipo e
        LEFT JOIN ti.Persona pa ON pa.PersonaId = e.PersonaAsignadaId
        WHERE LTRIM(RTRIM(e.Tag)) = ?
    """, (tag,))
    row = cur.fetchone()
    equipo = dict(zip([c[0] for c in cur.description], row)) if row else {}
    cur.close(); conn.close()
    return equipo



def equipo_upsert_completo(tag, marca, modelo, serial, ubicacion, persona_asignada,
                           cargador, maletin, mouse, teclado, observaciones,
                           impresora=0, lector=0):
    """
    Upsert del equipo SIN perder la persona asignada.
    Guarda además impresora y lector (bit).
    """
    conn = get_connection()
    cur = conn.cursor()

    # 1) Resolver PersonaAsignadaId (si se dio nombre)
    persona_id = None
    if persona_asignada and persona_asignada.strip():
        cur.execute("""
            DECLARE @PersonaId INT;
            EXEC ti.sp_Persona_Upsert @Nombre=?, @PersonaId=@PersonaId OUTPUT;
            SELECT @PersonaId;
        """, (persona_asignada.strip(),))
        row = cur.fetchone()
        persona_id = row[0] if row else None

    # 2) MERGE con 26 placeholders:
    #    1 (USING) + 12 (UPDATE) + 13 (INSERT) = 26
    sql = """
        MERGE ti.Equipo AS target
        USING (SELECT CAST(? AS NVARCHAR(50)) AS Tag) AS src
        ON (target.Tag = src.Tag)
        WHEN MATCHED THEN
            UPDATE SET
                Marca       = ?,
                Modelo      = ?,
                Serial      = ?,
                Ubicacion   = ?,
                PersonaAsignadaId = COALESCE(?, target.PersonaAsignadaId),
                Cargador    = ?,
                Maletin     = ?,
                Mouse       = ?,
                Teclado     = ?,
                Impresora   = ?,      -- NUEVO
                Lector      = ?,      -- NUEVO
                Observaciones = ?
        WHEN NOT MATCHED THEN
            INSERT (Tag, Marca, Modelo, Serial, Ubicacion, PersonaAsignadaId,
                    Cargador, Maletin, Mouse, Teclado, Impresora, Lector, Observaciones)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """

    params = (
        # USING (1)
        tag,
        # UPDATE (12)
        marca, modelo, serial, ubicacion, persona_id,
        cargador, maletin, mouse, teclado, impresora, lector, observaciones,
        # INSERT (13)
        tag, marca, modelo, serial, ubicacion, persona_id,
        cargador, maletin, mouse, teclado, impresora, lector, observaciones
    )

    cur.execute(sql, params)
    conn.commit()
    cur.close(); conn.close()


def archivo_principal_get(tag: str):
    conn = get_connection(); cur = conn.cursor()
    cur.execute("""
        SELECT TOP 1 Ruta, Nombre
        FROM ti.EquipoArchivo
        WHERE Tag = ? AND EsPrincipal = 1
        ORDER BY EquipoArchivoId DESC
    """, (tag,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if row:
        return {"ruta": row[0], "nombre": row[1]}
    return None

def archivo_principal_set(tag: str, ruta: str, nombre: str):
    conn = get_connection(); cur = conn.cursor()
    # Quita principal previo
    cur.execute("""
        UPDATE ti.EquipoArchivo
           SET EsPrincipal = 0
         WHERE Tag = ? AND EsPrincipal = 1
    """, (tag,))
    # Inserta nuevo principal
    cur.execute("""
        INSERT INTO ti.EquipoArchivo(Tag, Ruta, Nombre, EsPrincipal)
        VALUES(?, ?, ?, 1)
    """, (tag, ruta, nombre))
    conn.commit(); cur.close(); conn.close()
