from flask import Flask, render_template, request, redirect, url_for, flash, send_file, abort
from werkzeug.utils import secure_filename
import os, time, base64

from db import (
    sp_equipo_upsert,
    sp_equipo_agregar_cambio,
    sp_historial_por_persona,
    query_dispositivos,
    historial_por_equipo,
    sp_equipo_reasignar,
    sp_equipo_dar_baja,
    equipo_upsert_completo,
    obtener_equipo_por_tag,
    archivo_principal_get,
    archivo_principal_set,
)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'cambia_esto_mel'
app.url_map.strict_slashes = False
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['ENABLE_UPLOAD'] = False  # <<— si quieres habilitar subida directa, pon True

# =========================
#  Carpetas de red
# =========================
# Donde escaneamos/buscamos archivos ya existentes para VINCULARLOS (no subirlos).
SHARE_ROOT = r'\\itzamna\DATAUSERS\ADM Y SISTEMAS\Entrega Equipos'

# Donde (si quisieras) se pueden SUBIR/Reemplazar archivos principales manualmente.
UPLOAD_ROOT = r'\\itzamna\DATAUSERS\ADM Y SISTEMAS\Entrega Equipos'

ALLOWED_EXTS = {'pdf', 'jpg', 'jpeg', 'png', 'xlsx', 'xls'}

def _allowed(name: str) -> bool:
    return '.' in name and name.rsplit('.', 1)[-1].lower() in ALLOWED_EXTS

def _ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def _is_inside(path: str, base: str) -> bool:
    """Verifica que 'path' está dentro de 'base' (evita rutas fuera del repositorio)."""
    try:
        return os.path.commonpath([os.path.abspath(path), os.path.abspath(base)]) == os.path.abspath(base)
    except Exception:
        return False

def _encode_path(p: str) -> str:
    return base64.urlsafe_b64encode(p.encode('utf-8')).decode('ascii')

def _decode_path(tok: str) -> str:
    return base64.urlsafe_b64decode(tok.encode('ascii')).decode('utf-8')

def _scan_files_for_tag(tag: str):
    """
    Escanea SHARE_ROOT recursivamente para encontrar archivos que contengan el tag en el nombre.
    No guarda nada; solo ayuda a elegir el archivo correcto.
    """
    out = []
    tag_l = (tag or '').lower()
    if not tag_l or not os.path.isdir(SHARE_ROOT):
        return out
    for root, _, files in os.walk(SHARE_ROOT):
        for nm in files:
            if tag_l in nm.lower() and _allowed(nm):
                full = os.path.join(root, nm)
                try:
                    size = os.path.getsize(full)
                    mtime = time.strftime('%Y-%m-%d %H:%M', time.localtime(os.path.getmtime(full)))
                except OSError:
                    continue
                out.append({"name": nm, "path": full, "size": size, "mtime": mtime, "token": _encode_path(full)})
    # ordena recientes primero
    out.sort(key=lambda x: (x["mtime"], x["name"]), reverse=True)
    return out

# =========================
#  Subida opcional (si quieres además del vínculo)
# =========================
def save_equipo_file_principal(tag: str, file_storage):
    """
    Guarda un ÚNICO archivo principal por equipo en la carpeta raíz de UPLOAD_ROOT.
    Nombre final: <tag>.<ext>  (reemplaza si ya existe).
    También lo deja marcado como principal en BD.
    """
    if not file_storage or not file_storage.filename:
        return None, "Archivo vacío."
    fname = secure_filename(file_storage.filename)
    if not _allowed(fname):
        return None, "Tipo de archivo no permitido."

    _ensure_dir(UPLOAD_ROOT)
    ext = os.path.splitext(fname)[1].lower()
    stored = f"{tag}{ext}"
    path = os.path.join(UPLOAD_ROOT, stored)

    # tamaño (suave)
    try:
        file_storage.stream.seek(0, os.SEEK_END)
        size = file_storage.stream.tell()
        file_storage.stream.seek(0)
    except Exception:
        size = None
    if size is not None and size > 5 * 1024 * 1024:
        return None, "Máximo 5MB por archivo."

    file_storage.save(path)
    # Al subir, también marcamos ese archivo como principal por conveniencia
    archivo_principal_set(tag, path, stored)
    return stored, None

# =========================
#        Rutas
# =========================

@app.route('/')
def index():
    q = request.args.get('q', '').strip()
    person = request.args.get('person', '').strip()
    dispositivos = query_dispositivos(
        filtro_tag=q or None,
        filtro_persona=person or None,
        solo_activos=True
    )
    return render_template('index.html', dispositivos=dispositivos, q=q, person=person)

# --- Nuevo equipo (con vínculo/subida del archivo principal opcional) ---
@app.route('/device/new', methods=['GET', 'POST'])
@app.route('/new', methods=['GET', 'POST'])
def new_device():
    if request.method == 'POST':
        tag = request.form['tag']
        marca = request.form.get('marca')
        modelo = request.form.get('modelo')
        tipo_equipo = request.form.get('tipo_equipo')
        serial = request.form.get('serial')
        ubicacion = request.form.get('ubicacion')
        persona_asignada = request.form.get('persona_asignada')
        cargador = 1 if request.form.get('cargador') else 0
        maletin = 1 if request.form.get('maletin') else 0
        mouse = 1 if request.form.get('mouse') else 0
        teclado = 1 if request.form.get('teclado') else 0
        observaciones = request.form.get('observaciones')

        # 1) Guardamos/actualizamos equipo
        equipo_upsert_completo(
            tag, marca, modelo, serial, ubicacion, persona_asignada,
            cargador, maletin, mouse, teclado, observaciones
        )

        # 2) Procesamos archivo principal OPCIONAL
        ruta_manual = (request.form.get('ruta_principal') or '').strip()
        archivo_subido = request.files.get('archivo') if app.config.get('ENABLE_UPLOAD') else None

        if ruta_manual:
            # Validar que exista, que esté dentro del SHARE_ROOT y que su extensión sea permitida
            if not _is_inside(ruta_manual, SHARE_ROOT):
                flash('La ruta pegada no pertenece a la carpeta de red configurada.', 'danger')
            elif not os.path.isfile(ruta_manual):
                flash('La ruta pegada no existe como archivo.', 'danger')
            elif not _allowed(os.path.basename(ruta_manual)):
                flash('Extensión de archivo no permitida.', 'danger')
            else:
                archivo_principal_set(tag, ruta_manual, os.path.basename(ruta_manual))
                flash('Archivo principal vinculado correctamente.', 'success')
        elif archivo_subido and archivo_subido.filename:
            stored, err = save_equipo_file_principal(tag, archivo_subido)
            if err:
                flash(f'Archivo no subido: {err}', 'danger')
            else:
                flash('Archivo principal subido y vinculado.', 'success')

        flash('Equipo registrado/actualizado', 'success')
        return redirect(url_for('index'))

    return render_template('new_device.html', enable_upload=app.config['ENABLE_UPLOAD'])

# --- Vista dispositivo ---
@app.route('/device/<tag>', methods=['GET', 'POST'])
def device_view(tag):
    if request.method == 'POST':
        tipo = request.form['tipo']
        desc = request.form.get('descripcion') or None
        fecha = request.form.get('fecha') or None
        persona_rel = request.form.get('persona_rel') or None
        registrado_por = request.form.get('registrado_por') or ''
        sp_equipo_agregar_cambio(tag, tipo, desc, fecha, persona_rel, registrado_por)
        flash('Cambio registrado', 'success')
        return redirect(url_for('device_view', tag=tag))

    historial = historial_por_equipo(tag)
    equipo = obtener_equipo_por_tag(tag)

    # Archivo principal actual
    principal = archivo_principal_get(tag)  # {'ruta':..., 'nombre':...} o None

    # Si ya hay principal, ocultamos coincidencias
    if principal:
        files_preview = []
    else:
        files_preview = _scan_files_for_tag(tag)[:5]

    return render_template(
        'device.html',
        tag=tag,
        historial=historial,
        equipo=equipo,
        principal=principal,
        files_preview=files_preview,
        enable_upload=app.config['ENABLE_UPLOAD']
    )

# --- Vincular archivo existente de la red ---
@app.get('/device/<tag>/link')
def device_link(tag):
    candidates = _scan_files_for_tag(tag)
    principal = archivo_principal_get(tag)
    return render_template('device_link.html', tag=tag, candidates=candidates, principal=principal)

@app.post('/device/<tag>/link')
def device_link_save(tag):
    token = request.form.get('token') or ''
    ruta_manual = (request.form.get('ruta_manual') or '').strip()

    if token:
        full = _decode_path(token)
    else:
        full = ruta_manual

    if not full:
        flash('Debes seleccionar o pegar una ruta.', 'danger')
        return redirect(url_for('device_link', tag=tag))

    if not _is_inside(full, SHARE_ROOT):
        flash('La ruta no pertenece a la carpeta de red configurada.', 'danger')
        return redirect(url_for('device_link', tag=tag))

    if not os.path.isfile(full):
        flash('La ruta no existe como archivo.', 'danger')
        return redirect(url_for('device_link', tag=tag))

    nombre = os.path.basename(full)
    if not _allowed(nombre):
        flash('Extensión de archivo no permitida.', 'danger')
        return redirect(url_for('device_link', tag=tag))

    # Persistimos en BD solo la ruta del archivo principal (sin copiar el archivo)
    archivo_principal_set(tag, full, nombre)
    flash('Archivo principal vinculado correctamente.', 'success')
    return redirect(url_for('device_view', tag=tag))

# --- Descargar el principal vinculado ---
@app.get('/device/<tag>/principal/download')
def device_principal_download(tag):
    principal = archivo_principal_get(tag)
    if not principal:
        abort(404)
    path = principal.get('ruta')
    name = principal.get('nombre') or os.path.basename(path)
    if not (path and os.path.isfile(path)):
        abort(404)
    return send_file(path, as_attachment=True, download_name=name)

# --- Subir/Reemplazar principal (opcional) ---
@app.post('/device/<tag>/upload')
def device_upload(tag):
    f = request.files.get('archivo')
    if not f or not f.filename:
        flash('Debes seleccionar un archivo.', 'danger')
        return redirect(url_for('device_view', tag=tag))

    stored, err = save_equipo_file_principal(tag, f)
    if err:
        flash(err, 'danger')
    else:
        flash('Archivo principal subido y vinculado.', 'success')
    return redirect(url_for('device_view', tag=tag))

# --- Reasignar / Baja ---
@app.post('/device/<tag>/reasignar')
def device_reassign(tag):
    nueva = (request.form.get('nueva_persona') or '').strip()
    cargo = request.form.get('cargo') or None
    fecha = request.form.get('fecha') or None
    registrado_por = request.form.get('registrado_por') or ''
    desc = request.form.get('descripcion') or None
    if not nueva:
        flash('Debe indicar el nombre de la nueva persona.', 'danger')
        return redirect(url_for('device_view', tag=tag))
    sp_equipo_reasignar(tag, nueva, cargo, fecha, registrado_por, desc)
    flash(f'Equipo {tag} reasignado a {nueva}.', 'success')
    return redirect(url_for('device_view', tag=tag))

@app.post('/device/<tag>/baja')
def device_baja(tag):
    motivo = request.form.get('motivo') or None
    fecha = request.form.get('fecha') or None
    registrado_por = request.form.get('registrado_por') or ''
    sp_equipo_dar_baja(tag, motivo, fecha, registrado_por)
    flash(f'Equipo {tag} marcado en BAJA.', 'success')
    return redirect(url_for('device_view', tag=tag))

# --- Archivos listados (pantalla general de archivos) ---
@app.get('/device/<tag>/files')
def device_files(tag):
    principal = archivo_principal_get(tag)
    files = []
    if principal:
        p = principal['ruta']
        if os.path.isfile(p):
            files.append({
                'name': principal['nombre'],
                'size': os.path.getsize(p),
                'mtime': time.strftime('%Y-%m-%d %H:%M', time.localtime(os.path.getmtime(p))),
                'where': 'PRINCIPAL'
            })
    return render_template('device_files.html', tag=tag, files=files)

@app.get('/device/<tag>/files/PRINCIPAL/<path:fname>')
def device_download(tag, fname):
    principal = archivo_principal_get(tag)
    if not principal or principal['nombre'] != fname:
        abort(404)
    path = principal['ruta']
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=fname)

# --- Editar equipo ---
@app.route('/device/<tag>/edit', methods=['GET', 'POST'])
def edit_device(tag):
    equipo = obtener_equipo_por_tag(tag)
    if not equipo:
        flash('Equipo no encontrado.', 'danger')
        return redirect(url_for('device_view', tag=tag))

    if request.method == 'POST':
        marca = request.form.get('marca')
        modelo = request.form.get('modelo')
        serial = request.form.get('serial')
        ubicacion = request.form.get('ubicacion')
        persona_asignada = request.form.get('persona_asignada')
        cargador = 1 if request.form.get('cargador') else 0
        maletin = 1 if request.form.get('maletin') else 0
        mouse = 1 if request.form.get('mouse') else 0
        teclado = 1 if request.form.get('teclado') else 0
        observaciones = request.form.get('observaciones')

        equipo_upsert_completo(
            tag, marca, modelo, serial, ubicacion, persona_asignada,
            cargador, maletin, mouse, teclado, observaciones
        )
        flash('Equipo actualizado correctamente.', 'success')
        return redirect(url_for('device_view', tag=tag))

    return render_template('edit_device.html', equipo=equipo, tag=tag)

# --- Búsqueda por persona / listas auxiliares ---
@app.route('/search/person', methods=['GET'])
def search_person():
    nombre = request.args.get('nombre', '').strip()
    resultados = sp_historial_por_persona(nombre) if nombre else []
    return render_template('index.html', resultados_persona=resultados, q='', person=nombre, dispositivos=[])

@app.route('/bajas')
def bajas():
    dispositivos = query_dispositivos(solo_activos=False)
    dispositivos_baja = [d for d in dispositivos if (d.get('estado') or '').upper() == 'BAJA']
    return render_template('bajas.html', dispositivos=dispositivos_baja)

# --- Utilidades ---
@app.route('/_routes')
def _routes():
    return '<pre>' + '\n'.join(str(r) for r in app.url_map.iter_rules()) + '</pre>'

@app.errorhandler(404)
def err404(e):
    print(f"[404] PATH solicitado: {repr(request.path)}")
    return "Not Found", 404

if __name__ == '__main__':
    print('CWD:', os.getcwd())
    print('URL MAP:\n', app.url_map)
    app.run(debug=True)
