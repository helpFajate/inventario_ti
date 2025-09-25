from flask import Flask, render_template, request, redirect, url_for, flash
from db import (
    sp_equipo_upsert,
    sp_equipo_agregar_cambio,
    sp_historial_por_persona,
    query_dispositivos,
    historial_por_equipo,
)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'cambia_esto_mel'
app.url_map.strict_slashes = False  # acepta /ruta y /ruta/

@app.route('/')
def index():
    q = request.args.get('q', '').strip()
    person = request.args.get('person', '').strip()
    dispositivos = query_dispositivos(
        filtro_tag=q or None,
        filtro_persona=person or None
    )
    return render_template('index.html', dispositivos=dispositivos, q=q, person=person)

# -------- NUEVO: alias útiles ----------
@app.route('/device')              # /device -> /device/new
@app.route('/device/')             # /device/ -> /device/new
def device_alias():
    return redirect(url_for('new_device'))

# Formulario de nuevo equipo (con varios alias)
@app.route('/device/new', methods=['GET', 'POST'])
@app.route('/device/new/', methods=['GET', 'POST'])
@app.route('/new', methods=['GET', 'POST'])  # /new como acceso corto
def new_device():
    if request.method == 'POST':
        tag = request.form['tag'].strip()
        modelo = request.form.get('modelo') or None
        serial = request.form.get('serial') or None
        ubicacion = request.form.get('ubicacion') or None
        persona = request.form.get('persona') or None
        cargo = request.form.get('cargo') or None  # NUEVO
        eid = sp_equipo_upsert(tag, modelo, serial, ubicacion, persona, cargo)
        flash(f'Equipo {tag} registrado/actualizado (ID {eid})', 'success')
        return redirect(url_for('device_view', tag=tag))
    return render_template('new_device.html')

# Detalle / historial por equipo
@app.route('/device/<tag>', methods=['GET', 'POST'])
def device_view(tag):
    if request.method == 'POST':
        tipo = request.form['tipo']
        desc = request.form.get('descripcion') or None
        fecha = request.form.get('fecha') or None
        persona_rel = request.form.get('persona_rel') or None
        registrado_por = request.form.get('registrado_por') or 'Mel'
        sp_equipo_agregar_cambio(tag, tipo, desc, fecha, persona_rel, registrado_por)
        flash('Cambio registrado', 'success')
        return redirect(url_for('device_view', tag=tag))
    historial = historial_por_equipo(tag)
    return render_template('device.html', tag=tag, historial=historial)

# Búsqueda por persona
@app.route('/search/person', methods=['GET'])
def search_person():
    nombre = request.args.get('nombre', '').strip()
    resultados = sp_historial_por_persona(nombre) if nombre else []
    return render_template('index.html', resultados_persona=resultados, q='', person=nombre, dispositivos=[])

# Diagnóstico: lista de rutas
@app.route('/_routes')
def _routes():
    return '<pre>' + '\n'.join(str(r) for r in app.url_map.iter_rules()) + '</pre>'

# 404: muestra el path exacto (para detectar espacios o caracteres)
@app.errorhandler(404)
def err404(e):
    print(f"[404] PATH solicitado: {repr(request.path)}")
    return "Not Found", 404

if __name__ == '__main__':
    import os
    print('CWD:', os.getcwd())
    print('URL MAP:\n', app.url_map)
    app.run(debug=True)
