from flask import Flask, render_template, request, redirect, url_for, flash
from db import sp_equipo_reasignar, sp_equipo_dar_baja
from db import equipo_upsert_completo

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
        filtro_persona=person or None,
        solo_activos=True  # Solo activos por defecto
    )
       
    print('DISPOSITIVOS:', dispositivos)  # <-- Esto te mostrará los datos en consola
    return render_template('index.html', dispositivos=dispositivos, q=q, person=person)
    
   

# -------- NUEVO: alias útiles ----------
#@app.route('/device')              # /device -> /device/new
#@app.route('/device/')             # /device/ -> /device/new
#def device_alias():
   # return redirect(url_for('new_device'))

# Formulario de nuevo equipo (con varios alias)


@app.route('/device/new', methods=['GET', 'POST'])
@app.route('/new', methods=['GET', 'POST'])
def new_device():
    if request.method == 'POST':
        tag = request.form['tag']
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

        flash('Equipo registrado/actualizado', 'success')
        return redirect(url_for('index'))
    return render_template('new_device.html')
from db import obtener_equipo_por_tag

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
    equipo = obtener_equipo_por_tag(tag)
    return render_template('device.html', tag=tag, historial=historial, equipo=equipo)



@app.post('/device/<tag>/reasignar')
def device_reassign(tag):
    nueva = (request.form.get('nueva_persona') or '').strip()
    cargo = request.form.get('cargo') or None
    fecha = request.form.get('fecha') or None
    registrado_por = request.form.get('registrado_por') or 'Mel'
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
    registrado_por = request.form.get('registrado_por') or 'Mel'
    sp_equipo_dar_baja(tag, motivo, fecha, registrado_por)
    flash(f'Equipo {tag} marcado en BAJA.', 'success')
    return redirect(url_for('device_view', tag=tag))

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


@app.route('/bajas')
def bajas():
    dispositivos = query_dispositivos(solo_activos=False)
    # Filtra solo los que están en BAJA
    dispositivos_baja = [d for d in dispositivos if (d.get('estado') or '').upper() == 'BAJA']
    return render_template('bajas.html', dispositivos=dispositivos_baja)

@app.route('/device/<tag>/edit', methods=['GET', 'POST'])
def edit_device(tag):
    from db import obtener_equipo_por_tag
    equipo = obtener_equipo_por_tag(tag)
    if not equipo:
        flash('Equipo no encontrado.', 'danger')
        return redirect(url_for('index'))

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

if __name__ == '__main__':
    import os
    print('CWD:', os.getcwd())
    print('URL MAP:\n', app.url_map)
    app.run(debug=True)

