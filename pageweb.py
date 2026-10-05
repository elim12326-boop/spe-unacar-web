import os
import re
import time
import urllib.parse
import sqlite3
from io import BytesIO
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv
from google import genai
from google.genai import types

# 1. Cargar las variables del archivo .env en memoria
load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "SPE_UNACAR_SECURE_WEB_KEY_2026")

# Cliente dinámico y seguro de Gemini
def get_gemini_client():
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        return None
    try:
        return genai.Client(api_key=api_key)
    except Exception as e:
        print(f"[ERROR INICIALIZANDO CLIENTE GEMINI]: {e}", flush=True)
        return None

# Rutas de almacenamiento estático y de subidas
UPLOAD_FOLDER = os.path.join('static', 'fotos')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DB_NAME = 'spe_database.db'

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
FOTOS_FOLDER = os.path.join(UPLOADS_DIR, "fotos")
CREDENCIALES_FOLDER = os.path.join(UPLOADS_DIR, "credenciales")

os.makedirs(FOTOS_FOLDER, exist_ok=True)
os.makedirs(CREDENCIALES_FOLDER, exist_ok=True)

ALLOWED_IMG_EXTENSIONS = {"png", "jpg", "jpeg"}
ALLOWED_DOC_EXTENSIONS = {"png", "jpg", "jpeg", "pdf"}

# Catálogo oficial UNACAR
CATALOGO_CARRERAS = {
    # Facultad de Química (Filas 2 a 4)
    "Ingeniería Petrolera": ("Facultad de Química", "Fac_Química"),
    "Ingeniería Química": ("Facultad de Química", "Fac_Química"),
    "Ingeniería Geológica": ("Facultad de Química", "Fac_Química"),

    # Facultad de Ingeniería (Filas 5 a 10)
    "Ingeniería Mecánica": ("Facultad de Ingeniería", "Fac_Ingeniería"),
    "Ingeniería Mecatrónica": ("Facultad de Ingeniería", "Fac_Ingeniería"),
    "Ingeniería Civil": ("Facultad de Ingeniería", "Fac_Ingeniería"),
    "Ingeniería Geofísica": ("Facultad de Ingeniería", "Fac_Ingeniería"),
    "Ingeniería en Energía": ("Facultad de Ingeniería", "Fac_Ingeniería"),
    "Licenciatura en Arquitectura Sustentable": ("Facultad de Ingeniería", "Fac_Ingeniería"),

    # Facultad de Ciencias de la Información (Filas 11 a 14)
    "Ingeniería en Sistemas Computacionales": ("Facultad de Ciencias de la Información", "Fac_Informática"),
    "Ingeniería en Computación": ("Facultad de Ciencias de la Información", "Fac_Informática"),
    "Ingeniería en Diseño Multimedia": ("Facultad de Ciencias de la Información", "Fac_Informática"),
    "Ingeniería en Tecnologías de Cómputo y Comunica": ("Facultad de Ciencias de la Información", "Fac_Informática"),
    "Ingeniería en Tecnologías de Cómputo y Comunicaciones": ("Facultad de Ciencias de la Información", "Fac_Informática"),

    # Facultad de Ciencias de la Salud (Filas 15 a 20)
    "Licenciatura en Medicina": ("Facultad de Ciencias de la Salud", "Fac_Salud"),
    "Licenciatura en Enfermería": ("Facultad de Ciencias de la Salud", "Fac_Salud"),
    "Licenciatura en Nutrición": ("Facultad de Ciencias de la Salud", "Fac_Salud"),
    "Licenciatura en Fisioterapia": ("Facultad de Ciencias de la Salud", "Fac_Salud"),
    "Licenciatura en Psicología Clínica": ("Facultad de Ciencias de la Salud", "Fac_Salud"),
    "Licenciatura en Educación Física y Deporte": ("Facultad de Ciencias de la Salud", "Fac_Salud"),

    # Facultad de Ciencias Naturales (Fila 21)
    "Licenciatura en Biología Marina": ("Facultad de Ciencias Naturales", "Fac_CienciasNaturales"),

    # Facultad de Ciencias Económicas y Administrativas (Filas 22 a 26)
    "Licenciatura en Administración de Empresas": ("Facultad de Ciencias Económicas y Administrativas", "Fac_CEA"),
    "Licenciatura en Contaduría": ("Facultad de Ciencias Económicas y Administrativas", "Fac_CEA"),
    "Licenciatura en Negocios Internacionales": ("Facultad de Ciencias Económicas y Administrativas", "Fac_CEA"),
    "Licenciatura en Mercadotecnia": ("Facultad de Ciencias Económicas y Administrativas", "Fac_CEA"),
    "Licenciatura en Administración Turística": ("Facultad de Ciencias Económicas y Administrativas", "Fac_CEA"),

    # Facultad de Derecho (Filas 27 y 28)
    "Licenciatura en Derecho": ("Facultad de Derecho", "Fac_Derecho"),
    "Licenciatura en Criminología y Criminalística": ("Facultad de Derecho", "Fac_Derecho"),

    # Facultad de Ciencias Educativas (Filas 29 a 31)
    "Licenciatura en Educación": ("Facultad de Ciencias Educativas", "Fac_Educación"),
    "Licenciatura en Lengua Inglesa": ("Facultad de Ciencias Educativas", "Fac_Educación"),
    "Licenciatura en Comunicación y Gestión Cultural": ("Facultad de Ciencias Educativas", "Fac_Educación")
}

CONDICIONES_DEFAULT = """✔ Válido al presentar tu membresía física de SPE vigente.
✔ Aplica de forma presencial en el establecimiento participante.
✔ No acumulable con otras promociones o descuentos vigentes."""

def get_db_connection():
    conn = sqlite3.connect(DB_NAME, timeout=20.0)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Tabla de Socios Activos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS socios (
            id TEXT PRIMARY KEY,
            nombre TEXT,
            carrera TEXT,
            estatus TEXT,
            vigencia TEXT,
            foto TEXT
        )
    ''')
    
    # 2. Tabla de Solicitudes de Credencial Web
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS solicitudes_credencial (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_registro TEXT,
            nombre TEXT,
            matricula TEXT UNIQUE,
            correo TEXT,
            telefono TEXT,
            carrera TEXT,
            facultad TEXT,
            semestre TEXT,
            foto_rostro TEXT,
            credencial_inst TEXT,
            origen TEXT DEFAULT 'Web',
            estatus TEXT DEFAULT 'Pendiente de Validación',
            notas TEXT DEFAULT ''
        )
    ''')

    # 3. Tabla de Locales y Convenios
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS locales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT,
            descuento TEXT,
            categoria TEXT,
            direccion TEXT,
            descripcion TEXT,
            foto TEXT,
            foto_portada TEXT,
            foto_perfil TEXT,
            telefono TEXT,
            facebook TEXT,
            instagram TEXT,
            tiktok TEXT,
            mapa_url TEXT,
            condiciones TEXT,
            titulo_especiales TEXT,
            horarios TEXT,
            rango_precio TEXT,
            servicios_extra TEXT,
            archivo_menu TEXT,
            enlace_citas TEXT,
            fecha_registro TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fotos_locales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            local_id INTEGER,
            foto TEXT,
            FOREIGN KEY (local_id) REFERENCES locales (id) ON DELETE CASCADE
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS especiales_locales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            local_id INTEGER,
            titulo TEXT,
            descripcion TEXT,
            precio TEXT,
            foto TEXT,
            terminos TEXT,
            FOREIGN KEY (local_id) REFERENCES locales (id) ON DELETE CASCADE
        )
    ''')

    columnas_nuevas = [
        ('foto', 'TEXT'),
        ('foto_portada', 'TEXT'),
        ('foto_perfil', 'TEXT'),
        ('telefono', 'TEXT'),
        ('facebook', 'TEXT'),
        ('instagram', 'TEXT'),
        ('tiktok', 'TEXT'),
        ('mapa_url', 'TEXT'),
        ('condiciones', 'TEXT'),
        ('titulo_especiales', 'TEXT'),
        ('horarios', 'TEXT'),
        ('rango_precio', 'TEXT'),
        ('servicios_extra', 'TEXT'),
        ('archivo_menu', 'TEXT'),
        ('enlace_citas', 'TEXT'),
        ('fecha_registro', 'TEXT')
    ]
    for col_nombre, col_tipo in columnas_nuevas:
        try:
            cursor.execute(f'ALTER TABLE locales ADD COLUMN {col_nombre} {col_tipo}')
        except sqlite3.OperationalError:
            pass

    try:
        cursor.execute('ALTER TABLE especiales_locales ADD COLUMN terminos TEXT')
    except sqlite3.OperationalError:
        pass

    conn.commit()
    conn.close()

init_db()

def archivo_permitido(filename, extensiones):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in extensiones

def registrar_solicitud_db(datos, nombre_archivo_foto, nombre_archivo_cred):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        existente = cursor.execute('SELECT id FROM solicitudes_credencial WHERE matricula = ?', (datos["matricula"],)).fetchone()
        if existente:
            conn.close()
            return False, f"La matrícula {datos['matricula']} ya se encuentra registrada en la base de datos."

        facultad, _ = CATALOGO_CARRERAS[datos["carrera"]]
        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute('''
            INSERT INTO solicitudes_credencial (
                fecha_registro, nombre, matricula, correo, telefono,
                carrera, facultad, semestre, foto_rostro, credencial_inst,
                origen, estatus, notas
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            fecha_actual, datos["nombre"], datos["matricula"], datos["correo"],
            datos["telefono"], datos["carrera"], facultad, datos["semestre"],
            nombre_archivo_foto, nombre_archivo_cred, 'Web', 'Pendiente de Validación', ''
        ))
        conn.commit()
        conn.close()
        return True, "¡Tu solicitud ha sido registrada con éxito en el sistema!"
    except Exception as e:
        conn.close()
        return False, "El sistema de registro no se encuentra disponible en este momento. Por favor, intenta en otro momento."

def procesar_url_mapa(mapa_input, direccion=""):
    if not mapa_input or mapa_input.strip() == "":
        query = direccion if direccion else "Ciudad del Carmen Campeche"
        return f"https://maps.google.com/maps?q={urllib.parse.quote(query)}&t=&z=15&ie=UTF8&iwloc=&output=embed"

    mapa_input = mapa_input.strip()

    match_src = re.search(r'src=["\']([^"\']+)["\']', mapa_input)
    if match_src:
        return match_src.group(1)

    if "/maps/embed" in mapa_input:
        return mapa_input

    if "maps.app.goo.gl" in mapa_input or "google.com/maps" in mapa_input:
        query = direccion if direccion else mapa_input
    else:
        query = mapa_input

    return f"https://maps.google.com/maps?q={urllib.parse.quote(query)}&t=&z=15&ie=UTF8&iwloc=&output=embed"

@app.route('/')
def inicio():
    return render_template('index.html')

@app.route('/directorio')
def directorio():
    categoria_filtro = request.args.get('categoria', '')
    busqueda = request.args.get('q', '')

    conn = get_db_connection()
    query = 'SELECT * FROM locales WHERE 1=1'
    params = []

    if categoria_filtro:
        query += ' AND categoria = ?'
        params.append(categoria_filtro)
    if busqueda:
        query += ' AND (nombre LIKE ? OR descripcion LIKE ?)'
        params.append(f'%{busqueda}%')
        params.append(f'%{busqueda}%')

    rows = conn.execute(query, params).fetchall()
    cat_rows = conn.execute('SELECT DISTINCT categoria FROM locales WHERE categoria IS NOT NULL AND categoria != ""').fetchall()
    categorias = [row['categoria'] for row in cat_rows]
    conn.close()

    limite_nuevo = datetime.now() - timedelta(days=30)
    locales = []
    for r in rows:
        d = dict(r)
        es_nuevo = False
        f_str = d.get('fecha_registro')
        if f_str:
            try:
                f_dt = datetime.strptime(f_str[:10], '%Y-%m-%d')
                if f_dt >= limite_nuevo:
                    es_nuevo = True
            except Exception:
                pass
        else:
            es_nuevo = True
        d['es_nuevo'] = es_nuevo
        locales.append(d)

    return render_template('directorio.html', locales=locales, categorias=categorias, selected_categoria=categoria_filtro, search_query=busqueda)

@app.route('/perfil/<int:local_id>')
def perfil(local_id):
    conn = get_db_connection()
    local = conn.execute('SELECT * FROM locales WHERE id = ?', (local_id,)).fetchone()
    
    if local is None:
        conn.close()
        return redirect(url_for('directorio'))

    fotos_galeria = conn.execute('SELECT foto FROM fotos_locales WHERE local_id = ?', (local_id,)).fetchall()
    especiales = conn.execute('SELECT * FROM especiales_locales WHERE local_id = ?', (local_id,)).fetchall()
    
    locales_similares = conn.execute(
        'SELECT * FROM locales WHERE categoria = ? AND id != ? LIMIT 3',
        (local['categoria'], local_id)
    ).fetchall()

    conn.close()

    condiciones_texto = local['condiciones'] if local['condiciones'] and local['condiciones'].strip() != '' else CONDICIONES_DEFAULT
    condiciones_lista = [linea.strip() for linea in condiciones_texto.split('\n') if linea.strip() != '']

    servicios_lista = []
    if local['servicios_extra']:
        servicios_lista = [s.strip() for s in local['servicios_extra'].split(',') if s.strip()]

    nombre_enc = urllib.parse.quote(local['nombre'])
    mensaje_wpp = f"¡Hola! Vi su convenio de {nombre_enc} en la plataforma UniDescuentos SPE y me gustaría más información..."
    wpp_url = f"https://wa.me/{local['telefono']}?text={urllib.parse.quote(mensaje_wpp)}" if local['telefono'] else "#"

    if local['direccion']:
        maps_direct_url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(local['direccion'])}"
    else:
        maps_direct_url = local['mapa_url'] if local['mapa_url'] else "#"

    return render_template('perfil.html', 
                           local=local, 
                           fotos_galeria=fotos_galeria, 
                           especiales=especiales, 
                           condiciones_lista=condiciones_lista, 
                           servicios_lista=servicios_lista, 
                           wpp_url=wpp_url, 
                           maps_direct_url=maps_direct_url, 
                           locales_similares=locales_similares)

@app.route('/credencial', methods=['GET', 'POST'])
def credencial():
    if request.method == 'POST':
        nombre = request.form.get('nombre', '').strip().title()
        matricula = request.form.get('matricula', '').strip().replace(' ', '')
        correo = request.form.get('correo', '').strip().lower().replace(' ', '')
        telefono = request.form.get('telefono', '').strip()
        carrera = request.form.get('carrera', '').strip()
        semestre = request.form.get('semestre', '').strip()

        if not (nombre and matricula and correo and telefono and carrera and semestre):
            flash("Por favor completa todos los campos obligatorios.", "warning")
            return redirect(url_for('credencial'))

        if carrera not in CATALOGO_CARRERAS:
            flash("La carrera seleccionada no es válida.", "danger")
            return redirect(url_for('credencial'))

        archivo_foto = request.files.get('foto')
        archivo_credencial = request.files.get('credencial')

        if not archivo_foto or not archivo_credencial:
            flash("Ambos archivos (Foto y Credencial) son obligatorios.", "warning")
            return redirect(url_for('credencial'))

        if not archivo_permitido(archivo_foto.filename, ALLOWED_IMG_EXTENSIONS):
            flash("Formato de foto no válido. Solo se admiten JPG, JPEG o PNG.", "danger")
            return redirect(url_for('credencial'))

        if not archivo_permitido(archivo_credencial.filename, ALLOWED_DOC_EXTENSIONS):
            flash("Formato de credencial no válido. Solo se admiten JPG, PNG o PDF.", "danger")
            return redirect(url_for('credencial'))

        ext_foto = archivo_foto.filename.rsplit(".", 1)[1].lower()
        nombre_archivo_foto = f"foto_{matricula}_{int(datetime.now().timestamp())}.{ext_foto}"
        ruta_guardada_foto = os.path.join(FOTOS_FOLDER, nombre_archivo_foto)
        archivo_foto.save(ruta_guardada_foto)

        ext_cred = archivo_credencial.filename.rsplit(".", 1)[1].lower()
        nombre_archivo_cred = f"credencial_{matricula}_{int(datetime.now().timestamp())}.{ext_cred}"
        ruta_guardada_cred = os.path.join(CREDENCIALES_FOLDER, nombre_archivo_cred)
        archivo_credencial.save(ruta_guardada_cred)

        datos_alumno = {
            "nombre": nombre,
            "matricula": matricula,
            "correo": correo,
            "telefono": telefono,
            "carrera": carrera,
            "semestre": semestre
        }

        try:
            exito, mensaje = registrar_solicitud_db(datos_alumno, nombre_archivo_foto, nombre_archivo_cred)
            if exito:
                flash(mensaje, "success")
            else:
                tipo_alerta = "warning" if "ya se encuentra registrada" in mensaje else "danger"
                flash(mensaje, tipo_alerta)
        except Exception:
            flash("El sistema de registro no se encuentra disponible en este momento. Por favor, intenta en otro momento.", "danger")

        return redirect(url_for('credencial'))

    return render_template('credencial.html')

@app.route('/validar')
def validar_socio():
    socio_id = request.args.get('id')
    socio = None
    if socio_id:
        conn = get_db_connection()
        socio = conn.execute('SELECT * FROM socios WHERE id = ?', (socio_id,)).fetchone()
        conn.close()
    return render_template('validar.html', socio=socio)

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    conn = get_db_connection()

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'guardar_socio':
            socio_id = request.form['id']
            nombre = request.form['nombre']
            carrera = request.form['carrera']
            estatus = request.form['estatus']
            vigencia = request.form['vigencia']
            foto = request.files.get('foto')
            filename = None

            if foto and foto.filename != '':
                ext = os.path.splitext(foto.filename)[1]
                filename = secure_filename(f"{socio_id}{ext}")
                foto.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

            existente = conn.execute('SELECT foto FROM socios WHERE id = ?', (socio_id,)).fetchone()
            if existente and not filename:
                filename = existente['foto']

            conn.execute('''
                INSERT INTO socios (id, nombre, carrera, estatus, vigencia, foto)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    nombre=excluded.nombre, carrera=excluded.carrera,
                    estatus=excluded.estatus, vigencia=excluded.vigencia,
                    foto=COALESCE(excluded.foto, socios.foto)
            ''', (socio_id, nombre, carrera, estatus, vigencia, filename))
            conn.commit()

        elif action == 'eliminar_socio':
            socio_id = request.form['socio_id']
            row = conn.execute('SELECT foto FROM socios WHERE id = ?', (socio_id,)).fetchone()
            if row and row['foto']:
                foto_path = os.path.join(app.config['UPLOAD_FOLDER'], row['foto'])
                if os.path.exists(foto_path):
                    os.remove(foto_path)
            conn.execute('DELETE FROM socios WHERE id = ?', (socio_id,))
            conn.commit()

        elif action == 'guardar_local':
            local_id = request.form.get('local_id')
            nombre = request.form['nombre_local']
            descuento = request.form['descuento']
            categoria = request.form['categoria']
            direccion = request.form['direccion']
            descripcion = request.form['descripcion']
            telefono = request.form.get('telefono', '').strip()
            facebook = request.form.get('facebook', '').strip()
            instagram = request.form.get('instagram', '').strip()
            tiktok = request.form.get('tiktok', '').strip()
            mapa_raw = request.form.get('mapa_url', '').strip()
            mapa_url = procesar_url_mapa(mapa_raw, direccion)
            
            condiciones = request.form.get('condiciones', '').strip()
            if not condiciones:
                condiciones = CONDICIONES_DEFAULT

            titulo_especiales = request.form.get('titulo_especiales', '').strip()
            if not titulo_especiales:
                titulo_especiales = "Especiales para Estudiantes"

            horarios = request.form.get('horarios', '').strip()
            rango_precio = request.form.get('rango_precio', '').strip()
            servicios_extra = request.form.get('servicios_extra', '').strip()
            enlace_citas = request.form.get('enlace_citas', '').strip()

            foto_portada = request.files.get('foto_portada')
            foto_perfil = request.files.get('foto_perfil')
            archivo_menu_file = request.files.get('archivo_menu')
            archivos_galeria = request.files.getlist('fotos_galeria')

            filename_portada = None
            filename_perfil = None
            filename_menu = None
            safe_name = secure_filename(nombre.lower())

            if foto_portada and foto_portada.filename != '':
                ext = os.path.splitext(foto_portada.filename)[1]
                filename_portada = secure_filename(f"portada_{safe_name}{ext}")
                foto_portada.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_portada))

            if foto_perfil and foto_perfil.filename != '':
                ext = os.path.splitext(foto_perfil.filename)[1]
                filename_perfil = secure_filename(f"perfil_{safe_name}{ext}")
                foto_perfil.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_perfil))

            if archivo_menu_file and archivo_menu_file.filename != '':
                ext = os.path.splitext(archivo_menu_file.filename)[1]
                filename_menu = secure_filename(f"menu_{safe_name}{ext}")
                archivo_menu_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_menu))

            fecha_actual_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            if local_id and local_id.strip() != '':
                existente = conn.execute('SELECT foto_portada, foto_perfil, archivo_menu, fecha_registro FROM locales WHERE id = ?', (local_id,)).fetchone()
                if existente:
                    if not filename_portada:
                        filename_portada = existente['foto_portada']
                    if not filename_perfil:
                        filename_perfil = existente['foto_perfil']
                    if not filename_menu:
                        filename_menu = existente['archivo_menu']
                    fecha_reg = existente['fecha_registro'] if existente['fecha_registro'] else fecha_actual_str
                else:
                    fecha_reg = fecha_actual_str

                conn.execute('''
                    UPDATE locales 
                    SET nombre=?, descuento=?, categoria=?, direccion=?, descripcion=?, foto_portada=?, foto_perfil=?, telefono=?, facebook=?, instagram=?, tiktok=?, mapa_url=?, condiciones=?, titulo_especiales=?, horarios=?, rango_precio=?, servicios_extra=?, archivo_menu=?, enlace_citas=?, fecha_registro=?
                    WHERE id=?
                ''', (nombre, descuento, categoria, direccion, descripcion, filename_portada, filename_perfil, telefono, facebook, instagram, tiktok, mapa_url, condiciones, titulo_especiales, horarios, rango_precio, servicios_extra, filename_menu, enlace_citas, fecha_reg, local_id))
                target_local_id = int(local_id)
            else:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO locales (nombre, descuento, categoria, direccion, descripcion, foto_portada, foto_perfil, telefono, facebook, instagram, tiktok, mapa_url, condiciones, titulo_especiales, horarios, rango_precio, servicios_extra, archivo_menu, enlace_citas, fecha_registro)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (nombre, descuento, categoria, direccion, descripcion, filename_portada, filename_perfil, telefono, facebook, instagram, tiktok, mapa_url, condiciones, titulo_especiales, horarios, rango_precio, servicios_extra, filename_menu, enlace_citas, fecha_actual_str))
                target_local_id = cursor.lastrowid

            for idx, g_file in enumerate(archivos_galeria):
                if g_file and g_file.filename != '':
                    ext = os.path.splitext(g_file.filename)[1]
                    g_filename = secure_filename(f"galeria_{target_local_id}_{idx}_{g_file.filename}")
                    g_file.save(os.path.join(app.config['UPLOAD_FOLDER'], g_filename))
                    conn.execute('INSERT INTO fotos_locales (local_id, foto) VALUES (?, ?)', (target_local_id, g_filename))

            esp_titulos = request.form.getlist('esp_titulo[]')
            esp_descs = request.form.getlist('esp_desc[]')
            esp_precios = request.form.getlist('esp_precio[]')
            esp_terminos = request.form.getlist('esp_terminos[]')
            esp_fotos = request.files.getlist('esp_foto[]')
            esp_existentes = request.form.getlist('esp_foto_existente[]')

            conn.execute('DELETE FROM especiales_locales WHERE local_id = ?', (target_local_id,))

            for i in range(len(esp_titulos)):
                titulo_i = esp_titulos[i].strip()
                if titulo_i:
                    desc_i = esp_descs[i].strip() if i < len(esp_descs) else ""
                    precio_i = esp_precios[i].strip() if i < len(esp_precios) else ""
                    terminos_i = esp_terminos[i].strip() if i < len(esp_terminos) else ""
                    
                    foto_esp_name = esp_existentes[i] if i < len(esp_existentes) else ""
                    if i < len(esp_fotos) and esp_fotos[i] and esp_fotos[i].filename != '':
                        f_esp = esp_fotos[i]
                        ext_e = os.path.splitext(f_esp.filename)[1]
                        foto_esp_name = secure_filename(f"esp_{target_local_id}_{i}_{f_esp.filename}")
                        f_esp.save(os.path.join(app.config['UPLOAD_FOLDER'], foto_esp_name))

                    conn.execute('''
                        INSERT INTO especiales_locales (local_id, titulo, descripcion, precio, foto, terminos)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (target_local_id, titulo_i, desc_i, precio_i, foto_esp_name, terminos_i))

            conn.commit()

        elif action == 'eliminar_local':
            local_id = request.form['local_id']
            row = conn.execute('SELECT foto_portada, foto_perfil, archivo_menu FROM locales WHERE id = ?', (local_id,)).fetchone()
            if row:
                for img_key in ['foto_portada', 'foto_perfil', 'archivo_menu']:
                    if row[img_key]:
                        f_path = os.path.join(app.config['UPLOAD_FOLDER'], row[img_key])
                        if os.path.exists(f_path):
                            os.remove(f_path)
            
            g_rows = conn.execute('SELECT foto FROM fotos_locales WHERE local_id = ?', (local_id,)).fetchall()
            for g in g_rows:
                if g['foto']:
                    f_path = os.path.join(app.config['UPLOAD_FOLDER'], g['foto'])
                    if os.path.exists(f_path):
                        os.remove(f_path)

            e_rows = conn.execute('SELECT foto FROM especiales_locales WHERE local_id = ?', (local_id,)).fetchall()
            for e in e_rows:
                if e['foto']:
                    f_path = os.path.join(app.config['UPLOAD_FOLDER'], e['foto'])
                    if os.path.exists(f_path):
                        os.remove(f_path)

            conn.execute('DELETE FROM fotos_locales WHERE local_id = ?', (local_id,))
            conn.execute('DELETE FROM especiales_locales WHERE local_id = ?', (local_id,))
            conn.execute('DELETE FROM locales WHERE id = ?', (local_id,))
            conn.commit()

        conn.close()
        return redirect(url_for('admin'))

    socios = conn.execute('SELECT * FROM socios').fetchall()
    locales = conn.execute('SELECT * FROM locales').fetchall()

    especiales_map = {}
    for l in locales:
        rows = conn.execute('SELECT titulo, descripcion, precio, foto, terminos FROM especiales_locales WHERE local_id = ?', (l['id'],)).fetchall()
        especiales_map[l['id']] = [dict(r) for r in rows]

    conn.close()
    return render_template('admin.html', socios=socios, locales=locales, especiales_map=especiales_map, condiciones_default=CONDICIONES_DEFAULT)

@app.route('/admin/exportar_excel')
def exportar_excel():
    conn = get_db_connection()
    solicitudes = conn.execute('SELECT * FROM solicitudes_credencial ORDER BY id ASC').fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BD_General"

    encabezados = [
        "ID", "Fecha Registro", "Nombre Completo", "Matrícula / ID",
        "Correo Universitario", "Teléfono", "Carrera", "Facultad Asignada",
        "Semestre", "Foto Rostro", "Credencial Inst", "Origen", "Estatus Credencial", "Notas"
    ]
    ws.append(encabezados)

    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_num in range(1, len(encabezados) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for s in solicitudes:
        ws.append([
            s["id"], s["fecha_registro"], s["nombre"], s["matricula"],
            s["correo"], s["telefono"], s["carrera"], s["facultad"],
            s["semestre"], s["foto_rostro"], s["credencial_inst"],
            s["origen"], s["estatus"], s["notas"]
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    wb.close()

    fecha_descarga = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Reporte_Solicitudes_SPE_{fecha_descarga}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Caché en memoria para soportar alto tráfico sin saturar la base de datos
_CACHE_INSTRUCCIONES = {"texto": None, "timestamp": 0}

def obtener_instrucciones_asistente():
    """Lee la base de datos con caché de 300s para rendimiento masivo y aprendizaje autónomo."""
    ahora = time.time()
    if _CACHE_INSTRUCCIONES["texto"] and (ahora - _CACHE_INSTRUCCIONES["timestamp"] < 300):
        return _CACHE_INSTRUCCIONES["texto"]

    lista_comercios = ""
    try:
        conn = get_db_connection()
        locales = conn.execute("SELECT nombre, categoria, descuento, direccion, horarios, telefono, descripcion FROM locales").fetchall()
        conn.close()
        
        if locales:
            lineas = []
            for l in locales:
                lineas.append(f"- {l['nombre']} ({l['categoria']}): Descuento: {l['descuento']}. Ubicación: {l['direccion']}. Horario: {l['horarios']}. Teléfono/WhatsApp: {l['telefono']}. Detalles: {l['descripcion']}")
            lista_comercios = "\n".join(lineas)
        else:
            lista_comercios = "Actualmente se están actualizando los convenios en el sistema."
    except Exception as e:
        print(f"[AVISO ASISTENTE DB]: {e}", flush=True)
        lista_comercios = "Directorio de convenios afiliados a SPE UNACAR en Ciudad del Carmen."

    prompt = f"""
Eres el Asistente Virtual Oficial de SPE Student Chapter UNACAR (Universidad Autónoma del Carmen) en Ciudad del Carmen.
Tu labor es responder dudas a los estudiantes sobre convenios, descuentos y la credencial de membresía.

INFORMACIÓN DE LA PLATAFORMA:
1. Propósito: Brindar descuentos exclusivos a estudiantes universitarios afiliados mediante su credencial física oficial de SPE.
2. Trámite de membresía (/credencial):
   - Requisitos: Comprobante de inscripción o reinscripción UNACAR vigente, fotografía de rostro reciente y correo universitario activo (@mail.unacar.mx).
   - Datos solicitados: Nombre completo, matrícula, carrera, semestre y número de teléfono.
3. Uso de beneficios: El estudiante debe presentar su credencial física oficial de SPE antes de pagar o pedir la cuenta en el local participante.

COMERCIOS REGISTRADOS ACTUALMENTE:
{lista_comercios}

REGLAS PARA TUS RESPUESTAS:
- Sé amable, claro, conciso y con tono universitario formal pero accesible.
- Si te preguntan por un negocio en específico, indica su descuento, dirección y horario.
- Si te piden sugerencias (ejemplo: 'dónde reparar una pantalla' o 'servicios de belleza'), recomienda los comercios de esa categoría.
- No inventes descuentos ni negocios que no existan en la lista.
- Presenta tus respuestas con viñetas limpias y negritas en los nombres clave para facilitar la lectura.
"""
    _CACHE_INSTRUCCIONES["texto"] = prompt
    _CACHE_INSTRUCCIONES["timestamp"] = ahora
    return prompt

@app.route('/api/chat', methods=['POST'])
def api_chat():
    """Ruta con fallback dinámico entre modelos y reintentos para 100% de disponibilidad."""
    client = get_gemini_client()
    if not client:
        print("[ERROR]: No se detectó GEMINI_API_KEY o GOOGLE_API_KEY.", flush=True)
        return jsonify({'respuesta': 'El asistente no tiene configurada la clave de acceso a la API.'}), 500

    datos = request.get_json() or {}
    mensaje = datos.get('mensaje', '').strip()

    if not mensaje:
        return jsonify({'error': 'El mensaje está vacío'}), 400

    instrucciones = obtener_instrucciones_asistente()
    
    # Modelos oficiales de producción en Google AI Studio en orden de preferencia
    modelos_disponibles = ['gemini-2.5-flash', 'gemini-1.5-flash']
    
    for modelo in modelos_disponibles:
        for intento in range(2):
            try:
                respuesta_ia = client.models.generate_content(
                    model=modelo,
                    contents=mensaje,
                    config=types.GenerateContentConfig(
                        system_instruction=instrucciones,
                        temperature=0.3,
                        max_output_tokens=800
                    )
                )
                if respuesta_ia and respuesta_ia.text:
                    return jsonify({'respuesta': respuesta_ia.text})
            except Exception as e:
                error_str = str(e)
                print(f"[AVISO IA - {modelo} - Intento {intento+1}]: {error_str}", flush=True)
                time.sleep(1.0)
                continue

    return jsonify({'respuesta': 'En este momento hay varias consultas simultáneas en la plataforma. Por favor, realiza tu pregunta nuevamente en un momento.'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
