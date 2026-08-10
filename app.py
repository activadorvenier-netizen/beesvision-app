# app.py - VERSIÓN COMPLETA CORREGIDA
from __future__ import annotations
from pathlib import Path
from typing import Any
import pandas as pd
from flask import Flask, jsonify, render_template, request, session, send_file
from functools import wraps
import hashlib
from datetime import datetime
import io
from models import ReviewDatabase
import os
import time
import re

# =====================================================
# IMPORTS PARA GOOGLE SHEETS
# =====================================================
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False
    print("⚠️ gspread no instalado. La funcionalidad de Google Sheets no estará disponible.")
    print("   Instala con: pip install gspread oauth2client")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Base de datos
db = ReviewDatabase()

# Supervisores (SIN CONTRASEÑA)
SUPERVISORS = {
    14: "Bruno Del Popolo",
    17: "Franco Vivani", 
    41: "Claudio Raposo",
}

app = Flask(__name__)
app.secret_key = "clave_secreta_para_desarrollo"

# Cache de datos
_cached_df = None
_current_excel = None
_data_loaded = False
_current_mes = None

# =====================================================
# CONFIGURACIÓN DE GOOGLE SHEETS
# =====================================================

# ⚠️ CAMBIA ESTOS VALORES CON LOS TUYOS ⚠️
GOOGLE_SHEETS_CONFIG = {
    "sheet_id": "TU_SHEET_ID_AQUI",      # El ID de tu Google Sheet
    "sheet_name": "Clientes",             # Nombre de la hoja
    "credentials_file": "credentials.json" # Archivo de credenciales
}

# Cache de clientes
_client_cache = {}
_client_cache_time = 0
CACHE_TTL = 600  # 10 minutos

def get_google_sheet_client():
    """Obtener conexión a Google Sheets"""
    if not GOOGLE_SHEETS_AVAILABLE:
        return None
    
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            GOOGLE_SHEETS_CONFIG["credentials_file"], scope
        )
        return gspread.authorize(creds)
    except Exception as e:
        print(f"❌ Error al conectar con Google Sheets: {e}")
        return None

def load_client_master():
    """
    Cargar el maestro de clientes desde Google Sheets con caché.
    El sheet debe tener columnas: ClienteID, Nombre, Direccion
    """
    global _client_cache, _client_cache_time
    
    if not GOOGLE_SHEETS_AVAILABLE:
        return {}
    
    now = time.time()
    
    # Si el caché es válido, devolverlo
    if _client_cache and (now - _client_cache_time) < CACHE_TTL:
        return _client_cache
    
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return {}
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        worksheet = sheet.worksheet(GOOGLE_SHEETS_CONFIG["sheet_name"])
        
        # Obtener todos los datos
        records = worksheet.get_all_records()
        
        # Construir diccionario ClienteID -> datos
        client_master = {}
        for row in records:
            cliente_id = str(row.get("ClienteID", "")).strip()
            if cliente_id:
                client_master[cliente_id] = {
                    "nombre": row.get("Nombre", ""),
                    "direccion": row.get("Direccion", "")
                }
        
        _client_cache = client_master
        _client_cache_time = now
        
        print(f"✅ Clientes cargados desde Google Sheets: {len(client_master)}")
        return client_master
        
    except Exception as e:
        print(f"❌ Error al cargar maestro de clientes: {e}")
        print("   Verifica que:")
        print("   1. El archivo credentials.json existe y es válido")
        print("   2. El Google Sheet está compartido con la cuenta de servicio")
        print("   3. El sheet_id es correcto")
        return {}

def extract_short_poc_id(poc_id):
    """
    Extraer el código de cliente del POC ID completo.
    El prefijo fijo es '0538220000' (10 dígitos).
    También maneja POC IDs que comienzan con '53822' (sin el cero inicial).
    Ejemplos:
    - 05382200005533 -> 5533
    - 5382200005533 -> 5533
    - 05382200008232 -> 8232
    - 5382200008232 -> 8232
    """
    if not poc_id:
        return ""
    
    # Convertir a string
    poc_id = str(poc_id).strip()
    
    # Si tiene .0 al final, removerlo
    if poc_id.endswith('.0'):
        poc_id = poc_id[:-2]
    
    # Si tiene guiones o espacios, limpiar
    poc_id = poc_id.replace('-', '').replace(' ', '')
    
    # Prefijos posibles
    PREFIX_FULL = "0538220000"  # 10 dígitos
    PREFIX_SHORT = "53822"      # 5 dígitos (sin el cero inicial)
    
    # Si el ID comienza con el prefijo completo (10 dígitos)
    if poc_id.startswith(PREFIX_FULL):
        return poc_id[len(PREFIX_FULL):]
    
    # Si el ID comienza con el prefijo corto (5 dígitos)
    if poc_id.startswith(PREFIX_SHORT):
        return poc_id[len(PREFIX_SHORT):]
    
    # Si es más corto que el prefijo, devolverlo tal cual
    if len(poc_id) < len(PREFIX_FULL):
        return poc_id
    
    # Si no coincide con ningún prefijo, intentar extraer los últimos números
    match = re.search(r'0*(\d+)$', poc_id)
    if match:
        return match.group(1)
    
    return poc_id

def get_client_info(poc_id):
    """
    Obtener información de un cliente por POC ID.
    Primero extrae el código corto, luego busca en el maestro.
    """
    if not poc_id:
        return None
    
    # Extraer el código corto
    short_id = extract_short_poc_id(poc_id)
    if not short_id:
        return None
    
    # Buscar en el maestro de clientes
    master = load_client_master()
    return master.get(short_id)

# =====================================================
# DECORADORES
# =====================================================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'supervisor_id' not in session:
            return jsonify({"error": "No autorizado"}), 401
        return f(*args, **kwargs)
    return decorated_function

# =====================================================
# HELPERS
# =====================================================

def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()

def _is_visita_valida(value: Any) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, (int, float)):
        return float(value) == 1.0
    if isinstance(value, bool):
        return value
    text = str(value).strip().upper()
    return text in {"VERDADERO", "TRUE", "1", "1.0", "SI", "YES"}

def _build_task_id(row: pd.Series) -> str:
    """Crear ID único para cada tarea"""
    img_value = clean_text(row.get("Img", ""))
    poc_id = clean_text(row.get("POC ID", ""))
    fecha = str(row.get("Fecha", ""))
    
    parts = [img_value, poc_id, fecha]
    raw = "|".join(parts)
    return hashlib.md5(raw.encode()).hexdigest()[:12]

def get_mes_actual():
    """Obtener el mes actual en formato YYYYMM"""
    return datetime.now().strftime("%Y%m")

def formatear_fecha(fecha_valor):
    """
    Formatear fecha a DD/MM/YYYY.
    Maneja: datetime, string YYYY-MM-DD, YYYYMMDD, o timestamp de Excel
    """
    if pd.isna(fecha_valor):
        return ""
    
    # Si es datetime
    if isinstance(fecha_valor, (datetime, pd.Timestamp)):
        return fecha_valor.strftime("%d/%m/%Y")
    
    # Si es string
    fecha_str = str(fecha_valor).strip()
    
    # Si ya tiene barras (DD/MM/YYYY), devolver igual
    if '/' in fecha_str:
        return fecha_str
    
    # Si es YYYY-MM-DD (ej: 2026-08-03)
    if '-' in fecha_str:
        partes = fecha_str.split('-')
        if len(partes) == 3:
            año = partes[0]
            mes = partes[1]
            dia = partes[2].split(' ')[0]  # Quitar la hora si existe
            return f"{dia}/{mes}/{año}"
    
    # Si es YYYYMMDD (ej: 20260803)
    if len(fecha_str) >= 8 and fecha_str[:8].isdigit():
        año = fecha_str[0:4]
        mes = fecha_str[4:6]
        dia = fecha_str[6:8]
        return f"{dia}/{mes}/{año}"
    
    # Si es número de Excel
    try:
        fecha_int = int(float(fecha_str))
        if fecha_int > 40000:  # Número serial de Excel
            fecha_obj = datetime(1899, 12, 30) + pd.Timedelta(days=fecha_int)
            return fecha_obj.strftime("%d/%m/%Y")
    except:
        pass
    
    return fecha_str

# =====================================================
# LOAD DATA
# =====================================================

def _load_data(path: Path) -> pd.DataFrame:
    """Cargar y filtrar datos del Excel"""
    global _current_excel, _data_loaded, _current_mes
    
    df = pd.read_excel(path, engine="openpyxl")
    _current_excel = path.name
    _current_mes = get_mes_actual()
    
    print(f"📊 Columnas encontradas: {df.columns.tolist()}")
    print(f"📊 Total de filas: {len(df)}")
    
    # Verificar columnas requeridas
    required = [
        "Fecha", "Promotor", "POC ID", "Detalle Tarea", 
        "Img", "Completada", "Validada", "Visita Valida", 
        "Supervisor ID"
    ]
    
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"⚠️ Columnas faltantes: {missing}")
        raise ValueError(f"Columnas faltantes en el archivo: {missing}")
    
    df = df.copy()
    
    # Limpiar datos
    df["Completada"] = pd.to_numeric(df["Completada"], errors="coerce")
    df["Validada"] = pd.to_numeric(df["Validada"], errors="coerce")
    df["Supervisor ID"] = pd.to_numeric(df["Supervisor ID"], errors="coerce").astype("Int64")
    df["VisitaValidaBool"] = df["Visita Valida"].apply(_is_visita_valida)
    
    print(f"📊 Supervisores encontrados: {df['Supervisor ID'].unique().tolist()}")
    
    # Filtrar tareas: Completada=1, Validada=0, Visita Valida=VERDADERO
    filtered = df[
        (df["Completada"] == 1.0) &
        (df["Validada"] == 0.0) &
        (df["VisitaValidaBool"])
    ].copy()
    
    print(f"📊 Después de filtros básicos: {len(filtered)} filas")
    
    # Filtrar por supervisor en sesión
    if 'supervisor_id' in session:
        supervisor_id = session['supervisor_id']
        filtered = filtered[filtered["Supervisor ID"] == supervisor_id]
        print(f"📊 Filtrado por supervisor {supervisor_id}: {len(filtered)} filas")
    
    if filtered.empty:
        print("⚠️ No hay datos después del filtro.")
        print(f"   - Supervisor ID en sesión: {session.get('supervisor_id')}")
    
    # IDs únicos
    filtered = filtered.reset_index(drop=True)
    filtered["row_id"] = range(1, len(filtered) + 1)
    filtered["task_id"] = filtered.apply(_build_task_id, axis=1)
    
    _data_loaded = True
    
    return filtered

# =====================================================
# ROUTES
# =====================================================

@app.route("/")
def index():
    """Página principal"""
    return render_template("index.html", supervisors=SUPERVISORS)

@app.route("/api/login", methods=["POST"])
def api_login():
    """Login de supervisor (SIN CONTRASEÑA)"""
    data = request.json
    supervisor_id = data.get("supervisor_id")
    
    if supervisor_id not in SUPERVISORS:
        return jsonify({"error": "Supervisor no encontrado"}), 404
    
    session['supervisor_id'] = supervisor_id
    session['supervisor_name'] = SUPERVISORS[supervisor_id]
    
    # Forzar recarga de datos si existe archivo
    global _cached_df, _data_loaded
    file_path = UPLOAD_DIR / "data.xlsx"
    if file_path.exists():
        try:
            _cached_df = _load_data(file_path)
            _data_loaded = True
        except Exception as e:
            print(f"Error recargando datos: {e}")
    
    return jsonify({
        "ok": True,
        "supervisor_id": supervisor_id,
        "supervisor_name": SUPERVISORS[supervisor_id]
    })

@app.route("/api/logout", methods=["POST"])
def api_logout():
    """Cerrar sesión"""
    session.clear()
    return jsonify({"ok": True})

@app.route("/api/has_file")
@login_required
def api_has_file():
    """Verificar si hay archivo cargado"""
    global _data_loaded
    file_exists = (UPLOAD_DIR / "data.xlsx").exists()
    return jsonify({
        "has_file": file_exists and _data_loaded,
        "file_exists": file_exists,
        "data_loaded": _data_loaded,
        "mes_actual": _current_mes
    })

@app.route("/api/upload", methods=["POST"])
@login_required
def api_upload():
    """Subir archivo Excel"""
    global _cached_df, _data_loaded, _current_mes
    
    if "file" not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400
    
    f = request.files["file"]
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "El archivo debe ser .xlsx"}), 400
    
    file_path = UPLOAD_DIR / "data.xlsx"
    f.save(str(file_path))
    
    try:
        _cached_df = _load_data(file_path)
        _data_loaded = True
        _current_mes = get_mes_actual()
        
        import sqlite3
        with sqlite3.connect(db.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO uploaded_files (filename, row_count, supervisor_id, mes_revision)
                VALUES (?, ?, ?, ?)
            """, (file_path.name, len(_cached_df), session['supervisor_id'], _current_mes))
            conn.commit()
        
        return jsonify({
            "ok": True,
            "rows": len(_cached_df),
            "message": f"Archivo cargado con {len(_cached_df)} registros",
            "mes": _current_mes
        })
    except Exception as e:
        _cached_df = None
        _data_loaded = False
        return jsonify({"error": str(e)}), 500

@app.route("/api/tasks")
@login_required
def api_tasks():
    """Obtener tareas filtradas con estado de revisión"""
    global _cached_df, _data_loaded
    
    if not _data_loaded or _cached_df is None or _cached_df.empty:
        file_path = UPLOAD_DIR / "data.xlsx"
        if file_path.exists():
            try:
                _cached_df = _load_data(file_path)
                _data_loaded = True
            except Exception as e:
                return jsonify({
                    "error": f"Error cargando datos: {str(e)}",
                    "no_data": True
                }), 404
        else:
            return jsonify({
                "error": "No hay datos cargados. Por favor, sube un archivo Excel.",
                "no_data": True
            }), 404
    
    supervisor_id = session['supervisor_id']
    start_date = request.args.get("start_date", type=str)
    end_date = request.args.get("end_date", type=str)
    
    # Filtrar por supervisor
    result = _cached_df[_cached_df["Supervisor ID"] == supervisor_id].copy()
    
    if result.empty:
        return jsonify({
            "error": f"No hay tareas asignadas para {SUPERVISORS[supervisor_id]}",
            "no_tasks": True
        }), 404
    
    # Filtrar por fechas
    if start_date:
        try:
            # Convertir fecha a string para comparar
            result = result[result["Fecha"].astype(str) >= start_date]
        except:
            pass
    if end_date:
        try:
            result = result[result["Fecha"].astype(str) <= end_date]
        except:
            pass
    
    if result.empty:
        return jsonify({
            "error": "No hay tareas en el rango de fechas seleccionado",
            "no_tasks": True
        }), 404
    
    # Obtener estados de revisión
    task_ids = result["task_id"].tolist()
    pending_tasks = db.get_pending_tasks(supervisor_id, task_ids)
    
    # Preparar respuesta
    response_rows = []
    for _, row in result.iterrows():
        task_id = row["task_id"]
        is_reviewed = task_id not in pending_tasks
        
        review = None
        if is_reviewed:
            review = db.get_review_status(task_id, supervisor_id)
        
        # Obtener POC ID original
        raw_poc_id = clean_text(row.get("POC ID"))
        
        # Extraer el código corto
        short_poc_id = extract_short_poc_id(raw_poc_id)
        
        # Obtener información del cliente desde Google Sheets
        client_info = get_client_info(raw_poc_id) if raw_poc_id else None
        
        # Obtener fecha
        fecha_formateada = formatear_fecha(row.get("Fecha"))
        
        # Obtener URL de imagen (usar Img o TaskImageUrl)
        img_url = clean_text(row.get("Img", ""))
        if not img_url:
            img_url = clean_text(row.get("TaskImageUrl", ""))
        
        response_rows.append({
            "row_id": int(row["row_id"]),
            "task_id": task_id,
            "fecha": fecha_formateada,
            "promotor": clean_text(row.get("Promotor")),
            "poc_id": short_poc_id,
            "poc_id_completo": raw_poc_id,
            "cliente_id": short_poc_id,  # ClienteID desde Sheets
            "razon_social": client_info.get("nombre") if client_info else None,
            "direccion": client_info.get("direccion") if client_info else None,
            "detalle_tarea": clean_text(row.get("Detalle Tarea")),
            "imagen": img_url,
            "revisado": is_reviewed,
            "status": review.get("status") if review else None,
            "observaciones": review.get("observaciones") if review else "",
            "fecha_revision": review.get("fecha_revision") if review else None
        })
    
    return jsonify(response_rows)

@app.route("/api/save_review", methods=["POST"])
@login_required
def api_save_review():
    """Guardar revisión de una tarea"""
    data = request.json
    
    task_id = data.get("task_id")
    status = data.get("status")
    observaciones = data.get("observaciones", "")
    
    if not task_id or not status:
        return jsonify({"error": "Faltan datos"}), 400
    
    if status not in ["objeccion", "invalida", "fraude"]:
        return jsonify({"error": "Status inválido"}), 400
    
    if _cached_df is None or not _data_loaded:
        return jsonify({"error": "No hay datos cargados"}), 404
    
    task_row = _cached_df[_cached_df["task_id"] == task_id]
    if task_row.empty:
        return jsonify({"error": "Tarea no encontrada"}), 404
    
    row_id = int(task_row.iloc[0]["row_id"])
    supervisor_id = session['supervisor_id']
    supervisor_name = session['supervisor_name']
    
    success = db.save_review(
        task_id=task_id,
        row_id=row_id,
        supervisor_id=supervisor_id,
        supervisor_name=supervisor_name,
        status=status,
        observaciones=observaciones,
        excel_file=_current_excel or "data.xlsx",
        mes_revision=_current_mes or get_mes_actual()
    )
    
    if success:
        return jsonify({"ok": True, "message": "Revisión guardada correctamente"})
    else:
        return jsonify({"error": "Error al guardar la revisión"}), 500

@app.route("/api/delete_review/<task_id>", methods=["DELETE"])
@login_required
def api_delete_review(task_id):
    """Eliminar una revisión (para poder cambiarla)"""
    try:
        supervisor_id = session.get('supervisor_id')
        
        if _cached_df is None or not _data_loaded:
            return jsonify({"error": "No hay datos cargados"}), 404
        
        task_row = _cached_df[_cached_df["task_id"] == task_id]
        if task_row.empty:
            return jsonify({"error": "Tarea no encontrada"}), 404
        
        success = db.delete_review(task_id, supervisor_id)
        
        if success:
            return jsonify({"ok": True, "message": "Revisión eliminada correctamente"})
        else:
            return jsonify({"error": "Error al eliminar la revisión"}), 500
            
    except Exception as e:
        print(f"Error al eliminar revisión: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/stats")
@login_required
def api_stats():
    """Estadísticas del supervisor"""
    global _data_loaded
    
    supervisor_id = session['supervisor_id']
    mes = _current_mes or get_mes_actual()
    stats = db.get_supervisor_stats(supervisor_id, mes)
    
    if not _data_loaded or _cached_df is None:
        return jsonify({
            "total_revisados": stats["total"],
            "total_pendientes": 0,
            "by_status": stats["by_status"],
            "porcentaje": 0,
            "no_data": True,
            "message": "Sin datos cargados. Sube un archivo Excel para ver el progreso completo."
        })
    
    total_disponibles = len(_cached_df[_cached_df["Supervisor ID"] == supervisor_id])
    
    task_ids = _cached_df[_cached_df["Supervisor ID"] == supervisor_id]["task_id"].tolist()
    pending_tasks = db.get_pending_tasks(supervisor_id, task_ids)
    total_pendientes = len(pending_tasks)
    
    return jsonify({
        "total_revisados": stats["total"],
        "total_pendientes": total_pendientes,
        "by_status": stats["by_status"],
        "porcentaje": round((stats["total"] / total_disponibles * 100) if total_disponibles > 0 else 0, 1),
        "total_disponibles": total_disponibles,
        "has_data": True,
        "mes": mes
    })

@app.route("/api/supervisors")
def api_supervisors():
    """Lista de supervisores"""
    items = [{"id": sid, "name": name} for sid, name in SUPERVISORS.items()]
    return jsonify(items)


# =====================================================
# EXPORTAR REVISIONES (REPORTE SEMANAL)
# =====================================================

@app.route("/api/export_reviews", methods=["GET"])
@login_required
def api_export_reviews():
    """Exportar todas las revisiones a Excel con todos los campos"""
    try:
        supervisor_id = session.get('supervisor_id')
        supervisor_name = session.get('supervisor_name')
        mes = _current_mes or get_mes_actual()
        
        reviews = db.get_all_reviews(supervisor_id, mes)
        
        if not reviews:
            return jsonify({"error": "No hay revisiones para exportar"}), 404
        
        if _cached_df is None or not _data_loaded:
            export_data = []
            for review in reviews:
                export_data.append({
                    "Fecha Ejecución": "",
                    "POC ID": "",
                    "Cliente ID": "",
                    "Razón Social": "",
                    "Direccion": "",
                    "Detalle Tarea": "",
                    "Imagen": "",
                    "Promotor": "",
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        else:
            task_data = {}
            
            for _, row in _cached_df.iterrows():
                raw_poc_id = clean_text(row.get("POC ID", ""))
                short_poc_id = extract_short_poc_id(raw_poc_id)
                client_info = get_client_info(raw_poc_id) if raw_poc_id else None
                img_url = clean_text(row.get("Img", ""))
                if not img_url:
                    img_url = clean_text(row.get("TaskImageUrl", ""))
                
                task_data[row["task_id"]] = {
                    "fecha": formatear_fecha(row.get("Fecha")),
                    "promotor": row.get("Promotor", ""),
                    "poc_id": short_poc_id,
                    "poc_id_completo": raw_poc_id,
                    "cliente_id": short_poc_id,
                    "razon_social": client_info.get("nombre") if client_info else "",
                    "direccion": client_info.get("direccion") if client_info else "",
                    "detalle_tarea": row.get("Detalle Tarea", ""),
                    "imagen": img_url
                }
            
            export_data = []
            for review in reviews:
                task_id = review["task_id"]
                task_info = task_data.get(task_id, {})
                
                export_data.append({
                    "Fecha Ejecución": task_info.get("fecha", ""),
                    "POC ID": task_info.get("poc_id", ""),
                    "POC ID Completo": task_info.get("poc_id_completo", ""),
                    "Cliente ID": task_info.get("cliente_id", ""),
                    "Razón Social": task_info.get("razon_social", ""),
                    "Direccion": task_info.get("direccion", ""),
                    "Detalle Tarea": task_info.get("detalle_tarea", ""),
                    "Imagen": task_info.get("imagen", ""),
                    "Promotor": task_info.get("promotor", ""),
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, sheet_name='Revisiones', index=False)
            
            if not df_export.empty and 'Status' in df_export.columns:
                df_status = df_export.groupby('Status').size().reset_index(name='Cantidad')
                df_status = df_status.sort_values('Cantidad', ascending=False)
                total_row = pd.DataFrame({
                    'Status': ['TOTAL'],
                    'Cantidad': [df_status['Cantidad'].sum()]
                })
                df_status = pd.concat([df_status, total_row], ignore_index=True)
                df_status.to_excel(writer, sheet_name='Resumen por Status', index=False)
            
            if not df_export.empty and 'Promotor' in df_export.columns and 'Status' in df_export.columns:
                df_promotor = df_export.groupby(['Promotor', 'Status']).size().unstack(fill_value=0)
                df_promotor['Total'] = df_promotor.sum(axis=1)
                df_promotor = df_promotor.sort_values('Total', ascending=False)
                df_promotor.to_excel(writer, sheet_name='Resumen por Promotor')
            
            info_data = {
                'Supervisor': [supervisor_name],
                'ID Supervisor': [supervisor_id],
                'Mes': [mes],
                'Fecha Exportación': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                'Total Revisiones': [len(reviews)]
            }
            df_info = pd.DataFrame(info_data)
            df_info.to_excel(writer, sheet_name='Info Reporte', index=False)
            
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        
        filename = f"reporte_revisiones_{supervisor_name}_{mes}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return send_file(
            output,
            download_name=filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Error al exportar: {e}")
        return jsonify({"error": str(e)}), 500


# =====================================================
# CIERRE DE MES
# =====================================================

@app.route("/api/close_month", methods=["POST"])
@login_required
def api_close_month():
    """Exportar todas las revisiones con todos los campos y limpiar la base de datos"""
    global _cached_df, _data_loaded
    
    try:
        supervisor_id = session.get('supervisor_id')
        supervisor_name = session.get('supervisor_name')
        mes = _current_mes or get_mes_actual()
        
        reviews = db.get_all_reviews(supervisor_id, mes)
        
        if not reviews:
            return jsonify({"error": "No hay revisiones para exportar"}), 404
        
        if _cached_df is None or not _data_loaded:
            export_data = []
            for review in reviews:
                export_data.append({
                    "Fecha Ejecución": "",
                    "POC ID": "",
                    "Cliente ID": "",
                    "Razón Social": "",
                    "Direccion": "",
                    "Detalle Tarea": "",
                    "Imagen": "",
                    "Promotor": "",
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        else:
            task_data = {}
            
            for _, row in _cached_df.iterrows():
                raw_poc_id = clean_text(row.get("POC ID", ""))
                short_poc_id = extract_short_poc_id(raw_poc_id)
                client_info = get_client_info(raw_poc_id) if raw_poc_id else None
                img_url = clean_text(row.get("Img", ""))
                if not img_url:
                    img_url = clean_text(row.get("TaskImageUrl", ""))
                
                task_data[row["task_id"]] = {
                    "fecha": formatear_fecha(row.get("Fecha")),
                    "promotor": row.get("Promotor", ""),
                    "poc_id": short_poc_id,
                    "poc_id_completo": raw_poc_id,
                    "cliente_id": short_poc_id,
                    "razon_social": client_info.get("nombre") if client_info else "",
                    "direccion": client_info.get("direccion") if client_info else "",
                    "detalle_tarea": row.get("Detalle Tarea", ""),
                    "imagen": img_url
                }
            
            export_data = []
            for review in reviews:
                task_id = review["task_id"]
                task_info = task_data.get(task_id, {})
                
                export_data.append({
                    "Fecha Ejecución": task_info.get("fecha", ""),
                    "POC ID": task_info.get("poc_id", ""),
                    "POC ID Completo": task_info.get("poc_id_completo", ""),
                    "Cliente ID": task_info.get("cliente_id", ""),
                    "Razón Social": task_info.get("razon_social", ""),
                    "Direccion": task_info.get("direccion", ""),
                    "Detalle Tarea": task_info.get("detalle_tarea", ""),
                    "Imagen": task_info.get("imagen", ""),
                    "Promotor": task_info.get("promotor", ""),
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, sheet_name='Revisiones Cierre', index=False)
            
            if not df_export.empty and 'Status' in df_export.columns:
                df_status = df_export.groupby('Status').size().reset_index(name='Cantidad')
                df_status = df_status.sort_values('Cantidad', ascending=False)
                total_row = pd.DataFrame({
                    'Status': ['TOTAL'],
                    'Cantidad': [df_status['Cantidad'].sum()]
                })
                df_status = pd.concat([df_status, total_row], ignore_index=True)
                df_status.to_excel(writer, sheet_name='Resumen por Status', index=False)
            
            if not df_export.empty and 'Promotor' in df_export.columns and 'Status' in df_export.columns:
                df_promotor = df_export.groupby(['Promotor', 'Status']).size().unstack(fill_value=0)
                df_promotor['Total'] = df_promotor.sum(axis=1)
                df_promotor = df_promotor.sort_values('Total', ascending=False)
                df_promotor.to_excel(writer, sheet_name='Resumen por Promotor')
            
            info_data = {
                'Supervisor': [supervisor_name],
                'ID Supervisor': [supervisor_id],
                'Mes Cerrado': [mes],
                'Fecha Cierre': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                'Total Revisiones': [len(reviews)],
                'Nota': ['Este archivo contiene el cierre completo del mes']
            }
            df_info = pd.DataFrame(info_data)
            df_info.to_excel(writer, sheet_name='Info Cierre', index=False)
            
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        
        filename = f"cierre_mes_{supervisor_name}_{mes}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        db.clear_reviews(supervisor_id, mes)
        
        file_path = UPLOAD_DIR / "data.xlsx"
        if file_path.exists():
            file_path.unlink()
        
        _cached_df = None
        _data_loaded = False
        
        response = send_file(
            output,
            download_name=filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        return response
        
    except Exception as e:
        print(f"Error en cierre de mes: {e}")
        return jsonify({"error": str(e)}), 500


# =====================================================
# START
# =====================================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)