# app.py - VERSIÓN COMPLETA CORREGIDA (SOLO GUARDADO EN SHEETS)
from __future__ import annotations
from pathlib import Path
from typing import Any
import pandas as pd
from flask import Flask, jsonify, render_template, request, session, send_file
from functools import wraps
import hashlib
from datetime import datetime
import io
import json
import re
import os
from tempfile import NamedTemporaryFile

# =====================================================
# IMPORTS PARA GOOGLE SHEETS
# =====================================================
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GOOGLE_SHEETS_AVAILABLE = True
    print("✅ gspread importado correctamente")
except ImportError as e:
    GOOGLE_SHEETS_AVAILABLE = False
    print(f"⚠️ gspread no instalado: {e}")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

SUPERVISORS = {
    14: "Bruno Del Popolo",
    17: "Franco Vivani",
    41: "Claudio Raposo",
}

app = Flask(__name__)
app.secret_key = "clave_secreta_para_desarrollo"

_cached_df = None
_current_excel = None
_data_loaded = False
_current_mes = None

# =====================================================
# CONFIGURACIÓN DE GOOGLE SHEETS CON VARIABLES DE ENTORNO
# =====================================================

# Leer credenciales desde variable de entorno (Render)
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS")
SHEET_ID = os.environ.get("SHEET_ID", "12D-Ru8GNm0EE0NFg4kpcygqcPdqKpor9U4NOA-1TQRs")
SHEET_NAME = os.environ.get("SHEET_NAME", "Clientes")

# Crear archivo temporal con las credenciales
CREDENTIALS_FILE = "credentials.json"
if GOOGLE_CREDENTIALS_JSON:
    try:
        creds_data = json.loads(GOOGLE_CREDENTIALS_JSON)
        temp_creds = NamedTemporaryFile(mode='w', suffix='.json', delete=False)
        json.dump(creds_data, temp_creds)
        temp_creds.close()
        CREDENTIALS_FILE = temp_creds.name
        print("✅ Credenciales cargadas desde variable de entorno")
    except Exception as e:
        print(f"❌ Error al procesar credenciales: {e}")
        CREDENTIALS_FILE = "credentials.json"
else:
    print("⚠️ No se encontró GOOGLE_CREDENTIALS en variables de entorno")
    if os.path.exists("credentials.json"):
        CREDENTIALS_FILE = "credentials.json"
        print("✅ Usando credentials.json local")

GOOGLE_SHEETS_CONFIG = {
    "sheet_id": SHEET_ID,
    "sheet_name": SHEET_NAME,
    "credentials_file": CREDENTIALS_FILE
}

# =====================================================
# CARGAR CLIENTES DESDE JSON (FUNCIONA SIEMPRE)
# =====================================================

def cargar_clientes_json():
    """Cargar clientes desde clientes.json"""
    try:
        json_path = BASE_DIR / "clientes.json"
        if json_path.exists():
            with open(json_path, 'r', encoding='utf-8') as f:
                clientes = json.load(f)
            print(f"✅ Clientes cargados desde JSON: {len(clientes)}")
            return clientes
    except Exception as e:
        print(f"❌ Error cargando JSON: {e}")
    return {}

CLIENTES = cargar_clientes_json()

# =====================================================
# FUNCIONES PARA GOOGLE SHEETS - HOJA STATUS
# =====================================================

def get_google_sheet_client():
    """Obtener conexión a Google Sheets"""
    if not GOOGLE_SHEETS_AVAILABLE:
        print("❌ gspread no disponible")
        return None
    
    try:
        creds_file = GOOGLE_SHEETS_CONFIG["credentials_file"]
        print(f"🔍 Intentando conectar con credenciales: {creds_file}")
        
        if not os.path.exists(creds_file):
            print(f"❌ Archivo de credenciales no existe: {creds_file}")
            return None
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        client = gspread.authorize(creds)
        print("✅ Conexión a Google Sheets exitosa")
        return client
    except Exception as e:
        print(f"❌ Error al conectar con Google Sheets: {e}")
        return None

def guardar_status_en_sheets(fecha, id_imagen, status, supervisor, observaciones=""):
    """
    Guardar el status de una foto en la hoja 'Status' de Google Sheets
    Columnas: Fecha, ID (URL de la imagen), Status, Supervisor, Observaciones, Fecha Revision
    """
    try:
        print(f"🔍 Intentando guardar en Sheets:")
        print(f"   - Fecha: {fecha}")
        print(f"   - ID Imagen: {id_imagen}")
        print(f"   - Status: {status}")
        print(f"   - Supervisor: {supervisor}")
        
        gc = get_google_sheet_client()
        if gc is None:
            print("❌ No se pudo conectar a Google Sheets")
            return False
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        
        try:
            worksheet = sheet.worksheet("Status")
            print("✅ Hoja Status encontrada")
        except gspread.WorksheetNotFound:
            print("🔄 Hoja Status no existe, creando...")
            worksheet = sheet.add_worksheet(title="Status", rows=1000, cols=10)
            worksheet.append_row(["Fecha", "ID", "Status", "Supervisor", "Observaciones", "Fecha Revision"])
            print("✅ Hoja Status creada")
        
        id_column = worksheet.col_values(2)
        row_to_update = None
        if id_imagen in id_column:
            row_to_update = id_column.index(id_imagen) + 1
            print(f"📊 ID encontrado en fila {row_to_update} - Se actualizará")
        else:
            print(f"📊 ID no encontrado - Se agregará nuevo")
        
        fecha_revision = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        if row_to_update:
            worksheet.update(f'C{row_to_update}:F{row_to_update}', [[status, supervisor, observaciones, fecha_revision]])
            print(f"✅ Status actualizado en Sheets: {id_imagen} -> {status}")
        else:
            worksheet.append_row([fecha, id_imagen, status, supervisor, observaciones, fecha_revision])
            print(f"✅ Status guardado en Sheets: {id_imagen} -> {status}")
        
        return True
    except Exception as e:
        print(f"❌ Error al guardar en Sheets: {e}")
        import traceback
        traceback.print_exc()
        return False

def get_status_from_sheets(id_imagen):
    """
    Obtener el status de una foto desde la hoja 'Status'
    """
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return None
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        
        try:
            worksheet = sheet.worksheet("Status")
        except gspread.WorksheetNotFound:
            return None
        
        id_column = worksheet.col_values(2)
        if id_imagen in id_column:
            row_index = id_column.index(id_imagen) + 1
            row_data = worksheet.row_values(row_index)
            return {
                "status": row_data[2] if len(row_data) > 2 else None,
                "supervisor": row_data[3] if len(row_data) > 3 else None,
                "observaciones": row_data[4] if len(row_data) > 4 else None,
                "fecha_revision": row_data[5] if len(row_data) > 5 else None
            }
        return None
    except Exception as e:
        print(f"❌ Error al leer status desde Sheets: {e}")
        return None

def get_stats_from_sheets(supervisor_name):
    """
    Obtener estadísticas de la hoja 'Status' para un supervisor específico
    """
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return None
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        
        try:
            worksheet = sheet.worksheet("Status")
        except gspread.WorksheetNotFound:
            return None
        
        records = worksheet.get_all_records()
        
        # Filtrar por supervisor
        stats = {
            "total": 0,
            "by_status": {
                "objeccion": 0,
                "invalida": 0,
                "fraude": 0
            }
        }
        
        for row in records:
            supervisor = row.get("Supervisor", "")
            if supervisor == supervisor_name:
                status = row.get("Status", "")
                stats["total"] += 1
                if status in stats["by_status"]:
                    stats["by_status"][status] += 1
        
        return stats
    except Exception as e:
        print(f"❌ Error al leer stats desde Sheets: {e}")
        return None

# =====================================================
# FUNCIONES
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

def extract_short_poc_id(poc_id):
    """Extraer código corto del POC ID"""
    if not poc_id:
        return ""
    
    poc_id = str(poc_id).strip()
    if poc_id.endswith('.0'):
        poc_id = poc_id[:-2]
    poc_id = poc_id.replace('-', '').replace(' ', '')
    
    if poc_id.startswith("053822"):
        resultado = poc_id[6:].lstrip('0')
        return resultado if resultado else "0"
    if poc_id.startswith("53822"):
        resultado = poc_id[5:].lstrip('0')
        return resultado if resultado else "0"
    
    return poc_id.lstrip('0') or "0"

def get_client_info(poc_id):
    """Obtener datos del cliente desde JSON"""
    if not poc_id:
        return None
    
    short_id = extract_short_poc_id(poc_id)
    if not short_id or short_id == "0":
        return None
    
    if short_id in CLIENTES:
        return CLIENTES[short_id]
    
    return None

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'supervisor_id' not in session:
            return jsonify({"error": "No autorizado"}), 401
        return f(*args, **kwargs)
    return decorated_function

def _build_task_id(row: pd.Series) -> str:
    """Crear ID único para cada tarea usando POC ID + Fecha + Detalle Tarea"""
    poc_id = clean_text(row.get("POC ID", ""))
    fecha = str(row.get("Fecha", ""))
    detalle = clean_text(row.get("Detalle Tarea", ""))
    raw = f"{poc_id}|{fecha}|{detalle}"
    return hashlib.md5(raw.encode()).hexdigest()[:12]

def get_mes_actual():
    return datetime.now().strftime("%Y%m")

def formatear_fecha(fecha_valor):
    if pd.isna(fecha_valor):
        return ""
    
    if isinstance(fecha_valor, (datetime, pd.Timestamp)):
        return fecha_valor.strftime("%d/%m/%Y")
    
    fecha_str = str(fecha_valor).strip()
    
    if '/' in fecha_str:
        return fecha_str
    
    if '-' in fecha_str:
        partes = fecha_str.split('-')
        if len(partes) == 3:
            año = partes[0]
            mes = partes[1]
            dia = partes[2].split(' ')[0]
            return f"{dia}/{mes}/{año}"
    
    if len(fecha_str) >= 8 and fecha_str[:8].isdigit():
        año = fecha_str[0:4]
        mes = fecha_str[4:6]
        dia = fecha_str[6:8]
        return f"{dia}/{mes}/{año}"
    
    return fecha_str

def _load_data(path: Path) -> pd.DataFrame:
    global _current_excel, _data_loaded, _current_mes
    
    df = pd.read_excel(path, engine="openpyxl")
    _current_excel = path.name
    _current_mes = get_mes_actual()
    
    if "TaskImageUrl" in df.columns and "Img" not in df.columns:
        df = df.rename(columns={"TaskImageUrl": "Img"})
    
    required = ["Fecha", "Promotor", "POC ID", "Detalle Tarea", "Img", "Completada", "Validada", "Visita Valida", "Supervisor ID"]
    
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Columnas faltantes: {missing}")
    
    df = df.copy()
    
    df["Completada"] = pd.to_numeric(df["Completada"], errors="coerce")
    df["Validada"] = pd.to_numeric(df["Validada"], errors="coerce")
    df["Supervisor ID"] = pd.to_numeric(df["Supervisor ID"], errors="coerce").astype("Int64")
    df["VisitaValidaBool"] = df["Visita Valida"].apply(_is_visita_valida)
    
    filtered = df[
        (df["Completada"] == 1.0) &
        (df["Validada"] == 0.0) &
        (df["VisitaValidaBool"])
    ].copy()
    
    if 'supervisor_id' in session:
        supervisor_id = session['supervisor_id']
        filtered = filtered[filtered["Supervisor ID"] == supervisor_id]
    
    filtered = filtered.reset_index(drop=True)
    filtered["row_id"] = range(1, len(filtered) + 1)
    filtered["task_id"] = filtered.apply(_build_task_id, axis=1)
    
    print(f"📊 Task IDs generados: {filtered['task_id'].head().tolist()}")
    print(f"📊 Total tareas: {len(filtered)}")
    
    _data_loaded = True
    
    return filtered

# =====================================================
# ROUTES
# =====================================================

@app.route("/")
def index():
    return render_template("index.html", supervisors=SUPERVISORS)

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.json
    supervisor_id = data.get("supervisor_id")
    
    if supervisor_id not in SUPERVISORS:
        return jsonify({"error": "Supervisor no encontrado"}), 404
    
    session['supervisor_id'] = supervisor_id
    session['supervisor_name'] = SUPERVISORS[supervisor_id]
    
    # =====================================================
    # ELIMINADO: No cargar archivo automáticamente
    # Ahora el supervisor debe subir un archivo cada vez
    # =====================================================
    # global _cached_df, _data_loaded
    # file_path = UPLOAD_DIR / "data.xlsx"
    # if file_path.exists():
    #     try:
    #         _cached_df = _load_data(file_path)
    #         _data_loaded = True
    #     except Exception as e:
    #         print(f"Error recargando datos: {e}")
    # =====================================================
    
    return jsonify({
        "ok": True,
        "supervisor_id": supervisor_id,
        "supervisor_name": SUPERVISORS[supervisor_id]
    })

@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"ok": True})

@app.route("/api/has_file")
@login_required
def api_has_file():
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
    global _cached_df, _data_loaded, _current_mes
    
    if "file" not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400
    
    f = request.files["file"]
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "El archivo debe ser .xlsx"}), 400
    
    file_path = UPLOAD_DIR / "data.xlsx"
    f.save(str(file_path))
    
    try:
        print(f"📂 Procesando archivo subido: {file_path}")
        _cached_df = _load_data(file_path)
        _data_loaded = True
        _current_mes = get_mes_actual()
        
        print(f"✅ Archivo procesado: {len(_cached_df)} filas")
        
        return jsonify({
            "ok": True,
            "rows": len(_cached_df),
            "message": f"Archivo cargado con {len(_cached_df)} registros",
            "mes": _current_mes
        })
    except Exception as e:
        print(f"❌ Error al procesar archivo: {e}")
        import traceback
        traceback.print_exc()
        _cached_df = None
        _data_loaded = False
        return jsonify({"error": str(e)}), 500

@app.route("/api/tasks")
@login_required
def api_tasks():
    global _cached_df, _data_loaded
    
    try:
        print("=" * 50)
        print("🔍 INICIO api_tasks")
        print(f"📊 _data_loaded: {_data_loaded}")
        print(f"📊 _cached_df is None: {_cached_df is None}")
        
        # =====================================================
        # SI NO HAY DATOS CARGADOS, DEVOLVER ERROR
        # EL SUPERVISOR DEBE SUBIR UN ARCHIVO PRIMERO
        # =====================================================
        if not _data_loaded or _cached_df is None or _cached_df.empty:
            return jsonify({
                "error": "No hay datos cargados. Debes subir un archivo Excel primero.",
                "no_data": True
            }), 404
        # =====================================================
        
        supervisor_id = session.get('supervisor_id')
        supervisor_name = SUPERVISORS.get(supervisor_id, "Desconocido")
        print(f"👤 Supervisor ID: {supervisor_id}, Nombre: {supervisor_name}")
        
        start_date = request.args.get("start_date", type=str)
        end_date = request.args.get("end_date", type=str)
        print(f"📅 Filtros: desde={start_date}, hasta={end_date}")
        
        print(f"📊 Filtrando por Supervisor ID = {supervisor_id}")
        result = _cached_df[_cached_df["Supervisor ID"] == supervisor_id].copy()
        print(f"📊 Tareas del supervisor: {len(result)}")
        
        if result.empty:
            return jsonify({"error": f"No hay tareas asignadas para {supervisor_name}", "no_tasks": True}), 404
        
        result = result.sort_values(by="Fecha", ascending=True)
        
        if start_date or end_date:
            try:
                result['fecha_dt'] = pd.to_datetime(result['Fecha'], format='%d/%m/%Y', errors='coerce')
                if start_date:
                    start_dt = datetime.strptime(start_date, "%d/%m/%Y")
                    result = result[result['fecha_dt'] >= start_dt]
                if end_date:
                    end_dt = datetime.strptime(end_date, "%d/%m/%Y")
                    result = result[result['fecha_dt'] <= end_dt]
                result = result.drop(columns=['fecha_dt'])
                print(f"📊 Después de filtros de fecha: {len(result)}")
            except Exception as e:
                print(f"⚠️ Error en filtros de fecha: {e}")
                if start_date:
                    result = result[result["Fecha"].astype(str) >= start_date]
                if end_date:
                    result = result[result["Fecha"].astype(str) <= end_date]
        
        if result.empty:
            return jsonify({"error": "No hay tareas en el rango de fechas", "no_tasks": True}), 404
        
        print("🔄 Construyendo respuesta...")
        response_rows = []
        
        for _, row in result.iterrows():
            try:
                task_id = row["task_id"]
                
                img_url = clean_text(row.get("Img", ""))
                if not img_url:
                    img_url = clean_text(row.get("TaskImageUrl", ""))
                
                # Leer status desde Sheets
                try:
                    status_sheets = get_status_from_sheets(img_url)
                except Exception as e:
                    print(f"⚠️ Error al leer status desde Sheets: {e}")
                    status_sheets = None
                
                if status_sheets:
                    is_reviewed = True
                    review = {
                        "status": status_sheets["status"],
                        "observaciones": status_sheets.get("observaciones", ""),
                        "fecha_revision": status_sheets.get("fecha_revision", "")
                    }
                else:
                    is_reviewed = False
                    review = None
                
                raw_poc_id = clean_text(row.get("POC ID"))
                short_poc_id = extract_short_poc_id(raw_poc_id)
                client_info = get_client_info(raw_poc_id)
                
                response_rows.append({
                    "row_id": int(row["row_id"]),
                    "task_id": task_id,
                    "fecha": formatear_fecha(row.get("Fecha")),
                    "promotor": clean_text(row.get("Promotor")),
                    "poc_id": short_poc_id,
                    "poc_id_completo": raw_poc_id,
                    "cliente_id": short_poc_id,
                    "razon_social": client_info.get("nombre") if client_info else "SIN DATO",
                    "direccion": client_info.get("direccion") if client_info else "SIN DATO",
                    "detalle_tarea": clean_text(row.get("Detalle Tarea")),
                    "imagen": img_url,
                    "revisado": is_reviewed,
                    "status": review.get("status") if review else None,
                    "observaciones": review.get("observaciones") if review else "",
                    "fecha_revision": review.get("fecha_revision") if review else None
                })
                    
            except Exception as e:
                print(f"❌ Error procesando fila: {e}")
                continue
        
        print(f"✅ Respondiendo con {len(response_rows)} tareas")
        print("=" * 50)
        return jsonify(response_rows)
        
    except Exception as e:
        print(f"❌❌❌ ERROR GRAVE en api_tasks: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e), "no_data": True}), 500

@app.route("/api/save_review", methods=["POST"])
@login_required
def api_save_review():
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
    
    row = task_row.iloc[0]
    supervisor_name = session['supervisor_name']
    
    fecha_tarea = formatear_fecha(row.get("Fecha"))
    img_url = clean_text(row.get("Img", ""))
    if not img_url:
        img_url = clean_text(row.get("TaskImageUrl", ""))
    
    success = guardar_status_en_sheets(
        fecha=fecha_tarea,
        id_imagen=img_url,
        status=status,
        supervisor=supervisor_name,
        observaciones=observaciones
    )
    
    if success:
        return jsonify({"ok": True, "message": "Revisión guardada correctamente"})
    else:
        return jsonify({"error": "Error al guardar la revisión"}), 500

@app.route("/api/delete_review/<task_id>", methods=["DELETE"])
@login_required
def api_delete_review(task_id):
    try:
        supervisor_id = session.get('supervisor_id')
        
        if _cached_df is None or not _data_loaded:
            return jsonify({"error": "No hay datos cargados"}), 404
        
        task_row = _cached_df[_cached_df["task_id"] == task_id]
        if task_row.empty:
            return jsonify({"error": "Tarea no encontrada"}), 404
        
        try:
            img_url = clean_text(task_row.iloc[0].get("Img", ""))
            if not img_url:
                img_url = clean_text(task_row.iloc[0].get("TaskImageUrl", ""))
            
            gc = get_google_sheet_client()
            if gc:
                sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
                try:
                    worksheet = sheet.worksheet("Status")
                    id_column = worksheet.col_values(2)
                    if img_url in id_column:
                        row_index = id_column.index(img_url) + 1
                        worksheet.delete_rows(row_index)
                        print(f"✅ Revisión eliminada de Sheets: {img_url}")
                except gspread.WorksheetNotFound:
                    pass
        except Exception as e:
            print(f"⚠️ Error al eliminar de Sheets: {e}")
        
        return jsonify({"ok": True, "message": "Revisión eliminada correctamente"})
            
    except Exception as e:
        print(f"Error al eliminar revisión: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/stats")
@login_required
def api_stats():
    global _data_loaded
    
    try:
        supervisor_name = session.get('supervisor_name')
        
        # Si no hay datos cargados
        if not _data_loaded or _cached_df is None:
            return jsonify({
                "total_revisados": 0,
                "total_pendientes": 0,
                "by_status": {"objeccion": 0, "invalida": 0, "fraude": 0},
                "porcentaje": 0,
                "no_data": True
            })
        
        # Total de tareas disponibles para el supervisor
        supervisor_id = session.get('supervisor_id')
        total_disponibles = len(_cached_df[_cached_df["Supervisor ID"] == supervisor_id])
        
        # Obtener estadísticas desde Google Sheets
        stats = get_stats_from_sheets(supervisor_name)
        
        if stats:
            total_revisados = stats["total"]
            total_pendientes = total_disponibles - total_revisados
            porcentaje = round((total_revisados / total_disponibles * 100) if total_disponibles > 0 else 0, 1)
            
            return jsonify({
                "total_revisados": total_revisados,
                "total_pendientes": total_pendientes,
                "by_status": stats["by_status"],
                "porcentaje": porcentaje,
                "total_disponibles": total_disponibles,
                "has_data": True,
                "mes": _current_mes
            })
        else:
            # Si no se pudo obtener de Sheets
            return jsonify({
                "total_revisados": 0,
                "total_pendientes": total_disponibles,
                "by_status": {"objeccion": 0, "invalida": 0, "fraude": 0},
                "porcentaje": 0,
                "has_data": True,
                "mes": _current_mes,
                "mensaje": "No se pudieron obtener estadísticas de Sheets"
            })
            
    except Exception as e:
        print(f"❌ Error en api_stats: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({
            "total_revisados": 0,
            "total_pendientes": 0,
            "by_status": {"objeccion": 0, "invalida": 0, "fraude": 0},
            "porcentaje": 0,
            "no_data": True,
            "error": str(e)
        })

@app.route("/api/supervisors")
def api_supervisors():
    items = [{"id": sid, "name": name} for sid, name in SUPERVISORS.items()]
    return jsonify(items)

# =====================================================
# ENDPOINTS DE PRUEBA
# =====================================================

@app.route("/api/test_tasks_direct")
@login_required
def test_tasks_direct():
    """Prueba directa de api_tasks sin try/except"""
    global _cached_df, _data_loaded
    
    print("🔍 TEST DIRECT: Iniciando...")
    print(f"📊 _data_loaded: {_data_loaded}")
    print(f"📊 _cached_df is None: {_cached_df is None}")
    
    if _cached_df is not None:
        print(f"📊 _cached_df columns: {_cached_df.columns.tolist()}")
        print(f"📊 _cached_df rows: {len(_cached_df)}")
    
    # Verificar datos
    if not _data_loaded or _cached_df is None or _cached_df.empty:
        return jsonify({"error": "No hay datos cargados", "status": "no_data"})
    
    supervisor_id = session.get('supervisor_id')
    print(f"👤 Supervisor ID: {supervisor_id}")
    
    # Filtrar por supervisor
    result = _cached_df[_cached_df["Supervisor ID"] == supervisor_id].copy()
    print(f"📊 Tareas del supervisor: {len(result)}")
    
    if result.empty:
        return jsonify({"error": "No hay tareas para este supervisor", "status": "empty"})
    
    # Devolver datos básicos SIN acceder a Sheets
    tareas = []
    for _, row in result.iterrows():
        tareas.append({
            "row_id": int(row["row_id"]),
            "task_id": row["task_id"],
            "promotor": clean_text(row.get("Promotor")),
            "fecha": formatear_fecha(row.get("Fecha")),
            "poc_id": extract_short_poc_id(clean_text(row.get("POC ID"))),
            "detalle": clean_text(row.get("Detalle Tarea", "")),
            "imagen": clean_text(row.get("Img", ""))
        })
    
    return jsonify({
        "total": len(tareas),
        "tareas": tareas[:10]
    })

@app.route("/api/test_guardar")
@login_required
def test_guardar():
    """Probar guardado directo en Sheets"""
    try:
        fecha = "22/08/2026"
        id_imagen = "https://test.com/imagen_test.jpg"
        status = "objeccion"
        supervisor = "Bruno Del Popolo"
        observaciones = "Prueba desde el endpoint"
        
        print("🔍 Probando guardado directo...")
        resultado = guardar_status_en_sheets(fecha, id_imagen, status, supervisor, observaciones)
        
        if resultado:
            return jsonify({"ok": True, "mensaje": "Guardado exitoso en Sheets"})
        else:
            return jsonify({"ok": False, "mensaje": "Fallo el guardado en Sheets"}), 500
    except Exception as e:
        print(f"❌ Error en test_guardar: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/test_credenciales")
@login_required
def test_credenciales():
    """Verificar si existe el archivo de credenciales"""
    creds_file = GOOGLE_SHEETS_CONFIG["credentials_file"]
    existe = os.path.exists(creds_file)
    
    return jsonify({
        "archivo_credenciales": creds_file,
        "existe": existe,
        "directorio_actual": os.getcwd(),
        "archivos_en_directorio": os.listdir(".") if existe else []
    })

@app.route("/api/test_sheets_status")
@login_required
def test_sheets_status():
    """Probar conexión a la hoja Status"""
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return jsonify({"error": "No se pudo conectar a Google Sheets"}), 500
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        
        try:
            worksheet = sheet.worksheet("Status")
            records = worksheet.get_all_records()
            return jsonify({
                "conexion": "OK",
                "total_registros": len(records),
                "primeros_5": records[:5] if records else []
            })
        except gspread.WorksheetNotFound:
            return jsonify({
                "conexion": "OK",
                "total_registros": 0,
                "mensaje": "La hoja 'Status' no existe aún. Se creará al guardar la primera revisión."
            })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/debug_tasks")
@login_required
def debug_tasks():
    """Debug: ver qué está pasando con las tareas"""
    try:
        global _cached_df, _data_loaded
        
        resultado = {
            "data_loaded": _data_loaded,
            "cached_df_is_none": _cached_df is None,
            "cached_df_empty": _cached_df.empty if _cached_df is not None else True,
            "supervisor_id": session.get('supervisor_id'),
            "file_exists": (UPLOAD_DIR / "data.xlsx").exists(),
            "columnas": _cached_df.columns.tolist() if _cached_df is not None else None,
            "total_filas": len(_cached_df) if _cached_df is not None else 0
        }
        
        if _cached_df is not None:
            supervisor_id = session.get('supervisor_id')
            supervisor_tareas = _cached_df[_cached_df["Supervisor ID"] == supervisor_id]
            resultado["tareas_supervisor"] = len(supervisor_tareas)
            resultado["primeras_filas"] = supervisor_tareas.head(2).to_dict('records') if not supervisor_tareas.empty else []
        
        return jsonify(resultado)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/tasks_simple")
@login_required
def tasks_simple():
    """Versión simple de tasks para debug"""
    try:
        global _cached_df, _data_loaded
        
        if not _data_loaded or _cached_df is None or _cached_df.empty:
            return jsonify({"error": "No hay datos cargados", "status": "no_data"})
        
        supervisor_id = session.get('supervisor_id')
        result = _cached_df[_cached_df["Supervisor ID"] == supervisor_id].copy()
        
        if result.empty:
            return jsonify({"error": "No hay tareas", "status": "empty"})
        
        tareas = []
        for _, row in result.iterrows():
            tareas.append({
                "row_id": int(row["row_id"]),
                "task_id": row["task_id"],
                "promotor": clean_text(row.get("Promotor")),
                "fecha": formatear_fecha(row.get("Fecha")),
                "imagen": clean_text(row.get("Img", ""))
            })
        
        return jsonify({"total": len(tareas), "tareas": tareas[:5]})
        
    except Exception as e:
        import traceback
        error = traceback.format_exc()
        print(f"❌ Error: {error}")
        return jsonify({"error": str(e), "trace": error}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)