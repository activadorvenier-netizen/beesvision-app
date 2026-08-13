# app.py - VERSIÓN FINAL CORREGIDA
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
    print("⚠️ gspread no instalado")

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Base de datos
db = ReviewDatabase()

# Supervisores
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

GOOGLE_SHEETS_CONFIG = {
    "sheet_id": "12D-Ru8GNm0EE0NFg4kpcygqcPdqKpor9U4NOA-1TQRs",
    "sheet_name": "Clientes",
    "credentials_file": "credentials.json"
}

_client_cache = {}
_client_cache_time = 0
CACHE_TTL = 600

# =====================================================
# CLIENTES LOCALES (RESPALDO)
# =====================================================

CLIENTES_LOCALES = {
    "1875": {
        "nombre": "GRAU ADRIAN",
        "direccion": "CHABAS - ESPAÑA - 2151"
    },
    "502": {
        "nombre": "CLIENTE 502",
        "direccion": "DIRECCION 502"
    }
}

# =====================================================
# FUNCIONES
# =====================================================

def get_google_sheet_client():
    """Obtener conexión a Google Sheets"""
    if not GOOGLE_SHEETS_AVAILABLE:
        return None
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            GOOGLE_SHEETS_CONFIG["credentials_file"], scope
        )
        return gspread.authorize(creds)
    except Exception as e:
        print(f"❌ Error al conectar con Google Sheets: {e}")
        return None

def load_client_master():
    """Cargar el maestro de clientes desde Google Sheets con caché"""
    global _client_cache, _client_cache_time
    
    if not GOOGLE_SHEETS_AVAILABLE:
        return {}
    
    now = time.time()
    if _client_cache and (now - _client_cache_time) < CACHE_TTL:
        return _client_cache
    
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return {}
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        worksheet = sheet.worksheet(GOOGLE_SHEETS_CONFIG["sheet_name"])
        records = worksheet.get_all_records()
        
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
        return {}

def extract_short_poc_id(poc_id):
    """
    Extraer el código de cliente del POC ID completo.
    Maneja ambos formatos:
    - Con 0 inicial: 05382200001875 -> 1875
    - Sin 0 inicial: 5382200001875 -> 1875
    """
    if not poc_id:
        return ""
    
    poc_id = str(poc_id).strip()
    
    if poc_id.endswith('.0'):
        poc_id = poc_id[:-2]
    
    poc_id = poc_id.replace('-', '').replace(' ', '')
    
    # Prefijos posibles
    PREFIX_CON_CERO = "053822"    # 6 dígitos con cero
    PREFIX_SIN_CERO = "53822"     # 5 dígitos sin cero
    
    resultado = poc_id
    
    # Si empieza con el prefijo CON cero (053822)
    if poc_id.startswith(PREFIX_CON_CERO):
        resultado = poc_id[len(PREFIX_CON_CERO):]
    # Si empieza con el prefijo SIN cero (53822)
    elif poc_id.startswith(PREFIX_SIN_CERO):
        resultado = poc_id[len(PREFIX_SIN_CERO):]
    
    # Eliminar ceros a la izquierda
    resultado = resultado.lstrip('0')
    
    return resultado if resultado else "0"

def get_client_info(poc_id):
    """
    Obtener información de un cliente por POC ID.
    1. Extrae el código corto del POC ID (ej: 05382200001875 -> 1875)
    2. Busca ese código en Google Sheets
    3. Si no encuentra, busca en CLIENTES_LOCALES
    """
    if not poc_id:
        return None
    
    # Extraer el código corto del cliente
    short_id = extract_short_poc_id(poc_id)
    if not short_id or short_id == "0":
        return None
    
    print(f"🔍 Buscando cliente: {short_id} (POC ID: {poc_id})")
    
    # BUSCAR EN GOOGLE SHEETS
    master = load_client_master()
    if master and short_id in master:
        print(f"✅ Cliente encontrado en SHEETS: {short_id} -> {master[short_id]['nombre']}")
        return master[short_id]
    
    # BUSCAR EN CLIENTES LOCALES (RESPALDO)
    if short_id in CLIENTES_LOCALES:
        print(f"✅ Cliente encontrado en LOCAL: {short_id} -> {CLIENTES_LOCALES[short_id]['nombre']}")
        return CLIENTES_LOCALES[short_id]
    
    print(f"⚠️ Cliente NO ENCONTRADO: {short_id}")
    return None

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
    img_value = clean_text(row.get("Img", ""))
    poc_id = clean_text(row.get("POC ID", ""))
    fecha = str(row.get("Fecha", ""))
    parts = [img_value, poc_id, fecha]
    raw = "|".join(parts)
    return hashlib.md5(raw.encode()).hexdigest()[:12]

def get_mes_actual():
    return datetime.now().strftime("%Y%m")

def formatear_fecha(fecha_valor):
    """Formatear fecha a DD/MM/YYYY"""
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

# =====================================================
# LOAD DATA
# =====================================================

def _load_data(path: Path) -> pd.DataFrame:
    global _current_excel, _data_loaded, _current_mes
    
    df = pd.read_excel(path, engine="openpyxl")
    _current_excel = path.name
    _current_mes = get_mes_actual()
    
    # Renombrar TaskImageUrl a Img si existe
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
        _cached_df = _load_data(file_path)
        _data_loaded = True
        _current_mes = get_mes_actual()
        
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
                return jsonify({"error": f"Error cargando datos: {str(e)}", "no_data": True}), 404
        else:
            return jsonify({"error": "No hay datos cargados", "no_data": True}), 404
    
    supervisor_id = session['supervisor_id']
    start_date = request.args.get("start_date", type=str)
    end_date = request.args.get("end_date", type=str)
    
    result = _cached_df[_cached_df["Supervisor ID"] == supervisor_id].copy()
    
    if result.empty:
        return jsonify({"error": f"No hay tareas asignadas para {SUPERVISORS[supervisor_id]}", "no_tasks": True}), 404
    
    if start_date:
        try:
            result = result[result["Fecha"].astype(str) >= start_date]
        except:
            pass
    if end_date:
        try:
            result = result[result["Fecha"].astype(str) <= end_date]
        except:
            pass
    
    if result.empty:
        return jsonify({"error": "No hay tareas en el rango de fechas", "no_tasks": True}), 404
    
    task_ids = result["task_id"].tolist()
    pending_tasks = db.get_pending_tasks(supervisor_id, task_ids)
    
    response_rows = []
    for _, row in result.iterrows():
        task_id = row["task_id"]
        is_reviewed = task_id not in pending_tasks
        
        review = None
        if is_reviewed:
            review = db.get_review_status(task_id, supervisor_id)
        
        raw_poc_id = clean_text(row.get("POC ID"))
        
        # Extraer el código corto del cliente
        short_poc_id = extract_short_poc_id(raw_poc_id)
        
        # Buscar información del cliente usando el código corto
        client_info = get_client_info(raw_poc_id)
        
        img_url = clean_text(row.get("Img", ""))
        if not img_url:
            img_url = clean_text(row.get("TaskImageUrl", ""))
        
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
    
    return jsonify(response_rows)

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
            "no_data": True
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
    items = [{"id": sid, "name": name} for sid, name in SUPERVISORS.items()]
    return jsonify(items)

# =====================================================
# ENDPOINTS DE PRUEBA (OPCIONALES)
# =====================================================

@app.route("/api/test_poc/<poc_id>")
@login_required
def test_poc(poc_id):
    """Probar extracción de POC ID"""
    short = extract_short_poc_id(poc_id)
    return jsonify({
        "original": poc_id,
        "extraido": short
    })

@app.route("/api/test_sheets")
@login_required
def test_sheets():
    """Probar conexión a Google Sheets"""
    try:
        master = load_client_master()
        return jsonify({
            "total_clientes": len(master),
            "primeros_5": dict(list(master.items())[:5]),
            "todos_los_ids": list(master.keys())[:10]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/verificar_cliente/<poc_id>")
@login_required
def verificar_cliente(poc_id):
    """Verificar si un cliente existe en Sheets"""
    client_info = get_client_info(poc_id)
    short_id = extract_short_poc_id(poc_id)
    master = load_client_master()
    
    return jsonify({
        "poc_id_recibido": poc_id,
        "codigo_extraido": short_id,
        "cliente_encontrado": client_info is not None,
        "cliente": client_info,
        "existe_en_sheets": short_id in master if master else False,
        "primeros_ids_sheets": list(master.keys())[:10] if master else []
    })

# =====================================================
# EXPORTAR Y CIERRE DE MES
# =====================================================

@app.route("/api/export_reviews", methods=["GET"])
@login_required
def api_export_reviews():
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
                    "Fecha": "",
                    "Promotor": "",
                    "Cliente ID": "",
                    "Razón Social": "",
                    "Direccion": "",
                    "Tarea": "",
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
                    "Fecha": task_info.get("fecha", ""),
                    "Promotor": task_info.get("promotor", ""),
                    "Cliente ID": task_info.get("cliente_id", ""),
                    "Razón Social": task_info.get("razon_social", ""),
                    "Direccion": task_info.get("direccion", ""),
                    "Tarea": task_info.get("detalle_tarea", ""),
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, sheet_name='Revisiones', index=False)
            
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

@app.route("/api/close_month", methods=["POST"])
@login_required
def api_close_month():
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
                    "Fecha": "",
                    "Promotor": "",
                    "Cliente ID": "",
                    "Razón Social": "",
                    "Direccion": "",
                    "Tarea": "",
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
                    "Fecha": task_info.get("fecha", ""),
                    "Promotor": task_info.get("promotor", ""),
                    "Cliente ID": task_info.get("cliente_id", ""),
                    "Razón Social": task_info.get("razon_social", ""),
                    "Direccion": task_info.get("direccion", ""),
                    "Tarea": task_info.get("detalle_tarea", ""),
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, sheet_name='Revisiones Cierre', index=False)
            
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
        
        return send_file(
            output,
            download_name=filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Error en cierre de mes: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)