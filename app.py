# app.py - SOLUCIÓN DEFINITIVA
from __future__ import annotations
from pathlib import Path
from typing import Any
import pandas as pd
from flask import Flask, jsonify, render_template, request, session
from functools import wraps
import hashlib
from datetime import datetime
import json
import re
import os
from tempfile import NamedTemporaryFile

# =====================================================
# IMPORTS GOOGLE SHEETS
# =====================================================
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False

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

# =====================================================
# CONFIGURACIÓN GOOGLE SHEETS
# =====================================================

GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS")
SHEET_ID = os.environ.get("SHEET_ID", "12D-Ru8GNm0EE0NFg4kpcygqcPdqKpor9U4NOA-1TQRs")
SHEET_NAME = os.environ.get("SHEET_NAME", "Clientes")

CREDENTIALS_FILE = "credentials.json"
if GOOGLE_CREDENTIALS_JSON:
    try:
        creds_data = json.loads(GOOGLE_CREDENTIALS_JSON)
        temp_creds = NamedTemporaryFile(mode='w', suffix='.json', delete=False)
        json.dump(creds_data, temp_creds)
        temp_creds.close()
        CREDENTIALS_FILE = temp_creds.name
    except:
        pass

GOOGLE_SHEETS_CONFIG = {
    "sheet_id": SHEET_ID,
    "sheet_name": SHEET_NAME,
    "credentials_file": CREDENTIALS_FILE
}

# =====================================================
# CLIENTES DESDE JSON
# =====================================================

def cargar_clientes_json():
    try:
        json_path = BASE_DIR / "clientes.json"
        if json_path.exists():
            with open(json_path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    return {}

CLIENTES = cargar_clientes_json()

# =====================================================
# FUNCIONES GOOGLE SHEETS
# =====================================================

def get_google_sheet_client():
    if not GOOGLE_SHEETS_AVAILABLE:
        return None
    try:
        if not os.path.exists(GOOGLE_SHEETS_CONFIG["credentials_file"]):
            return None
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            GOOGLE_SHEETS_CONFIG["credentials_file"], scope
        )
        return gspread.authorize(creds)
    except:
        return None

def guardar_status_en_sheets(fecha, id_imagen, status, supervisor, observaciones=""):
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return False
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        
        try:
            worksheet = sheet.worksheet("Status")
        except:
            worksheet = sheet.add_worksheet(title="Status", rows=1000, cols=10)
            worksheet.append_row(["Fecha", "ID", "Status", "Supervisor", "Observaciones", "Fecha Revision"])
        
        id_column = worksheet.col_values(2)
        row_to_update = None
        if id_imagen in id_column:
            row_to_update = id_column.index(id_imagen) + 1
        
        fecha_revision = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        if row_to_update:
            worksheet.update(f'C{row_to_update}:F{row_to_update}', [[status, supervisor, observaciones, fecha_revision]])
        else:
            worksheet.append_row([fecha, id_imagen, status, supervisor, observaciones, fecha_revision])
        
        return True
    except:
        return False

def get_status_from_sheets(id_imagen):
    try:
        gc = get_google_sheet_client()
        if gc is None:
            return None
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        
        try:
            worksheet = sheet.worksheet("Status")
        except:
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
    except:
        return None

# =====================================================
# FUNCIONES BASE
# =====================================================

def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()

def extract_short_poc_id(poc_id):
    if not poc_id:
        return ""
    poc_id = str(poc_id).strip()
    if poc_id.endswith('.0'):
        poc_id = poc_id[:-2]
    poc_id = poc_id.replace('-', '').replace(' ', '')
    if poc_id.startswith("053822"):
        return poc_id[6:].lstrip('0') or "0"
    if poc_id.startswith("53822"):
        return poc_id[5:].lstrip('0') or "0"
    return poc_id.lstrip('0') or "0"

def get_client_info(poc_id):
    if not poc_id:
        return None
    short_id = extract_short_poc_id(poc_id)
    if short_id in CLIENTES:
        return CLIENTES[short_id]
    return None

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
            return f"{partes[2].split(' ')[0]}/{partes[1]}/{partes[0]}"
    if len(fecha_str) >= 8 and fecha_str[:8].isdigit():
        return f"{fecha_str[6:8]}/{fecha_str[4:6]}/{fecha_str[0:4]}"
    return fecha_str

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'supervisor_id' not in session:
            return jsonify({"error": "No autorizado"}), 401
        return f(*args, **kwargs)
    return decorated_function

# =====================================================
# RUTAS
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
    return jsonify({"ok": True, "supervisor_id": supervisor_id, "supervisor_name": SUPERVISORS[supervisor_id]})

@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"ok": True})

@app.route("/api/has_file")
@login_required
def api_has_file():
    file_exists = (UPLOAD_DIR / "data.xlsx").exists()
    return jsonify({"has_file": file_exists, "file_exists": file_exists})

@app.route("/api/upload", methods=["POST"])
@login_required
def api_upload():
    if "file" not in request.files:
        return jsonify({"error": "No se envió archivo"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "El archivo debe ser .xlsx"}), 400
    file_path = UPLOAD_DIR / "data.xlsx"
    f.save(str(file_path))
    
    # Verificar que se cargó correctamente
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        return jsonify({"ok": True, "rows": len(df), "message": f"Archivo cargado con {len(df)} registros"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/tasks")
@login_required
def api_tasks():
    file_path = UPLOAD_DIR / "data.xlsx"
    if not file_path.exists():
        return jsonify({"error": "No hay datos cargados", "no_data": True}), 404
    
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        supervisor_id = session.get('supervisor_id')
        
        # Filtrar por supervisor
        if 'Supervisor ID' in df.columns:
            df = df[df['Supervisor ID'] == supervisor_id]
        
        if df.empty:
            return jsonify({"error": "No hay tareas para este supervisor", "no_tasks": True}), 404
        
        # Ordenar por fecha
        if 'Fecha' in df.columns:
            df = df.sort_values(by="Fecha", ascending=True)
        
        # Renombrar columna de imagen
        if "TaskImageUrl" in df.columns and "Img" not in df.columns:
            df = df.rename(columns={"TaskImageUrl": "Img"})
        
        if "Img" not in df.columns:
            df["Img"] = ""
        
        response = []
        for _, row in df.iterrows():
            img_url = clean_text(row.get("Img", ""))
            poc_id = clean_text(row.get("POC ID", ""))
            
            # Buscar status en Sheets
            status_info = get_status_from_sheets(img_url) if img_url else None
            
            response.append({
                "row_id": int(_),
                "task_id": f"task_{_}",
                "fecha": formatear_fecha(row.get("Fecha")),
                "promotor": clean_text(row.get("Promotor")),
                "poc_id": extract_short_poc_id(poc_id),
                "poc_id_completo": poc_id,
                "cliente_id": extract_short_poc_id(poc_id),
                "razon_social": get_client_info(poc_id).get("nombre") if get_client_info(poc_id) else "SIN DATO",
                "direccion": get_client_info(poc_id).get("direccion") if get_client_info(poc_id) else "SIN DATO",
                "detalle_tarea": clean_text(row.get("Detalle Tarea")),
                "imagen": img_url,
                "revisado": status_info is not None,
                "status": status_info.get("status") if status_info else None,
                "observaciones": status_info.get("observaciones") if status_info else "",
                "fecha_revision": status_info.get("fecha_revision") if status_info else None
            })
        
        return jsonify(response)
        
    except Exception as e:
        print(f"❌ Error: {e}")
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
    
    file_path = UPLOAD_DIR / "data.xlsx"
    if not file_path.exists():
        return jsonify({"error": "No hay datos cargados"}), 404
    
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        row_idx = int(task_id.split("_")[1]) if "_" in task_id else 0
        if row_idx >= len(df):
            return jsonify({"error": "Tarea no encontrada"}), 404
        
        row = df.iloc[row_idx]
        supervisor_name = session['supervisor_name']
        fecha_tarea = formatear_fecha(row.get("Fecha"))
        img_url = clean_text(row.get("Img", row.get("TaskImageUrl", "")))
        
        success = guardar_status_en_sheets(fecha_tarea, img_url, status, supervisor_name, observaciones)
        
        if success:
            return jsonify({"ok": True, "message": "Revisión guardada correctamente"})
        else:
            return jsonify({"error": "Error al guardar en Sheets"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/supervisors")
def api_supervisors():
    return jsonify([{"id": sid, "name": name} for sid, name in SUPERVISORS.items()])

@app.route("/api/stats")
@login_required
def api_stats():
    try:
        supervisor_name = session.get('supervisor_name')
        gc = get_google_sheet_client()
        if gc is None:
            return jsonify({"total_revisados": 0, "total_pendientes": 0, "by_status": {"objeccion": 0, "invalida": 0, "fraude": 0}, "porcentaje": 0})
        
        sheet = gc.open_by_key(GOOGLE_SHEETS_CONFIG["sheet_id"])
        try:
            worksheet = sheet.worksheet("Status")
            records = worksheet.get_all_records()
        except:
            return jsonify({"total_revisados": 0, "total_pendientes": 0, "by_status": {"objeccion": 0, "invalida": 0, "fraude": 0}, "porcentaje": 0})
        
        stats = {"total": 0, "by_status": {"objeccion": 0, "invalida": 0, "fraude": 0}}
        for row in records:
            if row.get("Supervisor") == supervisor_name:
                status = row.get("Status", "")
                stats["total"] += 1
                if status in stats["by_status"]:
                    stats["by_status"][status] += 1
        
        file_path = UPLOAD_DIR / "data.xlsx"
        if file_path.exists():
            df = pd.read_excel(file_path, engine="openpyxl")
            supervisor_id = session.get('supervisor_id')
            if 'Supervisor ID' in df.columns:
                total_disponibles = len(df[df['Supervisor ID'] == supervisor_id])
            else:
                total_disponibles = len(df)
            pendientes = max(0, total_disponibles - stats["total"])
            porcentaje = round((stats["total"] / total_disponibles * 100) if total_disponibles > 0 else 0, 1)
        else:
            pendientes = 0
            porcentaje = 0
        
        return jsonify({
            "total_revisados": stats["total"],
            "total_pendientes": pendientes,
            "by_status": stats["by_status"],
            "porcentaje": porcentaje,
            "has_data": True
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)