# app.py
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
app.secret_key = "clave_secreta_para_desarrollo"  # Cambiar en producción

# Cache de datos
_cached_df = None
_current_excel = None
_data_loaded = False
_current_mes = None

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
    text = str(value).strip().upper()
    return text in {"VERDADERO", "TRUE", "1", "1.0"}

def _build_task_id(row: pd.Series) -> str:
    """Crear ID único para cada tarea"""
    parts = [
        clean_text(row.get("Imagen", "")),
        clean_text(row.get("POC ID", "")),
        str(row.get("Fecha", ""))
    ]
    raw = "|".join(parts)
    return hashlib.md5(raw.encode()).hexdigest()[:12]

def get_mes_actual():
    """Obtener el mes actual en formato YYYYMM"""
    return datetime.now().strftime("%Y%m")

# =====================================================
# LOAD DATA
# =====================================================

def _load_data(path: Path) -> pd.DataFrame:
    """Cargar y filtrar datos del Excel"""
    global _current_excel, _data_loaded, _current_mes
    
    df = pd.read_excel(path, engine="openpyxl")
    _current_excel = path.name
    _current_mes = get_mes_actual()
    
    required = [
        "Fecha", "Promotor", "POC ID", "Detalle Tarea", 
        "Imagen", "Completada", "Validada", "Visita Valida", 
        "Supervisor ID"
    ]
    
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Columnas faltantes: {missing}")
    
    df = df.copy()
    
    # Limpiar datos
    df["Fecha"] = pd.to_numeric(df["Fecha"], errors="coerce").astype("Int64")
    df["Completada"] = pd.to_numeric(df["Completada"], errors="coerce")
    df["Validada"] = pd.to_numeric(df["Validada"], errors="coerce")
    df["Supervisor ID"] = pd.to_numeric(df["Supervisor ID"], errors="coerce").astype("Int64")
    df["VisitaValidaBool"] = df["Visita Valida"].apply(_is_visita_valida)
    
    # Filtrar
    filtered = df[
        (df["Completada"] == 1.0) &
        (df["Validada"] == 0.0) &
        (df["VisitaValidaBool"]) &
        (df["Supervisor ID"].isin(SUPERVISORS.keys()))
    ].copy()
    
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
    
    # Guardar archivo
    file_path = UPLOAD_DIR / "data.xlsx"
    f.save(str(file_path))
    
    try:
        _cached_df = _load_data(file_path)
        _data_loaded = True
        _current_mes = get_mes_actual()
        
        # Registrar en base de datos
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
        return jsonify({
            "error": "No hay datos cargados. Por favor, sube un archivo Excel.",
            "no_data": True
        }), 404
    
    supervisor_id = session['supervisor_id']
    start_date = request.args.get("start_date", type=int)
    end_date = request.args.get("end_date", type=int)
    
    # Filtrar por supervisor
    result = _cached_df[_cached_df["Supervisor ID"] == supervisor_id].copy()
    
    if result.empty:
        return jsonify({
            "error": f"No hay tareas asignadas para {SUPERVISORS[supervisor_id]}",
            "no_tasks": True
        }), 404
    
    # Filtrar por fechas
    if start_date is not None:
        result = result[result["Fecha"] >= start_date]
    if end_date is not None:
        result = result[result["Fecha"] <= end_date]
    
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
        
        response_rows.append({
            "row_id": int(row["row_id"]),
            "task_id": task_id,
            "fecha": str(int(row["Fecha"])) if not pd.isna(row["Fecha"]) else "",
            "promotor": clean_text(row.get("Promotor")),
            "poc_id": clean_text(row.get("POC ID")),
            "detalle_tarea": clean_text(row.get("Detalle Tarea")),
            "imagen": clean_text(row.get("Imagen")),
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

# app.py - Reemplazar las funciones de exportación

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
        
        # Obtener todas las revisiones del supervisor
        reviews = db.get_all_reviews(supervisor_id, mes)
        
        if not reviews:
            return jsonify({"error": "No hay revisiones para exportar"}), 404
        
        # Obtener datos completos de las tareas revisadas
        if _cached_df is None or not _data_loaded:
            # Si no hay datos del Excel, solo exportar lo que está en la BD
            export_data = []
            for review in reviews:
                export_data.append({
                    "Fecha Ejecución": "",
                    "POC ID": "",
                    "Detalle Tarea": "",
                    "Imagen": "",
                    "Promotor": "",
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"],
                    "Nota": "Datos del Excel no disponibles"
                })
            df_export = pd.DataFrame(export_data)
        else:
            # Crear un diccionario para mapear task_id con datos del Excel
            task_data = {}
            for _, row in _cached_df.iterrows():
                task_data[row["task_id"]] = {
                    "fecha": row.get("Fecha", ""),
                    "promotor": row.get("Promotor", ""),
                    "poc_id": row.get("POC ID", ""),
                    "detalle_tarea": row.get("Detalle Tarea", ""),
                    "imagen": row.get("Imagen", "")
                }
            
            # Construir datos para el Excel con TODOS los campos
            export_data = []
            for review in reviews:
                task_id = review["task_id"]
                task_info = task_data.get(task_id, {})
                
                export_data.append({
                    "Fecha Ejecución": task_info.get("fecha", ""),
                    "POC ID": task_info.get("poc_id", ""),
                    "Detalle Tarea": task_info.get("detalle_tarea", ""),
                    "Imagen": task_info.get("imagen", ""),
                    "Promotor": task_info.get("promotor", ""),
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        
        # Crear archivo Excel en memoria
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Hoja 1: Detalle de revisiones con TODOS los campos
            df_export.to_excel(writer, sheet_name='Revisiones', index=False)
            
            # Hoja 2: Resumen por status
            if not df_export.empty and 'Status' in df_export.columns:
                df_status = df_export.groupby('Status').size().reset_index(name='Cantidad')
                df_status = df_status.sort_values('Cantidad', ascending=False)
                
                # Agregar total
                total_row = pd.DataFrame({
                    'Status': ['TOTAL'],
                    'Cantidad': [df_status['Cantidad'].sum()]
                })
                df_status = pd.concat([df_status, total_row], ignore_index=True)
                df_status.to_excel(writer, sheet_name='Resumen por Status', index=False)
            
            # Hoja 3: Resumen por promotor
            if not df_export.empty and 'Promotor' in df_export.columns and 'Status' in df_export.columns:
                df_promotor = df_export.groupby(['Promotor', 'Status']).size().unstack(fill_value=0)
                df_promotor['Total'] = df_promotor.sum(axis=1)
                df_promotor = df_promotor.sort_values('Total', ascending=False)
                df_promotor.to_excel(writer, sheet_name='Resumen por Promotor')
            
            # Hoja 4: Información del reporte
            info_data = {
                'Supervisor': [supervisor_name],
                'ID Supervisor': [supervisor_id],
                'Mes': [mes],
                'Fecha Exportación': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                'Total Revisiones': [len(reviews)]
            }
            df_info = pd.DataFrame(info_data)
            df_info.to_excel(writer, sheet_name='Info Reporte', index=False)
            
            # Autoajustar columnas en todas las hojas
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
# CIERRE DE MES (CON TODOS LOS CAMPOS) - CORREGIDO
# =====================================================

@app.route("/api/close_month", methods=["POST"])
@login_required
def api_close_month():
    """Exportar todas las revisiones con todos los campos y limpiar la base de datos"""
    # ✅ CORRECCIÓN: 'global' declarado al inicio de la función
    global _cached_df, _data_loaded
    
    try:
        supervisor_id = session.get('supervisor_id')
        supervisor_name = session.get('supervisor_name')
        mes = _current_mes or get_mes_actual()
        
        reviews = db.get_all_reviews(supervisor_id, mes)
        
        if not reviews:
            return jsonify({"error": "No hay revisiones para exportar"}), 404
        
        # Obtener datos completos de las tareas revisadas
        if _cached_df is None or not _data_loaded:
            # Si no hay datos del Excel, solo exportar lo que está en la BD
            export_data = []
            for review in reviews:
                export_data.append({
                    "Fecha Ejecución": "",
                    "POC ID": "",
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
            # Crear un diccionario para mapear task_id con datos del Excel
            task_data = {}
            for _, row in _cached_df.iterrows():
                task_data[row["task_id"]] = {
                    "fecha": row.get("Fecha", ""),
                    "promotor": row.get("Promotor", ""),
                    "poc_id": row.get("POC ID", ""),
                    "detalle_tarea": row.get("Detalle Tarea", ""),
                    "imagen": row.get("Imagen", "")
                }
            
            # Construir datos para el Excel con TODOS los campos
            export_data = []
            for review in reviews:
                task_id = review["task_id"]
                task_info = task_data.get(task_id, {})
                
                export_data.append({
                    "Fecha Ejecución": task_info.get("fecha", ""),
                    "POC ID": task_info.get("poc_id", ""),
                    "Detalle Tarea": task_info.get("detalle_tarea", ""),
                    "Imagen": task_info.get("imagen", ""),
                    "Promotor": task_info.get("promotor", ""),
                    "Status": review["status"],
                    "Observaciones": review.get("observaciones", ""),
                    "Supervisor": review["supervisor_name"],
                    "Fecha Revisión": review["fecha_revision"]
                })
            df_export = pd.DataFrame(export_data)
        
        # Crear archivo Excel en memoria
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Hoja 1: Detalle de revisiones con TODOS los campos
            df_export.to_excel(writer, sheet_name='Revisiones Cierre', index=False)
            
            # Hoja 2: Resumen por status
            if not df_export.empty and 'Status' in df_export.columns:
                df_status = df_export.groupby('Status').size().reset_index(name='Cantidad')
                df_status = df_status.sort_values('Cantidad', ascending=False)
                
                # Agregar total
                total_row = pd.DataFrame({
                    'Status': ['TOTAL'],
                    'Cantidad': [df_status['Cantidad'].sum()]
                })
                df_status = pd.concat([df_status, total_row], ignore_index=True)
                df_status.to_excel(writer, sheet_name='Resumen por Status', index=False)
            
            # Hoja 3: Resumen por promotor
            if not df_export.empty and 'Promotor' in df_export.columns and 'Status' in df_export.columns:
                df_promotor = df_export.groupby(['Promotor', 'Status']).size().unstack(fill_value=0)
                df_promotor['Total'] = df_promotor.sum(axis=1)
                df_promotor = df_promotor.sort_values('Total', ascending=False)
                df_promotor.to_excel(writer, sheet_name='Resumen por Promotor')
            
            # Hoja 4: Información del cierre
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
            
            # Autoajustar columnas en todas las hojas
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
        
        # Guardar una copia del archivo antes de limpiar
        filename = f"cierre_mes_{supervisor_name}_{mes}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Eliminar revisiones del mes actual
        db.clear_reviews(supervisor_id, mes)
        
        # Limpiar archivo Excel subido
        file_path = UPLOAD_DIR / "data.xlsx"
        if file_path.exists():
            file_path.unlink()
        
        # Resetear cache (ya declarado al inicio)
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