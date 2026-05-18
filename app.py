from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from flask import Flask, jsonify, render_template, request

BASE_DIR = Path(__file__).resolve().parent

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
UPLOADED_FILE = UPLOAD_DIR / "data.xlsx"

SUPERVISORS = {
    14: "Bruno Del Popolo",
    17: "Franco Vivani",
    41: "Claudio Raposo",
}

app = Flask(__name__)

# Cache del dataframe; None = todavía no hay archivo subido
_cached_df: pd.DataFrame | None = None


def _make_task_key(row: Any) -> str:
    """Genera una clave estable por tarea: Fecha_POCID_DetalleTarea."""
    fecha = "" if pd.isna(row["Fecha"]) else str(int(row["Fecha"]))
    poc = "" if pd.isna(row["POC ID"]) else str(row["POC ID"]).replace(".0", "")
    detalle = "" if pd.isna(row["Detalle Tarea"]) else str(row["Detalle Tarea"])
    return f"{fecha}_{poc}_{detalle}"


def _is_visita_valida(value: Any) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, (int, float)):
        return float(value) == 1.0
    text = str(value).strip().upper()
    return text in {"VERDADERO", "TRUE", "1", "1.0"}


def _load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    required = [
        "Fecha",
        "Promotor",
        "POC ID",
        "Detalle Tarea",
        "Imagen",
        "Completada",
        "Validada",
        "Visita Valida",
        "Supervisor ID",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Columnas faltantes en el Excel: {missing}")

    df = df.copy()
    df["Fecha"] = pd.to_numeric(df["Fecha"], errors="coerce").astype("Int64")
    df["Completada"] = pd.to_numeric(df["Completada"], errors="coerce")
    df["Validada"] = pd.to_numeric(df["Validada"], errors="coerce")
    df["Supervisor ID"] = pd.to_numeric(df["Supervisor ID"], errors="coerce").astype("Int64")
    df["VisitaValidaBool"] = df["Visita Valida"].apply(_is_visita_valida)

    filtered = df[
        (df["Completada"] == 1.0)
        & (df["Validada"] == 0.0)
        & (df["VisitaValidaBool"])
        & (df["Supervisor ID"].isin(SUPERVISORS.keys()))
    ].copy()

    filtered = filtered.reset_index().rename(columns={"index": "row_id"})
    filtered["task_key"] = filtered.apply(_make_task_key, axis=1)
    return filtered


def _get_df() -> pd.DataFrame | None:
    """Devuelve el dataframe cacheado, intentando cargarlo desde disco si aún no está en memoria."""
    global _cached_df
    if _cached_df is None and UPLOADED_FILE.exists():
        try:
            _cached_df = _load_data(UPLOADED_FILE)
        except Exception:
            pass
    return _cached_df


@app.route("/")
def index() -> str:
    return render_template("index.html", supervisors=SUPERVISORS)


@app.route("/api/has_file")
def api_has_file():
    return jsonify({"has_file": UPLOADED_FILE.exists()})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    global _cached_df
    if "file" not in request.files:
        return jsonify({"error": "No se envió ningún archivo."}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "El archivo debe ser un .xlsx"}), 400

    f.save(str(UPLOADED_FILE))
    try:
        _cached_df = _load_data(UPLOADED_FILE)
    except ValueError as e:
        UPLOADED_FILE.unlink(missing_ok=True)
        _cached_df = None
        return jsonify({"error": str(e)}), 422

    return jsonify({"ok": True, "rows": len(_cached_df)})


@app.route("/api/supervisors")
def api_supervisors():
    items = [{"id": sid, "name": name} for sid, name in SUPERVISORS.items()]
    return jsonify(items)


@app.route("/api/tasks")
def api_tasks():
    df = _get_df()
    if df is None:
        return jsonify({"error": "No hay archivo cargado. Subí el Excel primero."}), 404

    supervisor_id = request.args.get("supervisor_id", type=int)
    start_date = request.args.get("start_date", type=int)
    end_date = request.args.get("end_date", type=int)

    if supervisor_id not in SUPERVISORS:
        return jsonify({"error": "Invalid supervisor_id"}), 400

    result = df[df["Supervisor ID"] == supervisor_id]
    if start_date is not None:
        result = result[result["Fecha"] >= start_date]
    if end_date is not None:
        result = result[result["Fecha"] <= end_date]

    response_rows = []
    for row in result.to_dict(orient="records"):
        poc_raw = row.get("POC ID")
        fecha_raw = row.get("Fecha")
        response_rows.append(
            {
                "row_id": int(row.get("row_id")),
                "task_key": str(row.get("task_key", "")),
                "fecha": "" if pd.isna(fecha_raw) else str(int(fecha_raw)),
                "promotor": "" if pd.isna(row.get("Promotor")) else str(row.get("Promotor")),
                "poc_id": "" if pd.isna(poc_raw) else str(poc_raw).replace(".0", ""),
                "detalle_tarea": "" if pd.isna(row.get("Detalle Tarea")) else str(row.get("Detalle Tarea")),
                "imagen": "" if pd.isna(row.get("Imagen")) else str(row.get("Imagen")),
            }
        )

    return jsonify(response_rows)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)