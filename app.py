from __future__ import annotations
        "end_date",
        type=int
    )

    if supervisor_id not in SUPERVISORS:

        return jsonify({
            "error": "Invalid supervisor_id"
        }), 400

    result = df[
        df["Supervisor ID"] == supervisor_id
    ]

    if start_date is not None:
        result = result[result["Fecha"] >= start_date]

    if end_date is not None:
        result = result[result["Fecha"] <= end_date]

    response_rows = []

    for row in result.to_dict(orient="records"):

        poc_raw = row.get("POC ID")
        fecha_raw = row.get("Fecha")

        response_rows.append({

            "row_id": int(row["row_id"]),

            "fecha": (
                ""
                if pd.isna(fecha_raw)
                else str(int(fecha_raw))
            ),

            "promotor": clean_text(
                row.get("Promotor")
            ),

            "poc_id": (
                ""
                if pd.isna(poc_raw)
                else str(poc_raw).replace(".0", "")
            ),

            "detalle_tarea": clean_text(
                row.get("Detalle Tarea")
            ),

            "imagen": clean_text(
                row.get("Imagen")
            ),
        })

    return jsonify(response_rows)


# =====================================================
# START
# =====================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
