from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import io, os, traceback

from excel_parser import procesar_archivo

app = Flask(__name__, static_folder="static", static_url_path="/static")
CORS(app)

ACCESS_KEY = os.environ.get("ACCESS_KEY", "silver2024")
EXTENSIONES_PERMITIDAS = (".xlsx", ".xls", ".xlsm")


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


# ── Verificar clave (solo valida, no procesa nada) ──────────────────────────
@app.route("/verificar-clave", methods=["POST"])
def verificar_clave():
    key = request.form.get("key", "")
    if key != ACCESS_KEY:
        return jsonify({"ok": False}), 401
    return jsonify({"ok": True})


# ── Procesar archivo (limpieza Silver) ──────────────────────────────────────
@app.route("/procesar", methods=["POST"])
def procesar():
    key = request.form.get("key", "")
    if key != ACCESS_KEY:
        return jsonify({"error": "Sesión inválida. Vuelve a iniciar sesión."}), 401

    if "file" not in request.files:
        return jsonify({"error": "No se recibió ningún archivo"}), 400

    archivo = request.files["file"]
    nombre = archivo.filename or ""
    if not nombre:
        return jsonify({"error": "Nombre de archivo vacío"}), 400

    if not nombre.lower().endswith(EXTENSIONES_PERMITIDAS):
        return jsonify({"error": "Solo se aceptan archivos .xlsx, .xls o .xlsm"}), 400

    try:
        contenido = archivo.read()
        resultado = procesar_archivo(contenido, nombre)
        resumen = resultado["resumen"]

        nombre_base = nombre.rsplit(".", 1)[0]
        app.config["_ultimo_excel"] = resultado["excel_out"]
        app.config["_ultimo_nombre"] = f"{nombre_base}_SilverV2.xlsx"

        return jsonify({"ok": True, "resumen": resumen})

    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception:
        traceback.print_exc()
        return jsonify({
            "error": ("Error inesperado al procesar el archivo. Verifica que las hojas "
                      "\"GR'S\", \"GRS (RM)\" y \"SQ00\" existan y que los encabezados "
                      "estén en la fila 2.")
        }), 500


# ── Descargar reporte SilverV2 ───────────────────────────────────────────────
@app.route("/descargar")
def descargar():
    excel = app.config.get("_ultimo_excel")
    nombre = app.config.get("_ultimo_nombre", "SilverV2.xlsx")
    if not excel:
        return jsonify({"error": "No hay reporte disponible. Procesa un archivo primero."}), 404
    return send_file(
        io.BytesIO(excel),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=nombre,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)