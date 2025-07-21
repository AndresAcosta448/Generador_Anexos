# app.py
import os
import zipfile
import tempfile


from flask import (
    Flask, render_template, request,
    send_file, flash, redirect, url_for
)
from anexos_logic import run_all_anexos

app = Flask(__name__)
app.secret_key = "cambiar_por_una_clave_segura"

@app.route("/", methods=("GET","POST"))
def upload():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or not f.filename.lower().endswith(".xlsx"):
            flash("Por favor sube un archivo .xlsx válido.")
            return redirect(request.url)

        # Guardar XLSX en carpeta temporal
        tmpdir = tempfile.mkdtemp()
        in_path = os.path.join(tmpdir, f.filename)
        f.save(in_path)

        # Carpeta raíz de salida
        out_root = os.path.join(tmpdir, "Anexos_Todos")
        os.makedirs(out_root, exist_ok=True)

        try:
            run_all_anexos(in_path, out_root)
        except Exception as e:
            flash(f"Error durante el procesamiento: {e}")
            return redirect(request.url)

        # Empaquetar en ZIP
        zip_path = os.path.join(tmpdir, "anexos.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(out_root):
                for name in files:
                    full = os.path.join(root, name)
                    rel  = os.path.relpath(full, out_root)
                    zf.write(full, rel)

        return send_file(zip_path,
                         as_attachment=True,
                         download_name="anexos.zip")

    return render_template("upload.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
