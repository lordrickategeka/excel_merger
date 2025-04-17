# app.py
from flask import Flask, request, send_file
import os
import shutil
from merger import extract_zip, merge_excel_files, OUTPUT_FILE, EXTRACT_DIR

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return '''
        <h2>Upload ZIP of Excel Files</h2>
        <form method="post" action="/merge" enctype="multipart/form-data">
            <input type="file" name="zipfile" accept=".zip" required>
            <button type="submit">Upload & Merge</button>
        </form>
    '''

@app.route("/merge", methods=["POST"])
def merge():
    zip_file = request.files["zipfile"]
    zip_path = os.path.join(UPLOAD_FOLDER, zip_file.filename)
    zip_file.save(zip_path)

    if os.path.exists(EXTRACT_DIR):
        shutil.rmtree(EXTRACT_DIR)

    os.makedirs(EXTRACT_DIR, exist_ok=True)
    extract_zip(zip_path)
    output_file = merge_excel_files()

    if output_file:
        return send_file(output_file, as_attachment=True)
    else:
        return "No valid Excel files found or all had different headers.", 400

if __name__ == "__main__":
    app.run(debug=True)
