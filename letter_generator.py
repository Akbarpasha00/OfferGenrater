from flask import Flask, request, send_file, jsonify
import uuid
import os
import pandas as pd
import zipfile
from werkzeug.utils import secure_filename

# Ensure templating support
try:
    from docxtpl import DocxTemplate
except ImportError:
    DocxTemplate = None

app = Flask(__name__)

# Directory setup
UPLOAD_FOLDER = "uploads"
TEMPLATE_FOLDER = os.path.join(UPLOAD_FOLDER, "templates")
OUTPUT_FOLDER = "outputs"

os.makedirs(TEMPLATE_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET'])
def health_check():
    return jsonify({"status": "ok", "message": "Service running"}), 200

@app.route('/upload-template', methods=['POST'])
def upload_template():
    if 'company' not in request.form or 'template' not in request.files:
        return jsonify({"error": "Missing 'company' or 'template'"}), 400
    company = request.form['company'].strip()
    template_file = request.files['template']
    filename = secure_filename(f"{company}.docx")
    path = os.path.join(TEMPLATE_FOLDER, filename)
    template_file.save(path)
    return jsonify({"message": f"Template '{company}' uploaded successfully"}), 200

@app.route('/generate-letters', methods=['POST'])
def generate_letters():
    if not DocxTemplate:
        return jsonify({"error": "Template engine unavailable. Install 'docxtpl'."}), 500
    if 'company' not in request.form or 'file' not in request.files:
        return jsonify({"error": "Missing 'company' or data file"}), 400
    company = request.form['company'].strip()
    data_file = request.files['file']
    # Read data
    try:
        if data_file.filename.endswith('.xlsx'):
            df = pd.read_excel(data_file)
        else:
            df = pd.read_csv(data_file)
    except Exception as e:
        return jsonify({"error": f"Failed to read data: {e}"}), 400
    # Load template
    tpl_path = os.path.join(TEMPLATE_FOLDER, f"{company}.docx")
    if not os.path.exists(tpl_path):
        return jsonify({"error": f"Template for '{company}' not found"}), 404
    # Generate letters zip
    zip_name = f"letters_{company}_{uuid.uuid4().hex}.zip"
    zip_path = os.path.join(OUTPUT_FOLDER, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for idx, row in df.iterrows():
            context = {
                'Student_Name': row.get('Name of the student', '').strip(),
                'ref_id': row.get('Reference no', '').strip()
            }
            doc = DocxTemplate(tpl_path)
            doc.render(context)
            safe_name = context['Student_Name'].replace(' ', '_') or f"student_{idx}"
            docx_name = f"{safe_name}_{company}_{idx+1}.docx"
            out_path = os.path.join(OUTPUT_FOLDER, docx_name)
            doc.save(out_path)
            zipf.write(out_path, arcname=docx_name)
            os.remove(out_path)
    return send_file(zip_path, as_attachment=True, download_name=zip_name)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
