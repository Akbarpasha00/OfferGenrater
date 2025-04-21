# letter_generator.py
from flask import Flask, request, send_file, jsonify
import uuid
import os
import pandas as pd
import zipfile
from werkzeug.utils import secure_filename

# Check for templating support
try:
    from docxtpl import DocxTemplate
    DOCX_TEMPLATE_AVAILABLE = True
except ImportError:
    DOCX_TEMPLATE_AVAILABLE = False

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
TEMPLATE_FOLDER = os.path.join(UPLOAD_FOLDER, "templates")
OUTPUT_FOLDER = "outputs"

# Ensure folders exist
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    """
    Health check endpoint.
    Expected behavior: return status 200 and JSON with message.
    """
    return jsonify({"message": "Letter generation service is running."}), 200

@app.route('/upload-template', methods=['POST'])
def upload_template():
    """
    Upload a company template.
    Expects form-data with keys:
      - company: company name
      - template: .docx file
    Returns JSON with message and company on success.
    """
    if 'company' not in request.form or 'template' not in request.files:
        return jsonify({"error": "Missing company name or template file"}), 400
    company = request.form['company']
    template = request.files['template']
    filename = secure_filename(f"{company}.docx")
    path = os.path.join(TEMPLATE_FOLDER, filename)
    template.save(path)
    return jsonify({"message": "Template uploaded successfully", "company": company}), 200

@app.route('/generate-letters', methods=['POST'])
def generate_letters():
    """
    Generate letters for uploaded student data.
    Requires docxtpl; returns error if unavailable.
    Expects form-data with keys:
      - company: template to use
      - file: .csv or .xlsx of student data
    Returns a ZIP file of .docx letters.
    """
    if not DOCX_TEMPLATE_AVAILABLE:
        return jsonify({"error": "Templating support unavailable. Install docxtpl."}), 500
    if 'company' not in request.form or 'file' not in request.files:
        return jsonify({"error": "Missing company or student file"}), 400
    company = request.form['company']
    file = request.files['file']
    try:
        data = pd.read_excel(file) if file.filename.endswith(".xlsx") else pd.read_csv(file)
    except Exception as e:
        return jsonify({"error": f"Failed to read file: {e}"}), 400

    template_path = os.path.join(TEMPLATE_FOLDER, f"{company}.docx")
    if not os.path.exists(template_path):
        return jsonify({"error": f"Template not found for {company}"}), 404

    zip_name = f"letters_{uuid.uuid4().hex}.zip"
    zip_path = os.path.join(OUTPUT_FOLDER, zip_name)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for _, row in data.iterrows():
            context = row.to_dict()
            try:
                doc = DocxTemplate(template_path)
                doc.render(context)
                fname = f"{row.get('name','student')}_{company}_{uuid.uuid4().hex[:8]}.docx"
                out_path = os.path.join(OUTPUT_FOLDER, fname)
                doc.save(out_path)
                zipf.write(out_path, arcname=fname)
                os.remove(out_path)
            except Exception as e:
                return jsonify({"error": f"Row error: {e}"}), 500
    return send_file(zip_path, as_attachment=True, download_name=zip_name)

if __name__ == '__main__':
    app.run(debug=True)
