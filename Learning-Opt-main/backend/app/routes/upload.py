# backend/app/routes/upload.py
from flask import Blueprint, request, jsonify
import os
import json
import pandas as pd

bp = Blueprint('upload', __name__, url_prefix='/upload')

@bp.route('/excel', methods=['POST'])
def upload_excel():
    data = request.get_json()
    rows = data.get('rows', [])
    
    upload_path = os.path.join(os.path.dirname(__file__), '..', 'static', 'excel')
    os.makedirs(upload_path, exist_ok=True)
    file_path = os.path.join(upload_path, 'uploaded_data.json')

    if os.path.exists(file_path):
        os.remove(file_path)

    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(rows, f, indent=2)

    return jsonify({'message': 'Excel data uploaded successfully', 'rowCount': len(rows)})

@bp.route('/trainee', methods=['POST'])
def upload_trainee_file():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file uploaded'}), 400

    try:
        df = pd.read_excel(file)

        data = []
        for _, row in df.iterrows():
            data.append({
                'first_name': row.get('FIRST NAME', ''),
                'middle_name': row.get('MIDDLE NAME', ''),
                'last_name': row.get('LAST NAME', ''),
                'strand': row.get('STRAND', ''),
                'department': row.get('DEPARTMENT', ''),
                'school': row.get('SCHOOL', ''),
                'batch': row.get('BATCH', ''),
                'date_of_immersion': str(row.get('DATE OF IMMERSION', '')),
                'status': row.get('STATUS', '')
            })

        return jsonify({'students': data})

    except Exception as e:
        return jsonify({'error': str(e)}), 500