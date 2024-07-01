import os
import pandas as pd
from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp'
DOWNLOAD_FOLDER = '/tmp'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# Function to load, clean, and save the Excel file
def load_clean_and_save_excel(file_path):
    # Load the Excel file
    file_extension = file_path.split('.')[-1]
    if file_extension == 'xls':
        excel_data = pd.read_excel(file_path, engine='xlrd')
    else:
        excel_data = pd.read_excel(file_path, engine='openpyxl')

    sheet_data = pd.read_excel(file_path, sheet_name='復原_工作表1', engine='xlrd' if file_extension == 'xls' else 'openpyxl')

    data_start_row = 5
    cleaned_data = pd.read_excel(file_path, sheet_name='復原_工作表1', skiprows=data_start_row, engine='xlrd' if file_extension == 'xls' else 'openpyxl')

    cleaned_data = cleaned_data.dropna(axis=1, how='all').reset_index(drop=True)

    headers = ['Category', 'Item Code', 'Description', '', '', 'Qty', '', '', 'CN', 'OR', 'AMT']
    sales_data = cleaned_data.iloc[5:].reset_index(drop=True)
    sales_data.columns = headers
    sales_data = sales_data.dropna(subset=['Item Code', 'Description'])

    cleaned_file_path = os.path.join(os.path.dirname(file_path), f"cleaned_{os.path.basename(file_path)}")
    sales_data.to_excel(cleaned_file_path, index=False, engine='openpyxl')

    return cleaned_file_path

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    
    try:
        cleaned_file_path = load_clean_and_save_excel(file_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    download_url = request.url_root + 'download/' + os.path.basename(cleaned_file_path)
    
    response = {
        "message": "File processed successfully",
        "download_url": download_url
    }
    
    return jsonify(response), 200

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
