import os
import uuid
import pandas as pd
from flask import Flask, request, render_template, send_from_directory
from datetime import datetime

# Setup Flask app
app = Flask(__name__)

# Folder to save uploaded files
UPLOAD_FOLDER = 'download'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the uploads folder exists
if not os.path.exists(UPLOAD_FOLDER):
    try:
        os.makedirs(UPLOAD_FOLDER)
        print(f"Created directory: {UPLOAD_FOLDER}")
    except Exception as e:
        print(f"Error creating directory: {e}")
        raise

# Function to process the uploaded file
def load_sheets(file_path):
    attribute_mapping = pd.read_excel(file_path, sheet_name='attribute_mapping')
    product_file = pd.read_excel(file_path, sheet_name='product_file')
    option_sheet = pd.read_excel(file_path, sheet_name='option')

    product_file = product_file.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)
    return attribute_mapping, product_file, option_sheet


def create_pimattribute_with_filling(attribute_mapping, product_file):
    mapping_dict = attribute_mapping.groupby('PIMS Attr Name')['Product File Attr Name'].apply(list).to_dict()
    pimattribute_columns = {}
    for pims_attr, product_attrs in mapping_dict.items():
        primary_column = None
        for attr in product_attrs:
            if attr in product_file.columns:
                primary_column = product_file[attr].copy()
                break
        if primary_column is None:
            print(f"Warning: None of the attributes {product_attrs} found in product_file.")
            continue
        for additional_attr in product_attrs[1:]:
            if additional_attr in product_file.columns:
                primary_column = primary_column.combine_first(product_file[additional_attr])
        pimattribute_columns[pims_attr] = primary_column
    pimattribute = pd.concat(pimattribute_columns, axis=1)
    return pimattribute


def update_pimattribute(pimattribute, option_sheet):
    for _, row in option_sheet.iterrows():
        header = row['PIM_Attribute_Name']
        old_value = str(row['product_file_name_value']).strip()
        new_value = str(row['PIMS_Value']).strip()
        if header in pimattribute.columns:
            pimattribute[header] = pimattribute[header].fillna("").astype(str)
            if new_value:
                pimattribute[header] = pimattribute[header].apply(
                    lambda x: new_value if old_value.upper() == x.strip().upper() else x
                )
            else:
                pimattribute[header] = pimattribute[header].apply(
                    lambda x: "" if old_value.upper() == x.strip().upper() else x
                )
    return pimattribute


def save_to_excel(file_path, pimattribute):
    output_dir = app.config['UPLOAD_FOLDER']
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    unique_filename = f"processed_{uuid.uuid4().hex}_{os.path.basename(file_path)}"
    output_path = os.path.join(output_dir, unique_filename)

    with pd.ExcelWriter(output_path, mode='w', engine='openpyxl') as writer:
        pimattribute.to_excel(writer, sheet_name='pimattribute', index=False)
    
    print(f"File saved successfully: {output_path}")
    return output_path

# Route to display the upload form
@app.route('/')
def home():
    return render_template('upload.html')

# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    
    # Save the file to the uploads folder
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    print(f"Uploaded file saved to: {file_path}")


    start_time = datetime.now()
    print(f"Script started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    attribute_mapping, product_file, option_sheet = load_sheets(file_path)
    pimattribute = create_pimattribute_with_filling(attribute_mapping, product_file)
    pimattribute = update_pimattribute(pimattribute, option_sheet)
    output_file = save_to_excel(file_path, pimattribute)

        # Return the processed file as a download
    return send_from_directory(app.config['UPLOAD_FOLDER'], os.path.basename(output_file), as_attachment=True)
    
if __name__ == '__main__':
    app.run(debug=True, port=5001)
