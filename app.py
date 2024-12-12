import os
import uuid
import pandas as pd
import zipfile
import re
from io import BytesIO
from flask import Flask, request, render_template, send_from_directory
from datetime import datetime
from tkinter import filedialog
import tkinter as tk

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

# Function to allow file selection for df1 and df2
def select_file(title="Select File"):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title=title)
    return file_path

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


def create_template():
    # Load the first dataset (df1) with file selection
    file1 = select_file("Select the first dataset (df1) file")
    df1 = pd.read_excel(file1)

    # Load the second dataset (df2) with file selection
    file2 = select_file("Select the second dataset (df2) file")
    df2 = pd.read_excel(file2)

    # Get unique family codes from df1
    Family_Names = df1['Family_Name'].unique()

    # Directory for saving template files
    final_template_dir = r"C:/final_template"  # Change the path as per your needs
    os.makedirs(final_template_dir, exist_ok=True)

    # Function to sanitize file names
    def sanitize_filename(filename):
        return re.sub(r'[<>:"/\\|?*]', '_', filename)

    # List to store Family Names missing 'mfg_part'
    missing_mfg_part_families = []

    # Loop through each family code and create template files
    for Family_Name in Family_Names:
        # Filter df1 based on the current family code
        filtered_df1 = df1[df1['Family_Name'] == Family_Name]
        # Filter df2 based on the current family code (if applicable)
        filtered_df2 = df2[df2['Family_Name'] == Family_Name] if 'Family_Name' in df2.columns else df2
        
        # Create an empty DataFrame with headers from filtered_df2
        template_df = pd.DataFrame(columns=filtered_df2['Attribute_Code'])
        
        # Check if 'mfg_part' is missing in the columns
        if "mfg_part" not in template_df.columns:
            missing_mfg_part_families.append(Family_Name)
            print(f"'mfg_part' is missing in template for Family_Name: {Family_Name}")
        
        # Populate the template_df with values from filtered_df1
        for col in template_df.columns:
            if col in filtered_df1.columns:
                template_df[col] = filtered_df1[col].values

        # Add columns with default values if necessary
        if "mfg_part" not in template_df.columns:
            template_df["mfg_part"] = ""
        
        template_df["Enable_In_EGC_Supply"] = 0
        template_df["mfg_part"] = template_df["mfg_part"].astype(str)
        template_df["enabled"] = 1
        template_df["special_item"] = 0

        # Sanitize the Family Name for file name
        sanitized_family_name = sanitize_filename(Family_Name)
        file_path = os.path.join(final_template_dir, f'{sanitized_family_name}_template.xlsx')

        # Save the template DataFrame to an Excel file
        template_df.to_excel(file_path, index=False)

    # Create the ZIP file from the directory
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for foldername, subfolders, filenames in os.walk(final_template_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, final_template_dir)  # Relative path for the archive
                zip_file.write(filepath, arcname)

    # Reset the buffer's position to the beginning
    zip_buffer.seek(0)

    # Send the ZIP file as a response
    return send_file(zip_buffer, as_attachment=True, download_name="final_template.zip", mimetype="application/zip")


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


# Route to handle template creation and download
@app.route('/create_template', methods=['POST'])
def create_template_route():
    return create_template()


if __name__ == '__main__':
    app.run(debug=True)
