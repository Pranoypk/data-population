import os
import re
import pandas as pd
import zipfile
from flask import Flask, send_file
from io import BytesIO
from tkinter import filedialog
import tkinter as tk

app = Flask(__name__)

# Function to allow file selection for df1 and df2
def select_file(title="Select File"):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title=title)
    return file_path

@app.route('/upload', methods=['POST'])
def create_template():
    print("Creating template")

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

if __name__ == '__main__':
    app.run(debug=True)