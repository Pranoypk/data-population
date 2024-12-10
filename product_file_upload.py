import pandas as pd
from datetime import datetime

# Record the start time
start_time = datetime.now()
print(f"Script started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

# Function to load Excel sheets into DataFrames
def load_sheets(file_path):
    attribute_mapping = pd.read_excel(file_path, sheet_name='attribute_mapping')
    product_file = pd.read_excel(file_path, sheet_name='product_file')
    option_sheet = pd.read_excel(file_path, sheet_name='option')

    product_file = product_file.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)
    return attribute_mapping, product_file, option_sheet

# Function to create pimattribute sheet with deduplicated columns
def create_pimattribute_with_filling(attribute_mapping, product_file):
    # Create a dictionary to map each unique PIMS Attr Name to its corresponding Product File Attr Names
    mapping_dict = attribute_mapping.groupby('PIMS Attr Name')['Product File Attr Name'].apply(list).to_dict()

    # Dictionary to hold each processed column before concatenation
    pimattribute_columns = {}

    # Process each unique PIMS Attr Name
    for pims_attr, product_attrs in mapping_dict.items():
        # Check if the first product attribute exists in product_file
        primary_column = None
        for attr in product_attrs:
            if attr in product_file.columns:
                primary_column = product_file[attr].copy()
                break
        # Skip if none of the attributes exist in product_file
        if primary_column is None:
            print(f"Warning: None of the attributes {product_attrs} found in product_file.")
            continue

        # Fill NaN values in the primary column using other product attributes mapped to the same PIMS Attr Name
        for additional_attr in product_attrs[1:]:
            if additional_attr in product_file.columns:
                primary_column = primary_column.combine_first(product_file[additional_attr])

        # Add the filled column to the dictionary
        pimattribute_columns[pims_attr] = primary_column

    # Concatenate all columns at once
    pimattribute = pd.concat(pimattribute_columns, axis=1)

    return pimattribute

# Function to update pimattribute based on the option sheet with blank handling
# Updated function to update pimattribute based on the option sheet with strict matching
def update_pimattribute(pimattribute, option_sheet):
    for _, row in option_sheet.iterrows():
        header = row['PIM_Attribute_Name']
        old_value = str(row['product_file_name_value']).strip()  # Cleaned old value to look for
        new_value = str(row['PIMS_Value']).strip()  # Cleaned replacement value

        # Ensure the header exists in pimattribute columns
        if header in pimattribute.columns:
            # Fill NaNs and convert columns to string for consistency
            pimattribute[header] = pimattribute[header].fillna("").astype(str)

            # Apply replacement logic
            if new_value:
                # Replace old_value with new_value (case-insensitive)
                pimattribute[header] = pimattribute[header].apply(
                    lambda x: new_value if old_value.upper() == x.strip().upper() else x
                )
            else:
                # If new_value is blank, set matching old_value cells to blank (case-insensitive)
                pimattribute[header] = pimattribute[header].apply(
                    lambda x: "" if old_value.upper() == x.strip().upper() else x
                )

    return pimattribute

# Function to save the updated pimattribute sheet back to the Excel file
def save_to_excel(file_path, pimattribute):
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
        pimattribute.to_excel(writer, sheet_name='pimattribute', index=False)
    print("pimattribute sheet updated and saved successfully.")

def main(file_path):
    # Step 1: Load the necessary sheets
    attribute_mapping, product_file, option_sheet = load_sheets(file_path)
    
    # Step 2: Create the pimattribute sheet with deduplicated columns and NaN filling
    pimattribute = create_pimattribute_with_filling(attribute_mapping, product_file)
    
    # Step 3: Update the pimattribute sheet based on the option sheet
    pimattribute = update_pimattribute(pimattribute, option_sheet)
    
    # Step 4: Save the pimattribute sheet back to the Excel file
    save_to_excel(file_path, pimattribute)

# File path to your Excel file
file_path = r"C:\Users\pkayyala\Desktop\December\Sandvik_product_file_final.xlsx"
# Run the main function
main(file_path)

# Record the end time
end_time = datetime.now()
print(f"Script ended at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")

# Calculate and print the total duration
duration = end_time - start_time
print(f"Total execution time: {duration}")
