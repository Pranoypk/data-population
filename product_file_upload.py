from memory_profiler import profile

# Record the start time
from datetime import datetime
import pandas as pd

start_time = datetime.now()
print(f"Script started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

@profile
def load_sheets(file_path):
    attribute_mapping = pd.read_excel(file_path, sheet_name='attribute_mapping')
    product_file = pd.read_excel(file_path, sheet_name='product_file')
    option_sheet = pd.read_excel(file_path, sheet_name='option')

    product_file = product_file.apply(lambda col: col.str.strip() if col.dtype == 'object' else col)
    return attribute_mapping, product_file, option_sheet

@profile
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

@profile
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

@profile
def save_to_excel(file_path, pimattribute):
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
        pimattribute.to_excel(writer, sheet_name='pimattribute', index=False)
    print("pimattribute sheet updated and saved successfully.")

@profile
def main(file_path):
    attribute_mapping, product_file, option_sheet = load_sheets(file_path)
    pimattribute = create_pimattribute_with_filling(attribute_mapping, product_file)
    pimattribute = update_pimattribute(pimattribute, option_sheet)
    save_to_excel(file_path, pimattribute)

# File path to your Excel file
file_path = r"https://raw.githubusercontent.com/Pranoypk/data-population/main/Ceratizit_structured_file_final.xlsx"

main(file_path)

end_time = datetime.now()
print(f"Script ended at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
duration = end_time - start_time
print(f"Total execution time: {duration}")
