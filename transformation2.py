import json
from openpyxl import load_workbook, Workbook
from pprint import pprint
import sys

# Define file paths
input_file_path = "./test/testInputFiles/input1.xlsm"
output_file_path = "./test/testOutputFiles/output.xlsm"
mapping_file_path = "mapping2.json"

# Load the mapping schema globally
with open(mapping_file_path, 'r') as file:
    mapping_schema = json.load(file)

# Define transformation functions
def split_name(name, separator=' '):
    parts = [part for part in name.strip().split(separator) if part]
    if not parts:
        return '', ' '
    return parts[0], parts[1] if len(parts) > 1 else ' '

def capitalize(text):
    if isinstance(text, list):
        return [capitalize(sub_text) for sub_text in text]
    return text.capitalize()

def convert_to_integer(text):
    return int(text)

# Map transformation names to functions
transformation_functions = {
    'split_name': split_name,
    'capitalize': capitalize,
    'convert_to_integer': convert_to_integer
}

def apply_transformations(data, transformations):
    for transformation in transformations:
        func = transformation_functions.get(transformation)
        if func:
            if isinstance(data, list):
                transformed_data = []
                for item in data:
                    if isinstance(item, list):
                        result = [func(sub_item) for sub_item in item]
                    else:
                        result = func(item)
                    if isinstance(result, tuple):
                        transformed_data.append(list(result))
                    else:
                        transformed_data.append(result)
                data = transformed_data
            else:
                data = func(data)
    return data

def validate_data(data, validations):
    for validation in validations:
        if validation['type'] == 'required':
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, list):
                        if all(not sub_item.strip() for sub_item in item):
                            raise ValueError(validation['message'])
                    elif not item.strip():
                        raise ValueError(validation['message'])
            elif not data.strip():
                raise ValueError(validation['message'])

def read_data_from_input(input_path, mapping_schema):
    # Load the workbook and select the active sheet
    workbook = load_workbook(input_path, data_only=True)
    
    data_store = {}
    
    for mapping in mapping_schema['mappings']:
        field_name = mapping['field_name']
        source = mapping['source']
        sheet_name = source['sheet']
        cell_range = source['range']
        sheet = workbook[sheet_name]
        
        # Handle special cases for dynamic ranges
        if cell_range.endswith('_'):
            start_cell = cell_range.split(':')[0]
            col_letter = start_cell[0]
            start_row = int(start_cell[1:])
            data = []
            for row in sheet.iter_rows(min_row=start_row, min_col=ord(col_letter) - ord('A') + 1, max_col=ord(col_letter) - ord('A') + 1, values_only=True):
                if all(cell is None for cell in row):
                    continue
                data.append(row[0])
        else:
            # Handle fixed ranges and single cells
            if ':' in cell_range:
                start_cell, end_cell = cell_range.split(':')
                start_col = ord(start_cell[0]) - ord('A') + 1
                start_row = int(start_cell[1:])
                end_col = ord(end_cell[0]) - ord('A') + 1
                end_row = int(end_cell[1:])
                data = []
                for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col, values_only=True):
                    if all(cell is None for cell in row):
                        continue
                    data.extend([cell for cell in row if cell is not None])
            else:
                # Single cell case
                cell = sheet[cell_range]
                data = cell.value
        
        # Validate data if any validations are specified
        if 'validation' in mapping:
            print(f"Validating {field_name}")
            validate_data(data, mapping['validation'])
        
        # Apply transformations if any
        if 'transformations' in mapping:
            print(f"Applying transformations to {field_name}")
            data = apply_transformations(data, mapping['transformations'])
        
        # Store the data
        data_store[field_name] = data
        # Print the data
        # print(f"{field_name}: {data}")
    pprint(data_store)
    # sys.exit("Stopping the program!") # Currently stopped here for debugging
    return data_store

def use_mapping_generate_output(data_store, mapping_schema, output_file_path):
    # Load the existing workbook if it exists, otherwise create a new one
    try:
        output_workbook = load_workbook(output_file_path)
    except FileNotFoundError:
        output_workbook = Workbook()
    
    # Iterate over each mapping in the schema
    for mapping in mapping_schema['mappings']:
        field_name = mapping['field_name']
        destinations = mapping['destination']
        data = data_store.get(field_name, [])
        
        # Iterate over each destination for the current field
        for destination in destinations:
            sheet_name = destination['sheet']
            cell_ranges = destination['range'].split(',')  # e.g., "D6:D_,E6:E_"

            # Get or create the sheet in the output workbook
            if sheet_name in output_workbook.sheetnames:
                sheet = output_workbook[sheet_name]
            else:
                sheet = output_workbook.create_sheet(sheet_name)

            # Iterate over each cell range specified for the destination
            for idx, cell_range in enumerate(cell_ranges):
                if ':' in cell_range:
                    start_cell, end_cell = cell_range.split(':')
                    start_col = ord(start_cell[0]) - ord('A') + 1
                    start_row = int(start_cell[1:])
                    if end_cell.endswith('_'):
                        end_col = ord(end_cell[0]) - ord('A') + 1
                        end_row = start_row + len(data) - 1
                    else:
                        end_col = ord(end_cell[0]) - ord('A') + 1
                        end_row = int(end_cell[1:])
                    
                    # Write data to the range
                    for row_idx, row in enumerate(sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col)):
                        for cell_idx, cell in enumerate(row):
                            if isinstance(data, list) and isinstance(data[0], list):
                                if row_idx < len(data) and cell_idx < len(data[row_idx]):
                                    cell.value = data[row_idx][cell_idx]
                            else:
                                if row_idx < len(data):
                                    cell.value = data[row_idx]
                else:
                    # Handle single cell case
                    cell = sheet[cell_range]
                    if isinstance(data, list) and data:
                        cell.value = data[idx] if idx < len(data) else None
                    else:
                        cell.value = data

    # Save the workbook to the output file path
    output_workbook.save(output_file_path)

# Call the function to read data from input
data_store = read_data_from_input(input_file_path, mapping_schema)

# Call the function to generate the output file
use_mapping_generate_output(data_store, mapping_schema, output_file_path)