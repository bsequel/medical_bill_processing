import pandas as pd

def excel_to_list(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path, header=None)

    # Drop completely empty rows
    df = df.dropna(how='all')

    # Define column types: C (float), D (int), E (float)
    column_types = {2: float, 3: int, 4: float}

    # Function to convert values based on column type
    def convert_value(val, col):
        try:
            if col in column_types:
                return column_types[col](float(val))  # Ensure proper conversion
            return val  # Keep other columns unchanged
        except ValueError:
            return val  # Return as-is if conversion fails

    # Apply conversion to specified columns
    for col, dtype in column_types.items():
        df[col] = df[col].apply(lambda x: convert_value(x, col))

    # Convert DataFrame to list of lists with E column duplicated
    data_list = [row + [row[4]] for row in df.values.tolist()]  # Duplicate column E (index 4)

    return data_list




# Example usage
file_path = r"c:\Users\Dell\Desktop/13.xlsx"  # Update this if needed
result = excel_to_list(file_path)

# Print the result
for row in result:
    print(row)

# # Example usage
# file_path = r"c:\Users\Dell\Desktop/12.xlsx"  # Update this if needed
# result = excel_to_list(file_path)

# # Print the result
# for row in result:
#     print(row)

import pyautogui
