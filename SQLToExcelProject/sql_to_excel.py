import os
import re
import pandas as pd

def extract_column_names(sql_text):
    """
    Extract column names from the CREATE TABLE statement by processing the file line-by-line.
    It starts after the 'CREATE TABLE' line and collects column names from lines starting with '['.
    Stops when it reaches a line that begins with ')', indicating the end of the column definitions.
    """
    columns = []
    in_table_def = False
    for line in sql_text.splitlines():
        stripped = line.strip()
        # Look for the start of the CREATE TABLE statement
        if not in_table_def:
            if re.search(r'CREATE\s+TABLE', stripped, re.IGNORECASE):
                in_table_def = True
            continue

        # If we reached the end of the table definition, break out of the loop.
        if stripped.startswith(')'):
            break

        # If the line starts with a square bracket, assume it's a column definition.
        if stripped.startswith('['):
            # Extract the first bracketed item (the column name).
            m = re.match(r'\[([^\]]+)\]', stripped)
            if m:
                column_name = m.group(1)
                columns.append(column_name)
            else:
                print("Could not extract column name from:", stripped)
    return columns

def create_excel_from_sql(sql_file_path):
    try:
        # Open with errors='replace' to avoid UnicodeDecodeError issues.
        with open(sql_file_path, 'r', encoding='utf-8', errors='replace') as f:
            sql_text = f.read()
    except Exception as e:
        print(f"Error reading {sql_file_path}: {e}")
        return

    column_names = extract_column_names(sql_text)
    if not column_names:
        print(f"No columns found in the file: {sql_file_path}")
        return

    df = pd.DataFrame(columns=column_names)
    base_name = os.path.splitext(os.path.basename(sql_file_path))[0]
    excel_file_name = base_name + '.xlsx'
    df.to_excel(excel_file_name, index=False)
    print(f"Excel file '{excel_file_name}' created with columns: {column_names}")

def process_all_sql_files(directory):
    if not os.path.isdir(directory):
        print(f"Directory '{directory}' does not exist. Please check the path and create the folder if necessary.")
        exit(1)
    
    for file in os.listdir(directory):
        if file.lower().endswith('.sql'):
            sql_file_path = os.path.join(directory, file)
            create_excel_from_sql(sql_file_path)

if __name__ == "__main__":
    # Update the directory to your SQL files' folder
    sql_files_directory = r"YOURPATHGOESHERE"
    print("Current working directory:", os.getcwd())
    process_all_sql_files(sql_files_directory)
