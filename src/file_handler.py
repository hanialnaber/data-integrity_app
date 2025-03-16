import pandas as pd
import io
import os

def read_file(file):
    """
    Read a file and return its data
    """
    # Get file extension
    file_extension = os.path.splitext(file.name)[1].lower()

    # Initialize result dictionary
    result = {
        "name": file.name,
        "type": None,
        "data": None,
        "sheet_names": []
    }

    # Read Excel file
    if file_extension in ['.xlsx', '.xls']:
        result["type"] = "excel"

        # Read the file content
        file_content = file.read()

        # Create a BytesIO object
        excel_data = io.BytesIO(file_content)

        # Use pandas ExcelFile to get sheet names
        with pd.ExcelFile(excel_data) as xls:
            result["sheet_names"] = xls.sheet_names

            # Read each sheet into a dictionary
            sheets_data = {}
            for sheet_name in xls.sheet_names:
                sheets_data[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

            result["data"] = sheets_data

    # Read CSV file
    elif file_extension == '.csv':
        result["type"] = "csv"

        # Read the file content
        file_content = file.read()

        # Create a StringIO object
        csv_data = io.StringIO(file_content.decode('utf-8'))

        # Read CSV data
        result["data"] = pd.read_csv(csv_data)

    return result