import pandas as pd
import numpy as np
from collections import defaultdict

def compare_files(data1, data2):
    """
    Compare two files and return detailed report, summary report, and error details
    """
    detailed_report = []
    summary_report = []

    # Initialize error details structure
    error_details = {
        "missing_sheets": [],
        "extra_sheets": [],
        "column_differences": {},
        "row_differences": {},
        "value_differences": {}
    }

    # Compare file types
    if data1["type"] != data2["type"]:
        detailed_report.append(f"File types are different: {data1['type']} vs {data2['type']}")
        summary_report.append(f"File types are different: {data1['type']} vs {data2['type']}")

    # Compare sheet names (for Excel files)
    if data1["type"] == "excel" and data2["type"] == "excel":
        # Check for missing and extra sheets
        missing_sheets = set(data1["sheet_names"]) - set(data2["sheet_names"])
        extra_sheets = set(data2["sheet_names"]) - set(data1["sheet_names"])

        if missing_sheets:
            error_details["missing_sheets"] = list(missing_sheets)
            for sheet in missing_sheets:
                detailed_report.append(f"Sheet '{sheet}' is in file 1 but missing in file 2")
                summary_report.append(f"Sheet '{sheet}' is missing in file 2")

        if extra_sheets:
            error_details["extra_sheets"] = list(extra_sheets)
            for sheet in extra_sheets:
                detailed_report.append(f"Sheet '{sheet}' is in file 2 but missing in file 1")
                summary_report.append(f"Extra sheet '{sheet}' in file 2")

        # Compare common sheets
        common_sheets = set(data1["sheet_names"]) & set(data2["sheet_names"])

        for sheet in common_sheets:
            sheet_detailed_report, sheet_summary_report, sheet_error_details = compare_sheets(
                sheet, data1["data"][sheet], data2["data"][sheet]
            )

            detailed_report.extend(sheet_detailed_report)
            summary_report.extend(sheet_summary_report)

            # Update error details
            if sheet_error_details["column_differences"]:
                error_details["column_differences"][sheet] = sheet_error_details["column_differences"]

            if sheet_error_details["row_differences"]:
                error_details["row_differences"][sheet] = sheet_error_details["row_differences"]

            if sheet_error_details["value_differences"]:
                error_details["value_differences"][sheet] = sheet_error_details["value_differences"]

    # Compare CSV files
    elif data1["type"] == "csv" and data2["type"] == "csv":
        sheet_detailed_report, sheet_summary_report, sheet_error_details = compare_sheets(
            "data", data1["data"], data2["data"]
        )

        detailed_report.extend(sheet_detailed_report)
        summary_report.extend(sheet_summary_report)

        # Update error details
        if sheet_error_details["column_differences"]:
            error_details["column_differences"]["data"] = sheet_error_details["column_differences"]

        if sheet_error_details["row_differences"]:
            error_details["row_differences"]["data"] = sheet_error_details["row_differences"]

        if sheet_error_details["value_differences"]:
            error_details["value_differences"]["data"] = sheet_error_details["value_differences"]

    return detailed_report, summary_report, error_details

def compare_sheets(sheet_name, df1, df2):
    """
    Compare two dataframes and return detailed report, summary report, and error details
    """
    detailed_report = []
    summary_report = []

    # Initialize error details structure for this sheet
    error_details = {
        "column_differences": {},
        "row_differences": {},
        "value_differences": []
    }

    # Compare column names and order
    column_differences = compare_columns(df1, df2)
    error_details["column_differences"] = column_differences

    if column_differences["missing"]:
        for col in column_differences["missing"]:
            detailed_report.append(f"Column '{col}' in sheet '{sheet_name}' is in file 1 but missing in file 2")
            summary_report.append(f"Missing column '{col}' in sheet '{sheet_name}'")

    if column_differences["extra"]:
        for col in column_differences["extra"]:
            detailed_report.append(f"Column '{col}' in sheet '{sheet_name}' is in file 2 but missing in file 1")
            summary_report.append(f"Extra column '{col}' in sheet '{sheet_name}'")

    if column_differences["reordered"]:
        detailed_report.append(f"Column order in sheet '{sheet_name}' is different between files")
        summary_report.append(f"Column order in sheet '{sheet_name}' is different")

    # Get common columns for value comparison
    common_columns = [col for col in df1.columns if col in df2.columns]

    # Compare row counts
    row_differences = compare_rows(df1, df2, common_columns)
    error_details["row_differences"] = row_differences

    if row_differences["count_diff"]:
        detailed_report.append(f"Row count in sheet '{sheet_name}' is different: {row_differences['count_diff'][0]} rows in file 1 vs {row_differences['count_diff'][1]} rows in file 2")
        summary_report.append(f"Row count in sheet '{sheet_name}' is different: {row_differences['count_diff'][0]} vs {row_differences['count_diff'][1]}")

    if row_differences["missing_rows"]:
        detailed_report.append(f"{len(row_differences['missing_rows'])} rows in sheet '{sheet_name}' are in file 1 but missing in file 2")
        summary_report.append(f"{len(row_differences['missing_rows'])} rows missing in sheet '{sheet_name}'")

    if row_differences["extra_rows"]:
        detailed_report.append(f"{len(row_differences['extra_rows'])} rows in sheet '{sheet_name}' are in file 2 but missing in file 1")
        summary_report.append(f"{len(row_differences['extra_rows'])} extra rows in sheet '{sheet_name}'")

    # Compare values in common rows and columns
    value_differences = compare_values(df1, df2, common_columns, row_differences)
    error_details["value_differences"] = value_differences

    if value_differences:
        detailed_report.append(f"{len(value_differences)} value differences found in sheet '{sheet_name}'")
        summary_report.append(f"{len(value_differences)} value differences in sheet '{sheet_name}'")

        # Add detailed value differences
        for diff in value_differences:
            if "key" in diff:
                detailed_report.append(f"Value difference in sheet '{sheet_name}', key '{diff['key']}', column '{diff['column']}': '{diff['value1']}' vs '{diff['value2']}'")
            else:
                detailed_report.append(f"Value difference in sheet '{sheet_name}', row {diff['row']}, column '{diff['column']}': '{diff['value1']}' vs '{diff['value2']}'")

    return detailed_report, summary_report, error_details

def compare_columns(df1, df2):
    """
    Compare columns between two dataframes
    """
    # Get column names
    cols1 = list(df1.columns)
    cols2 = list(df2.columns)

    # Find missing and extra columns
    missing_cols = [col for col in cols1 if col not in cols2]
    extra_cols = [col for col in cols2 if col not in cols1]

    # Check if column order is different
    common_cols1 = [col for col in cols1 if col in cols2]
    common_cols2 = [col for col in cols2 if col in cols1]

    # Check if the order of common columns is different
    reordered = False
    if common_cols1 != common_cols2:
        # Need to check if the difference is just due to order, not content
        if set(common_cols1) == set(common_cols2):
            reordered = True

    return {
        "missing": missing_cols,
        "extra": extra_cols,
        "reordered": reordered
    }

def compare_rows(df1, df2, common_columns):
    """
    Compare rows between two dataframes
    """
    # Check if there are any common columns to use for comparison
    if not common_columns:
        return {
            "count_diff": [len(df1), len(df2)],
            "missing_rows": {},
            "extra_rows": {}
        }

    # Try to identify a key column (first column or index)
    key_column = common_columns[0]

    # Check if the key column has unique values
    if df1[key_column].duplicated().any() or df2[key_column].duplicated().any():
        # If key column has duplicates, use row indices
        missing_rows = {}
        extra_rows = {}

        # Row count difference
        count_diff = [len(df1), len(df2)]

        return {
            "count_diff": count_diff if count_diff[0] != count_diff[1] else None,
            "missing_rows": missing_rows,
            "extra_rows": extra_rows
        }

    # Use the key column to identify rows
    keys1 = set(df1[key_column].astype(str))
    keys2 = set(df2[key_column].astype(str))

    # Find missing and extra rows
    missing_keys = keys1 - keys2
    extra_keys = keys2 - keys1

    # Create dictionaries with key as the key and value as the row index
    missing_rows = {key: df1[df1[key_column].astype(str) == key].index[0] for key in missing_keys}
    extra_rows = {key: df2[df2[key_column].astype(str) == key].index[0] for key in extra_keys}

    # Row count difference
    count_diff = [len(df1), len(df2)]

    return {
        "count_diff": count_diff if count_diff[0] != count_diff[1] else None,
        "missing_rows": missing_rows,
        "extra_rows": extra_rows
    }

def compare_values(df1, df2, common_columns, row_differences):
    """
    Compare values in common rows and columns
    """
    value_differences = []

    # If there are no common columns, return empty list
    if not common_columns:
        return value_differences

    # Try to identify a key column (first column or index)
    key_column = common_columns[0]

    # Check if the key column has unique values
    if df1[key_column].duplicated().any() or df2[key_column].duplicated().any():
        # If key column has duplicates, use row indices for comparison
        # Get the minimum length to avoid index errors
        min_len = min(len(df1), len(df2))

        for i in range(min_len):
            for col in common_columns:
                # Convert values to strings for comparison to handle different types
                val1 = str(df1.iloc[i][col])
                val2 = str(df2.iloc[i][col])

                # Check if values are different
                if val1 != val2:
                    value_differences.append({
                        "row": i,
                        "column": col,
                        "value1": val1,
                        "value2": val2
                    })
    else:
        # Use the key column to identify common rows
        keys1 = set(df1[key_column].astype(str))
        keys2 = set(df2[key_column].astype(str))
        common_keys = keys1 & keys2

        for key in common_keys:
            # Get the rows with this key
            row1 = df1[df1[key_column].astype(str) == key].iloc[0]
            row2 = df2[df2[key_column].astype(str) == key].iloc[0]

            for col in common_columns:
                # Convert values to strings for comparison to handle different types
                val1 = str(row1[col])
                val2 = str(row2[col])

                # Check if values are different
                if val1 != val2:
                    value_differences.append({
                        "key": key,
                        "column": col,
                        "value1": val1,
                        "value2": val2
                    })

    return value_differences