import pandas as pd
import numpy as np
import streamlit as st
import time
import openpyxl
from openpyxl.styles import PatternFill, Font
from io import BytesIO
from functools import lru_cache

def compare_files(data1, data2):
    """
    Compare two files and return detailed report, summary report, and error details
    """
    start_time = time.time()
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
        return detailed_report, summary_report, error_details

    # Compare sheet names (for Excel files)
    if data1["type"] == "excel" and data2["type"] == "excel":
        # Use sets for faster comparison
        sheets1 = set(data1["sheet_names"])
        sheets2 = set(data2["sheet_names"])
        
        missing_sheets = sheets1 - sheets2
        extra_sheets = sheets2 - sheets1
        common_sheets = sheets1 & sheets2

        if missing_sheets:
            st.warning(f"⚠️ Found {len(missing_sheets)} missing sheets: {', '.join(missing_sheets)}")
            error_details["missing_sheets"] = list(missing_sheets)
            detailed_report.extend(f"Sheet '{sheet}' is in file 1 but missing in file 2" for sheet in missing_sheets)
            summary_report.extend(f"Sheet '{sheet}' is missing in file 2" for sheet in missing_sheets)

        if extra_sheets:
            st.warning(f"⚠️ Found {len(extra_sheets)} extra sheets: {', '.join(extra_sheets)}")
            error_details["extra_sheets"] = list(extra_sheets)
            detailed_report.extend(f"Sheet '{sheet}' is in file 2 but missing in file 1" for sheet in extra_sheets)
            summary_report.extend(f"Extra sheet '{sheet}' in file 2" for sheet in extra_sheets)
        
        # Process sheets in parallel using pandas
        import concurrent.futures
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(common_sheets))) as executor:
            # Create a dictionary to store futures
            future_to_sheet = {}
            
            # Submit tasks for each sheet
            for sheet in common_sheets:
                sheet_str = str(sheet)
                try:
                    # Check if the sheet exists in both data dictionaries
                    if sheet_str in data1["data"] and sheet_str in data2["data"]:
                        future = executor.submit(
                            compare_sheets, data1["data"][sheet_str], data2["data"][sheet_str]
                        )
                        future_to_sheet[future] = sheet_str
                    else:
                        st.error(f"❌ Sheet '{sheet}' exists in sheet_names but not in data dictionary")
                except Exception as e:
                    st.error(f"❌ Error submitting sheet '{sheet}' for comparison: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
            
            # Process results as they complete
            for future in concurrent.futures.as_completed(future_to_sheet):
                sheet = future_to_sheet[future]
                try:
                    sheet_detailed_report, sheet_summary_report, sheet_error_details = future.result()
                    
                    if any(sheet_error_details.values()):
                        detailed_report.extend(sheet_detailed_report)
                        summary_report.extend(sheet_summary_report)
                        
                        # Update error details only if there are differences
                        if sheet_error_details["column_differences"]:
                            error_details["column_differences"][sheet] = sheet_error_details["column_differences"]
                        if sheet_error_details["row_differences"]:
                            error_details["row_differences"][sheet] = sheet_error_details["row_differences"]
                        if sheet_error_details["value_differences"]:
                            error_details["value_differences"][sheet] = sheet_error_details["value_differences"]
                
                except Exception as e:
                    st.error(f"❌ Error analyzing sheet '{sheet}': {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

    # Compare CSV files
    elif data1["type"] == "csv" and data2["type"] == "csv":
        sheet_detailed_report, sheet_summary_report, sheet_error_details = compare_sheets(
            data1["data"], data2["data"]
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

    end_time = time.time()
    st.info(f"Comparison completed in {end_time - start_time:.2f} seconds")
    return detailed_report, summary_report, error_details

def compare_sheets(df1, df2, key_columns=None, chunk_size=10000):
    """
    Compare two dataframes and return a detailed report, summary report, and error details.
    
    Args:
        df1: First dataframe
        df2: Second dataframe
        key_columns: List of columns to use as keys for matching rows
        chunk_size: Size of chunks to process at a time
        
    Returns:
        detailed_report: Detailed report of differences
        summary_report: Summary report of differences
        error_details: Dictionary with details of errors
    """
    import pandas as pd
    import numpy as np
    import streamlit as st
    from datetime import datetime
    
    # Initialize reports and error details
    detailed_report = []
    summary_report = []
    error_details = {
        "column_differences": {
            "missing_columns": [],
            "extra_columns": [],
            "reordered_columns": []
        },
        "row_differences": {
            "count_diff": 0,
            "missing_rows": [],
            "extra_rows": []
        },
        "value_differences": []
    }
    
    # Quick check if dataframes are identical
    if df1.equals(df2):
        st.info("DataFrames are identical - no differences found")
        return detailed_report, summary_report, error_details
    
    # Check for column differences
    df1_cols = set(df1.columns)
    df2_cols = set(df2.columns)
    
    # Find missing and extra columns
    missing_cols = df1_cols - df2_cols
    extra_cols = df2_cols - df1_cols
    common_cols = df1_cols.intersection(df2_cols)
    
    # Log column differences
    if missing_cols:
        error_details["column_differences"]["missing_columns"] = list(missing_cols)
        detailed_report.append(f"Missing columns in second file: {', '.join(missing_cols)}")
        summary_report.append(f"Missing columns: {len(missing_cols)}")
        st.warning(f"Missing columns in second file: {', '.join(missing_cols)}")
    
    if extra_cols:
        error_details["column_differences"]["extra_columns"] = list(extra_cols)
        detailed_report.append(f"Extra columns in second file: {', '.join(extra_cols)}")
        summary_report.append(f"Extra columns: {len(extra_cols)}")
        st.warning(f"Extra columns in second file: {', '.join(extra_cols)}")
    
    # Check for reordered columns
    if list(df1.columns) != list(df2.columns):
        reordered_cols = [col for col in df1.columns if col in common_cols and list(df1.columns).index(col) != list(df2.columns).index(col)]
        if reordered_cols:
            error_details["column_differences"]["reordered_columns"] = reordered_cols
            detailed_report.append(f"Reordered columns: {', '.join(reordered_cols)}")
            summary_report.append(f"Reordered columns: {len(reordered_cols)}")
            st.warning(f"Reordered columns: {', '.join(reordered_cols)}")
    
    # Check row count differences
    row_diff = len(df1) - len(df2)
    if row_diff != 0:
        error_details["row_differences"]["count_diff"] = row_diff
        detailed_report.append(f"Row count difference: {row_diff} ({len(df1)} vs {len(df2)})")
        summary_report.append(f"Row count difference: {abs(row_diff)}")
        st.warning(f"Row count difference: {row_diff} ({len(df1)} vs {len(df2)})")
    
    # Compare values for common columns
    value_diffs = []
    
    # If key columns are provided, use them for matching rows
    if key_columns:
        # Log key columns being used
        st.info(f"Using key columns for matching rows: {key_columns}")
        
        # Ensure all key columns exist in both dataframes
        if all(col in df1.columns for col in key_columns) and all(col in df2.columns for col in key_columns):
            # Create dictionaries for faster lookups
            st.info("Creating dictionaries for key-based comparison")
            
            # Convert df1 to dictionary for faster lookups
            df1_dict = {}
            for _, row in df1.iterrows():
                # Create a tuple of key values
                key_values = tuple(str(row[col]) for col in key_columns)
                df1_dict[key_values] = row
            
            # Convert df2 to dictionary for faster lookups
            df2_dict = {}
            for _, row in df2.iterrows():
                # Create a tuple of key values
                key_values = tuple(str(row[col]) for col in key_columns)
                df2_dict[key_values] = row
            
            # Find common keys
            common_keys = set(df1_dict.keys()).intersection(set(df2_dict.keys()))
            missing_keys = set(df1_dict.keys()) - set(df2_dict.keys())
            extra_keys = set(df2_dict.keys()) - set(df1_dict.keys())
            
            st.info(f"Found {len(common_keys)} common keys, {len(missing_keys)} missing keys, {len(extra_keys)} extra keys")
            
            # Track missing and extra rows
            if missing_keys:
                error_details["row_differences"]["missing_rows"] = [dict(zip(key_columns, key)) for key in list(missing_keys)[:100]]  # Limit to 100 for performance
                detailed_report.append(f"Missing rows in second file: {len(missing_keys)}")
                summary_report.append(f"Missing rows: {len(missing_keys)}")
            
            if extra_keys:
                error_details["row_differences"]["extra_rows"] = [dict(zip(key_columns, key)) for key in list(extra_keys)[:100]]  # Limit to 100 for performance
                detailed_report.append(f"Extra rows in second file: {len(extra_keys)}")
                summary_report.append(f"Extra rows: {len(extra_keys)}")
            
            # Compare values for common keys and columns
            st.info(f"Comparing values for {len(common_keys)} common keys and {len(common_cols)} common columns")
            
            # Process in chunks to avoid memory issues
            keys_list = list(common_keys)
            total_diffs = 0
            
            for i in range(0, len(keys_list), chunk_size):
                chunk_keys = keys_list[i:i+chunk_size]
                chunk_diffs = []
                
                for key in chunk_keys:
                    row1 = df1_dict[key]
                    row2 = df2_dict[key]
                    
                    for col in common_cols:
                        # Convert values to strings for comparison to handle NaN and different types
                        val1 = str(row1[col]) if pd.notna(row1[col]) else "NaN"
                        val2 = str(row2[col]) if pd.notna(row2[col]) else "NaN"
                        
                        # Compare values
                        if val1 != val2:
                            diff = {
                                "key": dict(zip(key_columns, key)),
                                "column": col,
                                "value1": val1,
                                "value2": val2
                            }
                            chunk_diffs.append(diff)
                            
                            # Log for debugging (limit to first 10)
                            if total_diffs < 10:
                                st.info(f"Found difference: key={key}, col={col}, val1={val1}, val2={val2}")
                
                # Add chunk differences to total
                value_diffs.extend(chunk_diffs)
                total_diffs += len(chunk_diffs)
                
                # Log progress
                st.info(f"Processed chunk {i//chunk_size + 1}/{(len(keys_list) + chunk_size - 1)//chunk_size}, found {len(chunk_diffs)} differences")
            
            if value_diffs:
                error_details["value_differences"] = value_diffs
                detailed_report.append(f"Value differences: {len(value_diffs)}")
                summary_report.append(f"Value differences: {len(value_diffs)}")
                st.warning(f"Found {len(value_diffs)} value differences")
        else:
            missing_key_cols = [col for col in key_columns if col not in df1.columns or col not in df2.columns]
            st.error(f"Key columns not found in both dataframes: {missing_key_cols}")
            detailed_report.append(f"Key columns not found in both dataframes: {missing_key_cols}")
            summary_report.append("Unable to compare rows: key columns missing")
    else:
        # If no key columns, compare row by row
        st.info("No key columns provided, comparing row by row")
        
        # Get the minimum number of rows to compare
        min_rows = min(len(df1), len(df2))
        
        # Process in chunks to avoid memory issues
        for i in range(0, min_rows, chunk_size):
            chunk_end = min(i + chunk_size, min_rows)
            chunk_diffs = []
            
            try:
                # Get chunks of both dataframes
                df1_chunk = df1.iloc[i:chunk_end]
                df2_chunk = df2.iloc[i:chunk_end]
                
                # Compare values for common columns
                for col in common_cols:
                    # Convert values to strings for comparison
                    s1 = df1_chunk[col].astype(str)
                    s2 = df2_chunk[col].astype(str)
                    
                    # Find differences
                    diff_mask = (s1 != s2)
                    diff_indices = diff_mask[diff_mask].index
                    
                    for idx in diff_indices:
                        row_idx = df1.index.get_loc(idx)
                        val1 = str(df1.loc[idx, col])
                        val2 = str(df2.loc[idx, col])
                        
                        diff = {
                            "row": row_idx,
                            "column": col,
                            "value1": val1,
                            "value2": val2
                        }
                        chunk_diffs.append(diff)
                        
                        # Log for debugging (limit to first 10)
                        if len(value_diffs) + len(chunk_diffs) <= 10:
                            st.info(f"Found difference: row={row_idx}, col={col}, val1={val1}, val2={val2}")
                
                # Add chunk differences to total
                value_diffs.extend(chunk_diffs)
                
                # Log progress
                st.info(f"Processed chunk {i//chunk_size + 1}/{(min_rows + chunk_size - 1)//chunk_size}, found {len(chunk_diffs)} differences")
            
            except Exception as e:
                st.error(f"Error comparing chunk {i}-{chunk_end}: {str(e)}")
                detailed_report.append(f"Error comparing rows {i}-{chunk_end}: {str(e)}")
                summary_report.append("Error during row comparison")
        
        if value_diffs:
            error_details["value_differences"] = value_diffs
            detailed_report.append(f"Value differences: {len(value_diffs)}")
            summary_report.append(f"Value differences: {len(value_diffs)}")
            st.warning(f"Found {len(value_diffs)} value differences")
    
    return detailed_report, summary_report, error_details