import pandas as pd
import numpy as np
import io
import os
import streamlit as st
import concurrent.futures
from functools import lru_cache

def read_file(file):
    """
    Read a file and return its data with optimized memory usage
    """
    # Get file extension
    file_extension = os.path.splitext(file.name)[1].lower()
    file_size_mb = file.size / (1024 * 1024)  # Convert to MB

    # Initialize result dictionary
    result = {
        "name": file.name,
        "type": None,
        "data": None,
        "sheet_names": []
    }

    st.info(f"📂 Processing {file.name} ({file_size_mb:.1f} MB)")

    # Read Excel file
    if file_extension in ['.xlsx', '.xls']:
        result["type"] = "excel"
        
        # Create a BytesIO object
        excel_data = io.BytesIO(file.read())
        
        # Use pandas ExcelFile for better memory management
        with pd.ExcelFile(excel_data) as xls:
            # Convert sheet names to strings to ensure they're hashable
            result["sheet_names"] = [str(sheet) for sheet in xls.sheet_names]
            st.info(f"📊 Found {len(result['sheet_names'])} sheets")

            # Read each sheet with optimized settings
            sheets_data = {}
            total_sheets = len(result["sheet_names"])
            
            # Use parallel processing for multiple sheets
            if total_sheets > 1 and file_size_mb > 10:  # Only use parallel for larger files with multiple sheets
                with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, total_sheets)) as executor:
                    # Create a dictionary to store futures - ensure sheet_name is a string
                    future_to_sheet = {
                        executor.submit(
                            read_excel_sheet, xls, str(sheet_name), idx, total_sheets
                        ): str(sheet_name) for idx, sheet_name in enumerate(result["sheet_names"], 1)
                    }
                    
                    # Process results as they complete
                    for future in concurrent.futures.as_completed(future_to_sheet):
                        sheet_name = future_to_sheet[future]
                        try:
                            df = future.result()
                            # Use string representation of sheet name as dictionary key
                            sheets_data[sheet_name] = df
                        except Exception as e:
                            st.error(f"❌ Error reading sheet {sheet_name}: {str(e)}")
                            sheets_data[sheet_name] = pd.DataFrame()  # Empty DataFrame for failed sheets
            else:
                # Sequential processing for smaller files
                for idx, sheet_name in enumerate(result["sheet_names"], 1):
                    st.info(f"📑 Reading sheet {idx}/{total_sheets}: {sheet_name}")
                    
                    try:
                        df = read_excel_sheet(xls, sheet_name, idx, total_sheets)
                        # Use string representation of sheet name as dictionary key
                        sheets_data[sheet_name] = df
                    except Exception as e:
                        st.error(f"❌ Error reading sheet {sheet_name}: {str(e)}")
                        sheets_data[sheet_name] = pd.DataFrame()  # Empty DataFrame for failed sheets

            result["data"] = sheets_data

    # Read CSV file
    elif file_extension == '.csv':
        result["type"] = "csv"
        st.info("📊 Reading CSV file")

        try:
            # Determine optimal chunk size based on file size
            chunk_size = max(100000, min(1000000, int(file_size_mb * 10000)))  # Dynamic chunk size
            chunks = []
            
            # Create a StringIO object
            csv_data = io.BytesIO(file.read())
            
            # Try to detect encoding
            try:
                import chardet
                raw_data = csv_data.read(min(1000000, file.size))  # Read first MB or less
                encoding = chardet.detect(raw_data)['encoding']
                csv_data.seek(0)  # Reset position
            except:
                encoding = 'utf-8'  # Default to UTF-8
            
            # Try to detect delimiter
            try:
                import csv
                dialect = csv.Sniffer().sniff(csv_data.read(min(10000, file.size)).decode(encoding))
                delimiter = dialect.delimiter
                csv_data.seek(0)  # Reset position
            except:
                delimiter = ','  # Default to comma
            
            # Read chunks with optimized settings
            for chunk_num, chunk in enumerate(pd.read_csv(
                csv_data,
                chunksize=chunk_size,
                dtype=None,  # Let pandas infer types
                na_filter=True,  # Handle missing values efficiently
                low_memory=True,
                encoding=encoding,
                delimiter=delimiter,
                engine='c'  # Use the faster C engine
            ), 1):
                # Optimize each chunk
                chunk = optimize_dataframe(chunk)
                chunks.append(chunk)
                st.info(f"📑 Read chunk {chunk_num} ({len(chunk)} rows)")
            
            # Combine chunks
            if len(chunks) == 1:
                df = chunks[0]  # No need to concatenate if only one chunk
            else:
                df = pd.concat(chunks, ignore_index=True)
            
            result["data"] = df
            st.info(f"✅ Successfully read CSV ({len(df)} rows, {len(df.columns)} columns)")
            
        except Exception as e:
            st.error(f"❌ Error reading CSV: {str(e)}")
            result["data"] = pd.DataFrame()

    return result

def read_excel_sheet(xls, sheet_name, idx, total_sheets):
    """Helper function to read an Excel sheet with optimized settings"""
    st.info(f"📑 Reading sheet {idx}/{total_sheets}: {sheet_name}")
    
    # Try to read with optimized settings
    try:
        # Convert sheet_name to string to ensure it's hashable
        sheet_name_str = str(sheet_name)
        
        # Try to infer data types automatically
        df = pd.read_excel(
            xls,
            sheet_name=sheet_name_str,
            dtype=None,  # Let pandas infer types
            engine='openpyxl',
            na_filter=True  # Handle missing values efficiently
        )
        
        # Optimize memory usage
        df = optimize_dataframe(df)
        
        st.info(f"✅ Successfully read {sheet_name} ({len(df)} rows, {len(df.columns)} columns)")
        return df
    
    except Exception as e:
        st.error(f"❌ Error reading sheet {sheet_name}: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        raise

def optimize_dataframe(df):
    """
    Optimize DataFrame memory usage by selecting appropriate data types
    """
    # Make a copy to avoid modifying the original
    df = df.copy()
    
    # Handle missing values first
    df = df.fillna(pd.NA)  # Use pandas NA for consistent handling
    
    # Optimize numeric columns
    for col in df.select_dtypes(include=['int', 'float']).columns:
        # Skip if column has NaN values (can't be optimized to integer)
        if df[col].isna().any():
            if df[col].dtype == 'float':
                df[col] = pd.to_numeric(df[col], downcast='float')
            continue
            
        # Convert to numeric with best type
        if df[col].dtype == 'int':
            col_min = df[col].min()
            col_max = df[col].max()
            
            # Choose the smallest possible integer type
            if col_min >= 0:
                if col_max < 255:
                    df[col] = df[col].astype(np.uint8)
                elif col_max < 65535:
                    df[col] = df[col].astype(np.uint16)
                elif col_max < 4294967295:
                    df[col] = df[col].astype(np.uint32)
                else:
                    df[col] = df[col].astype(np.uint64)
            else:
                if col_min > -128 and col_max < 127:
                    df[col] = df[col].astype(np.int8)
                elif col_min > -32768 and col_max < 32767:
                    df[col] = df[col].astype(np.int16)
                elif col_min > -2147483648 and col_max < 2147483647:
                    df[col] = df[col].astype(np.int32)
                else:
                    df[col] = df[col].astype(np.int64)
        elif df[col].dtype == 'float':
            df[col] = pd.to_numeric(df[col], downcast='float')

    # Optimize string columns - only if there are a reasonable number of rows
    if len(df) < 1000000:  # Skip for very large dataframes
        for col in df.select_dtypes(include=['object']).columns:
            # Check if column contains only strings
            if df[col].dropna().map(type).eq(str).all():
                # Convert to category if less than 50% unique values and at least 10 rows
                if len(df) >= 10:
                    num_unique = df[col].nunique()
                    if num_unique / len(df) < 0.5 and num_unique < 1000:
                        df[col] = df[col].astype('category')

    return df