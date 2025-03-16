import pandas as pd
import numpy as np
from io import BytesIO
import traceback
import streamlit as st
import random
import string
from functools import lru_cache

@lru_cache(maxsize=100)
def generate_random_string(length: int) -> str:
    """Generate a random string with letters and digits."""
    chars = string.ascii_letters + string.digits
    return ''.join(random.choice(chars) for _ in range(length))

def create_sample_files(progress_callback=None):
    """
    Create sample Excel files with comprehensive differences for testing
    
    Args:
        progress_callback: Optional callback function to report progress (0-100)
    
    Returns:
        Tuple of (file1_bytes, file2_bytes)
    """
    try:
        st.info("🔄 Generating sample files...")
        
        # Create base data - simple structure with fewer rows for faster generation
        rows = 5000
        np.random.seed(42)  # For reproducibility
        
        # Initialize progress
        current_progress = 0
        if progress_callback:
            progress_callback(current_progress)
        
        # Create a simple DataFrame with various data types - optimize with numpy arrays
        st.info("📊 Creating base data...")
        data = {
            'ID': np.arange(1, rows + 1),
            'Name': [f'Item_{i}' for i in range(1, rows + 1)],
            'Category': np.random.choice(['A', 'B', 'C', 'D', 'E'], size=rows),
            'Value': np.random.uniform(0, 1000, size=rows),
            'Status': np.random.choice(['Active', 'Inactive', 'Pending'], size=rows),
            'Date': pd.date_range('2023-01-01', periods=rows).astype(str),
            'Amount': np.random.randint(1, 1000, size=rows),
            'Description': [f'Description for item {i}' for i in range(1, rows + 1)]
        }
        
        base_df = pd.DataFrame(data)
        
        # Update progress
        current_progress = 10
        if progress_callback:
            progress_callback(current_progress)
        
        output1 = BytesIO()
        output2 = BytesIO()

        # Generate File 1
        with pd.ExcelWriter(output1, engine='openpyxl') as writer1:
            # Sheet 1: Base sheet (identical in both files)
            st.info("📊 Generating Sheet1 (identical)...")
            base_df.to_excel(writer1, sheet_name='Sheet1', index=False)
            
            # Update progress
            current_progress = 20
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 2: Same in both but with some value differences
            st.info("📊 Generating Sheet2 (value differences)...")
            df2 = base_df.copy()
            # Modify more values with significant differences
            random_indices = np.random.choice(len(df2), size=1000, replace=False)  # Increased from 500 to 1000
            for idx in random_indices:
                # Make more significant changes to values
                if idx % 3 == 0:
                    df2.loc[idx, 'Value'] = df2.loc[idx, 'Value'] * 2.0  # Double the value
                    df2.loc[idx, 'Status'] = 'Significantly Modified'
                elif idx % 3 == 1:
                    df2.loc[idx, 'Value'] = df2.loc[idx, 'Value'] * 0.5  # Half the value
                    df2.loc[idx, 'Status'] = 'Reduced'
                else:
                    df2.loc[idx, 'Value'] = df2.loc[idx, 'Value'] + 100  # Add 100
                    df2.loc[idx, 'Status'] = 'Increased'
                
                # Also modify text fields for more obvious differences
                df2.loc[idx, 'Description'] = f'CHANGED: {df2.loc[idx, "Description"]}'
                
                # Modify dates occasionally
                if idx % 5 == 0:
                    df2.loc[idx, 'Date'] = '2024-01-01'  # Fixed different date
                
                # Modify amounts
                df2.loc[idx, 'Amount'] = df2.loc[idx, 'Amount'] + 500
            df2.to_excel(writer1, sheet_name='Sheet2', index=False)
            
            # Update progress
            current_progress = 30
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 3: Column order differences
            st.info("📊 Generating Sheet3 (column order differences)...")
            df3 = base_df.copy()
            # Shuffle columns
            columns = list(df3.columns)
            random.shuffle(columns)
            df3 = df3[columns]
            df3.to_excel(writer1, sheet_name='Sheet3', index=False)
            
            # Update progress
            current_progress = 40
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 4: Column name differences
            st.info("📊 Generating Sheet4 (column name differences)...")
            df4 = base_df.copy()
            # Rename some columns
            df4 = df4.rename(columns={
                'Value': 'Price',
                'Status': 'State',
                'Description': 'Details'
            })
            df4.to_excel(writer1, sheet_name='Sheet4', index=False)
            
            # Update progress
            current_progress = 50
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 5: Missing columns
            st.info("📊 Generating Sheet5 (missing columns)...")
            df5 = base_df.drop(['Description', 'Status'], axis=1)
            df5.to_excel(writer1, sheet_name='Sheet5', index=False)
            
            # Sheet 6: Unique to File 1
            st.info("📊 Generating Sheet6 (unique to File 1)...")
            df6 = base_df.head(1000).copy()
            df6['File1_Only'] = 'This column only exists in File 1'
            df6.to_excel(writer1, sheet_name='Sheet6', index=False)
            
            # Update progress
            current_progress = 60
            if progress_callback:
                progress_callback(current_progress)

        # Generate File 2
        with pd.ExcelWriter(output2, engine='openpyxl') as writer2:
            # Sheet 1: Identical to File 1
            base_df.to_excel(writer2, sheet_name='Sheet1', index=False)
            
            # Update progress
            current_progress = 70
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 2: Same structure but different values
            df2_2 = base_df.copy()
            # Apply matching modifications to file 2 for proper comparison
            for idx in random_indices:
                # Make corresponding changes to file 2 with different values
                if idx % 3 == 0:
                    # Original value was doubled, here we'll triple it for a clear difference
                    df2_2.loc[idx, 'Value'] = df2_2.loc[idx, 'Value'] * 3.0
                    df2_2.loc[idx, 'Status'] = 'Extremely Modified'
                elif idx % 3 == 1:
                    # Original value was halved, here we'll quarter it
                    df2_2.loc[idx, 'Value'] = df2_2.loc[idx, 'Value'] * 0.25
                    df2_2.loc[idx, 'Status'] = 'Severely Reduced'
                else:
                    # Original value had 100 added, here we'll add 200
                    df2_2.loc[idx, 'Value'] = df2_2.loc[idx, 'Value'] + 200
                    df2_2.loc[idx, 'Status'] = 'Greatly Increased'
                
                # Different text modification
                df2_2.loc[idx, 'Description'] = f'MODIFIED: {df2_2.loc[idx, "Description"]}'
                
                # Different date modification
                if idx % 5 == 0:
                    df2_2.loc[idx, 'Date'] = '2025-01-01'  # Different year
                
                # Different amount modification
                df2_2.loc[idx, 'Amount'] = df2_2.loc[idx, 'Amount'] + 1000
            df2_2.to_excel(writer2, sheet_name='Sheet2', index=False)
            
            # Update progress
            current_progress = 80
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 3: Different column order than File 1
            df3_2 = base_df.copy()
            # Reverse column order
            df3_2 = df3_2[df3_2.columns[::-1]]
            df3_2.to_excel(writer2, sheet_name='Sheet3', index=False)
            
            # Sheet 4: Different column names
            df4_2 = base_df.copy()
            # Different renaming
            df4_2 = df4_2.rename(columns={
                'Value': 'Cost',
                'Status': 'Condition',
                'Description': 'Notes'
            })
            df4_2.to_excel(writer2, sheet_name='Sheet4', index=False)
            
            # Update progress
            current_progress = 90
            if progress_callback:
                progress_callback(current_progress)
            
            # Sheet 5: Extra columns
            df5_2 = base_df.copy()
            df5_2['Extra1'] = np.random.rand(len(df5_2))
            df5_2['Extra2'] = [f'Extra_{i}' for i in range(len(df5_2))]
            df5_2.to_excel(writer2, sheet_name='Sheet5', index=False)
            
            # Sheet 7: Unique to File 2 (note: different sheet number)
            df7 = base_df.head(1000).copy()
            df7['File2_Only'] = 'This column only exists in File 2'
            df7.to_excel(writer2, sheet_name='Sheet7', index=False)
            
            # Sheet with special characters in name
            df_special = base_df.head(500).copy()
            df_special.to_excel(writer2, sheet_name='Special Sheet #1!', index=False)

        # Final progress update
        if progress_callback:
            progress_callback(100)
            
        st.success("✅ Sample files generated successfully!")
        
        # Return the files as bytes
        output1.seek(0)
        output2.seek(0)
        return output1.getvalue(), output2.getvalue()

    except Exception as e:
        error_msg = f"❌ Error generating sample files: {str(e)}\n{traceback.format_exc()}"
        st.error(error_msg)
        print(error_msg)
        
        # Return empty files in case of error
        empty1, empty2 = BytesIO(), BytesIO()
        with pd.ExcelWriter(empty1, engine='openpyxl') as writer:
            pd.DataFrame({"Error": ["Sample generation failed"]}).to_excel(writer, index=False)
        with pd.ExcelWriter(empty2, engine='openpyxl') as writer:
            pd.DataFrame({"Error": ["Sample generation failed"]}).to_excel(writer, index=False)
        empty1.seek(0)
        empty2.seek(0)
        return empty1.getvalue(), empty2.getvalue()