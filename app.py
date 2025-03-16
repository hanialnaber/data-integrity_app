import streamlit as st
import pandas as pd
import numpy as np
import os
import time
import gc
from src.file_handler import read_file
from src.comparison import compare_files
from src.highlighting import highlight_differences_excel, highlight_differences_csv
from src.sample_generator import create_sample_files
import io
import tempfile

# Set max upload size to 2GB (Streamlit's absolute maximum)
os.environ['STREAMLIT_SERVER_MAX_UPLOAD_SIZE'] = "2048"

# Import modules from src
from src.ui import (
    render_header, render_file_upload_section,
    render_comparison_results, render_download_section
)

# Configure Streamlit for better performance
st.set_page_config(
    page_title="Data Integrity Checker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize session state variables if they don't exist
if 'comparison_complete' not in st.session_state:
    st.session_state.comparison_complete = False
if 'detailed_report' not in st.session_state:
    st.session_state.detailed_report = None
if 'summary_report' not in st.session_state:
    st.session_state.summary_report = None
if 'error_details' not in st.session_state:
    st.session_state.error_details = None
if 'highlighted_excel' not in st.session_state:
    st.session_state.highlighted_excel = None

def main():
    """Main application function"""
    # Render header
    render_header()

    # Add tabs for different functionalities
    tab1, tab2 = st.tabs(["Compare Files", "Generate Sample Files"])

    with tab1:
        st.header("Upload Files")
        
        # Use columns for better layout
        col1, col2 = st.columns(2)
        
        with col1:
            file1 = st.file_uploader("Choose first file", type=['xlsx', 'xls', 'csv'])
        
        with col2:
            file2 = st.file_uploader("Choose second file", type=['xlsx', 'xls', 'csv'])

        if file1 is not None and file2 is not None:
            # Add a progress bar container
            progress_container = st.empty()
            
            # Add a compare button
            if st.button("Compare Files"):
                # Reset session state
                st.session_state.comparison_complete = False
                st.session_state.detailed_report = None
                st.session_state.summary_report = None
                st.session_state.error_details = None
                st.session_state.highlighted_excel = None
                
                # Create a progress bar
                progress_bar = progress_container.progress(0)
                
                # Read files
                with st.spinner("Reading files..."):
                    start_time = time.time()
                    data1 = read_file(file1)
                    progress_bar.progress(25)
                    data2 = read_file(file2)
                    progress_bar.progress(50)
                    read_time = time.time() - start_time
                    st.info(f"Files read in {read_time:.2f} seconds")

                # Compare files
                with st.spinner("Comparing files..."):
                    start_time = time.time()
                    detailed_report, summary_report, error_details = compare_files(data1, data2)
                    progress_bar.progress(75)
                    compare_time = time.time() - start_time
                    st.info(f"Comparison completed in {compare_time:.2f} seconds")
                
                # Store results in session state
                st.session_state.detailed_report = detailed_report
                st.session_state.summary_report = summary_report
                st.session_state.error_details = error_details
                
                # Create highlighted Excel file
                if data1["type"] == data2["type"]:
                    try:
                        st.info("Creating highlighted Excel file...")
                        
                        # Create a temporary file path
                        temp_output_path = os.path.join(tempfile.gettempdir(), "highlighted_comparison.xlsx")
                        
                        # Save uploaded files to temporary files
                        temp_file1_path = os.path.join(tempfile.gettempdir(), "temp_file1")
                        temp_file2_path = os.path.join(tempfile.gettempdir(), "temp_file2")
                        
                        with open(temp_file1_path, "wb") as f:
                            f.write(file1.getvalue())
                        
                        with open(temp_file2_path, "wb") as f:
                            f.write(file2.getvalue())
                        
                        if data1["type"] == "excel" and data2["type"] == "excel":
                            # Use the Excel highlighting function
                            highlighted_file = highlight_differences_excel(
                                file1_path=temp_file1_path,
                                file2_path=temp_file2_path,
                                output_path=temp_output_path,
                                error_details=error_details
                            )
                        elif data1["type"] == "csv" and data2["type"] == "csv":
                            # Use the CSV highlighting function
                            highlighted_file = highlight_differences_csv(
                                file1_path=temp_file1_path,
                                file2_path=temp_file2_path,
                                output_path=temp_output_path,
                                error_details=error_details
                            )
                        else:
                            st.warning("Highlighting is only available for Excel-Excel or CSV-CSV comparisons.")
                            highlighted_file = None
                        
                        if highlighted_file:
                            # Read the file into memory
                            with open(temp_output_path, "rb") as f:
                                highlighted_bytes = f.read()
                            
                            # Offer the file for download
                            st.download_button(
                                label="Download Highlighted Excel Report",
                                data=highlighted_bytes,
                                file_name="highlighted_comparison.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            # Clean up the temporary files
                            try:
                                os.remove(temp_output_path)
                                os.remove(temp_file1_path)
                                os.remove(temp_file2_path)
                            except:
                                pass
                    except Exception as e:
                        st.error(f"Error creating highlighted Excel report: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())
                else:
                    st.warning("Highlighting is only available for Excel-Excel or CSV-CSV comparisons.")
                
                # Mark comparison as complete
                st.session_state.comparison_complete = True
                progress_bar.progress(100)
                
                # Clear progress bar after completion
                time.sleep(0.5)
                progress_container.empty()
                
                # Force garbage collection to free memory
                gc.collect()
            
            # Display results if comparison is complete
            if st.session_state.comparison_complete:
                display_comparison_results()

    with tab2:
        st.header("Generate Sample Files")
        st.write("Generate sample Excel files with known differences for testing.")
        
        # Add a progress bar container
        sample_progress_container = st.empty()
        
        generate_button = st.button("Generate Sample Files")
        
        if generate_button:
            # Create a progress bar
            sample_progress_bar = sample_progress_container.progress(0)
            
            try:
                with st.spinner("Generating sample files... This may take a few minutes..."):
                    start_time = time.time()
                    file1_bytes, file2_bytes = create_sample_files(
                        progress_callback=lambda p: sample_progress_bar.progress(p)
                    )
                    generation_time = time.time() - start_time
                    st.info(f"Sample files generated in {generation_time:.2f} seconds")
                
                # Store the generated files in session state to prevent regeneration on download
                st.session_state.file1_bytes = file1_bytes
                st.session_state.file2_bytes = file2_bytes
                
                # Clear progress bar after completion
                time.sleep(0.5)
                sample_progress_container.empty()
                
                st.success("Sample files generated successfully! Click the buttons below to download.")
                
                # Display information about the generated files
                st.info("""
                📊 Sample Files Content:
                
                The files contain various types of differences:
                
                - Sheet1: Identical in both files (baseline)
                - Sheet2: Value differences with controlled error rate
                - Sheet3: Structural differences (column order, names)
                - Sheet4: Column name differences
                - Sheet5: Missing/extra columns
                - Sheet6: Unique to File 1
                - Sheet7: Unique to File 2
                - Special Sheet #1!: Sheet with special characters (only in File 2)
                """)
                
            except Exception as e:
                st.error(f"Error generating sample files: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
                
                # Clear progress bar on error
                sample_progress_container.empty()
        
        # Create download buttons outside the generation button to avoid regeneration
        if 'file1_bytes' in st.session_state and 'file2_bytes' in st.session_state:
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="Download Sample File 1",
                    data=st.session_state.file1_bytes,
                    file_name="sample1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_sample1"
                )
            
            with col2:
                st.download_button(
                    label="Download Sample File 2",
                    data=st.session_state.file2_bytes,
                    file_name="sample2.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_sample2"
                )

def display_comparison_results():
    """Display comparison results from session state"""
    if st.session_state.summary_report:
        st.header("Summary of Differences")
        for item in st.session_state.summary_report:
            st.warning(item)

        # Create downloadable detailed report
        if st.session_state.detailed_report:
            col1, col2 = st.columns(2)
            with col1:
                detailed_text = "\n".join(st.session_state.detailed_report)
                st.download_button(
                    label="Download Detailed Report",
                    data=detailed_text,
                    file_name="comparison_report.txt",
                    mime="text/plain"
                )
            
            with col2:
                if st.session_state.highlighted_excel:
                    st.download_button(
                        label="Download Highlighted Excel Report",
                        data=st.session_state.highlighted_excel,
                        file_name="highlighted_comparison.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Add color code explanation
                    st.markdown("### Color Code Legend")
                    legend_data = {
                        "Color": ["🟥 Red", "🟨 Yellow", "🟩 Green", "🟦 Blue"],
                        "Meaning": [
                            "Missing in File 2 (only in File 1)", 
                            "Value Difference", 
                            "Extra in File 2 (not in File 1)",
                            "Order Mismatch (columns or rows)"
                        ]
                    }
                    st.table(pd.DataFrame(legend_data))
                else:
                    st.warning("Highlighting only supported for Excel-Excel or CSV-CSV comparisons")

        # Display error details in an organized way
        if any(st.session_state.error_details.values()):
            # Use expander to save space
            with st.expander("Detailed Error Analysis", expanded=False):
                
                if st.session_state.error_details["missing_sheets"]:
                    st.subheader("Missing Sheets")
                    st.write(", ".join(st.session_state.error_details["missing_sheets"]))
                
                if st.session_state.error_details["extra_sheets"]:
                    st.subheader("Extra Sheets")
                    st.write(", ".join(st.session_state.error_details["extra_sheets"]))
                
                for sheet, details in st.session_state.error_details.get("column_differences", {}).items():
                    if details:
                        st.subheader(f"Column Differences in {sheet}")
                        if details.get("missing"):
                            st.write("Missing columns:", ", ".join(details["missing"]))
                        if details.get("extra"):
                            st.write("Extra columns:", ", ".join(details["extra"]))
                        if details.get("reordered"):
                            st.write("⚠️ Columns are in different order")
                
                if st.session_state.error_details.get("value_differences"):
                    st.subheader("Value Differences")
                    
                    # Limit the number of differences to display
                    all_diffs = []
                    for sheet, diffs in st.session_state.error_details["value_differences"].items():
                        for diff in diffs:
                            diff['sheet'] = sheet
                            all_diffs.append(diff)
                    
                    if all_diffs:
                        # Create a DataFrame with all differences
                        df = pd.DataFrame(all_diffs)
                        
                        # Limit to 1000 rows for performance
                        if len(df) > 1000:
                            st.warning(f"Showing only 1000 of {len(df)} value differences")
                            df = df.head(1000)
                        
                        # Display as an interactive table
                        st.dataframe(df)
    else:
        st.success("No differences found! The files are identical.")

if __name__ == "__main__":
    main()