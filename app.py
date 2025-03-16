import streamlit as st
import pandas as pd
import os

# Set max upload size to 2GB (Streamlit's absolute maximum)
os.environ['STREAMLIT_SERVER_MAX_UPLOAD_SIZE'] = "2048"

# Import modules from src
from src.ui import (
    setup_page, render_header, render_file_upload_section,
    render_comparison_results, render_download_section
)
from src.file_handler import read_file
from src.comparison import compare_files

def main():
    """Main application function"""
    # Setup page
    setup_page()

    # Render header
    render_header()

    # Render file upload section
    file1, file2, compare_clicked = render_file_upload_section()

    # Initialize session state for comparison results
    if "comparison_done" not in st.session_state:
        st.session_state.comparison_done = False
    if "detailed_report" not in st.session_state:
        st.session_state.detailed_report = []
    if "summary_report" not in st.session_state:
        st.session_state.summary_report = []
    if "error_details" not in st.session_state:
        st.session_state.error_details = {
            "missing_sheets": [],
            "extra_sheets": [],
            "column_differences": {},
            "row_differences": {},
            "value_differences": {}
        }
    if "data1" not in st.session_state:
        st.session_state.data1 = None
    if "data2" not in st.session_state:
        st.session_state.data2 = None

    # Compare files if both are uploaded and compare button is clicked
    if file1 and file2 and compare_clicked:
        with st.spinner("Comparing files..."):
            try:
                # Read files
                data1 = read_file(file1)
                data2 = read_file(file2)

                # Compare files
                detailed_report, summary_report, error_details = compare_files(data1, data2)

                # Store results in session state
                st.session_state.comparison_done = True
                st.session_state.detailed_report = detailed_report
                st.session_state.summary_report = summary_report
                st.session_state.error_details = error_details
                st.session_state.data1 = data1
                st.session_state.data2 = data2

                # Force a rerun to display the results
                st.experimental_rerun()
            except Exception as e:
                st.error(f"Error comparing files: {str(e)}")

    # Display comparison results if available
    if st.session_state.comparison_done:
        render_comparison_results(
            st.session_state.detailed_report,
            st.session_state.summary_report,
            st.session_state.error_details,
            st.session_state.data1
        )

        # Render download section
        render_download_section(
            st.session_state.data1,
            st.session_state.data2,
            st.session_state.error_details,
            st.session_state.detailed_report,
            st.session_state.summary_report
        )

if __name__ == "__main__":
    main()