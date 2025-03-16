import streamlit as st
import pandas as pd
from io import BytesIO
import sys
import os

# Make sure the assets module can be found
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import colors properly
from assets.colors import (
    PRIMARY_COLOR, SECONDARY_COLOR, ACCENT_COLOR,
    ERROR_COLOR, WARNING_COLOR, SUCCESS_COLOR
)

def setup_page():
    """Configure the page settings and styling"""
    st.set_page_config(page_title="Data Integrity Checker", layout="wide")

    # Custom CSS with explicit color values
    st.markdown(
        f"""
        <style>
        .main .block-container {{
            padding-top: 2rem;
            padding-bottom: 2rem;
        }}
        h1, h2, h3 {{
            color: {PRIMARY_COLOR};
        }}
        .stButton>button {{
            background-color: {SECONDARY_COLOR} !important;
            color: white !important;
            border: none !important;
        }}
        .stButton>button:hover {{
            background-color: {ACCENT_COLOR} !important;
            color: black !important;
        }}
        .report-container {{
            background-color: #f0f2f6;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
        }}
        .error {{
            color: {ERROR_COLOR} !important;
            font-weight: bold;
        }}
        .warning {{
            color: {WARNING_COLOR} !important;
            font-weight: bold;
        }}
        .success {{
            color: {SUCCESS_COLOR} !important;
            font-weight: bold;
        }}
        .stTabs [data-baseweb="tab-list"] {{
            gap: 24px;
        }}
        .stTabs [data-baseweb="tab"] {{
            height: 50px;
            white-space: pre-wrap;
            background-color: white;
            border-radius: 4px 4px 0px 0px;
            gap: 1px;
            padding-left: 10px;
            padding-right: 10px;
        }}
        .stTabs [aria-selected="true"] {{
            background-color: {PRIMARY_COLOR} !important;
            color: white !important;
        }}
        [data-testid="stFileUploader"] {{
            width: 100%;
        }}
        /* Additional styling for color emphasis */
        .stAlert {{
            background-color: {PRIMARY_COLOR}10;
            border-left-color: {PRIMARY_COLOR};
        }}
        .stAlert [data-testid="stText"] {{
            color: {PRIMARY_COLOR};
        }}
        /* Style for success messages */
        .element-container:has([data-testid="stAlert"]:contains("success")) {{
            background-color: {SUCCESS_COLOR}10;
            border-left: 4px solid {SUCCESS_COLOR};
            padding: 10px;
            border-radius: 4px;
        }}
        /* Style for warning messages */
        .element-container:has([data-testid="stAlert"]:contains("warning")) {{
            background-color: {WARNING_COLOR}10;
            border-left: 4px solid {WARNING_COLOR};
            padding: 10px;
            border-radius: 4px;
        }}
        /* Style for error messages */
        .element-container:has([data-testid="stAlert"]:contains("error")) {{
            background-color: {ERROR_COLOR}10;
            border-left: 4px solid {ERROR_COLOR};
            padding: 10px;
            border-radius: 4px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_header():
    """Render the application header with logo"""
    col1, col2 = st.columns([1, 5])
    with col1:
        st.image("https://snb.thesagenext.com/blog/wp-content/uploads/2023/07/Data-Integrity-Banner.webp", width=150)
    with col2:
        st.title("Data Integrity Checker")

    st.markdown("---")

def render_file_upload_section():
    """Render the file upload section"""
    st.header("Upload Files for Comparison")
    st.write("Upload two files (Excel or CSV) to compare their structure and data.")

    col1, col2 = st.columns(2)
    with col1:
        file1 = st.file_uploader("Upload the first file (Base)", type=["xlsx", "csv"],
                                accept_multiple_files=False)
    with col2:
        file2 = st.file_uploader("Upload the second file (Comparison)", type=["xlsx", "csv"],
                                accept_multiple_files=False)

    compare_clicked = st.button("Compare Files", type="primary", disabled=(not file1 or not file2))

    return file1, file2, compare_clicked

def render_comparison_results(detailed_report, summary_report, error_details, data1):
    """Render the comparison results in tabs"""
    st.markdown("---")
    st.header("Comparison Results")

    # Create tabs for different reports
    tab1, tab2, tab3 = st.tabs(["Summary Report", "Detailed Report", "Visual Comparison"])

    with tab1:
        render_summary_report(summary_report, data1)

    with tab2:
        render_detailed_report(detailed_report, data1)

    with tab3:
        render_visual_comparison(error_details)

def render_summary_report(summary_report, data1):
    """Render the summary report tab"""
    if not summary_report:
        st.success("No differences found! The files are identical.")
    else:
        st.warning(f"Found {len(summary_report)} differences between the files.")

        # Group summary by sheet
        sheet_differences = {}
        for item in summary_report:
            sheet = None
            for s in data1.get("sheet_names", []):
                if f"'{s}'" in item:
                    sheet = s
                    break

            if sheet:
                if sheet not in sheet_differences:
                    sheet_differences[sheet] = []
                sheet_differences[sheet].append(item)
            else:
                if "General" not in sheet_differences:
                    sheet_differences["General"] = []
                sheet_differences["General"].append(item)

        # Display grouped summary
        for sheet, items in sheet_differences.items():
            with st.expander(f"{sheet} ({len(items)} differences)", expanded=True):
                for i, item in enumerate(items):
                    st.markdown(f"**{i+1}.** {item}")

def render_detailed_report(detailed_report, data1):
    """Render the detailed report tab"""
    if not detailed_report:
        st.success("No differences found! The files are identical.")
    else:
        st.warning(f"Found {len(detailed_report)} detailed differences between the files.")

        # Group detailed report by sheet
        sheet_differences = {}
        for item in detailed_report:
            sheet = None
            for s in data1.get("sheet_names", []):
                if f"'{s}'" in item:
                    sheet = s
                    break

            if sheet:
                if sheet not in sheet_differences:
                    sheet_differences[sheet] = []
                sheet_differences[sheet].append(item)
            else:
                if "General" not in sheet_differences:
                    sheet_differences["General"] = []
                sheet_differences["General"].append(item)

        # Display grouped detailed report
        for sheet, items in sheet_differences.items():
            with st.expander(f"{sheet} ({len(items)} differences)", expanded=True):
                for i, item in enumerate(items):
                    st.markdown(f"**{i+1}.** {item}")

def render_visual_comparison(error_details):
    """Render the visual comparison tab"""
    st.subheader("Visual Comparison of Differences")

    # Display missing/extra sheets
    if error_details["missing_sheets"] or error_details["extra_sheets"]:
        st.markdown("### Sheet Structure Differences")

        col1, col2 = st.columns(2)
        with col1:
            if error_details["missing_sheets"]:
                st.markdown(f"<div class='error'>Sheets in File 1 but missing in File 2:</div>", unsafe_allow_html=True)
                for sheet in error_details["missing_sheets"]:
                    st.markdown(f"- {sheet}")

        with col2:
            if error_details["extra_sheets"]:
                st.markdown(f"<div class='warning'>Sheets in File 2 but missing in File 1:</div>", unsafe_allow_html=True)
                for sheet in error_details["extra_sheets"]:
                    st.markdown(f"- {sheet}")

    # Display column differences by sheet
    if error_details["column_differences"]:
        st.markdown("### Column Structure Differences")

        for sheet, diff in error_details["column_differences"].items():
            if diff["missing"] or diff["extra"] or diff["reordered"]:
                with st.expander(f"Column differences in '{sheet}'", expanded=True):
                    col1, col2 = st.columns(2)

                    with col1:
                        if diff["missing"]:
                            st.markdown(f"<div class='error'>Columns in File 1 but missing in File 2:</div>", unsafe_allow_html=True)
                            for col in diff["missing"]:
                                st.markdown(f"- {col}")

                    with col2:
                        if diff["extra"]:
                            st.markdown(f"<div class='warning'>Columns in File 2 but missing in File 1:</div>", unsafe_allow_html=True)
                            for col in diff["extra"]:
                                st.markdown(f"- {col}")

                    if diff["reordered"]:
                        st.markdown(f"<div class='warning'>Column order is different between files</div>", unsafe_allow_html=True)

    # Display row differences by sheet
    if error_details["row_differences"]:
        st.markdown("### Row Count Differences")

        for sheet, diff in error_details["row_differences"].items():
            if diff["count_diff"] or diff["missing_rows"] or diff["extra_rows"]:
                with st.expander(f"Row differences in '{sheet}'", expanded=True):
                    if diff["count_diff"]:
                        st.markdown(f"<div class='warning'>Row count mismatch: {diff['count_diff'][0]} rows in File 1 vs {diff['count_diff'][1]} rows in File 2</div>", unsafe_allow_html=True)

                    col1, col2 = st.columns(2)

                    with col1:
                        if diff["missing_rows"]:
                            st.markdown(f"<div class='error'>Rows in File 1 but missing in File 2 ({len(diff['missing_rows'])} rows):</div>", unsafe_allow_html=True)
                            # Show at most 10 missing rows to avoid cluttering the UI
                            for key in list(diff["missing_rows"])[:10]:
                                st.markdown(f"- Key: {key}")
                            if len(diff["missing_rows"]) > 10:
                                st.markdown(f"- ... and {len(diff['missing_rows']) - 10} more")

                    with col2:
                        if diff["extra_rows"]:
                            st.markdown(f"<div class='warning'>Rows in File 2 but missing in File 1 ({len(diff['extra_rows'])} rows):</div>", unsafe_allow_html=True)
                            # Show at most 10 extra rows to avoid cluttering the UI
                            for key in list(diff["extra_rows"])[:10]:
                                st.markdown(f"- Key: {key}")
                            if len(diff["extra_rows"]) > 10:
                                st.markdown(f"- ... and {len(diff['extra_rows']) - 10} more")

    # Display value differences by sheet
    if error_details["value_differences"]:
        st.markdown("### Value Differences")

        for sheet, diffs in error_details["value_differences"].items():
            if diffs:
                with st.expander(f"Value differences in '{sheet}' ({len(diffs)} differences)", expanded=True):
                    # Create a DataFrame to display the differences
                    diff_data = []
                    for diff in diffs[:100]:  # Limit to 100 differences to avoid performance issues
                        if "key" in diff:
                            diff_data.append({
                                "Identifier": f"Key: {diff['key']}",
                                "Column": diff["column"],
                                "Value in File 1": diff["value1"],
                                "Value in File 2": diff["value2"]
                            })
                        else:
                            diff_data.append({
                                "Identifier": f"Row: {diff['row']}",
                                "Column": diff["column"],
                                "Value in File 1": diff["value1"],
                                "Value in File 2": diff["value2"]
                            })

                    if diff_data:
                        diff_df = pd.DataFrame(diff_data)
                        st.dataframe(diff_df, use_container_width=True)

                        if len(diffs) > 100:
                            st.markdown(f"*Showing 100 of {len(diffs)} differences. Download the detailed report for all differences.*")

def render_download_section(data1, data2, error_details, detailed_report, summary_report):
    """Render the download section for highlighted files and reports"""
    st.markdown("---")
    st.header("Download Highlighted Files")

    from src.highlighting import highlight_differences_excel, highlight_differences_csv

    col1, col2 = st.columns(2)

    with col1:
        if data1["type"] == "excel":
            highlighted_file1 = highlight_differences_excel(data1, data2, error_details)
            if highlighted_file1:
                st.download_button(
                    label="Download File 1 with Highlights",
                    data=highlighted_file1,
                    file_name="file1_highlighted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        elif data1["type"] == "csv":
            highlighted_file1 = highlight_differences_csv(data1, data2, error_details)
            if highlighted_file1:
                st.download_button(
                    label="Download File 1 with Highlights",
                    data=highlighted_file1,
                    file_name="file1_highlighted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with col2:
        # Generate detailed report as Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Create a DataFrame for the summary report
            summary_df = pd.DataFrame({"Summary": summary_report})
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            # Create a DataFrame for the detailed report
            detailed_df = pd.DataFrame({"Details": detailed_report})
            detailed_df.to_excel(writer, sheet_name="Detailed", index=False)

        # Return the Excel file as bytes
        output.seek(0)
        report_bytes = output.getvalue()

        st.download_button(
            label="Download Detailed Report",
            data=report_bytes,
            file_name="comparison_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )