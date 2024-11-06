
import streamlit as st
from openpyxl import load_workbook, Workbook
from io import BytesIO
import os

st.title("District AFR Data Extractor")

# Define the cell mappings based on the template
cell_mapping = {
    "A": "Column A references",  # Placeholder to include Column A references if needed
    "C2": "Program 1",
    "D2": "Program 2",
    "E2": "Program 3",
    "F2": "Program 4",
    "G2": "Program 5",
    "H2": "Program 6",
    # Add mappings up to H60 for each program or reference
    "L2": "Additional Reference L2",
    "M2": "Additional Reference M2",
    "N2": "Additional Reference N2",
}

# Upload multiple AFR files
uploaded_files = st.file_uploader("Upload AFR Workbooks (2017-2024)", type="xlsx", accept_multiple_files=True)

# Process the uploaded files
if uploaded_files:
    # Initialize a new workbook to store the extracted data
    output_wb = Workbook()
    output_wb.remove(output_wb.active)  # Remove the default blank sheet

    for uploaded_file in uploaded_files:
        # Extract the name of the file without extension for sheet naming
        file_name = os.path.splitext(uploaded_file.name)[0]

        # Load the uploaded workbook
        wb = load_workbook(uploaded_file, data_only=True)

        # Check if "Summary" sheet exists in the workbook
        if "Summary" in wb.sheetnames:
            summary_sheet = wb["Summary"]
            # Create a new sheet in the output workbook named after the uploaded file
            output_sheet = output_wb.create_sheet(title=file_name)

            # Write headers in the new sheet
            output_sheet.append(["Cell", "Description", "Value"])

            # Extract data based on the cell mapping
            for cell, description in cell_mapping.items():
                cell_value = summary_sheet[cell].value if cell in summary_sheet else None  # Get the value or None if cell not found
                output_sheet.append([cell, description, cell_value])  # Append cell, description, and value

    # Save the output workbook to a BytesIO buffer for download
    output = BytesIO()
    output_wb.save(output)
    output.seek(0)

    # Provide a download button for the generated file
    st.download_button(
        label="Download Combined AFR Summary",
        data=output,
        file_name="Combined_AFR_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
