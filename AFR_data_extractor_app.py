import streamlit as st
from openpyxl import load_workbook, Workbook
from io import BytesIO
import os

st.title("District AFR Data Extractor and Template Filler")

# Upload AFR files and the template file
afr_files = st.file_uploader("Upload AFR Workbooks (2017-2024)", type="xlsx", accept_multiple_files=True)
template_file = st.file_uploader("Upload AFR Summary Template", type="xlsx")

# Start processing once all files are uploaded and button is clicked
if st.button("Start Processing") and afr_files and template_file:
    # Initialize a progress bar
    progress_bar = st.progress(0)
    total_files = len(afr_files)
    
    # Load the template workbook to populate with data
    template_wb = load_workbook(template_file, data_only=True)
    template_sheet = template_wb.active  # Use the first sheet in the template as base

    # Initialize a new workbook to store filled templates
    output_wb = Workbook()
    output_wb.remove(output_wb.active)  # Remove the default blank sheet

    # Process each uploaded AFR file
    for i, afr_file in enumerate(afr_files):
        # Extract the name of the file without extension for sheet naming
        file_name = os.path.splitext(afr_file.name)[0]

        # Load the AFR workbook
        afr_wb = load_workbook(afr_file, data_only=True)

        # Check if "Summary" sheet exists in the workbook
        if "Summary" in afr_wb.sheetnames:
            afr_summary_sheet = afr_wb["Summary"]
            
            # Create a new sheet in the output workbook named after the AFR file
            new_sheet = output_wb.create_sheet(title=file_name)

            # Copy the entire template structure to the new sheet
            for row in template_sheet.iter_rows():
                for cell in row:
                    # Copy each cell's value from the template to the new sheet
                    new_sheet[cell.coordinate].value = cell.value

                    # Check if the cell in the template contains a reference to replace
                    if isinstance(cell.value, str) and cell.value in afr_summary_sheet:
                        # Replace the reference with actual data from the AFR summary sheet
                        cell_value = afr_summary_sheet[cell.value].value
                        new_sheet[cell.coordinate].value = cell_value  # Replace with actual data

        # Update progress bar after processing each file
        progress_bar.progress((i + 1) / total_files)

    # Save the combined workbook to a BytesIO buffer for download
    output = BytesIO()
    output_wb.save(output)
    output.seek(0)

    # Provide a download button for the generated combined template
    st.download_button(
        label="Download Filled AFR Summary Template",
        data=output,
        file_name="AFR_Summary_Combined_Template_Filled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Display completion message
    st.success("Data extraction and template filling complete!")
