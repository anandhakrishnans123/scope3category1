import pandas as pd
import random
import streamlit as st
import io

def process_files(client_file, template_path, column_mapping):
    # Read the client sheet (assuming there's only one sheet in the client workbook)
    client_df = pd.read_excel(client_file, sheet_name=None)
    client_data = list(client_df.values())[0]  # Assumes the first sheet is the relevant one

    # Read the template workbook and specific sheet
    template_df = pd.read_excel(template_path, sheet_name=None)
    template_sheet_name = 'Import data file_Manufacturing'
    template_data = template_df[template_sheet_name]

    # Preserve the first row (assuming it contains headers)
    preserved_header = template_data.iloc[:0, :]

    # Create an empty DataFrame with the template columns, excluding the preserved header
    matched_data = pd.DataFrame(columns=template_data.columns)

    # Match and transfer data based on the column_mapping dictionary
    for client_col, template_col in column_mapping.items():
        if client_col in client_data.columns and template_col in template_data.columns:
            matched_data[template_col] = client_data[client_col]

    # Format the 'Res_Date' column if 'Start Date' exists in the client data
    if 'Job Date' in client_data.columns:
        matched_data['Res_Date'] = pd.to_datetime(matched_data['Res_Date']).dt.date

    matched_data['CF Standard'] = matched_data['CF Standard'].apply(lambda x: random.choice(['IATA']) if pd.isna(x) else x)
    matched_data['Gas'] = matched_data['Gas'].apply(lambda x: random.choice(['CO2']) if pd.isna(x) else x)

    # Combine the preserved header with the matched data
    final_data = pd.concat([preserved_header, matched_data], ignore_index=True)

    return final_data

st.title("Freight Data Processor")

# Upload client file
client_file = st.file_uploader("Upload Client Workbook", type=['xls', 'xlsx'])

# Define the path for the template file
template_path = 'Freight-Sample_scope3.xlsx'

if client_file:
    st.write("Processing files...")

    # Read client data to get column names for dynamic mapping
    client_df = pd.read_excel(client_file, sheet_name=None)
    client_data = list(client_df.values())[0]
    client_columns = client_data.columns.tolist()

    # Default column mapping dictionary (this can be customized by the user below)
    default_mapping = {
        'Job Date': 'Res_Date',
        'Consolidation Type': 'Facility',
        'POL': 'Departure',
        'POD': 'Arrival',
        'ATA': 'Start Date',
        'ATD': 'End Date',
        'Weight(Tons)': 'Weight Ton',
        'Weight(Kg)': 'Activity Unit'
    }

    # Allow the user to map the columns dynamically
    column_mapping = {}
    st.write("Map the client columns to the template columns:")

    for client_col, default_template_col in default_mapping.items():
        template_col = st.selectbox(f"Map '{client_col}' to template column", client_columns, index=client_columns.index(default_template_col) if default_template_col in client_columns else 0)
        column_mapping[client_col] = template_col

    # Process the files with the custom column mapping
    final_data = process_files(client_file, template_path, column_mapping)
    
    # Display the final data
    st.write("Processed Data:")
    st.dataframe(final_data)

    # Save the final data to an in-memory file and provide a download link
    output = io.BytesIO()
    final_data.to_excel(output, index=False)
    output.seek(0)
    
    st.download_button(
        label="Download Processed Data",
        data=output,
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.write("Please upload the client file to proceed.")
