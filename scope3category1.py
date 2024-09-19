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

    # Format the 'Res_Date' column if 'Job Date' exists in the client data
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

# Preset column mapping (you can modify these as the default mappings)
preset_column_mapping = {
    'Job Date': 'Res_Date',
    'Consolidation Type': 'Facility',
    'POL': 'Departure',
    'POD': 'Arrival',
    'ATA': 'Start Date',
    'ATD': 'End Date',
    'Weight(Tons)': 'Weight Ton',
    'Weight(Kg)': 'Activity Unit'
}

# If the client file is uploaded, show options to select or modify the template mapping
if client_file:
    # Load the client data to display options
    client_df = pd.read_excel(client_file, sheet_name=None)
    client_data = list(client_df.values())[0]  # Get the first sheet's data

    # Load the template to get template column options
    template_df = pd.read_excel(template_path, sheet_name=None)
    template_data = template_df['Import data file_Manufacturing']  # Assumes this is the relevant sheet

    st.write("Select columns in the template to map from the client data (preset values provided):")
    
    # Allow the user to edit the mapping by using a selectbox for each template column
    column_mapping = {}
    for client_col, preset_template_col in preset_column_mapping.items():
        # Let user select from available template columns for each client column
        column_mapping[client_col] = st.selectbox(
            f"Select template column for client column '{client_col}'", 
            template_data.columns, 
            index=template_data.columns.get_loc(preset_template_col) if preset_template_col in template_data.columns else 0
        )
    
    st.write("Processing files...")
    
    # Process the files using the updated column mapping
    final_data = process_files(client_file, template_path, column_mapping)

    # Display the processed data
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
