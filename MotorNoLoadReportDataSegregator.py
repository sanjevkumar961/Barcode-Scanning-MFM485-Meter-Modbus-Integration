import pandas as pd
import os

# Load the Excel file
excel_name = 'MotorNoLoadReportSegregate.xlsx'
excel_dir = os.path.join(os.environ['USERPROFILE'], 'Desktop')
input_file = os.path.join(excel_dir, excel_name)

# Read the data into a pandas DataFrame
data = pd.read_excel(input_file, sheet_name='Data')

# Create a new Excel writer to write back to the same file
with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # Write the original data back to a sheet named 'Data'
    data.to_excel(writer, sheet_name='Data', index=False)
    
    # Extract unique S.No prefixes and segregate data
    data['Prefix'] = data['S.No'].str.split('-').str[0]  # Extract the prefix from S.No
    unique_prefixes = data['Prefix'].unique()
    
    for prefix in unique_prefixes:
        # Filter the data for the current prefix
        filtered_data = data[data['Prefix'] == prefix].drop(columns=['Prefix'])
        
        # Sort the filtered data by S.No, Date, and Time
        sorted_data = filtered_data.sort_values(by=['Date', 'S.No'], ascending=True, na_position='last')
        # Write the sorted, filtered data to a new sheet named after the prefix
        sorted_data.to_excel(writer, sheet_name=prefix, index=False)
    
    # Read all sheet names
    sheet_data = pd.ExcelFile(input_file)
    sheet_names = sheet_data.sheet_names
    
    # Excluded sheet names
    excluded_sheets = {'Config', 'ModelAmp'}
    
    for sheet_name in sheet_names:
        # Skip sorting for excluded sheets
        if sheet_name in excluded_sheets:
            continue
        # Read the current sheet data
        sheet_df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Sort data by Date, S.No
        sorted_df = sheet_df.sort_values(by=['Date', 'S.No'], ascending=True, na_position='last')
        # Write the sorted data back to the same sheet
        sorted_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Data has been segregated, sorted, and saved in the same file: {input_file}.")

