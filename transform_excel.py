import pandas as pd

def transform_and_save_excel(input_file, output_file):
    # Read the input Excel file
    xl = pd.ExcelFile(input_file)

    # Create a writer object for the output file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Process each sheet
        for sheet_name in xl.sheet_names:
            # Read the sheet into a DataFrame with all columns as text
            df = xl.parse(sheet_name, dtype=str)

            # Melt the dataframe to transform columns into rows
            melted_df = pd.melt(df, var_name='Column Name', value_name='Value')

            # Remove blank rows
            melted_df.dropna(subset=['Value'], inplace=True)

            # Write the transformed DataFrame to the corresponding sheet in the output file
            melted_df.to_excel(writer, index=False, sheet_name=sheet_name)

            # Access the worksheet to set column formats
            worksheet = writer.sheets[sheet_name]

            # Ensure columns are formatted as text
            worksheet.set_column('A:A', None, None, {'num_format': '@'})
            worksheet.set_column('B:B', None, None, {'num_format': '@'})

# Specify the input and output file paths
input_file = 'dvd.xlsx'  # Replace with your input Excel file path
output_file = 'stockdvd.xlsx'  # Replace with your desired output Excel file path

# Transform the data and save to an Excel file
transform_and_save_excel(input_file, output_file)
