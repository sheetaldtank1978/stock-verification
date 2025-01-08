import pandas as pd

# Define input file paths
details_file = 'details.xlsx'
accreg_file = 'accreg.xlsx'
output_file = 'updated_details.xlsx'

# Read the Excel files
details_df = pd.read_excel(details_file, dtype={'Accno ': str})
accreg_df = pd.read_excel(accreg_file, dtype={'Accno': str})

# Print the columns of each dataframe to debug
print("Columns in details_df:", details_df.columns)
print("Columns in accreg_df:", accreg_df.columns)

# Strip any leading or trailing spaces from column names
details_df.columns = details_df.columns.str.strip()
accreg_df.columns = accreg_df.columns.str.strip()

# Print the trimmed columns of each dataframe to debug
print("Trimmed columns in details_df:", details_df.columns)
print("Trimmed columns in accreg_df:", accreg_df.columns)

# Ensure 'Accno' column is stored as string (object) in both dataframes
details_df['Accno'] = details_df['Accno'].astype(str)
accreg_df['Accno'] = accreg_df['Accno'].astype(str)

# Check if 'Accno' exists in both dataframes
if 'Accno' not in details_df.columns:
    raise KeyError("The column 'Accno' is not found in details_df")
if 'Accno' not in accreg_df.columns:
    raise KeyError("The column 'Accno' is not found in accreg_df")

# Merge the dataframes based on the 'Accno' column
merged_df = pd.merge(details_df, accreg_df, on='Accno', how='left', suffixes=('', '_accreg'))

# Print the merged dataframe for debugging
print("Merged DataFrame columns:", merged_df.columns)

# Update the columns in details_df with the corresponding columns from accreg_df
for column in accreg_df.columns:
    if column != 'Accno' and column + '_accreg' in merged_df.columns:
        merged_df[column] = merged_df[column + '_accreg']
        merged_df.drop(columns=[column + '_accreg'], inplace=True)

# Save the updated details_df to a new Excel file
merged_df.to_excel(output_file, index=False)

print(f"Updated details saved to {output_file}")
