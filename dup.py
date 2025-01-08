import pandas as pd

# Load the input Excel file
input_file = 'stockcd.xlsx'  # Replace with your input file name
output_file = 'duplicates_output.xlsx'  # Output file name

# Read the Excel file
df = pd.read_excel(input_file, dtype={'accno': str})

# Find duplicates based on all columns
df_duplicates = df[df.duplicated(keep=False)]

# Save the duplicates to a new Excel file
df_duplicates.to_excel(output_file, index=False)

print(f"Duplicates extracted and saved to {output_file}")
