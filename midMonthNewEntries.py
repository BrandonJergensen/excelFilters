import pandas as pd

# Paths to the two Excel files
file1_path = "unique_rows_from_workbook2.xlsx"
file2_path = "matched_filtered_fuzzy.xlsx"

# Sheet names (or use the first sheet by default)
sheet1_name = "Sheet1"
sheet2_name = "Sheet1"

# Read both Excel files
df1 = pd.read_excel(file1_path, sheet_name=sheet1_name)
df2 = pd.read_excel(file2_path, sheet_name=sheet2_name)

# Drop duplicate rows to prevent false matches
df1 = df1.drop_duplicates()
df2 = df2.drop_duplicates()

# Compare: Keep only rows in df1 that are not in df2 (row-wise)
df_diff = pd.merge(df1, df2, how='left', indicator=True)
df_only_in_df1 = df_diff[df_diff['_merge'] == 'left_only'].drop(columns=['_merge'])

# Save result
output_file = "unique_rows_from_workbook3.xlsx"
df_only_in_df1.to_excel(output_file, index=False)

print(f"Rows in '{file1_path}' but not in '{file2_path}' written to '{output_file}'.")
