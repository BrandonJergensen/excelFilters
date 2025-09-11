import pandas as pd

# Step 1: Load the CSVs from XRENS and XRENC
df1 = pd.read_csv('StudentCourseData.Rpt09.11.2025.csv')
df2 = pd.read_csv('CourseSectionsData.Rpt09.11.2025.csv', encoding='latin1')
demo = pd.read_csv('Demographic09.11.2025.csv')

# Step 2: Clean and align merge keys
for df in [df1, df2]:
    df['Course'] = df['Course'].astype(str).str.strip()
    df['SectionNumber'] = df['SectionNumber'].astype(str).str.strip()

# Step 3: Merge course section data into df1 using Course + SectionNumber
merged = df1.merge(
    df2[['Course', 'SectionNumber', 'SectionTitle', 'StartDate', 'EndDate']],
    on=['Course', 'SectionNumber'],
    how='left'
)

# Step 4: Merge demographic info on ID
merged = merged.merge(
    demo[['ID', 'Age', 'Gender', 'Ethnicity']],
    on='ID',
    how='left'
)

# Step 5: Save final result
merged.to_excel('merged_output09.11.xlsx', index=False)
print("✅ All merges complete. Final file saved as 'merged_output09.11.xlsx'")
