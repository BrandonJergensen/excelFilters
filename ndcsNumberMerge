import pandas as pd
from rapidfuzz import fuzz
from datetime import datetime

# --- Load Files ---
merged = pd.read_excel('merged_output09.08.xlsx')  # Your enriched student data
inmates = pd.read_excel('inmateDB.xlsx', sheet_name='Record Type 1')  # NDCS dataset

# --- Normalize inmate columns ---
inmates['FIRST NAME'] = inmates['FIRST NAME'].fillna('').str.upper().str.strip()
inmates['COMMITTED LAST NAME'] = inmates['COMMITTED LAST NAME'].fillna('').str.upper().str.strip()
inmates['GENDER'] = inmates['GENDER'].fillna('').str.upper().str.strip()
inmates['DATE OF BIRTH'] = pd.to_datetime(inmates['DATE OF BIRTH'], errors='coerce')

# --- Normalize merged columns ---
merged['FirstName'] = merged['FirstName'].fillna('').str.upper().str.strip()
merged['LastName'] = merged['LastName'].fillna('').str.upper().str.strip()
merged['Gender'] = merged['Gender'].fillna('').str.upper().str.strip()

# --- Define fuzzy similarity function ---
def similar(a, b, threshold=80):
    return fuzz.ratio(a, b) >= threshold

# --- Matching function ---
def get_matches(row):
    first = row['FirstName']
    last = row['LastName']
    gender = row['Gender'] if row['Gender'] in ['M', 'F'] else None
    age = row['Age']
    
    if pd.isnull(age):
        return ''

    # Estimate birth year range
    current_year = datetime.now().year
    birth_year_low = current_year - int(age) - 1
    birth_year_high = current_year - int(age) + 1

    # Filter first by DOB (within ±1 year) and optionally gender
    inmate_subset = inmates[
        inmates['DATE OF BIRTH'].dt.year.between(birth_year_low, birth_year_high)
    ]
    if gender:
        inmate_subset = inmate_subset[inmate_subset['GENDER'].isin(['MALE' if gender == 'M' else 'FEMALE'])]

    # Apply fuzzy match to names only within this reduced subset
    matches = inmate_subset[
        inmate_subset.apply(
            lambda r: similar(r['FIRST NAME'], first) and similar(r['COMMITTED LAST NAME'], last),
            axis=1
        )
    ]
    
    return ', '.join(matches['ID NUMBER'].astype(str)) if not matches.empty else ''

# --- Run matching ---
merged['MatchedID'] = merged.apply(get_matches, axis=1)

# --- Save result ---
merged.to_excel('matched_filtered_fuzzy2.xlsx', index=False)
print("✅ Filtered fuzzy match complete. Saved as 'matched_filtered_fuzzy1.xlsx'")
