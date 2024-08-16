import pandas as pd
import re
# Load the Excel file
df = pd.read_csv('Q1 2024.csv')

# Apply title case to only the part after the comma in column 'D'
def apply_title_after_comma(text):
    parts = text.split(',', 1)
    if len(parts) > 1:
        # Apply title case to the second part after the comma
       parts[0] = parts[0].strip().title()
       return ', '.join(parts)
    else:
       return text
# Function to title-case words before NL
def title_case_before_postal(text):
    parts = re.split(r'(\s*NL\s+)', text, 1)  # Split around the ' NL ' but keep it in the result
    if len(parts) > 1:
        # Only title-case the first part
        parts[0] = re.sub(r"(\b\w+(?:'\w+)?\b)", lambda match: match.group(0).capitalize(), parts[0])
    return ''.join(parts)
df['Address'] = df['Address'].apply(apply_title_after_comma)
df['Address'] = df['Address'].apply(title_case_before_postal)

#splitting occupation and NOC code
def split_occupation(occupation):
    if pd.isnull(occupation):  # Check for NaN entries
        return pd.Series([None, None])
    noc_code, occupation_desc = occupation.split('-', 1)
    noc_code = noc_code.strip().zfill(4)
    return pd.Series([noc_code.strip(), occupation_desc.strip()])

df["Employer"] = df["Employer"].str.upper().str.title()
df[['NOC Code', 'Occupation']] = df['Occupation'].apply(split_occupation)

# Save the updated DataFrame to a new Excel file
updated_file_path = 'D:/LMIA-Map-Project/updated_q1_2024.csv' 
df.to_csv('updated_q1_2024.csv', index= False)
print(df)

