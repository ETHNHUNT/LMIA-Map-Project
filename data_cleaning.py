import pandas as pd
import re

# Load the Excel file
df = pd.read_excel('2022.xlsx')
df2 = pd.read_excel('2023.xlsx')
df3= pd.read_excel('2024Q1.xlsx')

# Function to apply title case before the comma, handling apostrophes correctly
def apply_title_before_comma(text):
    if isinstance(text, str):  # Check if the value is a string
        parts = text.split(',', 1)  # Split the text at the first comma
        if len(parts) > 1:
            # Apply title case to the part before the comma
            parts[0] = ' '.join([word.capitalize() if "'" not in word else word[0].capitalize() + word[1:].lower() for word in parts[0].strip().split()])
            return ', '.join(parts)  # Join back the parts with a comma and space
        else:
            return text
    return text  # Return the value as is if it's not a string

df['Address'] = df['Address'].apply(apply_title_before_comma)
df["Employer"] = df["Employer"].str.upper().str.title()
df2['Address'] = df2['Address'].apply(apply_title_before_comma)
df2["Employer"] = df2["Employer"].str.upper().str.title()
df3['Address'] = df3['Address'].apply(apply_title_before_comma)
df3["Employer"] = df3["Employer"].str.upper().str.title()

# Save the updated DataFrame to a new Excel file
df.to_excel('Final_2022.xlsx', index= False)
df2.to_excel('Final_2023.xlsx', index=False)
df3.to_excel('Final_2024Q1.xlsx', index=False)
print(df3[['Address']].head(10))

