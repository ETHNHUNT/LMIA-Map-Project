import pandas as pd

# Load the Excel file
#file_path = 'D:/LMIA-Map-Project/Q1 2024.csv'
df = pd.read_csv('Q1 2024.csv')

# Delete the first row (index 0)
#df = df.drop(index=0)

# Delete the last 8 rows
#df = df[:-8]  # Slicing the DataFrame to exclude the last 8 rows

# Apply title case to only the part after the comma in column 'D'
def apply_title_after_comma(text):
    parts = text.split(',', 1)
    if len(parts) > 1:
        # Apply title case to the second part after the comma
        parts[0] = parts[0].strip().title()
        return ', '.join(parts)
    else:
        return text

df['Address'] = df['Address'].apply(apply_title_after_comma)


# Save the updated DataFrame to a new Excel file
updated_file_path = 'D:/LMIA-Map-Project/updated_q1_2024.csv' 
df.to_csv('updated_q1_2024.csv', index= False)
#print('u')