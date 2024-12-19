#CSV with changes

 

import pandas as pd

from google.colab import files

import re

from datetime import datetime

 

# Upload the file in Google Colab

print("Please upload your Excel file.")

uploaded = files.upload()

 

# Load the uploaded Excel file

input_file = list(uploaded.keys())[0]

try:

    data = pd.read_excel(input_file)

    print(f"File {input_file} loaded successfully.")

except Exception as e:

    print(f"Error loading the file: {e}")

    exit()

 

# Check if required columns exist (case-insensitive)

required_columns = ['PROJECT_DISPLAY_ID', 'PROJECT_DESCRIPTION']

data.columns = data.columns.str.strip()  # Clean up column names

missing_columns = [col for col in required_columns if col.lower() not in map(str.lower, data.columns)]

if missing_columns:

    print(f"Error: Missing required columns: {missing_columns}")

    print(f"Available columns: {data.columns.tolist()}")

    exit()

 

# Clean PROJECT_DISPLAY_ID

def clean_project_id(project_display_id):

    try:

        return int(project_display_id)

    except ValueError:

        print(f"Warning: Invalid PROJECT_DISPLAY_ID '{project_display_id}' replaced with NaN.")

        return None

 

# Extract multiple EX: entries

def extract_ex_info(description):

    """

    Extracts all 'EX: YYYY.MM.DD description' entries from project description.

    Handles multiple occurrences of 'EX:'.

    """

    try:

        description = str(description) if pd.notna(description) else ""

        # Regex to match all occurrences of 'EX:' followed by a date and comment

        matches = re.findall(r"EX:\s*(\d{4}\.\d{2}\.\d{2})\s*(.+?)(?=EX:|\Z)", description, re.DOTALL)

        return matches if matches else []

    except Exception as e:

        print(f"Error processing description: {description}. Error: {e}")

        return []

 

# Apply functions to data

data['PROJECT_DISPLAY_ID'] = data['PROJECT_DISPLAY_ID'].apply(clean_project_id)

data['Extracted Info'] = data['PROJECT_DESCRIPTION'].apply(extract_ex_info)

 

# Expand multiple EX: entries into separate rows

expanded_data = data.explode('Extracted Info')

 

# Split extracted info into Date and Executive Comment

if 'Extracted Info' in expanded_data and not expanded_data['Extracted Info'].isna().all():

    expanded_data = expanded_data.reset_index(drop=True)

    temp_df = pd.DataFrame(

        expanded_data['Extracted Info'].dropna().tolist(),

        columns=['Date', 'Executive Comment'],

        index=expanded_data['Extracted Info'].dropna().index

    ).reset_index(drop=True)

    expanded_data.loc[expanded_data['Extracted Info'].notna(), ['Date', 'Executive Comment']] = temp_df.values

else:

    expanded_data['Date'] = None

    expanded_data['Executive Comment'] = None

 

# Clean the Executive Comment column

def clean_executive_comment(comment):

    """

    Removes leading '=-' or '-' from the executive comment.

    Replaces blank comments with 'NULL'.

    """

    try:

        comment = str(comment).strip() if pd.notna(comment) else ""

        cleaned_comment = re.sub(r"^[-=]+", "", comment).strip()

        return cleaned_comment if cleaned_comment else "NULL"

    except Exception as e:

        print(f"Error cleaning comment: {comment}. Error: {e}")

        return "NULL"

expanded_data['Executive Comment'] = expanded_data['Executive Comment'].apply(clean_executive_comment)

 

# For IDs without EX:, add "NULL" for missing values

expanded_data['Date'] = expanded_data['Date'].fillna("NULL")

expanded_data['Executive Comment'] = expanded_data['Executive Comment'].fillna("NULL")

 

# Extract relevant columns

extracted_data = expanded_data[['PROJECT_DISPLAY_ID', 'Date', 'Executive Comment']]

 

# Save the extracted data to a new Excel file

output_file = 'Executive Updates.csv'

if not extracted_data.empty:

    try:

        extracted_data.to_csv(output_file, index=False)

        print(f"Data successfully extracted to {output_file}")

        files.download(output_file)

    except Exception as e:

        print(f"Error: Could not save the file. {e}")

else:

    print("No valid data to save. Check the input file.")

 
