# For Faculty Success in order to run and alter the data inside the raw files
# It will ask you for the pathway, run a tree on te excel files it finds, then run one after another, once the process is successful
# It will NOT override files

import pandas as pd  # Used for data manipulation and analysis
import openpyxl  # Used for reading and writing Excel files
import glob  # Match file patterns
import os  # Interacting with the operating system (directory and file management)
from datetime import datetime  # Used to format dates

# Function to build a directory tree including files
def build_directory_tree(base_path):
    tree = {}
    for item in os.listdir(base_path):
        item_path = os.path.join(base_path, item)
        if os.path.isdir(item_path):
            tree[item] = build_directory_tree(item_path)  # Recurse into subdirectories
        else:
            tree[item] = None  # Add file to tree
    return tree

# Function to print the directory tree with modification dates for directories only
def print_tree(tree, base_path, indent="", folder_name_padding=80):
    for key in tree:
        item_path = os.path.join(base_path, key)
        if tree[key] is not None:  # It's a folder
            mod_time = os.path.getmtime(item_path)  # Get last modification time
            mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
            # Pad folder name to align modification date to the right
            padded_folder_name = f"{indent}|-- {key}".ljust(folder_name_padding)
            print(f"{padded_folder_name} (Last modified: {mod_date})")
            print_tree(tree[key], item_path, indent + "    ", folder_name_padding)
        else:  # It's a file
            print(f"{indent}|-- {key}")

# Prompt user for the base directory
base_directory_path = input("Enter the base directory path: ")

# Validate if the directory exists
if not os.path.isdir(base_directory_path):
    print(f"The directory '{base_directory_path}' does not exist. Please check the path and try again.")
    exit()

# Build the directory tree
directory_tree = build_directory_tree(base_directory_path)

# Print the directory tree with last modified dates for folders only
print("\nDirectory Tree (Folders with Dates, Files without Dates):")
print_tree(directory_tree, base_directory_path)

# List all subdirectories within the base directory
subdirectories = [os.path.join(base_directory_path, d) for d in os.listdir(base_directory_path) if os.path.isdir(os.path.join(base_directory_path, d))]

# Find the most recently modified subdirectory, based on the tree and folder's modification date
most_recent_directory = max(subdirectories, key=os.path.getmtime)

# Set this directory as the directory path for file processing
directory_path = most_recent_directory

# File pattern for files with _AY in their names
file_pattern = '*_AY_*.xlsx'
file_paths = glob.glob(os.path.join(directory_path, file_pattern))

# Output the files matching the pattern
if file_paths:
    print("\nFiles matching the pattern:")
    for file in file_paths:
        print(file)
else:
    print("\nNO FILES MATCHING THE PATTERN WERE FOUND.")

# Define a mapping of file names to columns to keep
file_columns_mapping = {
    'applied_research': [
        'First Name', 'Last Name', 'Email', 'College (Most Recent)', 'Discipline (Most Recent)', 
        'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 'USERNAME', 'TITLE', 
        'DESC', 'STATUS', 'TYPE'
    ],
    'awards': [
        'First Name', 'Last Name', 'Email', 'College (Most Recent)', 'Discipline (Most Recent)', 
        'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 'USERNAME', 'NOMREC', 
        'NAME', 'ORG', 'SCOPE', 'SCOPE_LOCALE', 'DATE_START', 'DATE_END'
    ],
    'grants': [
        'CONGRANT_INVEST_1_FNAME', 'CONGRANT_INVEST_1_LNAME', 'USERNAME', 'TITLE', 
        'STATUS', 'SPONORG'
    ],
    'creative_works': [
        'First Name', 'Last Name', 'Email', 'College (Most Recent)', 'Discipline (Most Recent)', 
        'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 'USERNAME', 'TYPE', 
        'TITLE', 'STATUS', 'NAME', 'VENUE', 'CITY', 'STATE', 'COUNTRY'
    ],
    'high_impact_practices_scheduled_learning': [
        'First Name', 'Last Name', 'Discipline (Most Recent)', 'School (Most Recent)', 'IMPACT', 
        'USERNAME', 'Email', 'College (Most Recent)', 'Home Campus/Teaching Site (Most Recent)',
        'TYT_TERM', 'TYY_TERM', 'TERM_START', 'TERM_END', 'TYPE', 'DELIVERY_MODE','COURSEPRE', 'COURSENUM',
        'DIVISION', 'LEVEL', 'SECTION', 'SESSION', 'CLASS_NBR', 
        'LOAD_FACTOR', 'CIP', 'ENROLL', 'CHOURS_MIN', 'CHOURS_MAX', 'TOTALSCH'
    ],
    'high_impact_practices_directed_service_learning': [
        'First Name', 'Last Name', 'Email', 'College (Most Recent)', 'Discipline (Most Recent)', 
        'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 'USERNAME', 'TYPE', 
        'COMPSTAGE'
    ],
    'IP': [
        'First Name', 'Last Name', 'USERNAME', 'Email', 'College (Most Recent)', 'Discipline (Most Recent)', 
        'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 'FORMAT', 'TYPE', 
        'TITLE', 'NATIONALITY'
    ],
    'presentations': [
        'First Name', 'Last Name', 'Email', 'VENUE', 'CITY', 'STATE', 'COUNTRY', 'MEETING_TYPE', 
        'SCOPE', 'INVACC', 'ACADEMIC', 'College (Most Recent)', 'Discipline (Most Recent)', 
        'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 'USERNAME', 'TITLE', 
        'PRESENTATION_TYPE', 'NAME', 'DATE_START', 'DATE_END'
    ],
    'publications': [
        'First Name', 'Last Name', 'Email', 'INTELLCONT_AUTH_1_STUDENT_LEVEL', 'College (Most Recent)', 
        'Discipline (Most Recent)', 'Home Campus/Teaching Site (Most Recent)', 'School (Most Recent)', 
        'GROUP', 'USERNAME', 'CONTYPE', 'STATUS', 'TITLE', 'REFEREED'
    ]
}

# Function to clean up 'College', 'Discipline', and 'School' columns
def clean_column(value, column_name):
    if pd.isna(value):
        if column_name in ['City', 'State', 'Country', 'Impact']:
            return value
        else:
            return ' '
    
    # Strip and split by '|'
    parts = [part.strip() for part in value.split('|') if 'honors' not in part.lower()]
    cleaned_value = parts[0] if parts else 'Unknown'

    # Apply specific replacements based on column name
    if column_name == 'Home Campus/Teaching Site (Most Recent)':
        replacements = {
            'Gulf Park': 'Gulf Coast',
            'GCRL': 'Gulf Coast'
        }
        for key, new_value in replacements.items():
            if key in cleaned_value:
                cleaned_value = cleaned_value.replace(key, new_value)
    
    elif column_name == 'College (Most Recent)':
        if 'library' in cleaned_value.lower() or 'libraries' in cleaned_value.lower():
            cleaned_value = 'Education and Human Sciences'
    
    elif column_name == 'School (Most Recent)':
        if 'Center for STEM Education' in cleaned_value:
            cleaned_value = 'STEM Education'
        elif 'Dept of Aerosapce Studies' in cleaned_value:
            cleaned_value = 'Department of Aerospace Studies'

    elif column_name == 'Discipline (Most Recent)':
        if 'Aerospace Studies (Air Force ROTC)' in cleaned_value:
            cleaned_value = 'Aerospace Studies (AFROTC)'
            
    elif column_name in ['DATE_START', 'DATE_END']:
        # Remove time from the datetime value
        cleaned_value = cleaned_value.split(' ')[0]

    return cleaned_value

# Function to find the appropriate mapping key
def find_mapping_key(base_name):
    for key in file_columns_mapping:
        if key.replace(' ', '_').lower() in base_name.lower():
            return key
    return None

# Function to replace blanks or NaNs with 'Unknown' except for specified columns
def replace_blanks_with_unknown(df, exceptions=[]):
    return df.apply(lambda col: col.fillna('Unknown') if col.name not in exceptions else col)

# Process each file
for file_path in file_paths:
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(directory_path, f'{base_name}_filtered.xlsx')

    try:
        # Load the Excel file
        xls = pd.ExcelFile(file_path)
        print(f"\nProcessing file: {file_path}")

        # Create a dictionary to hold filtered data for each sheet
        filtered_sheets = {}

        # Iterate over each sheet in the Excel file
        for sheet_name in xls.sheet_names:
            print(f"  Processing sheet: {sheet_name}")

            # Determine the columns to keep based on the file's base name
            mapping_key = find_mapping_key(base_name)

            if mapping_key:
                try:
                    # Load data from the current sheet
                    data = pd.read_excel(file_path, sheet_name=sheet_name)

                    # Columns to keep for this sheet
                    columns_to_keep = file_columns_mapping[mapping_key]

                    # Check if all columns exist
                    missing_columns = [column for column in columns_to_keep if column not in data.columns]
                    if missing_columns:
                        print(f"    Warning: Columns not found in sheet '{sheet_name}': {', '.join(missing_columns)}")

                    # Filter the data to keep only the necessary columns that exist
                    filtered_data = data[[column for column in columns_to_keep if column in data.columns]]

                    # Apply the cleaning function specifically to the target columns
                    for col in ['Home Campus/Teaching Site (Most Recent)', 'College (Most Recent)', 'School (Most Recent)']:
                        if col in filtered_data.columns:
                            filtered_data.loc[:, col] = filtered_data.apply(lambda row: clean_column(row[col], col), axis=1)

                    # Apply specific transformations based on sheet name
                    if mapping_key == 'awards':
                        # Ensure 'DATE_START' and 'DATE_END' have only date (no time)
                        if 'DATE_START' in filtered_data.columns:
                            filtered_data['DATE_START'] = pd.to_datetime(filtered_data['DATE_START'], errors='coerce').dt.strftime('%Y-%m-%d')
                        if 'DATE_END' in filtered_data.columns:
                            filtered_data['DATE_END'] = pd.to_datetime(filtered_data['DATE_END'], errors='coerce').dt.strftime('%Y-%m-%d')

                    if mapping_key == 'high_impact_practices_scheduled_learning':
                    # Ensure 'TERM_START' and 'TERM_END' have only date (no time)
                        for term_col in ['TERM_START', 'TERM_END']:
                            if term_col in filtered_data.columns:
                                filtered_data.loc[:, term_col] = pd.to_datetime(filtered_data[term_col], errors='coerce').dt.strftime('%Y-%m-%d')

                    elif mapping_key == 'presentations':
                        # Fill blanks in specific columns
                        for col in ['VENUE', 'MEETING_TYPE', 'SCOPE', 'INVACC', 'ACADEMIC']:
                            if col in filtered_data.columns:
                                filtered_data.loc[:, col] = filtered_data[col].fillna('Unknown' if col in ['VENUE'] else 'Unknown')

                        # Ensure 'PRESENTATION_TYPE' and 'NAME' have 'Unknown' if blank
                        for col in ['PRESENTATION_TYPE', 'NAME']:
                            if col in filtered_data.columns:
                                filtered_data[col] = filtered_data[col].fillna('Unknown')
                        # Ensure 'DATE_START' and 'DATE_END' have only date (no time)
                        if 'DATE_START' in filtered_data.columns:
                            filtered_data.loc[:, 'DATE_START'] = pd.to_datetime(filtered_data['DATE_START'], errors='coerce').dt.strftime('%Y-%m-%d')
                        if 'DATE_END' in filtered_data.columns:
                            filtered_data.loc[:, 'DATE_END'] = pd.to_datetime(filtered_data['DATE_END'], errors='coerce').dt.strftime('%Y-%m-%d')

                    # Clean specific columns
                    for col in ['College (Most Recent)', 'Discipline (Most Recent)', 'School (Most Recent)']:
                        if col in filtered_data.columns:
                            filtered_data.loc[:, col] = filtered_data.apply(lambda row: clean_column(row[col], col), axis=1)

                    # Replace blanks or NaNs with 'Unknown' for all columns except the specified exceptions
                    exceptions = ['CITY', 'STATE', 'COUNTRY', 'IMPACT']
                    filtered_data = replace_blanks_with_unknown(filtered_data, exceptions)

                    # Add filtered data to the dictionary
                    filtered_sheets[sheet_name] = filtered_data

                except pd.errors.EmptyDataError:
                    print(f"    Warning: Sheet '{sheet_name}' is empty or has no readable data.")
                except pd.errors.ParserError:
                    print(f"    Error parsing data in sheet '{sheet_name}'.")
                except Exception as e:
                    print(f"    Error processing sheet '{sheet_name}': {e}")

        # Save filtered data to a new Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, filtered_data in filtered_sheets.items():
                filtered_data.to_excel(writer, sheet_name=sheet_name, index=False)

    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
    except PermissionError:
        print(f"Error: Permission denied when accessing '{file_path}'.")
    except Exception as e:
        print(f"Error processing file '{file_path}': {e}")

print("PROCESSING COMPLETE.")