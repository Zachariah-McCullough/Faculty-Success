**Faculty Success Excel Processor**

Automate the tedious work. Save a week (40 hours) of manual Excel processing in minutes.

**Overview:**

Faculty Success Excel Processor is a Python script designed to automate the cleaning, filtering, and processing of raw Faculty Success Excel files. Instead of manually opening, inspecting, and reformatting multiple Excel sheets every year=, this script does it for you—efficiently, accurately, and safely.

**With this tool, you can:**

  - Traverse a directory tree and visualize all subfolders and files, with folder modification dates.

  - Automatically identify and process Excel files matching _AY_ in their names.

  - Filter Excel sheets to keep only the relevant columns for each type of Faculty Success data.

  - Clean, standardize, and correct inconsistencies in key columns like College, Discipline, School, and Dates.

  - Replace missing or blank values with defaults to ensure clean datasets.

  - Save processed data as new Excel files—without overwriting the original files.

  - This automation can save a full week of manual work, reduce human errors, and streamline reporting workflows.

**Features:**

  - Directory Tree Visualization: Prints the folder structure of your base directory, including last modification dates for directories.

  - Automated File Processing: Finds Excel files in the most recently modified subdirectory and filters them based on a predefined pattern.

**Custom Column Cleaning:**

  - Standardizes College, School, Discipline, and Campus names.

  - Formats date fields to YYYY-MM-DD.

  - Replaces blank or missing values with "Unknown" where applicable.

  - Flexible Handling: Works across multiple sheet types and file structures without overwriting your original files.

Easy to Use: Simply run the script and provide a base directory path.

### Supported File Types and Columns

The script supports multiple Faculty Success file types, including:

| File Type                                        | Key Columns Kept |
|--------------------------------------------------|-----------------|
| applied_research                                 | First Name, Last Name, Email, College, Discipline, Home Campus, School, USERNAME, TITLE, DESC, STATUS, TYPE |
| awards                                           | First Name, Last Name, Email, College, Discipline, Home Campus, School, USERNAME, NOMREC, NAME, ORG, SCOPE, DATE_START, DATE_END |
| grants                                           | CONGRANT_INVEST_1_FNAME, CONGRANT_INVEST_1_LNAME, USERNAME, TITLE, STATUS, SPONORG |
| creative_works                                   | First Name, Last Name, Email, College, Discipline, Home Campus, School, USERNAME, TYPE, TITLE, STATUS, NAME, VENUE, CITY, STATE, COUNTRY |
| high_impact_practices_scheduled_learning         | First Name, Last Name, Discipline, School, IMPACT, USERNAME, Email, College, Home Campus, TERM_START, TERM_END, TYPE, DELIVERY_MODE, COURSEPRE, COURSENUM, DIVISION, LEVEL, SECTION, SESSION, CLASS_NBR, LOAD_FACTOR, CIP, ENROLL, CHOURS_MIN, CHOURS_MAX, TOTALSCH |
| high_impact_practices_directed_service_learning  | First Name, Last Name, Email, College, Discipline, Home Campus, School, USERNAME, TYPE, COMPSTAGE |
| IP                                               | First Name, Last Name, USERNAME, Email, College, Discipline, Home Campus, School, FORMAT, TYPE, TITLE, NATIONALITY |
| presentations                                    | First Name, Last Name, Email, VENUE, CITY, STATE, COUNTRY, MEETING_TYPE, SCOPE, INVACC, ACADEMIC, College, Discipline, Home Campus, School, USERNAME, TITLE, PRESENTATION_TYPE, NAME, DATE_START, DATE_END |
| publications                                     | First Name, Last Name, Email, INTELLCONT_AUTH_1_STUDENT_LEVEL, College, Discipline, Home Campus, School, GROUP, USERNAME, CONTYPE, STATUS, TITLE, REFEREED |

**Installation**

    Clone the repository:

git clone https://github.com/YOUR-USERNAME/faculty_success.git
cd faculty_success

    Install the dependencies:

pip install -r requirements.txt

Usage

Run the script:

python main.py

You will be prompted to enter the base directory path containing your Faculty Success Excel files.

The script will:

    Scan the directory tree and print folder and file structure.

    Find the most recently modified subdirectory.

    Identify all Excel files with _AY_ in their name.

    Process each file, filtering columns, cleaning data, and saving a new Excel file with _filtered appended to the name.

    Output warnings for missing columns or empty sheets.

**Benefits**

    Saves Time: Eliminates hours of repetitive Excel processing—what normally takes a week can now be done in minutes.

    Reduces Errors: Standardized cleaning and missing data handling ensures consistent, reliable datasets.

    Non-destructive: Original files are never overwritten.

    Easy to Extend: Column mappings can be updated to support new data types or sheets.

**Requirements**

    Python 3.8 or higher

    pandas

    openpyxl

**License**

    MIT License
