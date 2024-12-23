# Harding-Leovner

File Download and Zip Script
This Python script extracts file paths from an Excel sheet, downloads the files associated with those paths (based on today's date and specified Job_name), removes duplicates, and zips the downloaded files into a .zip archive.

Prerequisites
Before running the script, make sure you have Python installed on your system.

Requirements
Python 3.6 or higher
The following Python packages:
pandas
openpyxl
requests
The script will:

Read an Excel file with columns: Job_name, Job_type, file_path, and run_day.
Extract all unique file_path values for a specific Job_name and today's date (run_day).
Download the files.
Zip the downloaded files and save them to a specified destination.
Installation and Setup
Follow these steps to set up a virtual environment, install the required packages, and run the script.

1. Clone or Download the Repository
Clone this repository or download the script to your local machine.

2. Create a Virtual Environment
Run the following commands to create and activate a virtual environment:

bash
Copy code
# Navigate to the folder where the script is located
cd path/to/your/script

# Create a virtual environment (you can name it 'venv' or anything you like)
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
3. Install Required Packages
Once the virtual environment is activated, install the required packages:

# Install the required dependencies from requirements.txt
pip install -r requirements.txt

4. Run the Script
Now, you are ready to run the script. Make sure to update the excel_path and destination_folder variables in the script to point to your Excel file and desired folder where the downloaded and zipped files should be saved.

RUN THE SCRIPT:

python main.py
Script Configuration
excel_path: The full path to the Excel file containing the job details.
destination_folder: The folder where you want to store the downloaded files and the zip archive.
Example:

excel_path = r'C:\path\to\your\excel_file.xlsx'
destination_folder = r'C:\path\to\destination\folder'
The script will process each distinct Job_name in the Excel file, filter for today's run_day, and download the associated files (removing duplicates). All downloaded files will be zipped into a .zip file named after the Job_name and today's date.


SAMPLE OUTPUT:
Processing Job_name: Job1
Files zipped into C:\path\to\destination\folder\Job1_2024-12-23.zip



NOTES:
File Path: Ensure the file_path column in the Excel file contains either valid URLs or valid local file paths.
Error Handling: If any file cannot be found or downloaded, it will be skipped, and the script will proceed to the next file.
Duplication: The script removes duplicate records based on the file_path column to avoid downloading the same file multiple times.

TROUBLESHOOTING:
If you face issues with the Excel file not loading, make sure the file format is .xlsx (the script uses openpyxl for reading .xlsx files).
If there are issues with downloading files from URLs, ensure your internet connection is stable, and that the URLs are correct and accessible.

LICENSE:
This script is free to use. No warranty is provided. Please feel free to modify and extend it according to your needs.

SUMMARY OF THE PROCESS:
Create Virtual Environment: python -m venv venv
Activate Virtual Environment: source venv/bin/activate or venv\Scripts\activate
Install Dependencies: pip install -r requirements.txt
Run the Script: python main.py