import os
import pandas as pd
import zipfile
import requests
from datetime import datetime
from io import BytesIO

# Function to download the file from a given URL or local file path
def download_file(file_path, destination_folder):
    # If the file is a URL
    if file_path.startswith("http"):
        try:
            response = requests.get(file_path)
            response.raise_for_status()  # Will raise an error if the request failed
            filename = os.path.basename(file_path)
            with open(os.path.join(destination_folder, filename), 'wb') as f:
                f.write(response.content)
            return os.path.join(destination_folder, filename)
        except requests.exceptions.RequestException as e:
            print(f"Error downloading {file_path}: {e}")
            return None
    # If the file is a local path
    else:
        if os.path.exists(file_path):
            filename = os.path.basename(file_path)
            destination_path = os.path.join(destination_folder, filename)
            os.copy(file_path, destination_path)
            return destination_path
        else:
            print(f"File not found: {file_path}")
            return None

# Function to zip all the files in the destination folder
def zip_files(file_paths, zip_filename):
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in file_paths:
            zipf.write(file, os.path.basename(file))
    print(f"Files zipped into {zip_filename}")

# Main function to extract data from Excel and download files
def extract_and_download_files(excel_path, destination_folder):
    # Load the Excel sheet
    df = pd.read_excel(excel_path)

    # Get today's date
    today_date = datetime.today().strftime('%Y-%m-%d')

    # Extract distinct Job_names from the Excel file
    job_names = df['Job_name'].unique()

    # Loop through each distinct Job_name and process
    for job_name in job_names:
        print(f"Processing Job_name: {job_name}")

        # Filter rows for today's run_day and the current Job_name
        filtered_data = df[(df['run_day'] == today_date) & (df['Job_name'] == job_name)]

        # Drop duplicate rows based on file_path to avoid downloading the same file multiple times
        filtered_data = filtered_data.drop_duplicates(subset=['file_path'])

        if filtered_data.empty:
            print(f"No data found for Job_name: {job_name} on {today_date}")
            continue

        # Prepare to download files
        file_paths_to_download = []
        for _, row in filtered_data.iterrows():
            file_path = row['file_path']
            downloaded_file = download_file(file_path, destination_folder)
            if downloaded_file:
                file_paths_to_download.append(downloaded_file)

        # If we have any downloaded files, zip them
        if file_paths_to_download:
            zip_filename = os.path.join(destination_folder, f"{job_name}_{today_date}.zip")
            zip_files(file_paths_to_download, zip_filename)
        else:
            print(f"No files were downloaded for {job_name}.")

# Example usage
if __name__ == "__main__":
    # Specify the path to the Excel file
    excel_path = r'C:\path\to\your\excel_file.xlsx'
    destination_folder = r'C:\path\to\destination\folder'  # Path where files will be downloaded and zipped

    # Run the function
    extract_and_download_files(excel_path, destination_folder)
