# Script to Explore Zip File Structure

import zipfile
import os

# Function to display the structure of a zip file

def explore_zip_structure(zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # List of all files in the zip
        zip_contents = zip_ref.namelist()
        print(f'Contents of {zip_file_path}:')
        for file in zip_contents:
            print(file)

# Example usage
if __name__ == '__main__':
    explore_zip_structure('path_to_your_zip_file.zip')