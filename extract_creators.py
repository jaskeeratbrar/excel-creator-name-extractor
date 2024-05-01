import os
import openpyxl
import csv

def extract_creator_from_properties(directory):
    files = os.listdir(directory)
    excel_files = [file for file in files if file.endswith('.xlsx')]
    creators_info = []

    for file in excel_files:
        path = os.path.join(directory, file)
        try:
            workbook = openpyxl.load_workbook(path, read_only=True)
            properties = workbook.properties
            creator_name = properties.creator
            if creator_name:
                # Remove commas and then split
                parts = creator_name.replace(',', '').split()
                if len(parts) > 1:
                    first_name = ' '.join(parts[:-1])  # First name might be more than one part
                    last_name = parts[-1]
                else:
                    first_name = parts[0]
                    last_name = ""
                creators_info.append([file, first_name, last_name])
        except Exception as e:
            print(f"Failed to process {file}: {e}")

    with open('creators_output.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['File Name', 'First Name', 'Last Name'])
        writer.writerows(creators_info)

    return 'Data extracted and saved to creators_output.csv'

###########################################################################
###########################################################################
# Replace below directory path to your working directory
# where excel files are stored
###########################################################################
###########################################################################

directory_path = '/Users/jassibrar/Desktop/mgb msba ramon/School Submissions' 

###########################################################################
###########################################################################
###########################################################################
result = extract_creator_from_properties(directory_path)
print(result)
