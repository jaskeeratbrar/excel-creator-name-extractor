# Excel Creator Name Extractor

## Overview
The Excel Creator Name Extractor is a Python script designed to extract creator names from Excel files (.xlsx) and output them into a CSV file. This tool is useful for managing and documenting the metadata of Excel files, particularly in environments where tracking document authorship is necessary.

## Use Case
This script is intended for use in scenarios where users need to:
- Audit document authorship.
- Organize and document Excel file metadata for archival purposes.
- Extract author metadata from multiple files efficiently without opening each one manually.

## Features
- Extracts creator names directly from the document properties of Excel files.
- Handles names with multiple parts, including first names, middle names, and last names.
- Outputs the results in a CSV file, including the original file name, extracted first name, and last name.

## Prerequisites
Before running the script, ensure you have Python installed on your system along with the following packages:

- `openpyxl`: For reading Excel files.
- `csv`: For writing output data to CSV format.
  
- Place all your .xlsx files in a single directory.
- Modify the script to point to your directory by changing the directory_path variable to the path of your directory containing the Excel files.

## You can install the required package using pip:

```bash
pip install openpyxl

Modify your working directory where your excel files are stored

```bash
directory_path = 'path_to_your_directory'

```bash
python extract_creators.py


