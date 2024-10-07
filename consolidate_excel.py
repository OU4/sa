import pandas as pd
import xlwings as xw
import os
from pathlib import Path

def read_excel_file(file_path):
    if file_path.suffix == '.xls':
        with xw.App(visible=False) as app:
            book = app.books.open(file_path)
            sheet = book.sheets[0]
            data = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
            book.close()
        return data
    elif file_path.suffix == '.xlsx':
        return pd.read_excel(file_path, engine='openpyxl')
    else:
        raise ValueError(f"Unsupported file format: {file_path.suffix}")

def consolidate_excel_files(directory_path, output_file):
    # Get all .xls and .xlsx Excel files in the specified directory
    excel_files = list(Path(directory_path).glob('*.xls')) + list(Path(directory_path).glob('*.xlsx'))
    
    print(f"Directory being searched: {directory_path}")
    print(f"Number of Excel files found: {len(excel_files)}")
    
    if not excel_files:
        print("No Excel files found. Please check the directory path and file extensions.")
        return

    # Create an empty list to store individual dataframes
    dfs = []

    # Loop through each Excel file
    for file in excel_files:
        print(f"Processing file: {file}")
        try:
            # Read the Excel file
            df = read_excel_file(file)
            
            # Add a column with the filename (optional, for tracking purposes)
            df['Source_File'] = file.name
            
            # Append the dataframe to our list
            dfs.append(df)
        except Exception as e:
            print(f"Error processing {file}: {str(e)}")

    if not dfs:
        print("No data was read from the Excel files. Please check if the files are empty or corrupted.")
        return

    # Concatenate all dataframes in the list
    combined_df = pd.concat(dfs, ignore_index=True)

    # Save the combined dataframe to a new Excel file
    combined_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Combined data saved to {output_file}")

# Example usage
directory_path = '/Users/abdulazizdot/Desktop/tadawul'  # Replace with your actual directory path
output_file = 'combined_financial_data.xlsx'  # Name of the output file

consolidate_excel_files(directory_path, output_file)