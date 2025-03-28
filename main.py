import os
import pandas as pd
from pathlib import Path

def scan_and_merge_excel_files(input_folder, output_folder, max_rows=1000000):
    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Find all Excel files in the specified folder and its subfolders
    excel_files = [file for file in Path(input_folder).rglob("*.xlsx")]

    if not excel_files:
        print("No Excel files found in the specified folder.")
        return

    combined_data = []

    # Read and merge all Excel files
    for file in excel_files:
        try:
            print(f"Reading file: {file}")
            data = pd.read_excel(file)
            combined_data.append(data)
        except Exception as e:
            print(f"Error reading file {file}: {e}")

    if not combined_data:
        print("No data to merge.")
        return

    # Concatenate all dataframes
    merged_data = pd.concat(combined_data, ignore_index=True)

    # Save the merged data as a TSV file
    csv_output_file = os.path.join(output_folder, "combined_data.txt")
    merged_data.to_csv(csv_output_file, index=False, sep='\t')
    print(f"Saved TSV: {csv_output_file}")

    # Split the merged data into smaller files if necessary
    num_parts = (len(merged_data) // max_rows) + 1

    for part in range(num_parts):
        start_row = part * max_rows
        end_row = start_row + max_rows
        part_data = merged_data.iloc[start_row:end_row]

        output_file = os.path.join(output_folder, f"combined_part{part + 1}.xlsx")
        part_data.to_excel(output_file, index=False)
        print(f"Saved: {output_file}")

if __name__ == "__main__":
    # Set output folder path
    output_folder = "combined_output"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    input_folder = input("Enter the folder path containing Excel files: ").strip()
    output_folder = os.path.join(input_folder, output_folder)

    scan_and_merge_excel_files(input_folder, output_folder)