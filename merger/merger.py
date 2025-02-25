import os
import sys
import pandas as pd

def merge_xlsx_files(output_file, file_paths):
    merged_df = pd.DataFrame()
    
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"Warning: {file_path} does not exist and will be skipped.")
            continue
        df = pd.read_excel(file_path)
        merged_df = pd.concat([merged_df, df], ignore_index=True)
    
    # Remove duplicate rows based on 'cve' and 'Host' columns
    if 'cve' in merged_df.columns and 'Host' in merged_df.columns:
        merged_df = merged_df.drop_duplicates(subset=['cve', 'Host'])
    if "cvss3_base_score" in merged_df.columns:
        merged_df["cvss3_base_score"] = pd.to_numeric(merged_df["cvss3_base_score"], errors='coerce')
        merged_df.sort_values(by='cvss3_base_score', ascending=False, inplace=True)
    merged_df.to_excel(output_file, index=False)
    print(f"Merged file saved as {output_file}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <num_files> <output_filename>")
        sys.exit(1)
    
    num_files = int(sys.argv[1])
    output_file = sys.argv[2]
    file_paths = [input(f"Enter path for file {i+1}: ") for i in range(num_files)]
    
    merge_xlsx_files(output_file, file_paths)
