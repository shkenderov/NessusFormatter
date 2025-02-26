# Nessus Scan Report Generation Tools

This repository contains Python tools for processing Nessus scan data. These tools help convert Nessus scan outputs (in either .csv or .nessus formats) into easy-to-read and presentable Excel spreadsheets. 

The Nessus File Format Converter is particularly useful because the Nessus Scanner currently does not support sending PDF/CSV reports and `.nessus` files simultaneously and automatically (trough Email). This creates an inconvenience, as the `.nessus` file serves as an importable **backup** of the scan that can be uploaded back into the scanner platform. With this tool, the `.nessus` file can be the source of an Excel spreadsheet report, removing the need to manually export reports from the nessus scanner. 

Also, the included CSV converter tool can provide the same functionality, but converting from nessus CSV exports. This improves the presentation, as the default CSV is quite congested and hard to read.

## Tools Included

### **1. Nessus File Format to Excel Converter**
- Processes Nessus `.nessus` XML files and extracts vulnerability data.
- Exports results as a formatted Excel spreadsheet with:

    - Column sorting by CVSS scores.
    - Reorganized columns for better readability.
    - Styled headers and alternating row colors.

- Usage

    ```python nessusfile_To_Excel.py <input_file.nessus> <output_file.xlsx>```

    - Input: Path to the Nessus .nessus XML file.
    - Output: Path to the formatted Excel file (.xlsx).

### **2. Nessus CSV to Excel Converter**

- Converts Nessus scan outputs in CSV format to styled Excel spreadsheets.
- Features include:
    - Removing unwanted newlines within cells.
    - Sorting rows by the CVSS score column (if available).
    - Applying header styling and alternating row colors for better readability.

- Usage:

     ```python nessus_csv_converter.py <csv_file> <excel_file>```

    - Input: Path to the Nessus CSV file.
    - Output: Path to the formatted Excel file (.xlsx).

### **3. Excel Spreadsheet merger**
- Merges a number of excel output spreadsheets in case user wants results combined in one place
- Features include:
    - Drops Host and CVE duplicate pairs.
    - Sorting rows by the CVSSv3 score column (if available).
-Usage:
    ```Usage: python script.py <num_files> <output_filename>```
    The user is then prompted for path to each file

## **Prerequisites**

- Python 3.13 or later.
- Required Python libraries listed in requirements.txt:
    - pandas
    - openpyxl
    - bs4
    - xlsxwriter

    Install dependencies with:
```pip install -r requirements.txt```

## Important Notes

- Security Considerations: Nessus files often contain sensitive information. Ensure that your use of this tool complies with your organization's security policies and does not expose sensitive data.

- Ethical Use: This tool is intended for ethical use by security professionals or organizations conducting authorized vulnerability assessments. Do not use it for malicious purposes.

- Tenable Licensing: Ensure compliance with Tenable's licensing agreements when using this tool with Nessus scan data.
    

## Contribution

Contributions are welcome! Feel free to open issues or submit pull requests.

## License

This project is licensed under the MIT License.
