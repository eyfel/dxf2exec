# dxfexec
Extracting specific text data from DXF files and organizing it into an XML format such as:

Code: part-10-01
Material: S700
Thickness: 10 mm
Quantity: 2

The script reads the DXF file, extracts text data that begins with "Code:", and parses it into separate columns for "Code," "Material," "Thickness," and "Quantity." The processed data is then saved to an Excel file.

Usage:
1. Place your DXF file in the same directory as the script.
2. Update the 'dxf_file_path' variable with the name of your DXF file.
3. Run the script to extract and organize the text data.
4. The resulting data will be saved in an Excel file named 'part_list.xlsx' in the same directory.

Please make sure to install the required libraries (ezdxf and pandas) before running the script.

Feel free to customize the script and adapt it to your specific DXF files and data formats. Happy coding!

