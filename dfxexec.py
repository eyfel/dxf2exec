import ezdxf
import pandas as pd

# Function to extract attributes from text content
def extract_kod_text_attributes(text):
    attributes = {'Code': '', 'Material': '', 'Thickness': '', 'Quantity': ''}
    lines = text.split('\n')
    
    for line in lines:
        if line.startswith('Code:'):
            attributes['Code'] = line.replace('Code:', '').strip()
        elif line.startswith('Material:'):
            attributes['Material'] = line.replace('Material:', '').strip()
        elif line.startswith('Thickness:'):
            attributes['Thickness'] = line.replace('Thickness:', '').strip()
        elif line.startswith('Quantity:'):
            attributes['Quantity'] = line.replace('Quantity:', '').strip()
    
    return attributes

# Function to extract code texts from DXF file
def extract_code_texts_from_dxf(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    code_texts = []

    for text in msp.query('TEXT MTEXT'):
        text_content = text.plain_text().strip()
        if text_content.startswith("Code:"):
            code_attributes = extract_kod_text_attributes(text_content)
            code_texts.append(code_attributes)

    return code_texts

# Path to the DXF file
dxf_file_path = "example.dxf"

# Extract code texts from the DXF file
extracted_code_texts = extract_code_texts_from_dxf(dxf_file_path)

# Create a DataFrame from the extracted code texts
df = pd.DataFrame(extracted_code_texts)

# Define the path for the Excel file to save the data
excel_file_path = "part_list.xlsx"

# Save the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)

# Print a success message
print(f"Part List saved successfully to {excel_file_path}.")
