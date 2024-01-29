import ezdxf
import pandas as pd

# Function to extract attributes from text content
def extract_attributes(text):
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

# Function to extract code attributes from DXF file
def extract_code_attributes_from_dxf(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    code_attributes_list = []

    for text in msp.query('TEXT MTEXT'):
        text_content = text.plain_text().strip()
        if text_content.startswith("Code:"):
            code_attributes = extract_attributes(text_content)
            code_attributes_list.append(code_attributes)

    return code_attributes_list

# DXF file path
dxf_file_path = "ornek.dxf"

# Extract code attributes from the DXF file
extracted_code_attributes = extract_code_attributes_from_dxf(dxf_file_path)

# Create a DataFrame from the extracted code attributes
df = pd.DataFrame(extracted_code_attributes)

# Excel file path
excel_file_path = "parca_listesi.xlsx"

# Save the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)

print(f"Part List saved successfully to {excel_file_path}.")
