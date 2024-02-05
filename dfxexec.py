import ezdxf
import pandas as pd

def extract_attributes(text):
    attributes = {'Code': '', 'Material': '', 'Thickness': '', 'Quantity': ''}
    lines = text.split('\n')
    
    for line in lines:
        for key in attributes.keys():
            if line.startswith(f'{key}:'):
                attributes[key] = line.replace(f'{key}:', '').strip()
    
    return attributes


def extract_code_texts_from_dxf(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    code_texts = []

    for text in msp.query('TEXT MTEXT'):
        text_content = text.plain_text().strip()
        if text_content.startswith("Code:"):
            code_attributes = extract_attributes(text_content)
            code_texts.append(code_attributes)

    return code_texts

dxf_file_path = "example.dxf"
extracted_code_texts = extract_code_texts_from_dxf(dxf_file_path)

df = pd.DataFrame(extracted_code_texts)

excel_file_path = "part_list.xlsx"
df.to_excel(excel_file_path, index=False)

print(f"Part List successfully saved to {excel_file_path} file.")
