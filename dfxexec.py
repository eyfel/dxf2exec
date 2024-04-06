import ezdxf
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd

class DXFAttributeExtractorApp:
    def __init__(self, root):
        self.root = root
        self.setup_ui()

    def setup_ui(self):
        self.info_label = tk.Label(self.root, text="Please enter attribute names separated by commas (e.g., Code, Material, Thickness, Quantity)")
        self.info_label.pack(pady=10)

        self.attribute_entry = tk.Entry(self.root, width=50)
        self.attribute_entry.pack(pady=10)
        self.attribute_entry.insert(0, "")

        self.file_select_button = ttk.Button(self.root, text="Select DXF File", command=self.select_file)
        self.file_select_button.pack(pady=10)

        self.run_button = ttk.Button(self.root, text="Extract Attributes and Save to Excel", command=self.run_extraction)
        self.run_button.pack(pady=10)

        self.file_path = ""

    def select_file(self):
        self.file_path = filedialog.askopenfilename(title="Select a DXF file", filetypes=[("DXF files", "*.dxf")])

    def run_extraction(self):
        if self.file_path:
            attribute_names = [attr.strip() for attr in self.attribute_entry.get().split(',')]
            extracted_code_texts = self.extract_code_texts_from_dxf(self.file_path, attribute_names)

            df = pd.DataFrame(extracted_code_texts, columns=attribute_names)
            excel_file_path = "part_list.xlsx"
            df.to_excel(excel_file_path, index=False)

            print(f"Part List successfully saved to {excel_file_path} file.")
        else:
            print("No file selected.")

    def extract_code_texts_from_dxf(self, file_path, attribute_names):
        doc = ezdxf.readfile(file_path)
        msp = doc.modelspace()

        all_texts = []

        for text in msp.query('TEXT MTEXT'):
            text_content = text.plain_text().strip()
            text_dict = {name: '' for name in attribute_names}
            for line in text_content.split('\n'):
                key, sep, value = line.partition(':')
                if key.strip() in text_dict:
                    text_dict[key.strip()] = value.strip()
            all_texts.append(text_dict)

        return all_texts

if __name__ == "__main__":
    root = tk.Tk()
    root.title("DXF Attribute Extractor")
    app = DXFAttributeExtractorApp(root)
    root.mainloop()
