from docx import Document
import pandas as pd
from tkinter import Tk, filedialog

def write_multiple_docx(data, index, save_path):
    doc = Document()
    for key, value in data.items():
        doc.add_paragraph(f"{key}:{value}")
    name = f"Document{index + 1}.docx"
    doc.save(f"{save_path}/{name}")
    print(f"Document {name} created successfully")

def write_single_docx(data, doc):
    for key, value in data.items():
        doc.add_paragraph(f"{key}:{value}")
    doc.add_paragraph()
    return doc

def process_files(path, save_path, option):
    if option == '1':
        d = pd.read_excel(path)
        for index, row in d.iterrows():
            data = row.to_dict()
            write_multiple_docx(data, index, save_path)
    elif option == '2':
        single_doc = Document()
        d = pd.read_excel(path)
        for index, row in d.iterrows():
            data = row.to_dict()
            single_doc = write_single_docx(data, single_doc)
        single_doc.save(f"{save_path}/SingleDocument.docx")
        print("Single Document created successfully")

def main():
    root = Tk()
    print("1. Select the Excel file")
    print("2. Exit")
    choice = int(input("Enter your choice: "))
    
    if choice == 1:
        file_path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            save_path = filedialog.askdirectory(title="Select the folder to save the documents")
            if save_path:
                print("1. Generate multiple documents")
                print("2. Generate a single document")
                option = input("Enter your choice (1/2): ")
                if option in ['1', '2']:
                    process_files(file_path, save_path, option)
                else:
                    print("Invalid option. Please enter 1 or 2.")
            else:
                print("No folder selected")
        else:
            print("No file selected")
            
    elif choice == 2:
        print("Exiting...")
        exit()
    else:
        print("Invalid choice. Please enter 1 or 2.")
    
    root.destroy()  # Close the Tkinter window after use

if __name__ == "__main__":
    main()
