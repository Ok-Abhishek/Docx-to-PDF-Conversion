import os
from docx2pdf import convert

def convert_docx_to_pdf(source_folder, destination_folder):
    # Check if the destination folder exists; if not, create it
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    # Loop through all files in the source folder
    for filename in os.listdir(source_folder):
        if filename.endswith(".docx"):  # Only consider .docx files
            # Full path of the docx file
            docx_path = os.path.join(source_folder, filename)
            
            # Create the destination pdf file path
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_path = os.path.join(destination_folder, pdf_filename)
            
            # Convert the .docx to .pdf
            convert(docx_path, pdf_path)
            print(f"Converted: {filename} to {pdf_filename}")

# Usage
source_folder = "C:/Users/abhis/Downloads/03.10.24 Updated BOS DEPARTMENT OF COMPUTER SCIENCE ENGINEERING-20241004T165214Z-001/03.10.24 Updated BOS DEPARTMENT OF COMPUTER SCIENCE ENGINEERING/2024-25"  # Folder containing .docx files
destination_folder = "C:/Users/abhis/Downloads/03.10.24 Updated BOS DEPARTMENT OF COMPUTER SCIENCE ENGINEERING-20241004T165214Z-001/03.10.24 Updated BOS DEPARTMENT OF COMPUTER SCIENCE ENGINEERING/2024-25"  # Folder to save converted .pdf files

convert_docx_to_pdf(source_folder, destination_folder)
