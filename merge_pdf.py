import os
from PyPDF2 import PdfMerger

# Directory path for input PDFs and output PDF
input_pdf_folder = os.path.join(os.getcwd(), 'PDF Files')
output_pdf_folder = os.path.join(os.getcwd(), 'Merge PDFs')
output_pdf_file = os.path.join(output_pdf_folder, 'Merged.pdf')

# Create Merge PDFs folder if it doesn't exist
if not os.path.exists(input_pdf_folder):
    os.makedirs(input_pdf_folder)

# Create PDF folder if it doesn't exist
if not os.path.exists(output_pdf_folder):
    os.makedirs(output_pdf_folder)

# Get a list of PDF files in the folder
pdf_files = [f for f in os.listdir(input_pdf_folder) if f.lower().endswith('.pdf')]

# Function to merge PDFs
def merge_pdfs(input_folder, output_file):
    # Create an instance of PdfFileMerger
    pdf_merger = PdfMerger()

    for pdf_file in pdf_files:
        # Declare the path
        pdf_path = os.path.join(input_folder, pdf_file)

        # Merge all pages of the current PDF
        pdf_merger.append(pdf_path)
    
    # Write the merged PDF to the output file
    with open(output_file, 'wb') as merged_pdf:
        pdf_merger.write(merged_pdf)

    print(f"Merged PDF file saved at: {output_file}")

# Call the function
merge_pdfs(input_pdf_folder, output_pdf_file)
