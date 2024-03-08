import os
from fpdf import FPDF

# Directory path for image files and output PDF
image_folder = os.path.join(os.getcwd(), 'Image Files')
pdf_folder = os.path.join(os.getcwd(), 'PDF Files')

# Create PDF folder if it doesn't exist
if not os.path.exists(image_folder):
    os.makedirs(image_folder)

# Create PDF folder if it doesn't exist
if not os.path.exists(pdf_folder):
    os.makedirs(pdf_folder)

# Function to convert image files to PDF
def convert_images_to_pdf(image_folder, output_pdf_file):

    # Get a list of image files in the folder
    image_files = [f for f in os.listdir(image_folder) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif'))]

    for image_file in image_files:
        image_path = os.path.join(image_folder, image_file)

        # Create a PDF for each image
        pdf = FPDF()
        pdf.add_page()

        pdf.image(image_path, x=0, y=0, w=212)

        pdf_output_file = os.path.join(pdf_folder, os.path.splitext(image_file)[0] + '.pdf')
        pdf.output(pdf_output_file, 'F')
        print(f"PDF file saved at: {output_pdf_file}")

# Provide the complete file path including the filename
output_pdf_file = os.path.join(pdf_folder, 'output.pdf')

# Call the function with the correct file path
convert_images_to_pdf(image_folder, output_pdf_file)
