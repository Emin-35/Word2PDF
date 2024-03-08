from subprocess import run

# Function to convert Word files to PDF
def convert_word_to_pdf():
    run(["python", "word_to_pdf.py"])

# Function to convert Word files to PDF
def convert_pdf_to_word():
    run(["python", "pdf_to_word.py"])

# Function to convert images to PDF
def convert_images_to_pdf():
    run(["python", "img_to_pdf.py"])

# Function to merge multiple PDFs
def merge_pdf():
    run(["python", "merge_pdf.py"])

# Menu to choose functionality
def main():
    while True:
        print("Choose an option:")
        print("1. Convert Word files to PDF")
        print("2. Convert PDF files to Word")
        print("3. Convert images to PDF")
        print("4. Merge multiple PDFs")
        print("5. Exit")

        choice = input("Enter your choice (1, 2.. or 5): ")

        if choice == "1":
            convert_word_to_pdf()
        elif choice == "2":
            convert_pdf_to_word()
        elif choice == "3":
            convert_images_to_pdf()
        elif choice == "4":
            merge_pdf()
        elif choice == "5":
            break
        else:
            print("Invalid choice. Please enter 1, 2.. or 5")

if __name__ == "__main__":
    main()
