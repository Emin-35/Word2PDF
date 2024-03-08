import os
import threading
from pdf2docx import Converter

# Directory paths
pdf_folder = os.path.join(os.getcwd(), 'PDF Files')
word_folder = os.path.join(os.getcwd(), 'Word Files')

# Create Word folder if it doesn't exist
if not os.path.exists(pdf_folder):
    os.makedirs(pdf_folder)

# Create Word folder if it doesn't exist
if not os.path.exists(word_folder):
    os.makedirs(word_folder)

# Function to convert PDF file to Word
def convert_to_word(pdf_file, word_file):
    try:
        # Create a Converter object
        cv = Converter(pdf_file)

        # Convert PDF to Word
        cv.convert(word_file, start=0, end=None)

        # Close the converter
        cv.close()
        print(f"Conversion completed for '{pdf_file}'. Word file saved at: {word_file}")
    except Exception as e:
        print(f"Error converting '{pdf_file}' to Word: {e}")


# Function to convert a range of PDF files to Word using threads
def convert_range_to_word(thread_id, total_threads, pdf_files):
    print(f"Thread {thread_id} working...")
    total_files = len(pdf_files)
    files_per_thread = total_files // total_threads

    start_range = (thread_id - 1) * files_per_thread
    end_range = thread_id * files_per_thread

    for i in range(start_range, end_range):
        if i < total_files:
            pdf_filename = pdf_files[i]
            pdf_filepath = os.path.join(pdf_folder, pdf_filename)

            # Word file path (in the Word folder)
            word_filename = os.path.splitext(pdf_filename)[0] + '.docx'
            word_filepath = os.path.join(word_folder, word_filename)

            # Convert PDF to Word
            convert_to_word(pdf_filepath, word_filepath)

            print(f"Thread {thread_id}: Conversion completed for '{pdf_filename}'. Word file saved at: {word_filepath}")
        else:
            print(f"Thread {thread_id}: No more files to process. Exiting.")

    print(f"Thread {thread_id} finished.")


# Get a list of PDF files in the folder
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]

# Determine the total number of files dynamically
total_files = len(pdf_files)

# Calculate the number of threads
if total_files < 10:
    # If the number of files is less than 10, use the actual file count as the number of threads
    threads_count = total_files
else:
    # Use a maximum thread count (e.g., 20) or a percentage of total_files, whichever is bigger
    threads_count = int(max(20, total_files * 0.1))

# Create and start threads
threads = []
files_per_thread = total_files // threads_count

for i in range(threads_count):
    start_range = i * files_per_thread + 1
    end_range = min((i + 1) * files_per_thread, total_files)
    thread = threading.Thread(target=convert_range_to_word, args=(start_range, end_range, pdf_files))
    threads.append(thread)
    thread.start()

# Wait for all threads to complete
for thread in threads:
    thread.join()

print("All conversions completed.")
