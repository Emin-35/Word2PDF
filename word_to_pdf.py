import os
import threading
from win32com import client
import pythoncom

# Directory paths
word_folder = os.path.join(os.getcwd(), 'Word Files')
pdf_folder = os.path.join(os.getcwd(), 'PDF Files')

# Create Word folder if it doesn't exist
if not os.path.exists(word_folder):
    os.makedirs(word_folder)

# Create PDF folder if it doesn't exist
if not os.path.exists(pdf_folder):
    os.makedirs(pdf_folder)

# Function to convert Word file to PDF
def convert_to_pdf(word_file, pdf_file):
    pythoncom.CoInitialize()

    word = client.DispatchEx("Word.Application")
    word.Visible = False
    doc = None

    try:
        doc = word.Documents.Open(word_file)
        doc.SaveAs(pdf_file, FileFormat=17)  # 17 is the PDF file format
    except Exception as e:
        print(f"Error converting '{word_file}' to PDF: {e}")
    finally:
        if doc:
            doc.Close()
        word.Quit()
        pythoncom.CoUninitialize()

# Function to convert a range of Word files to PDF using threads
def convert_range_to_pdf(thread_id, total_threads, word_files):
    print(f"Thread {thread_id} working...")
    total_files = len(word_files)
    files_per_thread = total_files // total_threads

    start_range = (thread_id - 1) * files_per_thread
    end_range = thread_id * files_per_thread

    for i in range(start_range, end_range):
        if i < total_files:
            filename = word_files[i]
            word_filepath = os.path.join(word_folder, filename)

            # PDF file path (in the PDF folder)
            pdf_filename = os.path.splitext(filename)[0] + '.pdf'
            pdf_filepath = os.path.join(pdf_folder, pdf_filename)

            # Convert Word file to PDF
            convert_to_pdf(word_filepath, pdf_filepath)

            print(f"Thread {thread_id}: Conversion completed for '{filename}'. PDF file saved at: {pdf_filepath}")
        else:
            print(f"Thread {thread_id}: No more files to process. Exiting.")

    print(f"Thread {thread_id} finished.")

# Determine the total number of files dynamically
word_files = [f for f in os.listdir(word_folder) if f.lower().endswith(('.doc', '.docx'))]
total_files = len(word_files)

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
    thread = threading.Thread(target=convert_range_to_pdf, args=(start_range, end_range, word_files))
    threads.append(thread)
    thread.start()

# Wait for all threads to complete
for thread in threads:
    thread.join()

print("All conversions completed.")
