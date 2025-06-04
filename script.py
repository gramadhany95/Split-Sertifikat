import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import PyPDF2
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import re
from tqdm import tqdm  # Import tqdm for the progress bar

# Function to open file dialog for selecting multiple PDF files
def select_pdf_files():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF Files", "*.pdf")])
    return file_paths

# Function to select output folder for saving the split files
def select_output_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title="Select Output Folder")
    return folder_path

# Function to prompt the user for a custom name for the certificates
def get_custom_name():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    custom_name = simpledialog.askstring("Input", "Enter custom name for certificates:")
    if not custom_name:
        custom_name = "Certificate"  # Default name if nothing is provided
    return custom_name

# Select PDF files and output folder using file dialogs
input_pdfs = select_pdf_files()
output_folder = select_output_folder()

# Get custom name for the certificate
custom_name = get_custom_name()

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Initialize a list to store extracted data
data = []

# Function to extract ID from the text using regular expression
def extract_id(text):
    # Regular expression to find the ID format: "ID: 1105210003"
    match = re.search(r"ID:\s*(\d+)", text)
    if match:
        return match.group(1)  # Extract the numeric ID
    return "Unknown"

# Function to extract Name from the text
def extract_name(text):
    # Look for the name by searching between "We hereby confirm that" and "ID:"
    name_start = text.find("We hereby confirm that") + len("We hereby confirm that")
    name_end = text.find("ID:")
    if name_start != -1 and name_end != -1:
        return text[name_start:name_end].strip()
    return "Unknown"

# Count the total number of pages across all selected PDFs
total_pages = 0
for pdf_file in input_pdfs:
    reader = PdfReader(pdf_file)
    total_pages += len(reader.pages)

# Initialize a progress bar with the total number of pages
with tqdm(total=total_pages, desc="Processing Pages", unit="page") as pbar:
    pages_read = 0  # Counter for pages read
    # Process each selected PDF file
    for pdf_file in input_pdfs:
        reader = PdfReader(pdf_file)
        
        # Process each page in the PDF
        for i, page in enumerate(reader.pages):
            writer = PdfWriter()

            # Extract text from the page
            text = page.extract_text()

            # Extract ID and Name from the text
            user_id = extract_id(text)
            name = extract_name(text)

            # Prepare filename based on the custom name, ID, and Name
            filename = f"{custom_name}_{user_id}_{name}.pdf"

            # Add the page to the writer
            writer.add_page(page)

            # Save the page as a new PDF file
            with open(os.path.join(output_folder, filename), "wb") as output_pdf:
                writer.write(output_pdf)

            # Append the extracted data to the list
            data.append({
                "ID": user_id,
                "Name": name,
                "Filename": filename,
                "Source PDF": os.path.basename(pdf_file)
            })

            # Update the progress bar after each page is processed
            pbar.update(1)
            pages_read += 1  # Increment the pages read counter

# Create a DataFrame to save as an Excel file
df = pd.DataFrame(data)

# Save the data to an Excel file
spreadsheet_output = os.path.join(output_folder, "extracted_data.xlsx")
df.to_excel(spreadsheet_output, index=False)

# Check if the number of extracted files matches the total page count
extracted_files_count = len(data)
if extracted_files_count == total_pages:
    messagebox.showinfo("Extraction Complete", 
                        f"All files have been successfully extracted!\nTotal Pages Read: {total_pages}\nTotal Pages Processed: {extracted_files_count}")
else:
    messagebox.showwarning("Extraction Incomplete", 
                           f"Extraction may be incomplete.\nTotal Pages Read: {total_pages}\nTotal Pages Processed: {extracted_files_count}")

print(f"PDF pages have been split and saved in {output_folder}")
print(f"Extracted data has been saved in {spreadsheet_output}")
