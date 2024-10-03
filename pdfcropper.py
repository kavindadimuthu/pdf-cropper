import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


def extract_slides_from_pdf(input_pdf, output_pdf, left_margin, right_margin, top_margin, bottom_margin, horizontal_gap, vertical_gap):
    doc = fitz.open(input_pdf)  # Open the input PDF
    output_doc = fitz.open()    # Create a new PDF for the output
    slide_counter = 1

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)

        # Get the page size
        page_width = page.rect.width
        page_height = page.rect.height

        # Calculate the usable space
        usable_width = page_width - (left_margin + right_margin) - horizontal_gap
        usable_height = page_height - (top_margin + bottom_margin) - (2 * vertical_gap)

        # Calculate each slide's width and height
        slide_width = usable_width / 2
        slide_height = usable_height / 3

        # Loop through 3 rows and 2 columns to extract each slide
        for row in range(3):
            for col in range(2):
                # Define the rectangle area of the slide (left, top, right, bottom) considering margins and gaps
                left = left_margin + col * (slide_width + horizontal_gap)
                top = top_margin + row * (slide_height + vertical_gap)
                right = left + slide_width
                bottom = top + slide_height

                # Define the clip rectangle
                clip = fitz.Rect(left, top, right, bottom)

                # Create a new page for the extracted slide
                slide_page = output_doc.new_page(width=clip.width, height=clip.height)

                # Copy the slide from the original page
                slide_page.show_pdf_page(fitz.Rect(0, 0, clip.width, clip.height), doc, page_num, clip=clip)

                slide_counter += 1

    # Save the final PDF
    output_doc.save(output_pdf)
    output_doc.close()
    print(f"Slides extracted and saved into '{output_pdf}'")


# Your PDF processing function goes here
def process_pdfs(input_folder, output_folder):
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Get all PDF files from the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            input_pdf = os.path.join(input_folder, filename)  # Full input PDF path
            output_pdf = os.path.join(output_folder, f"output_{filename}")  # Define the output file path
            
            print(f"Processing {input_pdf}...")  # For debugging
            
            # Call the function for each file
            extract_slides_from_pdf(
                input_pdf, 
                output_pdf, 
                left_margin=42, 
                right_margin=42, 
                top_margin=70, 
                bottom_margin=70, 
                horizontal_gap=29, 
                vertical_gap=46
            )
            
            print(f"Saved output to {output_pdf}\n")  # Confirm saving
    messagebox.showinfo("Success", "PDFs processed successfully!")

# GUI for selecting folders and triggering processing
def create_gui():
    def browse_input_folder():
        folder = filedialog.askdirectory()
        input_folder_entry.delete(0, tk.END)
        input_folder_entry.insert(0, folder)
    
    def browse_output_folder():
        folder = filedialog.askdirectory()
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder)
    
    def start_processing():
        input_folder = input_folder_entry.get()
        output_folder = output_folder_entry.get()
        
        if not input_folder or not output_folder:
            messagebox.showwarning("Input Error", "Please select both input and output folders.")
            return
        
        process_pdfs(input_folder, output_folder)

    # Main window
    root = tk.Tk()
    root.title("PDF Slide Extractor")

    # Input folder selection
    tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    input_folder_entry = tk.Entry(root, width=40)
    input_folder_entry.grid(row=0, column=1, padx=10, pady=10)
    browse_input_button = tk.Button(root, text="Browse", command=browse_input_folder)
    browse_input_button.grid(row=0, column=2, padx=10, pady=10)

    # Output folder selection
    tk.Label(root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    output_folder_entry = tk.Entry(root, width=40)
    output_folder_entry.grid(row=1, column=1, padx=10, pady=10)
    browse_output_button = tk.Button(root, text="Browse", command=browse_output_folder)
    browse_output_button.grid(row=1, column=2, padx=10, pady=10)

    # Start button
    start_button = ttk.Button(root, text="Start Processing", command=start_processing)
    start_button.grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()

# Call the function to create the GUI
create_gui()
