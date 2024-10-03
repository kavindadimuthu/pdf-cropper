import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def extract_slides_from_pdf(input_pdf, output_pdf, rows, columns, left_margin, right_margin, top_margin, bottom_margin, horizontal_gap, vertical_gap):
    doc = fitz.open(input_pdf)  # Open the input PDF
    output_doc = fitz.open()    # Create a new PDF for the output
    slide_counter = 1

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)

        # Get the page size
        page_width = page.rect.width
        page_height = page.rect.height

        # Calculate the usable space
        usable_width = page_width - (left_margin + right_margin) - ((columns - 1) * horizontal_gap)
        usable_height = page_height - (top_margin + bottom_margin) - ((rows - 1) * vertical_gap)

        # Calculate each slide's width and height based on user-defined rows and columns
        slide_width = usable_width / columns
        slide_height = usable_height / rows

        # Loop through rows and columns to extract each slide
        for row in range(rows):
            for col in range(columns):
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
def process_pdfs(input_folder, output_folder, rows, columns, left_margin, right_margin, top_margin, bottom_margin, horizontal_gap, vertical_gap):
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
                rows=rows,
                columns=columns,
                left_margin=left_margin, 
                right_margin=right_margin, 
                top_margin=top_margin, 
                bottom_margin=bottom_margin, 
                horizontal_gap=horizontal_gap, 
                vertical_gap=vertical_gap
            )
            
            print(f"Saved output to {output_pdf}\n")  # Confirm saving
    messagebox.showinfo("Success", "PDFs processed successfully!")




def show_user_manual():
    manual_text = """
    User Manual:

    1. Select Input Folder:
       - Click 'Browse Input Folder' to select the folder with your input PDFs.

    2. Set Margins and Gaps:
       - Adjust the Left, Right, Top, Bottom margins and gaps between slides.

    3. Set Number of Slides:
       - Define how many Rows and Columns of slides you want to extract per page.

    4. Select Output Folder:
       - Click 'Browse Output Folder' to select where the output PDFs will be saved.

    5. Run the Program:
       - Click 'Process PDFs' to start extracting slides into separate pages.

    Output:
    - Extracted slides will be saved in the selected output folder.
    
    Illustration:
    - The diagram on the right represents a PDF page with slides, margins, and gaps. 
      This helps you understand how the values you input will affect the slides on the page.
    """

    manual_frame = tk.Frame(root, padx=10, pady=10)
    manual_frame.grid(row=0, column=2, rowspan=8, padx=10, sticky='n')

    manual_label = tk.Label(manual_frame, text=manual_text, justify='left', anchor='w')
    manual_label.pack(fill=tk.BOTH)


def create_illustration(canvas, horizontal_gap=31, vertical_gap=50):
    # Clear previous drawings
    canvas.delete("all")
    
    # Drawing a representation of a PDF page with margins, gaps, and slides
    canvas.create_rectangle(50, 50, 300, 450, outline='black', width=2)  # PDF page outline

    # Margins
    canvas.create_rectangle(75, 75, 275, 425, outline='blue', width=1, dash=(5, 2))  # Margins area
    canvas.create_text(60, 90, text="Left Margin", anchor="e", fill="blue", angle=90)  # Left Margin (Vertical)
    canvas.create_text(290, 90, text="Right Margin", anchor="e", fill="blue", angle=90)  # Right Margin (Vertical)
    canvas.create_text(160, 35, text="Top Margin", fill="blue")
    canvas.create_text(160, 440, text="Bottom Margin", fill="blue")
    
    # Calculate slide dimensions based on gaps
    slide_width = (275 - 75 - horizontal_gap) / 2  # Adjusting for horizontal gaps
    slide_height = (425 - 75 - vertical_gap * 2) / 2  # Adjusting for vertical gaps

    # Slide 1 (Top-Left)
    canvas.create_rectangle(75, 75, 75 + slide_width, 75 + slide_height, outline='green', width=1)  # Top-left slide
    canvas.create_text(75 + slide_width / 2, 75 + slide_height / 2, text="Slide 1", fill="green")

    # Slide 2 (Top-Right)
    canvas.create_rectangle(75 + slide_width + horizontal_gap, 75, 
                            75 + 2 * slide_width + horizontal_gap, 75 + slide_height, outline='green', width=1)  # Top-right slide
    canvas.create_text(75 + slide_width + horizontal_gap + slide_width / 2, 75 + slide_height / 2, text="Slide 2", fill="green")

    # Slide 3 (Bottom-Left)
    canvas.create_rectangle(75, 75 + slide_height + vertical_gap, 
                            75 + slide_width, 75 + 2 * slide_height + vertical_gap, outline='green', width=1)  # Bottom-left slide
    canvas.create_text(75 + slide_width / 2, 75 + slide_height + vertical_gap + slide_height / 2, text="Slide 3", fill="green")

    # Slide 4 (Bottom-Right)
    canvas.create_rectangle(75 + slide_width + horizontal_gap, 75 + slide_height + vertical_gap, 
                            75 + 2 * slide_width + horizontal_gap, 75 + 2 * slide_height + vertical_gap, outline='green', width=1)  # Bottom-right slide
    canvas.create_text(75 + slide_width + horizontal_gap + slide_width / 2, 
                       75 + slide_height + vertical_gap + slide_height / 2, text="Slide 4", fill="green")

    # Vertical Gap
    canvas.create_line(75, 75 + slide_height + vertical_gap, 275, 75 + slide_height + vertical_gap, fill='orange', dash=(4, 2))  # Horizontal gap line
    canvas.create_text(120, 75 + slide_height + vertical_gap - 20 , text="Vertical Gap", fill="orange")

    # Horizontal Gap
    canvas.create_line(75 + slide_width + horizontal_gap, 75, 75 + slide_width + horizontal_gap, 425, fill='orange', dash=(4, 2))  # Vertical gap line
    canvas.create_text(50 + slide_width + horizontal_gap + 10, 270, text="Horizontal Gap", fill="orange",  anchor="e", angle=90)


def create_gui():
    global root
    root = tk.Tk()
    root.title("PDF Slide Extractor")

    # Input and Output Folder Selection
    tk.Label(root, text="Input Folder:").grid(row=0, column=0, sticky='e')
    input_folder_entry = tk.Entry(root)
    input_folder_entry.grid(row=0, column=1)
    tk.Button(root, text="Browse Input Folder", command=lambda: select_folder(input_folder_entry)).grid(row=0, column=2)

    tk.Label(root, text="Output Folder:").grid(row=1, column=0, sticky='e')
    output_folder_entry = tk.Entry(root)
    output_folder_entry.grid(row=1, column=1)
    tk.Button(root, text="Browse Output Folder", command=lambda: select_folder(output_folder_entry)).grid(row=1, column=2)

    # Margins and Gap Settings
    tk.Label(root, text="Left Margin:").grid(row=2, column=0, sticky='e')
    left_margin_entry = tk.Entry(root)
    left_margin_entry.insert(0, "43")
    left_margin_entry.grid(row=2, column=1)

    tk.Label(root, text="Right Margin:").grid(row=3, column=0, sticky='e')
    right_margin_entry = tk.Entry(root)
    right_margin_entry.insert(0, "43")
    right_margin_entry.grid(row=3, column=1)

    tk.Label(root, text="Top Margin:").grid(row=4, column=0, sticky='e')
    top_margin_entry = tk.Entry(root)
    top_margin_entry.insert(0, "72")
    top_margin_entry.grid(row=4, column=1)

    tk.Label(root, text="Bottom Margin:").grid(row=5, column=0, sticky='e')
    bottom_margin_entry = tk.Entry(root)
    bottom_margin_entry.insert(0, "72")
    bottom_margin_entry.grid(row=5, column=1)

    tk.Label(root, text="Horizontal Gap:").grid(row=6, column=0, sticky='e')
    horizontal_gap_entry = tk.Entry(root)
    horizontal_gap_entry.insert(0, "31")
    horizontal_gap_entry.grid(row=6, column=1)

    tk.Label(root, text="Vertical Gap:").grid(row=7, column=0, sticky='e')
    vertical_gap_entry = tk.Entry(root)
    vertical_gap_entry.insert(0, "50")
    vertical_gap_entry.grid(row=7, column=1)

    # Slide Rows and Columns
    tk.Label(root, text="Rows (Slides per Page):").grid(row=8, column=0, sticky='e')
    rows_entry = tk.Entry(root)
    rows_entry.insert(0, "3")
    rows_entry.grid(row=8, column=1)

    tk.Label(root, text="Columns (Slides per Page):").grid(row=9, column=0, sticky='e')
    columns_entry = tk.Entry(root)
    columns_entry.insert(0, "2")
    columns_entry.grid(row=9, column=1)

    # User Manual
    show_user_manual()

    # Illustration Canvas
    illustration_canvas = tk.Canvas(root, width=400, height=500, bg='white')
    illustration_canvas.grid(row=0, column=3, rowspan=10, padx=20)
    create_illustration(illustration_canvas)

    # Process Button
    tk.Button(root, text="Process PDFs", command=lambda: process_pdfs(input_folder_entry, output_folder_entry, 
                                                                     left_margin_entry, right_margin_entry, 
                                                                     top_margin_entry, bottom_margin_entry, 
                                                                     horizontal_gap_entry, vertical_gap_entry, 
                                                                     rows_entry, columns_entry)).grid(row=10, column=0, columnspan=2, pady=10)

    root.mainloop()



# Call the function to create the GUI
create_gui()
