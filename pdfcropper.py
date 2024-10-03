import fitz  # PyMuPDF
import os

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



# Task 2: Process all PDFs in a folder
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

# Define input and output folders
input_folder = "./input"
output_folder = "./output"

# Process all PDFs in the input folder
process_pdfs(input_folder, output_folder)
