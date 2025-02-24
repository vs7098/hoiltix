import PyPDF2

def is_blank_page(page):
    # Check if the page contains only whitespace
    return not page.extract_text().strip()

def remove_blank_and_crop(input_pdf, output_pdf, crop_width):
    # Open the input PDF file
    with open(input_pdf, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()

        # Loop through all pages
        for page_number in range(len(reader.pages)):
            page = reader.pages[page_number]
            
            # Add non-blank pages
            if not is_blank_page(page):
                # Get the original dimensions of the page
                media_box = page.mediabox
                original_width = media_box.width
                original_height = media_box.height

                # Crop 4 inches from the right side of the page
                media_box.upper_right = (original_width - crop_width, original_height)

                # Add the cropped page to the writer
                writer.add_page(page)
        
        # Write the result to the output PDF file
        with open(output_pdf, 'wb') as output_file:
            writer.write(output_file)

# Example usage
input_pdf = "Batch3.pdf"
output_pdf = "GA-1000-3.pdf"
crop_width = 102  # 4 inches (72 points per inch)
remove_blank_and_crop(input_pdf, output_pdf, crop_width)

