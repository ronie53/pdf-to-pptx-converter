import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt
import os

def pdf_to_pptx(pdf_file, output_file):
    # Check if the PDF file exists
    if not os.path.exists(pdf_file):
        print(f"Error: File '{pdf_file}' not found.")
        return
    
    try:
        # Open the PDF
        with pdfplumber.open(pdf_file) as pdf:
            # Create a new PowerPoint presentation
            presentation = Presentation()
            
            # Loop through each page in the PDF
            for page_number, page in enumerate(pdf.pages):
                # Extract text from the page
                text = page.extract_text()
                
                # Add a new slide to the presentation
                slide = presentation.slides.add_slide(presentation.slide_layouts[1])
                
                # Add title (page number)
                title = slide.shapes.title
                title.text = f"Page {page_number + 1}"
                
                # Add text content
                content = slide.placeholders[1]
                text_frame = content.text_frame
                text_frame.text = text if text else "No text found on this page"
                
            # Save the presentation
            presentation.save(output_file)
            print(f"✓ Conversion complete! Saved as {output_file}")
    except Exception as e:
        print(f"Error processing PDF: {e}")

# Run the converter
if __name__ == "__main__":
    pdf_file = input("Enter full path to PDF file (with .pdf): ")
    output_file = pdf_file.replace('.pdf', '.pptx')
    
    pdf_to_pptx(pdf_file, output_file)




