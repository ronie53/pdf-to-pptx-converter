# PDF to PPTX Converter

A Python script that converts PDF documents into PowerPoint presentations (.pptx format).

## Features

- **PDF to PPTX Conversion**: Extracts content from PDF files and creates a PowerPoint presentation
- **Text Extraction**: Automatically extracts text from PDF pages
- **Error Handling**: Gracefully handles invalid or corrupted PDF files
- **Simple Interface**: Easy-to-use command-line interface

## Requirements

- Python 3.7+
- `pdfplumber` - for PDF parsing and text extraction
- `python-pptx` - for creating PowerPoint presentations

## Installation

1. Clone or download this repository
2. Install the required dependencies:
   ```bash
   pip install pdfplumber python-pptx
   ```

## Usage

1. Run the script:
   ```bash
   python main.py
   ```

2. When prompted, enter the **full path** to your PDF file:
   ```
   Enter full path to PDF file (with .pdf): C:\Users\YourUsername\Desktop\document.pdf
   ```

3. The script will generate a PowerPoint file with the same name as the PDF (but with a `.pptx` extension) in the same directory as the PDF.

### Example

```bash
$ python main.py
Enter full path to PDF file (with .pdf): C:\Users\Niessen\Desktop\sample.pdf
PDF successfully converted to: C:\Users\Niessen\Desktop\sample.pptx
```

## Troubleshooting

### "No /Root object! - Is this really a PDF?" Error

**Cause**: The file is not a valid PDF document. This typically happens when:
- The file is corrupted or incomplete
- The file is not actually a PDF (e.g., a ZIP file renamed with a `.pdf` extension)
- The file format is unsupported

**Solution**:
1. Verify the file is a real PDF by opening it in a PDF viewer (Adobe Acrobat, Chrome, Edge, etc.)
2. Check the file properties: Right-click > Properties > General tab should show "PDF File" as the type
3. Try converting a different, known-good PDF file
4. If the file is corrupted, obtain a fresh copy from the source

### File Not Found Error

**Cause**: The system cannot locate the PDF file at the provided path

**Solution**:
1. Use the full absolute path (e.g., `C:\Users\YourUsername\Desktop\file.pdf`)
2. Verify the file exists by checking in File Explorer
3. Ensure the path doesn't contain typos
4. If the path has spaces, include it exactly as shown in the file properties

## How It Works

1. Opens and reads the PDF file using `pdfplumber`
2. Extracts text from each page
3. Creates a new PowerPoint presentation
4. Adds each PDF page's content as a slide in the presentation
5. Saves the presentation with a `.pptx` extension

## Limitations

- Requires a valid PDF file with an actual PDF structure (not just a renamed file)
- Text extraction quality depends on PDF encoding and formatting
- Complex PDFs with images, tables, or special formatting may require manual formatting adjustments in PowerPoint
- Scanned PDFs (image-based) may not extract text properly

## Output

The generated PowerPoint file will be created in the same directory as the input PDF file and will have the same name with a `.pptx` extension.

## License

This project is provided as-is for personal use.
