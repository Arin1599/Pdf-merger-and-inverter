# PDF Wizard

A powerful Python-based graphical user interface (GUI) tool for comprehensive PDF management. PDF Wizard allows users to merge multiple PDF files, reverse page order, sign documents, and features an advanced visual page editor with drag-and-drop functionality for rearranging pages and inserting images.

## Features:

### Core Features:
- **Merge PDFs**: Select and combine multiple PDFs into a single document with customizable order
- **Reverse PDF**: Reverse the pages of an individual PDF and save it as a new file
- **Sign PDF**: Add digital signatures to any PDF page with visual drag-and-drop placement
- **Visual Page Editor**: Advanced editor window with thumbnail previews of all pages
- **Drag & Drop Reordering**: Click and drag page thumbnails to rearrange pages visually
- **Insert Images**: Add JPG, PNG, GIF, BMP, or TIFF images as new PDF pages
- **Insert PDFs**: Add pages from other PDF files at any position
- **Delete Pages**: Remove unwanted pages from your PDF
- **Real-time Preview**: See thumbnail previews of all pages before saving

### User Interface:
- **Intuitive GUI**: Clean, modern interface with easy file selection and operations
- **Progress Feedback**: Real-time progress bar while merging PDFs
- **Visual Feedback**: Highlighted selections and hover effects for better user experience
- **Insert Buttons**: Convenient "+" buttons between pages for quick content insertion
- **Page Numbering**: Clear page number badges on each thumbnail
- **Multi-Page Signing**: Add signatures to multiple pages in a single session

## Technologies Used:
- **Python 3.x**
- **Tkinter**: For creating the GUI
- **PyPDF2**: For PDF manipulation (merging, reversing, and page operations)
- **PyMuPDF (fitz)**: For high-quality PDF rendering and thumbnail generation
- **Pillow (PIL)**: For image processing and conversion
- **os, ctypes, tempfile**: For file handling and system operations

## Installation:

1. Clone the repository:
```bash
git clone https://github.com/Arin1599/Pdf-merger-and-inverter.git
cd Pdf-merger-and-inverter
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python pdf_inverter.py
```

## How to Use:

### Basic Operations:

1. **Merge PDFs**: 
   - Click "Add PDFs" to select multiple PDF files
   - Use "Move Up" / "Move Down" to reorder files
   - Click "Merge" to combine them into a single PDF

2. **Reverse a PDF**: 
   - Click "Browse" in the Reverse PDF section
   - Select a PDF file
   - Click "Reverse" to save a reversed copy

3. **Sign a PDF**:
   - Click "Browse PDF" in the Sign PDF section
   - Select the PDF you want to sign
   - Click "Open Signer" to launch the signature tool

### Advanced Page Editor:

1. **Open Page Editor**:
   - Add PDFs to the merge list
   - Click "Edit Pages" to open the visual editor

2. **Rearrange Pages**:
   - Click and drag any page thumbnail to a new position
   - Visual feedback shows where the page will be dropped
   - Release to reorder

3. **Insert Content**:
   - Click any "+" button between pages
   - Choose "Insert Image" or "Insert PDF"
   - Select your file to add it at that position

4. **Delete Pages**:
   - Click on a page thumbnail to select it
   - Click "Delete Selected" to remove the page

5. **Save Your Work**:
   - Click "Save PDF" to export your edited document
   - Choose a location and filename

### PDF Signature Tool:

1. **Select Signature Image**:
   - Click "Select Signature" button
   - Choose an image file (PNG, JPG, etc.) containing your signature

2. **Place Signature**:
   - Click and drag on the PDF page to create a signature box
   - Adjust the size and position as needed
   - Click "Add to This Page" to confirm placement

3. **Navigate Pages**:
   - Use "Previous" and "Next" buttons to move between pages
   - Add signatures to multiple pages as needed
   - Status bar shows which pages have signatures

4. **Save Signed PDF**:
   - Click "Save Signed PDF" when done
   - Choose output location and filename
   - Your signed PDF is ready!

## Supported File Formats:
- **PDF**: .pdf
- **Images**: .png, .jpg, .jpeg, .gif, .bmp, .tiff

## Requirements:
- Python 3.7 or higher
- Windows OS (for console minimization feature)
- See `requirements.txt` for complete package dependencies
