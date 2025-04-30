import os
import re
import PyPDF2
import docx
import argparse
from pathlib import Path
import io
import json
import base64
from PIL import Image, ImageDraw, ImageFont
import google.generativeai as genai
import fitz  # PyMuPDF
import tempfile
import subprocess
import sys
import textwrap  # Add this import

# Configure the Generative AI SDK with your API key
genai.configure(api_key='AIzaSyAhtesttttutututuH_S-LqxBnkSg')

# Initialize the Gemini 2.0 Flash model
model = genai.GenerativeModel('gemini-2.0-flash')

def encode_image_to_base64(image):
    """
    Encode an image (PIL Image or path) to base64 string
    """
    try:
        if isinstance(image, str):  # It's a path
            with Image.open(image) as img:
                # Convert to RGB if needed
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Resize if too large
                max_size = (1600, 1600)
                if img.width > max_size[0] or img.height > max_size[1]:
                    img.thumbnail(max_size, Image.LANCZOS)
                
                buffered = io.BytesIO()
                img.save(buffered, format='JPEG', quality=95)
                return base64.b64encode(buffered.getvalue()).decode('utf-8')
        else:  # It's a PIL Image or bytes
            if isinstance(image, bytes):
                # Convert bytes to PIL Image
                img = Image.open(io.BytesIO(image))
            else:
                img = image
                
            # Convert to RGB if needed
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Resize if too large
            max_size = (1600, 1600)
            if img.width > max_size[0] or img.height > max_size[1]:
                img.thumbnail(max_size, Image.LANCZOS)
            
            buffered = io.BytesIO()
            img.save(buffered, format='JPEG', quality=95)
            return base64.b64encode(buffered.getvalue()).decode('utf-8')
    except Exception as e:
        print(f"Error encoding image: {e}")
        return None

def extract_text_with_gemini(image, document_type="document"):
    """
    Extract text from an image using Gemini API
    """
    try:
        # Get base64 encoded image
        if isinstance(image, str):  # Image path
            base64_image = encode_image_to_base64(image)
        else:  # PIL Image or bytes
            base64_image = encode_image_to_base64(image)
        
        if not base64_image:
            return "Error: Could not encode image"
        
        # Create prompt
        prompt_text = (
            f"Extract all text content from this {document_type} image. "
            "Preserve the layout and structure as much as possible. "
            "Return just the plain text content, with line breaks where appropriate."
        )
        
        # Prepare the image part
        image_part = {
            "mime_type": "image/jpeg", 
            "data": base64_image
        }
        
        # Generate content
        response = model.generate_content(
            contents=[prompt_text, image_part],
            generation_config=genai.GenerationConfig(
                temperature=0.1,
                response_mime_type="text/plain"
            )
        )
        
        # Return extracted text
        return response.text.strip()
        
    except Exception as e:
        print(f"Error extracting text with Gemini: {e}")
        return f"Error: {str(e)}"

def convert_docx_to_images(docx_path):
    """Convert DOCX directly to images for OCR processing"""
    images = []
    
    try:
        # Method 1: Try using MS Word via COM automation (Windows only)
        if sys.platform == "win32":
            try:
                import win32com.client
                import win32api
                
                print("Converting DOCX to PDF using MS Word COM automation...")
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                
                # Convert full path to format Word likes
                abs_path = os.path.abspath(docx_path)
                doc = word.Documents.Open(abs_path)
                
                # Save as PDF
                pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
                doc.Close()
                word.Quit()
                
                # Now convert PDF to images using PyMuPDF
                if os.path.exists(pdf_path):
                    doc = fitz.open(pdf_path)
                    for page_num in range(len(doc)):
                        page = doc[page_num]
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_data = pix.tobytes("png")
                        img = Image.open(io.BytesIO(img_data))
                        images.append(img)
                    doc.close()
                    return images
            except Exception as e:
                print(f"Error using Word automation: {e}")
        
        # Method 2: Try using docx2pdf
        try:
            from docx2pdf import convert
            print("Converting DOCX to PDF using docx2pdf...")
            pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
            convert(docx_path, pdf_path)
            
            if os.path.exists(pdf_path):
                doc = fitz.open(pdf_path)
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    images.append(img)
                doc.close()
                return images
        except ImportError:
            print("docx2pdf not installed. Trying alternative method...")
        
        # Method 3: Direct approach - render DOCX content using python-docx and Pillow
        print("Converting DOCX to images directly...")
        doc = docx.Document(docx_path)
        
        # Create image with the document content
        width, height = 1700, 2200  # A4 size at 200 DPI
        background_color = (255, 255, 255)  # White
        
        # Create one image per 40 paragraphs (approximate page)
        paragraphs_per_page = 40
        all_paragraphs = list(doc.paragraphs)
        
        for i in range(0, len(all_paragraphs), paragraphs_per_page):
            img = Image.new('RGB', (width, height), background_color)
            draw = ImageDraw.Draw(img)
            
            # Try to use a nice font, fall back to default if not available
            try:
                font = ImageFont.truetype("arial.ttf", 24)
            except IOError:
                font = ImageFont.load_default()
            
            y_position = 100
            
            # Add paragraphs to the image
            for para in all_paragraphs[i:i+paragraphs_per_page]:
                if para.text.strip():
                    # Wrap text to fit the image width
                    wrapped_text = textwrap.fill(para.text, width=70)
                    draw.text((100, y_position), wrapped_text, fill=(0, 0, 0), font=font)
                    y_position += 30 + wrapped_text.count('\n') * 30
                
                # If we're running out of space, move to next image
                if y_position > height - 100:
                    break
            
            images.append(img)
        
        return images
        
    except Exception as e:
        print(f"Error converting DOCX to images: {e}")
        return []

def extract_text_from_docx(docx_path):
    """
    Extract text from DOCX by converting pages to images and using OCR.
    """
    text = []
    
    try:
        # Get images from the DOCX file
        print(f"Converting DOCX to images: {docx_path}")
        images = convert_docx_to_images(docx_path)
        
        if not images:
            print("Falling back to text extraction only...")
            doc = docx.Document(docx_path)
            for para in doc.paragraphs:
                if para.text.strip():
                    text.append(para.text)
            
            # Extract text from tables
            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text for cell in row.cells]
                    text.append(" | ".join(row_text))
            
            return "\n\n".join(text)
        
        # Process each image with Gemini OCR
        for i, img in enumerate(images):
            text.append(f"\n----- Page {i + 1} -----\n")
            print(f"Processing page {i + 1} with Gemini OCR...")
            page_text = extract_text_with_gemini(img, "document page")
            text.append(page_text)
        
        return "\n".join(text)
    except Exception as e:
        return f"Error processing DOCX file: {str(e)}"

def extract_text_from_pdf_as_images(pdf_path):
    """Extract text from PDF by converting all pages to images and using OCR."""
    text = []
    
    try:
        # Open the PDF
        doc = fitz.open(pdf_path)
        
        # Process each page
        for page_num in range(len(doc)):
            # Add page number as a marker
            text.append(f"\n----- Page {page_num + 1} -----\n")
            
            # Render page to an image (higher resolution for better OCR)
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            
            # Convert to PIL Image
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # Use Gemini to extract text from image
            print(f"Processing page {page_num + 1} with Gemini OCR...")
            page_text = extract_text_with_gemini(img, "document page")
            text.append(page_text)
        
        doc.close()
        return "\n".join(text)
    except Exception as e:
        return f"Error extracting text from PDF as images: {str(e)}"

def check_if_pdf_has_text(pdf_path, page_num):
    """Check if a PDF page contains actual text or is image-based."""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        
        # If the page has less than 10 characters, it's likely an image
        return len(text.strip()) > 10

def extract_image_from_pdf_page(pdf_path, page_num):
    """Extract image from PDF page using PyMuPDF (fitz)."""
    try:
        # Open the PDF
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        
        # Render page to an image (higher resolution for better OCR)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        
        # Convert to PIL Image
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        
        doc.close()
        return img
    except Exception as e:
        print(f"Error extracting image from PDF page: {e}")
        return None

def extract_text_from_pdf(pdf_path):
    """
    Extract text from PDF, intelligently handling both text and image-based content.
    For mixed PDFs, extracts text where possible and uses OCR for image-based pages.
    """
    text = []
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Count how many pages have text vs. need OCR
            total_pages = len(pdf_reader.pages)
            text_pages = 0
            image_pages = 0
            
            for page_num in range(total_pages):
                if check_if_pdf_has_text(pdf_path, page_num):
                    text_pages += 1
                else:
                    image_pages += 1
            
            # If most pages are image-based (>75%), use the full image approach for consistency
            if image_pages > total_pages * 0.75:
                print(f"PDF is primarily image-based ({image_pages}/{total_pages} pages). Using full OCR approach...")
                return extract_text_from_pdf_as_images(pdf_path)
            
            # Otherwise, use hybrid approach (text extraction + OCR where needed)
            print(f"Using hybrid approach: {text_pages} text pages, {image_pages} image pages")
            
            # Extract text from each page
            for page_num in range(total_pages):
                # Add page number as a marker
                text.append(f"\n----- Page {page_num + 1} -----\n")
                
                # Check if page has text or needs OCR
                if check_if_pdf_has_text(pdf_path, page_num):
                    # Page has text - extract normally
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    text.append(page_text)
                    print(f"Page {page_num + 1}: Extracted as text")
                else:
                    # Page is image-based - use OCR with Gemini
                    print(f"Page {page_num + 1}: Using Gemini OCR (image-based)")
                    
                    # Get image from page using PyMuPDF
                    page_image = extract_image_from_pdf_page(pdf_path, page_num)
                    
                    if page_image:
                        # Use Gemini to extract text from image
                        page_text = extract_text_with_gemini(page_image, "PDF page")
                        text.append(page_text)
                    else:
                        text.append("Error: Could not extract image from PDF page")
                
        return "\n".join(text)
    except Exception as e:
        return f"Error extracting text from PDF: {str(e)}"

def save_text_to_file(text, output_path):
    """Save extracted text to a text file."""
    try:
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(text)
        return True
    except Exception as e:
        print(f"Error saving text to file: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Extract text from PDF and DOCX files')
    parser.add_argument('input_file', nargs='?', 
                        default=r"C:\Users\asingh50\Documents\OCR\img1.jpg",
                        help='Path to the input PDF or DOCX file')
    parser.add_argument('--output', help='Path to the output text file')
    parser.add_argument('--force-ocr', action='store_true', 
                        help='Force OCR for all content, even if text is detected')
    
    args = parser.parse_args()
    
    input_path = args.input_file
    
    # Determine output path if not provided
    if args.output:
        output_path = args.output
    else:
        input_file = Path(input_path)
        output_path = str(input_file.with_suffix('.txt'))
    
    # Check if input file exists
    if not os.path.isfile(input_path):
        print(f"Error: Input file '{input_path}' does not exist.")
        return
    
    # Determine file type and extract text
    file_ext = os.path.splitext(input_path)[1].lower()
    
    print(f"Processing file: {input_path}")
    print(f"File type: {file_ext}")
    
    if file_ext == '.pdf':
        print("Extracting text from PDF...")
        text = extract_text_from_pdf(input_path)
    elif file_ext in ['.docx', '.doc']:
        print("Processing DOCX with full document OCR...")
        text = extract_text_from_docx(input_path)
    else:
        print(f"Error: Unsupported file type '{file_ext}'. Please provide a PDF or DOCX file.")
        return
    
    # Save extracted text to file
    if save_text_to_file(text, output_path):
        print(f"Text extracted successfully and saved to '{output_path}'")
    else:
        print("Failed to save extracted text.")

if __name__ == "__main__":
    main()
