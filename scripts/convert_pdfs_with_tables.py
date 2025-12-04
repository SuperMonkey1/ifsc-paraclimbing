"""
PDF to Markdown converter with enhanced table extraction support.
Uses PyMuPDF (fitz) for both text and table extraction.
Falls back to OCR for pages with vector graphics (drawings) but no text.
"""

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io
import os
import re


def extract_table_as_markdown(table_data):
    """Convert extracted table data to markdown format."""
    if not table_data or len(table_data) == 0:
        return ""
    
    # Clean up cell values
    def clean_cell(cell):
        if cell is None:
            return ""
        text = str(cell).replace('\n', ' ').strip()
        text = text.replace('|', '\\|')
        return text
    
    # Filter out completely empty rows
    non_empty_rows = []
    for row in table_data:
        cells = [clean_cell(c) for c in row]
        if any(c.strip() for c in cells):
            non_empty_rows.append(cells)
    
    if len(non_empty_rows) < 1:
        return ""
    
    # Build markdown table
    md_lines = []
    
    for i, cells in enumerate(non_empty_rows):
        md_lines.append("| " + " | ".join(cells) + " |")
        
        # Add header separator after first row
        if i == 0:
            separators = ["---" for _ in cells]
            md_lines.append("| " + " | ".join(separators) + " |")
    
    return "\n".join(md_lines)


def find_tables_with_pymupdf(page):
    """Extract tables from a page using PyMuPDF's table finder."""
    tables = []
    try:
        # Use PyMuPDF's table finder
        tab_finder = page.find_tables()
        for table in tab_finder:
            # Extract the table data
            table_data = table.extract()
            if table_data:
                tables.append({
                    'data': table_data,
                    'bbox': table.bbox
                })
    except Exception as e:
        pass
    return tables


def extract_text_excluding_tables(page, table_bboxes):
    """Extract text from page excluding table regions."""
    if not table_bboxes:
        return page.get_text("text")
    
    # Get all text blocks
    blocks = page.get_text("dict")["blocks"]
    text_parts = []
    
    for block in blocks:
        if block.get("type") == 0:  # Text block
            block_rect = fitz.Rect(block["bbox"])
            
            # Check if block overlaps with any table
            in_table = False
            for bbox in table_bboxes:
                table_rect = fitz.Rect(bbox)
                if block_rect.intersects(table_rect):
                    in_table = True
                    break
            
            if not in_table:
                # Extract text from this block
                for line in block.get("lines", []):
                    line_text = ""
                    for span in line.get("spans", []):
                        line_text += span.get("text", "")
                    if line_text.strip():
                        text_parts.append(line_text.strip())
    
    return "\n".join(text_parts)


def page_to_image(page, dpi=300):
    """Convert a PDF page to a PIL Image."""
    # Render page at specified DPI
    zoom = dpi / 72  # 72 is the default PDF DPI
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    
    # Convert to PIL Image
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))
    return img


def ocr_page(page, dpi=300):
    """Perform OCR on a PDF page and return text."""
    try:
        img = page_to_image(page, dpi)
        
        # Configure tesseract for table extraction
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(img, config=custom_config)
        
        return text
    except Exception as e:
        print(f"    OCR error: {e}")
        return ""


def ocr_page_with_tables(page, dpi=300):
    """Perform OCR on a PDF page and try to extract tables."""
    try:
        img = page_to_image(page, dpi)
        
        # Get text with bounding boxes for potential table detection
        # First, try to get structured data
        custom_config = r'--oem 3 --psm 6'
        
        # Get data with bounding boxes
        data = pytesseract.image_to_data(img, config=custom_config, output_type=pytesseract.Output.DICT)
        
        # Also get plain text
        text = pytesseract.image_to_string(img, config=custom_config)
        
        return text, data
    except Exception as e:
        print(f"    OCR error: {e}")
        return "", {}


def format_ocr_text_as_table(text):
    """Attempt to format OCR text as markdown tables where appropriate."""
    if not text:
        return ""
    
    lines = text.strip().split('\n')
    formatted_lines = []
    
    # Try to detect table-like structures
    # Look for lines with multiple values separated by whitespace
    potential_table = []
    
    for line in lines:
        line = line.strip()
        if not line:
            # Empty line - might be end of table
            if potential_table:
                table_md = try_convert_to_table(potential_table)
                if table_md:
                    formatted_lines.append(table_md)
                else:
                    formatted_lines.extend(potential_table)
                potential_table = []
            formatted_lines.append("")
            continue
        
        # Check if this line looks like a table row (has multiple space-separated values)
        parts = line.split()
        
        # If line has numbers or currency-like values, it might be a table row
        has_numbers = any(re.match(r'^[\d,.\-]+$', p) or 
                         re.match(r'^[\(\)0-9,.\-]+$', p) or
                         p.startswith("CHF") or p.startswith("EUR") or p.startswith("USD")
                         for p in parts)
        
        if has_numbers and len(parts) >= 2:
            potential_table.append(line)
        else:
            # Flush any pending table
            if potential_table:
                table_md = try_convert_to_table(potential_table)
                if table_md:
                    formatted_lines.append(table_md)
                else:
                    formatted_lines.extend(potential_table)
                potential_table = []
            formatted_lines.append(line)
    
    # Handle any remaining potential table
    if potential_table:
        table_md = try_convert_to_table(potential_table)
        if table_md:
            formatted_lines.append(table_md)
        else:
            formatted_lines.extend(potential_table)
    
    return '\n'.join(formatted_lines)


def try_convert_to_table(lines):
    """Try to convert a list of lines into a markdown table."""
    if len(lines) < 2:
        return None
    
    # Split each line by multiple spaces (common in OCR'd tables)
    rows = []
    for line in lines:
        # Split by 2+ spaces to separate columns
        cells = re.split(r'\s{2,}', line.strip())
        if cells:
            rows.append(cells)
    
    if not rows:
        return None
    
    # Find the maximum number of columns
    max_cols = max(len(row) for row in rows)
    
    # If most rows don't have consistent column counts, it's probably not a table
    col_counts = [len(row) for row in rows]
    most_common_cols = max(set(col_counts), key=col_counts.count)
    
    if most_common_cols < 2:
        return None
    
    # Normalize rows to have the same number of columns
    normalized_rows = []
    for row in rows:
        while len(row) < max_cols:
            row.append("")
        normalized_rows.append(row[:max_cols])
    
    # Build markdown table
    md_lines = []
    for i, row in enumerate(normalized_rows):
        cells = [c.replace('|', '\\|') for c in row]
        md_lines.append("| " + " | ".join(cells) + " |")
        if i == 0:
            md_lines.append("| " + " | ".join(["---"] * len(cells)) + " |")
    
    return "\n".join(md_lines)


def clean_text(text):
    """Clean up extracted text."""
    if not text:
        return ""
    
    # Fix common issues with extracted text
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if line:
            cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)


def convert_pdf_to_markdown(pdf_path, output_path):
    """Convert a PDF file to markdown with table support."""
    try:
        doc = fitz.open(pdf_path)
        all_content = []
        
        for page_num, page in enumerate(doc, 1):
            page_content = []
            
            # Check if page has text content or needs OCR
            text = page.get_text("text")
            drawings = page.get_drawings()
            
            # If page has many drawings but no text, use OCR
            if len(text.strip()) < 50 and len(drawings) > 100:
                print(f"    Page {page_num}: Using OCR (vector graphics detected)")
                ocr_text = ocr_page(page, dpi=300)
                ocr_text = format_ocr_text_as_table(ocr_text)
                if ocr_text:
                    page_content.append(ocr_text)
            else:
                # Find tables on this page
                tables = find_tables_with_pymupdf(page)
                table_bboxes = [t['bbox'] for t in tables]
                
                # Extract text (excluding table areas if tables were found)
                if tables:
                    text = extract_text_excluding_tables(page, table_bboxes)
                
                text = clean_text(text)
                
                if text:
                    page_content.append(text)
                
                # Add tables as markdown
                for table_info in tables:
                    table_md = extract_table_as_markdown(table_info['data'])
                    if table_md:
                        page_content.append("\n" + table_md + "\n")
            
            if page_content:
                all_content.append('\n\n'.join(page_content))
            
            all_content.append("\n\n---\n\n")  # Page separator
        
        doc.close()
        
        # Get filename for title
        filename = os.path.basename(pdf_path)
        title = filename.replace('.pdf', '').replace('.PDF', '').upper()
        
        # Write markdown file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f'# {title}\n\n')
            f.write('\n'.join(all_content))
        
        return True
        
    except Exception as e:
        print(f"Error converting {pdf_path}: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    base_path = 'c:/PYTHON/ifsc-paraclimbing/'
    input_folder = base_path + 'official_documents/'
    output_folder = base_path + 'official_documents_md_format/'
    
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Get all PDF files from the input folder
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("No PDF files found in the input folder.")
        return
    
    print(f"Found {len(pdf_files)} PDF files to convert.\n")
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        md_name = pdf_file.replace('.pdf', '.md').replace('.PDF', '.md')
        output_path = os.path.join(output_folder, md_name)
        
        if os.path.exists(pdf_path):
            print(f"Converting: {pdf_file}...")
            if convert_pdf_to_markdown(pdf_path, output_path):
                print(f"  -> {md_name} ✓")
            else:
                print(f"  -> Failed ✗")
        else:
            print(f"Not found: {pdf_file}")
    
    print("\nConversion complete!")


if __name__ == "__main__":
    main()
