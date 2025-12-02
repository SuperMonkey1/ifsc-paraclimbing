from playwright.sync_api import sync_playwright
import os
import glob

# Define the Requests folder path
requests_folder = 'c:/PYTHON/ifsc-paraclimbing/Requests'

# Find all HTML files in the Requests folder
html_files = glob.glob(os.path.join(requests_folder, '*.html'))

if not html_files:
    print('No HTML files found in Requests folder.')
else:
    print(f'Found {len(html_files)} HTML file(s) to convert to PDF.')
    
    with sync_playwright() as p:
        browser = p.chromium.launch()
        
        for html_path in html_files:
            # Create output PDF path
            pdf_path = os.path.splitext(html_path)[0] + '.pdf'
            
            page = browser.new_page()
            page.goto(f'file:///{html_path}')
            page.pdf(path=pdf_path, format='A4', margin={'top': '20mm', 'bottom': '20mm', 'left': '15mm', 'right': '15mm'})
            page.close()
            
            print(f'Created: {pdf_path}')
        
        browser.close()

print('\nDone!')
