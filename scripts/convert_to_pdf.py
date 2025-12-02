from playwright.sync_api import sync_playwright
import os

html_path = 'c:/PYTHON/ifsc-paraclimbing/VI_Requests_Rule_Changes.html'
pdf_path = 'c:/PYTHON/ifsc-paraclimbing/VI_Requests_Rule_Changes.pdf'

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto(f'file:///{html_path}')
    page.pdf(path=pdf_path, format='A4', margin={'top': '20mm', 'bottom': '20mm', 'left': '15mm', 'right': '15mm'})
    browser.close()

print(f'Created: {pdf_path}')
