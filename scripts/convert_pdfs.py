import fitz
import os

base_path = 'c:/PYTHON/ifsc-paraclimbing/'
input_folder = base_path + 'official_documents/'
output_folder = base_path + 'official_documents_md_format/'

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Get all PDF files from the input folder
pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

for pdf_file in pdf_files:
    pdf_path = input_folder + pdf_file
    if os.path.exists(pdf_path):
        doc = fitz.open(pdf_path)
        text = ''
        for page in doc:
            text += page.get_text() + '\n\n---\n\n'
        doc.close()
        
        md_name = pdf_file.replace('.pdf', '.md')
        with open(output_folder + md_name, 'w', encoding='utf-8') as f:
            f.write('# ' + pdf_file.replace('.pdf', '').upper() + '\n\n')
            f.write(text)
        print('Converted: ' + pdf_file + ' -> ' + md_name)
    else:
        print('Not found: ' + pdf_file)
