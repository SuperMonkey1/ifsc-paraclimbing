import markdown
import os
import glob

# Define the Requests folder path
requests_folder = 'c:/PYTHON/ifsc-paraclimbing/Requests'

# Find all markdown files in the Requests folder
md_files = glob.glob(os.path.join(requests_folder, '*.md'))

if not md_files:
    print('No markdown files found in Requests folder.')
else:
    print(f'Found {len(md_files)} markdown file(s) to convert.')

for md_file in md_files:
    # Read the markdown file
    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    # Get the base filename for the title
    base_name = os.path.splitext(os.path.basename(md_file))[0]
    
    # Convert to HTML with tables extension
    html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])
    
    # Create styled HTML
    styled_html = f'''<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{base_name}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 800px;
            margin: 40px auto;
            padding: 20px;
            line-height: 1.6;
            color: #333;
        }}
        h1 {{
            color: #1a5276;
            border-bottom: 3px solid #1a5276;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #2874a6;
            margin-top: 30px;
            border-bottom: 1px solid #ddd;
            padding-bottom: 5px;
        }}
        h3 {{
            color: #1a5276;
        }}
        blockquote {{
            background: #f9f9f9;
            border-left: 4px solid #2874a6;
            margin: 15px 0;
            padding: 10px 20px;
            font-style: italic;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 10px;
            text-align: left;
        }}
        th {{
            background-color: #2874a6;
            color: white;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        strong {{
            color: #1a5276;
        }}
        hr {{
            border: none;
            border-top: 2px solid #eee;
            margin: 30px 0;
        }}
        code {{
            background: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
        }}
    </style>
</head>
<body>
{html_content}
</body>
</html>'''
    
    # Create output HTML path
    html_output = os.path.splitext(md_file)[0] + '.html'
    
    # Save as HTML
    with open(html_output, 'w', encoding='utf-8') as f:
        f.write(styled_html)
    
    print(f'Created: {html_output}')

print('\nDone! You can open HTML files in a browser and print to PDF (Ctrl+P -> Save as PDF)')
