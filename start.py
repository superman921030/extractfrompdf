import fitz
import pandas as pd

def extract_tick_box_selections(page):
    tick_box_selections = []
    checkbox_instances = page.search_for('checkbox')

    for inst in checkbox_instances:
        x, y, _, _ = inst
        selected = any(page.get_image_xobjects(full=True, clip=(x, y, x+10, y+10)))
        if selected:
            tick_box_selections.append("Selected")
        else:
            tick_box_selections.append("Not Selected")
    
    return tick_box_selections

# Initialize a DataFrame to store the extracted data
pdf_data = pd.DataFrame(columns=['File Name', 'Text Content', 'Image Paths', 'Tick Box Selections'])

# List of PDF files to process
pdf_files = ['file1.pdf']

for file_name in pdf_files:
    pdf_doc = fitz.open(file_name)
    text_content = ''
    image_paths = []
    tick_box_selections = []

    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text_content += page.get_text()

        for img_info in page.get_images(full=True):
            xref = img_info[0]  # xref is the first item in the tuple
            img_data = pdf_doc.extract_image(xref)  # Pass the correct xref
            img_path = f'{file_name}_page_{page_num}_image_{xref}.png'
            with open(img_path, 'wb') as img_file:
                img_file.write(img_data['image'])
            image_paths.append(img_path)

        tick_box_selections.extend(extract_tick_box_selections(page))

    pdf_data = pdf_data.append({'File Name': file_name, 
                                'Text Content': text_content,
                                'Image Paths': ','.join(image_paths), 
                                'Tick Box Selections': ','.join(tick_box_selections)}, 
                                ignore_index=True)

pdf_data.to_excel('pdf_data.xlsx', index=False)