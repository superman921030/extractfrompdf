import fitz
import pandas as pd
import os

def extract_data_from_page(page):
    text = page.get_text("text")
    
    data = {}
    
    # print(text)
    
    fields = ['Reference', 'Adresse', 'PARCELLE', 'ILOT', 'Type', 'Date de redaction', 'Denominations successives', 'EMJ / MP', 'Date de l\'enquete', 'Iconographie generale', 'Galerie photographique', 'Description generale', 'Interet patrimonial', 'Points faibles', 'Etat structurel', 'Etat sanitaire', 'Dossier SEM', 'Travaux', 'Plan de Sauvegarde et de Mise en Valeur', 'Prescription archeo', 'Photos']
    value = [1, 1, -1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    for i, field in enumerate(fields):
        start_idx = text.find(field)
        
        if start_idx != -1:
            if value[i] == 1:
                end_idx = text.find('\n', start_idx + len(field) + 1)
                data[field] = text[start_idx+len(field) + 1:end_idx].strip()
            else:
                previous_start = text.rfind('\n', 0, start_idx - 1)
                data[field] = text[previous_start + 1: start_idx - 1].strip()

            print(i, field, data[field])
    

    return data

def extract_tick_box_selections(page):
    tick_box_selections = []
    checkbox_instances = page.search_for('checkbox')

    for inst in checkbox_instances:
        x, y, _, _ = inst
        selected = any(page.get_image_xobjects(full=True, clip=(x, y, x+10, y+10)))
        tick_box_selections.append("Selected" if selected else "Not Selected")
    
    return tick_box_selections

# Initialize a DataFrame to store the extracted data
pdf_data = pd.DataFrame(columns=['File Name', 'Text Content', 'Image Paths', 'Tick Box Selections', 'Reference', 'Adresse', 'PARCELLE', 'ILOT', 'Type', 'Date de redaction', 'Denominations successives', 'EMJ / MP', 'Date de l\'enquete', 'Iconographie generale', 'Galerie photographique', 'Description generale', 'Interet patrimonial', 'Points faibles', 'Etat structurel', 'Etat sanitaire', 'Dossier SEM', 'Travaux', 'Plan de Sauvegarde et de Mise en Valeur', 'Prescription archeo', 'Photos'])

# List of PDF files to process
pdf_files = ['206-BD TRACr14-16.pdf']

# Create the base directory if it doesn't exist
base_directory = os.path.join("DB", "Galerie")
os.makedirs(base_directory, exist_ok=True)

for file_name in pdf_files:
    file_directory = os.path.join(base_directory, os.path.splitext(file_name)[0])
    os.makedirs(file_directory, exist_ok=True)
    
    pdf_doc = fitz.open(file_name)

    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text_content = page.get_text()
        image_paths = []
        tick_box_selections = extract_tick_box_selections(page)
        data = extract_data_from_page(page)

        for img_info in page.get_images(full=True):
            xref = img_info[0]
            img_data = pdf_doc.extract_image(xref)
            img_path = os.path.join(file_directory, f'page_{page_num}_image_{xref}.png')
            with open(img_path, 'wb') as img_file:
                img_file.write(img_data['image'])
            image_paths.append(img_path)

        # Append the extracted data to the DataFrame
        pdf_data = pdf_data.append({'File Name': file_name, 
                                    'Text Content': text_content,
                                    'Image Paths': ','.join(image_paths), 
                                    'Tick Box Selections': ','.join(tick_box_selections),
                                    **data},  # Unpack the data dictionary
                                    ignore_index=True)

    pdf_doc.close()

# Specify the path to save the Excel file in the DB directory
output_path = os.path.join(os.getcwd(), "DB", "db.xlsx")

# Save the DataFrame to the Excel file in the DB directory
pdf_data.to_excel(output_path, index=False)