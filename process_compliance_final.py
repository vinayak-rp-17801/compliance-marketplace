import os
import zipfile
from pathlib import Path
from docx import Document
from docx.shared import Inches
import re

def sanitize_folder_name(doc_name):
    name = re.sub(r'^\d+_', '', doc_name)
    name = name.replace('.docx', '')
    name = name.replace(' ', '_')
    name = name.replace('.', '_')
    return name

def extract_zips(docs_zip, logos_zip, extract_dir):
    docs_extract = os.path.join(extract_dir, 'docs')
    logos_extract = os.path.join(extract_dir, 'logos')
    
    os.makedirs(docs_extract, exist_ok=True)
    os.makedirs(logos_extract, exist_ok=True)
    
    print("Extracting compliance_docs.zip...")
    with zipfile.ZipFile(docs_zip, 'r') as zip_ref:
        zip_ref.extractall(docs_extract)
    
    print("Extracting compliance logo marketplace.zip...")
    with zipfile.ZipFile(logos_zip, 'r') as zip_ref:
        zip_ref.extractall(logos_extract)
    
    return docs_extract, logos_extract

def find_images(logos_extract, folder_name):
    folder_path = os.path.join(logos_extract, folder_name)
    
    if not os.path.exists(folder_path):
        return None, None
    
    logo_path = None
    thumbnail_path = None
    
    for file in os.listdir(folder_path):
        if '180x180' in file or '180_180' in file:
            logo_path = os.path.join(folder_path, file)
        elif '740x340' in file or '740_340' in file:
            thumbnail_path = os.path.join(folder_path, file)
    
    return logo_path, thumbnail_path

def add_images_to_document(doc_path, logo_path, thumbnail_path, output_path):
    try:
        doc = Document(doc_path)
        
        if logo_path and os.path.exists(logo_path):
            paragraph = doc.add_paragraph()
            run = paragraph.add_run()
            run.add_picture(logo_path, width=Inches(1.875), height=Inches(1.875))
            paragraph.alignment = 1
        
        if thumbnail_path and os.path.exists(thumbnail_path):
            paragraph = doc.add_paragraph()
            run = paragraph.add_run()
            run.add_picture(thumbnail_path, width=Inches(7.708), height=Inches(3.542))
            paragraph.alignment = 1
        
        doc.save(output_path)
        return True
    except Exception as e:
        print(f"  Error: {str(e)}")
        return False

def main():
    docs_zip = 'compliance_docs.zip'
    logos_zip = 'compliance logo marketplace.zip'
    extract_dir = 'extracted_files'
    output_dir = 'output_docs'
    
    os.makedirs(output_dir, exist_ok=True)
    
    print("=" * 70)
    print("STEP 1: Extracting zip files...")
    print("=" * 70)
    docs_extract, logos_extract = extract_zips(docs_zip, logos_zip, extract_dir)
    
    print("\n" + "=" * 70)
    print("STEP 2: Processing compliance documents...")
    print("=" * 70 + "\n")
    
    docx_files = sorted(list(Path(docs_extract).rglob("*.docx")))
    total_files = len(docx_files)
    successful = 0
    
    print(f"Found {total_files} DOCX files\n")
    
    for idx, doc_path in enumerate(docx_files, 1):
        doc_name = doc_path.name
        folder_name = sanitize_folder_name(doc_name)
        
        print(f"[{idx}/{total_files}] {doc_name}")
        
        logo_path, thumbnail_path = find_images(logos_extract, folder_name)
        
        output_path = os.path.join(output_dir, doc_name)
        if add_images_to_document(str(doc_path), logo_path, thumbnail_path, output_path):
            successful += 1
    
    print("\n" + "=" * 70)
    print("COMPLETE: " + str(successful) + "/" + str(total_files) + " documents processed")
    print("=" * 70)

if __name__ == "__main__":
    main()
