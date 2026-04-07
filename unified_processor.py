import os
import zipfile
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Inches
import re

print("="*70)
print("STEP 1: READING ZIP CONTENTS")
print("="*70)

# First, let's see what's ACTUALLY in the logos zip
with zipfile.ZipFile('compliance logo marketplace.zip', 'r') as z:
    all_files = z.namelist()
    print(f"Total items in logos zip: {len(all_files)}\n")
    
    # Extract unique folder names
    folders = set()
    for item in all_files:
        parts = item.strip('/').split('/')
        if len(parts) > 1 and parts[0] not in ['__MACOSX', '']:
            folders.add(parts[0])
    
    print(f"Found {len(folders)} folders in logos zip:\n")
    folders_list = sorted(folders)
    for i, folder in enumerate(folders_list[:15], 1):
        print(f"  {i}. {folder}")
    
    if len(folders_list) > 15:
        print(f"  ... and {len(folders_list) - 15} more\n")

print("\n" + "="*70)
print("STEP 2: EXTRACTING FILES")
print("="*70)

docs_extract = 'extracted_docs'
logos_extract = 'extracted_logos'

os.makedirs(docs_extract, exist_ok=True)
os.makedirs(logos_extract, exist_ok=True)

print("Extracting compliance_docs.zip...")
with zipfile.ZipFile('compliance_docs.zip', 'r') as z:
    z.extractall(docs_extract)

print("Extracting compliance logo marketplace.zip...")
with zipfile.ZipFile('compliance logo marketplace.zip', 'r') as z:
    z.extractall(logos_extract)

print("\n" + "="*70)
print("STEP 3: CHECKING EXTRACTED STRUCTURE")
print("="*70)

# Check what we extracted
print("\nContents of extracted_logos:")
for root, dirs, files in os.walk(logos_extract):
    level = root.replace(logos_extract, '').count(os.sep)
    if level < 2:  # Only show first 2 levels
        indent = ' ' * 2 * level
        print(f'{indent}{os.path.basename(root)}/')
        subindent = ' ' * 2 * (level + 1)
        for d in dirs[:5]:
            print(f'{subindent}{d}/')
        if len(dirs) > 5:
            print(f'{subindent}... and {len(dirs)-5} more folders')
        break

print("\n" + "="*70)
print("STEP 4: MATCHING DOCUMENTS WITH IMAGES")
print("="*70)

# Get all DOCX files
all_docx = list(Path(docs_extract).rglob("*.docx"))
docx_files = sorted([f for f in all_docx if not f.name.startswith('._')])

print(f"\nFound {len(docx_files)} documents\n")

# Get all actual folder names in logos
logos_base = logos_extract
actual_folders = []
for item in os.listdir(logos_base):
    full_path = os.path.join(logos_base, item)
    if os.path.isdir(full_path) and not item.startswith('._') and not item.startswith('__MACOSX'):
        actual_folders.append(item)

print(f"Actual logo folders available: {len(actual_folders)}\n")

# Try to match each document with a folder
print("Matching documents to folders:")
print("-" * 70)

matches_found = 0
for i, doc in enumerate(docx_files[:10], 1):  # Show first 10
    doc_name = doc.name
    doc_base = doc_name.replace('.docx', '').replace('001_', '').replace('002_', '').replace('003_', '').replace('004_', '').replace('005_', '').replace('006_', '').replace('007_', '').replace('008_', '').replace('009_', '').replace('010_', '')
    
    # Remove leading number and underscore
    doc_base = re.sub(r'^\d+_', '', doc_name).replace('.docx', '')
    
    # Try to find matching folder
    match = None
    for folder in actual_folders:
        # Check various matching strategies
        if folder.lower() == doc_base.lower() or \
           folder.lower().replace(' ', '_').replace('.', '_') == doc_base.lower().replace(' ', '_').replace('.', '_') or \
           folder.lower().replace('-', '').replace('_', '') == doc_base.lower().replace('-', '').replace('_', ''):
            match = folder
            break
    
    if match:
        print(f"✓ {doc_name} → {match}")
        matches_found += 1
    else:
        print(f"✗ {doc_name} → NO MATCH (looked for: {doc_base})")

print(f"\nMatches found: {matches_found}/10")

print("\n" + "="*70)
print("Please share this output!")
print("="*70)
