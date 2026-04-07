import os
import zipfile
from pathlib import Path

# Extract both zips
print("Extracting zips...")
with zipfile.ZipFile('compliance_docs.zip', 'r') as z:
    z.extractall('extracted_docs')

with zipfile.ZipFile('compliance logo marketplace.zip', 'r') as z:
    z.extractall('extracted_logos')

# Show docs structure
print("\n" + "="*70)
print("DOCS STRUCTURE:")
print("="*70)
for root, dirs, files in os.walk('extracted_docs'):
    level = root.replace('extracted_docs', '').count(os.sep)
    indent = ' ' * 2 * level
    print(f'{indent}{os.path.basename(root)}/')
    subindent = ' ' * 2 * (level + 1)
    for file in files[:3]:
        print(f'{subindent}{file}')
    if len(files) > 3:
        print(f'{subindent}... and {len(files) - 3} more files')

# Show logos structure
print("\n" + "="*70)
print("LOGOS STRUCTURE:")
print("="*70)
for root, dirs, files in os.walk('extracted_logos'):
    level = root.replace('extracted_logos', '').count(os.sep)
    indent = ' ' * 2 * level
    if level < 3:  # Only show first 3 levels
        print(f'{indent}{os.path.basename(root)}/')
        subindent = ' ' * 2 * (level + 1)
        for file in files:
            print(f'{subindent}{file}')

# Show document names
print("\n" + "="*70)
print("FIRST 5 DOCUMENT NAMES:")
print("="*70)
docs = list(Path('extracted_docs').rglob("*.docx"))
for doc in docs[:5]:
    print(f"Doc: {doc.name}")

# Show folder names in logos
print("\n" + "="*70)
print("FOLDER NAMES IN LOGOS (first 5):")
print("="*70)
logo_dirs = [d for d in os.listdir('extracted_logos') if os.path.isdir(os.path.join('extracted_logos', d))]
for folder in logo_dirs[:5]:
    print(f"Folder: {folder}")
