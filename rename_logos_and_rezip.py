import os
import zipfile
import shutil
import re
from pathlib import Path

def sanitize_folder_name(doc_name):
    """Convert document name to folder name format"""
    name = re.sub(r'^\d+_', '', doc_name)
    name = name.replace('.docx', '')
    name = name.replace(' ', '_')
    name = name.replace('.', '_')
    return name

# Step 1: Extract both zips
print("="*70)
print("STEP 1: Extracting files...")
print("="*70)

docs_dir = 'extracted_docs_for_rename'
logos_dir = 'extracted_logos_for_rename'

os.makedirs(docs_dir, exist_ok=True)
os.makedirs(logos_dir, exist_ok=True)

print("Extracting compliance_docs.zip...")
with zipfile.ZipFile('compliance_docs.zip', 'r') as z:
    z.extractall(docs_dir)

print("Extracting compliance logo marketplace.zip...")
with zipfile.ZipFile('compliance logo marketplace.zip', 'r') as z:
    z.extractall(logos_dir)

# Step 2: Get all actual document names
print("\n" + "="*70)
print("STEP 2: Building mapping of documents to logos...")
print("="*70)

docx_files = sorted([f.name for f in Path(docs_dir).rglob("*.docx") if not f.name.startswith('._')])
print(f"\nFound {len(docx_files)} documents\n")

# Step 3: List what we have in logos
print("="*70)
print("STEP 3: Listing current logo folders...")
print("="*70)

logos_base = os.path.join(logos_dir, 'compliance logo marketplace') if os.path.exists(os.path.join(logos_dir, 'compliance logo marketplace')) else logos_dir

current_folders = []
for item in os.listdir(logos_base):
    item_path = os.path.join(logos_base, item)
    if os.path.isdir(item_path) and not item.startswith('._') and not item.startswith('__MACOSX'):
        current_folders.append(item)

print(f"\nCurrent folders in logos: {len(current_folders)}")
print("First 5 current folder names:")
for folder in sorted(current_folders)[:5]:
    print(f"  - {folder}")

# Step 4: Build mapping
print("\n" + "="*70)
print("STEP 4: Creating mapping...")
print("="*70)

mapping = {}
for doc in docx_files:
    expected_name = sanitize_folder_name(doc)
    mapping[doc] = expected_name

print(f"\nSample mappings:")
for doc, expected in list(mapping.items())[:5]:
    print(f"  {doc} → {expected}")

# Step 5: Find matching folders and rename
print("\n" + "="*70)
print("STEP 5: Renaming logo folders...")
print("="*70)

renamed_count = 0
for current_folder in current_folders:
    # Try to find which document this folder matches
    for doc, expected_name in mapping.items():
        # Try exact match or partial match
        if current_folder.lower() == expected_name.lower() or \
           current_folder.lower().replace('-', '').replace('_', '') == expected_name.lower().replace('-', '').replace('_', ''):
            
            old_path = os.path.join(logos_base, current_folder)
            new_path = os.path.join(logos_base, expected_name)
            
            if old_path != new_path:
                if os.path.exists(new_path):
                    shutil.rmtree(new_path)
                shutil.move(old_path, new_path)
                print(f"✓ Renamed: {current_folder} → {expected_name}")
                renamed_count += 1
            break

print(f"\nRenamed {renamed_count} folders")

# Step 6: Verify structure
print("\n" + "="*70)
print("STEP 6: Verifying structure...")
print("="*70)

verification = []
for doc in docx_files[:10]:
    expected_folder = sanitize_folder_name(doc)
    folder_path = os.path.join(logos_base, expected_folder)
    exists = os.path.exists(folder_path)
    
    if exists:
        files = [f for f in os.listdir(folder_path) if not f.startswith('._')]
        verification.append((doc, expected_folder, True, files))
        print(f"✓ {doc} → {expected_folder} (Found: {files})")
    else:
        verification.append((doc, expected_folder, False, []))
        print(f"✗ {doc} → {expected_folder} (NOT FOUND)")

# Step 7: Re-zip
print("\n" + "="*70)
print("STEP 7: Creating new zip files...")
print("="*70)

# Remove old zips
backup_docs = 'compliance_docs_backup.zip'
backup_logos = 'compliance_logo_marketplace_backup.zip'

if os.path.exists('compliance_docs.zip'):
    shutil.move('compliance_docs.zip', backup_docs)
    print(f"Backed up original docs zip to {backup_docs}")

if os.path.exists('compliance logo marketplace.zip'):
    shutil.move('compliance logo marketplace.zip', backup_logos)
    print(f"Backed up original logos zip to {backup_logos}")

# Create new docs zip
print("\nCreating new compliance_docs.zip...")
with zipfile.ZipFile('compliance_docs.zip', 'w', zipfile.ZIP_DEFLATED) as z:
    for file in Path(docs_dir).rglob("*"):
        if file.is_file() and not file.name.startswith('._'):
            arcname = str(file.relative_to(docs_dir))
            z.write(file, arcname)

# Create new logos zip
print("Creating new compliance logo marketplace.zip...")
with zipfile.ZipFile('compliance logo marketplace.zip', 'w', zipfile.ZIP_DEFLATED) as z:
    for file in Path(logos_dir).rglob("*"):
        if file.is_file() and not file.name.startswith('._'):
            arcname = str(file.relative_to(logos_dir))
            z.write(file, arcname)

print("\n" + "="*70)
print("COMPLETE!")
print("="*70)
print("✓ Backup files created")
print("✓ Logo folders renamed to match document names")
print("✓ New zip files created")
print("\nNow you can run the process_compliance_final.py script!")
print("="*70)
