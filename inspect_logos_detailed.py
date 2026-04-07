import os
import zipfile
from pathlib import Path

print("="*70)
print("INSPECTING LOGOS ZIP STRUCTURE")
print("="*70)

# Extract
with zipfile.ZipFile('compliance logo marketplace.zip', 'r') as z:
    z.extractall('inspect_logos')

print("\nDirect contents of extracted folder:")
print("-" * 70)
for item in os.listdir('inspect_logos'):
    path = os.path.join('inspect_logos', item)
    if os.path.isdir(path):
        print(f"📁 FOLDER: {item}")
        # Show contents of folder
        for subitem in os.listdir(path)[:3]:
            print(f"   - {subitem}")
    else:
        print(f"📄 FILE: {item}")

print("\n\nSearching for all .png files:")
print("-" * 70)
png_files = list(Path('inspect_logos').rglob("*.png"))
print(f"Total PNG files found: {len(png_files)}\n")

# Show first 10
for i, png_file in enumerate(png_files[:10], 1):
    relative_path = str(png_file.relative_to('inspect_logos'))
    print(f"{i}. {relative_path}")

print("\n\nLooking for folder names (one level deep):")
print("-" * 70)
base = 'inspect_logos'
folders = []
for item in os.listdir(base):
    full_path = os.path.join(base, item)
    if os.path.isdir(full_path):
        folders.append(item)

print(f"Total folders: {len(folders)}\n")
for folder in sorted(folders)[:15]:
    folder_path = os.path.join(base, folder)
    files_in_folder = os.listdir(folder_path)
    png_count = len([f for f in files_in_folder if f.endswith('.png')])
    print(f"📁 {folder} ({png_count} PNG files)")

print("\n" + "="*70)
