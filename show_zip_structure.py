import zipfile

print("="*70)
print("READING compliance logo marketplace.zip DIRECTLY")
print("="*70)

with zipfile.ZipFile('compliance logo marketplace.zip', 'r') as z:
    all_files = z.namelist()
    
    print(f"\nTotal items in ZIP: {len(all_files)}\n")
    
    print("First 50 items in ZIP:")
    print("-" * 70)
    for i, item in enumerate(all_files[:50], 1):
        print(f"{i}. {item}")
    
    print("\n" + "="*70)
    print("SEARCHING FOR UNIQUE FOLDER NAMES")
    print("="*70)
    
    folders = set()
    for item in all_files:
        # Extract folder name (first level)
        parts = item.strip('/').split('/')
        if len(parts) > 1 and parts[0] != '__MACOSX':
            folders.add(parts[0])
    
    print(f"\nFound {len(folders)} unique folders:\n")
    for folder in sorted(folders)[:20]:
        print(f"  - {folder}")
    
    if len(folders) > 20:
        print(f"  ... and {len(folders) - 20} more")
