#!/usr/bin/env python3
"""
Clean external references from encrypted Excel file
This script removes problematic external reference metadata that causes openpyxl to fail
"""

import msoffcrypto
import io
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import os
from dotenv import load_dotenv

# Load password
load_dotenv()
password = os.getenv('password', 'Laxmi7541')

print("Cleaning Asset.xlsx...")

# Step 1: Decrypt the file
print("1. Decrypting file...")
decrypted = io.BytesIO()
with open('Asset.xlsx', 'rb') as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)

decrypted.seek(0)

# Step 2: Clean the zip contents
print("2. Removing external references...")
cleaned = io.BytesIO()

with ZipFile(decrypted, 'r') as zin:
    with ZipFile(cleaned, 'w') as zout:
        for item in zin.infolist():
            # Skip external link files completely
            if 'externalLink' in item.filename:
                print(f"   Skipping: {item.filename}")
                continue
            
            data = zin.read(item.filename)
            
            # Clean workbook.xml to remove externalReferences element
            if item.filename == 'xl/workbook.xml':
                print(f"   Cleaning: {item.filename}")
                try:
                    # Parse XML with namespace handling
                    root = ET.fromstring(data)
                    
                    # Find and remove externalReferences elements (with any namespace)
                    for elem in list(root):
                        if 'externalReferences' in elem.tag:
                            print(f"      Removing externalReferences element")
                            root.remove(elem)
                    
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                except Exception as e:
                    print(f"   Warning: Could not clean workbook.xml: {e}")
            
            # Clean workbook.xml.rels to remove external reference relationships
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                print(f"   Cleaning: {item.filename}")
                try:
                    root = ET.fromstring(data)
                    # Remove relationships with Type containing "externalLink"
                    for elem in list(root):
                        if 'Type' in elem.attrib and 'externalLink' in elem.attrib['Type']:
                            print(f"      Removing external link relationship: {elem.attrib.get('Target', 'unknown')}")
                            root.remove(elem)
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                except Exception as e:
                    print(f"   Warning: Could not clean workbook.xml.rels: {e}")
            
            zout.writestr(item, data)

# Step 3: Save the cleaned file
print("3. Saving cleaned file as Asset_cleaned.xlsx...")
cleaned.seek(0)
with open('Asset_cleaned.xlsx', 'wb') as f:
    f.write(cleaned.read())

print("\n✓ Done! Created Asset_cleaned.xlsx")
print("\nYou can now:")
print("  1. Open Asset_cleaned.xlsx in Excel (it's NOT password protected)")
print("  2. Add password protection back if needed: File → Protect Workbook")
print("  3. Save as Asset.xlsx")
print("  4. Or just use Asset_cleaned.xlsx directly with: python process_assets.py --file Asset_cleaned.xlsx --normalize")
