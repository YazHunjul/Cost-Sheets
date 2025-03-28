import zipfile
import os
import shutil
from pathlib import Path

def clean_excel_template():
    # Define paths
    template_path = "resources/Halton Cost Sheet Jan 2025.xlsx"
    temp_dir = "temp_folder"
    temp_xlsx = "temp.xlsx"
    
    print(f"Starting cleanup of {template_path}")
    
    try:
        # Create temp directory
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            print("Created temporary directory")
        
        # Extract Excel file
        print("Extracting Excel file...")
        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Remove drawings folder
        drawings_path = os.path.join(temp_dir, 'xl', 'drawings')
        if os.path.exists(drawings_path):
            shutil.rmtree(drawings_path)
            print("Removed drawings folder")
            
            # Update the drawing relationships file
            rels_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
            if os.path.exists(rels_path):
                # Create backup
                shutil.copy2(rels_path, rels_path + '.bak')
                
                # Read and modify relationships file to remove drawing references
                with open(rels_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Remove drawing relationship entries
                import re
                content = re.sub(r'<Relationship [^>]*Target="drawings/drawing[^>]*?/>', '', content)
                
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(content)
        
        # Create new zip file
        print("Creating cleaned Excel file...")
        with zipfile.ZipFile(temp_xlsx, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # First, add [Content_Types].xml
            content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
            if os.path.exists(content_types_path):
                zipf.write(content_types_path, '[Content_Types].xml')
            
            # Then add _rels folder
            rels_dir = os.path.join(temp_dir, '_rels')
            if os.path.exists(rels_dir):
                for root, dirs, files in os.walk(rels_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)
            
            # Then add xl folder
            xl_dir = os.path.join(temp_dir, 'xl')
            if os.path.exists(xl_dir):
                for root, dirs, files in os.walk(xl_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)
            
            # Add any remaining files
            for item in os.listdir(temp_dir):
                if item not in ['[Content_Types].xml', '_rels', 'xl']:
                    item_path = os.path.join(temp_dir, item)
                    if os.path.isfile(item_path):
                        zipf.write(item_path, item)
        
        # Make backup of original file
        backup_path = template_path + '.backup'
        if os.path.exists(template_path):
            shutil.copy2(template_path, backup_path)
            print(f"Created backup at {backup_path}")
        
        # Replace original file
        print("Replacing original file...")
        shutil.move(temp_xlsx, template_path)
        
        # Cleanup
        print("Cleaning up temporary files...")
        shutil.rmtree(temp_dir)
        
        print("Cleanup completed successfully!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        # Cleanup in case of error
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        if os.path.exists(temp_xlsx):
            os.remove(temp_xlsx)
        # Restore from backup if needed
        backup_path = template_path + '.backup'
        if os.path.exists(backup_path) and not os.path.exists(template_path):
            shutil.move(backup_path, template_path)

if __name__ == "__main__":
    clean_excel_template() 