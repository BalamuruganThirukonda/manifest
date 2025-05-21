import os
from openpyxl import Workbook


base_dir = input("Enter the base directory path: ")
if not os.path.exists(base_dir):
    print(f"Directory {base_dir} does not exist.")
    exit(1)
    
# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active
ws.title = "Manifest"
ws.append(["Year", "File Name", "Path", "Staining"])

for folder_name in os.listdir(base_dir):
    folder_path = os.path.join(base_dir, folder_name)   
    
    if os.path.isdir(folder_path):
        print(f"Scanning folder: {folder_path}")
        for file_name in os.listdir(folder_path):            
            if file_name.endswith((".svs", ".ndpi")):
                print(f"Found file: {file_name}")
                if file_name.startswith(('E', 'R')) and len(file_name) > 6:
                    year = file_name[1:5]
                else:
                    year = "Unknown"
                # Construct the full path
                path = os.path.join(folder_path, file_name)
                
                   
                # Extract staining
                parts = file_name.rsplit("-", 1)
                staining = parts[-1].split('.')[0] if len(parts) > 1 else "Unknown"
                                
                # Append the data to the worksheet
                ws.append([year, file_name, path, staining])
                
# Save the workbook and remove existing file if it exists
output_path = os.path.join(base_dir, "manifest.xlsx")
if os.path.exists(output_path):
    os.remove(output_path)        
    


wb.save(output_path)
print(f"Excel file created at: {output_path}")