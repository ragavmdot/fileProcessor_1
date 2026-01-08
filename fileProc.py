import os
import shutil
from openpyxl import Workbook

def filenames_to_excel(folder_path, output_file, archive_path):
    # Create a new workbook and select active sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "FileWords"

    # Write fixed headers
    headers = ["Col1", "Col2", "Col3", "Col4"]
    ws.append(headers)

    # Iterate over files in the folder
    for file in os.listdir(folder_path):
        full_path = os.path.join(folder_path, file)
        if os.path.isfile(os.path.join(folder_path, file)):
            # Remove file extension
            filename = os.path.splitext(file)[0]
            # Split by underscore
            words = filename.split("_")
            # Ensure only 4 columns (pad with blanks if shorter)
            row = words[:4] + [""] * (4 - len(words))
            ws.append(row)

            # Move the file to archive folder
            shutil.move(full_path, os.path.join(archive_path, file))

    # Save the Excel file to the chosen path
    wb.save(output_file)
    print(f"✅ Excel file created at: {output_file}")

    


# Example usage
if __name__ == "__main__":
    folder = "/Users/ragavmohankumar/Ragav/FileProcAppPath/SourceFolder" 
    output = "/Users/ragavmohankumar/Ragav/FileProcAppPath/TargetFolder/filenames.xlsx"
    archive = "/Users/ragavmohankumar/Ragav/FileProcAppPath/ArchiveFolder" 
    filenames_to_excel(folder, output, archive)