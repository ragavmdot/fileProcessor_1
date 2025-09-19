import os
import shutil
import pandas as pd
from pathlib import Path
from typing import List, Tuple


class FileProcessor:
    def __init__(self, source_folder: str, excel_output_path: str, archive_folder: str):
        """
        Initialize the FileProcessor with source folder, output path, and archive folder.
        
        Args:
            source_folder (str): Path to the folder containing files to process
            excel_output_path (str): Path where the Excel file will be saved
            archive_folder (str): Path to the folder where processed files will be moved
        """
        self.source_folder = Path(source_folder)
        self.excel_output_path = Path(excel_output_path)
        self.archive_folder = Path(archive_folder)
        
        self.archive_folder.mkdir(parents=True, exist_ok=True)
        self.excel_output_path.parent.mkdir(parents=True, exist_ok=True)
    
    def get_files_from_folder(self) -> List[Path]:
        """
        Get all files from the source folder (excluding directories).
        
        Returns:
            List[Path]: List of file paths in the source folder
        """
        if not self.source_folder.exists():
            raise FileNotFoundError(f"Source folder '{self.source_folder}' does not exist")
        
        files = [f for f in self.source_folder.iterdir() if f.is_file()]
        return files
    
    def split_filename(self, filename: str) -> List[str]:
        """
        Split filename at underscore delimiters and return up to 4 parts.
        
        Args:
            filename (str): The filename to split (without extension)
            
        Returns:
            List[str]: List of split parts, padded to 4 elements with empty strings
        """
        name_without_ext = Path(filename).stem
        
        parts = name_without_ext.split('_')
        
        while len(parts) < 4:
            parts.append('')
        
        return parts[:4]
    
    def process_files_to_excel(self) -> Tuple[int, List[str]]:
        """
        Process all files in the source folder and create an Excel file.
        
        Returns:
            Tuple[int, List[str]]: Number of files processed and list of processed filenames
        """
        files = self.get_files_from_folder()
        
        if not files:
            print(f"No files found in '{self.source_folder}'")
            return 0, []
        
        data = []
        processed_files = []
        
        for file_path in files:
            filename = file_path.name
            split_parts = self.split_filename(filename)
            data.append(split_parts)
            processed_files.append(filename)
        
        df = pd.DataFrame(data, columns=['Col1', 'Col2', 'Col3', 'Col4'])
        
        df.to_excel(self.excel_output_path, index=False)
        
        print(f"Excel file created successfully at: {self.excel_output_path}")
        print(f"Processed {len(files)} files")
        
        return len(files), processed_files
    
    def move_files_to_archive(self, processed_files: List[str]) -> int:
        """
        Move processed files from source folder to archive folder.
        
        Args:
            processed_files (List[str]): List of filenames that were processed
            
        Returns:
            int: Number of files successfully moved
        """
        moved_count = 0
        
        for filename in processed_files:
            source_file = self.source_folder / filename
            archive_file = self.archive_folder / filename
            
            try:
                counter = 1
                original_archive_file = archive_file
                while archive_file.exists():
                    name_part = original_archive_file.stem
                    ext_part = original_archive_file.suffix
                    archive_file = self.archive_folder / f"{name_part}_{counter}{ext_part}"
                    counter += 1
                
                shutil.move(str(source_file), str(archive_file))
                moved_count += 1
                print(f"Moved: {filename} -> {archive_file.name}")
                
            except Exception as e:
                print(f"Error moving {filename}: {str(e)}")
        
        print(f"Successfully moved {moved_count} files to archive folder")
        return moved_count
    
    def process_all(self) -> dict:
        """
        Complete processing workflow: create Excel file and move files to archive.
        
        Returns:
            dict: Summary of processing results
        """
        try:
            files_processed, processed_filenames = self.process_files_to_excel()
            
            if files_processed == 0:
                return {
                    'success': True,
                    'files_processed': 0,
                    'files_moved': 0,
                    'excel_path': str(self.excel_output_path),
                    'message': 'No files found to process'
                }
            
            files_moved = self.move_files_to_archive(processed_filenames)
            
            return {
                'success': True,
                'files_processed': files_processed,
                'files_moved': files_moved,
                'excel_path': str(self.excel_output_path),
                'archive_path': str(self.archive_folder),
                'message': f'Successfully processed {files_processed} files and moved {files_moved} to archive'
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'message': f'Error during processing: {str(e)}'
            }


def main():
    """
    Example usage of the FileProcessor class.
    Customize the paths according to your needs.
    """
    source_folder = "./source_files"  # Folder containing files to process
    excel_output_path = "./output/processed_files.xlsx"  # Path for the Excel output
    archive_folder = "./archive"  # Folder to move processed files
    
    processor = FileProcessor(source_folder, excel_output_path, archive_folder)
    
    result = processor.process_all()
    
    print("\n" + "="*50)
    print("PROCESSING SUMMARY")
    print("="*50)
    
    if result['success']:
        print(f"✓ {result['message']}")
        if result['files_processed'] > 0:
            print(f"✓ Excel file saved to: {result['excel_path']}")
            print(f"✓ Files archived to: {result['archive_path']}")
    else:
        print(f"✗ {result['message']}")
    
    return result


if __name__ == "__main__":
    main()
