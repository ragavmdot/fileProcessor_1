# File Processor

A Python utility to process files in a folder by splitting their names at underscore delimiters and exporting the results to an Excel file, then archiving the processed files.

## Features

- Reads all files from a specified source folder
- Splits filenames at underscore (`_`) delimiters
- Creates an Excel file with fixed column headers: `Col1`, `Col2`, `Col3`, `Col4`
- Each row represents the split words from one filename
- Moves processed files to a customizable archive folder
- Handles file name conflicts in the archive folder
- Customizable paths for source, output, and archive locations

## Installation

1. Clone this repository:
```bash
git clone https://github.com/ragavmdot/fileProcessor_1.git
cd fileProcessor_1
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```python
from file_processor import FileProcessor

# Initialize the processor with your custom paths
processor = FileProcessor(
    source_folder="./source_files",           # Folder with files to process
    excel_output_path="./output/results.xlsx", # Where to save Excel file
    archive_folder="./archive"                # Where to move processed files
)

# Process all files
result = processor.process_all()
print(result['message'])
```

### Advanced Usage

```python
from file_processor import FileProcessor

# Step-by-step processing
processor = FileProcessor(
    source_folder="/path/to/your/files",
    excel_output_path="/path/to/output/report.xlsx",
    archive_folder="/path/to/archive"
)

try:
    # Get list of files
    files = processor.get_files_from_folder()
    print(f"Found {len(files)} files")
    
    # Create Excel file
    count, filenames = processor.process_files_to_excel()
    print(f"Processed {count} files to Excel")
    
    # Move files to archive
    moved = processor.move_files_to_archive(filenames)
    print(f"Moved {moved} files to archive")
    
except Exception as e:
    print(f"Error: {e}")
```

### Running the Example

```bash
python example_usage.py
```

This will create sample files and demonstrate the processing workflow.

## How It Works

### Filename Processing

The script processes filenames by:

1. **Removing file extensions**: `report_2024_january_sales.txt` → `report_2024_january_sales`
2. **Splitting at underscores**: `report_2024_january_sales` → `['report', '2024', 'january', 'sales']`
3. **Mapping to columns**: 
   - Col1: `report`
   - Col2: `2024` 
   - Col3: `january`
   - Col4: `sales`

### Excel Output

The Excel file will have this structure:

| Col1     | Col2 | Col3      | Col4    |
|----------|------|-----------|---------|
| report   | 2024 | january   | sales   |
| data     | analysis | quarterly | summary |
| user     | feedback | survey  | results |
| inventory| stock | levels   |         |

### File Archiving

After creating the Excel file, all processed files are moved to the archive folder. If a file with the same name already exists in the archive, it will be renamed with a counter (e.g., `file_1.txt`, `file_2.txt`).

## Example File Structure

```
project/
├── source_files/           # Your files to process
│   ├── report_2024_january_sales.txt
│   ├── data_analysis_quarterly_summary.csv
│   └── user_feedback_survey_results.pdf
├── output/
│   └── processed_files.xlsx  # Generated Excel file
└── archive/                # Processed files moved here
    ├── report_2024_january_sales.txt
    ├── data_analysis_quarterly_summary.csv
    └── user_feedback_survey_results.pdf
```

## Customization

### Paths

All paths are fully customizable:

```python
processor = FileProcessor(
    source_folder="/home/user/documents/files",
    excel_output_path="/home/user/desktop/analysis.xlsx", 
    archive_folder="/home/user/backup/processed"
)
```

### Column Headers

The Excel file uses fixed headers (`Col1`, `Col2`, `Col3`, `Col4`). If you need different headers, modify the `process_files_to_excel` method in `file_processor.py`.

### Number of Columns

Currently supports up to 4 columns. Files with fewer underscore-separated parts will have empty cells in the remaining columns. Files with more than 3 underscores will only use the first 4 parts.

## Requirements

- Python 3.6+
- pandas >= 1.5.0
- openpyxl >= 3.0.0

## Error Handling

The script handles common errors:

- **Source folder doesn't exist**: Raises `FileNotFoundError`
- **Permission issues**: Catches and reports file operation errors
- **File conflicts**: Automatically renames files in archive folder
- **Empty folders**: Reports when no files are found to process

## Return Values

The `process_all()` method returns a dictionary with processing results:

```python
{
    'success': True,
    'files_processed': 5,
    'files_moved': 5,
    'excel_path': '/path/to/output.xlsx',
    'archive_path': '/path/to/archive',
    'message': 'Successfully processed 5 files and moved 5 to archive'
}
```

## License

This project is open source and available under the MIT License.
