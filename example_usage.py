#!/usr/bin/env python3
"""
Example script showing how to use the FileProcessor class with custom paths.
"""

from file_processor import FileProcessor
import os


def example_usage():
    """
    Example demonstrating how to use FileProcessor with custom paths.
    """
    
    print("Example 1: Basic usage")
    print("-" * 30)
    
    processor1 = FileProcessor(
        source_folder="./test_files",
        excel_output_path="./output/file_analysis.xlsx", 
        archive_folder="./processed_archive"
    )
    
    result1 = processor1.process_all()
    print(f"Result: {result1['message']}\n")
    
    
    print("Example 2: Absolute paths")
    print("-" * 30)
    
    home_dir = os.path.expanduser("~")
    processor2 = FileProcessor(
        source_folder=f"{home_dir}/Documents/files_to_process",
        excel_output_path=f"{home_dir}/Desktop/file_report.xlsx",
        archive_folder=f"{home_dir}/Documents/archived_files"
    )
    
    result2 = processor2.process_all()
    print(f"Result: {result2['message']}\n")
    
    
    print("Example 3: Step-by-step processing")
    print("-" * 30)
    
    processor3 = FileProcessor(
        source_folder="./sample_data",
        excel_output_path="./reports/detailed_analysis.xlsx",
        archive_folder="./backup"
    )
    
    try:
        files = processor3.get_files_from_folder()
        print(f"Found {len(files)} files to process")
        
        count, filenames = processor3.process_files_to_excel()
        print(f"Created Excel file with {count} entries")
        
        moved = processor3.move_files_to_archive(filenames)
        print(f"Moved {moved} files to archive")
        
    except FileNotFoundError as e:
        print(f"Source folder not found: {e}")
    except Exception as e:
        print(f"Error: {e}")


def create_sample_files():
    """
    Create some sample files for testing the FileProcessor.
    """
    import os
    
    test_dir = "./test_files"
    os.makedirs(test_dir, exist_ok=True)
    
    sample_files = [
        "report_2024_january_sales.txt",
        "data_analysis_quarterly_summary.csv", 
        "user_feedback_survey_results.pdf",
        "inventory_stock_levels.xlsx",
        "meeting_notes_team_sync.docx",
        "project_timeline_milestone_review.pptx",
        "budget_forecast_annual.txt",
        "customer_complaints_resolved.csv"
    ]
    
    for filename in sample_files:
        file_path = os.path.join(test_dir, filename)
        with open(file_path, 'w') as f:
            f.write(f"Sample content for {filename}")
    
    print(f"Created {len(sample_files)} sample files in {test_dir}")
    return test_dir


if __name__ == "__main__":
    print("FileProcessor Example Usage")
    print("=" * 50)
    
    sample_dir = create_sample_files()
    print(f"Sample files created in: {sample_dir}\n")
    
    example_usage()
