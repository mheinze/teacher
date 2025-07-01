#!/usr/bin/env python3
"""
Batch processing script for AIG Class List Processor
Handles multiple PDF, Excel, and Word document file combinations
"""

import os
import glob
from aig_processor import AIGClassListProcessor
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def find_files():
    """
    Find all PDF, Excel, and Word files in the input directory
    
    Returns:
        tuple: (pdf_files, excel_files, word_files)
    """
    # Look in input directory for files
    input_dir = "input"
    if not os.path.exists(input_dir):
        input_dir = "."  # Fallback to current directory
    
    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    excel_files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    word_files = glob.glob(os.path.join(input_dir, "*.docx"))
    
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]  # Exclude temp files
    word_files = [f for f in word_files if not os.path.basename(f).startswith('~$')]   # Exclude temp files
    
    return pdf_files, excel_files, word_files

def batch_process():
    """
    Process all combinations of PDF, Excel, and Word files found
    Updated for new single-file output format
    """
    pdf_files, excel_files, word_files = find_files()
    
    if not pdf_files:
        logger.error("No PDF files found in input directory")
        return
    
    if not excel_files:
        logger.error("No Excel files found in input directory")
        return
    
    logger.info(f"Found {len(pdf_files)} PDF files, {len(excel_files)} Excel files, and {len(word_files)} Word files")
    
    for pdf_file in pdf_files:
        for excel_file in excel_files:
            # Find matching Word file (optional)
            word_file = word_files[0] if word_files else None
            
            logger.info(f"Processing: {pdf_file} + {excel_file}" + (f" + {word_file}" if word_file else ""))
            
            # Create safe output directory name
            pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
            excel_name = os.path.splitext(os.path.basename(excel_file))[0]
            output_dir = f"output/batch_{pdf_name}_{excel_name}".replace(" ", "_")
            
            try:
                processor = AIGClassListProcessor(pdf_file, excel_file, word_file, output_dir)
                processor.process()
                logger.info(f"✅ Successfully processed {pdf_file} + {excel_file}" + (f" + {word_file}" if word_file else ""))
                
            except Exception as e:
                logger.error(f"❌ Error processing {pdf_file} + {excel_file}: {e}")
                continue

def main():
    """
    Main function for batch processing
    """
    print("🔄 AIG Class List Processor - Batch Mode")
    print("=" * 50)
    print("This will process all PDF and Excel files in the current directory")
    print()
    
    # Show what files will be processed
    pdf_files, excel_files, word_files = find_files()
    
    print("PDF files found:")
    for i, pdf in enumerate(pdf_files, 1):
        print(f"  {i}. {pdf}")
    
    print(f"\nExcel files found:")
    for i, excel in enumerate(excel_files, 1):
        print(f"  {i}. {excel}")
    
    print(f"\nWord files found:")
    for i, word in enumerate(word_files, 1):
        print(f"  {i}. {word}")
    
    print(f"\nThis will create {len(pdf_files) * len(excel_files)} processing combinations.")
    
    # Ask for confirmation
    response = input("\nProceed with batch processing? (y/N): ").strip().lower()
    
    if response in ['y', 'yes']:
        batch_process()
        print("\n🎉 Batch processing complete!")
    else:
        print("Batch processing cancelled.")

if __name__ == "__main__":
    main()
