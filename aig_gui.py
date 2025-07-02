#!/usr/bin/env python3
"""
macOS GUI Application for AIG Class List Processor

This creates a simple macOS application that prompts for PDF, Word, and Excel documents
and generates the AIG output files as specified in prompt.md.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import os
import sys
import threading
from aig_processor import AIGClassListProcessor
import logging

# Configure logging for GUI
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AIGProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AIG Class List Processor")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # File paths
        self.pdf_file = tk.StringVar()
        self.excel_file = tk.StringVar()
        self.word_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        
        # Processing status
        self.processing_success = False
        
        # Set default output directory
        here = Path.cwd()
        self.output_dir.set(self.find_directory_upward())
        
        self.setup_ui()

    def find_directory_upward(self):
        current = Path.cwd()
        while current != current.parent:
            potential_target = current / 'output'
            if potential_target.is_dir():
               return potential_target
            current = current.parent
        return Path.cwd()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="AIG Class List Processor", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # PDF file selection
        ttk.Label(main_frame, text="PDF File (AIG Roster):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.pdf_file, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_pdf).grid(row=1, column=2, pady=5)
        
        # Excel file selection
        ttk.Label(main_frame, text="Excel File (Class Lists):").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_file, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_excel).grid(row=2, column=2, pady=5)
        
        # Word file selection (optional)
        ttk.Label(main_frame, text="Word File (Optional):").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.word_file, width=50).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_word).grid(row=3, column=2, pady=5)
        
        # Output directory selection
        ttk.Label(main_frame, text="Output Directory:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_output_dir).grid(row=4, column=2, pady=5)
        
        # Process button
        self.process_button = ttk.Button(main_frame, text="Process Files", command=self.process_files)
        self.process_button.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Buttons frame for completion actions (initially hidden)
        self.completion_frame = ttk.Frame(main_frame)
        self.completion_frame.grid(row=5, column=0, columnspan=3, pady=20)
        self.completion_frame.grid_remove()  # Hide initially
        
        # New Process button
        self.new_process_button = ttk.Button(self.completion_frame, text="Process New Files", command=self.new_process)
        self.new_process_button.grid(row=0, column=0, padx=10)
        
        # Close button
        self.close_button = ttk.Button(self.completion_frame, text="Close Application", command=self.close_application)
        self.close_button.grid(row=0, column=1, padx=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status text
        self.status_text = tk.Text(main_frame, height=15, width=70)
        self.status_text.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=7, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(7, weight=1)
        
        # Initial status
        self.log_message("Ready to process AIG class lists...")
        self.log_message("Please select PDF, Excel, and optionally Word files.")
        
    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Select PDF File (AIG Roster)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_file.set(filename)
            self.log_message(f"Selected PDF: {os.path.basename(filename)}")
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File (Class Lists)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            self.log_message(f"Selected Excel: {os.path.basename(filename)}")
    
    def browse_word(self):
        filename = filedialog.askopenfilename(
            title="Select Word File (Optional AIG List)",
            filetypes=[("Word files", "*.docx *.doc"), ("All files", "*.*")]
        )
        if filename:
            self.word_file.set(filename)
            self.log_message(f"Selected Word: {os.path.basename(filename)}")
    
    def browse_output_dir(self):
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir.set(dirname)
            self.log_message(f"Output directory: {dirname}")
    
    def log_message(self, message):
        """Add a message to the status text widget"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def validate_inputs(self):
        """Validate that required files are selected"""
        if not self.pdf_file.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return False
        
        if not self.excel_file.get():
            messagebox.showerror("Error", "Please select an Excel file")
            return False
        
        if not os.path.exists(self.pdf_file.get()):
            messagebox.showerror("Error", f"PDF file not found: {self.pdf_file.get()}")
            return False
        
        if not os.path.exists(self.excel_file.get()):
            messagebox.showerror("Error", f"Excel file not found: {self.excel_file.get()}")
            return False
        
        if self.word_file.get() and not os.path.exists(self.word_file.get()):
            messagebox.showerror("Error", f"Word file not found: {self.word_file.get()}")
            return False
        
        return True
    
    def process_files(self):
        """Process the selected files"""
        if not self.validate_inputs():
            return
        
        # Disable the process button and start progress bar
        self.process_button.config(state='disabled')
        self.progress.start()
        
        # Run processing in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        """Run the AIG processing in a separate thread"""
        try:
            self.log_message("\n" + "="*50)
            self.log_message("Starting AIG Class List Processing...")
            self.log_message("="*50)
            
            # Get file paths
            pdf_path = self.pdf_file.get()
            excel_path = self.excel_file.get()
            word_path = self.word_file.get() if self.word_file.get() else None
            output_path = self.output_dir.get()
            
            self.log_message(f"PDF File: {os.path.basename(pdf_path)}")
            self.log_message(f"Excel File: {os.path.basename(excel_path)}")
            if word_path:
                self.log_message(f"Word File: {os.path.basename(word_path)}")
            else:
                self.log_message("Word File: None (optional)")
            self.log_message(f"Output Directory: {output_path}")
            
            # Create processor and run
            processor = AIGClassListProcessor(pdf_path, excel_path, word_path, output_path)
            
            # Redirect logging to our GUI
            class GUILogHandler(logging.Handler):
                def __init__(self, gui):
                    super().__init__()
                    self.gui = gui
                
                def emit(self, record):
                    msg = self.format(record)
                    self.gui.log_message(msg)
            
            # Add GUI handler
            gui_handler = GUILogHandler(self)
            gui_handler.setLevel(logging.INFO)
            logger.addHandler(gui_handler)
            
            # Process files
            processor.process()
            
            # Remove GUI handler
            logger.removeHandler(gui_handler)
            
            # Success message
            self.log_message("\n" + "="*50)
            self.log_message("✅ Processing completed successfully!")
            self.log_message(f"📁 Output files saved to: {output_path}")
            self.log_message("📊 Files generated:")
            self.log_message("  • updated_class_lists.xlsx (Main file with AIG columns)")
            self.log_message("  • updated_class_lists_AIG_Only.xlsx (AIG students only)")
            self.log_message("  • students_not_in_excel.xlsx (Missing students)")
            self.log_message("  • aig_statistics_report.md (Statistics report)")
            self.log_message("="*50)
            
            # Show success dialog
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", 
                f"Processing completed successfully!\n\nOutput files saved to:\n{output_path}"
            ))
            
            # Mark as successful completion
            self.processing_success = True
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.log_message(f"\n❌ {error_msg}")
            logger.error(error_msg, exc_info=True)
            
            # Show error dialog
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            
            # Mark as failed completion
            self.processing_success = False
        
        finally:
            # Re-enable the process button and stop progress bar
            self.root.after(0, self.processing_complete)
    
    def processing_complete(self):
        """Called when processing is complete"""
        self.progress.stop()
        
        if self.processing_success:
            # Processing completed successfully - show completion options
            self.process_button.grid_remove()  # Hide the process button
            self.completion_frame.grid()  # Show the completion buttons
            self.log_message("\n🎉 Processing complete! You can process new files or close the application.")
        else:
            # Processing failed - re-enable process button for retry
            self.process_button.config(state='normal')
            self.log_message("\n⚠️ Processing failed. Please check your files and try again.")
    
    def new_process(self):
        """Reset the interface for a new processing session"""
        # Hide completion buttons and show process button
        self.completion_frame.grid_remove()
        self.process_button.grid()
        self.process_button.config(state='normal')
        
        # Clear the status text
        self.status_text.delete(1.0, tk.END)
        
        # Reset processing status
        self.processing_success = False
        
        # Add initial status message
        self.log_message("Ready to process new AIG class lists...")
        self.log_message("Please select PDF, Excel, and optionally Word files.")
    
    def close_application(self):
        """Close the application"""
        self.root.quit()
        self.root.destroy()

def main():
    """Main function to run the GUI application"""
    # Create the root window
    root = tk.Tk()
    
    # Create the application
    app = AIGProcessorGUI(root)
    
    # Set window icon (if available)
    try:
        # Try to set a reasonable window icon for macOS
        if sys.platform == "darwin":  # macOS
            root.call('wm', 'iconphoto', root._w, tk.PhotoImage())
    except:
        pass  # Ignore if icon setting fails
    
    # Center the window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    # Start the GUI event loop
    try:
        root.mainloop()
    except KeyboardInterrupt:
        root.quit()

if __name__ == "__main__":
    main()
