"""
Kohler Image Automation - GUI Application
Desktop application for inserting product images into Excel workbook.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pathlib import Path

# Import existing automation modules
from python.main import ExcelAutomation
from python.gui_worker import WorkerThread


class KohlerImageAutomationGUI:
    """
    Main GUI window for Kohler Image Automation
    Provides user-friendly interface for image insertion
    """

    def __init__(self):
        """Initialize the GUI application"""
        self.root = tk.Tk()
        self.root.title("Kohler Image Automation")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f7fa")
        
        # Set app icon
        icon_path = Path(__file__).parent.parent / "assets" / "kohler_icon.png"
        if icon_path.exists():
            try:
                icon = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, icon)
            except:
                pass  # If icon fails to load, just use default
        
        # Set minimum window size
        self.root.minsize(800, 650)

        # Variables to store the file paths
        self.excel_path = tk.StringVar()
        self.pdf_path = tk.StringVar()
        self.save_as_new = tk.BooleanVar(value=False)

        # Track processing state
        self.is_processing = False

        # Build the user interface
        self.setup_ui()

    def setup_ui(self):
        """Create all UI elements."""

        # ===== HEADER SECTION =====
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=110)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)

        # Company/App name
        title_label = tk.Label(
            header_frame,
            text="KOHLER",
            font=("Helvetica", 16, "bold"),
            bg="#2c3e50",
            fg="#ecf0f1"
        )
        title_label.pack(pady=(15, 0))
        
        # Main title
        main_title = tk.Label(
            header_frame,
            text="Product Image Automation",
            font=("Helvetica", 24, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        main_title.pack(pady=(5, 5))
        
        # Subtitle
        subtitle_label = tk.Label(
            header_frame,
            text="Seamlessly insert product images into Excel workbooks from PDF catalogs",
            font=("Helvetica", 11),
            bg="#2c3e50",
            fg="#95a5a6"
        )
        subtitle_label.pack(pady=(0, 15))

        # ===== FILE SELECTION SECTION ======
        file_frame = tk.LabelFrame(
            self.root,
            text="",
            font=("Helvetica", 13, "bold"),
            padx=30,
            pady=25,
            bg="white",
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#dfe6e9",
            highlightcolor="#dfe6e9"
        )
        file_frame.pack(fill=tk.X, padx=30, pady=(30, 15))
        
        # Section title
        section_title = tk.Label(
            file_frame,
            text="Select Input Files",
            font=("Helvetica", 14, "bold"),
            bg="white",
            fg="#2c3e50"
        )
        section_title.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 20))

        # Excel file selection
        excel_label = tk.Label(
            file_frame, 
            text="Excel Workbook (.xlsx, .xlsm)", 
            font=("Helvetica", 11),
            bg="white",
            fg="#34495e"
        )
        excel_label.grid(row=1, column=0, sticky=tk.W, pady=10)

        excel_entry = tk.Entry(
            file_frame,
            textvariable=self.excel_path,
            width=50,
            font=("Helvetica", 11),
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#bdc3c7",
            highlightcolor="#3498db"
        )
        excel_entry.grid(row=1, column=1, padx=20, pady=10, sticky=tk.EW)

        excel_button = tk.Button(
            file_frame,
            text="Browse Files",
            command=self.browse_excel,
            width=14,
            bg="#3498db",
            fg="white",
            font=("Helvetica", 10, "bold"),
            relief=tk.SOLID,
            borderwidth=1,
            cursor="hand2",
            activebackground="#2980b9",
            activeforeground="white",
            pady=8,
            highlightthickness=0
        )
        excel_button.grid(row=1, column=2, pady=10, padx=(0, 5))

        # PDF file selection
        pdf_label = tk.Label(
            file_frame, 
            text="PDF Pricebook (.pdf)", 
            font=("Helvetica", 11),
            bg="white",
            fg="#34495e"
        )
        pdf_label.grid(row=2, column=0, sticky=tk.W, pady=10)

        pdf_entry = tk.Entry(
            file_frame,
            textvariable=self.pdf_path,
            width=50,
            font=("Helvetica", 11),
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#bdc3c7",
            highlightcolor="#3498db"
        )
        pdf_entry.grid(row=2, column=1, padx=20, pady=10, sticky=tk.EW)

        pdf_button = tk.Button(
            file_frame,
            text="Browse Files",
            command=self.browse_pdf,
            width=14,
            bg="#3498db",
            fg="white",
            font=("Helvetica", 10, "bold"),
            relief=tk.SOLID,
            borderwidth=1,
            cursor="hand2",
            activebackground="#2980b9",
            activeforeground="white",
            pady=8,
            highlightthickness=0
        )
        pdf_button.grid(row=2, column=2, pady=10, padx=(0, 5))
        
        # Make column 1 expandable
        file_frame.columnconfigure(1, weight=1)

        # ===== OPTIONS SECTION =====
        options_frame = tk.Frame(
            self.root,
            bg="white",
            highlightthickness=1,
            highlightbackground="#dfe6e9",
            highlightcolor="#dfe6e9"
        )
        options_frame.pack(fill=tk.X, padx=30, pady=15)

        # Options title
        options_title = tk.Label(
            options_frame,
            text="Output Options",
            font=("Helvetica", 14, "bold"),
            bg="white",
            fg="#2c3e50"
        )
        options_title.pack(anchor=tk.W, padx=30, pady=(20, 15))

        # Save as Checkbox with better styling
        checkbox_frame = tk.Frame(options_frame, bg="white")
        checkbox_frame.pack(anchor=tk.W, padx=30, pady=(0, 10))
        
        save_as_checkbox = tk.Checkbutton(
            checkbox_frame,
            text="Save as new file (preserves original)",
            variable=self.save_as_new,
            font=("Helvetica", 11, "bold"),
            bg="white",
            activebackground="white",
            selectcolor="#ecf0f1",
            fg="#2c3e50"
        )
        save_as_checkbox.pack(side=tk.LEFT)

        # Help text with icon
        help_frame = tk.Frame(options_frame, bg="#ecf0f1")
        help_frame.pack(fill=tk.X, padx=30, pady=(0, 20))
        
        help_label = tk.Label(
            help_frame,
            text="  UNCHECKED: Overwrites original file (faster) | CHECKED: Creates new file with '_with_images' suffix (safer for order changes)",
            font=("Helvetica", 10),
            fg="#7f8c8d",
            justify=tk.LEFT,
            bg="#ecf0f1",
            pady=8,
            padx=10
        )
        help_label.pack(fill=tk.X)

        # ===== ACTION BUTTONS SECTION =====
        button_frame = tk.Frame(self.root, pady=20, bg="#f5f7fa")
        button_frame.pack(fill=tk.X, padx=30)

        # Fill Images Button - Large and prominent
        self.fill_images_button = tk.Button(
            button_frame,
            text="START IMAGE PROCESSING",
            command=self.run_fill_images,
            bg="#27ae60",
            fg="white",
            font=("Helvetica", 15, "bold"),
            cursor="hand2",
            relief=tk.RAISED,
            borderwidth=2,
            activebackground="#229954",
            activeforeground="white",
            pady=18,
            padx=40,
            highlightthickness=0
        )
        self.fill_images_button.pack(expand=True, fill=tk.X, ipady=5)

        # ===== PROGRESS SECTION =====
        progress_frame = tk.Frame(
            self.root,
            bg="white",
            highlightthickness=1,
            highlightbackground="#dfe6e9",
            highlightcolor="#dfe6e9"
        )
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(15, 20))

        # Progress title
        progress_title = tk.Label(
            progress_frame,
            text="Activity Log",
            font=("Helvetica", 14, "bold"),
            bg="white",
            fg="#2c3e50"
        )
        progress_title.pack(anchor=tk.W, padx=30, pady=(20, 10))

        # Status text with icon
        status_frame = tk.Frame(progress_frame, bg="white")
        status_frame.pack(anchor=tk.W, padx=30, pady=(0, 10), fill=tk.X)
        
        self.status_label = tk.Label(
            status_frame,
            text="Ready to process files",
            font=("Helvetica", 11, "bold"),
            fg="#27ae60",
            bg="white",
            anchor=tk.W
        )
        self.status_label.pack(side=tk.LEFT)
        
        # Progress Bar - modern flat design
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Modern.Horizontal.TProgressbar", 
                       thickness=8,
                       background='#27ae60',
                       troughcolor='#ecf0f1',
                       borderwidth=0,
                       relief='flat')
        
        progress_container = tk.Frame(progress_frame, bg="white")
        progress_container.pack(fill=tk.X, padx=30, pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(
            progress_container,
            mode="determinate",
            style="Modern.Horizontal.TProgressbar",
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X)
        
        # Progress percentage label
        self.progress_percent = tk.Label(
            progress_container,
            text="0%",
            font=("Helvetica", 9),
            fg="#7f8c8d",
            bg="white"
        )
        self.progress_percent.pack(pady=(5, 0))

        # Log output (scrollable text area) with modern design
        log_container = tk.Frame(progress_frame, bg="white")
        log_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 20))
        
        log_scroll = tk.Scrollbar(log_container, width=12)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(
            log_container,
            height=8,
            font=("Consolas", 10),
            bg="#2c3e50",
            fg="#ecf0f1",
            wrap=tk.WORD,
            yscrollcommand=log_scroll.set,
            relief=tk.FLAT,
            borderwidth=0,
            padx=15,
            pady=12,
            insertbackground="#ecf0f1"
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)

        # Make log read-only
        self.log_text.config(state=tk.DISABLED)

        # ===== STATUS BAR =====
        self.status_bar = tk.Label(
            self.root,
            text="Ready | Kohler Image Automation v1.0",
            relief=tk.FLAT,
            anchor=tk.W,
            font=("Helvetica", 10),
            bg="#34495e",
            fg="#ecf0f1",
            padx=15,
            pady=8
        )
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def browse_excel(self):
        """Open file dialog to select Excel workbook."""
        filename = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls *.xlsb"), ("All Files", "*.*")],
            initialdir=os.path.expanduser("~/Desktop")  
        )
        if filename:
            self.excel_path.set(filename)
            self.log(f"Selected Excel file: {os.path.basename(filename)}")

    def browse_pdf(self):
        """Open file dialog to select PDF pricebook."""
        filename = filedialog.askopenfilename(
            title="Select PDF Pricebook",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            initialdir=os.path.expanduser("~/Desktop")  
        )
        if filename:
            self.pdf_path.set(filename)
            self.log(f"Selected PDF file: {os.path.basename(filename)}")

    def log(self, message):
        """
        Add a message to the log text area.
        Args:
            message: text to display in the log
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def clear_log(self):
        """Clear all text from the log area."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def validate_files(self, require_pdf=False):
        """
        Validate that required files are selected and exist.

        Args:
            require_pdf: Whether PDF file is required for this operation

        Returns:
            bool: True if validation passes, False otherwise
        """
        excel = self.excel_path.get()
        pdf = self.pdf_path.get()

        # Check Excel file
        if not excel:
            messagebox.showerror("Missing File", "Please select an Excel workbook.")
            return False
        
        if not os.path.exists(excel):
            messagebox.showerror("File Not Found", f"Excel file not found:\n{excel}")
            return False
        
        # Check Excel file format
        excel_ext = os.path.splitext(excel)[1].lower()
        if excel_ext == '.xlsb':
            messagebox.showerror(
                "Unsupported File Format",
                "Excel Binary Workbook (.xlsb) format is not supported.\n\n"
                "Please convert your file to one of these formats:\n"
                "• .xlsx (Excel Workbook) - Recommended\n"
                "• .xlsm (Excel Macro-Enabled Workbook)\n\n"
                "To convert:\n"
                "1. Open the file in Excel\n"
                "2. File → Save As\n"
                "3. Choose .xlsx format\n\n"
                "Note: Use .xlsx format for best compatibility.\n"
                "VBA macros can be run separately after image insertion."
            )
            return False
        
        if excel_ext not in ['.xlsx', '.xlsm']:
            messagebox.showwarning(
                "File Format Warning",
                f"File extension '{excel_ext}' may not be supported.\n\n"
                "Recommended formats: .xlsx or .xlsm"
            )
        
        # Check PDF file (if required)
        if require_pdf:
            if not pdf:
                messagebox.showerror("Missing File", "Please select a PDF pricebook.")
                return False
            
            if not os.path.exists(pdf):
                messagebox.showerror("File Not Found", f"PDF file not found:\n{pdf}")
                return False
        
        return True
    
    def set_processing_state(self, is_processing):
        """
        Update UI to show processing or idle state.

        Args:
            is_processing: True if operation is running, False if idle
        """
        self.is_processing = is_processing

        # Disable/Enable button
        state = tk.DISABLED if is_processing else tk.NORMAL
        self.fill_images_button.config(state=state)

        # Show/hide progress bar
        if is_processing:
            self.progress_bar['value'] = 0
            self.progress_percent.config(text="0%")
            self.status_bar.config(text="Processing... Please wait", bg="#f39c12", fg="white")
        else:
            self.progress_bar['value'] = 0
            self.progress_percent.config(text="0%")
            self.status_bar.config(text="Ready | Kohler Image Automation v1.0", bg="#34495e", fg="#ecf0f1")
    
    def update_progress(self, current, total, message=""):
        """
        Update progress bar and status.
        
        Args:
            current: Current progress value
            total: Total value for completion
            message: Optional status message
        """
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_bar['value'] = percentage
            self.progress_percent.config(text=f"{percentage}%")
            
            if message:
                self.status_label.config(text=message)
            else:
                self.status_label.config(text=f"Processing... {current}/{total}")
            
            # Force GUI update
            self.root.update_idletasks()

    def run_fill_images(self):
        """Execute the Fill Images operation in background thread."""

        # Validate files
        if not self.validate_files(require_pdf=True):
            return
        
        # check if already processing
        if self.is_processing:
            messagebox.showwarning("Operation in Progress", "Please wait for the current operation to finish.")
            return
        
        # Clear previous log
        self.clear_log()
        self.log("=" * 60)
        self.log("FILL IMAGES OPERATION")
        self.log("=" * 60)

        # Determine output path
        excel_path = self.excel_path.get()
        pdf_path = self.pdf_path.get()

        if self.save_as_new.get():
            # Create new filename with suffix
            path = Path(excel_path)
            output_path = path.parent / f"{path.stem}_with_images{path.suffix}"
            self.log(f"Output mode: Save as new file")
            self.log(f"Output path: {output_path.name}")
        else:
            output_path = excel_path
            self.log(f"Output mode: Overwrite original")

        # If saving as new, copy original file first
        if self.save_as_new.get():
            import shutil
            try:
                shutil.copy2(excel_path, output_path)
                self.log(f"Copied original file to: {output_path.name}")
            except Exception as e:
                messagebox.showerror("File Copy Error", f"Failed to copy file:\n{str(e)}")
                return
        
        # Define progress callback
        def on_progress(current, total, message=""):
            self.root.after(0, self.update_progress, current, total, message)
        
        # Define success callback
        def on_success(output):
            self.log(output)
            self.progress_bar['value'] = 100
            self.progress_percent.config(text="100%")
            self.status_label.config(text="Completed successfully!", fg="#27ae60")
            self.set_processing_state(False)
            messagebox.showinfo(
                "Success",
                f"Images inserted successfully!\n\nOutput file:\n{os.path.basename(str(output_path))}"
            )

        # Define error callback
        def on_error(error_msg):
            self.log(error_msg)
            self.status_label.config(text="Operation failed", fg="#e74c3c")
            self.set_processing_state(False)
            messagebox.showerror("Error", "Operation failed. Check log for details.")
            
        # Create automation instance
        automation = ExcelAutomation(progress_callback=on_progress)

        # Update status
        self.status_label.config(text="Processing images... Please wait", fg="#f39c12")
        self.set_processing_state(True)

        # Define task function
        def task():
            return automation.fill_images_from_pdf(str(output_path), pdf_path)

        # Start background thread
        worker = WorkerThread(task, on_success, on_error, on_progress)
        worker.start()

    def run(self):
        """Start the GUI main loop."""
        self.root.mainloop()

def main():
    """Start the GUI application."""
    app = KohlerImageAutomationGUI()
    app.run()

if __name__ == "__main__":
    main()

