#!/usr/bin/env python3
"""
MD to DOCX Converter
A simple desktop application to convert Markdown files to DOCX format using pandoc.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import sys
from pathlib import Path
import threading


class MarkdownToDocxConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("MD to DOCX Converter")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Configure style
        self.setup_styles()
        
        # Create GUI elements
        self.create_widgets()
        
        # Check if pandoc is available
        self.check_pandoc()
    
    def setup_styles(self):
        """Configure the application styling"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Success.TLabel', foreground='green')
        style.configure('Error.TLabel', foreground='red')
    
    def create_widgets(self):
        """Create and arrange GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Markdown to DOCX Converter", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file selection
        ttk.Label(main_frame, text="Select Markdown File:", style='Heading.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=(0, 5)
        )
        
        self.input_file_var = tk.StringVar()
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_file_var, font=('Arial', 10))
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(input_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=1)
        
        # Output file selection
        ttk.Label(main_frame, text="Output DOCX File:", style='Heading.TLabel').grid(
            row=3, column=0, sticky=tk.W, pady=(0, 5)
        )
        
        self.output_file_var = tk.StringVar()
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file_var, font=('Arial', 10))
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(output_frame, text="Browse", command=self.browse_output_file).grid(row=0, column=1)
        
        # Conversion options
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding="10")
        options_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        options_frame.columnconfigure(1, weight=1)
        
        self.toc_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Include Table of Contents", variable=self.toc_var).grid(
            row=0, column=0, sticky=tk.W
        )
        
        self.number_sections_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Number Sections", variable=self.number_sections_var).grid(
            row=1, column=0, sticky=tk.W
        )
        
        self.highlight_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Syntax Highlighting", variable=self.highlight_var).grid(
            row=2, column=0, sticky=tk.W
        )
        
        # Template file
        ttk.Label(options_frame, text="Reference Template (optional):").grid(
            row=3, column=0, sticky=tk.W, pady=(10, 5)
        )
        
        self.template_file_var = tk.StringVar()
        template_frame = ttk.Frame(options_frame)
        template_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))
        template_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(template_frame, textvariable=self.template_file_var, font=('Arial', 10)).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(template_frame, text="Browse", command=self.browse_template_file).grid(row=0, column=1)
        
        # Convert button
        self.convert_button = ttk.Button(
            main_frame, text="Convert to DOCX", command=self.convert_file,
            style='Accent.TButton'
        )
        self.convert_button.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=8, column=0, columnspan=3)
        
        # Log text area
        log_frame = ttk.LabelFrame(main_frame, text="Conversion Log", padding="5")
        log_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(9, weight=1)
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, font=('Consolas', 9))
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def check_pandoc(self):
        """Check if pandoc is installed and available"""
        try:
            result = subprocess.run(['pandoc', '--version'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version_line = result.stdout.split('\n')[0]
                self.log_message(f"✓ Pandoc found: {version_line}")
                self.status_var.set("Ready - Pandoc available")
            else:
                raise subprocess.CalledProcessError(result.returncode, 'pandoc')
        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
            self.log_message("✗ Pandoc not found. Please install pandoc first.")
            self.status_var.set("Error - Pandoc not installed")
            self.convert_button.configure(state='disabled')
            messagebox.showerror(
                "Pandoc Not Found",
                "Pandoc is required but not found on your system.\n\n"
                "Please install pandoc using:\n"
                "winget install pandoc\n\n"
                "Or download from: https://pandoc.org/installing.html"
            )
    
    def browse_input_file(self):
        """Browse for input markdown file"""
        filename = filedialog.askopenfilename(
            title="Select Markdown File",
            filetypes=[
                ("Markdown files", "*.md *.markdown *.mdown *.mkd"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.input_file_var.set(filename)
            # Auto-suggest output filename
            if not self.output_file_var.get():
                output_path = Path(filename).with_suffix('.docx')
                self.output_file_var.set(str(output_path))
    
    def browse_output_file(self):
        """Browse for output DOCX file"""
        filename = filedialog.asksaveasfilename(
            title="Save DOCX File As",
            defaultextension=".docx",
            filetypes=[
                ("Word documents", "*.docx"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.output_file_var.set(filename)
    
    def browse_template_file(self):
        """Browse for template DOCX file"""
        filename = filedialog.askopenfilename(
            title="Select Reference Template",
            filetypes=[
                ("Word documents", "*.docx"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.template_file_var.set(filename)
    
    def log_message(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def convert_file(self):
        """Convert markdown file to DOCX"""
        # Validate inputs
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input markdown file.")
            return
        
        if not output_file:
            messagebox.showerror("Error", "Please specify an output DOCX file.")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"Input file does not exist: {input_file}")
            return
        
        # Start conversion in separate thread
        self.convert_button.configure(state='disabled')
        self.progress.start()
        self.status_var.set("Converting...")
        
        conversion_thread = threading.Thread(target=self._perform_conversion, args=(input_file, output_file))
        conversion_thread.daemon = True
        conversion_thread.start()
    
    def _perform_conversion(self, input_file, output_file):
        """Perform the actual conversion (runs in separate thread)"""
        try:
            # Build pandoc command
            cmd = ['pandoc', input_file, '-o', output_file]
            
            # Add options
            if self.toc_var.get():
                cmd.append('--toc')
            
            if self.number_sections_var.get():
                cmd.append('--number-sections')
            
            if self.highlight_var.get():
                cmd.extend(['--highlight-style', 'tango'])
            
            template_file = self.template_file_var.get().strip()
            if template_file and os.path.exists(template_file):
                cmd.extend(['--reference-doc', template_file])
            
            # Log the command
            self.root.after(0, self.log_message, f"Running: {' '.join(cmd)}")
            
            # Execute pandoc
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                # Success
                self.root.after(0, self._conversion_success, output_file)
            else:
                # Error
                error_msg = result.stderr or "Unknown error occurred"
                self.root.after(0, self._conversion_error, error_msg)
                
        except subprocess.TimeoutExpired:
            self.root.after(0, self._conversion_error, "Conversion timed out")
        except Exception as e:
            self.root.after(0, self._conversion_error, str(e))
    
    def _conversion_success(self, output_file):
        """Handle successful conversion"""
        self.progress.stop()
        self.convert_button.configure(state='normal')
        self.status_var.set("Conversion completed successfully!")
        self.log_message(f"✓ Successfully created: {output_file}")
        
        # Ask if user wants to open the file
        if messagebox.askyesno("Success", f"Conversion completed!\n\nWould you like to open the output file?\n{output_file}"):
            try:
                os.startfile(output_file)  # Windows
            except AttributeError:
                subprocess.run(['open', output_file])  # macOS
            except Exception:
                subprocess.run(['xdg-open', output_file])  # Linux
    
    def _conversion_error(self, error_msg):
        """Handle conversion error"""
        self.progress.stop()
        self.convert_button.configure(state='normal')
        self.status_var.set("Conversion failed")
        self.log_message(f"✗ Error: {error_msg}")
        messagebox.showerror("Conversion Failed", f"An error occurred during conversion:\n\n{error_msg}")


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = MarkdownToDocxConverter(root)
    
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Start the GUI
    root.mainloop()


if __name__ == "__main__":
    main()
