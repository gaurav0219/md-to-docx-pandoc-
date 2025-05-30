# MD to DOCX Converter - Executable Distribution

## About
This is a standalone executable version of the MD to DOCX Converter application.
It converts Markdown (.md) files to Microsoft Word (.docx) format.

## System Requirements
- Windows 7/8/10/11 (64-bit)
- Pandoc must be installed on your system

## Installation

### Step 1: Install Pandoc
This application requires Pandoc to be installed on your system.

**Option 1: Using Windows Package Manager (Recommended)**
Open Command Prompt or PowerShell as Administrator and run:
```
winget install pandoc
```

**Option 2: Manual Download**
1. Visit https://pandoc.org/installing.html
2. Download the Windows installer
3. Run the installer and follow the instructions

### Step 2: Run the Application
1. Double-click `MD-to-DOCX-Converter.exe` to start the application
2. Or use the provided `Run_MD_to_DOCX_Converter.bat` for a guided start

## Usage
1. Click "Browse" next to "Select Markdown File" to choose your .md file
2. Choose where to save the output .docx file
3. Configure conversion options if needed:
   - Include Table of Contents
   - Number Sections
   - Syntax Highlighting
   - Reference Template (optional)
4. Click "Convert to DOCX"

## Features
- User-friendly graphical interface
- Real-time conversion progress
- Comprehensive logging
- Support for conversion options
- Template-based formatting
- Cross-platform compatibility

## Troubleshooting

### "Pandoc not found" error
- Make sure Pandoc is properly installed
- Restart your computer after installing Pandoc
- Check that Pandoc is in your system PATH

### Conversion fails
- Check that the input Markdown file is valid
- Ensure you have write permissions to the output directory
- Check the conversion log for detailed error messages

## Support
If you encounter issues, please check the conversion log in the application
for detailed error messages.

## Version Information
Created with Python and PyInstaller
Requires Pandoc for file conversion
