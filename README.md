# MD to DOCX Converter

A simple desktop application for converting Markdown files to DOCX format using pandoc.

## Features

- **Easy-to-use GUI** built with tkinter
- **File browser integration** for selecting input and output files
- **Conversion options**:
  - Table of Contents generation
  - Section numbering
  - Syntax highlighting for code blocks
  - Custom Word template support
- **Real-time conversion log** with progress indication
- **Error handling** with helpful messages
- **Cross-platform** support (Windows, macOS, Linux)

## Prerequisites

### Required Software

1. **Python 3.7+** - Download from [python.org](https://www.python.org/downloads/)
2. **Pandoc** - The conversion engine

### Installing Pandoc

#### Windows
```powershell
winget install pandoc
```

#### macOS
```bash
brew install pandoc
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get install pandoc
```

## Installation

1. **Clone or download** this project to your computer
2. **Navigate** to the project directory
3. **Install dependencies** (optional, tkinter is included with Python):
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

#### Method 1: Double-click
- Simply double-click `main.py` if Python is properly associated with `.py` files

#### Method 2: Command Line
```bash
python main.py
```

#### Method 3: VS Code
- Open the project in VS Code
- Press `F5` or use `Run > Start Debugging`

### Using the Converter

1. **Select Input File**: Click "Browse" next to "Select Markdown File" and choose your `.md` file
2. **Choose Output Location**: Click "Browse" next to "Output DOCX File" to specify where to save the converted file
3. **Configure Options** (optional):
   - ✅ Include Table of Contents
   - ✅ Number Sections  
   - ✅ Syntax Highlighting
   - 📄 Reference Template (use a custom Word template for styling)
4. **Click "Convert to DOCX"** to start the conversion
5. **Monitor Progress** in the log area
6. **Open Result** when prompted

### Conversion Options Explained

- **Table of Contents**: Automatically generates a TOC based on your markdown headings
- **Number Sections**: Adds automatic numbering to headings (1.1, 1.2, etc.)
- **Syntax Highlighting**: Applies color coding to code blocks in the output
- **Reference Template**: Use an existing Word document as a style template

## Supported File Types

### Input Formats
- `.md` - Standard Markdown
- `.markdown` - Markdown
- `.mdown` - Markdown variant
- `.mkd` - Markdown variant

### Output Format
- `.docx` - Microsoft Word Document

## Troubleshooting

### Common Issues

#### "Pandoc not found"
- **Solution**: Install pandoc using the instructions above
- **Check**: Run `pandoc --version` in your terminal to verify installation

#### "Permission denied" errors
- **Solution**: Make sure you have write permissions to the output directory
- **Try**: Choose a different output location (like your Documents folder)

#### "File not found" errors
- **Solution**: Verify the input file exists and the path is correct
- **Check**: The file path doesn't contain special characters

#### Application won't start
- **Solution**: Make sure Python 3.7+ is installed
- **Check**: Run `python --version` to verify Python installation

### Getting Help

If you encounter issues:

1. **Check the conversion log** in the application for detailed error messages
2. **Verify pandoc installation** by running `pandoc --version` in terminal
3. **Test with a simple markdown file** first
4. **Check file permissions** for both input and output locations

## Advanced Usage

### Custom Templates

To use a custom Word template:

1. Create a Word document with your desired styles, fonts, and formatting
2. Save it as a `.docx` file
3. In the application, browse and select this file as your "Reference Template"
4. The converted document will use these styles

### Batch Conversion

For converting multiple files, you can:

1. Use the application multiple times, or
2. Use pandoc directly from command line:
   ```bash
   # Convert all .md files in current directory
   for file in *.md; do pandoc "$file" -o "${file%.md}.docx"; done
   ```

## Technical Details

### Dependencies
- **Python 3.7+**: Core runtime
- **tkinter**: GUI framework (included with Python)
- **subprocess**: For running pandoc commands
- **pathlib**: Modern path handling
- **threading**: Non-blocking conversions

### Architecture
- **Single-file application** for easy distribution
- **Threaded conversion** to keep GUI responsive
- **Error handling** with user-friendly messages
- **Cross-platform compatibility**

## License

This project is open source and available under the MIT License.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this application!
