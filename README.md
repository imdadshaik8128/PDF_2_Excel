# PDF to Excel Converter

A standalone desktop application that converts PDF files with structured hierarchical format into well-organized Excel files with merged cells. The application features a Python backend for powerful text extraction and Excel generation, combined with a user-friendly web-based interface built with HTML and JavaScript.

## Features

* Convert PDF files to Excel format
* Preserve hierarchical structure with merged cells
* User-friendly web-based interface
* Standalone desktop application (no installation required)
* Automatic fallback between PDF parsing libraries for maximum compatibility

## Prerequisites

Before building the executable, ensure you have Python installed on your development machine.

## Installation & Setup

### 1. Python Packages (Libraries)

Install the core dependencies required for the Python backend:

```bash
pip install -r requirements.txt
```

#### Required Libraries:

* **`eel`**: Main library that enables communication between Python script and web UI
* **`PyPDF2`** or **`PyMuPDF` (fitz)**: For PDF text extraction
  * `PyPDF2`: Stable and reliable for general text extraction
  * `PyMuPDF`: Faster and better at handling complex layouts (primary choice with PyPDF2 fallback)
* **`openpyxl`**: Library for creating and styling Excel `.xlsx` files

#### requirements.txt
```txt
eel
PyPDF2
PyMuPDF
openpyxl
```

### 2. Project File Structure

Ensure your project follows this exact structure for PyInstaller compatibility:

```
your_project_folder/
├── main_app.py
├── requirements.txt
├── README.md
└── web/
    └── index.html
```

**File Descriptions:**
* `main_app.py`: Python script containing Eel setup and core parsing logic
* `web/`: Dedicated folder for all web assets (HTML, CSS, JavaScript)
* `index.html`: Main HTML file for the user interface
* `requirements.txt`: List of Python dependencies
* `README.md`: This documentation file

### 3. Development Environment Setup

1. **Clone or download the project files**
2. **Navigate to the project directory:**
   ```bash
   cd your_project_folder
   ```
3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
4. **Install PyInstaller** (required for building executable):
   ```bash
   pip install pyinstaller
   ```

## Building the Executable

### Build Command

From the project root directory (`your_project_folder`), run:

```bash
pyinstaller --onefile --add-data "web;web" main_app.py
```

#### Command Options Explained:
* `--onefile`: Bundles everything into a single `.exe` file
* `--add-data "web;web"`: Copies the `web` folder contents into the final application
  * Format: `"source_path;destination_path"`
* `main_app.py`: The main Python script to build from

### Build Output

After successful build, you'll find:
* `dist/main_app.exe`: Your standalone executable
* `build/`: Temporary build files (can be deleted)
* `main_app.spec`: PyInstaller specification file

## Usage

### For End Users
1. Double-click `main_app.exe` to launch the application
2. The web interface will open in your default browser
3. Select your PDF file using the file picker
4. Click "Convert" to generate the Excel file
5. The converted Excel file will be saved to your chosen location

### For Developers
1. Run the application in development mode:
   ```bash
   python main_app.py
   ```
2. Make changes to the code as needed
3. Rebuild the executable using the PyInstaller command

## Technical Details

### PDF Processing
* Primary: PyMuPDF (fitz) for fast and accurate text extraction
* Fallback: PyPDF2 for compatibility with various PDF formats
* Handles structured hierarchical data extraction

### Excel Generation
* Uses openpyxl for creating `.xlsx` files
* Implements cell merging for hierarchical data representation
* Maintains formatting and structure from original PDF

### User Interface
* Web-based UI using HTML, CSS, and JavaScript
* Communicates with Python backend via Eel
* Responsive design for better user experience

## Troubleshooting

### Common Issues

**Build fails with "module not found" error:**
* Ensure all dependencies are installed: `pip install -r requirements.txt`
* Check that you're in the correct directory when running PyInstaller

**Executable doesn't start:**
* Verify the `web` folder is included in the build
* Check that `index.html` exists in the `web` folder

**PDF conversion fails:**
* Ensure the PDF has extractable text (not scanned images)
* Try with a different PDF file to isolate the issue

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Review the project structure requirements
3. Ensure all dependencies are properly installed

---

**Note**: This application creates a self-contained executable that can be distributed to end-users without requiring them to install Python or any dependencies.
