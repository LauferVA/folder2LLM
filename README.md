# folder2LLM
prepares contents of ./dir and its subdirs by converting to sanitized plain text for consumption by a (non RAG) LLM.

# File-to-Text Conversion Script

This script extracts text from a wide variety of file types—DOCX, PDF, Excel, ODT, Jupyter notebooks, and more—and saves the resulting text to `.txt` files. It's a quick way to convert mixed document collections into plain-text format suitable for downstream NLP or large language model processing.

## Features

1. **Direct Text Extraction** From Several File Types  
   - Supports extraction of text from a variety of plain text files: ".txt", ".md", ".py", ".json", ".csv", ".tsv", ".log", ".xml", ".yaml", ".yml", ".html", ".htm", ".css", ".js", ".jsx", ".ts", ".tsx", ".sh", ".cmd", ".ps1", ".swift", ".kt", ".go", ".rs", ".lua", ".pl", ".r", ".m", ".vb", ".cs", ".asm", ".dart", ".php", ".rb", ".sql" 
2. **File Fmt Conversion Followed By Extraction** From Several Additional File Types: '.doc', '.docx', '.ipynb', '.odp', '.ods', '.odt', '.pdf',\ '.ppt', '.pptx', '.rtf', '.xls', '.xlsm', '.xlsx'
2. **Sanitization**  
   - Removes odd or non-printable characters, normalizes Unicode to NFC, and ensures UTF-8 output.
3. **Flexible** Allows skipping files above a certain size (`max_file_size`) or hidden files (optional).
4. **Local & Private**  No network calls. All extraction happens entirely on your machine.

## Requirements

- **Python 3.7+** (for best compatibility with libraries)
- **Optional Libraries** (install only what you need):
  - `python-docx` (for `.docx`)
  - `PyMuPDF` (or `PyPDF2`) (for `.pdf`)
  - `openpyxl` (for `.xlsx`)
  - `odfpy` (for `.odt`)
  - plus any other libraries you desire for advanced parsing

You can install them via pip, for example:
```bash
pip install python-docx PyMuPDF PyPDF2 openpyxl odfpy
```
*(You only need PyMuPDF **or** PyPDF2 for PDF support.)*

## Usage

1. **Clone or Download** this repository/script.  
2. **Install** any required libraries as desired (see [Requirements](#requirements)).
3. **Run** the script with:
   ```bash
   python files_2_zip.py /path/to/source /path/to/destination
   ```

This will recursively scan `/path/to/source` for supported file types, extract text (if possible), sanitize it, and write `.txt` files into `/path/to/destination`, preserving subdirectory structure.

### Script Arguments

- **`input_dir`**: Directory containing the original files to be converted.
- **`output_dir`**: Directory where `.txt` files will be written.
- **`max_file_size`** (adjustable in the code): Skip files above this size (in bytes). Defaults to 5 MB.  
- **`skip_hidden`** (adjustable in the code): If `True`, skip hidden files/directories.

## Example

Suppose you have:
```
my_docs/
   report.pdf
   notes.docx
   data.xlsx
```
Run:
```bash
python files_2_zip.py my_docs converted_txt
```
Result:
```
converted_txt/
   report.txt
   notes.txt
   data.txt
```
Each `.txt` file contains the extracted, sanitized, and UTF-8 encoded text from the corresponding source file.

## Customizing the Sanitization

Inside the script, the `sanitize_text` function removes non-printable control characters and normalizes Unicode. If you need stricter rules (e.g., removing everything non-ASCII) or more lenient rules (e.g., keeping emojis), you can edit that function accordingly.

## License

This script is distributed under GPL3. Feel free to modify and incorporate it into your own projects as needed. contact: laufer@openchromatin.com

