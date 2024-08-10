# multiple_txt_to_one_docx

# Text Files to Word Document Converter

This script combines multiple `.txt` files into a single Microsoft Word document. It formats each file with a title that includes an appendix letter, the site name extracted from the hostname, and the hostname itself. The script automatically handles the title formatting and ensures that each file's content is properly separated.

## Features
#### see requirements in ** docs/junction scipt instructions.docx **
- Combines multiple `.txt` files into a single `.docx` file.
- Titles each section with an appendix letter, site name, and hostname.
- Formats titles and content with specified font styles and sizes.
- Searches for a specific string in each file and prints an error if not found.

## Requirements

- Python 3.x
- `docx` library (for creating Word documents)
- `tkinter` library (for file selection dialog)

## Installation

1. Ensure you have Python 3.x installed.
2. Install the required library with pip:

    ```sh
    pip install python-docx
    ```

## How to Use

1. **Save the Script**:
   - Save the provided script as `convert_txt_to_docx.py`.

2. **Run the Script**:
   - Execute the script using Python. You can run it from the command line:

    ```sh
    python convert_txt_to_docx.py
    ```

3. **Select Files**:
   - A file dialog will appear prompting you to select the `.txt` files you want to include in the Word document. You can select multiple files at once.

4. **Processing**:
   - The script will process the selected files and create a Word document named `combined_document.docx` in the same directory as the script.

## Configuration

- **Title Formatting**:
  - The title of each section in the Word document will be formatted as `Appendix X Configuration for SiteName: Hostname`, where `X` is an incrementing letter starting from 'D', and `SiteName` and `Hostname` are extracted from the content of each file.
  - The title will be bold, with the font set to Calibri and size 16.

- **Content Formatting**:
  - The body text will be formatted with font size 8 and single spacing.

- **Error Handling**:
  - The script checks for the presence of the string `no ip http server` in each file. If the string is not found, an error message will be printed in the console.

## Example

If you select the following files:

- `controller tester.txt` (hostname: `Transistor_contr`)
- `controller tester_1.1.txt` (hostname: `Switch_001`)
- `controller tester_1.2.txt` (hostname: `Router_XYZ`)

The generated Word document will contain:

- **Appendix D Configuration for Transistor: Transistor_contr**
  - Content of `controller_tester.txt`
  
- **Appendix E Configuration for Transistor: Switch_001**
  - Content of `controller_tester_1.1.txt`

- **Appendix F Configuration for Transistor: Router_XYZ**
  - Content of `controller_tester_1.2.txt`

## Notes

- Ensure that the `.txt` files follow the expected format for hostname extraction to work correctly.
- The script uses a simple regex to extract the hostname. If the format varies, the regex may need adjustments.

## Troubleshooting

- **Error: No files selected**: Make sure you have selected at least one file when the file dialog appears.
- **Formatting Issues**: Ensure the `.txt` files are formatted correctly and contain the necessary strings for the script to process.



