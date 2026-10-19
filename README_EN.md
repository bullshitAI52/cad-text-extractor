# CAD Text Extractor

Python tools for reading AutoCAD DWG text through the AutoCAD COM interface and exporting or translating TEXT/MTEXT content to Excel-compatible files.

## Typical workflow

1. Install the Python dependencies from `requirements.txt`.
2. Make sure AutoCAD and the required COM environment are available on Windows.
3. Configure the input DWG files and output paths.
4. Run the appropriate Python script.
5. Review the exported spreadsheet and translated text before writing changes back to drawings.

The tool is intended for batch processing; keep backups of source drawings and do not expose API keys stored in environment files.
