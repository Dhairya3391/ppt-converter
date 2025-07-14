# ppt-converter

Convert Office files (PPT, DOC, XLS, etc.) to PDF using the Google Drive API and a service account.

## Setup

1. Clone the repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Obtain a Google Cloud service account JSON key and save it (default: `service-account.json`).
4. Enable the Google Drive API for your project in the Google Cloud Console.

## Usage

Prepare your input files in a directory (e.g., `input_files`).

Run the converter:

```bash
python ToPdf.py --input input_files --output output_pdfs --service-account service-account.json
```

- `--input` or `-i`: Directory with files to convert
- `--output` or `-o`: Directory to save PDFs
- `--service-account` or `-s`: Path to service account JSON (default: `service-account.json`)
- `--log` or `-l`: Logging level (default: INFO)

## Supported Formats

- .doc, .docx
- .ppt, .pptx
- .xls, .xlsx

## Notes

- Files are uploaded to Google Drive, converted, exported as PDF, and deleted from Drive automatically.
- Make sure your service account has access to the Drive API.
