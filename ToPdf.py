import os
import argparse
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/drive']

def authenticate_drive(service_account_file):
    """Authenticate using a service account and create a Google Drive service client."""
    try:
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=SCOPES)
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        logging.error(f"Failed to authenticate with Google Drive: {e}")
        raise

def upload_and_convert_to_pdf(file_path, output_dir, drive_service):
    """Upload a file to Google Drive, convert to Google format, export as PDF, and download."""
    extension = os.path.splitext(file_path)[1].lower()
    mime_types = {
        '.doc': 'application/msword',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.ppt': 'application/vnd.ms-powerpoint',
        '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        '.xls': 'application/vnd.ms-excel',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    google_mime_types = {
        '.doc': 'application/vnd.google-apps.document',
        '.docx': 'application/vnd.google-apps.document',
        '.ppt': 'application/vnd.google-apps.presentation',
        '.pptx': 'application/vnd.google-apps.presentation',
        '.xls': 'application/vnd.google-apps.spreadsheet',
        '.xlsx': 'application/vnd.google-apps.spreadsheet'
    }
    file_mime_type = mime_types.get(extension)
    google_mime_type = google_mime_types.get(extension)
    if not file_mime_type or not google_mime_type:
        logging.warning(f"Unsupported file type: {extension}")
        return
    file_metadata = {
        'name': os.path.basename(file_path),
        'mimeType': google_mime_type
    }
    media = MediaFileUpload(file_path, mimetype=file_mime_type)
    file_id = None
    try:
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        file_id = file.get('id')
        logging.info(f"Uploaded file with ID: {file_id}")
        pdf_name = os.path.splitext(os.path.basename(file_path))[0] + '.pdf'
        pdf_path = os.path.join(output_dir, pdf_name)
        request = drive_service.files().export_media(fileId=file_id, mimeType='application/pdf')
        with open(pdf_path, 'wb') as f:
            f.write(request.execute())
        logging.info(f"Downloaded PDF to: {pdf_path}")
    except HttpError as e:
        logging.error(f"Google API error: {e}")
    except Exception as e:
        logging.error(f"Error processing {file_path}: {e}")
    finally:
        if file_id:
            try:
                drive_service.files().delete(fileId=file_id).execute()
                logging.info(f"Deleted file with ID: {file_id} from Google Drive")
            except Exception as e:
                logging.warning(f"Failed to delete file with ID: {file_id}: {e}")

def process_files(input_dir, output_dir, drive_service):
    """Process all supported files in the input directory."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    supported_extensions = ['.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx']
    for filename in os.listdir(input_dir):
        if any(filename.lower().endswith(ext) for ext in supported_extensions):
            file_path = os.path.join(input_dir, filename)
            logging.info(f"Processing {file_path}...")
            upload_and_convert_to_pdf(file_path, output_dir, drive_service)

def main():
    parser = argparse.ArgumentParser(description="Convert Office files to PDF using Google Drive API.")
    parser.add_argument('--input', '-i', required=True, help='Input directory containing files to convert')
    parser.add_argument('--output', '-o', required=True, help='Output directory for PDFs')
    parser.add_argument('--service-account', '-s', default='service-account.json', help='Path to service account JSON file')
    parser.add_argument('--log', '-l', default='INFO', help='Logging level (DEBUG, INFO, WARNING, ERROR)')
    args = parser.parse_args()
    logging.basicConfig(level=getattr(logging, args.log.upper(), None), format='%(levelname)s: %(message)s')
    try:
        drive_service = authenticate_drive(args.service_account)
        process_files(args.input, args.output, drive_service)
    except Exception as e:
        logging.error(f"Fatal error: {e}")

if __name__ == '__main__':
    main()
