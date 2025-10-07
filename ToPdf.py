#!/usr/bin/env python3
"""
Zero-touch Office-to-PDF converter using Google Drive + OAuth user flow.

Usage:
  Just run: python ToPdf.py

Behavior:
  - Uses fixed 'input' and 'output' directories located next to this script.
  - On first run opens browser for Google OAuth consent and stores token.json.
  - Subsequent runs reuse / silently refresh token; no browser unless revoked.
  - Converts .doc .docx .ppt .pptx .xls .xlsx to PDF via Drive import/export.

IMPORTANT (Cannot be bypassed):
  Google requires a valid OAuth 2.0 Client ID & Secret. They must be embedded
  below (CLIENT_ID / CLIENT_SECRET). Without real values the flow cannot succeed.
  Replace ONLY once before distribution if you control the environment.

SECURITY:
  Do NOT commit real credentials or token.json to public repositories.

Directories (auto-created if missing):
  ./input   - place source Office files here
  ./output  - PDFs will appear here

Token file:
  ./token.json

"""

import os
import sys
import json
import time
import logging
from typing import Optional
import concurrent.futures

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow

# ---------------------------------------------------------------------------
# Embed (placeholder) OAuth client credentials.
# You MUST set real values for these two constants for the script to function.
# ---------------------------------------------------------------------------
CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")

# OAuth scopes (broad). For tighter scope you could switch drive -> drive.file
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TOKEN_PATH = os.path.join(BASE_DIR, "token.json")
# Max parallel conversions (bounded to avoid too many simultaneous HTTP requests)
MAX_WORKERS = min(8, (os.cpu_count() or 4))

SUPPORTED_MIME = {
    ".doc": ("application/msword", "application/vnd.google-apps.document"),
    ".docx": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/vnd.google-apps.document",
    ),
    ".ppt": (
        "application/vnd.ms-powerpoint",
        "application/vnd.google-apps.presentation",
    ),
    ".pptx": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "application/vnd.google-apps.presentation",
    ),
    ".xls": ("application/vnd.ms-excel", "application/vnd.google-apps.spreadsheet"),
    ".xlsx": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.google-apps.spreadsheet",
    ),
}


def ensure_directories():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def validate_embedded_credentials():
    if (
        not CLIENT_ID
        or CLIENT_ID.startswith("REPLACE_ME_")
        or not CLIENT_SECRET
        or CLIENT_SECRET.startswith("REPLACE_ME_")
    ):
        msg = (
            "Embedded OAuth CLIENT_ID / CLIENT_SECRET are placeholders.\n"
            "Set valid Google OAuth 2.0 Desktop Client credentials directly in ToPdf.py\n"
            "before running. (Cannot proceed without them.)"
        )
        raise SystemExit(msg)


def load_or_authorize() -> Credentials:
    """
    Load cached credentials if present; refresh or start browser flow as needed.
    """
    creds: Optional[Credentials] = None

    if os.path.exists(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
            logging.debug("Loaded stored token.json")
        except Exception as e:
            logging.warning("Could not parse token.json (%s). Re-authenticating.", e)

    if creds and creds.expired and creds.refresh_token:
        try:
            logging.info("Refreshing expired access token...")
            creds.refresh(Request())
            save_credentials(creds)
            return creds
        except Exception as e:
            logging.warning("Refresh failed (%s); launching browser auth.", e)

    if not creds or not creds.valid:
        logging.info(
            "Launching browser for Google OAuth consent (first-time or invalid token)."
        )
        flow = InstalledAppFlow.from_client_config(
            {
                "installed": {
                    "client_id": CLIENT_ID,
                    "client_secret": CLIENT_SECRET,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": ["http://localhost"],
                }
            },
            SCOPES,
        )
        creds = flow.run_local_server(
            host="localhost",
            port=0,
            open_browser=True,
            authorization_prompt_message="Opening browser for authorization...",
            success_message="Authorization complete. You may close this tab.",
            access_type="offline",
            prompt="consent",
        )
        logging.info("Authorization succeeded.")
        save_credentials(creds)

    return creds


def save_credentials(creds: Credentials):
    data = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }
    try:
        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        try:
            os.chmod(TOKEN_PATH, 0o600)
        except Exception:
            pass
        logging.debug("Stored credentials at %s", TOKEN_PATH)
    except Exception as e:
        logging.error("Failed saving credentials: %s", e)


def build_drive(creds: Credentials):
    # cache_discovery=False suppresses the oauth2client file_cache warning
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def convert_one(drive, path: str):
    ext = os.path.splitext(path)[1].lower()
    mapping = SUPPORTED_MIME.get(ext)
    if not mapping:
        logging.debug("Skipping unsupported file: %s", path)
        return
    src_mime, tgt_mime = mapping
    filename = os.path.basename(path)
    pdf_name = os.path.splitext(filename)[0] + ".pdf"
    out_path = os.path.join(OUTPUT_DIR, pdf_name)

    start = time.time()
    file_id = None
    try:
        media = MediaFileUpload(path, mimetype=src_mime, resumable=False)
        metadata = {"name": filename, "mimeType": tgt_mime}
        created = (
            drive.files().create(body=metadata, media_body=media, fields="id").execute()
        )
        file_id = created["id"]
        logging.info("Uploaded %s (id=%s)", filename, file_id)

        request = drive.files().export_media(fileId=file_id, mimeType="application/pdf")
        with open(out_path, "wb") as f:
            f.write(request.execute())

        elapsed = time.time() - start
        logging.info("Saved PDF: %s (%.2fs)", out_path, elapsed)
    except HttpError as he:
        logging.error("Google API error on %s: %s", filename, he)
    except Exception as e:
        logging.error("Unexpected error on %s: %s", filename, e, exc_info=True)
    finally:
        if file_id:
            try:
                drive.files().delete(fileId=file_id).execute()
                logging.debug("Deleted temp file id=%s", file_id)
            except Exception as de:
                logging.warning("Could not delete temp file id=%s: %s", file_id, de)


def process_all(drive):
    # Collect supported files
    entries = [os.path.join(INPUT_DIR, n) for n in sorted(os.listdir(INPUT_DIR))]
    files = [
        p
        for p in entries
        if os.path.isfile(p) and os.path.splitext(p)[1].lower() in SUPPORTED_MIME
    ]
    if not files:
        logging.warning("Input directory is empty or no supported files: %s", INPUT_DIR)
        return
    start_batch = time.time()
    logging.info("Starting batch: %d files with %d workers", len(files), MAX_WORKERS)
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(convert_one, drive, f): f for f in files}
        for fut in concurrent.futures.as_completed(futures):
            f = futures[fut]
            try:
                fut.result()
            except Exception as e:
                logging.error("Failed converting %s: %s", f, e)
    logging.info(
        "Batch complete: %d files in %.2fs", len(files), time.time() - start_batch
    )


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )
    try:
        ensure_directories()
        validate_embedded_credentials()
        creds = load_or_authorize()
        drive = build_drive(creds)
        logging.info("Using up to %d concurrent workers", MAX_WORKERS)
        process_all(drive)
        logging.info("Done.")
        logging.info("Place additional files in '%s' and run again.", INPUT_DIR)
    except SystemExit as se:
        logging.error(str(se))
        sys.exit(1)
    except KeyboardInterrupt:
        logging.warning("Interrupted.")
        sys.exit(130)
    except Exception as e:
        logging.error("Fatal error: %s", e, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
