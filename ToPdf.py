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
import threading
from random import uniform

# Load environment variables from .env file
try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass  # dotenv not available - user must set environment variables manually

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
    "https://www.googleapis.com/auth/drive.file",
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TOKEN_PATH = os.path.join(BASE_DIR, "token.json")
# Max parallel conversions (bounded to avoid too many simultaneous HTTP requests)
MAX_WORKERS = min(4, (os.cpu_count() or 4))
# Thread-local storage for per-thread Drive service
_thread_local = None  # will initialize lazily after importing threading

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
            # prompt removed to avoid forced re-consent
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


def convert_one(creds, path: str):
    """
    Convert a single file with retries.
    Returns: 'success' | 'failed' | 'skipped'
    """
    global _thread_local
    if _thread_local is None:
        _thread_local = threading.local()

    ext = os.path.splitext(path)[1].lower()
    mapping = SUPPORTED_MIME.get(ext)
    if not mapping:
        logging.debug("Skipping unsupported file: %s", path)
        return "skipped"

    filename = os.path.basename(path)
    pdf_name = os.path.splitext(filename)[0] + ".pdf"
    out_path = os.path.join(OUTPUT_DIR, pdf_name)

    # Skip if existing PDF is newer or same mtime
    if os.path.exists(out_path):
        try:
            if os.path.getmtime(out_path) >= os.path.getmtime(path):
                logging.info("Skipping (up-to-date): %s", filename)
                return "skipped"
        except OSError:
            pass

    def thread_drive():
        if not hasattr(_thread_local, "drive"):
            _thread_local.drive = build_drive(creds)
        return _thread_local.drive

    size = os.path.getsize(path)
    resumable = size > 5 * 1024 * 1024  # Use resumable uploads for files >5MB

    attempts = 3
    backoff = 1.0
    for attempt in range(1, attempts + 1):
        start = time.time()
        file_id = None
        try:
            src_mime, tgt_mime = mapping
            media = MediaFileUpload(path, mimetype=src_mime, resumable=resumable)
            metadata = {"name": filename, "mimeType": tgt_mime}
            drive = thread_drive()
            created = (
                drive.files()
                .create(body=metadata, media_body=media, fields="id")
                .execute()
            )
            file_id = created["id"]
            logging.info(
                "Uploaded %s (id=%s)%s",
                filename,
                file_id,
                " [resumable]" if resumable else "",
            )

            request = drive.files().export_media(
                fileId=file_id, mimeType="application/pdf"
            )
            with open(out_path, "wb") as f:
                f.write(request.execute())

            elapsed = time.time() - start
            logging.info("Saved PDF: %s (%.2fs)", out_path, elapsed)

            # Best-effort cleanup
            try:
                drive.files().delete(fileId=file_id).execute()
            except Exception:
                pass
            return "success"

        except HttpError as he:
            status = getattr(he, "status_code", None) or getattr(
                getattr(he, "resp", None), "status", None
            )
            if status == 400 and attempt == attempts:
                logging.error(
                    "Permanent HTTP 400 on %s after %d attempts: %s",
                    filename,
                    attempt,
                    he,
                )
                return "failed"
            logging.warning(
                "HTTP error on %s attempt %d/%d: %s", filename, attempt, attempts, he
            )
        except (BrokenPipeError, OSError) as ioe:
            if attempt == attempts:
                logging.error(
                    "I/O error on %s after %d attempts: %s", filename, attempt, ioe
                )
                return "failed"
            logging.warning(
                "I/O error on %s attempt %d/%d: %s", filename, attempt, attempts, ioe
            )
        except Exception as e:
            if attempt == attempts:
                logging.error(
                    "Unexpected error on %s after %d attempts: %s",
                    filename,
                    attempt,
                    e,
                    exc_info=True,
                )
                return "failed"
            logging.warning(
                "Retryable error on %s attempt %d/%d: %s",
                filename,
                attempt,
                attempts,
                e,
            )
        finally:
            if file_id:
                # Attempt cleanup between retries
                try:
                    thread_drive().files().delete(fileId=file_id).execute()
                except Exception:
                    pass

        # Exponential backoff with jitter
        time.sleep(backoff + uniform(0, 0.2))
        backoff *= 2

    return "failed"


def process_all(creds):
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
    results = {"success": 0, "failed": 0, "skipped": 0}
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(convert_one, creds, f): f for f in files}
        for fut in concurrent.futures.as_completed(futures):
            f = futures[fut]
            try:
                status = fut.result()
                if status in results:
                    results[status] += 1
            except Exception as e:
                logging.error("Failed converting %s: %s", f, e)
                results["failed"] += 1
    elapsed = time.time() - start_batch
    logging.info(
        "Batch complete in %.2fs | success=%d skipped=%d failed=%d",
        elapsed,
        results["success"],
        results["skipped"],
        results["failed"],
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
        logging.info("Using up to %d concurrent workers", MAX_WORKERS)
        process_all(creds)
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
