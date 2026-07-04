"""
drive_helper.py
Uploads captured/selected member photos to a shared Google Drive folder,
and lets the app browse/download photos already sitting in that folder.

Uses a real user OAuth login (not the service account used for Sheets)
because service accounts have ZERO personal storage quota - uploading a
new file with one fails with a 403 "Service Accounts do not have storage
quota" error, even into a folder someone else shared with it, unless that
folder lives in a Google Workspace Shared Drive. For a regular/free Gmail
account there is no Shared Drive option, so instead we sign in once as the
actual Google account (a browser window opens for that), and the uploaded
files then count against that account's normal 15GB quota like any file
you'd upload by hand.

One-time setup required (see README section in this docstring):
1. In Google Cloud Console (same project as the service account, or any
   project), create an OAuth Client ID of type "Desktop app".
2. Download its JSON and save it next to this file as
   "oauth_client_secret.json".
3. Run the app and do any Drive action (scan/upload/browse a photo) - a
   browser window opens once for you to sign in and grant access. After
   that, a refresh token is cached in "drive_oauth_token.json" so you
   won't need to sign in again.
"""

import io
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials as UserCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from network_utils import with_retry

DRIVE_FOLDER_ID = "1X6OtZxVXiUVk7LgGtzVPFvvoTi1dtTRk"

OAUTH_CLIENT_SECRET_FILE = "oauth_client_secret.json"
OAUTH_TOKEN_FILE = "drive_oauth_token.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]

_service = None


def _get_user_credentials():
    """Loads cached OAuth credentials (refreshing if needed), or runs the
    one-time browser sign-in flow if none are cached yet."""
    creds = None
    if os.path.exists(OAUTH_TOKEN_FILE):
        try:
            creds = UserCredentials.from_authorized_user_file(OAUTH_TOKEN_FILE, SCOPES)
        except Exception:
            creds = None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open(OAUTH_TOKEN_FILE, "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            pass  # refresh failed (e.g. revoked access) - fall through to a fresh sign-in

    if not os.path.exists(OAUTH_CLIENT_SECRET_FILE):
        raise RuntimeError(
            f"Google Drive isn't connected yet: '{OAUTH_CLIENT_SECRET_FILE}' is missing. "
            "Create a Desktop-app OAuth Client ID in Google Cloud Console, download its "
            f"JSON, and save it next to the app as '{OAUTH_CLIENT_SECRET_FILE}'."
        )

    flow = InstalledAppFlow.from_client_secrets_file(OAUTH_CLIENT_SECRET_FILE, SCOPES)
    creds = flow.run_local_server(port=0)
    with open(OAUTH_TOKEN_FILE, "w") as f:
        f.write(creds.to_json())
    return creds


def _build_service():
    creds = _get_user_credentials()
    return build("drive", "v3", credentials=creds)


def _get_service():
    global _service
    if _service is None:
        _service = with_retry(_build_service)
    return _service


def _do_upload(local_path, folder_id):
    service = _get_service()
    filename = os.path.basename(local_path)
    file_metadata = {"name": filename, "parents": [folder_id]}
    media = MediaFileUpload(local_path, resumable=False)
    file = service.files().create(
        body=file_metadata, media_body=media, fields="id, webViewLink"
    ).execute()
    return file.get("id"), file.get("webViewLink")


def upload_photo(local_path, folder_id=DRIVE_FOLDER_ID):
    """Uploads a local photo file to the Drive folder.
    Returns (file_id, web_view_link)."""
    return with_retry(_do_upload, local_path, folder_id)


def _do_list_photos(folder_id):
    service = _get_service()
    query = (
        f"'{folder_id}' in parents and trashed = false "
        "and mimeType contains 'image/'"
    )
    files = []
    page_token = None
    while True:
        results = service.files().list(
            q=query, fields="nextPageToken, files(id, name)", pageSize=1000,
            orderBy="name", pageToken=page_token,
        ).execute()
        files.extend(results.get("files", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return files


def list_photos(folder_id=DRIVE_FOLDER_ID):
    """Returns a list of {'id', 'name'} dicts for image files in the folder."""
    return with_retry(_do_list_photos, folder_id)


def _do_download(file_id, dest_path):
    service = _get_service()
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest_path


def download_photo(file_id, dest_path):
    """Downloads a Drive file to dest_path. Returns dest_path."""
    return with_retry(_do_download, file_id, dest_path)
