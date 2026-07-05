"""
drive_helper.py
Uploads captured/selected member photos to a shared Google Drive folder,
and lets the app browse/download photos already sitting in that folder.
Uses the same service account credentials as the Sheets integration.
"""

import io
import os

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from network_utils import with_retry

CREDENTIALS_FILE = "credentials.json"
DRIVE_FOLDER_ID = "1X6OtZxVXiUVk7LgGtzVPFvvoTi1dtTRk"

SCOPES = ["https://www.googleapis.com/auth/drive"]

_service = None


def _build_service():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
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
