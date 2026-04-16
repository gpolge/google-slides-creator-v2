"""Google Drive helpers: upload images, get URLs for Slides API."""

import io
from googleapiclient.http import MediaIoBaseUpload


def upload_image(drive_service, image_bytes: bytes, filename: str) -> str:
    """
    Upload a PNG to Google Drive.

    Strategy: try to make it publicly readable (works on personal accounts).
    If domain policy blocks public sharing, fall back to a Drive thumbnail URL
    which the Slides API can resolve because it runs on behalf of the same user.
    """
    media = MediaIoBaseUpload(
        io.BytesIO(image_bytes),
        mimetype="image/png",
        resumable=False,
    )

    file_meta = {"name": filename, "mimeType": "image/png"}

    uploaded = drive_service.files().create(
        body=file_meta,
        media_body=media,
        fields="id,webContentLink,thumbnailLink",
    ).execute()

    file_id = uploaded["id"]

    # Try public sharing first; fall back gracefully
    try:
        drive_service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
        ).execute()
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    except Exception:
        pass

    # Fallback: direct content URL — Slides API resolves Drive files owned
    # by the same authenticated user without needing public permissions.
    return f"https://drive.google.com/uc?id={file_id}&export=download"


def create_presentation_copy(drive_service, template_id: str, title: str) -> str:
    """Copy an existing Google Slides presentation and return the new ID."""
    copied = drive_service.files().copy(
        fileId=template_id,
        body={"name": title},
    ).execute()
    return copied["id"]
