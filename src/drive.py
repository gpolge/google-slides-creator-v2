"""Google Drive helpers: upload images, get URLs for Slides API."""

import io
from googleapiclient.http import MediaIoBaseUpload

# All decks are always created in this shared folder
OUTPUT_FOLDER_ID = "190SHiRi4BDFzcPzIUljwKtsZCZQKipjG"


def copy_template(drive_service, template_id: str, title: str) -> str:
    """
    Copy a template presentation into the shared output folder.
    Returns the new presentation ID.
    """
    copied = drive_service.files().copy(
        fileId=template_id,
        body={
            "name": title,
            "parents": [OUTPUT_FOLDER_ID],
        },
    ).execute()
    return copied["id"]


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

    try:
        drive_service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
        ).execute()
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    except Exception:
        pass

    return f"https://drive.google.com/uc?id={file_id}&export=download"


def move_to_output_folder(drive_service, file_id: str):
    """Move an existing file into the shared output folder."""
    # Get current parents first
    f = drive_service.files().get(fileId=file_id, fields="parents").execute()
    previous_parents = ",".join(f.get("parents", []))
    drive_service.files().update(
        fileId=file_id,
        addParents=OUTPUT_FOLDER_ID,
        removeParents=previous_parents,
        fields="id,parents",
    ).execute()
