"""Google OAuth2 authentication for Slides + Drive APIs."""

import os
import pickle
from pathlib import Path

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

_BASE = Path(__file__).parent.parent
CREDENTIALS_FILE = _BASE / "credentials.json"
TOKEN_FILE = _BASE / "token.pickle"


def get_credentials():
    """Load or refresh OAuth2 credentials. Prompts browser flow if needed."""
    creds = None

    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=8080)

        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    return creds


def get_services():
    """Return authenticated Slides, Drive, and Sheets service clients."""
    creds = get_credentials()
    slides   = build("slides",   "v1",  credentials=creds)
    drive    = build("drive",    "v3",  credentials=creds)
    sheets   = build("sheets",   "v4",  credentials=creds)
    return slides, drive, sheets
