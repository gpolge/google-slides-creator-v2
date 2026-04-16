"""Delete all test presentations and chart sheets created during debugging."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services

KEEP_PRESENTATION_ID = "1TaIhFZuUpL3iSWp8ju9ObKZGKPdHs1gT_09YtGlLdOo"

def main():
    _, drive_svc, _ = get_services()

    deleted = 0

    # Delete all Slides presentations named "Maze Demo Deck — API Build" except the keeper
    res = drive_svc.files().list(
        q="name='Maze Demo Deck — API Build' and mimeType='application/vnd.google-apps.presentation' and trashed=false",
        fields="files(id, name)",
    ).execute()
    for f in res.get("files", []):
        if f["id"] != KEEP_PRESENTATION_ID:
            drive_svc.files().delete(fileId=f["id"]).execute()
            print(f"Deleted presentation: {f['id']}")
            deleted += 1

    # Delete all chart spreadsheets (named "chart_*")
    res = drive_svc.files().list(
        q="name contains 'chart_' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
        fields="files(id, name)",
        pageSize=100,
    ).execute()
    for f in res.get("files", []):
        drive_svc.files().delete(fileId=f["id"]).execute()
        print(f"Deleted sheet: {f['name']} ({f['id']})")
        deleted += 1

    print(f"\nDone. {deleted} files deleted.")
    print(f"Kept: https://docs.google.com/presentation/d/{KEEP_PRESENTATION_ID}/edit")

if __name__ == "__main__":
    main()
