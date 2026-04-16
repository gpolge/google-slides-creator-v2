"""
One-time script: create the persistent "AI Slides" presentation and save its ID.
Run once: python3 init_deck.py
"""

import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services
from src import slides_api as api

CONFIG_FILE = Path(__file__).parent / "deck_config.json"


def main():
    if CONFIG_FILE.exists():
        cfg = json.loads(CONFIG_FILE.read_text())
        print(f"Already initialised. AI Slides ID: {cfg['presentation_id']}")
        print(f"https://docs.google.com/presentation/d/{cfg['presentation_id']}/edit")
        return

    print("Creating persistent 'AI Slides' presentation...")
    slides_svc, _, _ = get_services()
    pres = api.create_presentation(slides_svc, "AI Slides")
    pres_id = pres["presentationId"]

    CONFIG_FILE.write_text(json.dumps({"presentation_id": pres_id}, indent=2))
    print(f"Created! ID saved to deck_config.json")
    print(f"https://docs.google.com/presentation/d/{pres_id}/edit")


if __name__ == "__main__":
    main()
