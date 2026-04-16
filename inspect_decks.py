"""Inspect the reference deck and template to understand structure and design."""

import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services

REFERENCE_ID = "1oKxEiHgXZPvje9eTdVpGF2X0jjA9gUo-bAOcY4-7foU"
TEMPLATE_ID  = "1KpYl6uTcUMBUFCHnMGkRXJVuxhcdsHrhDnzbZRQKNmo"


def rgb_to_hex(rgb):
    if not rgb:
        return None
    r = int(rgb.get("red", 0) * 255)
    g = int(rgb.get("green", 0) * 255)
    b = int(rgb.get("blue", 0) * 255)
    return f"#{r:02X}{g:02X}{b:02X}"


def describe_element(el):
    out = {}
    t = el.get("shape", {})
    if t:
        st = t.get("shapeType", "")
        out["type"] = st
        # Text
        text_els = t.get("text", {}).get("textElements", [])
        texts = []
        for te in text_els:
            pr = te.get("textRun", {})
            if pr.get("content", "").strip():
                style = pr.get("style", {})
                fg = style.get("foregroundColor", {}).get("opaqueColor", {}).get("rgbColor")
                texts.append({
                    "text": pr["content"].strip(),
                    "bold": style.get("bold"),
                    "size": style.get("fontSize", {}).get("magnitude"),
                    "color": rgb_to_hex(fg),
                })
        if texts:
            out["texts"] = texts
        # Fill
        fill = t.get("shapeProperties", {}).get("shapeBackgroundFill", {})
        solid = fill.get("solidFill", {}).get("color", {}).get("rgbColor")
        if solid:
            out["fill"] = rgb_to_hex(solid)
    # Image
    img = el.get("image", {})
    if img:
        out["type"] = "IMAGE"
        out["contentUrl"] = img.get("contentUrl", "")[:80]
    # Position/size
    tf = el.get("transform", {})
    size = el.get("size", {})
    if tf and size:
        out["x_in"] = round(tf.get("translateX", 0) / 914400, 2)
        out["y_in"] = round(tf.get("translateY", 0) / 914400, 2)
        out["w_in"] = round(size.get("width", {}).get("magnitude", 0) / 914400, 2)
        out["h_in"] = round(size.get("height", {}).get("magnitude", 0) / 914400, 2)
    return out


def inspect(slides_svc, pres_id, name):
    print(f"\n{'='*60}")
    print(f"  {name}")
    print(f"  ID: {pres_id}")
    print(f"{'='*60}")
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    print(f"Title: {pres.get('title')}")
    print(f"Slides: {len(pres.get('slides', []))}")

    # Page size
    ps = pres.get("pageSize", {})
    w = ps.get("width", {}).get("magnitude", 0) / 914400
    h = ps.get("height", {}).get("magnitude", 0) / 914400
    print(f"Canvas: {w:.2f} x {h:.2f} inches")

    for i, slide in enumerate(pres.get("slides", [])):
        bg = slide.get("slideProperties", {}).get("pageBackgroundFill", {})
        bg_solid = bg.get("solidFill", {}).get("color", {}).get("rgbColor")
        bg_hex = rgb_to_hex(bg_solid) if bg_solid else "inherited"
        print(f"\n  Slide {i+1} (id={slide['objectId']}) bg={bg_hex}")
        for el in slide.get("pageElements", []):
            d = describe_element(el)
            if d:
                texts_preview = ""
                if "texts" in d:
                    texts_preview = " | ".join(
                        f'"{t["text"][:40]}" sz={t["size"]} bold={t["bold"]} col={t["color"]}'
                        for t in d["texts"][:3]
                    )
                print(f"    [{d.get('type','?')}] pos=({d.get('x_in')},{d.get('y_in')}) "
                      f"size=({d.get('w_in')}x{d.get('h_in')}) fill={d.get('fill')} {texts_preview}")

    # Save raw JSON for deep inspection
    out_path = Path(__file__).parent / f"inspect_{name.replace(' ','_').lower()}.json"
    out_path.write_text(json.dumps(pres, indent=2))
    print(f"\n  Full JSON saved to: {out_path.name}")


def main():
    slides_svc, _, _ = get_services()
    inspect(slides_svc, REFERENCE_ID, "Reference Deck")
    inspect(slides_svc, TEMPLATE_ID, "Template")


if __name__ == "__main__":
    main()
