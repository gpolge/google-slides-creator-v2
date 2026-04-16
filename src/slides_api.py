"""
Low-level Google Slides API request builders.
All functions return batchUpdate request dicts.
"""

import uuid
from . import colors as C


def new_id() -> str:
    return uuid.uuid4().hex[:16]


# ─────────────────────────────────────────────
# PRESENTATION
# ─────────────────────────────────────────────

def create_presentation(slides_service, title: str) -> dict:
    return slides_service.presentations().create(
        body={"title": title}
    ).execute()


def batch_update(slides_service, presentation_id: str, requests: list) -> dict:
    if not requests:
        return {}
    return slides_service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": requests},
    ).execute()


def get_presentation(slides_service, presentation_id: str) -> dict:
    return slides_service.presentations().get(
        presentationId=presentation_id
    ).execute()


# ─────────────────────────────────────────────
# SLIDE MANAGEMENT
# ─────────────────────────────────────────────

def add_slide(slide_id: str = None, layout: str = "BLANK", insertion_index: int = None) -> dict:
    slide_id = slide_id or new_id()
    req = {
        "createSlide": {
            "objectId": slide_id,
            "slideLayoutReference": {"predefinedLayout": layout},
        }
    }
    if insertion_index is not None:
        req["createSlide"]["insertionIndex"] = insertion_index
    return req


def delete_slide(slide_id: str) -> dict:
    return {"deleteObject": {"objectId": slide_id}}


def set_slide_background(slide_id: str, hex_color: str) -> dict:
    return {
        "updatePageProperties": {
            "objectId": slide_id,
            "pageProperties": {
                "pageBackgroundFill": C.solid_fill(hex_color)
            },
            "fields": "pageBackgroundFill",
        }
    }


# ─────────────────────────────────────────────
# SHAPES & TEXT BOXES
# ─────────────────────────────────────────────

def emu(inches: float) -> int:
    """Convert inches to EMU (English Metric Units). 1 inch = 914400 EMU."""
    return int(inches * 914400)


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU. 1 pt = 12700 EMU."""
    return int(pt * 12700)


def add_shape(
    obj_id: str,
    slide_id: str,
    shape_type: str,   # e.g. "RECTANGLE", "ELLIPSE", "ROUND_RECTANGLE"
    x: float, y: float, w: float, h: float,  # inches
    fill_color: str = None,
    border_color: str = None,
    border_width_pt: float = 0,
) -> list:
    """Add a shape to a slide. Returns list of request dicts."""
    reqs = [{
        "createShape": {
            "objectId": obj_id,
            "shapeType": shape_type,
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {
                    "width":  {"magnitude": emu(w), "unit": "EMU"},
                    "height": {"magnitude": emu(h), "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": emu(x), "translateY": emu(y),
                    "unit": "EMU",
                },
            },
        }
    }]

    props = {}
    if fill_color:
        props["shapeBackgroundFill"] = C.solid_fill(fill_color)
    if border_color and border_width_pt > 0:
        props["outline"] = {
            "outlineFill": {"solidFill": {"color": {"rgbColor": C.hex_to_rgb(border_color)}}},
            "weight": {"magnitude": pt_to_emu(border_width_pt), "unit": "EMU"},
        }
    else:
        props["outline"] = {"propertyState": "NOT_RENDERED"}

    if props:
        reqs.append({
            "updateShapeProperties": {
                "objectId": obj_id,
                "shapeProperties": props,
                "fields": ",".join(props.keys()),
            }
        })

    return reqs


_ALIGN_MAP = {"LEFT": "START", "RIGHT": "END", "CENTER": "CENTER", "JUSTIFIED": "JUSTIFIED",
              "START": "START", "END": "END"}

def add_text_box(
    obj_id: str,
    slide_id: str,
    text: str,
    x: float, y: float, w: float, h: float,  # inches
    font_size_pt: float = 14,
    bold: bool = False,
    italic: bool = False,
    color: str = None,
    bg_color: str = None,
    align: str = "START",   # START | CENTER | END | JUSTIFIED
    v_align: str = "MIDDLE",
) -> list:
    color = color or C.TEXT_PRIMARY
    reqs = [{
        "createShape": {
            "objectId": obj_id,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {
                    "width":  {"magnitude": emu(w), "unit": "EMU"},
                    "height": {"magnitude": emu(h), "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": emu(x), "translateY": emu(y),
                    "unit": "EMU",
                },
            },
        }
    }, {
        "insertText": {
            "objectId": obj_id,
            "text": text,
            "insertionIndex": 0,
        }
    }, {
        "updateTextStyle": {
            "objectId": obj_id,
            "textRange": {"type": "ALL"},
            "style": {
                "fontSize": {"magnitude": font_size_pt, "unit": "PT"},
                "bold": bold,
                "italic": italic,
                "foregroundColor": C.rgb_dict(color),
                "fontFamily": "Inter",
            },
            "fields": "fontSize,bold,italic,foregroundColor,fontFamily",
        }
    }, {
        "updateParagraphStyle": {
            "objectId": obj_id,
            "textRange": {"type": "ALL"},
            "style": {
                "alignment": _ALIGN_MAP.get(align, align),
                "spaceAbove": {"magnitude": 0, "unit": "PT"},
                "spaceBelow": {"magnitude": 0, "unit": "PT"},
            },
            "fields": "alignment,spaceAbove,spaceBelow",
        }
    }]

    shape_props = {"contentAlignment": v_align, "outline": {"propertyState": "NOT_RENDERED"}}
    fields = "contentAlignment,outline"
    if bg_color:
        shape_props["shapeBackgroundFill"] = C.solid_fill(bg_color)
        fields += ",shapeBackgroundFill"
    else:
        shape_props["shapeBackgroundFill"] = {"propertyState": "NOT_RENDERED"}
        fields += ",shapeBackgroundFill"

    reqs.append({
        "updateShapeProperties": {
            "objectId": obj_id,
            "shapeProperties": shape_props,
            "fields": fields,
        }
    })

    return reqs


# ─────────────────────────────────────────────
# IMAGE
# ─────────────────────────────────────────────

def add_image(
    obj_id: str,
    slide_id: str,
    url: str,
    x: float, y: float, w: float, h: float,  # inches
) -> dict:
    return {
        "createImage": {
            "objectId": obj_id,
            "url": url,
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {
                    "width":  {"magnitude": emu(w), "unit": "EMU"},
                    "height": {"magnitude": emu(h), "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": emu(x), "translateY": emu(y),
                    "unit": "EMU",
                },
            },
        }
    }


def add_sheets_chart(
    obj_id: str,
    slide_id: str,
    spreadsheet_id: str,
    chart_id: int,
    x: float, y: float, w: float, h: float,  # inches
    linking_mode: str = "LINKED",
) -> dict:
    """Embed a Google Sheets chart directly into a slide (no external URL needed)."""
    return {
        "createSheetsChart": {
            "objectId": obj_id,
            "spreadsheetId": spreadsheet_id,
            "chartId": chart_id,
            "linkingMode": linking_mode,
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {
                    "width":  {"magnitude": emu(w), "unit": "EMU"},
                    "height": {"magnitude": emu(h), "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": emu(x), "translateY": emu(y),
                    "unit": "EMU",
                },
            },
        }
    }


# ─────────────────────────────────────────────
# REPLACE TEXT (template placeholders)
# ─────────────────────────────────────────────

def replace_all_text(find: str, replace: str, match_case: bool = False) -> dict:
    """Replace all occurrences of {{find}} with replace across the entire presentation."""
    return {
        "replaceAllText": {
            "containsText": {"text": f"{{{{{find}}}}}", "matchCase": match_case},
            "replaceText": replace,
        }
    }
