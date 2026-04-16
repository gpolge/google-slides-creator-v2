"""Maze brand color system."""

# --- Primary palette ---
MAZE_PURPLE  = "#A366FF"
MAZE_DARK    = "#1A0533"
MAZE_PINK    = "#FC1258"
MAZE_CYAN    = "#79DEC0"
MAZE_YELLOW  = "#FFC300"
MAZE_LIME    = "#B9E20E"
MAZE_GRAY    = "#6B6B8A"
MAZE_BORDER  = "#E5E0F5"
MAZE_MUTED   = "#F4ECFE"
MAZE_WHITE   = "#FFFFFF"

# --- Chart series palette (ordered) ---
SERIES = [
    MAZE_PURPLE,
    MAZE_CYAN,
    MAZE_PINK,
    MAZE_YELLOW,
    MAZE_LIME,
    "#C095F9",
    "#FD6189",
    MAZE_DARK,
]

# --- Slide background ---
SLIDE_BG     = MAZE_WHITE
SLIDE_DARK   = MAZE_DARK

# --- Typography colors ---
TEXT_PRIMARY   = MAZE_DARK
TEXT_SECONDARY = MAZE_GRAY
TEXT_ON_DARK   = MAZE_WHITE
TEXT_ACCENT    = MAZE_PURPLE


def hex_to_rgb(hex_color: str) -> dict:
    """Convert hex color to Google Slides API RGB dict (0-1 floats)."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


def rgb_dict(hex_color: str) -> dict:
    """Wrap hex_to_rgb in the opaqueColor structure for Slides API."""
    return {"opaqueColor": {"rgbColor": hex_to_rgb(hex_color)}}


def solid_fill(hex_color: str) -> dict:
    """Return a solidFill object for shape background use."""
    return {"solidFill": {"color": {"rgbColor": hex_to_rgb(hex_color)}}}
