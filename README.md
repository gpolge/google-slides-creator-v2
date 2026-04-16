# Google Slides Creator v2

Programmatically build Maze-branded QBR and partnership review decks using the Google Slides, Sheets, and Drive APIs — no manual editing required.

## What it does

- **Copies the official Maze QBR template** preserving all design (fonts, colors, layouts, master slides)
- **Fills all content placeholders** with customer-specific data (goals, engagement stats, achievements, spotlight, industry benchmarks)
- **Creates and embeds live charts** (bar, line, donut, stacked bar, horizontal bar) via Google Sheets — linked natively into Slides
- **One deck per run** — no Drive clutter

## Project structure

```
├── src/
│   ├── auth.py           # OAuth2 flow (browser login, token cache)
│   ├── colors.py         # Maze brand color system
│   ├── slides_api.py     # Low-level Slides API request builders
│   ├── sheets_charts.py  # Chart creation via Google Sheets API
│   ├── builder.py        # High-level DeckBuilder class
│   └── drive.py          # Drive helpers
├── qbr_builder.py        # Generic QBR builder (edit CUSTOMER dict)
├── qbr_jlr.py            # Jaguar Land Rover partnership review example
├── demo.py               # Full demo deck (bar, line, donut, stacked bar charts)
├── init_deck.py          # One-time: create persistent "AI Slides" presentation
├── cleanup.py            # Remove test presentations from Drive
├── requirements.txt
└── credentials.example.json
```

## Setup

### 1. Google Cloud project

Enable these APIs in your Google Cloud project:
- [Google Slides API](https://console.developers.google.com/apis/api/slides.googleapis.com)
- [Google Sheets API](https://console.developers.google.com/apis/api/sheets.googleapis.com)
- [Google Drive API](https://console.developers.google.com/apis/api/drive.googleapis.com)

### 2. OAuth credentials

1. Create an OAuth 2.0 Desktop App credential in Google Cloud Console
2. Download the JSON and save as `credentials.json` (see `credentials.example.json` for format)

### 3. Install dependencies

```bash
pip3 install -r requirements.txt
```

### 4. Authenticate

Run any script — it will open a browser for Google login on first run and cache the token locally.

## Usage

### Generate a QBR deck for a customer

Edit the `CUSTOMER` dict in `qbr_builder.py`, then:

```bash
python3 qbr_builder.py
```

Outputs a Google Slides URL. The deck is a copy of the Maze QBR template with all placeholders filled and a live usage chart injected.

### Run the JLR partnership review

```bash
python3 qbr_jlr.py
```

### Run the full demo deck

```bash
python3 demo.py
```

## Chart types

All charts are created as native Google Sheets charts and embedded directly into Slides via `createSheetsChart`.

| Function | Chart type |
|---|---|
| `sc.create_bar_chart()` | Vertical bar (column) |
| `sc.create_hbar_chart()` | Horizontal bar |
| `sc.create_line_chart()` | Line (multi-series) |
| `sc.create_stacked_bar_chart()` | Stacked column |
| `sc.create_donut_chart()` | Donut / pie |

## Maze brand colors

| Token | Hex | Use |
|---|---|---|
| `MAZE_PURPLE` | `#A366FF` | Primary accent, charts |
| `MAZE_DARK` | `#1A0533` | Dark backgrounds |
| `MAZE_CYAN` | `#79DEC0` | Secondary accent |
| `MAZE_PINK` | `#FC1258` | Highlights |
| `MAZE_LIME` | `#B9E20E` | Tertiary |
| `MAZE_YELLOW` | `#FFC300` | Badges |

## Template

The QBR template ID is hardcoded in `qbr_builder.py` and `qbr_jlr.py`:
```
1KpYl6uTcUMBUFCHnMGkRXJVuxhcdsHrhDnzbZRQKNmo
```
