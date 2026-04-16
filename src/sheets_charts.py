"""
Create charts in Google Sheets with Maze brand colors.
Returns (spreadsheet_id, chart_id) tuples for use with Slides createSheetsChart.
"""

from . import colors as C

# ── Maze color sequences for chart series ─────────────────────────────────────
_SERIES_COLORS = [
    C.MAZE_PURPLE, C.MAZE_CYAN, C.MAZE_PINK, C.MAZE_YELLOW,
    C.MAZE_LIME, "#C095F9", "#FD6189",
]


def _rgb_obj(hex_color: str) -> dict:
    h = hex_color.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


def _cell_range(sheet_id: int, r1: int, c1: int, r2: int, c2: int) -> dict:
    return {
        "sheetId": sheet_id,
        "startRowIndex": r1, "endRowIndex": r2,
        "startColumnIndex": c1, "endColumnIndex": c2,
    }


def _source(sheet_id: int, r1: int, c1: int, r2: int, c2: int) -> dict:
    return {"sourceRange": {"sources": [_cell_range(sheet_id, r1, c1, r2, c2)]}}


def _text_fmt(hex_color: str, size: float = 10, bold: bool = False) -> dict:
    return {"textFormat": {"foregroundColor": _rgb_obj(hex_color),
                           "fontSize": size, "bold": bold}}


def _base_chart_spec(title: str, chart_type: str, legend: str = "BOTTOM_LEGEND") -> dict:
    # Sheets API uses NO_LEGEND not NONE
    legend_val = "NO_LEGEND" if legend == "NONE" else legend
    return {
        "title": title,
        "titleTextFormat": {
            "foregroundColor": _rgb_obj(C.MAZE_DARK),
            "fontSize": 12,
            "bold": True,
        },
        "backgroundColor": _rgb_obj(C.MAZE_WHITE),
        "basicChart": {
            "chartType": chart_type,
            "legendPosition": legend_val,
            "headerCount": 1,
            "axis": [
                {"position": "BOTTOM_AXIS"},
                {"position": "LEFT_AXIS"},
            ],
        },
    }


# ─────────────────────────────────────────────────────────────────────────────
# Sheet creation helpers
# ─────────────────────────────────────────────────────────────────────────────

def _create_sheet(sheets_service, title: str) -> tuple:
    """Create a new spreadsheet; return (spreadsheet_id, sheet_id=0)."""
    ss = sheets_service.spreadsheets().create(
        body={"properties": {"title": title}, "sheets": [{"properties": {"title": "Data"}}]}
    ).execute()
    spreadsheet_id = ss["spreadsheetId"]
    sheet_id = ss["sheets"][0]["properties"]["sheetId"]
    return spreadsheet_id, sheet_id


def _write_rows(sheets_service, spreadsheet_id: str, rows: list):
    """Write rows to Sheet starting at A1."""
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Data!A1",
        valueInputOption="RAW",
        body={"values": rows},
    ).execute()


def _add_chart_to_sheet(sheets_service, spreadsheet_id: str, sheet_id: int,
                        chart_spec: dict) -> int:
    """Add a chart to the sheet, return chart_id."""
    resp = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{
            "addChart": {
                "chart": {
                    "spec": chart_spec,
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": sheet_id,
                                "rowIndex": 0,
                                "columnIndex": 3,
                            },
                            "widthPixels": 900,
                            "heightPixels": 500,
                        }
                    },
                }
            }
        }]}
    ).execute()
    return resp["replies"][0]["addChart"]["chart"]["chartId"]


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC CHART BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def create_bar_chart(
    sheets_service,
    categories: list,
    values: list,
    title: str = "",
    color: str = None,
) -> tuple:
    """Vertical bar (COLUMN) chart. Returns (spreadsheet_id, chart_id)."""
    color = color or C.MAZE_PURPLE
    spreadsheet_id, sheet_id = _create_sheet(sheets_service, f"chart_{title[:30]}")

    rows = [["Category", "Value"]] + [[c, v] for c, v in zip(categories, values)]
    _write_rows(sheets_service, spreadsheet_id, rows)

    n = len(categories)
    spec = _base_chart_spec(title, "COLUMN", "NONE")
    spec["basicChart"]["domains"] = [{"domain": _source(sheet_id, 0, 0, n + 1, 1)}]
    spec["basicChart"]["series"] = [{
        "series": _source(sheet_id, 0, 1, n + 1, 2),
        "targetAxis": "LEFT_AXIS",
        "color": _rgb_obj(color),
    }]

    chart_id = _add_chart_to_sheet(sheets_service, spreadsheet_id, sheet_id, spec)
    return spreadsheet_id, chart_id


def create_hbar_chart(
    sheets_service,
    categories: list,
    values: list,
    title: str = "",
    color: str = None,
) -> tuple:
    """Horizontal bar chart. Returns (spreadsheet_id, chart_id)."""
    color = color or C.MAZE_PURPLE
    spreadsheet_id, sheet_id = _create_sheet(sheets_service, f"chart_{title[:30]}")

    rows = [["Category", "Value"]] + [[c, v] for c, v in zip(categories, values)]
    _write_rows(sheets_service, spreadsheet_id, rows)

    n = len(categories)
    spec = _base_chart_spec(title, "BAR", "NONE")
    spec["basicChart"]["domains"] = [{"domain": _source(sheet_id, 0, 0, n + 1, 1)}]
    spec["basicChart"]["series"] = [{
        "series": _source(sheet_id, 0, 1, n + 1, 2),
        "targetAxis": "BOTTOM_AXIS",
        "color": _rgb_obj(color),
    }]

    chart_id = _add_chart_to_sheet(sheets_service, spreadsheet_id, sheet_id, spec)
    return spreadsheet_id, chart_id


def create_stacked_bar_chart(
    sheets_service,
    categories: list,
    series: dict,    # {"Series A": [v1,...], ...}
    title: str = "",
    palette: list = None,
    percent: bool = False,
) -> tuple:
    """Stacked column chart. Returns (spreadsheet_id, chart_id)."""
    palette = palette or _SERIES_COLORS
    spreadsheet_id, sheet_id = _create_sheet(sheets_service, f"chart_{title[:30]}")

    series_names = list(series.keys())
    header_row = ["Category"] + series_names
    rows = [header_row]
    for i, cat in enumerate(categories):
        rows.append([cat] + [series[s][i] for s in series_names])
    _write_rows(sheets_service, spreadsheet_id, rows)

    n = len(categories)
    m = len(series_names)
    chart_type = "COLUMN"
    spec = _base_chart_spec(title, chart_type, "BOTTOM_LEGEND")
    spec["basicChart"]["stackedType"] = "PERCENT_STACKED" if percent else "STACKED"
    spec["basicChart"]["domains"] = [{"domain": _source(sheet_id, 0, 0, n + 1, 1)}]
    spec["basicChart"]["series"] = [
        {
            "series": _source(sheet_id, 0, c + 1, n + 1, c + 2),
            "targetAxis": "LEFT_AXIS",
            "color": _rgb_obj(palette[c % len(palette)]),
        }
        for c in range(m)
    ]

    chart_id = _add_chart_to_sheet(sheets_service, spreadsheet_id, sheet_id, spec)
    return spreadsheet_id, chart_id


def create_line_chart(
    sheets_service,
    x_labels: list,
    series: dict,    # {"Series A": [v1,...], ...}
    title: str = "",
    palette: list = None,
) -> tuple:
    """Line chart. Returns (spreadsheet_id, chart_id)."""
    palette = palette or _SERIES_COLORS
    spreadsheet_id, sheet_id = _create_sheet(sheets_service, f"chart_{title[:30]}")

    series_names = list(series.keys())
    header_row = ["Period"] + series_names
    rows = [header_row]
    for i, x in enumerate(x_labels):
        rows.append([x] + [series[s][i] for s in series_names])
    _write_rows(sheets_service, spreadsheet_id, rows)

    n = len(x_labels)
    m = len(series_names)
    spec = _base_chart_spec(title, "LINE", "BOTTOM_LEGEND" if m > 1 else "NONE")
    spec["basicChart"]["domains"] = [{"domain": _source(sheet_id, 0, 0, n + 1, 1)}]
    spec["basicChart"]["series"] = [
        {
            "series": _source(sheet_id, 0, c + 1, n + 1, c + 2),
            "targetAxis": "LEFT_AXIS",
            "color": _rgb_obj(palette[c % len(palette)]),
            "lineStyle": {"width": 3},
            "pointStyle": {"shape": "CIRCLE", "size": 6},
        }
        for c in range(m)
    ]

    chart_id = _add_chart_to_sheet(sheets_service, spreadsheet_id, sheet_id, spec)
    return spreadsheet_id, chart_id


def create_donut_chart(
    sheets_service,
    labels: list,
    values: list,
    title: str = "",
    palette: list = None,
) -> tuple:
    """Pie/donut chart. Returns (spreadsheet_id, chart_id)."""
    palette = palette or _SERIES_COLORS
    spreadsheet_id, sheet_id = _create_sheet(sheets_service, f"chart_{title[:30]}")

    rows = [["Label", "Value"]] + [[l, v] for l, v in zip(labels, values)]
    _write_rows(sheets_service, spreadsheet_id, rows)

    n = len(labels)
    spec = {
        "title": title,
        "titleTextFormat": {
            "foregroundColor": _rgb_obj(C.MAZE_DARK),
            "fontSize": 12,
            "bold": True,
        },
        "backgroundColor": _rgb_obj(C.MAZE_WHITE),
        "pieChart": {
            "legendPosition": "RIGHT_LEGEND",
            "domain": _source(sheet_id, 0, 0, n + 1, 1),
            "series": _source(sheet_id, 0, 1, n + 1, 2),
            "pieHole": 0.45,
            "threeDimensional": False,
        },
    }

    chart_id = _add_chart_to_sheet(sheets_service, spreadsheet_id, sheet_id, spec)
    return spreadsheet_id, chart_id
