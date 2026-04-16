"""
Demo: Build a Maze-branded Google Slides deck from scratch using native Google APIs.

Charts are created in Google Sheets (with Maze brand colors) and linked into the
presentation via createSheetsChart — no external image hosting required.

Run:  python3 demo.py
"""

import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services
from src.builder import DeckBuilder
from src import sheets_charts as sc

CONFIG_FILE = Path(__file__).parent / "deck_config.json"


def main():
    print("Authenticating with Google...")
    slides_svc, drive_svc, sheets_svc = get_services()

    # Always append to the single persistent "AI Slides" presentation
    cfg = json.loads(CONFIG_FILE.read_text())
    pres_id = cfg["presentation_id"]
    print(f"Using persistent deck: https://docs.google.com/presentation/d/{pres_id}/edit")

    builder = DeckBuilder(
        slides_svc,
        drive_svc,
        sheets_svc,
        presentation_id=pres_id,
    )

    try:
        # ── 1. COVER ──────────────────────────────────────────────────────────
        print("Building cover slide...")
        builder.add_cover(
            title="Research Engagement Review",
            subtitle="April 2026 · Built with Maze Slides API",
            account="Acme Corp",
        )

        # ── 2. SECTION INTRO ─────────────────────────────────────────────────
        print("Adding section divider...")
        builder.add_section_title(
            section="Engagement Overview",
            description="How your team is using Maze to ship better products",
        )

        # ── 3. STAT ROW ───────────────────────────────────────────────────────
        print("Adding stat row...")
        builder.add_stat_row(
            title="Snapshot: Last 6 Months",
            eyebrow="Key Metrics",
            stats=[
                {"value": "1,082",  "label": "Credits used",   "context": "+18% vs prior period"},
                {"value": "57",     "label": "Researchers",    "context": "Across 4 teams"},
                {"value": "100",    "label": "Studies run",    "context": "Avg 10.8 credits each"},
                {"value": "84%",    "label": "Adoption rate",  "context": "of licensed seats active"},
            ],
        )

        # ── 4. BAR CHART SLIDE ────────────────────────────────────────────────
        print("Creating bar chart in Google Sheets...")
        ss_id, chart_id = sc.create_bar_chart(
            sheets_svc,
            categories=["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
            values=[142, 168, 195, 210, 176, 191],
            title="Monthly Credit Consumption",
        )
        builder.add_chart_slide(
            title="Credit Usage — Monthly Trend",
            eyebrow="Engagement",
            spreadsheet_id=ss_id,
            chart_id=chart_id,
            stats=[
                {"value": "1,082",  "label": "Total credits"},
                {"value": "+18%",   "label": "YoY growth"},
                {"value": "180",    "label": "Avg / month"},
            ],
        )

        # ── 5. LINE CHART SLIDE ───────────────────────────────────────────────
        print("Creating line chart in Google Sheets...")
        ss_line, cid_line = sc.create_line_chart(
            sheets_svc,
            x_labels=["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
            series={
                "Active researchers": [34, 38, 44, 52, 49, 57],
                "New researchers":    [8, 6, 9, 12, 5, 11],
            },
            title="Researcher Growth",
        )
        builder.add_chart_slide(
            title="Researcher Growth Over Time",
            eyebrow="Deep Dive",
            spreadsheet_id=ss_line,
            chart_id=cid_line,
            stats=[
                {"value": "57",   "label": "Active this month"},
                {"value": "+68%", "label": "Growth since Oct"},
            ],
        )

        # ── 6. DONUT + STACKED BAR (TWO CHARTS) ──────────────────────────────
        print("Creating donut chart in Google Sheets...")
        ss_donut, cid_donut = sc.create_donut_chart(
            sheets_svc,
            labels=["Usability Testing", "Concept Testing", "Card Sorting",
                    "Tree Testing", "Surveys"],
            values=[38, 27, 16, 12, 7],
            title="Study Type Mix",
        )
        print("Creating stacked bar chart in Google Sheets...")
        ss_stack, cid_stack = sc.create_stacked_bar_chart(
            sheets_svc,
            categories=["Product", "Research", "Design", "Marketing"],
            series={
                "Usability Testing": [28, 45, 18, 5],
                "Concept Testing":   [15, 20, 30, 8],
                "Card Sorting":      [8, 12, 6, 2],
                "Other":             [5, 8, 4, 3],
            },
            title="Credits by Team × Study Type",
        )
        builder.add_two_chart_slide(
            title="Research Mix & Team Breakdown",
            eyebrow="Deep Dive",
            left_spreadsheet_id=ss_donut,
            left_chart_id=cid_donut,
            right_spreadsheet_id=ss_stack,
            right_chart_id=cid_stack,
            label_left="Study type distribution",
            label_right="Credits by team and study type",
        )

        # ── 7. HORIZONTAL BAR (TOP RESEARCHERS) ──────────────────────────────
        print("Creating horizontal bar chart in Google Sheets...")
        ss_hbar, cid_hbar = sc.create_hbar_chart(
            sheets_svc,
            categories=["Sarah M.", "James T.", "Priya K.", "Lena W.", "Tom A."],
            values=[124, 98, 87, 72, 61],
            title="Top 5 Researchers by Credits",
        )
        builder.add_chart_slide(
            title="Top Researchers",
            eyebrow="Spotlight",
            spreadsheet_id=ss_hbar,
            chart_id=cid_hbar,
            stats=[
                {"value": "5",    "label": "Power users"},
                {"value": "442",  "label": "Combined credits"},
                {"value": "41%",  "label": "Share of total"},
            ],
        )

        # ── 8. QUOTE ──────────────────────────────────────────────────────────
        print("Adding quote slide...")
        builder.add_quote_slide(
            quote="Since rolling out Maze across our product teams, "
                  "we've cut our research cycle from 3 weeks to 4 days.",
            attribution="Head of Product Research, Acme Corp",
        )

        # ── 9. SECTION: NEXT STEPS ────────────────────────────────────────────
        builder.add_section_title(
            section="Next Steps & Recommendations",
            description="How we can help you get even more from Maze",
            accent_color="#79DEC0",
        )

        builder.add_text_slide(
            title="Recommendations for Q2 2026",
            eyebrow="Next Steps",
            body="1.  Expand Maze access to the Marketing team — high demand, currently no licenses\n\n"
                 "2.  Set up study templates for recurring usability cycles\n\n"
                 "3.  Connect Maze to your Jira workflow via the native integration\n\n"
                 "4.  Schedule a deep-dive session on Card Sorting best practices",
            accent_color="#79DEC0",
        )

        # ── DONE ──────────────────────────────────────────────────────────────
        url = builder.build()
        print(f"\nOpen your deck: {url}")
        return url

    except Exception as e:
        print(f"\nBuild failed: {e}")
        print("Cleaning up partial files from Drive...")
        builder.cleanup_on_failure()
        raise


if __name__ == "__main__":
    main()
