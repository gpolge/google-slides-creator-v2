"""
High-level slide builder: composes slides from layout blocks.
Handles the full lifecycle: create → populate → add charts → return URL.
"""

from . import slides_api as api
from . import colors as C
from . import sheets_charts as sc
from .drive import OUTPUT_FOLDER_ID

# Google Slides canvas: 10 × 5.625 inches (widescreen 16:9)
SLIDE_W = 10.0
SLIDE_H = 5.625
MARGIN  = 0.4


class DeckBuilder:
    """
    Build a Google Slides presentation programmatically.

    Usage:
        builder = DeckBuilder(slides_service, drive_service, "My Deck")
        builder.add_cover("Account Review", "Q1 2026", "Acme Corp")
        builder.add_section_title("Engagement Overview")
        builder.add_chart_slide("Credits by Month", chart_png_bytes, stats=[...])
        url = builder.build()
    """

    def __init__(self, slides_service, drive_service, sheets_service,
                 title: str = None, presentation_id: str = None):
        """
        Pass `presentation_id` to append slides to an existing presentation.
        Pass `title` to create a new one.
        """
        self.slides  = slides_service
        self.drive   = drive_service
        self.sheets  = sheets_service
        self._reqs   = []
        self._slides_order = []
        self._chart_sheet_ids = []

        if presentation_id:
            # Append to existing deck
            self.pres_id = presentation_id
            self.title   = api.get_presentation(slides_service, presentation_id).get("title", "AI Slides")
        else:
            # Create new deck then move it into the shared output folder
            self.title = title or "AI Slides"
            pres = api.create_presentation(slides_service, self.title)
            self.pres_id = pres["presentationId"]
            # Move into shared folder
            try:
                f = drive_service.files().get(fileId=self.pres_id, fields="parents").execute()
                previous_parents = ",".join(f.get("parents", []))
                drive_service.files().update(
                    fileId=self.pres_id,
                    addParents=OUTPUT_FOLDER_ID,
                    removeParents=previous_parents,
                    fields="id,parents",
                ).execute()
            except Exception:
                pass  # non-fatal if folder move fails
            # Delete the default blank slide that Google adds automatically
            default_slide = pres["slides"][0]["objectId"]
            self._flush([api.delete_slide(default_slide)])

    # ── Internal helpers ──────────────────────────────────────────────────

    def _flush(self, reqs: list):
        """Send a batch of requests immediately."""
        api.batch_update(self.slides, self.pres_id, reqs)

    def _new_slide(self, bg: str = C.SLIDE_BG) -> str:
        slide_id = api.new_id()
        reqs = [api.add_slide(slide_id, layout="BLANK")]
        reqs.append(api.set_slide_background(slide_id, bg))
        self._flush(reqs)
        self._slides_order.append(slide_id)
        return slide_id

    def _embed_chart(self, slide_id: str, spreadsheet_id: str, chart_id: int,
                     x: float, y: float, w: float, h: float):
        """Embed a Sheets chart into the current slide."""
        if spreadsheet_id not in self._chart_sheet_ids:
            self._chart_sheet_ids.append(spreadsheet_id)
        req = api.add_sheets_chart(api.new_id(), slide_id,
                                   spreadsheet_id, chart_id, x, y, w, h)
        self._flush([req])

    def cleanup_on_failure(self):
        """Delete the presentation and all chart sheets if the build failed."""
        try:
            self.drive.files().delete(fileId=self.pres_id).execute()
        except Exception:
            pass
        for ss_id in self._chart_sheet_ids:
            try:
                self.drive.files().delete(fileId=ss_id).execute()
            except Exception:
                pass

    # ── Accent bar (left edge decoration) ────────────────────────────────

    def _accent_bar(self, slide_id: str, color: str = C.MAZE_PURPLE) -> list:
        bar_id = api.new_id()
        return api.add_shape(bar_id, slide_id, "RECTANGLE",
                             x=0, y=0, w=0.04, h=SLIDE_H, fill_color=color)

    # ── SLIDE TYPES ───────────────────────────────────────────────────────

    def add_cover(
        self,
        title: str,
        subtitle: str = "",
        account: str = "",
        accent_color: str = C.MAZE_PURPLE,
    ) -> str:
        """Dark cover slide with large title."""
        slide_id = self._new_slide(bg=C.MAZE_DARK)
        reqs = []

        # Purple accent blob (background decoration)
        blob_id = api.new_id()
        reqs += api.add_shape(blob_id, slide_id, "ELLIPSE",
                              x=6.8, y=-1.2, w=4.5, h=4.5,
                              fill_color=accent_color)
        # Make blob semi-transparent via opacity — not directly supported in
        # Slides API, so we use a muted tint instead
        blob2_id = api.new_id()
        reqs += api.add_shape(blob2_id, slide_id, "ELLIPSE",
                              x=7.5, y=2.8, w=3.0, h=3.0,
                              fill_color="#2D1060")

        # Eyebrow
        if account:
            reqs += api.add_text_box(
                api.new_id(), slide_id, account.upper(),
                x=MARGIN, y=0.55, w=6, h=0.35,
                font_size_pt=9, bold=True, color=accent_color,
            )

        # Title
        reqs += api.add_text_box(
            api.new_id(), slide_id, title,
            x=MARGIN, y=1.05, w=7.5, h=1.8,
            font_size_pt=40, bold=True, color=C.MAZE_WHITE,
        )

        # Subtitle
        if subtitle:
            reqs += api.add_text_box(
                api.new_id(), slide_id, subtitle,
                x=MARGIN, y=2.9, w=6.5, h=0.6,
                font_size_pt=16, color="#C4B5E0",
            )

        # Bottom bar
        bar_id = api.new_id()
        reqs += api.add_shape(bar_id, slide_id, "RECTANGLE",
                              x=0, y=SLIDE_H - 0.06, w=SLIDE_W, h=0.06,
                              fill_color=accent_color)

        self._flush(reqs)
        return slide_id

    def add_section_title(
        self,
        section: str,
        description: str = "",
        accent_color: str = C.MAZE_PURPLE,
    ) -> str:
        """Section divider slide."""
        slide_id = self._new_slide(bg=C.MAZE_DARK)
        reqs = []

        # Left accent bar
        reqs += self._accent_bar(slide_id, accent_color)

        # Pill label
        pill_id = api.new_id()
        reqs += api.add_shape(pill_id, slide_id, "ROUND_RECTANGLE",
                              x=MARGIN + 0.12, y=1.8, w=1.4, h=0.36,
                              fill_color=accent_color)
        reqs += api.add_text_box(
            api.new_id(), slide_id, "SECTION",
            x=MARGIN + 0.12, y=1.8, w=1.4, h=0.36,
            font_size_pt=8, bold=True, color=C.MAZE_WHITE, align="CENTER",
        )

        # Section title
        reqs += api.add_text_box(
            api.new_id(), slide_id, section,
            x=MARGIN + 0.12, y=2.25, w=7.5, h=1.2,
            font_size_pt=34, bold=True, color=C.MAZE_WHITE,
        )

        if description:
            reqs += api.add_text_box(
                api.new_id(), slide_id, description,
                x=MARGIN + 0.12, y=3.55, w=6.5, h=0.7,
                font_size_pt=14, color="#C4B5E0",
            )

        self._flush(reqs)
        return slide_id

    def add_text_slide(
        self,
        title: str,
        body: str,
        eyebrow: str = "",
        accent_color: str = C.MAZE_PURPLE,
        dark: bool = False,
    ) -> str:
        """Simple titled text slide."""
        bg = C.MAZE_DARK if dark else C.SLIDE_BG
        slide_id = self._new_slide(bg=bg)
        reqs = []
        reqs += self._accent_bar(slide_id, accent_color)
        text_color = C.MAZE_WHITE if dark else C.MAZE_DARK
        sub_color  = "#C4B5E0" if dark else C.MAZE_GRAY

        if eyebrow:
            reqs += api.add_text_box(
                api.new_id(), slide_id, eyebrow.upper(),
                x=MARGIN + 0.12, y=0.32, w=9, h=0.28,
                font_size_pt=8, bold=True, color=accent_color,
            )

        reqs += api.add_text_box(
            api.new_id(), slide_id, title,
            x=MARGIN + 0.12, y=0.55, w=9, h=0.7,
            font_size_pt=22, bold=True, color=text_color,
        )

        # Divider line
        line_id = api.new_id()
        reqs += api.add_shape(line_id, slide_id, "RECTANGLE",
                              x=MARGIN + 0.12, y=1.28, w=9.2 - MARGIN, h=0.025,
                              fill_color=C.MAZE_BORDER if not dark else "#2D1060")

        reqs += api.add_text_box(
            api.new_id(), slide_id, body,
            x=MARGIN + 0.12, y=1.45, w=9.2, h=3.8,
            font_size_pt=13, color=sub_color,
        )

        self._flush(reqs)
        return slide_id

    def add_stat_row(
        self,
        title: str,
        stats: list,    # [{"value": "84%", "label": "Adoption rate", "context": "..."}, ...]
        eyebrow: str = "",
        accent_color: str = C.MAZE_PURPLE,
        dark: bool = False,
    ) -> str:
        """Slide with a row of 2–4 KPI stat boxes."""
        bg = C.MAZE_DARK if dark else C.SLIDE_BG
        slide_id = self._new_slide(bg=bg)
        reqs = []
        reqs += self._accent_bar(slide_id, accent_color)
        text_color = C.MAZE_WHITE if dark else C.MAZE_DARK
        card_bg    = "#2D1060" if dark else C.MAZE_MUTED

        if eyebrow:
            reqs += api.add_text_box(
                api.new_id(), slide_id, eyebrow.upper(),
                x=MARGIN + 0.12, y=0.32, w=9, h=0.28,
                font_size_pt=8, bold=True, color=accent_color,
            )

        reqs += api.add_text_box(
            api.new_id(), slide_id, title,
            x=MARGIN + 0.12, y=0.55, w=9, h=0.65,
            font_size_pt=22, bold=True, color=text_color,
        )

        n = len(stats)
        card_w = (SLIDE_W - MARGIN * 2 - 0.12 - (n - 1) * 0.18) / n
        card_x = MARGIN + 0.12
        card_y = 1.55

        for stat in stats:
            box_id = api.new_id()
            reqs += api.add_shape(box_id, slide_id, "ROUND_RECTANGLE",
                                  x=card_x, y=card_y, w=card_w, h=2.9,
                                  fill_color=card_bg)

            val_color = accent_color
            reqs += api.add_text_box(
                api.new_id(), slide_id, stat["value"],
                x=card_x + 0.15, y=card_y + 0.5, w=card_w - 0.3, h=1.0,
                font_size_pt=32, bold=True, color=val_color, align="CENTER",
            )
            reqs += api.add_text_box(
                api.new_id(), slide_id, stat["label"],
                x=card_x + 0.1, y=card_y + 1.55, w=card_w - 0.2, h=0.55,
                font_size_pt=11, bold=True, color=text_color, align="CENTER",
            )
            if stat.get("context"):
                reqs += api.add_text_box(
                    api.new_id(), slide_id, stat["context"],
                    x=card_x + 0.1, y=card_y + 2.15, w=card_w - 0.2, h=0.55,
                    font_size_pt=9, color=C.MAZE_GRAY if not dark else "#9B82C4",
                    align="CENTER",
                )

            card_x += card_w + 0.18

        self._flush(reqs)
        return slide_id

    def add_chart_slide(
        self,
        title: str,
        spreadsheet_id: str,
        chart_id: int,
        eyebrow: str = "",
        stats: list = None,         # optional [{"value":..,"label":..}] on the right
        accent_color: str = C.MAZE_PURPLE,
    ) -> str:
        """
        Slide with a Sheets chart (left/full width) and optional stat column (right).
        """
        slide_id = self._new_slide(bg=C.SLIDE_BG)
        reqs = []
        reqs += self._accent_bar(slide_id, accent_color)

        if eyebrow:
            reqs += api.add_text_box(
                api.new_id(), slide_id, eyebrow.upper(),
                x=MARGIN + 0.12, y=0.32, w=9, h=0.28,
                font_size_pt=8, bold=True, color=accent_color,
            )

        reqs += api.add_text_box(
            api.new_id(), slide_id, title,
            x=MARGIN + 0.12, y=0.55, w=9, h=0.65,
            font_size_pt=22, bold=True, color=C.MAZE_DARK,
        )

        self._flush(reqs)

        if stats:
            chart_w, chart_x = 6.8, MARGIN + 0.12
            stat_x = chart_x + chart_w + 0.25
            stat_w = SLIDE_W - stat_x - MARGIN
        else:
            chart_w, chart_x = SLIDE_W - MARGIN * 2 - 0.12, MARGIN + 0.12
            stat_x, stat_w = None, None

        self._embed_chart(slide_id, spreadsheet_id, chart_id,
                          x=chart_x, y=1.35,
                          w=chart_w, h=SLIDE_H - 1.35 - 0.25)

        if stats and stat_x:
            stat_reqs = []
            n = len(stats)
            cell_h = (SLIDE_H - 1.35 - 0.2) / n
            for i, stat in enumerate(stats):
                sy = 1.35 + i * cell_h
                card_id = api.new_id()
                stat_reqs += api.add_shape(card_id, slide_id, "ROUND_RECTANGLE",
                                           x=stat_x, y=sy + 0.05,
                                           w=stat_w, h=cell_h - 0.12,
                                           fill_color=C.MAZE_MUTED)
                stat_reqs += api.add_text_box(
                    api.new_id(), slide_id, stat["value"],
                    x=stat_x + 0.1, y=sy + 0.18, w=stat_w - 0.2, h=cell_h * 0.48,
                    font_size_pt=22, bold=True, color=accent_color, align="CENTER",
                )
                stat_reqs += api.add_text_box(
                    api.new_id(), slide_id, stat["label"],
                    x=stat_x + 0.05, y=sy + cell_h * 0.55, w=stat_w - 0.1, h=cell_h * 0.38,
                    font_size_pt=9, color=C.MAZE_GRAY, align="CENTER",
                )
            self._flush(stat_reqs)

        return slide_id

    def add_two_chart_slide(
        self,
        title: str,
        left_spreadsheet_id: str,
        left_chart_id: int,
        right_spreadsheet_id: str,
        right_chart_id: int,
        label_left: str = "",
        label_right: str = "",
        eyebrow: str = "",
        accent_color: str = C.MAZE_PURPLE,
    ) -> str:
        """Two charts side-by-side."""
        slide_id = self._new_slide(bg=C.SLIDE_BG)
        reqs = []
        reqs += self._accent_bar(slide_id, accent_color)

        if eyebrow:
            reqs += api.add_text_box(
                api.new_id(), slide_id, eyebrow.upper(),
                x=MARGIN + 0.12, y=0.32, w=9, h=0.28,
                font_size_pt=8, bold=True, color=accent_color,
            )

        reqs += api.add_text_box(
            api.new_id(), slide_id, title,
            x=MARGIN + 0.12, y=0.55, w=9, h=0.65,
            font_size_pt=22, bold=True, color=C.MAZE_DARK,
        )

        divider_id = api.new_id()
        reqs += api.add_shape(divider_id, slide_id, "RECTANGLE",
                              x=SLIDE_W / 2, y=1.35, w=0.02, h=SLIDE_H - 1.6,
                              fill_color=C.MAZE_BORDER)

        self._flush(reqs)

        chart_w = (SLIDE_W - MARGIN * 2 - 0.12 - 0.2) / 2
        for i, (ss_id, c_id, label) in enumerate([
            (left_spreadsheet_id,  left_chart_id,  label_left),
            (right_spreadsheet_id, right_chart_id, label_right),
        ]):
            cx = MARGIN + 0.12 + i * (chart_w + 0.2)
            self._embed_chart(slide_id, ss_id, c_id,
                              x=cx, y=1.45, w=chart_w, h=SLIDE_H - 1.85)

            if label:
                self._flush(api.add_text_box(
                    api.new_id(), slide_id, label,
                    x=cx, y=SLIDE_H - 0.42, w=chart_w, h=0.28,
                    font_size_pt=9, color=C.MAZE_GRAY, align="CENTER",
                ))

        return slide_id

    def add_quote_slide(
        self,
        quote: str,
        attribution: str = "",
        accent_color: str = C.MAZE_PURPLE,
    ) -> str:
        """Large pull-quote slide."""
        slide_id = self._new_slide(bg=C.MAZE_DARK)
        reqs = []

        # Giant quote mark
        reqs += api.add_text_box(
            api.new_id(), slide_id, "\u201c",
            x=MARGIN, y=0.3, w=1.2, h=1.2,
            font_size_pt=72, bold=True, color=accent_color,
        )

        reqs += api.add_text_box(
            api.new_id(), slide_id, quote,
            x=MARGIN + 0.1, y=1.1, w=9.1, h=3.0,
            font_size_pt=20, italic=True, color=C.MAZE_WHITE,
        )

        if attribution:
            reqs += api.add_text_box(
                api.new_id(), slide_id, f"\u2014 {attribution}",
                x=MARGIN + 0.1, y=4.25, w=9.1, h=0.5,
                font_size_pt=11, color=accent_color, bold=True,
            )

        # Bottom bar
        bar_id = api.new_id()
        reqs += api.add_shape(bar_id, slide_id, "RECTANGLE",
                              x=0, y=SLIDE_H - 0.06, w=SLIDE_W, h=0.06,
                              fill_color=accent_color)

        self._flush(reqs)
        return slide_id

    def fill_template_placeholders(self, replacements: dict):
        """
        Replace {{key}} placeholders across the entire presentation.
        replacements = {"Account_name": "Acme Corp", "Period": "Q1 2026", ...}
        """
        reqs = [api.replace_all_text(k, str(v)) for k, v in replacements.items()]
        self._flush(reqs)

    # ── Finalise ─────────────────────────────────────────────────────────

    def url(self) -> str:
        return f"https://docs.google.com/presentation/d/{self.pres_id}/edit"

    def build(self) -> str:
        """Return the Google Slides URL."""
        print(f"\n✅ Presentation ready: {self.url()}\n")
        return self.url()
