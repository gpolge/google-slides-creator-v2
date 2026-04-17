"""
Engagement Demo Deck — 6 designer-quality engagement slides
Inspired by the reference deck (clean, statement headlines, data-forward)
Uses full Maze brand system with varied chart types.

Run: python3 engagement_demo.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services
from src import slides_api as api
from src import sheets_charts as sc
from src.drive import OUTPUT_FOLDER_ID

# ── Design tokens ──────────────────────────────────────────────────────────────
WHITE     = "#FFFFFF"
DARK      = "#1A0533"
PURPLE    = "#A366FF"
LIGHT_BG  = "#F5F0FF"   # light purple tint — card backgrounds
CYAN      = "#79DEC0"   # callout highlight
PINK      = "#FC1258"   # zero / at-risk values
YELLOW    = "#FFC300"
LIME      = "#B9E20E"
GRAY      = "#6B6B8A"
BORDER    = "#E5E0F5"
MID_GRAY  = "#C4BAD8"

# ── Canvas ─────────────────────────────────────────────────────────────────────
W, H  = 10.0, 5.625
ML    = 0.38    # left margin

# ── Demo data ──────────────────────────────────────────────────────────────────
PERIOD   = "OCT 2025 – APR 2026"
ACCOUNT  = "Acme Corp"
TODAY    = "Apr 16, 2026"
TOTAL_SLIDES = 6

MONTHS   = ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
CREDITS  = [28, 34, 42, 51, 68, 44]

STUDY_TYPES   = ["Usability Testing", "Survey", "Card Sort", "Concept Testing", "Moderated", "AI Moderator"]
STUDY_COUNTS  = [58, 28, 8, 4, 2, 0]

CADENCE_SERIES = {
    "Usability Testing": [8, 10, 12, 14, 9, 5],
    "Survey":            [3,  4,  5,  7,  5, 4],
    "Card Sort":         [1,  1,  2,  2,  1, 1],
}

RESEARCHERS = [
    {"num": "18", "name": "Sarah M.",       "detail": "acme.com",             "highlight": True},
    {"num": "12", "name": "James T.",       "detail": "acme.com",             "highlight": False},
    {"num": "8",  "name": "Priya K.",       "detail": "acme.com",             "highlight": False},
    {"num": "6",  "name": "Alex R.",        "detail": "acme.com · 1 moderated","highlight": False},
    {"num": "4",  "name": "Lisa C. / Tom B. / Nina P.", "detail": "+5 others at this level", "highlight": False},
    {"num": "1–3","name": "29 researchers", "detail": "Remainder of active cohort", "highlight": False},
]

CREDITS_USED  = 210
CREDITS_TOTAL = 250


# ── Low-level helpers ──────────────────────────────────────────────────────────

def _flush(svc, pres_id, reqs):
    if reqs:
        api.batch_update(svc, pres_id, reqs)

def _new_slide(svc, pres_id):
    sid = api.new_id()
    _flush(svc, pres_id, [
        api.add_slide(sid, "BLANK"),
        api.set_slide_background(sid, WHITE),
    ])
    return sid


# ── Reusable layout primitives ─────────────────────────────────────────────────

def _section_pill(slide_id, number, label):
    """Top-left tag: '03 · Research Methods'"""
    reqs = api.add_shape(
        api.new_id(), slide_id, "ROUND_RECTANGLE",
        x=ML, y=0.22, w=2.3, h=0.3,
        fill_color=WHITE, border_color=BORDER, border_width_pt=0.75,
    )
    reqs += api.add_text_box(
        api.new_id(), slide_id, f"{number} · {label}",
        x=ML + 0.06, y=0.22, w=2.2, h=0.3,
        font_size_pt=8, color=GRAY, align="CENTER", v_align="MIDDLE",
    )
    return reqs


def _eyebrow(slide_id, text, color=PURPLE):
    return api.add_text_box(
        api.new_id(), slide_id, text.upper(),
        x=ML, y=0.66, w=9.2, h=0.22,
        font_size_pt=7.5, bold=True, color=color,
    )


def _headline(slide_id, text, size=28, y=0.88, h=1.0, w=9.2):
    return api.add_text_box(
        api.new_id(), slide_id, text,
        x=ML, y=y, w=w, h=h,
        font_size_pt=size, bold=True, color=DARK,
    )


def _body(slide_id, text, x=ML, y=2.08, w=4.25, h=2.4, size=11):
    return api.add_text_box(
        api.new_id(), slide_id, text,
        x=x, y=y, w=w, h=h,
        font_size_pt=size, color=GRAY,
    )


def _callout(slide_id, label, text, x=ML, y=3.55, w=4.25, h=1.4, bg=DARK):
    """Dark callout box with a bold label and supporting text."""
    tc = WHITE
    sub_color = CYAN if bg == DARK else DARK
    reqs = api.add_shape(
        api.new_id(), slide_id, "ROUND_RECTANGLE",
        x=x, y=y, w=w, h=h, fill_color=bg,
    )
    if label:
        reqs += api.add_text_box(
            api.new_id(), slide_id, label,
            x=x + 0.18, y=y + 0.14, w=w - 0.36, h=0.28,
            font_size_pt=9, bold=True, color=sub_color,
        )
    reqs += api.add_text_box(
        api.new_id(), slide_id, text,
        x=x + 0.18, y=y + 0.38, w=w - 0.36, h=h - 0.5,
        font_size_pt=10, color=tc,
    )
    return reqs


def _footer(slide_id, page_num):
    reqs = api.add_text_box(
        api.new_id(), slide_id, TODAY,
        x=ML, y=5.35, w=2.5, h=0.2,
        font_size_pt=8, color=MID_GRAY,
    )
    reqs += api.add_text_box(
        api.new_id(), slide_id, f"{page_num} of {TOTAL_SLIDES}",
        x=8.5, y=5.35, w=1.12, h=0.2,
        font_size_pt=8, color=MID_GRAY, align="END",
    )
    return reqs


def _kpi_card(slide_id, x, y, w, h, value, label, sublabel="", dark=False, alert=False):
    """Large KPI card — dark or light variant."""
    bg        = DARK if dark else LIGHT_BG
    val_color = WHITE if dark else (PINK if alert else PURPLE)
    lbl_color = CYAN  if dark else DARK
    sub_color = "#9B82C4" if dark else GRAY
    reqs = api.add_shape(
        api.new_id(), slide_id, "ROUND_RECTANGLE",
        x=x, y=y, w=w, h=h, fill_color=bg,
    )
    reqs += api.add_text_box(
        api.new_id(), slide_id, value,
        x=x + 0.18, y=y + 0.22, w=w - 0.36, h=h * 0.44,
        font_size_pt=44, bold=True, color=val_color, align="START",
    )
    reqs += api.add_text_box(
        api.new_id(), slide_id, label.upper(),
        x=x + 0.18, y=y + h * 0.58, w=w - 0.36, h=0.28,
        font_size_pt=8, bold=True, color=lbl_color,
    )
    if sublabel:
        reqs += api.add_text_box(
            api.new_id(), slide_id, sublabel,
            x=x + 0.18, y=y + h * 0.75, w=w - 0.36, h=0.28,
            font_size_pt=8, color=sub_color,
        )
    return reqs


def _stat_box(slide_id, x, y, w, h, value, label, val_color=PURPLE):
    """Compact stat box — for the right-side column on chart slides."""
    reqs = api.add_shape(
        api.new_id(), slide_id, "ROUND_RECTANGLE",
        x=x, y=y, w=w, h=h, fill_color=LIGHT_BG,
    )
    reqs += api.add_text_box(
        api.new_id(), slide_id, value,
        x=x + 0.12, y=y + 0.1, w=w - 0.24, h=h * 0.52,
        font_size_pt=22, bold=True, color=val_color, align="CENTER",
    )
    reqs += api.add_text_box(
        api.new_id(), slide_id, label,
        x=x + 0.08, y=y + h * 0.58, w=w - 0.16, h=h * 0.36,
        font_size_pt=8.5, color=GRAY, align="CENTER",
    )
    return reqs


def _leaderboard_row(slide_id, x, y, w, num, name, detail, highlight=False):
    """Single leaderboard row."""
    reqs = []
    if highlight:
        reqs += api.add_shape(
            api.new_id(), slide_id, "ROUND_RECTANGLE",
            x=x, y=y, w=w, h=0.72, fill_color=DARK,
        )
        num_c, name_c, det_c = WHITE, WHITE, CYAN
    else:
        reqs += api.add_shape(
            api.new_id(), slide_id, "RECTANGLE",
            x=x, y=y + 0.7, w=w, h=0.012, fill_color=BORDER,
        )
        num_c, name_c, det_c = PURPLE, DARK, GRAY

    # Big number
    reqs += api.add_text_box(
        api.new_id(), slide_id, num,
        x=x + 0.12, y=y + 0.06, w=0.7, h=0.55,
        font_size_pt=26, bold=True, color=num_c,
    )
    # "consumed" micro-label
    reqs += api.add_text_box(
        api.new_id(), slide_id, "consumed",
        x=x + 0.78, y=y + 0.38, w=0.65, h=0.22,
        font_size_pt=6.5, color=det_c,
    )
    # Name
    reqs += api.add_text_box(
        api.new_id(), slide_id, name,
        x=x + 1.35, y=y + 0.08, w=w - 1.48, h=0.3,
        font_size_pt=11.5, bold=True, color=name_c,
    )
    # Detail
    reqs += api.add_text_box(
        api.new_id(), slide_id, detail,
        x=x + 1.35, y=y + 0.39, w=w - 1.48, h=0.22,
        font_size_pt=8.5, color=det_c,
    )
    return reqs


# ── Slide builders ─────────────────────────────────────────────────────────────

def slide_engagement_summary(svc, pres_id, page):
    """Slide 1: 3-KPI card layout — studies, credits, researchers."""
    sid = _new_slide(svc, pres_id)
    reqs = _section_pill(sid, "01", "Engagement Overview")
    reqs += _eyebrow(sid, f"ENGAGEMENT SUMMARY · {PERIOD}")
    reqs += _headline(sid,
        "100 studies consumed, 1,082 credits used\nacross 57 active researchers",
        size=26, y=0.88, h=1.1,
    )
    # 3 KPI cards: dark (hero) + 2 light
    cw, ch = 2.9, 3.15
    gap = 0.18
    reqs += _kpi_card(sid, x=ML,            y=2.2, w=cw, h=ch, value="100", label="Studies Consumed",  sublabel=f"{PERIOD}",         dark=True)
    reqs += _kpi_card(sid, x=ML+cw+gap,     y=2.2, w=cw, h=ch, value="1,082", label="Credits Used",  sublabel="Completed panel testers")
    reqs += _kpi_card(sid, x=ML+2*(cw+gap), y=2.2, w=cw, h=ch, value="57",  label="Active Researchers", sublabel="Active in the last 60 days")
    reqs += _footer(sid, page)
    _flush(svc, pres_id, reqs)
    return sid


def slide_monthly_trend(svc, pres_id, sheets_svc, page):
    """Slide 2: Line chart — monthly credit consumption trend."""
    sid = _new_slide(svc, pres_id)
    reqs = _section_pill(sid, "02", "Credit Trend")
    reqs += _eyebrow(sid, f"CREDIT CONSUMPTION · {PERIOD}")
    reqs += _headline(sid,
        "Credit spend climbed steadily — February peaked at 68,\nthen tapered as sprint cycles slowed",
        size=22, y=0.88, h=1.12,
    )
    reqs += _footer(sid, page)
    _flush(svc, pres_id, reqs)

    # Line chart (multi-series for visual richness — usage vs rolling avg)
    rolling_avg = [round(sum(CREDITS[:i+1]) / (i+1)) for i in range(len(CREDITS))]
    ss_id, c_id = sc.create_line_chart(
        sheets_svc,
        x_labels=MONTHS,
        series={
            "Credits used":    CREDITS,
            "Rolling average": rolling_avg,
        },
        title="",
        palette=[PURPLE, CYAN],
    )
    # Chart (left 6.8in)
    _flush(svc, pres_id, [api.add_sheets_chart(
        api.new_id(), sid, ss_id, c_id,
        x=ML, y=2.18, w=6.7, h=3.12,
    )])

    # 3 stat boxes (right column)
    peak_idx = CREDITS.index(max(CREDITS))
    sx = ML + 6.7 + 0.2
    sw = W - sx - 0.2
    sh = (3.12 - 0.18) / 3
    stats = [
        (str(max(CREDITS)),       f"Peak credits\n({MONTHS[peak_idx]})", PURPLE),
        (str(round(sum(CREDITS)/len(CREDITS))), "Monthly average",       CYAN),
        (str(sum(CREDITS)),       "Total this period",                   DARK),
    ]
    stat_reqs = []
    for i, (val, lbl, vc) in enumerate(stats):
        stat_reqs += _stat_box(sid, sx, 2.18 + i * (sh + 0.09), sw, sh, val, lbl, val_color=vc)
    _flush(svc, pres_id, stat_reqs)
    return sid, ss_id


def slide_research_methods(svc, pres_id, sheets_svc, page):
    """Slide 3: Horizontal bar chart — study type breakdown, split layout."""
    sid = _new_slide(svc, pres_id)
    reqs = _section_pill(sid, "03", "Research Methods")
    reqs += _eyebrow(sid, "METHOD BREAKDOWN · STUDIES CREATED")
    reqs += _headline(sid,
        "Usability testing dominates — AI moderated\nand moderated studies remain untapped",
        size=22, y=0.88, h=1.12,
    )
    reqs += _body(sid,
        "Prototype and usability testing is the team's default method. "
        "Surveys support discovery work effectively. Card sorting covers IA. "
        "Both moderated formats remain virtually unused.",
        y=2.08, w=4.25, h=1.35,
    )
    reqs += _callout(sid,
        label="The gap",
        text="0 AI Moderated studies and just 2 moderated — the two methods "
             "your team identified as priorities for Q2 research.",
        y=3.55, w=4.25, h=1.72, bg=DARK,
    )
    reqs += _footer(sid, page)
    _flush(svc, pres_id, reqs)

    # Horizontal bar chart — reverse order so biggest is at top
    ss_id, c_id = sc.create_hbar_chart(
        sheets_svc,
        categories=list(reversed(STUDY_TYPES)),
        values=list(reversed(STUDY_COUNTS)),
        title="",
        color=PURPLE,
    )
    _flush(svc, pres_id, [api.add_sheets_chart(
        api.new_id(), sid, ss_id, c_id,
        x=4.82, y=0.65, w=4.8, h=4.7,
    )])
    return sid, ss_id


def slide_top_researchers(svc, pres_id, page):
    """Slide 4: Leaderboard — top researchers by studies consumed."""
    sid = _new_slide(svc, pres_id)
    reqs = _section_pill(sid, "04", "Top Researchers")
    reqs += _eyebrow(sid, "RESEARCHERS · STUDIES CONSUMED")
    reqs += _headline(sid,
        "Sarah M. leads with 18 consumed — 8 researchers\nat 4 or more studies this period",
        size=22, y=0.88, h=1.12,
    )
    reqs += _body(sid,
        "Consumption is well distributed across teams. "
        "Growing this top tier from 8 to 15+ researchers is the "
        "lever for compounding ROI ahead of renewal.",
        y=2.08, w=4.25, h=1.28,
    )
    # Callout: opportunity framing
    reqs += _callout(sid,
        label="The opportunity",
        text="29 researchers used 1–3 studies. Converting 10 of them to "
             "4+ studies each would add 30+ studies to your annual total.",
        y=3.45, w=4.25, h=1.82, bg=PURPLE,
    )
    reqs += _footer(sid, page)

    # Leaderboard rows (right column)
    rx = 4.82
    rw = W - rx - 0.2
    ry = 0.62
    row_h = 0.77
    for r in RESEARCHERS:
        reqs += _leaderboard_row(sid, rx, ry, rw,
                                 r["num"], r["name"], r["detail"],
                                 highlight=r["highlight"])
        ry += row_h

    _flush(svc, pres_id, reqs)
    return sid


def slide_credit_utilisation(svc, pres_id, sheets_svc, page):
    """Slide 5: Donut chart — credit utilisation vs commitment."""
    used = CREDITS_USED
    total = CREDITS_TOTAL
    pct = round(used / total * 100)

    sid = _new_slide(svc, pres_id)
    reqs = _section_pill(sid, "05", "Credit Utilisation")
    reqs += _eyebrow(sid, "CREDITS · UTILISATION VS COMMITMENT")
    reqs += _headline(sid,
        f"{pct}% of credits consumed — on track\nbut little headroom for Q2 expansion",
        size=24, y=0.88, h=1.12,
    )
    reqs += _footer(sid, page)
    _flush(svc, pres_id, reqs)

    # Donut chart
    ss_id, c_id = sc.create_donut_chart(
        sheets_svc,
        labels=["Credits Used", "Remaining"],
        values=[used, total - used],
        title="",
        palette=[PURPLE, LIGHT_BG],
    )
    _flush(svc, pres_id, [api.add_sheets_chart(
        api.new_id(), sid, ss_id, c_id,
        x=ML, y=2.08, w=5.2, h=3.18,
    )])

    # 3 stat boxes + a note
    sx = ML + 5.2 + 0.3
    sw = W - sx - 0.2
    sh = 0.88
    stat_reqs = []
    stat_reqs += _stat_box(sid, sx, 2.08,           sw, sh, str(used),         "Credits used",      val_color=PURPLE)
    stat_reqs += _stat_box(sid, sx, 2.08 + sh + 0.12, sw, sh, str(total),       "Total commitment",  val_color=DARK)
    stat_reqs += _stat_box(sid, sx, 2.08 + 2*(sh+0.12), sw, sh, str(total-used), "Remaining credits", val_color=CYAN)

    # Utilisation bar
    bar_total_w = sw
    bar_used_w  = bar_total_w * (used / total)
    bar_y = 2.08 + 3 * (sh + 0.12) + 0.1
    stat_reqs += api.add_shape(api.new_id(), sid, "ROUND_RECTANGLE",
                               x=sx, y=bar_y, w=bar_total_w, h=0.18, fill_color=BORDER)
    stat_reqs += api.add_shape(api.new_id(), sid, "ROUND_RECTANGLE",
                               x=sx, y=bar_y, w=bar_used_w,  h=0.18, fill_color=PURPLE)
    stat_reqs += api.add_text_box(api.new_id(), sid, f"{pct}% used",
                                  x=sx, y=bar_y + 0.22, w=sw, h=0.22,
                                  font_size_pt=8, color=GRAY, align="CENTER")
    _flush(svc, pres_id, stat_reqs)
    return sid, ss_id


def slide_study_cadence(svc, pres_id, sheets_svc, page):
    """Slide 6: Stacked bar chart — study types by month, full width."""
    sid = _new_slide(svc, pres_id)
    reqs = _section_pill(sid, "06", "Monthly Cadence")
    reqs += _eyebrow(sid, "STUDY CADENCE · STUDY TYPES BY MONTH")
    reqs += _headline(sid,
        "Two study types drove all activity — survey usage doubled in January "
        "while usability testing peaked in February",
        size=20, y=0.88, h=1.0, w=9.2,
    )
    reqs += _footer(sid, page)
    _flush(svc, pres_id, reqs)

    # Stacked bar chart — full width
    ss_id, c_id = sc.create_stacked_bar_chart(
        sheets_svc,
        categories=MONTHS,
        series=CADENCE_SERIES,
        title="",
        palette=[PURPLE, CYAN, YELLOW],
    )
    _flush(svc, pres_id, [api.add_sheets_chart(
        api.new_id(), sid, ss_id, c_id,
        x=ML, y=2.05, w=W - ML * 2, h=3.22,
    )])
    return sid, ss_id


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("Authenticating with Google...")
    slides_svc, drive_svc, sheets_svc = get_services()

    # Create presentation in shared folder
    title = f"Maze QBR — Engagement Slides Demo"
    pres = api.create_presentation(slides_svc, title)
    pres_id = pres["presentationId"]

    # Move to shared folder + delete the default blank slide
    f = drive_svc.files().get(fileId=pres_id, fields="parents").execute()
    drive_svc.files().update(
        fileId=pres_id,
        addParents=OUTPUT_FOLDER_ID,
        removeParents=",".join(f.get("parents", [])),
        fields="id,parents",
    ).execute()
    api.batch_update(slides_svc, pres_id, [api.delete_slide(pres["slides"][0]["objectId"])])

    chart_sheet_ids = []

    try:
        print("Building slide 1: Engagement Summary...")
        slide_engagement_summary(slides_svc, pres_id, page=1)

        print("Building slide 2: Monthly Credit Trend...")
        _, ss = slide_monthly_trend(slides_svc, pres_id, sheets_svc, page=2)
        chart_sheet_ids.append(ss)

        print("Building slide 3: Research Methods...")
        _, ss = slide_research_methods(slides_svc, pres_id, sheets_svc, page=3)
        chart_sheet_ids.append(ss)

        print("Building slide 4: Top Researchers...")
        slide_top_researchers(slides_svc, pres_id, page=4)

        print("Building slide 5: Credit Utilisation...")
        _, ss = slide_credit_utilisation(slides_svc, pres_id, sheets_svc, page=5)
        chart_sheet_ids.append(ss)

        print("Building slide 6: Study Cadence...")
        _, ss = slide_study_cadence(slides_svc, pres_id, sheets_svc, page=6)
        chart_sheet_ids.append(ss)

        url = f"https://docs.google.com/presentation/d/{pres_id}/edit"
        print(f"\n✅  Deck ready: {url}\n")
        return url

    except Exception as e:
        print(f"\n❌  Build failed: {e}")
        print("Cleaning up...")
        try:
            drive_svc.files().delete(fileId=pres_id).execute()
        except Exception:
            pass
        for ss_id in chart_sheet_ids:
            try:
                drive_svc.files().delete(fileId=ss_id).execute()
            except Exception:
                pass
        raise


if __name__ == "__main__":
    main()
