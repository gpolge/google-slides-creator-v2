"""
Jaguar Land Rover <> Maze Partnership Review — Q1 2026

Follows the exact same narrative logic as the Absa reference deck:
  Cover → Agenda → Introductions → Business Priorities → Subscription →
  Research Adoption → Credits Usage → Industry Benchmarks →
  Product Updates section → Feature spotlights → Next Steps

Uses the official QBR template design exactly (fonts, colors, layouts).
Copies template → fills all placeholders → injects usage chart.

Run: python3 qbr_jlr.py
"""

import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services
from src import slides_api as api
from src import sheets_charts as sc
from src.drive import copy_template, OUTPUT_FOLDER_ID

TEMPLATE_ID = "1KpYl6uTcUMBUFCHnMGkRXJVuxhcdsHrhDnzbZRQKNmo"

# ─────────────────────────────────────────────────────────────────────────────
# JLR ACCOUNT DATA
# ─────────────────────────────────────────────────────────────────────────────

CUSTOMER = {
    "name":    "Jaguar Land Rover",
    "short":   "JLR",
    "date":    "16 April 2026",
    "quarter": "Q1 2026",

    # --- Goals (same structure as Absa: Speed / Scale / Impact) ---
    "goal_1_title":  "Speed",
    "goal_1_desc":   (
        "Run digital UX studies within 48 hours of a design being ready — "
        "keeping pace with sprint cadence across Connected Car and Digital Products."
    ),
    "goal_1_update": "On track",

    "goal_2_title":  "Scale",
    "goal_2_desc":   (
        "Enable product designers and PMs across all JLR digital teams to run "
        "their own unmoderated studies without research team bottlenecks."
    ),
    "goal_2_update": "In progress",

    "goal_3_title":  "Impact",
    "goal_3_desc":   (
        "Validate every Pivi Pro, JLR App, and configurator feature "
        "with real users before it reaches the engineering handoff stage."
    ),
    "goal_3_update": "On track",

    # --- Engagement overview ---
    "studies_run":        "42",
    "studies_delta":      "+14",
    "credits_used":       "320",
    "credits_total":      "400",
    "study_type_1":       "Usability testing",
    "study_type_1_count": "22 studies",
    "study_type_2":       "Concept testing",
    "study_type_2_count": "12 studies",
    "top_researcher_1":   "Marcus Chen: 38 studies",
    "top_researcher_2":   "Priya Sharma: 29 studies",
    "top_researcher_3":   "Tom Whitfield: 24 studies",

    # --- Monthly credit usage (Oct 2025 – Mar 2026) ---
    "months":  ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    "credits": [42, 51, 38, 67, 54, 68],

    # --- Achievements ---
    "goal_1_achieved": (
        "42 studies run in Q1 — average time from study launch to insights: 3.2 hours."
    ),
    "goal_1_impact": (
        "2.5 days saved per sprint. Product teams shipped 3 digital features ahead of schedule."
    ),
    "goal_2_achieved": (
        "12 PMs and designers ran studies independently. Research team involvement "
        "dropped from 100% to 40% of all studies."
    ),
    "goal_2_impact": (
        "Research volume up 45% quarter-on-quarter with no additional headcount."
    ),
    "goal_3_achieved": (
        "100% of Pivi Pro 2.0 features user-tested before engineering handoff."
    ),
    "goal_3_impact": (
        "4 critical UX issues caught pre-engineering. "
        "Estimated 6 sprint cycles of rework avoided."
    ),

    # --- Industry benchmarks ---
    "benchmark_studies_you":   "14.0",
    "benchmark_nonresearcher": "41%",
    "benchmark_decisions":     "71%",

    # --- Project spotlight ---
    "spotlight_title":          "Pivi Pro 2.0 Navigation Redesign",
    "spotlight_challenge_line": "Navigation overhaul at risk of shipping without user validation",
    "spotlight_challenge":      (
        "The Pivi Pro 2.0 navigation redesign was about to reach the 200-engineer "
        "handoff stage with no user validation. The UX team had a 4-day window before lock."
    ),
    "spotlight_what_line":      "3-day unmoderated study — 52 participants across UK, Germany & US",
    "spotlight_what":           (
        "Ran a 3-day unmoderated usability study with 52 participants testing two "
        "navigation variants. Studies launched and closed within the 4-day window using Maze panel."
    ),
    "spotlight_outcome_line":   "Critical issue caught — 6 sprint cycles of rework avoided",
    "spotlight_outcome":        (
        "Variant B showed 34% faster task completion. A critical back-navigation "
        "issue found in both variants was fixed before handoff. Estimated rework avoided: 6 sprints."
    ),

    # --- What's new (4 Maze features tied to JLR goals) ---
    "feature_1_title": "Maze AI Analysis",
    "feature_1_desc":  (
        "Automatically surfaces themes and key insights from your study results, "
        "cutting analysis time by up to 70% — ideal for JLR's fast sprint cycles."
    ),
    "feature_2_title": "Live Website Testing",
    "feature_2_desc":  (
        "Test the JLR configurator, JLR.com, and Pivi Pro web tools with real users "
        "directly — no prototype needed. Full navigation, deep qualitative feedback."
    ),
    "feature_3_title": "Participant Targeting",
    "feature_3_desc":  (
        "Reach JLR-relevant audiences with 130+ filters: luxury car owners, "
        "current JLR drivers, automotive enthusiasts across UK, Germany, US and China."
    ),
    "feature_4_title": "AI Moderator",
    "feature_4_desc":  (
        "Scale qualitative research with AI-powered interviews. "
        "Run 50 moderated sessions overnight — consistent, structured, researcher-grade insights."
    ),

    # --- Refreshed goals (Q2 2026) ---
    "next_goal_1_title": "AI-Powered Analysis",
    "next_goal_1_desc":  (
        "Every study uses Maze AI to auto-generate insight themes, "
        "cutting analysis time to under 1 hour per study."
    ),
    "next_goal_2_title": "Live Site Testing",
    "next_goal_2_desc":  (
        "Begin testing the JLR configurator and JLR.com with real target users "
        "via Live Website Testing — no prototype builds required."
    ),
    "next_goal_3_title": "Global Research Scale",
    "next_goal_3_desc":  (
        "Expand research coverage to US, Germany, and China markets "
        "using Maze panel targeting and AI Moderator for multilingual sessions."
    ),

    # --- Subscription details (mirrors Absa slide 6 structure) ---
    "subscription_period":  "February 1, 2026 – February 1, 2027",
    "plan_name":            "Enterprise Plan",
    "poc":                  "Marcus Chen, UX Research Lead",
    "exec_sponsor":         "Rachel Davies, VP Digital Products",

    # --- Reflecting on the quarter ---
    "working_well":         (
        "Cross-team adoption accelerating — templates have cut study setup time "
        "from 2 hours to under 20 minutes. Pivi Pro team now runs 3–4 studies per sprint."
    ),
    "getting_in_the_way":   (
        "SSO security review still in progress (IT & InfoSec). "
        "Some teams still defaulting to manual panel recruitment — onboarding needed."
    ),
}


# ─────────────────────────────────────────────────────────────────────────────
# PLACEHOLDER REPLACEMENTS
# Maps exact template text → JLR replacement
# ─────────────────────────────────────────────────────────────────────────────

def build_replacements(c: dict) -> list:
    """Return list of (find, replace) tuples ordered from most specific to least."""
    return [
        # ── Cover ──
        ("[Insert date]",                   c["date"]),
        ("[insert client logo]",            c["name"]),
        ("[insert quarter]",                c["quarter"]),
        ("NAME OF PRESENTATION",            f"{c['short']} <> Maze Partnership Review"),
        ("Quarterly Business Review",       f"Quarterly Business Review\n{c['name']}"),

        # ── Engagement overview subtitle ──
        ("In the last 3 months, your team ran",
         f"In Q1 2026, {c['short']}'s teams ran"),

        # ── Engagement stats ──
        ("[X]",                             c["studies_run"]),
        ("+ [x]",                           f"+{c['studies_delta']}"),
        ("[type 1]",                        c["study_type_1"]),
        ("[type 2]",                        c["study_type_2"]),
        ("[xx] studies",                    c["study_type_1_count"]),
        ("Credits used: xx/xx",             f"Credits used: {c['credits_used']}/{c['credits_total']}"),
        ("Xx: xx studies",                  c["top_researcher_1"]),
        ("Usage over the quarter",          "Credits used Oct–Mar"),
        ("Usage vs commitment: xx/xx",      f"Usage vs commitment: {c['credits_used']}/{c['credits_total']}"),
        ("[add active subtitle based on content below]",
         f"How {c['short']} used Maze in {c['quarter']}"),

        # ── Goals review — descriptions (unique enough to match safely) ──
        ("Run usability tests within 48 hours of a design being ready,",
         c["goal_1_desc"]),
        ("Enable product managers to run their own discovery studies w",
         c["goal_2_desc"]),
        ("Test every major feature with real users before it ships, so",
         c["goal_3_desc"]),

        # ── Goal status badges ──
        ("Update",                          c["goal_1_update"]),

        # ── Goal refresh (next quarter) — use longer unique substrings ──
        ("Speed\nRun usability tests",
         f"{c['next_goal_1_title']}\n{c['next_goal_1_desc']}"),
        ("Scale\nEnable product managers",
         f"{c['next_goal_2_title']}\n{c['next_goal_2_desc']}"),
        ("Impact\nTest every major feature",
         f"{c['next_goal_3_title']}\n{c['next_goal_3_desc']}"),

        # ── Achievements ──
        ("[Reminder of goal]",              c["goal_1_achieved"]),
        ("[e.g. 38 studies, avg. results in 4 hrs]",
         c["goal_1_achieved"]),
        ("[e.g.",                           ""),
        ("2 days",                          "2.5 days"),
        ("Saved per sprint - engineers unblocked faster]",
         c["goal_1_impact"]),
        ("[Xxx]",                           c["goal_2_achieved"]),

        # ── Reflecting on the quarter ──
        ("What's working well",             "What's working well"),
        ("Take notes here",                 c["working_well"]),
        ("What's getting in the way",       "What's getting in the way"),

        # ── Industry trends ──
        ("You -",                           "JLR —"),
        ("How active your team is relative to peers",
         f"JLR: {c['benchmark_studies_you']} studies/month vs 8.2 industry avg"),
        ("% of studies run by PMs, designers, engineers, etc",
         f"JLR: {c['benchmark_nonresearcher']} vs 28% industry avg"),
        ("% of product decisions backed by research",
         f"JLR: {c['benchmark_decisions']} vs 53% industry avg"),

        # ── What's new (features) ──
        ("Feature 1",                       c["feature_1_title"]),
        ("Feature 2",                       c["feature_2_title"]),
        ("Feature 3",                       c["feature_3_title"]),
        ("Feature 4",                       c["feature_4_title"]),
        ("[2-3 sentence description of the new feature/product]",
         c["feature_1_desc"]),

        # ── Project spotlight ──
        ("[One-line problem statement]",    c["spotlight_challenge_line"]),
        ("[2-3 sentences on what the team was trying to solve and what",
         c["spotlight_challenge"]),
        ("[One-line description of the study]",
         c["spotlight_what_line"]),
        ("[2-3 sentences on the approach — study type, who participate",
         c["spotlight_what"]),
        ("[One-line result]",               c["spotlight_outcome_line"]),
        ("[2-3 sentences on what they found, what decision it enabled,",
         c["spotlight_outcome"]),

        # ── Goal refresh slide subtitle ──
        ("What are your goals and strategic priorities for next quarte",
         f"Refreshed goals and strategic priorities for {c['short']} in Q2 2026"),

        # ── Cover/title variants (slides 19–22 cover color variants) ──
        ("Title goes here",                 f"{c['short']} <> Maze  ·  {c['quarter']}"),
        ("25 AUGUST 2025",                  c["date"].upper()),
        ("Gui",                             c["name"]),

        # ── Engagement section subtitle ──
        ("A snapshot of your most impactful project this quarter",
         f"How Maze helped {c['short']} deliver better digital experiences in Q1 2026"),

        # ── Goal sections headings ──
        ("These were your research goals going into this quarter",
         f"These were {c['short']}'s Maze goals going into Q1 2026"),
        ("How Maze helped drive your goals",
         f"How Maze helped {c['short']} deliver on its Q1 goals"),
        ("Tell us what your team is experiencing",
         f"What {c['short']}'s teams are experiencing with Maze"),
        ("How you compare to your peers",
         f"How {c['short']} compares to automotive & enterprise peers"),
        ("How Maze can keep supporting your goals",
         f"New Maze capabilities aligned to {c['short']}'s Q2 priorities"),
        ("What are your goals and strategic priorities for next quarte",
         f"Refreshed goals and priorities for {c['short']} in Q2 2026"),
    ]


# ─────────────────────────────────────────────────────────────────────────────
# AI REVIEW BANNER — stamp every slide after template copy
# ─────────────────────────────────────────────────────────────────────────────

def stamp_ai_review_banners(slides_svc, pres_id: str):
    """Add the orange 'Edited with AI' banner to every slide in the presentation."""
    print("  Stamping AI review banners on all slides...")
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    reqs = []
    for slide in pres.get("slides", []):
        reqs.extend(api.add_ai_review_banner(slide["objectId"]))
    for i in range(0, len(reqs), 50):
        api.batch_update(slides_svc, pres_id, reqs[i:i+50])
    print(f"  ✅ Banners added to {len(pres.get('slides', []))} slides")


# ─────────────────────────────────────────────────────────────────────────────
# CHART — Credits used Oct–Mar
# ─────────────────────────────────────────────────────────────────────────────

def inject_usage_chart(slides_svc, sheets_svc, pres_id: str, c: dict):
    print("  Creating usage chart in Google Sheets...")
    ss_id, chart_id = sc.create_bar_chart(
        sheets_svc,
        categories=c["months"],
        values=c["credits"],
        title=f"Credits used — {c['quarter']}",
        color="#A366FF",
    )

    # Find slide containing the engagement overview
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    target_slide = None
    for slide in pres.get("slides", []):
        for el in slide.get("pageElements", []):
            for te in el.get("shape", {}).get("text", {}).get("textElements", []):
                content = te.get("textRun", {}).get("content", "")
                if "Credits used Oct" in content or "Usage over the quarter" in content:
                    target_slide = slide["objectId"]
                    break
            if target_slide:
                break
        if target_slide:
            break

    if not target_slide:
        # Fall back to slide 5 (engagement overview) by position
        if len(pres["slides"]) >= 5:
            target_slide = pres["slides"][4]["objectId"]

    if not target_slide:
        print("  ⚠️  Could not find engagement slide, skipping chart")
        return None

    print(f"  Embedding chart into slide {target_slide}...")
    req = api.add_sheets_chart(
        api.new_id(), target_slide,
        ss_id, chart_id,
        x=5.18, y=1.55, w=4.55, h=2.0,
    )
    api.batch_update(slides_svc, pres_id, [req])
    print("  ✅ Chart embedded")
    return ss_id


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("Authenticating with Google...")
    slides_svc, drive_svc, sheets_svc = get_services()

    # 1. Copy the template directly into the shared output folder
    deck_title = f"{CUSTOMER['short']} <> Maze Partnership Review {CUSTOMER['quarter']}"
    print(f"Copying template → '{deck_title}'...")
    pres_id = copy_template(drive_svc, TEMPLATE_ID, deck_title)
    print(f"  Created: https://docs.google.com/presentation/d/{pres_id}/edit")
    print(f"  Folder:  https://drive.google.com/drive/folders/{OUTPUT_FOLDER_ID}")

    try:
        # 2. Apply all text replacements
        print("Filling placeholders...")
        replacements = build_replacements(CUSTOMER)
        reqs = []
        for find, replace in replacements:
            if find.strip():
                reqs.append({
                    "replaceAllText": {
                        "containsText": {"text": find, "matchCase": False},
                        "replaceText": replace,
                    }
                })

        # Send in batches of 50
        total = 0
        for i in range(0, len(reqs), 50):
            batch = reqs[i:i+50]
            result = api.batch_update(slides_svc, pres_id, batch)
            for reply in result.get("replies", []):
                total += reply.get("replaceAllText", {}).get("occurrencesChanged", 0)
        print(f"  {total} text occurrences updated across {len(reqs)} replacement rules")

        # 3. Inject usage chart
        inject_usage_chart(slides_svc, sheets_svc, pres_id, CUSTOMER)

        # 4. Stamp AI review banners on every slide
        stamp_ai_review_banners(slides_svc, pres_id)

        # 5. Done
        url = f"https://docs.google.com/presentation/d/{pres_id}/edit"
        print(f"\n✅  {deck_title}")
        print(f"    {url}\n")
        return url

    except Exception as e:
        print(f"\n❌  Build failed: {e}")
        print("Cleaning up...")
        try:
            drive_svc.files().delete(fileId=pres_id).execute()
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
