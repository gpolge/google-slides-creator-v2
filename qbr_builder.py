"""
QBR Builder — copies the Maze QBR template and populates it with customer data.

Strategy:
  1. Copy the template presentation (preserves all master formatting, images, brand)
  2. replaceAllText to fill every [placeholder]
  3. Inject a Sheets usage chart into the engagement overview slide
  4. Move the finished deck to the persistent "AI Slides" presentation (append slides)

Usage:
  python3 qbr_builder.py
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
# CUSTOMER DATA  —  edit this section for each account
# ─────────────────────────────────────────────────────────────────────────────

CUSTOMER = {
    "name":       "Acme Corp",
    "date":       "16 April 2026",
    "quarter":    "Q1 2026",

    # Goals
    "goal_1_title": "Speed",
    "goal_1_desc":  "Run usability tests within 48 hours of a design being ready, cutting the feedback loop on every sprint.",
    "goal_1_status": "On track",

    "goal_2_title": "Scale",
    "goal_2_desc":  "Enable product managers to run their own discovery studies without research team bottlenecks.",
    "goal_2_status": "In progress",

    "goal_3_title": "Impact",
    "goal_3_desc":  "Test every major feature with real users before it ships, so no feature launches blind.",
    "goal_3_status": "Needs focus",

    # Engagement overview
    "studies_run":        "38",
    "studies_delta":      "+12",
    "credits_used":       "210",
    "credits_total":      "250",
    "study_type_1":       "Usability testing",
    "study_type_1_count": "22 studies",
    "study_type_2":       "Concept testing",
    "study_type_2_count": "16 studies",
    "top_researcher_1":   "Sarah M.: 18 studies",
    "top_researcher_2":   "James T.: 12 studies",
    "top_researcher_3":   "Priya K.: 8 studies",

    # Monthly usage data for chart
    "months":  ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    "credits": [28, 34, 42, 51, 38, 44],

    # Achievements
    "goal_1_achieved": "Ran 38 studies this quarter — avg. results delivered in under 4 hours.",
    "goal_1_impact":   "2 days saved per sprint. Engineers unblocked faster, reducing re-work cycles.",
    "goal_2_achieved": "15 PMs ran their own studies using templates. 0 research-team bottlenecks this quarter.",
    "goal_2_impact":   "Research volume up 30% without adding headcount.",
    "goal_3_achieved": "100% of major launches were user-tested before shipping.",
    "goal_3_impact":   "3 critical UX issues caught pre-launch, avoiding post-ship rework estimated at 3 sprints.",

    # Industry benchmarks
    "benchmark_studies_you":     "12.7",
    "benchmark_nonresearcher":   "41%",
    "benchmark_decisions":       "67%",

    # What's new (4 features)
    "feature_1_title": "Maze AI",
    "feature_1_desc":  "Automatically surfaces key themes and insights from your study results, cutting analysis time by up to 70%.",
    "feature_2_title": "Live Website Testing",
    "feature_2_desc":  "Test your live product with real users — no prototypes needed. Works on any web URL.",
    "feature_3_title": "Participant Targeting",
    "feature_3_desc":  "Reach the exact audience you need with 130+ demographic filters across the Maze panel.",
    "feature_4_title": "Jira Integration",
    "feature_4_desc":  "Link Maze insights directly to Jira tickets so every decision has a research trail.",

    # Refreshed goals (next quarter)
    "next_goal_1_title": "Democratise Research",
    "next_goal_1_desc":  "Every product team runs at least 2 studies per sprint without research team involvement.",
    "next_goal_2_title": "Faster Insights",
    "next_goal_2_desc":  "Average time from study launch to insight delivery under 48 hours.",
    "next_goal_3_title": "Research ROI",
    "next_goal_3_desc":  "Quantify and report the business impact of research on at least 3 shipped features.",

    # Project spotlight
    "spotlight_challenge":  "The payments redesign was about to ship without user validation. The team had 4 days.",
    "spotlight_challenge_title": "Payments redesign at risk of shipping without user testing",
    "spotlight_what_we_did": "Ran a 3-day unmoderated usability test with 40 participants via Maze. Tested 3 prototype variants.",
    "spotlight_what_title":  "3-day unmoderated usability study, 40 participants",
    "spotlight_outcome":    "Identified a critical navigation issue missed in design review. Fix shipped pre-launch. NPS on payments flow +18 pts post-launch.",
    "spotlight_outcome_title": "Critical issue caught — NPS +18 pts post-launch",
}


# ─────────────────────────────────────────────────────────────────────────────
# PLACEHOLDER MAP
# Maps template [placeholder text] → replacement value
# ─────────────────────────────────────────────────────────────────────────────

def build_replacements(c: dict) -> dict:
    return {
        "[Insert date]":                    c["date"],
        "[insert client logo]":             c["name"],
        "[insert quarter]":                 c["quarter"],
        "NAME OF PRESENTATION":             f"{c['name']} QBR {c['quarter']}",
        "Quarterly Business Review":        f"Quarterly Business Review\n{c['name']}",

        # Goals slide
        "Speed\nRun usability tests within 48 hours of a design being ready,":
            f"{c['goal_1_title']}\n{c['goal_1_desc']}",
        "Scale\nEnable product managers to run their own discovery studies w":
            f"{c['goal_2_title']}\n{c['goal_2_desc']}",
        "Impact\nTest every major feature with real users before it ships, so":
            f"{c['goal_3_title']}\n{c['goal_3_desc']}",
        "Update":                           c["goal_1_status"],   # first occurrence = goal 1

        # Engagement overview stats
        "[X]":                              c["studies_run"],
        "+ [x]":                            f"+{c['studies_delta']}",
        "[type 1]":                         c["study_type_1"],
        "[type 2]":                         c["study_type_2"],
        "[xx] studies":                     c["study_type_1_count"],
        "Credits used: xx/xx":              f"Credits used: {c['credits_used']}/{c['credits_total']}",
        "Xx: xx studies":                   c["top_researcher_1"],
        "Usage vs commitment: xx/xx":       f"Usage vs commitment: {c['credits_used']}/{c['credits_total']}",

        # Achievements
        "[Reminder of goal]":               c["goal_1_achieved"],   # first card
        "[e.g. 38 studies, avg. results in 4 hrs]": c["goal_1_achieved"],
        "[e.g.":                            "",
        "2 days":                           "2 days",
        "Saved per sprint - engineers unblocked faster]": c["goal_1_impact"],
        "[Xxx]":                            c["goal_2_achieved"],   # will replace all remaining

        # Reflecting (leave as note-taking space — no replacements needed)

        # Industry trends
        "You -":                            f"You — {c['benchmark_studies_you']}",

        # What's new
        "Feature 1":                        c["feature_1_title"],
        "Feature 2":                        c["feature_2_title"],
        "Feature 3":                        c["feature_3_title"],
        "Feature 4":                        c["feature_4_title"],
        "[2-3 sentence description of the new feature/product]": c["feature_1_desc"],

        # Goal refresh
        "[goal title]":                     c["next_goal_1_title"],
        "[1-2 sentences description]":      c["next_goal_1_desc"],
        "I[goal title]":                    c["next_goal_3_title"],

        # Spotlight
        "THE CHALLENGE":                    "THE CHALLENGE",
        "[One-line problem statement]":     c["spotlight_challenge_title"],
        "[2-3 sentences on what the team was trying to solve and what": c["spotlight_challenge"],
        "WHAT WE DID":                      "WHAT WE DID",
        "[One-line description of the study]":  c["spotlight_what_title"],
        "[2-3 sentences on the approach — study type, who participate": c["spotlight_what_we_did"],
        "OUTCOME":                          "OUTCOME",
        "[One-line result]":                c["spotlight_outcome_title"],
        "[2-3 sentences on what they found, what decision it enabled,": c["spotlight_outcome"],

        # Agenda / section labels (leave section labels as-is, they look good)
        "Goal review":                      "Goal review",
        "[add active subtitle based on content below]":
            f"How {c['name']} used Maze in {c['quarter']}",

        # Footer / title area
        "Gui":                              c["name"],
        "Title goes here":                  f"{c['name']} QBR {c['quarter']}",
        "25 AUGUST 2025":                   c["date"].upper(),
    }


# ─────────────────────────────────────────────────────────────────────────────
# CHART INJECTION — Usage over the quarter
# ─────────────────────────────────────────────────────────────────────────────

def find_slide_with_text(slides_svc, pres_id: str, search_text: str):
    """Return the objectId of the first slide containing search_text."""
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    for slide in pres.get("slides", []):
        for el in slide.get("pageElements", []):
            for te in el.get("shape", {}).get("text", {}).get("textElements", []):
                if search_text in te.get("textRun", {}).get("content", ""):
                    return slide["objectId"]
    return None


def inject_usage_chart(slides_svc, sheets_svc, pres_id: str, customer: dict):
    """Create a Sheets chart and embed it in the engagement overview slide."""
    print("  Creating usage chart in Google Sheets...")
    ss_id, chart_id = sc.create_bar_chart(
        sheets_svc,
        categories=customer["months"],
        values=customer["credits"],
        title=f"Credits used — {customer['quarter']}",
        color="#A366FF",
    )

    # Find the engagement overview slide (contains "Usage over the quarter")
    slide_id = find_slide_with_text(slides_svc, pres_id, "Usage over the quarter")
    if not slide_id:
        print("  ⚠️  Could not find engagement slide, skipping chart injection")
        return

    print(f"  Injecting chart into slide {slide_id}...")
    # Position: right column of the engagement overview card area
    # Template card is at x≈5.03, y≈1.04 (the "Usage over the quarter" card)
    req = api.add_sheets_chart(
        api.new_id(), slide_id,
        ss_id, chart_id,
        x=5.18, y=1.55,
        w=4.55, h=2.0,
    )
    api.batch_update(slides_svc, pres_id, [req])
    print("  ✅ Chart injected")
    return ss_id  # return so we can clean up later if needed


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("Authenticating with Google...")
    slides_svc, drive_svc, sheets_svc = get_services()

    # 1. Copy the template directly into the shared output folder
    deck_title = f"{CUSTOMER['name']} QBR {CUSTOMER['quarter']}"
    print(f"Copying template into shared folder...")
    pres_id = copy_template(drive_svc, TEMPLATE_ID, deck_title)
    print(f"  Created: https://docs.google.com/presentation/d/{pres_id}/edit")
    print(f"  Folder:  https://drive.google.com/drive/folders/{OUTPUT_FOLDER_ID}")

    try:
        # 2. Replace all text placeholders
        print("Filling placeholders...")
        replacements = build_replacements(CUSTOMER)
        reqs = []
        for find, replace in replacements.items():
            reqs.append({
                "replaceAllText": {
                    "containsText": {"text": find, "matchCase": False},
                    "replaceText": replace,
                }
            })
        # Send in batches of 50 to avoid payload limits
        for i in range(0, len(reqs), 50):
            api.batch_update(slides_svc, pres_id, reqs[i:i+50])
        print(f"  Applied {len(reqs)} text replacements")

        # 3. Inject usage chart
        inject_usage_chart(slides_svc, sheets_svc, pres_id, CUSTOMER)

        # 4. Done
        url = f"https://docs.google.com/presentation/d/{pres_id}/edit"
        print(f"\n✅  QBR deck ready: {url}\n")
        return url

    except Exception as e:
        print(f"\n❌  Build failed: {e}")
        print("Cleaning up...")
        drive_svc.files().delete(fileId=pres_id).execute()
        raise


if __name__ == "__main__":
    main()
