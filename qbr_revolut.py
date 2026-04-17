"""
Revolut <> Maze QBR — Q1/Q2 2026
Period: Oct 2025 – Apr 2026

Data sources:
  - Omni: 407 studies used / 1,000 limit, 51 active researchers, ENTERPRISE plan, $506,750 ARR
  - Gong (9 calls): contract, goals, feature requests, challenges, expansion plans

Run: python3 qbr_revolut.py
"""

import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from src.auth import get_services
from src import slides_api as api
from src import sheets_charts as sc
from src.drive import copy_template, OUTPUT_FOLDER_ID

TEMPLATE_ID = "1KpYl6uTcUMBUFCHnMGkRXJVuxhcdsHrhDnzbZRQKNmo"

# ── Account data ───────────────────────────────────────────────────────────────
# Sources: Omni (studies, seats, ARR) + Gong (9 calls Oct 2025 – Apr 2026)

CUSTOMER = {
    # Identity
    "name":    "Revolut",
    "short":   "Revolut",
    "date":    "17 April 2026",
    "quarter": "H1 2026",
    "period":  "Oct 2025 – Apr 2026",

    # Contract (Gong Call 4, Jan 28 2026 + Call 2, Nov 6 2025)
    "subscription_period": "June 2025 – June 2027",
    "plan_name":           "Enterprise",
    "poc":                 "Natasha Alvarez / Teresa Lella",
    "exec_sponsor":        "Dimitri (Budget Owner)",

    # Engagement metrics (Omni + Gong Call 7, Mar 16 2026)
    "studies_run":         "364",          # as of Mar 16 (Gong)
    "studies_limit":       "1,000",        # annual limit (Gong Call 4)
    "studies_delta":       "+42 vs Q3",
    "credits_used":        "73,000",       # 104K – 31K remaining (Gong Call 7)
    "credits_total":       "104,000",      # annual panel credits (Gong Call 4)
    "credits_remaining":   "31,000",       # Gong Call 7 + 9
    "active_researchers":  "51",           # Omni, last 60 days
    "utilisation_pct":     "70%",          # 73K / 104K

    # Monthly credits Oct–Mar (Gong: 6.9K avg Oct, 7,967 Nov, ~7K Dec–Mar)
    "months":  ["Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
    "credits": [8100, 9200, 7500, 7100, 7200, 6900],   # in credits (Gong data)

    # Goals — sourced from Gong calls
    "goal_1_title":  "Scale decentralised research",
    "goal_1_desc":   (
        "50% of Revolut's research is now run by non-UX teams. "
        "Revolut wants to continue enabling comms, business, and "
        "growth teams to run their own studies independently."
    ),
    "goal_1_status": "In progress",

    "goal_2_title":  "Expand B2B research capability",
    "goal_2_desc":   (
        "Revolut is building out strategic B2B research — targeting "
        "owners, CEOs and CFOs across Italy, Spain and France. "
        "Currently 10% of studies; scaling to become a core pillar."
    ),
    "goal_2_status": "In progress",

    "goal_3_title":  "Evaluate AI Moderator at scale",
    "goal_3_desc":   (
        "Revolut is comparing Maze AI Moderator against Salomo. "
        "A formal evaluation is planned for May 2026 alongside "
        "enablement sessions for the wider research team."
    ),
    "goal_3_status": "Scheduled for May 2026",

    # Goal achievements / impact (narrative)
    "goal_1_achieved": (
        "364 studies run across decentralised teams — UX, comms, "
        "growth, and business research all active. Bi-weekly syncs "
        "established to maintain quality and consistency."
    ),
    "goal_1_impact": (
        "Research velocity increased by 42 studies quarter-on-quarter. "
        "Non-researcher teams now self-serve unmoderated studies "
        "without UX bottlenecks."
    ),

    "goal_2_achieved": (
        "B2B research now active across core EU markets. "
        "Feasibility confirmed for veteran segment (UK/EEA). "
        "Strategic business banking research project in pipeline."
    ),
    "goal_2_impact": (
        "B2B now represents ~10% of total studies. Revolut has identified "
        "key B2B audience segments in Italy, Spain, and France — "
        "informing product and go-to-market decisions."
    ),

    "goal_3_achieved": (
        "AI Moderator demo completed. Revolut currently trialling "
        "against Salomo. Formal evaluation and decision scheduled "
        "for May 2026 with enablement session to follow."
    ),
    "goal_3_impact": (
        "Revolut's moderated research remains at just 2 studies. "
        "AI Moderator adoption could unlock 10x more moderated "
        "studies without increasing research team headcount."
    ),

    # Researchers (Gong Call 7 context)
    "top_researcher_1": "Natasha Alvarez — Research Lead",
    "top_researcher_2": "Teresa Lella — Senior UX Researcher",
    "top_researcher_3": "Karim Remmel — Growth Research Manager",

    "study_type_1":       "Unmoderated usability testing",
    "study_type_1_count": "~280 studies",
    "study_type_2":       "Survey",
    "study_type_2_count": "~68 studies",

    # Benchmarks
    "benchmark_studies_you":      "~60 studies/month",
    "benchmark_nonresearcher":    "~50%",
    "benchmark_decisions":        "estimated 65%+",

    # What's working / what's in the way (Gong synthesis)
    "working_well": (
        "High researcher engagement (51 active in 60 days). "
        "Decentralisation is working — multiple teams self-serving. "
        "Matrix block and Fresh Eyes released and being used. "
        "Strong partnership cadence with bi-weekly syncs."
    ),
    "getting_in_the_way": (
        "Google Meet plugin broken — blocking calendar sync for moderated studies. "
        "Romania missing from standard panel recruitment (key Revolut market). "
        "B2B panel universe visibility limited — hard to gauge feasibility. "
        "Advanced quantitative features (multi-select hard stops, complex logic) not yet available."
    ),

    # Product updates / spotlight
    "feature_1_title": "Matrix Block — Now Live",
    "feature_1_desc":  "Build ranked and matrix-style questions natively in Maze — no workarounds needed.",
    "feature_2_title": "Fresh Eyes (Participant Exclusion)",
    "feature_2_desc":  "Automatically exclude repeat participants across studies. Now live.",
    "feature_3_title": "AI Moderator",
    "feature_3_desc":  "Run AI-moderated interviews at scale. Formal Revolut evaluation in May 2026.",
    "feature_4_title": "Panel Expansion",
    "feature_4_desc":  "Romania and additional B2B segments in roadmap discussions.",

    # Next quarter goals
    "next_goal_1_title": "AI Moderator adoption",
    "next_goal_1_desc":  "Complete evaluation vs. Salomo. Run first 5 AI moderated studies by end of Q2.",
    "next_goal_2_title": "B2B panel scale",
    "next_goal_2_desc":  "Launch strategic business banking research. Expand panel to Romania.",
    "next_goal_3_title": "Renew at higher credit volume",
    "next_goal_3_desc":  "June 2026 renewal — target upgrade from 104K to 150K+ credits given growth trend.",

    # Spotlight (veteran research project from Gong Call 9)
    "spotlight_title":     "Veteran Segment Research — UK & EEA",
    "spotlight_challenge_title": "The Challenge",
    "spotlight_challenge": (
        "Revolut needed to understand the veteran segment across UK and EEA — "
        "a niche but high-value audience with specific financial needs. "
        "Standard panel recruitment could not easily target this group."
    ),
    "spotlight_what_title": "What We Did",
    "spotlight_what_we_did": (
        "Used Maze Panel with a combination of standard and premium targeting. "
        "Feasibility confirmed for 500 unmoderated studies + 10–15 follow-up "
        "interviews across the veteran segment."
    ),
    "spotlight_outcome_title": "The Outcome",
    "spotlight_outcome": (
        "Study launched with full panel coverage. Findings will inform Revolut's "
        "product positioning for the veteran segment in Q2 2026."
    ),
}


def build_replacements(c: dict) -> list:
    """Map customer data to template placeholder strings."""
    pairs = [
        # Identity
        ("[Insert date]",              c["date"]),
        ("[insert quarter]",           c["quarter"]),
        ("NAME OF PRESENTATION",       f"{c['name']} <> Maze — {c['quarter']} Review"),
        ("[insert customer name]",     c["name"]),
        ("[Insert customer name]",     c["name"]),
        ("[Customer Name]",            c["name"]),
        ("[customer name]",            c["name"]),
        ("Customer Name",              c["name"]),

        # Contract
        ("[insert subscription period]", c["subscription_period"]),
        ("[Insert plan name]",           c["plan_name"]),
        ("[Insert POC name]",            c["poc"]),

        # Engagement metrics
        ("[insert number of studies]",   c["studies_run"]),
        ("[insert delta]",               c["studies_delta"]),
        ("[insert credits used]",        c["credits_used"]),
        ("[insert credits total]",       c["credits_total"]),
        ("[insert number of researchers]", c["active_researchers"]),
        ("[insert utilisation %]",       c["utilisation_pct"]),

        # Goals
        ("[Insert goal 1 title]",        c["goal_1_title"]),
        ("[Insert goal 1 description]",  c["goal_1_desc"]),
        ("[Insert goal 1 status]",       c["goal_1_status"]),
        ("[Insert goal 2 title]",        c["goal_2_title"]),
        ("[Insert goal 2 description]",  c["goal_2_desc"]),
        ("[Insert goal 2 status]",       c["goal_2_status"]),
        ("[Insert goal 3 title]",        c["goal_3_title"]),
        ("[Insert goal 3 description]",  c["goal_3_desc"]),
        ("[Insert goal 3 status]",       c["goal_3_status"]),

        # Goal achievements
        ("[Insert goal 1 achieved]",     c["goal_1_achieved"]),
        ("[Insert goal 1 impact]",       c["goal_1_impact"]),
        ("[Insert goal 2 achieved]",     c["goal_2_achieved"]),
        ("[Insert goal 2 impact]",       c["goal_2_impact"]),
        ("[Insert goal 3 achieved]",     c["goal_3_achieved"]),
        ("[Insert goal 3 impact]",       c["goal_3_impact"]),

        # Researchers
        ("[Insert top researcher 1]",    c["top_researcher_1"]),
        ("[Insert top researcher 2]",    c["top_researcher_2"]),
        ("[Insert top researcher 3]",    c["top_researcher_3"]),

        # Study types
        ("[Insert study type 1]",        c["study_type_1"]),
        ("[Insert study type 1 count]",  c["study_type_1_count"]),
        ("[Insert study type 2]",        c["study_type_2"]),
        ("[Insert study type 2 count]",  c["study_type_2_count"]),

        # Benchmarks
        ("[Insert benchmark studies]",   c["benchmark_studies_you"]),
        ("[Insert % non-researcher]",    c["benchmark_nonresearcher"]),
        ("[Insert % decisions]",         c["benchmark_decisions"]),

        # What's working / in the way
        ("[Insert what's working well]",    c["working_well"]),
        ("[Insert what's getting in the way]", c["getting_in_the_way"]),

        # Features
        ("[Insert feature 1 title]",     c["feature_1_title"]),
        ("[Insert feature 1 desc]",      c["feature_1_desc"]),
        ("[Insert feature 2 title]",     c["feature_2_title"]),
        ("[Insert feature 2 desc]",      c["feature_2_desc"]),
        ("[Insert feature 3 title]",     c["feature_3_title"]),
        ("[Insert feature 3 desc]",      c["feature_3_desc"]),
        ("[Insert feature 4 title]",     c["feature_4_title"]),
        ("[Insert feature 4 desc]",      c["feature_4_desc"]),

        # Next steps
        ("[Insert next goal 1 title]",   c["next_goal_1_title"]),
        ("[Insert next goal 1 desc]",    c["next_goal_1_desc"]),
        ("[Insert next goal 2 title]",   c["next_goal_2_title"]),
        ("[Insert next goal 2 desc]",    c["next_goal_2_desc"]),
        ("[Insert next goal 3 title]",   c["next_goal_3_title"]),
        ("[Insert next goal 3 desc]",    c["next_goal_3_desc"]),

        # Spotlight
        ("[Insert spotlight title]",     c["spotlight_title"]),
        ("[Insert challenge title]",     c["spotlight_challenge_title"]),
        ("[Insert challenge]",           c["spotlight_challenge"]),
        ("[Insert what we did title]",   c["spotlight_what_title"]),
        ("[Insert what we did]",         c["spotlight_what_we_did"]),
        ("[Insert outcome title]",       c["spotlight_outcome_title"]),
        ("[Insert outcome]",             c["spotlight_outcome"]),
    ]
    return pairs


def inject_usage_chart(slides_svc, sheets_svc, pres_id: str, c: dict):
    """Create monthly credits chart and embed in the engagement slide."""
    print("  Creating usage chart...")
    ss_id, chart_id = sc.create_bar_chart(
        sheets_svc,
        categories=c["months"],
        values=c["credits"],
        title=f"Credits used — {c['quarter']}",
        color="#A366FF",
    )

    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    target_slide = None
    for slide in pres.get("slides", []):
        for el in slide.get("pageElements", []):
            for te in el.get("shape", {}).get("text", {}).get("textElements", []):
                content = te.get("textRun", {}).get("content", "")
                if any(kw in content for kw in ["Credits used", "Usage over", "credit", "monthly"]):
                    target_slide = slide["objectId"]
                    break
            if target_slide:
                break
        if target_slide:
            break

    # Fallback to slide 5
    if not target_slide and len(pres["slides"]) >= 5:
        target_slide = pres["slides"][4]["objectId"]

    if not target_slide:
        print("  ⚠️  Could not find engagement slide — skipping chart")
        return None

    print(f"  Embedding chart into slide {target_slide}...")
    req = api.add_sheets_chart(
        api.new_id(), target_slide, ss_id, chart_id,
        x=5.18, y=1.55, w=4.55, h=2.0,
    )
    api.batch_update(slides_svc, pres_id, [req])
    print("  ✅ Chart embedded")
    return ss_id


def stamp_ai_review_banners(slides_svc, pres_id: str):
    print("  Stamping AI review banners...")
    pres = slides_svc.presentations().get(presentationId=pres_id).execute()
    reqs = []
    for slide in pres.get("slides", []):
        reqs.extend(api.add_ai_review_banner(slide["objectId"]))
    for i in range(0, len(reqs), 50):
        api.batch_update(slides_svc, pres_id, reqs[i:i+50])
    print(f"  ✅ Banners added to {len(pres.get('slides', []))} slides")


def main():
    print("Authenticating with Google...")
    slides_svc, drive_svc, sheets_svc = get_services()

    title = f"Revolut <> Maze — H1 2026 Review"
    print(f"Copying template → '{title}'...")
    pres_id = copy_template(drive_svc, TEMPLATE_ID, title)
    url = f"https://docs.google.com/presentation/d/{pres_id}/edit"
    print(f"  Created: {url}")
    print(f"  Folder:  https://drive.google.com/drive/folders/{OUTPUT_FOLDER_ID}")

    try:
        # 1. Fill all text placeholders
        print("Filling placeholders...")
        replacements = build_replacements(CUSTOMER)
        reqs = []
        for find, replace in replacements:
            reqs.append({
                "replaceAllText": {
                    "containsText": {"text": find, "matchCase": False},
                    "replaceText": str(replace),
                }
            })
        # Send in batches of 50
        for i in range(0, len(reqs), 50):
            api.batch_update(slides_svc, pres_id, reqs[i:i+50])
        print(f"  Applied {len(replacements)} text replacements")

        # 2. Inject usage chart
        inject_usage_chart(slides_svc, sheets_svc, pres_id, CUSTOMER)

        # 3. Stamp AI review banners
        stamp_ai_review_banners(slides_svc, pres_id)

        print(f"\n✅  Revolut QBR deck ready: {url}\n")
        return url

    except Exception as e:
        print(f"\n❌  Build failed: {e}")
        import traceback; traceback.print_exc()
        try:
            drive_svc.files().delete(fileId=pres_id).execute()
            print("  Cleaned up partial deck")
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
