#!/usr/bin/env python3
"""
enrich_2108.py — Enrich OWP_2108_JCR_Cortex_v2.xlsx with key players + populate
sparse tabs.  Pattern mirrors enrich_2020.py but adapted to the 17-numbered-tab
schema used by live projects.

Inputs:
  • 2108_data.json            — Sage JDR parse
  • 2108_dashboard_arrays.json — 16 arrays mirroring index.html PROJECTS['2108']

Outputs (in-place):
  • OWP_2108_JCR_Cortex_v2.xlsx — Job Info tab rebuilt with sectioned key-players
    + project metadata; Overview tab gains a Project Team block; status text
    corrected from the cancelled-template default to "ON HOLD".

Project context (2026-04-27):
  R&G Apartments · Braseth Construction · 263 units · ON HOLD (Summer 2026?)
  $91,402 design/takeoff billed (the only revenue so far)
  No GDrive folder available; team data comes from index.html PROJECT_TEAMS['2108']
  + the sparse JDR identity block.
"""
import json
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).parent
JSON_FILE  = SCRIPT_DIR / "2108_data.json"
DASH_FILE  = SCRIPT_DIR / "2108_dashboard_arrays.json"
WORKBOOK   = SCRIPT_DIR / "OWP_2108_JCR_Cortex_v2.xlsx"

# ============================================================================
# Project-team metadata
# ============================================================================
META_2108 = {
    "job_id":          "2108",
    "name":            "R&G Apartments",
    "short_name":      "R&G",
    "gc":              "Braseth Construction",
    "gc_pm":           "TBD (Braseth) — primary contact pending",
    "gc_sup":          "TBD",
    "gc_pe":           "TBD",
    "owp_ri_foreman":  "TBD · project on hold",
    "owp_trim_foreman":"TBD",
    "owp_signatory":   "Richard Donelson",
    "owp_estimator":   "Jeffrey S. Gerard / Jordan E. Gerard (takeoff)",
    "owner":           "TBD",
    "developer":       "TBD",
    "architect":       "TBD",
    "structural":      "TBD",
    "mep_engineer":    "TBD (Franklin Engineering involved per AP — design fees $90.6k)",
    "civil":           "TBD",
    "landscape":       "TBD",
    "site_address":    "TBD",
    "permit":          "Pre-permit",
    "insurance":       "Standard (COI) — non-Wrap",
    "lien_position":   "Not yet recorded — no AR billed",
    "warranty":        "Pending project activation",
    "delivery_route":  "TBD",
    "contract_doc":    "Not yet executed — Braseth has not committed start date",
    "gdrive_status":   "No GDrive folder created yet (project hasn't reached active stage). "
                        "All team data from index.html PROJECT_TEAMS['2108'] + Sage JDR identity block.",
    "status_text":     "ON HOLD · Summer 2026?",
    "status_severity": "WATCH",
    "ar_billed":       91402,
    "design_cost":     117161.97,
    "carrying_loss":   25759.97,
    "units":           263,
    "fixtures":        "TBD (no fixture schedule yet)",
    "project_type":    "Multi-family apartments (pre-construction · on-hold)",
    "schedule":        "Design/takeoff complete · awaiting Braseth go-ahead · est. start Summer 2026 / Q1 2027",
    "expected_start":  "Q1 2027 (estimated)",
    "expected_finish": "TBD",
}

# ============================================================================
# Style helpers
# ============================================================================
INK         = Font(name="Arial", size=10, color="1F1F1F")
INK_BOLD    = Font(name="Arial", size=10, color="1F1F1F", bold=True)
SECTION_HDR = Font(name="Arial", size=11, color="1F1F1F", bold=True)
TITLE_FONT  = Font(name="Arial", size=14, color="1F1F1F", bold=True)
GREY_FONT   = Font(name="Arial", size=9, color="6B6B6B", italic=True)
WHITE_BOLD  = Font(name="Arial", size=10, color="FFFFFF", bold=True)
WARN_FONT   = Font(name="Arial", size=10, color="9A4F02", bold=True)

HDR_FILL    = PatternFill("solid", start_color="2C3E50")
SECTION_FILL= PatternFill("solid", start_color="ECF0F1")
ROW_ALT     = PatternFill("solid", start_color="FAFAFA")
TEAM_FILL   = PatternFill("solid", start_color="FFF8E7")
WARN_FILL   = PatternFill("solid", start_color="FFF3CD")
THIN        = Side(border_style="thin", color="D5D8DC")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP        = Alignment(wrap_text=True, vertical="top")
CENTER      = Alignment(horizontal="center", vertical="center")


def clear_range(ws, r1, r2, c1=1, c2=10):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).value = None
            ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)
            ws.cell(row=r, column=c).border = Border()


# ============================================================================
# JOB INFO tab — full rebuild with sectioned layout
# ============================================================================
def rebuild_job_info(ws):
    # Wipe whatever's there
    clear_range(ws, 1, ws.max_row + 5, 1, 6)

    ws["A1"] = f"JOB #{META_2108['job_id']} · PROJECT INFORMATION & KEY PLAYERS"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")
    ws["A2"] = ("Pre-construction record — project is on-hold. Team data sourced from "
                "index.html PROJECT_TEAMS['2108'] + Sage JDR identity block. "
                "GDrive folder not created yet (will populate when Braseth confirms start).")
    ws["A2"].font = GREY_FONT
    ws.merge_cells("A2:F2")

    sections = [
        ("IDENTITY", [
            ("Job Number",    META_2108["job_id"]),
            ("Job Name",      META_2108["name"]),
            ("Project Type",  META_2108["project_type"]),
            ("Site Address",  META_2108["site_address"]),
            ("Permit",        META_2108["permit"]),
            ("Status",        META_2108["status_text"]),
        ]),
        ("SCHEDULE", [
            ("Design / Takeoff",  "Complete (Apr 2026)"),
            ("Project Start",     META_2108["expected_start"]),
            ("Project End",       META_2108["expected_finish"]),
            ("Schedule Note",     META_2108["schedule"]),
        ]),
        ("FINANCIAL POSTURE (PRE-CONSTRUCTION)", [
            ("Original Contract", "Not yet executed"),
            ("Revised Contract",  "Not yet executed"),
            ("AR Billed to Date", f"${META_2108['ar_billed']:,}"),
            ("Direct Cost (sunk)",f"${META_2108['design_cost']:,.0f}"),
            ("Carrying Loss",     f"$({abs(META_2108['carrying_loss']):,.0f})"),
            ("Retainage",         "$0 (no AR)"),
            ("Insurance",         META_2108["insurance"]),
            ("Lien Position",     META_2108["lien_position"]),
            ("Warranty",          META_2108["warranty"]),
            ("Contract on File",  META_2108["contract_doc"]),
        ]),
        ("SCOPE", [
            ("Plumbing Units",  META_2108["units"]),
            ("Total Fixtures",  META_2108["fixtures"]),
            ("Floors",          "TBD"),
            ("Fixture Counts",  "TBD"),
        ]),
        ("PROJECT TEAM — GENERAL CONTRACTOR", [
            ("General Contractor", META_2108["gc"]),
            ("GC Project Manager", META_2108["gc_pm"]),
            ("GC Superintendent",  META_2108["gc_sup"]),
            ("GC Project Engineer", META_2108["gc_pe"]),
        ]),
        ("PROJECT TEAM — OWP STAFF", [
            ("OWP Roughin Foreman", META_2108["owp_ri_foreman"]),
            ("OWP Trim Foreman",    META_2108["owp_trim_foreman"]),
            ("OWP Estimator (takeoff)", META_2108["owp_estimator"]),
            ("OWP Signatory",       META_2108["owp_signatory"]),
        ]),
        ("PROJECT TEAM — OWNER & DEVELOPMENT", [
            ("Owner of Record", META_2108["owner"]),
            ("Developer",       META_2108["developer"]),
        ]),
        ("PROJECT TEAM — DESIGN", [
            ("Architect",          META_2108["architect"]),
            ("Structural Engineer", META_2108["structural"]),
            ("MEP / Plumbing Engineer", META_2108["mep_engineer"]),
            ("Civil Engineer",     META_2108["civil"]),
            ("Landscape",          META_2108["landscape"]),
        ]),
        ("DOCUMENT META (current)", [
            ("Pay Apps (filed)",   "0"),
            ("Executed COs",       "0"),
            ("CORs",               "0"),
            ("RFIs",               "0"),
            ("Submittals",         "0"),
            ("POs",                "0"),
            ("Permit Count",       "0 (pre-permit)"),
            ("Notes",              "Project hasn't reached document-generation phase. "
                                   "Only the OWP estimating team has touched the file (takeoff + "
                                   "design coordination)."),
        ]),
        ("DATA SOURCES", [
            ("JDR PDF",            "2108 Job Detail Report (Sage Timberline · Apr 3 2026 run)"),
            ("Parsed Data",        "2108_data.json"),
            ("Dashboard Arrays",   "2108_dashboard_arrays.json (16 arrays sourced from index.html PROJECTS['2108'])"),
            ("GDrive Folder",      "Not yet created"),
            ("GDrive Status",      META_2108["gdrive_status"]),
        ]),
        ("RISK FLAGS", [
            ("Carrying $91k design cost", "OWP has $117k of sunk costs. If project cancels, "
                                          "OWP eats the design + takeoff time. Reach out to Braseth "
                                          "PM to confirm summer 2026 start before Q4 2026."),
            ("Top-vendor concentration",  "Franklin Engineering = 92.5% of $97k AP. All design fee."),
            ("New GC (Braseth)",          "OWP has no closed-job history with Braseth. Once project "
                                          "activates, watch payment cadence + change-order behavior closely."),
        ]),
    ]

    r = 4
    for section_name, fields in sections:
        ws.cell(row=r, column=1, value=section_name).font = SECTION_HDR
        ws.cell(row=r, column=1).fill = SECTION_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(row=r, column=1).border = BORDER
        r += 1
        for label, val in fields:
            ws.cell(row=r, column=1, value=label).font = INK_BOLD
            cell_v = ws.cell(row=r, column=2, value=val)
            cell_v.font = INK
            cell_v.alignment = WRAP
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.cell(row=r, column=1).border = BORDER
            cell_v.border = BORDER
            if section_name.startswith("PROJECT TEAM") or section_name == "RISK FLAGS":
                ws.cell(row=r, column=1).fill = TEAM_FILL
                cell_v.fill = TEAM_FILL
            r += 1
        r += 1

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 30


# ============================================================================
# OVERVIEW tab — fix CANCELLED status; add Project Team block
# ============================================================================
def patch_overview_team_block(ws):
    """Find any cell containing 'CANCELLED' on rows 1–10 and replace with the
    correct status, then add a project-team block below the existing header."""
    for r in range(1, 12):
        for c in range(1, 12):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "CANCELLED" in v:
                ws.cell(row=r, column=c).value = v.replace(
                    "CANCELLED", META_2108["status_text"])
                ws.cell(row=r, column=c).font = WARN_FONT

    # Add Project Team block at row 40+
    start = 40
    # Wipe any prior content in the team-block area
    clear_range(ws, start, start + 30, 1, 8)

    ws.cell(row=start, column=1, value="PROJECT TEAM · KEY PLAYERS").font = SECTION_HDR
    ws.cell(row=start, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=7)

    ws.cell(row=start+1, column=1, value=(
        "Pre-construction. Project on hold pending Braseth go-ahead for Summer 2026 start. "
        "Team data from index.html PROJECT_TEAMS + JDR identity block."
    )).font = GREY_FONT
    ws.merge_cells(start_row=start+1, start_column=1, end_row=start+1, end_column=7)

    rows = [
        ("General Contractor",  META_2108["gc"]),
        ("GC Project Manager",  META_2108["gc_pm"]),
        ("OWP Estimator",       META_2108["owp_estimator"]),
        ("OWP Signatory",       META_2108["owp_signatory"]),
        ("Owner",               META_2108["owner"]),
        ("Architect",           META_2108["architect"]),
        ("MEP / Plumbing Engineer", META_2108["mep_engineer"]),
        ("Insurance",           META_2108["insurance"]),
        ("Status",              META_2108["status_text"]),
        ("AR Billed to Date",   f"${META_2108['ar_billed']:,}"),
        ("Direct Cost (sunk)",  f"${META_2108['design_cost']:,.0f}"),
        ("Carrying Loss",       f"$({abs(META_2108['carrying_loss']):,.0f})"),
    ]
    r = start + 3
    ws.cell(row=r, column=1, value="ROLE").font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HDR_FILL
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=2, value="VALUE").font = WHITE_BOLD
    ws.cell(row=r, column=2).fill = HDR_FILL
    ws.cell(row=r, column=2).border = BORDER
    r += 1
    for role, val in rows:
        ws.cell(row=r, column=1, value=role).font = INK_BOLD
        ws.cell(row=r, column=2, value=val).font = INK
        ws.cell(row=r, column=1).fill = TEAM_FILL
        ws.cell(row=r, column=2).fill = TEAM_FILL
        ws.cell(row=r, column=1).border = BORDER
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=2).alignment = WRAP
        r += 1


# ============================================================================
# PREDICTIVE SIGNALS tab — expand from 4 rows to richer set
# ============================================================================
def patch_predictive_signals(ws, dash):
    signals = dash.get("predictiveSignals", [])
    clear_range(ws, 1, 30, 1, 6)
    ws["A1"] = f"Predictive Signals — Job #{META_2108['job_id']} (pre-construction)"
    ws["A1"].font = TITLE_FONT

    headers = ["#", "Signal", "Current Value", "Threshold", "Severity"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_BOLD
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BORDER

    # Combine the existing dashboard signals with pre-construction-specific ones
    extra = [
        ["Project status",            META_2108["status_text"], "ACTIVE",          "WATCH"],
        ["Sunk design cost (carrying)", f"${META_2108['design_cost']:,.0f}", "$0",      "WATCH"],
        ["AR billed",                 f"${META_2108['ar_billed']:,}",   ">$1M for medium job", "INFO"],
        ["Pay apps filed",            "0",                       ">=1",                "INFO"],
        ["GC closed-job history",     "0 jobs with Braseth",      ">=1",                "WATCH"],
    ]
    rows = [s if isinstance(s, list) else list(s) for s in signals] + extra
    r = 4
    for i, s in enumerate(rows, 1):
        if not isinstance(s, list) or len(s) < 4: continue
        signal, val, thr, sev = s[0], s[1], s[2], s[3]
        ws.cell(row=r, column=1, value=i).font = INK
        ws.cell(row=r, column=2, value=signal).font = INK
        ws.cell(row=r, column=3, value=val).font = INK
        ws.cell(row=r, column=4, value=thr).font = INK
        sev_cell = ws.cell(row=r, column=5, value=sev)
        sev_cell.font = INK_BOLD
        if sev == "WARN" or sev == "WATCH": sev_cell.fill = WARN_FILL
        elif sev == "CRIT" or sev == "CRITICAL": sev_cell.fill = PatternFill("solid", start_color="F8D7DA")
        elif sev == "HEALTHY": sev_cell.fill = PatternFill("solid", start_color="D4EDDA")
        elif sev == "INFO":  sev_cell.fill = PatternFill("solid", start_color="D1ECF1")
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDER
        r += 1

    widths = {1: 5, 2: 38, 3: 22, 4: 24, 5: 12}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w


# ============================================================================
# Patch CHANGE LOG with proper change-event log
# Note: "17 Change Log" is the workbook change-log in 2108/2109 schema.
# Since there are no project change events on a pre-construction job, leave the
# workbook change-log as-is and add a row noting the enrichment pass.
# ============================================================================
def patch_change_log(ws):
    # Find the next empty row after row 3 header
    last_row = ws.max_row
    r = last_row + 1
    ws.cell(row=r, column=1, value="2026-04-27").font = INK
    ws.cell(row=r, column=2, value="v1.1 · enriched").font = INK_BOLD
    ws.cell(row=r, column=3, value=(
        "Enriched Job Info tab with 11-section sectioned layout (identity / schedule / "
        "financial posture / scope / project team — 4 sub-sections / document meta / "
        "data sources / risk flags). Corrected status from 'CANCELLED' (template "
        "default) to 'ON HOLD · Summer 2026?'. Predictive Signals expanded with "
        "pre-construction-specific signals."
    )).font = INK
    ws.cell(row=r, column=3).alignment = WRAP


# ============================================================================
# Main
# ============================================================================
def main():
    print(f"Loading {WORKBOOK.name}...")
    wb = load_workbook(WORKBOOK)
    data = json.loads(JSON_FILE.read_text()) if JSON_FILE.exists() else {}
    dash = json.loads(DASH_FILE.read_text())

    print("→ Rebuilding Job Info tab (sectioned layout)...")
    rebuild_job_info(wb["02 Job Info"])

    print("→ Patching Overview status + adding Project Team block...")
    patch_overview_team_block(wb["01 Overview"])

    print("→ Expanding Predictive Signals tab...")
    patch_predictive_signals(wb["14 Predictive Signals"], dash)

    print("→ Logging enrichment pass to Change Log...")
    patch_change_log(wb["17 Change Log"])

    wb.save(WORKBOOK)
    print(f"\n✓ Saved {WORKBOOK.name} ({len(wb.sheetnames)} tabs)")


if __name__ == "__main__":
    main()
