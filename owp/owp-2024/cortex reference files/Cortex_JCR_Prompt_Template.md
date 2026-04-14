# Cortex JCR Summary — Prompt Template (v2 — Live Predictive Edition)

Use this template every time you start a new OWP project file. Copy the **Prompt** section at the bottom, fill in the bracketed fields, and paste into a fresh Claude chat.

**Version 2 changes:** the JCR now includes 17 tabs covering the full change ecosystem (Change Log, Root Cause Analysis, Predictive Signals, Metric Registry) and every metric is tagged with `data_label`, `confidence_level`, and `source_document(s)` for programmatic Cortex ingestion.

---

## Step 1 — Gather Inputs

Before pasting the prompt, have these ready:

- **Job Number** (e.g. 2012)
- **Project Name** (e.g. Exxel 8th Ave Apartments)
- **Job Detail Report PDF** — uploaded or available in the project folder
- **Project folder path**
- **Unit Count** and **Fixture Count** (verified from permit/P-tag list)
- **Duration in months**
- **Contract Value** and **Revenue**
- **Change ecosystem folders:** ASI-RFI, Invoices/64-Backcharges, Invoices/38-T&M, Change Orders (if present)

---

## Step 2 — Paste the Prompt

### Project Inputs

- Job Number: [JOB_NUMBER]
- Project Name: [PROJECT_NAME]
- JCR PDF Path: [PATH_TO_JCR.pdf]
- Project Folder: [PATH_TO_PROJECT_FOLDER]
- Unit Count: [N]
- Fixture Count: [N]
- Duration (months): [N]
- Contract Value: $[N]
- Revenue: $[N]

### Prompt

I'm building Cortex, a predictive intelligence platform for OWP (One Way Plumbing). Build a 17-tab JCR + Live-Intelligence spreadsheet for Job #[JOB_NUMBER] ([PROJECT_NAME]) following the exact same format, structure, and aesthetic as Job #2012 (reference: OWP_2012_JCR_Cortex_v2.xlsx).

#### Workflow

1. Extract all cost codes from the JCR PDF using pdfplumber/pypdf. Parse "Cost Code Totals" lines and "Payroll Hours" lines per code. Extract per-worker PR transactions and per-vendor AP transactions.
2. Scan the project folder for change-ecosystem documents: ASI-RFI/, Invoices/64-Backcharges/, Invoices/38-T&M/, Submittals/ (for ASI references), Emails Saved/, Meetings-Schedules/. Build a Change Log row for each artifact found.
3. Build the spreadsheet with 17 tabs in this exact order: Overview, Budget vs Actual, Cost Breakdown, Material, Crew & Labor, Crew Analytics, Productivity, PO Commitments, Billing & SOV, Insights, Benchmark KPIs, Vendors, **Change Log**, **Root Cause Analysis**, **Predictive Signals**, **Metric Registry**, Reconciliation.
4. Every roll-up must tie out exactly to the JCR within $1 rounding. The JCR PDF is the canonical source of truth.
5. Every metric across every tab MUST carry three metadata attributes: `data_label` (snake_case machine key), `confidence_level` (Verified / Medium / Low), `source_document(s)` (file path or section reference).

#### Tab Requirements

Tabs 1-12 (same as v1 with additions):

- **Overview:** KPI cards, project meta, eyebrow label
- **Budget vs Actual:** All cost codes with Status tags (OVER, CRITICAL, UNDER, ON, UNBUDGETED)
- **Cost Breakdown:** PR / AP / GL rollups with vendor and worker splits
- **Material:** 2xx + 039 codes with $/Unit, $/Fixture, % of Total, Status columns
- **Crew & Labor:** Tier bands, blended rates, burden multiplier
- **Crew Analytics:** Per-worker table (Tier: Lead ≥ $32, Journeyman ≥ $22, Apprentice ≥ $15, Helper < $15)
- **Productivity:** Hours, $, hrs/unit, hrs/fixture, $/unit, $/fixture per phase
- **PO Commitments:** All scheduled POs with vendor, amount, date, cost code
- **Billing & SOV:** Pay applications with retainage tracking
- **Insights:** Variance narrative and risk callouts
- **Benchmark KPIs:** Six columns — KPI, data_label, Value, Category, **Confidence**, **Source Document**
- **Vendors:** Ranked vendor list with AP spend, invoices, tie-out

Tabs 13-16 (NEW — predictive / live intelligence layer):

- **Change Log:** Master register of all change events (RFI, ASI, COP, CO, Backcharge, T&M directive). Columns: Event ID, Type, Date, Subject, Originator, Linked Event(s), Cost Impact, Schedule days, Status, Root Cause, Responsible Party, Source Doc. Cross-reference IDs so `RFI-023 → ASI-015 → COP-007 → CO-003` traces cleanly.
- **Root Cause Analysis:** Category rollup using this taxonomy: Design error, Design substitution, Owner directive, Field condition, Coordination conflict (MEP, trade interface), Scope gap/ambiguity, Unforeseen condition, Acceleration, Cleanup/damage, Rework, Combined/multi-cause. Each category with # events, net $ impact, % of events, predominant responsibility, and notes. Plus a Responsibility Attribution matrix with bid-risk signals.
- **Predictive Signals:** Leading indicators (RFI velocity, RFI aging, ASI-to-COP lag, RFI-to-ASI conversion, cumulative CO ratio, T&M burn rate, backcharge frequency, submittal resubmission rate). Plus Forecast Models block (projected final CO %, projected EAC, projected margin, projected completion slip, composite risk score 0-100, GC scorecard A-F).
- **Metric Registry:** Machine-readable catalog of every metric. Columns: #, Data Label (snake_case), Human Label, Value, Unit, Source Tab, Confidence, Source Document(s), Formula/Derivation. This is the Cortex API surface — everything downstream consumes this.

Tab 17:

- **Reconciliation:** Cross-sheet formula tie-outs for all 12 analytical tabs. Sections A-J with TIES/WITHIN/OFF status. Auto-tag: TIES if |diff| ≤ $1, WITHIN if ≤ 5% or 5 hrs, else OFF.

#### Metric Metadata Schema (applied to EVERY metric)

Every metric cell or row must be accompanied by:

- `data_label` — snake_case programmatic identifier (e.g., `project_revenue`, `rfi_avg_cycle_days`, `co_ratio_pct`). Stable across all projects so Cortex can benchmark longitudinally.
- `confidence_level` — one of: **Verified** (green — directly sourced from a ground-truth document), **Medium** (yellow — derived or partially inferred), **Low** (red — approximated, missing source, or flagged for field verification).
- `source_document(s)` — file path, section reference, or named artifact (e.g., "JDR Cost Code 038", "Submittals/Kavela ASI 25.pdf", "Aggregate Billing row 6").

Apply via (a) inline columns on Benchmark KPIs and Metric Registry, and (b) a complete inventory in the Metric Registry tab.

#### Design System

- Font: Arial throughout. White background, no gridlines.
- Notion-inspired muted palette (red/green/blue/purple/pink/gray/brown/yellow/orange backgrounds with paired text colors)
- Confidence color convention: Verified = green_bg (#E2EFDA), Medium = yellow_bg (#FFF2CC), Low = red_bg (#FFE6E6)
- Status tag convention: TIES/HEALTHY/POSITIVE = green; WITHIN/NEUTRAL/Medium = yellow; OFF/ELEVATED/CAUTION/Low = red
- Generous row heights (26-34px), soft borders (#BFBFBF), alternating row fill

#### Must-Haves

- Use Excel formulas, not hardcoded calculated values
- Cross-sheet formulas (e.g., `='Benchmark KPIs'!D12`) — always single-quote sheet names containing spaces
- In Metric Registry, formula-description cells (col J) are TEXT — strip leading `=` if referencing a formula so LibreOffice doesn't try to evaluate
- Every tab ends with a Sources footer citing the JCR PDF and document references
- Run LibreOffice recalc (`python scripts/recalc.py`) at the end — confirm zero formula errors
- Verify Reconciliation shows all TIES before finishing
- Save as `OWP_[JOB]_JCR_Cortex_v2.xlsx` and return a computer:// link

#### Live-Project Capture Mode

When Cortex runs on a LIVE project (not a retroactive analysis), the Change Log + Predictive Signals tabs become the primary interface. New change events append as rows. Predictive Signals recalculates weekly. Metric Registry updates as data sources arrive. Low-confidence metrics get flagged for PM review.

---

## Reference

The canonical reference implementation is Job #2012 (Exxel 8th Ave Apartments). The build scripts are `gen_2012.py` (base 13-tab JCR) + `extend_jcr_cortex_v2.py` (adds Change Log, Root Cause, Predictive Signals, Metric Registry). The reference output is `OWP_2012_JCR_Cortex_v2.xlsx`. For any ambiguity, mirror that file's structure exactly.
