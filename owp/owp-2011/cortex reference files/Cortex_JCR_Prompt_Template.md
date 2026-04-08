# Cortex JCR Summary — Prompt Template

Use this template every time you start a new OWP project file. Copy the **Prompt** section at the bottom, fill in the bracketed fields, and paste into a fresh Claude chat.

---

## Step 1 — Gather Inputs

Before pasting the prompt, have these ready:

- **Job Number** (e.g. 2012)
- **Project Name** (e.g. Exxel 8th Ave Apartments)
- **Job Detail Report PDF** — uploaded or available in the project folder
- **Project folder path**
- **Unit Count**
- **Fixture Count**
- **Duration in months**
- **Contract Value**
- **Revenue** (if different from contract)

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

I'm building Cortex, an intelligence platform for OWP (One Way Plumbing). Build a Notion-style JCR summary spreadsheet for Job #[JOB_NUMBER] ([PROJECT_NAME]) following the exact same format, structure, and aesthetic as Job #2012.

#### Workflow

1. Extract all cost codes from the JCR PDF using pdfplumber. Parse "Cost Code Totals" lines (format: Budget, Revised Budget, Variance, Actual as the 4th number) and "Payroll Hours" lines per code. Also extract per-worker PR transactions (hours, OT, wages) and per-vendor AP transactions (invoices, dollars).
2. Build the spreadsheet with 10 tabs in this exact order: Overview, Budget vs Actual, Material, Cost Breakdown, Crew & Labor, Insights, Productivity, Benchmark KPIs, Crew Analytics, Reconciliation.
3. Every roll-up must tie out exactly to the JCR within $1 rounding. The JCR PDF is the canonical source of truth.

#### Tab Requirements

- **Overview:** KPI cards, project meta, eyebrow label
- **Budget vs Actual:** All cost codes (labor + material + overhead + burden). Columns: Phase, Budget, Actual, Variance $, Variance %, Hours, $/Hr, Status. Status tags: OVER, CRITICAL, UNDER, ON, UNBUDGETED
- **Material:** Only 2xx + 039 codes with $/Unit, $/Fixture, % of Total, Status columns, plus insights callouts
- **Cost Breakdown:** PR / AP / GL rollups with vendor and worker splits
- **Crew & Labor:** Tier bands, blended rates, burden multiplier
- **Insights:** Variance narrative and risk callouts
- **Productivity:** Hours, $, hrs/unit, hrs/fixture, $/unit, $/fixture per phase
- **Benchmark KPIs:** Five columns — KPI, Data Name (snake_case), Value, Category, Notes. Categories: Profile, Financial, Labor, Crew, Throughput, Cost Mix, Estimating, Material
- **Crew Analytics:** Per-worker table — Name, ID, Reg Hrs, OT Hrs, Total, OT %, Wages, $/Hr, Codes touched, Tier tag. Tiers: Lead ≥ $32, Journeyman ≥ $22, Apprentice ≥ $15, Helper < $15
- **Reconciliation:** Side-by-side tie-out proof with sections for Grand Totals, Labor, Material, Overhead + Burden, Crew Analytics, Cross-Tab Consistency, Vendor AP. Each row: Metric, JCR Source (hardcoded), Workbook (cross-sheet formula), Difference, Status. Auto-tag: TIES if |diff| ≤ $1, WITHIN if ≤ 50 for crew hours, else OFF

#### Design System (Notion Aesthetic)

- Font: Arial, white background, no gridlines on any sheet
- Muted Notion palette:
    - red_bg #FBE4E4
    - green_bg #DDEDEA
    - blue_bg #DDEBF1
    - purple_bg #EAE4F2
    - pink_bg #F4DFEB
    - gray_bg #EBECED
    - brown_bg #EEE0DA
    - yellow_bg #FBF3DB
    - orange_bg #FBECDC
- All tag backgrounds paired with matching text colors
- Generous row heights (26–34px), soft borders (#EAEAEA), alternating row fill (white / #FAFAFA)
- Section headers in #F7F7F5 card background
- Status and category tags as filled pills with paired bg/text colors
- KPI card titles in 9pt uppercase muted text; values in 28pt bold

#### Must-Haves

- Use Excel formulas, not hardcoded calculated values (totals, variances, per-unit metrics)
- Every tab ends with a Sources footer block citing the JCR PDF and explaining where each column's data came from
- Run LibreOffice recalc at the end and confirm zero formula errors (#REF!, #DIV/0!, #VALUE!, #NAME?)
- Verify the Reconciliation tab shows all TIES before finishing
- Save the final file to the project folder as OWP_[JOB]_JCR_Summary_Notion.xlsx and return a computer:// link

---

## Reference

The canonical reference implementation is Job #2012 (Exxel 8th Ave Apartments). The build script lives at build_notion_jcr.py and the output at OWP_2012_JCR_Summary_Notion.xlsx. For any ambiguity, mirror that file's structure exactly.
