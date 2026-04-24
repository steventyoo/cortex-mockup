# Cortex JCR Portfolio Dashboard — Project Knowledge Base

> This file is read by Claude at the start of every session. Keep it current.
> Last updated: 2026-04-24

## Active/live projects (last updated 2026-04-24)
Eleven live projects are now wired into the dashboard (status: 'active'):

- **#2098 Chinn Trailside 2** (170 units · $6.92M · 100% billed · retention-release · forecast 25% margin · retention $345k)
- **#2103 Compass General Northgate M2** (234 units · $6.63M · 99.75% billed · trim/punch · forecast 21.9% vs 28% target · retention $305k)
- **#2104 BMDC Big 1410** (24 units · $1.20M · 99.76% billed · in trim · foreman Tommy/Ben)
- **#2111 Compass Northgate M3** (186 units · $4.51M · 48.8% billed · roughin · Franklin Engineering MEP · $89.9k COs / 7 executed)
- **#2112 BMDC Big 3** (20 units · $696k · 74.4% billed · trim · Robison Engineering · $32.3k COs / 9 executed)
- **#2114 Holland Ballard Blossom** (239 units · $5.67M · 1.4% billed · pre-construction · **NEW GC Holland** · Robison MEP · Sep 2025 design → Jul 2026 foundations → Apr 2028 TCO · OCIP · 0 COs)
- **#2115 Chinn Elowen** (2-building multifamily · Lake Stevens WA · $3.94M · design/DD · **Chinn's 7th OWP job** · Charlie Poggemann PM · North Cove Investors owner · design signed Dec 3, 2025 · $80.4k design fee)
- **#2116 SRM The V** (194 units · 7-story Seattle · $3.26M · GMP executed Feb 24, 2026 · SRM Spokane · Erik Benzel PM (same as #2020) · NYL owner · ECE MEP · COR#2 in flight · heat pump upsize +$142.8k pending)
- **#2117 Ravenna Partners Lachlan** (~104 units · 8-story + penthouse · 3421 Woodland Park Ave N Seattle · design/DD · **NEW GC Ravenna Partners** (owner-GC) · Franklin Engineering MEP · job ticket Jan 14 2026 · $11k design-phase expenses · gas water heating · standard COI)
- **#2118 Exxel Edmonds Behar** (409 units · 6-story + 2 towers + 3 below-grade · Edmonds WA · **$7.18M total plumbing budget** — OWP's largest live-pipeline bid · 22-month bid history Mar 2024 → Jan 2026 · Robison design contract signed Mar 16, 2026 ($125k) · Tommy Booth PM Exxel · **no JDR yet** — data.json is a stub, parse script tolerates missing PDF · will auto-populate when Sage publishes a JDR)
- **#2119 Exxel The Frank (508 Fremont)** (earliest live-pipeline project · Exxel Pacific · Graham developer · Michael Fisher PM (Bellingham) · Seattle WA · **pre-bid / mockup-only phase** · only $4,700 mockup-install proposal Nov 13 2024 · Exxel bid set received but OWP full proposal pending · no JDR, no contract value yet · stub data.json matching 2118 pattern)

Added 2026-04-23: #2098 and #2103 (both via the same `build/notionize/parse_XXXX.py` pipeline as 2104/2111/2112, output in `owp/owp-<id>-live/cortex output files/`). Both ship the full `.live` block (burnCurve, phases, anomalies, recommends) alongside the hero fields.

Added 2026-04-24: #2114 Holland Ballard Blossom. First live project in pre-construction / design-DD state — bid $5.67M accepted May 2025, $80k design billing to date, field foundations scheduled Jul 2026. New GC (Holland) and new MEP partner (Robison) — no historical BvA baseline. `.live` block uses a "notstarted" phase ladder and a baseline-only burnCurve to represent pre-construction state honestly. Full pipeline in `owp/owp-2114-live/cortex output files/` (parse_2114.py + build_2114.py + notionize_2114.py, modeled off 2098 as the richest live template). Output: 17-tab JCR + 2-tab Notion summary, zero formula errors. Holland becomes OWP's 10th GC.

Added 2026-04-24 (same batch as 2114): **#2115 Chinn Elowen** and **#2116 SRM The V**, both pre-construction. Pipeline cloned from 2114 (parse/build/notionize) and META-patched for each project:
- **2115 Elowen** (Lake Stevens, WA) — Chinn's 7th OWP engagement. $3.94M base bid signed Dec 3, 2025 by Matthew Burton (North Cove Investors). 2-building project (Bldg A $2.80M + Bldg B $1.06M). OWP design-build (2-D CAD, not 3-D BIM). $12.9k expenses (design consulting + 26 hrs Gerard takeoff). Output: 17-tab JCR (19KB) + 2-tab Notion (9KB), zero formula errors. No GDrive contract subcontract yet filed.
- **2116 The V** (Seattle, WA) — 194 units, 7-story. SRM Construction (Spokane, Erik Benzel PM — same PM as closed #2020). GMP executed Feb 24 2026 (NYL). $3.26M base bid. MEP: Emerald City Engineers. COR#2 Plumbing Changes already in flight (Mar 25). Fixture schedule captured: 194 Toto toilets, 103 tub/shower stalls, 60 Mincey pans, 166 kitchen sinks, 26 Sanco2 heat pumps. Heat pump system flagged as undersized — upsize alternate +$142,800 pending. Output: 17-tab JCR + 2-tab Notion (13KB), zero formula errors. OWP subcontract under GMP not yet filed.

All three (#2114/2115/2116) use the same pre-construction template: burnCurve with one early actual point (design/takeoff) and baseline-only forward points; phases ladder with "notstarted" status for field phases; anomalies recording bid milestones + proposal dates + contract execution; recommends surfacing what needs to happen before field work. 2115's folder name uses a hyphen ("owp-2115-live") to match dashboardize.py convention; originally had a space which broke auto-detection — renamed during wiring.

Added 2026-04-24 (third batch — 2117/2118): **#2117 Ravenna Partners Lachlan** and **#2118 Exxel Edmonds Behar**, both pre-construction. 2117 has a thin JDR (2 pages, $11k design-phase expenses, 3 hrs Gerard takeoff); 2118 has NO JDR yet — data.json is a stub placeholder matching the full v2.2 schema so build/notionize scripts produce the same 17-tab JCR + 2-tab Notion shape as peers, just with zero actuals. Key facts:
- **2117 Lachlan** (Seattle, Woodland Park) — 104-unit studio apartment, 8-story + penthouse. First OWP engagement with Ravenna Partners (owner-GC). Franklin Engineering MEP. Fixture schedule locked Jan 14 2026 (Gerber Maxwell toilets, Kohler K-2035-1 wall-hung lavs, Sterling Vikrell tubs + shower pans, Peerless matte black trims, gas water heating). Contract value TBD — bid sheet not yet filed to Drive. Output: 17-tab JCR (19KB) + 2-tab Notion (7.6KB), zero formula errors.
- **2118 Edmonds Behar** (Edmonds WA) — 409-unit, 6-story + 2 towers + 3 below-grade. **OWP's largest live-pipeline bid** at $7.18M total plumbing budget (design $173.8k + permits $77.5k + construction $6.33M + fixtures $606.8k). Exxel Pacific GC with Tommy Booth as PM (same PM as closed #2017 Mack Urban + #2022). 22-month bid history since Mar 2024. Robison Engineering MEP design contract signed Mar 16 2026 for $125k. GMP set dated Nov 27 2024. No Sage JDR yet — parse_2118.py is a placeholder; 2118_data.json is a hand-written stub matching schema. Output: 17-tab JCR (17.8KB) + 2-tab Notion (6.7KB), zero formula errors even on empty data.

Ravenna Partners becomes OWP's 11th GC in the dashboard (joining Chinn/SRM/Exxel/Compass/BMDC/Blueprint/Natural & Built/Marpac/Shelter/Holland/GRE). 2117/2118 folders originally had spaces ("owp-2117 live", "owp-2118 live") — renamed to hyphen form to match dashboardize.py convention.

Added 2026-04-24 (fourth batch — 2119): **#2119 Exxel Pacific, The Frank (508 Fremont)** — earliest-stage project in entire live pipeline. No Sage JDR yet; no full OWP construction bid yet. Only a $4,700 mockup-install proposal sent Nov 13, 2024 to Michael Fisher (Exxel Bellingham PM). Exxel bid set received per email reminder; OWP's full plumbing proposal + bid sheet are pending. Developer identified as Graham (per GDrive folder naming). Same "pre-bid mockup" pattern as #2113 Yonder. data.json is a stub matching v2.2 schema; build/notionize still produce the 17-tab + 2-tab outputs cleanly. Output: 17-tab JCR (17.8KB) + 2-tab Notion (6.7KB), zero formula errors. Folder originally "owp-2119 live" (space); renamed to hyphen form.

Live project pipeline is nearly identical to closed-job pipeline but:
- Uses 'revised' budget as canonical basis (not 'original') — `rev_contract`, `rev_expense` in build script
- Forecast margin = (rev_contract - rev_expense) / rev_contract (not actual)
- `pct_complete` = billed / rev_contract
- PROJECTS entries include extra hero fields: `pctComplete`, `contractOrig`, `contractRevised`, `coPct`, `billedToDate`, `earnedLessRetention`, `balanceToFinish`
- Separate `PROJECTS['XXXX'].live = { ... }` block with `burnCurve`, `phases`, `anomalies`, `recommends` arrays — required for Job Health view to render (burn-curve chart, phase ladder, anomaly feed, action cards)
- PROJECT_TEAMS entries use clay-colored Status field like "ACTIVE · 48.8% billed"
- Document Pipeline cards use LIVE pattern: `border: 1.5px solid var(--clay); background: rgba(184,92,62,0.04);` with "LIVE" chip in kicker
- Atlas view chip shows `11 active` now (was 10)
- Job Health view chip shows `11 active jobs` now (was 10)

Live build scripts: `/sessions/gracious-relaxed-pascal/mnt/cortex-mockup/owp/owp-<id>-live/cortex output files/build_<id>.py` + `notionize_<id>.py` (same code with different META dict). Output: 17-tab JCR + 2-tab Notion summary, zero formula errors verified.

## Owner
Steven Yoo (steventyoo@gmail.com) — One Way Plumbing (OWP)

## What This Project Is
Cortex is a comprehensive Job Cost Report (JCR) portfolio system for OWP's plumbing projects. It has two main outputs:

1. **Dashboard** (`index.html`) — A single-file HTML/JS dashboard hosted on Vercel (auto-deploys from GitHub `steventyoo/cortex-mockup` on `main` branch). Contains all 32 OWP projects with interactive cards, charts, cost analysis, crew data, and change order tracking.

2. **Per-project JCR workbooks** — Excel files (openpyxl) with 17 tabs: Overview, Job Info, Contract, SOV/Pay Apps, Change Orders, Cost Codes, Cost Categories, BVA, Vendor Analysis, Crew Roster, Wage Tiers, Productivity, Benchmarks, Predictive Signals, Reconciliation, Metric Registry, Change Log.

## Project List (32 projects)
| ID | GC | Project Name | Units |
|----|-----|-------------|-------|
| 2001 | SRM | 230 Broadway | 234 |
| 2008 | SRM | 101 Taylor | 258 |
| 2009 | Chinn | Greenwood Ave | 56 |
| 2010 | Chinn | Old Town | 149 |
| 2011 | Chinn | Phinney Ridge | 117 |
| 2012 | Exxel | 12th & Madison (The Lyric) | 162 |
| 2015 | Exxel | Portage Bay | 197 |
| 2016 | Natural & Built | Admiral Way | 80 |
| 2017 | Exxel | Mack Urban (Station 7) | 119 |
| 2018 | SRM | Merrill Gardens First Hill | 211 |
| 2019 | Exxel | Wolff (Anthem) | 211 |
| 2020 | SRM | Pillar (Vox) | 256 |
| 2021 | Exxel | Intra-Corp (Kinects) | 136 |
| 2022 | Exxel | Mack Urban (Station 7 Ph2) | 18 |
| 2023 | Chinn | Legacy Apartments | 209 |
| 2024 | SRM | Ballard Merrill Gardens | 104 |
| 2025 | Exxel | The Cora | 75 |
| 2026 | Synergy | Fox & Finch (525 Boren) | 49 |
| 2027 | Exxel | Zig Apts (550 Broadway) | 170 |
| 2028 | Exxel | Reserve Lynnwood | 296 |
| 2029 | Exxel | Parsonage | 83 |
| 2030 | Exxel | East Union | 144 |
| 2031 | Exxel | Issaquah Gateway | 398 |
| 2032 | Blueprint | Luna Apts-Roosevelt (6921 Roosevelt Way NE) | 71 |
| 2033 | Compass Harbor | Compass Vuecrest (Bellevue) | 137 |
| 2034 | Compass Harbor | Compass Park Lane Apts (Kirkland) | 128 |
| 2035 | Natural & Built | 162 Ten (Redmond) | 92 |
| 2036 | Exxel | Westridge | 31 |
| 2037 | Marpac | University Apartments (MYSA) | 122 |
| 2038 | Compass Harbor | 2nd & John | 80 |
| 2039 | Shelter Holdings | Ravello Gas Piping (HOA remediation) | 1 |
| 2040 | Blueprint | Brooklyn 65 | 55 |

## Architecture

### Dashboard (index.html)
- Single HTML file with embedded JS
- `PROJECTS` object: main data for each project (id, name, gc, contract, expenses, profit, margin, status, etc.)
- `PROJECT_TEAMS` object: team HTML and unit counts, merged into PROJECTS at runtime
- `PROJECT_ORDER` array: controls rendering order
- **16 extended data arrays per project** (assigned as `PROJECTS['XXXX'].arrayName`):
  - `sovData` — Statement of Values summary
  - `payApps` — Payment application history
  - `allVendors` — AP vendor list with spend
  - `phases` — Phase breakdown by hours
  - `insights` — 10 key narrative insights
  - `changeLog` — Change order/RFI log entries
  - `changeMeta` — CO/ASI/RFI summary counts
  - `rootCauses` — Root cause analysis
  - `responsibility` — Responsibility matrix
  - `predictiveSignals` — Health indicators
  - `crewRoster` — Worker hours/wages
  - `tierDist` — Wage tier distribution
  - `wageStats` — Wage statistics
  - `costCodes` — Sage cost code detail [code, desc, budget, actual, variance]
  - `costCats` — Category rollup [category, count, budget, actual, pct, perUnit]
  - `bva` — Bid vs Actual [label, budget, actual, variance, pct, null, null, status]
- BVA status tags: "CRITICAL" (>50% over), "OVER" (>10% over), "ON" (±10%), "UNDER" (>10% under), "UNBUDGETED"

### Per-Project Pipeline (3 scripts per project)
Located in each `owp-XXXX/cortex output files/` folder:
1. `parse_XXXX.py` — Parses the Sage Timberline JDR PDF → `XXXX_data.json`
2. `build_XXXX.py` — Builds the 17-tab JCR workbook from parsed data
3. `notionize_XXXX.py` — Builds the Notion-styled summary workbook

### Data Sources
- **JDR PDFs** — Sage Timberline Job Detail Reports (one per project, in each owp-XXXX folder)
- **Google Drive folders** — Mounted as `owp-XXXX` folders, contain COs, RFIs, submittals, permits, POs, daily reports, pay apps, O&M docs
- **Budget Transfer spreadsheets** — In GDrive CO folders, contain actual CO amounts

## Sage Cost Code Taxonomy
- **Labor codes**: 011, 100, 101, 110, 111, 112, 113, 120, 130, 140, 141, 142, 143, 145, 150
- **Material codes**: 039, 210, 211, 212, 213, 220, 230, 240, 241, 242, 243, 244, 245
- **Overhead codes**: 600, 601, 602, 603, 604, 607
- **Burden codes**: 995, 998
- **Sales/Revenue**: 999 (skip in cost analysis)

Key codes: 120 = Roughin Labor (always largest), 230 = Finish Material, 995 = Payroll Burden, 998 = Payroll Taxes

## Common Pitfalls & Fixes (learned from experience)

### AR Invoice Regex
The JDR parser regex for invoices must use `(?:\s+([\d,.\-]+))?` for the optional retainage group — NOT `\s+([\d,.\-]+)?`. The `\s+` must be inside the optional group, otherwise lines without retainage fail to match.

### Template Cloning Errors
When creating a new project by copying from an existing one, ALWAYS check for:
- Wrong unit counts (per-unit calculations)
- Wrong GC names
- Wrong project names/addresses
- Empty CO lists that need real data
- Wrong document counts (files, permits, POs, submittals, RFIs)
- Placeholder CO amounts that don't match Budget Transfer spreadsheets

### BVA Format
The dashboard JS expects 8-column arrays: `[label, budget, actual, variance, pct, null, null, status]`
Some older projects (2029, 2030) have 9-column format with orig+rev budget split — this shifts column indices.
Projects 2027 and 2028 have object format `{origBudget, revBudget, actual, variance, pctVar}` — doesn't render in the BVA tab.

### Deployment
- GitHub repo: `steventyoo/cortex-mockup`
- Branch: `main`
- Vercel auto-deploys on push to main
- Just `git add index.html && git commit && git push origin main`

## Folder Structure
```
cortex-mockup/           ← GitHub repo, Vercel source
├── index.html           ← THE dashboard (single file, ~9000 lines)
├── owp/                 ← Per-project folders (cortex output files inside)
│   ├── owp-2001/
│   ├── owp-2008/
│   └── ...
├── vercel.json
└── .vercel/

owp-XXXX/                ← Google Drive mounted folders (one per project)
├── XXXX Job Detail Report.pdf
├── XXXX-GC, Project Name/    ← GDrive subfolder with all project docs
│   ├── CO's/
│   ├── RFI's/
│   ├── Submittals/
│   ├── Permits/
│   ├── PO's/
│   └── ...
├── cortex output files/      ← Generated outputs
│   ├── parse_XXXX.py
│   ├── build_XXXX.py
│   ├── notionize_XXXX.py
│   ├── XXXX_data.json
│   ├── OWP_XXXX_JCR_Cortex_v2.xlsx
│   └── OWP_XXXX_JCR_Summary_Notion_v2.xlsx
└── cortex reference files/   ← Templates and reference scripts
```

## Current Status (as of 2026-04-17)
- All 32 projects fully wired: 18/18 data blocks each
- **2035-2040** were the most recent batch added (6 projects, via `gen2035/` generator pipeline, source JDRs from "Gemini Accounting Reports" GDrive folder)
- **Two new GCs** introduced with this batch: **Marpac** (via #2037) and **Shelter Holdings** (via #2039). Atlas SVG updated with both new GC rects at (1540,74) and (196,74) respectively.
- 2035 (Natural & Built, 162 Ten, Redmond, 92 units, $722,900, 49.7% margin) — $359,133 net profit, 24 workers / 5,205 hrs, retainage $36,115 outstanding 7+ yrs; additive CO profile (+$5,500 / +0.8%)
- 2036 (Exxel Pacific, Westridge, 31 units, $418,677, 34.0% margin) — $142,468 net profit, 30 workers / 4,098 hrs, +$30,162 COs (+7.8%)
- 2037 (Marpac, University Apartments (MYSA), 122 units, $2,076,783, 41.4% margin) — $859,548 net profit, 50 workers / 14,799 hrs, +$155,933 COs (+8.1%) — largest of the new batch, NEW GC
- 2038 (Compass Harbor, 2nd & John, 80 units, $1,195,660, 32.5% margin) — $388,347 net profit, 36 workers / 11,814 hrs, +$20,282 COs (+1.7%); third Compass Harbor job (with 2033, 2034)
- 2039 (Shelter Holdings, Ravello Gas Piping, HOA remediation, $54,900, 64.4% margin) — $35,329 net profit, 12 workers / 238 hrs; small HOA gas piping service job, NEW GC, 0 COs
- 2040 (Blueprint, Brooklyn 65, 55 units, $624,171, 35.2% margin) — $219,653 net profit, 32 workers / 5,454 hrs, +$20,171 COs (+3.3%); second Blueprint job (with 2032)
- 2034 (Compass Harbor, Park Lane Apts, Kirkland, 128 units, $1.90M, 32.4% margin) — $615,703 net profit, OCIP (Wrap) insurance via Compass developer, 13 executed COs net +$58,651 ADDITIVE, retainage $92,133 outstanding 7+ years; 45 workers / 21,638 hrs; Thomas McCabe top worker (1,992 hrs); Rosen Park Lane top vendor (44% / $179k / 121 invoices)
- 2034 Project Team: GC PMs/Sup/PE TBD, OWP RI Foreman Thaddeus Waites, Owner Kirkland Main Street LP
- 2034 critical BVA flags: Garage Labor 111 (+70%), Water Main/Insulation 141 (+61%), Condensation Drains 143 (+718%, emerged via CO#02b), Gas Labor 140 (+231%); offset by material procurement sweep (Roughin Material 220 -50%, Mech Room 242 -68%)
- 2033 (Compass Harbor, Vuecrest, 137 units, $2.34M, 43.1% margin) — $1,010,508 net profit, OCIP (Wrap) via Continental developer, 18 executed COs net -$161,412, retainage $114,309 outstanding 7+ years
- 2033 Project Team: GC PM Ryan Ames/Vince Dennison, GC Sup Todd Mortenson, GC PE Dan Peck, OWP RI Foreman Victor
- Known issue: 2027 and 2028 have object-format BVA instead of array format (BVA tab won't render correctly for those two)
- No projects 2013 or 2014 exist (IDs were never used)
- Atlas SVG now shows **50 nodes / 78 edges** (was 42/64) — 6 new project circles at (340,210), (1190,210), (1600,210), (1340,600), (130,210), (1220,380)

## Notion Integration
The `notionize_XXXX.py` scripts generate summary workbooks styled with Notion design tokens. These use specific color palettes, typography, and layout conventions. The Notion summaries are 2-tab workbooks meant for quick stakeholder review.

## Key Metrics Steven Cares About
- Margin profile (profit %, profit/unit)
- Vendor concentration (top vendor % of total AP)
- Roughin labor share (% of total labor hours)
- Budget variance by cost code (especially CRITICAL flags)
- CO reconciliation (documented vs JDR-implied gap)
- Fixture counts and density (fixtures/unit)
- Crew productivity (hours/unit, revenue/hour)
