# Cortex JCR Portfolio Dashboard — Project Knowledge Base

> This file is read by Claude at the start of every session. Keep it current.
> Last updated: 2026-04-25

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

## Closed projects added since the table below (last updated 2026-04-25)

Beyond the 32-project table that follows, **closed projects 2041–2051 have been added** to the dashboard (PROJECT_ORDER now contains 42 IDs). Most recent addition:

- **#2051 Compass, Vail Apartments** (**Shoreline WA — NOT Bellevue area** per executed contract 18-0622-0500; site 17962 Midvale Ave N; 163 units, $2,588,280 final / $2,672,000 original / **-$83,720 net COs (-3.1%)** — unusual credit-net CO posture; 50.0% gross margin = **highest in OWP's closed Compass portfolio** vs 43.1% Vuecrest, 32.4% Park Lane, 32.5% 2nd & John, 46.5% 2100 Madison; $1,293,193 net profit; 45 workers / 19,549 hrs / 25 cost codes; OWP's **earliest closed Compass job** by start date (Mar 2018 → 2023, ~22 mo field + retention tail); GC PM Justin Anderson, GC Sup Will Fenton, GC PE Jeff Seeb, GC contract signatory Ryan Ames VP, OWP RI Foreman Bob, OWP signatory Richard Donelson VP, Owner AAA Management LLC dba ADC Ridge at Sun Valley (San Diego CA), MEP Franklin Engineering, Insurance Not Wrap; top vendors Ferguson #3007 17.7% / Keller 17.5% / Consolidated 15.3% / Rosen Kirkland 15.2% / California Hydronics 12.7% — top 5 = 78% of $453k AP across 13 vendors / 252 invoices; **retention $126,114 still held 7+ yrs post-completion** per JDR snapshot 04/03/2026 — same aged-retention pattern as Compass siblings #2033 / #2034; 17 SCOs + 7 CORs documented in GDrive Change Orders folder). Pipeline: parse_2051.py + build_2051.py + notionize_2051.py + **enrich_2051.py** in `owp/owp-2051/cortex output files/`. **dashboardize.py was patched** to try both `owp-XXXX-live` and `owp-XXXX` folder paths. Output: 17-tab JCR (60+ key-player rows on Job Info tab after enrich) + 2-tab Notion summary, zero formula errors.

- **#2052 Farrell-McKenna, Fifth & Roy** (Lower Queen Anne Seattle, 701 5th Ave N, 107-unit mixed-use multifamily + Blarney Stone Irish pub retail; $2,231,035 final / $2,134,000 original / **+$97,035 net COs (+4.5%)** — healthy additive posture; **45.4% gross margin** — strong margin profile, well above OWP closed-portfolio median ~34%; $1,013,824 net profit; 39 workers / 15,402 hrs / 27 cost codes; Feb 2018 → Dec 2019 (~22 months); Subcontract #22-0000 (GC Project #17-100) executed Feb 9, 2018; **OWP's first and only Farrell-McKenna engagement**; GC Farrell-McKenna Construction LLC (Burien WA, 17786 Des Moines Memorial Dr); Owner Fifth North & Roy LLC (same address — sister entity); OWP contract signatory Michael Donelson; full design team — Architect **Hewitt Architects**, Civil **KPFF Consulting**, Structural **Bykonen Carter Quinn**, MEP **Rushing Engineering**; Bond NOT Required; top vendors Rosen Kirkland 22.7% / Keller 20.3% / Consolidated 15.8% / California Hydronics 11.3% / Ferguson #3007 11.2% — top 5 = 81.3% of $485k AP across 322 invoices / 20 vendors; **retention $111,552 still held 6+ yrs post-completion**; only BVA flag was 141 Water Main/Insulation Labor +136% — same code that runs hot on most OWP projects; CO #2 added duplex booster, RFI #009 foundation wall collector, RFI #210 upper roof gutters; sub coordination meetings + Daily Reports in GDrive). Pipeline: parse_2052.py + build_2052.py + notionize_2052.py + enrich_2052.py in `owp/owp-2052/cortex output files/`. Farrell-McKenna becomes the 12th GC in the closed-portfolio dashboard. Output: 17-tab JCR (77-row Job Info tab after enrich) + 2-tab Notion summary, zero formula errors. **PROJECT_ORDER now contains 43 IDs** (was 42).

## Closed projects 2053-2060 (added 2026-04-25 batch)

Eight closed projects added in one batch. Each: parse → build → notionize → wire to dashboard → enrich Job Info → fixture schedule. **PROJECT_ORDER now contains 51 IDs** (was 43). All zero formula errors verified.

- **#2053 Compass Harbor, Bellevue Parkside (BLU Apartments)** (Bellevue WA · 75 102nd Ave NE · 135 units · 1,579 itemized fixtures · $2,332,680 final / $2,241,309 orig / +$91,371 COs +4.1% / 40.0% margin · $113,684 retention · OCIP/Wrap-up · Subcontract #22-0500 effective May 23, 2018 · GC Compass Harbor (Kirkland WA, signatory Ryan Ames VP) · Owner Bellevue Parkside LP · Architect Encore Architects · MEP Franklin Engineering · 11 SCOs + 10 CORs + 50 RFIs + 3 ASIs · 167 toilets + 167 heated seats Bemis H1900NL · 5 HTP Phoenix PH199-119 boilers · key feats: dog wash backflow, retail grease trap, 4 yard hydrants, 167+136 condensate stacks).
- **#2054 Blueprint, 800 5th Apartments** (Seattle WA · 800 5th Ave N · 68 units Live/Work · 487 itemized fixtures · $843,087 final / $785,600 orig / +$57,487 COs +7.3% / 14.5% margin — small project, lower margin · $0 retention released · Subcontract #03-08-6505 effective Feb 5, 2018 · GC Blueprint Capital Services LLC (Seattle PO Box 16309) · Owner Blueprint 800 LLC · matte black trim throughout — Delta Trinsic 559HA-BL-DST + Brizo 85521 handshowers).
- **#2055 Blueprint, Dockside Apartments** (Greenlake Seattle WA · 6860 E Green Lake Way N · 98 units · 515 itemized fixtures · $1,315,429 final / $1,142,340 orig / +$173,089 COs +15.2% / 25.8% margin · $0 retention released · GC Blueprint · matte black trim · 4 Bock OptiTherm OT-199N water heaters · Schier GB250 grease interceptor for retail).
- **#2056 Exxel Pacific, Acme Farms** (Central District Seattle WA · 1029 S Jackson Street · **321 units — large project** · 1,345 itemized fixtures · $6,034,876 final / $5,877,400 orig / +$157,476 COs +2.7% / **43.8% margin** — strong on a large bid · $296,894 retention held · Subcontract L1.220000 effective April 18, 2018 · GC Exxel Pacific · features GROOMER'S BEST 58" ADA WALK-THROUGH dog wash tub · Elkay LZWSSM bottle filler with ECH8 chiller in fitness · Mustee mop sink + T&S Brass bike room faucet · pressure-assist Kohler K-3519 retail toilets).
- **#2057 Natural & Built, Plaza Apartments** (Kirkland WA · 330 4th Street · **111 units (10 full apartments + 101 SRO bedrooms with private baths + 6 shared common-area kitchens)** — initially mis-counted as "10 units" because the parse step only captured full-kitchen apartments and missed the 101 SRO bedrooms; corrected 2026-04-26 batch · 350 verified fixtures (354 itemized first-pass) · 3.15 fixtures/unit · $10,216 revenue/unit (sane post-correction; was showing $113k as 10-unit outlier) · $1,133,967 final / $1,098,897 orig / +$35,070 COs +3.2% / **49.4% margin** · $56,698 retention · Owner Kirkland Sustainable Investments · GC Natural & Built · 101 Niagara N7716 0.8gpf SRO toilets + 96 Dayton D11516 bar sinks + 10 full-kitchen units · CO #10 added water meters at studio units · contract not in GDrive folder).
- **#2058 Blueprint, Howell Apartments** (Capitol Hill Seattle WA · 600 E Howell Street · 76 units (+2 added Live/Work via CO #07) · 464 itemized fixtures · $961,556 final / $928,800 orig / +$32,756 COs +3.5% / 27.3% margin · $0 retention released · GC Blueprint · CO #07 converted retail to Live/Work units adding 2 ADA bathroom packages · matte black trim throughout · contract not in GDrive folder).
- **#2059 GRE Construction, Meeker** (Kent WA · 2030 West Meeker Street · 107 units Building A — multi-building project · 773 itemized fixtures · $4,492,453 final / $4,378,750 orig / +$113,703 COs +2.6% / 32.9% margin · $224,623 retention · MasterSubcontract 78472 effective Feb 27, 2019 · GC GRE Construction LLC · per-unit electric water heaters Ruud ProE50-T2 50-gallon · 33 Moen ARC-4200 disposal air switches · Eemax on-demand WHs at amenities · Building A ticket only in this enrichment — Building B+ in subsequent revisions).
- **#2060 Compass, Marina Square (MSQ)** (Bremerton WA · 280 Washington Ave · **254 units across 2 towers (Tower 1 = 129 apt + Tower 2 = 125 hotel→extended-stay conversion)** · 1,176 itemized fixtures · $6,156,577 final / $5,307,510 orig / **+$849,067 COs +16.0%** — major scope additions · **58.2% margin — OWP's highest margin on a closed-portfolio job over $5M** · $300,554 retention · Subcontract MSQ + SCO #2 Full Subcontract for Tower 2 conversion · GC Compass Construction · 2019 → Jul 2022 (~36 months) · ADA Delta items: 59424-18-PK handshowers + 41836 wall bars + 50560 wall elbows · Tower 1 Amerisink AS3337 SS sinks + Tower 2 Elkay ECTRU12179T undercounter SS · Kohler K-3609 toilets in Tower 2 hotel-converted units).

Compass becomes OWP's largest GC by closed-job count — now 9 closed Compass jobs across the portfolio (#2033 Vuecrest, #2034 Park Lane, #2038 2nd & John, #2050 2100 Madison, #2051 Vail, #2053 Bellevue Parkside, #2060 MSQ + 2 future). Blueprint count is now 5 (#2032 Luna, #2040 Brooklyn 65, #2054 800 5th, #2055 Dockside, #2058 Howell). GRE Construction becomes the 13th closed-portfolio GC (joining via #2059 Meeker). Pipeline pattern: each project has parse_XXXX.py, build_XXXX.py, notionize_XXXX.py in `owp/owp-XXXX/cortex output files/`. All META updated with project-specific values. Single batch enrichment via `/tmp/enrich_2053_2060.py` patched 02 Job Info tabs with 9-section identity/team/insurance/contract/CO/document grids + verified fixture schedules.

## #2061 Exxel Pacific, Alta Columbia City (real closed project — NOT a backtest)

- **#2061 Exxel Pacific, Alta Columbia City** (3717 South Alaska Street, Seattle WA 98118 · Columbia City · **243-unit Mixed-use multifamily + 6 retail spaces** · 7-story over below-grade garage · Pavilion + Skylounge amenity · 1,093 trim fixtures (4.50/unit per Job Ticket) · $4,430,088 final / $4,305,550 original lump sum / **+$124,538 net COs (+2.9%)** — modest additive posture · **41.4% gross margin** — exactly matches the OWP house median (41.4%) so this job is the canonical "median" closed project · $1,832,588 net profit · 58 workers / 28,516 hrs / 26 cost codes · **Code 120 Roughin Labor 14,973 hrs = 53% of all labor** · Jun 2019 → May 2022 (~35 months) · Subcontract **L1.220000 executed Aug 8, 2019** (AGC Washington 2009 form modified) · GC Exxel Pacific (PM Brian Christensen) · Owner **Gateway Alta Rainer Owner LLC** · Architect **Johnston Architects PLLC** · MEP **Franklin Engineering** · Insurance OCIP (Wrap-up) — Exxel-administered · OWP signatory Richard Donelson · OWP License ONEWAWP895BU · top vendors **Rosen Supply Kirkland 37.9% / 223 invoices** · Keller 22.2% · Consolidated 8.7% · California Hydronics 7.5% · Franklin Engineering 6.9% — top 5 = 83.2% of $1.0M+ AP across 23 vendors · 39 RFIs + 26 ASIs + 13 SCOs + 14 COR series resolved · **retention $221,504 released 2022-06-23** (clean closeout — no aged retention) · 65 O&M close-out documents on file · serves as the **anchor case study for the bid-accuracy backtest** (`2061_Bid_Accuracy_Test.xlsx`, `2061_Bid_Intelligence_Case_Study.xlsx`) — the bid tool's $1,837/fx default is calibrated TO this project, hence the case-study files; the project itself is a real, fully-billed, fully-closed Exxel job).

#2061 is the 52nd ID in PROJECT_ORDER and brings the closed-portfolio count to 52 (was 51). Pipeline files: `owp/owp-2061/cortex output files/` (parse_2061.py reads the 171-page JDR PDF from `owp/Gemini Accounting Reports/2061 Job Detail Report.pdf`, schema matches 2020 gold standard + 2049 v2.2 extensions). Output: 17-tab JCR + 2-tab Notion summary, zero formula errors. **Important context:** because the bid intelligence tool was calibrated against #2061's actual cost profile (1,093 fx · $1,837/fx total · 41.4% margin), this project shows up in the tool as the canonical "what good looks like" example. The case-study Excel files are derivative artifacts, not the source of truth — the source is the JDR pipeline same as every other closed job.

## #2063 Compass Aria Flats (added 2026-04-29 — skipped #2062)

- **#2063 Compass General Construction, Aria Flats** (Redmond WA · 16760 Redmond Way · 186-unit mixed-use multifamily + retail · 5–7 story · internal name "Redmond Flats GMP" · 1,513 itemized fixtures · $1,653,409 final / $1,653,409 original / **18 executed COs net $0** — heavy add-deduct activity, fully balanced · **32.6% gross margin** · $539,146 net profit · $1,114,263 direct cost · 29 workers / 13,264 hrs / 27 cost codes · Oct 2019 → Oct 2022 (~36 months) · Subcontract **Compass 22-0500-Aria** · GC Compass General Construction I, LLC · MEP Franklin Engineering · OWP PM Richard Donelson · Insurance OCIP (Wrap-up) — Compass developer · top vendors **Rosen Supply Kirkland 42.9% / 216 invoices / $151,872** — single-vendor concentration approaching the 35% threshold · 41 RFIs + 100 submittals + 30 total change events · **retention $79,842 still held** per JDR snapshot · notable systems: 3 Bock 125-gal condensing gas WHs, retail grease interceptor, duplex booster pump, 14 irrigation POCs).

**Skipped #2062 Marpac 206 Place (Origin 206)** — design-only engagement where the construction scope was transferred to JJ Plumbing on 2022-08-02. OWP only billed $50,927 in design fees and recorded a $10,599 net loss on the design phase. Pipeline files exist (`owp/owp-2062/cortex output files/`) but the project was deliberately not added to the dashboard because (a) it's not a representative closed job, (b) the small contract + loss would dilute the benchmark table's "what good looks like" signal, and (c) the JJ Plumbing handoff makes it an outlier even if wired.

#2063 is the 53rd ID in PROJECT_ORDER and brings the closed-portfolio count to 53. Compass becomes OWP's largest GC by closed-job count — now **10 closed Compass jobs** (#2033 Vuecrest, #2034 Park Lane, #2038 2nd & John, #2050 2100 Madison, #2051 Vail, #2053 Bellevue Parkside, #2060 MSQ, #2063 Aria Flats + 2 future). Pipeline: parse_2063.py + build_2063.py + notionize_2063.py + audit_2063.py in `owp/owp-2063/cortex output files/`. **Schema note:** 2063_data.json uses the v2.2 schema (`cost_codes`, `workers`, `vendors` as top-level keys) which differs from `dashboardize.py`'s expected v2.0 schema (`cost_code_summaries`, `worker_wages`); a one-shot adapter script `/tmp/wire_2063.py` was used to generate the JS data arrays. Future closed jobs with the v2.2 schema will need either dashboardize.py patched to handle both schemas, or another one-shot adapter. Output: 17-tab JCR + 2-tab Notion summary, zero formula errors.

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
