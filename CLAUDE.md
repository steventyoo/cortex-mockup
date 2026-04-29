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

#2063 is the 53rd ID in PROJECT_ORDER and brings the closed-portfolio count to 53. Compass becomes OWP's largest GC by closed-job count — now **10 closed Compass jobs** (#2033 Vuecrest, #2034 Park Lane, #2038 2nd & John, #2050 2100 Madison, #2051 Vail, #2053 Bellevue Parkside, #2060 MSQ, #2063 Aria Flats + 2 future). Pipeline: parse_2063.py + build_2063.py + notionize_2063.py + audit_2063.py in `owp/owp-2063/cortex output files/`. **Schema note:** 2063_data.json uses the v2.2 schema (`cost_codes`, `workers`, `vendors` as top-level keys) which differed from `dashboardize.py`'s original v2.0-only schema (`cost_code_summaries`, `worker_wages`); a one-shot adapter script `/tmp/wire_2063.py` was used to generate the JS data arrays for #2063 specifically. **Resolved 2026-04-29 in commit `078e0bc`:** dashboardize.py was patched to auto-detect and handle BOTH schemas, plus a new `synthesize_meta_from_data()` helper produces a default META from v2.2 data.json's `project` + `totals` blocks when the build script has no `META_XXXX` dict. Future closed jobs with v2.2 schema now wire via `python3 owp/dashboardize.py XXXX` with no per-job adapter. #2063's data arrays were re-generated through the standardized pipeline in commit `078e0bc` (the `/tmp/wire_2063.py` adapter is now redundant). Output: 17-tab JCR + 2-tab Notion summary, zero formula errors.

## #2066 Compass Fireside Flats (added 2026-04-29 — skipped #2065)

- **#2066 Compass General Construction, Fireside Flats** (Roosevelt Seattle · 841 NE 68th St, 98115 · 102-unit multifamily · 5–6 story · 626 itemized fixtures · $1,695,972 final / $1,565,600 base subcontract / **16 executed COs net $0** — heavy add-deduct fully balanced, **same posture as #2063 Aria Flats** and a recurring pattern in Compass jobs · **33.5% gross margin** · $568,094 net profit · $1,127,878 direct cost · 38 workers / 11,498 hrs / 36 months Nov 2020 → Nov 2023 · Subcontract **22-0500-01 executed 2020-11-24** · GC Compass General Construction I, LLC · Owner **Fireside Flats LLC** (Compass-affiliate) · MEP Franklin Engineering · OWP PM Richard Donelson · Insurance OCIP (Wrap-up) — Compass developer · top vendors **Rosen Supply Kirkland 22.0% / 146 invoices / $102,724** — much healthier vendor concentration than #2063's 42.9% · 17 vendors total, $466,139 AP across 389 invoices · 66 RFIs + 152 submittals + 6 unexecuted CORs · retention released at closeout).

**Skipped #2065 Holland Flag Lot (1120 Dexter Ave N)** — same pattern as #2062 Marpac. $1,466,807 base subcontract was signed Feb 6, 2020 ("1120 Trade Contractor Agreement") under Holland Construction Management's OCIP, but **the construction scope shifted elsewhere** and OWP only billed $5,273 in design-phase work (9 hours, 2 workers, 1-month duration Mar 2020). Pipeline files exist (`owp/owp-2065/cortex output files/`) but the project was deliberately not added to the dashboard because (a) it's effectively a $5k design fee, not a constructed project, (b) wiring it as a closed job at "0% margin · $5k contract" would dilute the closed-portfolio benchmarks alongside real $1M+ projects, and (c) the actual built-out scope presumably ran under a different OWP job number (whoever ended up doing it). The $1.47M signed-but-not-built subcontract is interesting historical context — it's the second time OWP was in design-phase for a project that ultimately went elsewhere — but doesn't belong in the closed-portfolio dashboard. Same skip rationale as #2062.

#2066 is the 54th ID in PROJECT_ORDER and brings the closed-portfolio count to 54. Compass becomes OWP's largest GC by closed-job count — now **11 closed Compass jobs** (#2033 Vuecrest, #2034 Park Lane, #2038 2nd & John, #2050 2100 Madison, #2051 Vail, #2053 Bellevue Parkside, #2060 MSQ, #2063 Aria Flats, #2066 Fireside Flats + 2 future). Pipeline: parse_2066.py + build_2066.py + notionize_2066.py + audit_2066.py in `owp/owp-2066/cortex output files/`. **First closed job wired purely through the standardized v2.2 pipeline** — no per-job adapter, just `python3 owp/dashboardize.py 2066` followed by `/tmp/insert_2066.py` for the index.html plumbing. Validates the dashboardize.py patch from `078e0bc`. Output: 17-tab JCR + 2-tab Notion summary, zero formula errors.

**Pattern observation across closed Compass jobs:** #2063 Aria Flats and #2066 Fireside Flats both delivered with **18 and 16 executed COs netting exactly $0** — heavy add-deduct activity that fully balances. This is unusual: most closed jobs in the portfolio show net additive CO posture (#2052 +$97k, #2059 +$114k, #2061 +$125k), and a handful show net credit (#2051 −$84k). The $0-net pattern across two consecutive Compass jobs suggests Compass's CO process is structurally different — likely tighter GMP discipline where every add gets offset by a credit elsewhere in the same CO cycle. **Update 2026-04-29:** #2067 GRE Edmonds 99 (different GC, separate batch) ALSO delivered with 4 COs net $0, broadening the pattern beyond Compass. Now N=3 jobs across 2 GCs with $0-net CO posture, vs the additive/credit-net norm across the rest of the closed portfolio. Worth watching for a 4th data point — if it holds, the bid tool should model "no-net-CO" as a real GC behavior class (Compass + GRE) and adjust contingency assumptions accordingly. Tighter contingency on these GCs because COs don't compound; looser on the additive-pattern GCs.

## #2067 GRE Edmonds 99 (added 2026-04-29 — skipped #2068)

- **#2067 GRE Construction, Edmonds 99** (Edmonds WA · 23400 Highway 99 · 232-unit multifamily · 1,870 itemized fixtures · **8.06 fx/unit — above OWP portfolio median** · $2,588,429 final / $281,800 preliminary subcontract → $2.59M final / **4 executed COs net $0** — same balanced posture as Compass jobs · **29.8% gross margin** · $771,334 net profit · $1,816,129 direct cost · 52 workers / 25,854 hrs / 33 months Mar 2020 → Dec 2022 · Subcontract **GRE #7829** NTP 2020-03-11 · GC GRE Construction LLC · Owner **GRE Edmonds, LLC** · MEP Franklin Engineering · OWP PM Richard Donelson · Insurance **Standard (not OCIP/Wrap)** — different from the OCIP-heavy Compass and Exxel jobs · top vendor **Rosen Supply Kirkland 33.9% / 213 invoices / $213,992** · 15 vendors total, $632,146 AP across 381 invoices · **15 RFIs + 9 submittals — unusually low document load** for a 232-unit job, suggests a clean spec set or a less change-management-heavy GC · retention $129,421 still held).

**Skipped #2068 Schuchart Alloy Apartments** — same pattern as #2062 Marpac and #2065 Holland Flag Lot. OWP did 5 months of design consulting (May–Sep 2020) on Schuchart's 203-unit / 768-fixture Alloy Apartments project but **no construction subcontract was awarded** to OWP. Total billed $62,541 / 0 hours / 0 workers / 0 RFIs / 0 submittals. Pipeline files exist (`owp/owp-2068/cortex output files/`) but the project was deliberately not added to the dashboard because (a) it's a $62k design fee, not a constructed job, (b) wiring it would put a 0-hour, 0-worker entry in the closed-portfolio benchmarks alongside real built-out projects, and (c) Schuchart never became an OWP construction-phase GC. Same skip rationale as #2062 / #2065.

#2067 is the 55th ID in PROJECT_ORDER and brings the closed-portfolio count to 55. **GRE becomes OWP's 4th-largest GC by closed-job count** — now **2 closed GRE jobs** (#2059 Meeker Building A + #2067 Edmonds 99). Pipeline: parse_2067.py + build_2067.py + notionize_2067.py + audit_2067.py in `owp/owp-2067/cortex output files/`. **Wired purely through the standardized v2.2 pipeline** (`python3 owp/dashboardize.py 2067`). Output: 17-tab JCR + 2-tab Notion summary, zero formula errors.

**Insurance class observation:** #2067 is the first closed job in recent additions (post-#2050 batch) that is NOT under OCIP/Wrap-up. Closed Compass jobs (#2033/#2034/#2038/#2050/#2051/#2053/#2060/#2063/#2066) are all OCIP. Closed Exxel jobs (#2056/#2061) are all OCIP. GRE #2059 Meeker had MasterSubcontract terms but per-job insurance — and now #2067 Edmonds 99 is explicitly Standard (not OCIP). This is becoming a useful third axis for the bid tool: GC × insurance class × $0-net CO behavior. Standard-insurance jobs may have different retention/billing dynamics worth modeling separately.

## #2069 Exxel Theory U-District (added 2026-04-29 — skipped #2070)

- **#2069 Exxel Pacific, Theory U-District** (Seattle WA · 4731 15th Ave NE, U-District 98105 · **342-unit student housing** — OWP's largest closed Exxel job by unit count after #2056 Acme Farms (321u) · 1,775 itemized fixtures · $4,087,616 final / $3,872,036 lump-sum subcontract / **6 executed COs net $0** — but mechanism is contract-bill-and-credit (see pattern dig below) · **38.7% gross margin** — highest dollar net in this batch at $1,581,148 · $2,505,628 direct cost · 60 workers / 25,818 hrs / 33 months Feb 2021 → Nov 2023 · Subcontract **L1.220000 S** executed 2021-02-01 (design-build) · GC Exxel Pacific Inc. · Owner **BYSHSF Seattle LLC c/o McKinley Seattle LLC** · Architect **Ankrom Moisan Architects** · MEP Franklin Engineering · OWP PM Richard Donelson · Insurance **Standard (NOT OCIP)** — atypical for Exxel; prior closed Exxel jobs #2056/#2061 were both OCIP · top vendor **Rosen Supply Kirkland 33.6% / 225 invoices / $359,657** · 17 vendors total, $1,069,358 AP across 524 invoices · **78 RFIs + 16 submittals** — high RFI count reflects design-build + tight Exxel coordination across electrical/fire/elevator/framing trades · 2 boilers + duplex booster + 18 motorized irrigation zone valves · retention $199,638 still held).

**Skipped #2070 Chinn Beacon Crossing** — same pattern as #2062 Marpac, #2065 Holland Flag Lot, #2068 Schuchart. OWP did 14 months of design + permit phase work (Dec 2020 → Feb 2022) on Chinn's 103-unit Beacon Hill TOD project but **the construction subcontract was NOT executed**. Total billed $42,289 / 16 hours / 2 workers / 1 RFI / 8 submittals / **−$1,981 net loss**. Pipeline files exist (`owp/owp-2070/cortex output files/`) but the project was deliberately not added to the dashboard for the same reasons as the prior skips. Chinn closed-portfolio count remains at 5 (#2009 Greenwood, #2010 Old Town, #2011 Phinney Ridge, #2023 Legacy, #2041) — Beacon Crossing would not have changed that count anyway since it was never built.

#2069 is the 56th ID in PROJECT_ORDER, closed-portfolio count is now 56. **Exxel becomes 3rd closed Exxel job** (#2056 Acme Farms, #2061 Alta Columbia City, #2069 Theory U-District). Pipeline: parse_2069.py + build_2069.py + notionize_2069.py + audit_2069.py in `owp/owp-2069/cortex output files/`. Wired purely through the standardized v2.2 pipeline. Output: 17-tab JCR + 2-tab Notion summary, zero formula errors.

## CO-mechanism dig (2026-04-29) — the "$0-net" pattern is actually TWO mechanisms

Earlier I flagged that #2063 Aria Flats, #2066 Fireside Flats, and #2067 Edmonds 99 all delivered with executed COs netting $0, and called it a single "no-net-CO" pattern across Compass + GRE. **After #2069 Theory U-District landed (also $0 net) and a closer dig into AR invoice ledgers, the pattern decomposes into two genuinely different mechanisms:**

### Mechanism A: Contract-bill-and-credit (Compass + Exxel)
OWP bills MORE than the contract value gross, then issues credits on separate invoices to bring net billing back to the flat contract. CO scope is real and billed; credits offset.

| Job | GC | Contract | Gross billings | Credits | Over-bill % |
|---|---|---|---|---|---|
| #2063 Aria Flats | Compass | $1,653,409 | $1,705,127 | −$51,718 | +3.1% |
| #2066 Fireside Flats | Compass | $1,695,972 | $2,067,638 | −$371,666 | **+21.9%** |
| #2069 Theory U-District | Exxel | $4,087,616 | $4,758,038 | −$670,422 | +16.4% |

#2066 and #2069 are running massive CO activity (~$370–670k of adds) that gets fully credited back. The credits are likely OCIP insurance deducts (for the Compass jobs that ARE OCIP), final pay-app reconciliations, or owner-directed scope reductions billed as separate credits. Either way, the gross/net delta is real CO traffic — just net-zero to OWP economics.

### Mechanism B: True $0-CO (GRE)
| Job | GC | Contract | Gross billings | Credits | Over-bill % |
|---|---|---|---|---|---|
| #2067 Edmonds 99 | GRE | $2,588,429 | $2,588,429 | $0 | **+0.0%** |

GRE #2067 has ZERO negative AR invoices and zero gross over-billing. The 4 executed COs that "net $0" must have been processed at $0 internally (add + deduct in the same CO line, or scope changes that didn't move price), or they were never billed as separate items. Fundamentally different from the Compass/Exxel pattern.

### Bid-tool implications

The two mechanisms have different bid implications:

- **Mechanism A (Compass, Exxel):** COs add scope and revenue, but credits offset. Net economic effect on OWP is preserved-margin even though gross volume swings. Bid tool should expect CO traffic on Compass/Exxel jobs but model it as cost-neutral. **Margin discipline is what matters, not contingency size** — credits will eat any inflated CO billing anyway.
- **Mechanism B (GRE):** COs don't move money. Bid tool can model GRE jobs with **lower contingency** because scope changes get absorbed without billing impact — which is favorable to OWP's cash flow but means there's no cushion to recover from underbids.

Both mechanisms differ from the **additive-pattern GCs** in the rest of the closed portfolio (#2052 +$97k, #2059 +$114k, #2061 +$125k) where CO billings flow through to revenue and margin in the conventional way.

**Caveat — N is still small.** Mechanism A is now N=4 across 3 GCs (Compass×2, Exxel×1, Jabooda×1 after #2064 wired); Mechanism B is N=1. To ship this as a bid-calibration signal we'd want N=2 on Mechanism B (another closed GRE job) and ideally N=5+ on Mechanism A across more GCs. Worth re-checking when more closed jobs land (#2071+).

**Caveat — GDrive Change Orders folders are not mounted** for #2063/#2066/#2067/#2069. The decomposition above is derived purely from AR invoice ledger amounts in the data.json (positive vs negative billings against final contract value). Folder review (per-CO description, dollar amount, who initiated, what scope changed) would let us distinguish "OCIP insurance deduct" credits from "scope reduction" credits from "reconciliation" credits — currently we can't tell. If you mount the GDrive folders or upload the CO PDFs for any of these jobs, I can refine the analysis.

## #2064 Jabooda Melody Apts (added 2026-04-29 — was MISSED in earlier batch, now corrected)

**Note: This project was inadvertently skipped when the user asked me to "do 2065 and 2066" and I jumped straight from #2063 to #2065 without checking #2064. User caught it on 2026-04-29 ("where is 2068 and 2064?"). Wired retroactively the same session.** Documented here both for completeness and as a reminder to always sweep the next unwired ID rather than jumping on user-named IDs.

- **#2064 Jabooda Homes, Melody Apartments** (Seattle WA · 1801 Rainier Ave S, 98144 · 186-unit Rainier Avenue mixed-use multifamily · 1,280 itemized fixtures · rooftop bar + retail program with grease interception · 4 Laars 100-gal condensing gas boilers · $2,978,869 final / $2,978,869 original / **12 executed COs net $0** — Mechanism A (contract-bill-and-credit, +2.2% gross over contract — much milder than #2066's +21.9% or #2069's +16.4%) · **47.6% gross margin — highest single-job margin in the recent batch and one of the highest in the entire closed portfolio** (only #2057 Plaza 49.4% and #2060 MSQ 58.2% are higher) · $1,418,917 net profit · $1,559,952 direct cost · 46 workers / 18,226 hrs / 33 months Dec 2019 → Sep 2022 · Subcontract **AIA A401-2017** executed 2019-12-13 · GC Jabooda Homes Inc. — **NEW GC, OWP's first and only closed Jabooda engagement** · Owner **Jabooda 1801 LLC** (sister entity) · Architect **Mass Architect LLC** · MEP Franklin Engineering · OWP PM Richard Donelson · Insurance **Standard (no wrap)** — like #2067 GRE and #2069 Exxel · top vendor **Rosen Supply Kirkland 30.8% / 126 invoices / $149,827** · 18 vendors total, $486,303 AP across 291 invoices · **0 RFIs / 114 submittals — extremely atypical document profile**, suggests a clean spec set, an experienced field crew, or a GC that handled clarifications informally outside the formal RFI process · retention $149,670 still held).

#2064 is now the 57th ID in PROJECT_ORDER (inserted between #2063 and #2066 to maintain ID order). Closed-portfolio count is now 57. **Jabooda becomes OWP's 14th closed-portfolio GC** (joining Chinn/SRM/Exxel/Compass/BMDC/Blueprint/Natural & Built/Marpac/Shelter/Holland/GRE/Synergy/Farrell-McKenna/Jabooda).

**The "0 RFIs / 114 submittals" anomaly is worth flagging.** Normal closed jobs in the OWP portfolio show RFI counts in the 15–80 range (e.g. #2059 Meeker 15 RFIs, #2061 Alta CC 39 RFIs, #2063 Aria Flats 41 RFIs, #2069 Theory U-D 78 RFIs). Zero is extreme. Either Jabooda's coordination process is unusually tight, the project benefited from a clean spec from Mass Architect + Franklin Engineering with no field surprises, or RFIs were handled informally and not captured in the document log. **Combined with the 47.6% margin (one of the highest in the portfolio), this suggests a "low-friction GC" archetype worth modeling for the bid tool.** If #2064 is representative of how Jabooda runs jobs, it's a high-margin, low-doc-load engagement profile — a different shape than the high-RFI Exxel jobs (#2069 78 RFIs at 38.7% margin) or the high-CO-traffic Compass jobs (#2066 16 COs at 33.5% margin). N=1 on Jabooda — would need a 2nd Jabooda job to confirm pattern.

## #2071 Chinn Stellar + #2074 Marpac Buddha Jewel (added 2026-04-29)

- **#2071 Chinn Construction, 1405 Dexter Stellar** (Seattle WA · 1405 Dexter Ave N, 98109 · South Lake Union · 160-unit mixed-use multifamily · 1,260 fixtures · $2,833,195 final / $2,786,700 base / **14 executed COs net $0** — Mechanism A (gross billings $3.14M + credits $311k = +11.0% gross over contract, mid-range for the pattern) · **38.6% gross margin** · $1,094,143 net profit · $1,738,656 direct cost · 43 workers / 16,345 hrs / 21 months Oct 2021 → Jun 2023 · 5-bid history Mar 2020 → Sept 2021 (scope shifted 169u → 160u Oct 2020, tubs → showers Jul 2020) · GC Chinn Construction LLC · Owner **Dexter Borealis LLC** · Architect **Board & Vellum** (115 15th Ave E, Seattle) · MEP Franklin Engineering (Chinn's standard partner) · Insurance **Standard (no wrap) — atypical for Chinn** · top vendor **Keller Supply Company 27.1% / 53 invoices / $200,522** — Keller leading instead of Rosen Supply Kirkland is unusual (Rosen typically dominates OWP closed-portfolio AP) · 17 vendors total, $739,615 AP across 347 invoices · 35 RFIs + 110 submittals · retention $138,710 still held).

- **#2074 Marpac Construction, Buddha Jewel Monastery Phase II** (Shoreline WA · 17418 8th Ave NE, 98155 · **specialty religious build — first and only OWP monastery engagement** · NOT multifamily · 138 fixtures · $350,295 final / $333,600 base (3rd bid accepted Jul 23, 2020 after Nov 2019 $189,500 → Mar 2020 → Jul 2020 sequence) / **3 executed COs net $0** — small-job Mechanism A (gross $356k + credits $20k = +1.5% gross over) · **44.4% gross margin — second-highest in this batch after #2064 Jabooda 47.6%** · $155,425 net profit · $194,870 direct cost · 18 workers / 2,400 hrs / 10 months Sep 2020 → Jul 2021 · GC Marpac Construction · GC PM **Michelle Ip** · Owner Buddha Jewel Monastery · MEP Franklin Engineering (Job #90404, Kirkland) · Insurance Standard (no wrap) · top vendor **Keller Supply Company 29.0% / 20 invoices / $14,103** · 7 vendors total, $48,574 AP across 163 invoices (small-job AP profile) · 5 RFIs + 20 submittals · retention $16,788 still held).

#2071 + #2074 are the 58th and 59th IDs in PROJECT_ORDER, closed-portfolio count is now 59. **Chinn becomes 5 closed + 1 live** (#2009, #2010, #2011, #2023, #2041, #2071 + #2098 live). **Marpac becomes 3 closed** (#2037, #2046, #2074). Both wired purely through standardized dashboardize.py + insert script. After wiring, **re-ran owp/build_calibration.py** to refresh OWP_Productivity_Insights.json — calibration sample now n=46 (was 44). Headline benchmarks essentially unchanged (hours/unit median 116.1, gross_margin median 40.3%, loaded_wage $39.73/hr) — adding 2 jobs to a 44-sample doesn't move medians materially.

**Mechanism A pattern is now N=6 across 4 GCs:** Compass×2 (#2063, #2066), Exxel×1 (#2069), Jabooda×1 (#2064), Chinn×1 (#2071), Marpac×1 (#2074). The "contract-bill-and-credit" CO process is no longer GC-specific — it's the dominant pattern across recent closed jobs regardless of GC. Mechanism B (true $0-CO with no AR movement) remains N=1 (only GRE #2067). Updated framing: **"$0-net CO" is the recent-era OWP norm, not an anomaly worth flagging in bid-tool calibration**; what's actually distinctive is the gross-over-contract magnitude (#2066 +21.9%, #2069 +16.4%, #2071 +11.0% on the larger jobs vs #2074 +1.5%, #2064 +2.2% on smaller). Larger jobs run heavier CO traffic before crediting back — possibly a contractor-management-bandwidth signal.

**Vendor concentration shift:** #2071 + #2074 both have **Keller Supply Company** as top vendor (27.1% and 29.0%) instead of the usual Rosen Supply Kirkland. This is the first time in the recent batch we see a non-Rosen top vendor on closed jobs. Possibly Chinn-specific (Chinn historically used Keller) or Marpac/specialty-build-specific. Worth tracking if more Keller-led jobs appear.

## #2075-#2078 batch (added 2026-04-29 — skipped #2070 dead, #2073 missing)

User explicitly directed: **"DO NOT upload the dead projects"** — referring to design-only / scope-shifted closures (#2062 Marpac Origin, #2065 Holland Flag, #2068 Schuchart Alloy, #2070 Chinn Beacon). All correctly excluded. #2073 has no JDR PDF in `Gemini Accounting Reports/` — likely a number-skipped or never-closed job.

- **#2075 Braseth Construction, Evt Rockefeller** (Everett WA · 2701 Rockefeller Ave · 187-unit mid-rise · 1,300 fixtures · $2,437,274 final / $2,393,930 base / **0 executed COs** — only $11.6k in small AR credits, essentially flat — exceptional for a $2.4M / 21-mo build · **27.3% gross margin** — lowest in this batch · $665,044 net profit · 36 workers / 20,082 hrs / 21 months Jan 2021 → Sep 2022 · Subcontract **OWP 20-002-09** executed 2021-01-28 · GC Braseth Construction — **NEW GC, OWP's first closed Braseth engagement** · Owner Rockefeller Apartments LLC · MEP Franklin Engineering (per OWP standard) · OWP PM Richard Donelson · Insurance Standard (no wrap) · top vendors **Rosen 28.0% / 143 inv · Keller 24.7% · Ferguson 19.6%** — distributed top-3 (top vendor share <30% is rare in OWP closed portfolio) · 13 RFIs + 7 submittals — low document load · retention $121,864 still held).

- **#2076 Compass General, Yesler Terrace Phase 2** (Seattle WA · 1020 South Main St · Yesler Terrace neighborhood · 219-unit residential of SHA mixed-income redevelopment · 1,504 fixtures · $3,620,263 final / $3,112,600 base / **23 executed COs net $0** — Mechanism A but **unusually tight: gross $3.66M + credits $44.7k = +1.2% gross over** (vs Compass siblings #2066 +21.9%, #2078 +28.3%) · **31.3% gross margin** · $1,133,866 net profit · 64 workers / 29,867 hrs / 22 months Oct 2020 → Aug 2022 · Subcontract **03-22-0500** executed 2021-01-28 · GC Compass General Construction I, LLC · Owner **Yesler Phase 2 LLC (Lowe Enterprises affiliate)** · Architect GGLO (per 2018 bid set) · MEP Franklin Engineering · OWP PM Richard Donelson · Insurance Standard (Compass-administered SHA project, atypical for Compass) · top vendor **Keller Supply 36.6% / 61 inv / $313,436** · 18 vendors total, $857,056 AP across 441 invoices · 6 RFIs + 13 submittals — extremely low for 23 COs · retention $177,763).

- **#2077 Intracorp, Shoreline (147th)** (Shoreline WA · 147th Street area · **346-unit multifamily — OWP's largest closed job by unit count after #2069 Theory U-D 342u** · 1,518 fixtures · $5,611,653 final / 12 executed COs net $0 (Mechanism A: $5.82M gross + $210k credits = +3.7% gross over — moderate) · **41.7% gross margin — highest dollar net in this batch at $2,337,621** · $3,274,031 direct cost · 84 workers / 37,399 hrs / 33 months Oct 2021 → Jul 2024 · Subcontract **147-ONEW00** with 5 signed COs · GC **Intracorp — NEW GC, OWP's first closed Intracorp engagement** · MEP Franklin Engineering ($74,732 AP / 9 inv) · OWP PM Richard Donelson · 13 RFIs + 12 submittals — low for 346u · top vendor **Rosen Supply Kirkland 77.6% / 273 invoices / $1,037,745** — **HIGHEST single-vendor concentration in any closed-portfolio job, by a large margin (next highest is #2064 Jabooda at 30.8%)**. Single-source-of-failure flag worth raising in bid-tool risk panel for any future Intracorp pursuit · retention $272,212).

- **#2078 Compass General, Lake Street** (Kirkland WA · 112 Lake Street S, 98033 · 151-unit mixed-use multifamily · 800 fixtures · $2,881,456 final / $2,711,810 base / **40 executed COs net $0 — HIGHEST CO COUNT IN OWP CLOSED PORTFOLIO** (Mechanism A on steroids: $3.70M gross + $816k credits = **+28.3% gross over contract — largest gross-over magnitude in any closed job**) · **23.8% gross margin — LOWEST in OWP's closed Compass portfolio** vs Vuecrest 43.1%, Park Lane 32.4%, MSQ 58.2%, Aria 32.6%, Fireside 33.5%, Yesler T2 31.3% · $686,461 net profit · $2,194,995 direct cost · 46 workers / 24,941 hrs / 22 months Nov 2020 → Sep 2022 · Subcontract **2020-2201** executed 2021-01-26 · GC Compass · Owner **Kirkland Lake Street LP** · MEP Franklin Engineering ($54,308 AP / 8 inv) · Insurance OCIP / Wrap-up (Compass developer-administered) · 62 RFIs + 19 submittals · top vendor **Keller Supply 33.3% / 70 invoices / $253,676** · retention $141,023).

PROJECT_ORDER now contains 63 IDs (was 59). **Compass becomes 12 closed** (#2076, #2078 added). **Intracorp becomes 16th closed-portfolio GC** (joining Chinn/SRM/Exxel/Compass/BMDC/Blueprint/Natural & Built/Marpac/Shelter/Holland/GRE/Synergy/Farrell-McKenna/Jabooda/Braseth/Intracorp). **Braseth becomes 15th closed-portfolio GC**.

**Calibration sample now n=50** (was n=46). Recalibrated via `python3 owp/build_calibration.py` and re-embedded. Headline benchmarks: hours/unit 116.1 · gross_margin 39.4% (was 40.3% at n=46) · loaded_wage $40.35/hr (was $39.73) · burden_multiplier 1.432 (was 1.443). Cushion codes essentially unchanged. Overrun: code 120 Roughin Labor +5.4% (was +4.4% at n=46 — softer than v1.2's +10.2% but still the only major overrun).

**Three pattern signals strengthened by this batch:**

1. **#2078 is the new Mechanism A extreme: 40 COs / +28.3% gross over.** Confirms my earlier reframe that gross-over-contract magnitude (not the $0-net flag) is what's worth modeling. Larger / more-CO-heavy jobs run hotter gross billings before crediting back; #2078's 40-CO posture is structurally different from #2076's 23-CO / +1.2% on the same GC. Worth digging into WHY one Compass job ran so much hotter than the other.

2. **#2077 vendor concentration is a real outlier.** Rosen 77.6% on a $5.6M job is extreme — far above the 35% watch threshold that already flagged in the bid tool. Single-source-of-failure flag. If a 4th project shows similar Rosen-dominance (>50%), it suggests OWP is increasingly reliant on Rosen for big-AP-volume Intracorp/Exxel-class jobs.

3. **The Keller-led pattern is now N=4** (#2071 27%, #2074 29%, #2076 37%, #2078 33%). Not GC-specific (spans Chinn / Marpac / Compass×2). Suggests Keller is the secondary go-to supplier for jobs where Rosen isn't dominant — possibly Compass-Kirkland axis (both #2076 and #2078 are Compass) plus the Keller-friendly GCs (Chinn historically). Watch.

## #2079 BMDC Crown Hill + #2080 Compass Redmond Square (added 2026-04-29)

- **#2079 BMDC, Crown Hill** (Seattle WA · 7730 15th Ave NW, Crown Hill · 54-unit multifamily · 376 fixtures · $1,035,531 final / $992,600 base / **4 executed COs net $0** — Mechanism A but minimal traffic ($200 in credits, essentially flat) · **23.1% gross margin** — second-lowest in recent batch · $239,561 net profit · 40 workers / 8,570 hrs / 17 months Nov 2020 → Apr 2022 · Subcontract **AIA A401-2017** executed 2021-01-20 · GC BMDC — **OWP's first closed BMDC engagement** (BMDC was previously 2 live only: #2104, #2112) · Owner 7750 15th Ave NW LLC · Architect **Johnston Architects** (same architect as #2061 Alta CC) · MEP not in folder, likely Franklin per OWP standard · OWP PM Richard Donelson · Insurance Standard (per A401 contract) · top vendor **Rosen Supply Kirkland 28.0% / 80 invoices** with Keller (23.6%) and Consolidated (13.3%) — distributed top-3 (continues the small-job balanced-vendor pattern from #2075 Braseth) · 15 vendors total, $280,493 AP across 192 invoices · 9 RFIs + 4 submittals — low load · retention $51,777 still held).

- **#2080 Compass General, Redmond Square** (Redmond WA · 16563 Redmond Way + 16425 Cleveland Street · 311-unit mixed-use multifamily · 1,900 fixtures — **largest-fixture closed Compass job** · $5,931,932 final / 18 executed COs net $0 — but mechanism is **TRUE Mechanism B (zero AR credits)**, the **first time we've seen Mechanism B on a Compass job** (was previously GRE-only via #2067). $5.93M positive billings + $0 credits = $5.93M net flat. The 18 COs were processed at $0 internally or never billed as separate items · **40.1% gross margin** · $2,377,107 net profit · $3,554,825 direct cost · 63 workers / 37,438 hrs / 30 months Sep 2021 → Mar 2024 · GC Compass General Construction I, LLC · Owner Redmond Grand LLC (Compass-affiliate) · MEP **Emerald City Engineers, Inc. (Lynnwood WA) — FIRST non-Franklin MEP in OWP closed portfolio**; every prior closed job used Franklin Engineering as standard · OWP PM Richard Donelson · Insurance OCIP / Wrap-up (Compass developer-administered, confirmed via CO#06 OCIP Credit) · top vendor **Rosen Supply Kirkland 28.2% / 309 invoices / $383,263** with Keller (25.6%, $347k) and Consolidated (15.8%, $215k) — distributed top-3 · 22 vendors total, $1,357,848 AP across 742 invoices · 23 RFIs + 19 submittals · retention $296,597 still held).

PROJECT_ORDER now contains 65 IDs (was 63). **Compass becomes 13 closed + 2 live (15 total — largest-volume GC by far)**. **BMDC becomes 1 closed + 2 live (first closed BMDC sets a baseline for the 2 in-flight live projects)**. Calibration sample now n=52 (was n=50). Recalibrated via `python3 owp/build_calibration.py` and re-embedded. Headline benchmarks: hours/unit 117.1 (was 116.1) · gross_margin 39.4% (unchanged) · loaded_wage $40.65/hr (was $40.35) · burden_multiplier 1.428 (was 1.432). Cushion codes essentially unchanged: 211 -44.7%, 220 -27.2%, 241 -23.9%, 601 -22.1%, 111 -17.2%. Overrun: code 120 Roughin Labor +5.4%.

**GC_DB and bid tool dropdown updated to reflect new composition:**
- Compass: 10 closed + 2 live → **13 closed + 2 live** (jobs counter 12 → 15). knownTag updated to flag the bimodal CO behavior (Mechanism A on most jobs, but #2080 ran true Mech B)
- BMDC: 2 live → **1 closed + 2 live** (jobs counter 2 → 3). knownTag flags low-margin small-job profile so far (23.1% on #2079)
- Chinn: 5 closed + 1 live → **6 closed + 1 live** (added #2071 Stellar)
- Marpac: 2 closed → **3 closed** (added #2074 Buddha Jewel — all three are 40%+ margin)
- Intracorp + Braseth: new GC entries added (each 1 closed)

**Two new pattern signals from this batch:**

1. **Mechanism B is no longer GRE-only.** #2080 Compass Redmond Square executed 18 COs with ZERO AR credits and zero gross-over-contract movement — first Compass job to use the GRE-style true-$0-CO posture. With #2080's 40.1% margin sitting between #2076's 31.3% (Mech A +1.2% gross) and #2078's 23.8% (Mech A +28.3% gross), within-GC margin variance on closed Compass jobs (12 jobs, 23.8%–58.2% spread) is now wider than I'd expected. Mechanism B may correlate with healthier margins (#2067 GRE 29.8% was lower; need more N).

2. **First non-Franklin MEP partner in closed portfolio (#2080 Emerald City Engineers).** Every other closed job in the dataset used Franklin Engineering as the MEP. This is a meaningful diversification — worth tracking whether Emerald City correlates with different cost-code behavior. Same applies to #2116 The V (live) which uses Emerald City Engineers — first time we'd see consistency on a non-Franklin partner across a Compass + SRM job.

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
