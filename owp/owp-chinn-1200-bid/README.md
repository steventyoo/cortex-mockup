# Chinn 1200 116th Ave NE — Live Bid Workspace

Working folder for OWP's first live bid in the dashboard. **OWP project number TBD** — using `CHINN-1200` as a placeholder string ID until assigned.

## Project snapshot

| Field | Value |
|---|---|
| Job name | 1200 116th Ave NE |
| Address | 1200 116th Ave NE · Bellevue WA 98004 (Downtown / 98004 core) |
| Site area | 38,341 SF (dense urban) |
| GC | Chinn Construction (job E25-07) |
| Owner | High Street Northwest Development (Trammell Crow) |
| Architect | Weber Thompson |
| Civil | Coughlin Porter Lundeen |
| MEP | TBD |
| Building | 8 above + 3 below grade · Type IIIA wood / Type IA podium / Type IA garage |
| Units | 215 (market + affordable, ~736 sf avg) |
| Permit fixtures | 1,648 (1,502 upper + 146 lower per ESTIMATE tab) |
| City permit-fee count | 1,675 (1,509 std + 156 special + 6 BFP + 4 WS) |
| Hot water | Heat-pump central plant — 4 NYLE WHs ($448k) + 5 storage tanks ($95k) + 23 support boilers ($80.5k) + controls ($12k) |
| Submeters | Hot-only + transceivers + repeaters — priced at **$0** in 1st bid (deferred to owner scope or oversight) |
| Special | Sauna + frost locker (vendor-supplied AmFinn) · 2 dog wash tubs · LEED Silver |
| Bid set | **SD Pricing Set 11/18/2025** (NOT CD) |

## Bid history

| Milestone | Date | Amount | Notes |
|---|---|---|---|
| ADR Submittal | Oct 17, 2025 | — | City of Bellevue design review |
| SD Pricing Set | Nov 18, 2025 | — | Schematic Design — NOT Construction Documents |
| **1st bid** | **Dec 11, 2025** | **$4,700,000** | $21,860/unit · $2,852/fixture · 12% O&P |
| 2nd bid | May 2026 | TBD | In flight |

## Files in this folder

| File | Purpose |
|---|---|
| `README.md` | This file |
| `bid_blind_post_mortem.md` | Full write-up: 1st blind bid $7.8M miss, root-cause analysis, corrected $4.73M methodology |
| `chinn_1200_metadata.json` | Structured project data (units, fixtures, scope, comps) |
| `bid_sheet_estimate_summary.csv` | Extracted ESTIMATE tab line items from OWP bid sheet (Dec 11, 2025) |
| `chinn_1200_hero_block.js` | Dashboard hero block (PROJECTS['CHINN-1200']) — exported from index.html for reference |
| `chinn_1200_live_block.js` | Dashboard `.live` block (burnCurve, phases, anomalies, recommends) |
| `calibration_v1_4_deltas.md` | Summary of calibration JSON changes shipped 2026-04-29 |

## Linked dashboard entry

The Chinn 1200 entry is wired into `index.html` at PROJECT_ORDER position 80 (last entry), keyed as `"CHINN-1200"`. Pre-construction state with `.live` block (4 action items in recommends panel).

## Related closed-portfolio comps (anchors for calibration)

| Comp | Why it anchors |
|---|---|
| #2099 KOZ Apartments (Kirtley-Cole) | Closest system match — heat-pump central plant, Trane top vendor at 30.4% AP |
| #2071 Chinn Stellar (1405 Dexter) | Same GC (Chinn), dense urban Seattle, similar 160u scale |
| #2061 Alta Columbia City (Exxel) | OWP house median anchor, 7-story, 41.4% margin |
| #2087 Exxel Northgate Roosevelt | Larger but Bellevue-area mid-rise, 47.6% margin |
| #2107 Chinn 68th & Roosevelt (live) | Same Trammell Crow / High Street NW owner family — but DIFFERENT project |

## What changed in calibration v1.4 (driven by this bid)

The 1st blind bid attempt missed actual by **$3.1M / 66%** because the calibration used $/fixture × CLOSED-final values + phantom percentage premiums (Bellevue +10%, SD-set +6%). Root cause analysis produced calibration v1.4 with:

- New `orig_bid_per_unit` benchmark (PRIMARY anchor): **$14,978/u portfolio median, P25 $11,342, P75 $18,244, n=76**
- New `orig_bid_per_fixture` benchmark (secondary, density-sensitive)
- New `co_traffic_pct` benchmark (median +2.7% on additive-CO jobs)
- New rule **OWP-BID-000** naming per-unit (NOT per-fixture) as PRIMARY anchor for new-bid calibration

See `bid_blind_post_mortem.md` for full write-up.

## Status

- [x] 1st bid submitted to Chinn (Dec 11, 2025)
- [x] Wired into dashboard as `PROJECTS["CHINN-1200"]`
- [x] Calibration v1.4 shipped (orig_bid_per_unit, orig_bid_per_fixture, co_traffic_pct)
- [ ] 2nd bid round (May 2026) — pending
- [ ] Submeter scope verified with Chinn (currently $0 — deferred to owner or oversight?)
- [ ] Chinn PM / PE / Sup names captured for PROJECT_TEAMS enrichment
- [ ] Chinn awards subcontract → switch to construction phase
- [ ] OWP assigns real project number → swap `CHINN-1200` → `#XXXX` everywhere
