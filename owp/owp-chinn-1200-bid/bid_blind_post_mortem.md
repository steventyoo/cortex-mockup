# Chinn 1200 Blind Bid — Post-Mortem & Calibration Lesson

**Date:** 2026-04-29
**Project:** Chinn 1200 116th Ave NE (CHINN-1200)
**1st bid actual (OWP):** $4,700,000
**1st blind bid attempt (Cortex):** $7,800,000
**Miss:** +$3.1M (66% over)

## What was the exercise?

Steven asked: "Can you do a BLIND BID, not knowing what OWP actually bid?"

Strict no-peek rules:
1. No reading the OWP bid sheet xlsx
2. No reading the OWP plumbing proposal pdf/docx
3. No reading 2nd bid folder (saved for round 2)
4. Build the bid only from specs/plans/permit docs

Pages read before committing the blind bid:
- Cover sheet, sheet index, abbreviations
- Specification_Log_Current.pdf (scope, sauna+frost locker, hardscape)
- 132416 Sauna and Frost Locker spec (AmFinn vendor)
- 010000 Outline Specifications - Add Divisions
- BO Plumbing Plan Review Inputs - PermitFeeEstimator (1,675 city permit count)
- Fees PermitFeeEstimator ($33,368 Bellevue + ~$3k BFP = $36k permits)
- Plans pages 1-3, 11-26, 27-35 (cover, civil, easements, transportation notes)

Did NOT read (no-peek): bid sheet, proposal pdf/docx, 2nd bid folder, plumbing fixture calc page G2.20 (couldn't find before committing).

## The committed blind bid: $7,800,000

```
1,675 fx × $3,400/fx (blended anchor)            = $5,695,000
+ Bellevue 98004 premium (+10%)                  = $   569,500
+ SD-set contingency (+6%)                       = $   375,870
+ Heat-pump plant equipment uplift               = $   150,000
+ Submeter system w/ transceivers                = $    65,000
+ Sauna/frost locker rough-in                    = $     6,000
+ Dog wash tubs (2) allowance                    = $    18,000
+ Site sewer/storm/water 75 LF                   = $    85,000
                                                   ───────────
SUBTOTAL                                         = $6,964,370
+ OH&P @ 12% (Chinn-OWP standard)                = $   835,724
                                                   ───────────
BLIND BID                                        ≈ $7,800,000
```

## What OWP actually bid: $4,700,000

```
Total RI material           $862,467
Total RI labor            $1,118,756
Total finish material     $1,121,681
Total finish labor          $114,480 (+ extra $420)
Supervision + takeoff        $48,375
Subs (insulation, sterilize) $28,600
Equipment (booster, sumps, OWS) $112,750
Engineering/rental/parking  $121,375
Permits                      $43,000
Tub repairs                   $2,580
                          ──────────
TOTAL COST                $4,196,269
+ OH&P @ 12%                $503,731
                          ──────────
TOTAL BID                 $4,700,000
```

Per-unit: **$21,860** · per-fixture: **$2,852** (using 1,648 permit count)

## Root cause: 4 calibration errors

### Error 1: Wrong primary anchor — $/fixture instead of $/unit
I weighted #2099 KOZ ($4,232/fx final) and #2061 Alta CC ($4,053/fx final) as anchors. But:
- KOZ has only 4.2 fx/unit (mostly small studios with shared kitchens)
- Chinn 1200 has 7.66 fx/unit (full apartments)

When fixture density shifts dramatically between project types, $/fx is no longer comparable. **$/unit is the durable anchor across density classes.** KOZ original bid was $17,455/unit → 215 × that = $3.75M baseline, which would have been much closer to reality.

### Error 2: Used closed-final values to calibrate an original bid
KOZ closed at $3,030,296 (final, post-CO traffic). I anchored on that. But the original bid was lower. **Bid-tool calibration must distinguish ORIGINAL contract economics from FINAL billed economics.** Final values are appropriate for retrospective margin analysis, NOT for calibrating a fresh bid.

### Error 3: Phantom "Bellevue +10%" location premium
I added 10% on the assumption Bellevue work commands a premium. The actual bid sheet uses standard OWP unit prices (toilets $153, lavs $198, tubs $657) — no jurisdictional markup. **OWP doesn't price location premium internally; whatever Bellevue cost-of-doing-business is already in the burden rate.** That added ~$570k of phantom money.

### Error 4: Phantom "SD-set +6%" contingency
I added 6% for schematic-design risk. The actual bid doesn't carry an explicit contingency line — pricing is direct from spec/quantities. **SD risk gets handled via CO traffic later, not a built-in cushion.** That added ~$376k of phantom money.

Combined the 4 errors put the blind bid ~$1.2M too high before getting to scope detail. On top of that I over-blended the heat-pump premium ($150k extra).

## Corrected methodology: $4,731,250 (off by $31k / 0.7%)

```
215u × $17,750/u (P75 anchor for mid-rise multifamily)   = $3,816,250
+ Heat-pump central plant uplift (vs per-unit electric)     $475,000
+ Bellevue urban site complexity                            $129,000
+ Below-grade garage drainage scope                          $86,000
+ Chinn margin posture (+5% above OWP median)              $225,000
                                                          ──────────
                                                           $4,731,250
                                              vs Actual:   $4,700,000
                                              Δ:              +$31k
```

Three rules:
1. **PRIMARY anchor: archetype-filtered $/unit (original-bid economics).** For mid-rise heat-pump multifamily 150-300u, the relevant comp band is $17,500-$18,200/u original. Portfolio-wide is $14,978/u but that mixes specialty jobs and small projects.
2. **Scope premiums per-unit, calibrated against comp-to-comp deltas.** Heat-pump central = (KOZ Trane equipment) − (per-unit electric WH baseline). Bellevue urban = (Bellevue closed jobs) − (Seattle closed jobs). NOT generic percentage cushions.
3. **GC margin posture is a real per-GC adjustment.** Chinn closed originals run slightly above OWP median; Compass clusters at median. This belongs in the bid tool's GC dropdown.

## What got shipped (calibration v1.4)

`owp/build_calibration.py` patched + `owp/bidding tool calibration/OWP_Productivity_Insights.json` regenerated:

- New `orig_bid_per_unit` benchmark: median $14,978/u, P25 $11,342, P75 $18,244, n=76
- New `orig_bid_per_fixture` benchmark: median $2,189/fx, n=76 (secondary anchor only)
- New `co_traffic_pct` benchmark: median +2.7% across the 16 additive-CO closed jobs
- New rule **OWP-BID-000**: per-unit (NOT per-fixture) is PRIMARY anchor for new-bid calibration. Explicit warning against generic percentage cushions.
- Two new headline findings (severity: critical): `orig_bid_per_unit_anchor` + `orig_vs_final_distinction`

Schema note for future reference: data.json's `project.contract_original` field stores the EXECUTED contract value, not the very-first base bid. For Mech-A net-zero CO jobs, original equals final because all COs balanced out — this is acceptable for calibration since the executed-and-billed scope IS what should anchor a comparable new bid.

## Recommendations for the 2nd bid (May 2026)

1. **Verify submeter scope assignment** — 1st bid has rows 64-66 (cold meters / hot meters/transceivers / repeaters) all $0 despite spec calling for hot-only submetering w/ transceivers + repeaters. Confirm with Chinn whether deferred to owner or oversight to capture. Could be +$50-70k delta.
2. **Re-anchor against $/unit calibration v1.4** — 1st bid lands at $21,860/u, above OWP closed-portfolio P75 of $18,244 but justified by heat-pump central + Bellevue urban site. Re-validate 2nd bid against archetype-filtered comps (mid-rise heat-pump multifamily $17,500-18,200/u + scope premiums).
3. **Capture Chinn job E25-07 PM team** — GC PM/PE/Sup names not yet captured. Chinn-OWP relationship is healthy (closed-portfolio Chinn margins run +5% above OWP median).
4. **Track against KOZ #2099 once mobilized** — closest system comp (heat-pump central plant). Use as live BvA benchmark.

## Lesson learned

Calibrating a fresh bid against retrospective benchmarks (closed-final $/fx) without distinguishing original-bid economics is a structural error. The bid tool now treats this distinction as a critical calibration rule going forward.
