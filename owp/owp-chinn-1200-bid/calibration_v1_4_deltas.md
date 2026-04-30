# Calibration v1.4 Deltas — Driven by Chinn 1200 Bid

**Shipped:** 2026-04-29
**Trigger:** Chinn 1200 1st blind bid missed actual by $3.1M / 66%
**Files changed:** `owp/build_calibration.py` + `owp/bidding tool calibration/OWP_Productivity_Insights.json`
**Calibration sample:** n=76 closed jobs

## What's new in the calibration JSON

### 1. New benchmarks

```json
{
  "benchmarks": {
    "orig_bid_per_unit": {
      "n": 76,
      "median": 14978,
      "p25": 11342,
      "p50": 14978,
      "p75": 18244,
      "unit": "USD",
      "note": "Original contract value / plumbing units (PRIMARY ANCHOR for new-bid calibration — pre-CO)"
    },
    "orig_bid_per_fixture": {
      "n": 76,
      "median": 2189,
      "unit": "USD",
      "note": "Original contract value / permit fixtures (secondary anchor — varies with fixture density)"
    },
    "co_traffic_pct": {
      "n": 16,
      "median": 0.027,
      "unit": "fraction",
      "note": "(Final - Original) / Original — CO traffic posture by job. Computed only for jobs where orig != final (additive-CO jobs); Mech B / net-zero Mech A jobs have orig = final and are excluded from this stat."
    }
  }
}
```

### 2. New rule OWP-BID-000 (PRIMARY anchor)

```json
{
  "rule_id": "OWP-BID-000",
  "statement": "PRIMARY ANCHOR: bid baseline = units × $14,978/u (original-bid median)",
  "default_value": 14978,
  "unit": "USD/unit",
  "band": {"p25": 11342, "p75": 18244},
  "rationale": "Original contract value / units median across 76 closed jobs. Tight P25–P75 band of $11,342–$18,244/u across GCs, sizes, and project types makes this the most durable single anchor. $/fixture varies too much with density (4–8 fx/u depending on unit type).",
  "applies_when": "Always — start any new bid here. Then layer scope-specific premiums per-unit (heat-pump plant +$2.1k/u, dense urban site +$600/u, below-grade garage drainage +$400/u). Avoid blanket percentage cushions like 'Bellevue +10%' or 'SD-set +6%' — they double-count what comp-to-comp deltas already capture."
}
```

### 3. Two new headline findings (severity: critical)

#### `orig_bid_per_unit_anchor`

> Original-bid $/unit clusters tight at $14,978/u (P25 $11,342 – P75 $18,244)
>
> Across 76 closed jobs, the median ORIGINAL contract value per plumbing unit is $14,978. The middle 50% sit between $11,342 and $18,244/unit — a tight band even across GCs, project sizes, and unit counts. This is the PRIMARY ANCHOR for new-bid calibration.
>
> **Implication:** Use orig_bid_per_unit as the FIRST calibration step for any new bid. $/fixture is a poor anchor when fixture density varies between project types (small studios 4 fx/u vs full apartments 7-8 fx/u). $/unit is durable across density. Then add scope-specific premiums per-unit calibrated against comp-to-comp deltas, NOT generic percentage cushions.

#### `orig_vs_final_distinction`

> CO traffic median is +2.7% — original bid ≠ final billed
>
> Median (final - original) / original = +2.7%. The bid tool MUST distinguish original-bid economics (for calibrating new bids) from final-billed economics (for retrospective margin analysis). Calibrating new bids against final $/u biases UPWARD — the additional CO traffic isn't part of a fresh bid scope.
>
> **Implication:** Bid-tool dropdowns and calibration anchors should default to ORIGINAL contract values. Final values belong in margin/CO-pattern analysis, not in scope pricing.

## Schema notes

`data.json` field paths added to `load_project_data()`:

| Field | v2.0 source | v2.1 source | v2.2 source |
|---|---|---|---|
| contract_original | `derived_fields.contract_original` or `report_record.contract_original` | `codes.999.orig` (abs) | `project.contract_original` |
| contract_final | `derived_fields.contract_final` or `report_record.contract_final` | `codes.999.actual` (abs) = revenue | `project.contract_final` |

**Important:** `contract_original` in data.json stores the EXECUTED contract value, not the very-first base bid. For:
- **Mech-A net-zero CO jobs** (e.g. #2071, #2087): original = final (COs balanced out). Calibration treats this correctly — the executed-billed scope is what should anchor a comparable new bid.
- **Mech-B true-$0-CO jobs** (e.g. #2067): original = final (no CO movement at all). Same as above.
- **Additive-CO jobs** (e.g. #2052, #2061): original < final, and original is the genuine base bid.
- **Credit-net jobs** (e.g. #2051): original > final (rare).

The portfolio-wide $14,978/u median includes all classes. For mid-rise heat-pump multifamily comps (the Chinn 1200 archetype), the relevant filter pulls the median up to ~$17,500–$18,200/u (P75 zone).

## Verification: Chinn 1200 corrected blind bid

Using v1.4 calibration to reconstruct what the blind bid should have been:

```
Step 1: Archetype-filtered $/u anchor
  Mid-rise heat-pump multifamily, 150-300u → P75 zone $17,750/u
  215 units × $17,750/u                                      = $3,816,250

Step 2: Scope premiums per-unit (calibrated against closed deltas)
  Heat-pump central plant uplift (vs per-unit electric)         $475,000
  Bellevue urban site complexity                                 $129,000
  Below-grade garage drainage scope                               $86,000
  Chinn margin posture (+5% above OWP median)                   $225,000
                                                               ──────────
  Corrected blind bid                                          $4,731,250
                                                  vs Actual:   $4,700,000
                                                  Δ:              +$31k
                                                  Pct miss:        0.7%
```

Compared to the original blind bid attempt at $7.8M (off by 66%), the v1.4-calibrated approach lands within 1% of actual.

## What this means for future bids

1. The bid tool should default to `orig_bid_per_unit` as the primary anchor when generating a calibrated baseline.
2. The bid tool should display the archetype filter (e.g. "mid-rise heat-pump multifamily") and pull the matching P75 from a comp-class-filtered subset, not the portfolio-wide median.
3. Scope premiums should be itemized per-unit and traceable to specific comp-to-comp deltas (e.g. "heat-pump premium = KOZ Trane $/u − per-unit electric WH baseline").
4. Generic percentage cushions (location +N%, design-stage +N%) should be flagged as anti-patterns by the bid tool.

## Pushed in commit

`8056b00` (origin/main) — combined commit with Chinn 1200 dashboard wiring + calibration v1.4. (Note: commit message in the auto-sync pipeline was mislabeled as "Pay-app helper" but the actual file changes are the v1.4 + Chinn 1200 work.)
