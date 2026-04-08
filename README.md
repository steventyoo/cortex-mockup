# Cortex — Specialty Sub Intelligence

A single-file HTML mockup of Cortex, an intelligence platform for specialty subcontractors. Built for One Way Plumbing LLC as a design prototype.

## What it is

`index.html` is a standalone, self-contained report with no build step. It includes:

- **Home / Project Report** — editorial intelligence report for Job #2009 Chinn Greenwood Apartments, with 11 clickable KPI cards (Financial Overview, Budget vs Actual, Cost Breakdown, Material Spend, Crew & Labor, Productivity, PO Commitments, Billings & SOV, Change Order Log, Insights, Benchmark KPIs) that each open full drill-down modals populated with real JCR data
- **Portfolio Overview** — four real OWP jobs (2009, 2010, 2011, 2012) as clickable summary cards
- **Benchmark Comparison** — 41-metric 4-way comparison table with best/worst shading and averages, sourced from `OWP_4Way_Comparison.xlsx`
- **Document Pipeline** — 4-stage ingestion view (ingest → classify → parse → index)
- **Operator Workbench** — 6 analyst workflows with status (live/pending)

Clicking a project in the left sidebar swaps the Home report to that job's real data (hero, KPI-01, exec summary, workforce, blueprint, chips) pulled from the canonical 4-Way benchmark workbook.

## Stack

- Plain HTML / CSS / JavaScript
- Tailwind via CDN (no build step)
- Inline SVG for icons, blueprint elevation, and charts
- Typography: Instrument Serif (italic display), Inter (body), JetBrains Mono (numerics and labels)
- Palette: Claude paper (`#F5F1EB` cream, `#1F1E1D` ink, `#B85C3E` clay accent)

## Running locally

Just open `index.html` in any browser. No server, no dependencies, no build.

```
open index.html
```

## Deploying

Because it's a single HTML file, both GitHub Pages and Vercel work with zero configuration. Vercel is recommended for faster iteration:

1. Push to GitHub
2. vercel.com → Add New Project → Import this repo → Deploy

Every future push auto-redeploys.

## Data sources

All numbers in the report are real, extracted from OWP's Sage Timberline exports:

- `OWP_2009_JCR_Summary.xlsx` — primary drill-down data for Job #2009
- `OWP_2012_JCR_Summary_Notion.xlsx` — Job #2012 Exxel 8th Ave
- `2010 Job Detail Report.pdf` — Job #2010 Chinn Old Town (revenue, profit, margin from Job Totals)
- `2011 Job Detail Report.pdf` — Job #2011 Chinn Barrett Park (revenue, profit, margin from Job Totals)
- `OWP_4Way_Comparison.xlsx` — canonical 4-way benchmark workbook, source for the Benchmark Comparison view

## Status

Design prototype · v0.2 · April 2026
