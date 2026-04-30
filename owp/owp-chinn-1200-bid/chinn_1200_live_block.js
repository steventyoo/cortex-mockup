// Chinn 1200 .live block — exported from index.html on 2026-04-29
// Lives outside the PROJECTS dict, attached after PROJECTS is defined
// Drives Job Health view: burnCurve, phases, anomalies, recommends

if (PROJECTS['CHINN-1200']) PROJECTS['CHINN-1200'].live = {
    lastSync: 'today',
    title: '1200 116th Ave NE',
    subtitle: '215-unit · Chinn Construction · Bellevue 98004 (downtown core) · SD-set bid',
    chipContract: 'Bid: $4,700,000 (1st)',
    chipOrigMargin: 'O&P target: 12% ($504k)',
    chipForecastMargin: '○ Forecast margin: ~32% gross / 12% net',
    pctComplete: '0%',
    pctSub: 'SD pricing · awaiting Chinn acceptance · 2nd bid May 2026',
    cortexStatus: '○ SD-set bid · not CD',
    cortexStatusBody: 'Chinn job E25-07 · 215u / 1,648 permit fx / $4.70M 1st bid / $21,860 per unit / $2,852 per fixture. Heat-pump central plant (4 NYLE + 5 storage tanks + 23 support boilers + controls = ~$640k equipment). Owner High Street NW Development (Trammell Crow). Architect Weber Thompson. SD pricing 11/18/2025 — contingency posture is implicit in equipment line items, no explicit SD-risk cushion. Submeter system priced at $0 (deferred to owner scope). 2nd bid round in flight May 2026.',
    eta: 'TBD',
    etaSub: 'schedule locks at subcontract execution · typical 22-26 mo for 215u Bellevue',
    hoursVar: '—',
    hoursVarSub: 'no field hours yet · SD-pricing phase',
    costVar: '$0',
    costVarSub: 'no AP yet · design + bid only',
    schedVar: '—',
    schedVarSub: 'NTP / mobilization not scheduled',
    forecastEac: '$4.70M',
    forecastEacSub: 'EAC matches 1st bid · will revise on 2nd bid',
    burnCurve: [
      {pct: 0, actual: 0, baseline: 0}
    ],
    phases: [
      {phase:'SD pricing / takeoff',  pct:100, actual:1, baseline:1, status:'on'},
      {phase:'2nd bid round (May)',   pct:0,  actual:0, baseline:0, status:'notstarted'},
      {phase:'Subcontract execution', pct:0,  actual:0, baseline:0, status:'notstarted'},
      {phase:'Underground rough-in',  pct:0,  actual:0, baseline:0, status:'notstarted'},
      {phase:'Roughin',               pct:0,  actual:0, baseline:0, status:'notstarted'},
      {phase:'Trim & fixtures',       pct:0,  actual:0, baseline:0, status:'notstarted'},
      {phase:'Test & inspection',     pct:0,  actual:0, baseline:0, status:'notstarted'}
    ],
    anomalies: [
      {date:'Oct 17, 2025', title:'ADR Submittal filed', body:'City of Bellevue design review submittal. Architect Weber Thompson, civil Coughlin Porter Lundeen. 215 units, 8 above + 3 below grade.', severity:'INFO'},
      {date:'Nov 18, 2025', title:'SD Pricing Set issued', body:'Schematic Design Pricing Set — NOT Construction Documents. SD-stage bids carry implicit scope risk (vs CD bids). Permit Fee Estimator pinned 1,675 fixtures (1,509 std + 156 special + 6 BFP + 4 WS). Bellevue permit fee ~$33k.', severity:'INFO'},
      {date:'Dec 11, 2025', title:'1st bid submitted to Chinn', body:'$4,700,000 base bid. RI mat $862k + RI lab $1.12M + finish mat $1.12M + finish lab $114k + equip $113k + eng/rental/parking $121k + permits $43k = $4.20M direct cost. 12% O&P = $504k = $4.70M total. Heat-pump central plant ($555k NYLE+tanks+controls). Submeter system $0 (deferred to owner scope).', severity:'WARN'},
      {date:'Apr 2026', title:'2nd bid round in flight', body:'Chinn requested 2nd bid round May 2026 — typical for SD-set bids as design progresses to DD/CD. Watch for: (1) submeter scope re-assignment, (2) parking-deck drainage scope change, (3) heat-pump-system right-sizing.', severity:'WATCH'}
    ],
    recommends: [
      {urgency:'urgent', label:'ACTION 01 · PRE-2ND-BID', title:'Verify submeter scope assignment with Chinn', body:'1st bid priced submeter system at $0 (rows 64-66 in ESTIMATE: cold meters, hot meters/transceivers, repeaters all zero). Spec calls for hot-only submetering w/ transceivers + repeaters — confirm with Chinn whether this is owner scope (deferred) or oversight to capture in 2nd bid. Could be +$50-70k delta.', owner:'Estimating', ownerLabel:'Owner', extra:'Bid delta', extraVal:'+$50-70k if added back'},
      {urgency:'urgent', label:'ACTION 02 · PRE-2ND-BID', title:'Re-anchor against $/unit calibration v1.4', body:'1st bid lands at $21,860/u — above OWP closed-portfolio P75 of $18,244 but justified by heat-pump central + Bellevue urban site. v1.4 calibration just shipped with orig_bid_per_unit benchmark; re-validate 2nd bid against archetype-filtered comps (mid-rise heat-pump multifamily $17,500-18,200/u + scope premiums).', owner:'Estimating', ownerLabel:'Owner', extra:'Calibration', extraVal:'v1.4 (29 Apr 2026)'},
      {urgency:'urgent', label:'ACTION 03 · PRE-2ND-BID', title:'Capture Chinn job E25-07 PM team', body:'GC PM / PE / Sup names not yet captured. Chinn 1200 is OWP\'s 8th Chinn engagement (7 closed + 1 in design). Chinn-OWP relationship is healthy — closed-portfolio Chinn margins run +5% above OWP median. Capture team for dashboard PROJECT_TEAMS enrichment.', owner:'OWP Admin', ownerLabel:'Owner', extra:'Deliverable', extraVal:'Team data populated'},
      {urgency:'strategic', label:'ACTION 04 · POST-AWARD', title:'Track Chinn 1200 against KOZ #2099 heat-pump benchmark', body:'KOZ Apartments (closed) is the closest system comp — heat-pump central plant, Trane vendor, 169u $3.03M / 32.4% margin. Chinn 1200 should track similar margin profile if scope stays stable. Use as live BvA benchmark once mobilized.', owner:'PM', ownerLabel:'Owner', extra:'Comp', extraVal:'#2099 KOZ Apts'}
    ]
  };
