#!/usr/bin/env python3
"""Format fragments.json into JS text blocks ready for injection."""
import json
from pathlib import Path

DATA = json.loads(Path('/sessions/gracious-relaxed-pascal/gen2035/fragments.json').read_text())

def js(v):
    """JSON serialize (gives us valid JS for primitives/arrays/objects)."""
    if v is None: return 'null'
    if isinstance(v, bool): return 'true' if v else 'false'
    if isinstance(v, (int, float)):
        # Avoid 1.0 -> 1.0 is OK in JS
        return json.dumps(v)
    if isinstance(v, str):
        return json.dumps(v)
    if isinstance(v, list):
        return '[' + ', '.join(js(x) for x in v) + ']'
    if isinstance(v, dict):
        return '{' + ', '.join(f'{json.dumps(k)}: {js(val)}' for k, val in v.items()) + '}'
    raise ValueError(f'cannot serialize {type(v)}')

def fmt_project_entry(pid, pe):
    """Format as a multi-line PROJECTS entry matching 2034 style."""
    lines = [f'    "{pid}": {{']
    lines.append(f'      reportNum: {js(pe["reportNum"])},')
    chips_str = ','.join(js(c) for c in pe['chips'])
    lines.append(f'      chips: [{chips_str}],')
    lines.append(f'      jobTop: {js(pe["jobTop"])}, jobBottom: {js(pe["jobBottom"])},')
    lines.append(f'      jobNum: {js(pe["jobNum"])}, location: {js(pe["location"])}, gc: {js(pe["gc"])},')
    lines.append(f'      gcShort: {js(pe["gcShort"])},')
    lines.append(f'      dateRange: {js(pe["dateRange"])},')
    lines.append(f'      unitsDesc: {js(pe["unitsDesc"])},')
    lines.append(f'      revenue: {js(pe["revenue"])}, revenueFoot: {js(pe["revenueFoot"])},')
    lines.append(f'      profit: {js(pe["profit"])}, profitFoot: {js(pe["profitFoot"])},')
    lines.append(f'      expenses: {js(pe["expenses"])}, expensesFoot: {js(pe["expensesFoot"])},')
    lines.append(f'      retainage: {js(pe["retainage"])}, retainageFoot: {js(pe["retainageFoot"])},')
    lines.append(f'      execHeadline: {js(pe["execHeadline"])},')
    lines.append(f'      execBody: {js(pe["execBody"])},')
    lines.append(f'      jobNum2: {js(pe["jobNum2"])}, workforce: {js(pe["workforce"])},')
    lines.append(f'      bpLeft: {js(pe["bpLeft"])}, bpRight: {js(pe["bpRight"])},')
    lines.append(f'      k01_revenue: {js(pe["k01_revenue"])}, k01_margin: {js(pe["k01_margin"])}, k01_origContract: {js(pe["k01_origContract"])},')
    lines.append(f'      k01_co: {js(pe["k01_co"])}, k01_directCost: {js(pe["k01_directCost"])}, k01_profit: {js(pe["k01_profit"])},')
    lines.append(f'      footerReport: {js(pe["footerReport"])},')
    lines.append(f'      sbName: {js(pe["sbName"])}, skyName: {js(pe["skyName"])},')
    sky = pe['skyline']
    lines.append(f'      skyline: {{ floors: {sky["floors"]}, cols: {sky["cols"]}, w: {sky["w"]} }},')
    num = pe['num']
    num_parts = ', '.join(f'{k}: {js(v)}' for k, v in num.items())
    lines.append(f'      num: {{ {num_parts} }},')
    vm = pe['vendorMeta']
    vm_parts = ', '.join(f'{k}: {js(v)}' for k, v in vm.items())
    lines.append(f'      vendorMeta: {{ {vm_parts} }},')
    lines.append('      vendors: [')
    for v in pe['vendors']:
        lines.append(f'        {{ name: {js(v["name"])}, invoices: {v["invoices"]}, spend: {js(v["spend"])}, pct: {js(v["pct"])} }},')
    lines.append('      ]')
    lines.append('    }')
    return '\n'.join(lines)

def fmt_team(pid, team):
    html = team['teamHTML']
    units = team['jobInfoUnits']
    return f'    "{pid}": {{"teamHTML": {json.dumps(html)}, "jobInfoUnits": {units}}}'

def fmt_arrays(pid, arr):
    """Format the PROJECTS['pid'].xxx = ... block."""
    lines = []
    lines.append(f'  PROJECTS[\'{pid}\'].sovData = {js(arr["sovData"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].payApps = {js(arr["payApps"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].allVendors = {js(arr["allVendors"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].phases = {js(arr["phases"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].insights = {js(arr["insights"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].changeLog = {js(arr["changeLog"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].changeMeta = {js(arr["changeMeta"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].rootCauses = {js(arr["rootCauses"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].responsibility = {js(arr["responsibility"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].predictiveSignals = {js(arr["predictiveSignals"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].crewRoster = {js(arr["crewRoster"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].tierDist = {js(arr["tierDist"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].wageStats = {js(arr["wageStats"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].costCodes = {js(arr["costCodes"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].costCats = {js(arr["costCats"])};')
    lines.append(f'  PROJECTS[\'{pid}\'].bva = {js(arr["bva"])};')
    return '\n'.join(lines)

# Assemble output blocks
out_projects = []
out_teams = []
out_arrays = []
for pid in sorted(DATA.keys()):
    d = DATA[pid]
    out_projects.append(fmt_project_entry(pid, d['project_entry']))
    out_teams.append(fmt_team(pid, d['team_entry']))
    out_arrays.append(fmt_arrays(pid, d['arrays']))

Path('/sessions/gracious-relaxed-pascal/gen2035/projects_block.js').write_text(',\n'.join(out_projects))
Path('/sessions/gracious-relaxed-pascal/gen2035/teams_block.js').write_text(',\n'.join(out_teams))
Path('/sessions/gracious-relaxed-pascal/gen2035/arrays_block.js').write_text('\n\n'.join(out_arrays))
print('Wrote projects_block.js, teams_block.js, arrays_block.js')
