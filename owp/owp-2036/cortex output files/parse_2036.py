#!/usr/bin/env python3
"""Generic JDR parser for 2035-2040: adapted from parse_2033.py"""
import re, json, sys
from collections import defaultdict
from pathlib import Path

PID = sys.argv[1]
BASE = Path(f'/sessions/gracious-relaxed-pascal/mnt/cortex-mockup/owp/owp-{PID}/cortex output files')
JDR_TXT = BASE / f'{PID}_jdr.txt'
OUT_JSON = BASE / f'{PID}_data.json'

text = JDR_TXT.read_text()
lines = text.split('\n')

codes_dict = {}
totals_re = re.compile(r'Cost Code Totals\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)')
hours_re = re.compile(r'Payroll Hours:\s+([\d,.\-]+)\s+\(\s*Reg:\s+([\d,.\-]+)\s+O/T:\s+([\d,.\-]+)\s+D/T:\s+([\d,.\-]+)\s+Other:\s+([\d,.\-]+)')

def num(s):
    s = s.strip().replace(',', '')
    if s.endswith('-'):
        try: return -float(s[:-1])
        except: return 0.0
    if s == '' or s == '0.00':
        return 0.0
    try: return float(s)
    except: return 0.0

i = 0
while i < len(lines):
    line = lines[i]
    m = re.match(r'^(\d{3})\s{10,}(.+?)\s*$', line)
    if m:
        code = m.group(1)
        desc = m.group(2).strip()
        codes_dict[code] = {'desc': desc, 'orig': 0, 'rev': 0, 'var': 0, 'actual': 0,
                            'net': 0, 'ret': 0, 'hrs_total': 0, 'hrs_reg': 0,
                            'hrs_ot': 0, 'hrs_dt': 0, 'hrs_other': 0}
        for j in range(i+1, min(i+3000, len(lines))):
            t = totals_re.search(lines[j])
            if t:
                codes_dict[code]['orig']   = num(t.group(1))
                codes_dict[code]['rev']    = num(t.group(2))
                codes_dict[code]['var']    = num(t.group(3))
                codes_dict[code]['actual'] = num(t.group(4))
                codes_dict[code]['net']    = num(t.group(5))
                codes_dict[code]['ret']    = num(t.group(6))
                if j+1 < len(lines):
                    h = hours_re.search(lines[j+1])
                    if h:
                        codes_dict[code]['hrs_total'] = num(h.group(1))
                        codes_dict[code]['hrs_reg']   = num(h.group(2))
                        codes_dict[code]['hrs_ot']    = num(h.group(3))
                        codes_dict[code]['hrs_dt']    = num(h.group(4))
                        codes_dict[code]['hrs_other'] = num(h.group(5))
                i = j
                break
    i += 1

workers = defaultdict(lambda: {'name': '', 'hours': 0.0, 'amount': 0.0, 'days': set()})
pr_re = re.compile(r'^PR\s+\d+\s+(\d{2}/\d{2}/\d{2})\s+([A-Z]{2}\d{2})\s+(.+?)\s*$')
hour_line_re = re.compile(r'^\s+(\d{2}/\d{2}/\d{2})\s+Regular:\s+([\d.]+)\s+hours\s+([\d,.\-]+)')

for i, line in enumerate(lines):
    m = pr_re.match(line)
    if m and i+1 < len(lines):
        wid = m.group(2); wname = m.group(3).strip()
        h = hour_line_re.match(lines[i+1])
        if h:
            hrs = num(h.group(2)); amt = num(h.group(3))
            workers[wid]['name'] = wname
            workers[wid]['hours'] += hrs
            workers[wid]['amount'] += amt
            workers[wid]['days'].add(h.group(1))

ap_re = re.compile(r'^AP\s+\d+\s+(\d{2}/\d{2}/\d{2})\s+(\S+)\s+(.+?)\s*$')
ap_detail_re = re.compile(r'^\s+\d{2}/\d{2}/\d{2}\s+(.+?)\s+([\d,.\-]+)\s*$')
vendors = defaultdict(lambda: {'name': '', 'total': 0.0, 'count': 0})
for i, line in enumerate(lines):
    m = ap_re.match(line)
    if m and i+1 < len(lines):
        vid = m.group(2); vname = m.group(3).strip()
        detail = lines[i+1]
        d = ap_detail_re.match(detail)
        if d:
            nums = re.findall(r'(?:^|\s)([\d,]+\.\d{2}-?)', detail)
            if nums:
                amt = num(nums[-1])
                vendors[vid]['name'] = vname
                vendors[vid]['total'] += amt
                vendors[vid]['count'] += 1

invoices = defaultdict(lambda: {'total': 0.0, 'retainage': 0.0, 'date': '', 'lines': 0})
ar_header_re = re.compile(r'^AR\s+\d+\s+(\d{2}/\d{2}/\d{2})\s+\S+\s+')
for i, line in enumerate(lines):
    h = ar_header_re.match(line)
    if h and i+1 < len(lines):
        date = h.group(1)
        inv_line = lines[i+1]
        m = re.search(r'Invoice\s+(\d+)\s+([\d,.\-]+)(?:\s+([\d,.\-]+))?', inv_line)
        if m:
            inv = m.group(1)
            amt = num(m.group(2))
            ret = num(m.group(3)) if m.group(3) else 0
            if not invoices[inv]['date']:
                invoices[inv]['date'] = date
            invoices[inv]['total'] += amt
            invoices[inv]['retainage'] += ret
            invoices[inv]['lines'] += 1

total_actual = sum(c['actual'] for c in codes_dict.values())
total_orig = sum(c['orig'] for c in codes_dict.values() if c['orig'] > 0)
total_rev = sum(c['rev'] for c in codes_dict.values() if c['rev'] > 0)
total_hours = sum(c['hrs_total'] for c in codes_dict.values() if c['hrs_total'] > 0)

print(f'=== {PID} summary ===')
print(f'codes: {len(codes_dict)} | orig={total_orig:,.0f} rev={total_rev:,.0f} actual={total_actual:,.0f} hours={total_hours:,.0f}')
print(f'workers: {len(workers)} | vendors: {len(vendors)} | invoices: {len(invoices)}')
# Top 5 codes by actual
top = sorted(codes_dict.items(), key=lambda kv: kv[1]['actual'], reverse=True)[:8]
for c,v in top:
    print(f'  {c} {v["desc"][:28]:28} Orig={v["orig"]:>10,.0f} Rev={v["rev"]:>10,.0f} Act={v["actual"]:>10,.0f} Hrs={v["hrs_total"]:>6.0f}')

out = {
    'codes': codes_dict,
    'workers': {k: {**v, 'days': len(v['days'])} for k, v in workers.items()},
    'vendors': {k: {'name': v['name'], 'total': v['total'], 'count': v['count']} for k, v in vendors.items()},
    'invoices': {k: {'date': v['date'], 'total': v['total'], 'retainage': v['retainage'], 'lines': v['lines']} for k, v in invoices.items()},
}
OUT_JSON.write_text(json.dumps(out, indent=2, default=str))
print(f'Saved {OUT_JSON}')
