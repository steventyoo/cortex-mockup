#!/usr/bin/env python3
"""Parse 2026 JDR: cost codes, payroll workers, AP vendors, AR invoices."""
import re, json
from collections import defaultdict

text = open('/sessions/keen-determined-mccarthy/work/2026_jdr.txt').read()
lines = text.split('\n')

# Find cost code sections
codes = []  # list of (code, desc, orig, rev, variance, actual, net_due, retainage, hours_reg, hours_ot, hours_dt, hours_other)
current_code = None
current_desc = None
in_code = False

# Regex: cost code header like "100           Supervision"
code_hdr_re = re.compile(r'^(\d{3})\s{2,}(.+?)\s*$')
# "Cost Code Totals       7,840.00       7,840.00            5,761.50-      2,078.50          0.00              0.00"
totals_re = re.compile(r'Cost Code Totals\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)')
hours_re = re.compile(r'Payroll Hours:\s+([\d,.\-]+)\s+\(\s*Reg:\s+([\d,.\-]+)\s+O/T:\s+([\d,.\-]+)\s+D/T:\s+([\d,.\-]+)\s+Other:\s+([\d,.\-]+)')

def num(s):
    s = s.strip().replace(',', '')
    if s.endswith('-'):
        return -float(s[:-1])
    if s == '' or s == '0.00':
        return 0.0
    try: return float(s)
    except: return 0.0

codes_dict = {}

i = 0
while i < len(lines):
    line = lines[i]
    # Match "NNN  Description" header - must not start w/ whitespace much; it's at col0
    m = re.match(r'^(\d{3})\s{10,}(.+?)\s*$', line)
    if m:
        code = m.group(1)
        desc = m.group(2).strip()
        # Look ahead for totals + hours
        codes_dict[code] = {'desc': desc, 'orig': 0, 'rev': 0, 'var': 0, 'actual': 0, 'net': 0, 'ret': 0, 'hrs_total': 0, 'hrs_reg': 0, 'hrs_ot': 0, 'hrs_dt': 0, 'hrs_other': 0}
        # find next cost code totals within ~200 lines
        for j in range(i+1, min(i+3000, len(lines))):
            t = totals_re.search(lines[j])
            if t:
                codes_dict[code]['orig'] = num(t.group(1))
                codes_dict[code]['rev'] = num(t.group(2))
                codes_dict[code]['var'] = num(t.group(3))
                codes_dict[code]['actual'] = num(t.group(4))
                codes_dict[code]['net'] = num(t.group(5))
                codes_dict[code]['ret'] = num(t.group(6))
                # hours may be on next line
                if j+1 < len(lines):
                    h = hours_re.search(lines[j+1])
                    if h:
                        codes_dict[code]['hrs_total'] = num(h.group(1))
                        codes_dict[code]['hrs_reg'] = num(h.group(2))
                        codes_dict[code]['hrs_ot'] = num(h.group(3))
                        codes_dict[code]['hrs_dt'] = num(h.group(4))
                        codes_dict[code]['hrs_other'] = num(h.group(5))
                i = j
                break
    i += 1

# Workers: "PR   178 03/11/16 GE69    Gerard, Jeffrey S" ... next line "... Regular: 1.00 hours   43.00"
workers = defaultdict(lambda: {'name': '', 'hours': 0.0, 'amount': 0.0, 'days': set()})
pr_re = re.compile(r'^PR\s+\d+\s+(\d{2}/\d{2}/\d{2})\s+([A-Z]{2}\d{2})\s+(.+?)\s*$')
hour_line_re = re.compile(r'^\s+(\d{2}/\d{2}/\d{2})\s+Regular:\s+([\d.]+)\s+hours\s+([\d,.\-]+)')

for i, line in enumerate(lines):
    m = pr_re.match(line)
    if m and i+1 < len(lines):
        wid = m.group(2)
        wname = m.group(3).strip()
        h = hour_line_re.match(lines[i+1])
        if h:
            hrs = num(h.group(2))
            amt = num(h.group(3))
            workers[wid]['name'] = wname
            workers[wid]['hours'] += hrs
            workers[wid]['amount'] += amt
            workers[wid]['days'].add(h.group(1))

# AP vendors: "AP    123 04/04/16  ROSE    Rosen Supply ... Invoice XXX  1,234.56"
# multiline. Pattern: "AP  NNN MM/DD/YY  VENDORID    VendorName"  next line "... Invoice NNN  amount  net_due  ret"
ap_re = re.compile(r'^AP\s+\d+\s+(\d{2}/\d{2}/\d{2})\s+(\S+)\s+(.+?)\s*$')
ap_detail_re = re.compile(r'^\s+\d{2}/\d{2}/\d{2}\s+(.+?)\s+([\d,.\-]+)\s*$')

vendors = defaultdict(lambda: {'name': '', 'total': 0.0, 'count': 0, 'invoices': []})
for i, line in enumerate(lines):
    m = ap_re.match(line)
    if m and i+1 < len(lines):
        vid = m.group(2)
        vname = m.group(3).strip()
        detail = lines[i+1]
        d = ap_detail_re.match(detail)
        if d:
            # amount is last number
            # Match numbers preceded by whitespace (avoid "S7745594.001" being read as 7,745,594)
            nums = re.findall(r'(?:^|\s)([\d,]+\.\d{2}-?)', detail)
            if nums:
                amt = num(nums[-1])  # last whitespace-prefixed number is the amount
                vendors[vid]['name'] = vname
                vendors[vid]['total'] += amt
                vendors[vid]['count'] += 1
                inv_ref = d.group(1).strip()
                vendors[vid]['invoices'].append((m.group(1), inv_ref, amt))

# AR invoices
ar_re = re.compile(r'Invoice\s+(\d+)\s+([\d,.\-]+)\s+([\d,.\-]+)?')
invoices = defaultdict(lambda: {'total': 0.0, 'retainage': 0.0, 'date': '', 'lines': 0})
ar_header_re = re.compile(r'^AR\s+\d+\s+(\d{2}/\d{2}/\d{2})\s+\S+\s+')
for i, line in enumerate(lines):
    h = ar_header_re.match(line)
    if h and i+1 < len(lines):
        date = h.group(1)
        inv_line = lines[i+1]
        m = re.search(r'Invoice\s+(\d+)\s+([\d,.\-]+)\s+([\d,.\-]+)?', inv_line)
        if m:
            inv = m.group(1)
            amt = num(m.group(2))
            ret = num(m.group(3)) if m.group(3) else 0
            if not invoices[inv]['date']:
                invoices[inv]['date'] = date
            invoices[inv]['total'] += amt
            invoices[inv]['retainage'] += ret
            invoices[inv]['lines'] += 1

# Summary
print("=== COST CODES ===")
total_actual = 0
for code in sorted(codes_dict.keys()):
    c = codes_dict[code]
    total_actual += c['actual']
    print(f"{code} {c['desc'][:30]:30} Orig={c['orig']:>12,.2f} Rev={c['rev']:>12,.2f} Act={c['actual']:>12,.2f} Hrs={c['hrs_total']:>7.2f}")
print(f"Total actual (incl AR line): {total_actual:,.2f}")

print("\n=== WORKERS ===")
for wid, w in sorted(workers.items(), key=lambda x: -x[1]['hours']):
    print(f"{wid} {w['name'][:28]:28} Hrs={w['hours']:>8.2f} Amt={w['amount']:>10,.2f} Days={len(w['days'])}")

print(f"\nTotal workers: {len(workers)}")

print("\n=== VENDORS (top 15) ===")
for vid, v in sorted(vendors.items(), key=lambda x: -x[1]['total'])[:15]:
    print(f"{vid:10} {v['name'][:35]:35} Total={v['total']:>12,.2f} #={v['count']}")
print(f"Total unique vendors: {len(vendors)}")

print("\n=== INVOICES (count/sum) ===")
ar_total = 0
for inv in sorted(invoices.keys()):
    iv = invoices[inv]
    # Sum net of sign for net AR
    ar_total += iv['total']
print(f"Unique invoices: {len(invoices)}")
print(f"AR total (sum of all lines, signed): {ar_total:,.2f}")
print(f"Invoice list: {sorted(invoices.keys())}")

# Save
out = {
    'codes': codes_dict,
    'workers': {k: {**v, 'days': len(v['days'])} for k, v in workers.items()},
    'vendors': {k: {'name': v['name'], 'total': v['total'], 'count': v['count']} for k, v in vendors.items()},
    'invoices': {k: {'date': v['date'], 'total': v['total'], 'retainage': v['retainage'], 'lines': v['lines']} for k, v in invoices.items()},
}
with open('/sessions/keen-determined-mccarthy/work/2026_data.json', 'w') as f:
    json.dump(out, f, indent=2, default=str)
print("Saved /sessions/keen-determined-mccarthy/work/2026_data.json")
