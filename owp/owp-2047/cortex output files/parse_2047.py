#!/usr/bin/env python3
"""
parse_2047.py — Parse Job 2047 JDR PDF → 2047_data.json

Structured parse of the 94-page Sage Timberline Job Detail Report for
Job 2047 (GRE Metro Line Flats). Mirrors Sam's v4 Test-Labels
schema field-for-field so downstream Cortex builders can consume the same shape.

Customer code: 2047SH · Client: Boren and Spruce, LLC · GC: GRE.
"""
import json
import re
from pathlib import Path

import pdfplumber

PDF = Path('/sessions/gracious-relaxed-pascal/2047_Job_Detail_Report.pdf')
OUT = Path('/sessions/gracious-relaxed-pascal/2047_data.json')

PROJECT_META = {
    'job_number': '2047',
    'job_name': 'GRE Greenlake',
    'general_contractor': 'GRE',
    'report_date': '2026-04-03',
}

CAT = {}
for c in range(100, 200): CAT[str(c)] = 'Labor'
for c in range(200, 300): CAT[str(c)] = 'Material'
for c in range(600, 700): CAT[str(c)] = 'Overhead'
CAT.update({'995': 'Burden', '998': 'Burden', '999': 'Revenue'})


def _num(s):
    if s is None or s == '':
        return 0.0
    s = s.strip().replace(',', '')
    neg = s.endswith('-')
    if neg:
        s = s[:-1]
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


CODE_HDR_RE = re.compile(r'^(\d{3}) ([A-Z][A-Za-z &/\-.,]+)$')
PAGE_HDR_MARKERS = (
    'OWP, LLC Page:', 'Job Detail Report', 'for Job "2047" only', 'Job Number:2047',
    'Job Name: GRE Greenlake', 'Legal:',
    'Code Description Org Budget Rev.Budget',
    'Src Ref # Post / Doc Number Amount Net Due Retainage',
)
COST_TOTALS_RE = re.compile(
    r'^Cost Code Totals\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)$'
)
PAYROLL_HRS_RE = re.compile(
    r'Payroll Hours:\s+([\d,.\-]+)\s+\(\s+Reg:\s+([\d,.\-]+)\s+O/T:\s+([\d,.\-]+)\s+D/T:\s+([\d,.\-]+)\s+Other:\s+([\d,.\-]+)\s+\)'
)
PR1_RE = re.compile(r'^PR\s+(\S+)\s+(\d{2}/\d{2}/\d{2})\s+(\S+)\s+(.+)$')
PR2_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{2})\s+(Regular|Overtime|Doubletime|Other):\s+([\d,.\-]+)\s+hours\s+([\d,.\-]+)\s+Ck\s*#:\s*(\S+)$'
)
PR2_OH_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{2})\s+Overhead\s+(\d+)\s+([\d,.\-]+)$'
)
AP1_RE = re.compile(r'^AP\s+(\S+)\s+(\d{2}/\d{2}/\d{2})\s+(\S+)\s+(.+)$')
AP2_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{2})\s+(Inv|Credit):\s+(\S+)\s+([\d,.\-]+)(?:\s+([\d,.\-]+))?(?:\s+([\d,.\-]+))?$'
)
GL1_RE = re.compile(r'^GL\s+(\S+)\s+(\d{2}/\d{2}/\d{2})\s+(\S+)(?:\s+(.+))?$')
GL2_RE = re.compile(r'^(\d{2}/\d{2}/\d{2})(?:\s+(.+))?$')
GL3_RE = re.compile(r'^([\d,.\-]+)$')
AR1_RE = re.compile(r'^AR\s+(\S+)\s+(\d{2}/\d{2}/\d{2})\s+(\S+)\s+(.+)$')
AR2_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{2})\s+(Invoice|Credit|Payment)\s+(\S+)\s+([\d,.\-]+)(?:\s+([\d,.\-]+))?$'
)
JOB_TOTALS_RE = re.compile(
    r'^Job Totals Revenues:\s+([\d,.\-]+)\s+Expenses:\s+([\d,.\-]+)\s+Net:\s+([\d,.\-]+)\s+([\d,.\-]+)\s+([\d,.\-]+)$'
)
BY_SOURCE_RE = re.compile(
    r'^by Source:\s+GL:\s+([\d,.\-]+)\s+AP:\s+([\d,.\-]+)\s+PR:\s+([\d,.\-]+)\s+AR:\s+([\d,.\-]+)$'
)


def is_page_header(s):
    return any(m in s for m in PAGE_HDR_MARKERS)


def load_pdf_lines():
    raw = []
    with pdfplumber.open(str(PDF)) as pdf:
        for pg in pdf.pages:
            txt = pg.extract_text() or ''
            raw.extend(txt.split('\n'))
    return raw


def parse():
    lines = load_pdf_lines()
    clean = [L for L in lines if not is_page_header(L)]

    line_items = []
    code_summaries = []
    job_totals = {}
    current_code = None
    current_desc = None
    i = 0
    while i < len(clean):
        L = clean[i].rstrip()

        m = CODE_HDR_RE.match(L)
        if m and not is_page_header(L):
            current_code = m.group(1)
            current_desc = m.group(2).strip()
            i += 1
            continue

        m = COST_TOTALS_RE.match(L)
        if m:
            orig_b = _num(m.group(1))
            rev_b = _num(m.group(2))
            pm = _num(m.group(3))
            actual = _num(m.group(4))
            net_due = _num(m.group(5))
            retainage = _num(m.group(6))
            reg_h = ot_h = dt_h = other_h = 0.0
            if i + 1 < len(clean):
                hm = PAYROLL_HRS_RE.match(clean[i + 1].strip())
                if hm:
                    reg_h = _num(hm.group(2))
                    ot_h = _num(hm.group(3))
                    dt_h = _num(hm.group(4))
                    other_h = _num(hm.group(5))
                    i += 1
            code_summaries.append({
                'code': current_code, 'description': current_desc,
                'category': CAT.get(current_code, 'Unknown'),
                'original_budget': orig_b, 'current_budget': rev_b,
                'plus_minus_budget': pm, 'actual_amount': actual,
                'net_due': net_due, 'retainage': retainage,
                'regular_hours': reg_h, 'overtime_hours': ot_h,
                'doubletime_hours': dt_h, 'other_hours': other_h,
            })
            current_code = None
            current_desc = None
            i += 1
            continue

        jm = JOB_TOTALS_RE.match(L)
        if jm:
            job_totals = {
                'revenues': _num(jm.group(1)),
                'expenses': _num(jm.group(2)),
                'net': _num(jm.group(3)),
                'net_due': _num(jm.group(4)),
                'retainage': _num(jm.group(5)),
            }
            i += 1
            continue

        bm = BY_SOURCE_RE.match(L)
        if bm:
            job_totals['by_source'] = {
                'GL': _num(bm.group(1)), 'AP': _num(bm.group(2)),
                'PR': _num(bm.group(3)), 'AR': -abs(_num(bm.group(4))),
            }
            i += 1
            continue

        if current_code is not None:
            if L.startswith('PR ') and (i + 1) < len(clean):
                m1 = PR1_RE.match(L)
                nxt = clean[i + 1].strip()
                m2 = PR2_RE.match(nxt)
                m2_oh = PR2_OH_RE.match(nxt)
                if m1 and m2:
                    ref, post_date, code_id, name = m1.groups()
                    doc_date, kind, hrs, amt, ck = m2.groups()
                    reg_h = _num(hrs) if kind == 'Regular' else 0
                    ot_h = _num(hrs) if kind == 'Overtime' else 0
                    dt_h = _num(hrs) if kind == 'Doubletime' else 0
                    other_h = _num(hrs) if kind == 'Other' else 0
                    reg_a = _num(amt) if kind == 'Regular' else 0
                    ot_a = _num(amt) if kind == 'Overtime' else 0
                    actual = _num(amt)
                    line_items.append({
                        'cost_code': current_code, 'description': current_desc,
                        'source': 'PR', 'ref_number': ref, 'document_date': doc_date,
                        'posted_date': post_date, 'number': code_id, 'name': name.strip(),
                        'regular_hours': reg_h, 'overtime_hours': ot_h,
                        'doubletime_hours': dt_h, 'other_hours': other_h,
                        'regular_amount': reg_a, 'overtime_amount': ot_a,
                        'actual_amount': actual, 'check_number': ck,
                        'net_due': 0, 'retainage': 0,
                    })
                    i += 2
                    continue
                if m1 and m2_oh:
                    ref, post_date, code_id, name = m1.groups()
                    doc_date, oh_bucket, amt = m2_oh.groups()
                    line_items.append({
                        'cost_code': current_code, 'description': current_desc,
                        'source': 'PR', 'ref_number': ref, 'document_date': doc_date,
                        'posted_date': post_date, 'number': code_id, 'name': name.strip(),
                        'regular_hours': 0, 'overtime_hours': 0,
                        'doubletime_hours': 0, 'other_hours': 0,
                        'regular_amount': 0, 'overtime_amount': 0,
                        'actual_amount': _num(amt), 'check_number': None,
                        'overhead_bucket': int(oh_bucket),
                        'net_due': 0, 'retainage': 0,
                    })
                    i += 2
                    continue

            if L.startswith('AP ') and (i + 1) < len(clean):
                m1 = AP1_RE.match(L)
                m2 = AP2_RE.match(clean[i + 1].strip())
                if m1 and m2:
                    ref, post_date, vendor_id, vendor_name = m1.groups()
                    doc_date, kind, inv_num, amt, nd, ret = m2.groups()
                    line_items.append({
                        'cost_code': current_code, 'description': current_desc,
                        'source': 'AP', 'ref_number': ref, 'document_date': doc_date,
                        'posted_date': post_date, 'number': vendor_id,
                        'name': vendor_name.strip(),
                        'regular_hours': None, 'overtime_hours': None,
                        'regular_amount': None, 'overtime_amount': None,
                        'actual_amount': _num(amt), 'check_number': inv_num,
                        'net_due': _num(nd) if nd else 0,
                        'retainage': _num(ret) if ret else 0,
                    })
                    i += 2
                    continue

            if L.startswith('GL ') and (i + 2) < len(clean):
                m1 = GL1_RE.match(L)
                m2 = GL2_RE.match(clean[i + 1].strip())
                m3 = GL3_RE.match(clean[i + 2].strip())
                if m1 and m2 and m3:
                    ref, post_date, ref2, name = m1.groups()
                    name = name or ''
                    doc_date, gl_desc = m2.groups()
                    amt, = m3.groups()
                    line_items.append({
                        'cost_code': current_code, 'description': current_desc,
                        'source': 'GL', 'ref_number': ref, 'document_date': doc_date,
                        'posted_date': post_date, 'number': ref2, 'name': name.strip(),
                        'regular_hours': None, 'overtime_hours': None,
                        'regular_amount': None, 'overtime_amount': None,
                        'actual_amount': _num(amt), 'check_number': None,
                        'net_due': 0, 'retainage': 0,
                    })
                    i += 3
                    continue

            if L.startswith('AR ') and (i + 1) < len(clean):
                m1 = AR1_RE.match(L)
                m2 = AR2_RE.match(clean[i + 1].strip())
                if m1 and m2:
                    ref, post_date, client_id, client_name = m1.groups()
                    doc_date, kind, inv_num, amt, ret = m2.groups()
                    line_items.append({
                        'cost_code': current_code, 'description': current_desc,
                        'source': 'AR', 'ref_number': ref, 'document_date': doc_date,
                        'posted_date': post_date, 'number': client_id,
                        'name': client_name.strip(),
                        'regular_hours': None, 'overtime_hours': None,
                        'regular_amount': None, 'overtime_amount': None,
                        'actual_amount': _num(amt), 'check_number': inv_num,
                        'net_due': 0,
                        'retainage': _num(ret) if ret else 0,
                    })
                    i += 2
                    continue

        i += 1

    return line_items, code_summaries, job_totals


def build_reconciliation(line_items, code_summaries):
    out = []
    by_code = {}
    for it in line_items:
        by_code.setdefault(it['cost_code'], 0)
        by_code[it['cost_code']] += (it['actual_amount'] or 0)
    for cs in code_summaries:
        code = cs['code']
        pdf_total = round(cs['actual_amount'], 2)
        parsed = round(by_code.get(code, 0), 2)
        diff = round(parsed - pdf_total, 2)
        out.append({
            'code': code, 'description': cs['description'],
            'pdf_total': pdf_total, 'parsed_sum': parsed,
            'difference': diff, 'status': 'PASS' if abs(diff) < 0.02 else 'FAIL',
        })
    return out


def build_worker_wages(line_items):
    by_name = {}
    for it in line_items:
        if it['source'] != 'PR':
            continue
        n = it['name']
        w = by_name.setdefault(n, {
            'name': n, 'regular_hours': 0.0, 'overtime_hours': 0.0,
            'doubletime_hours': 0.0, 'other_hours': 0.0,
            'regular_amount': 0.0, 'overtime_amount': 0.0,
        })
        w['regular_hours'] += it.get('regular_hours') or 0
        w['overtime_hours'] += it.get('overtime_hours') or 0
        w['doubletime_hours'] += it.get('doubletime_hours') or 0
        w['other_hours'] += it.get('other_hours') or 0
        w['regular_amount'] += it.get('regular_amount') or 0
        w['overtime_amount'] += it.get('overtime_amount') or 0

    workers = []
    for w in by_name.values():
        rh = w['regular_hours']
        ra = w['regular_amount']
        w['nominal_rate'] = (ra / rh) if rh > 0 else None
        rate = w['nominal_rate']
        if rh == 0 and w['overtime_hours'] > 0:
            w['tier'] = 'OT-ONLY'
        elif rate is None:
            w['tier'] = 'OT-ONLY'
        elif rate < 20:
            w['tier'] = 'APPRENTICE/HELPER'
        elif rate < 36:
            w['tier'] = 'JOURNEYMAN'
        else:
            w['tier'] = 'LEAD/SUPERVISOR'
        workers.append(w)
    tier_ord = {'APPRENTICE/HELPER': 0, 'JOURNEYMAN': 1, 'LEAD/SUPERVISOR': 2, 'OT-ONLY': 3}
    workers.sort(key=lambda x: (tier_ord[x['tier']], x['nominal_rate'] or 999))
    return workers


def percentile(vals, p):
    if not vals:
        return 0
    vals = sorted(vals)
    k = (len(vals) - 1) * (p / 100)
    f = int(k)
    c = min(f + 1, len(vals) - 1)
    if f == c:
        return vals[f]
    return vals[f] + (vals[c] - vals[f]) * (k - f)


def build_derived(line_items, code_summaries, workers, job_totals):
    pr_amt = sum((it['actual_amount'] or 0) for it in line_items if it['source'] == 'PR')
    ap_amt = sum((it['actual_amount'] or 0) for it in line_items if it['source'] == 'AP')
    gl_amt = sum((it['actual_amount'] or 0) for it in line_items if it['source'] == 'GL')
    ar_amt = sum((it['actual_amount'] or 0) for it in line_items if it['source'] == 'AR')
    direct = pr_amt + ap_amt + gl_amt
    revenue = abs(ar_amt)

    total_hours = 0
    reg_hours_sum = 0
    reg_amt_sum = 0
    for it in line_items:
        if it['source'] != 'PR':
            continue
        total_hours += (it.get('regular_hours') or 0) + (it.get('overtime_hours') or 0) + \
                       (it.get('doubletime_hours') or 0) + (it.get('other_hours') or 0)
        reg_hours_sum += it.get('regular_hours') or 0
        reg_amt_sum += it.get('regular_amount') or 0

    burden_total = sum(cs['actual_amount'] for cs in code_summaries if cs['code'] in ('995', '998'))
    material_codes = [cs for cs in code_summaries if cs['category'] == 'Material']
    mat_spend = sum(cs['actual_amount'] for cs in material_codes)
    mat_budget = sum(cs['current_budget'] for cs in material_codes)
    mat_over = [cs for cs in material_codes if cs['plus_minus_budget'] > 0]
    mat_under = [cs for cs in material_codes if cs['plus_minus_budget'] <= 0]

    expense_codes = [cs for cs in code_summaries if cs['code'] != '999']
    phases_over = [cs for cs in expense_codes if cs['plus_minus_budget'] > 0]
    phases_under = [cs for cs in expense_codes if cs['plus_minus_budget'] <= 0]

    labor_codes = [cs for cs in code_summaries if cs['category'] == 'Labor']
    labor_jtd = sum(cs['actual_amount'] for cs in labor_codes)

    total_budget = sum(cs['current_budget'] for cs in expense_codes)
    nominal_rates = [w['nominal_rate'] for w in workers if w['nominal_rate'] is not None]

    ap_by_vendor = {}
    for it in line_items:
        if it['source'] == 'AP':
            ap_by_vendor[it['name']] = ap_by_vendor.get(it['name'], 0) + (it['actual_amount'] or 0)
    top_vendor = max(ap_by_vendor.items(), key=lambda x: x[1]) if ap_by_vendor else ('(none)', 0)

    largest_overrun = max(expense_codes, key=lambda x: x['plus_minus_budget']) if expense_codes else None
    largest_savings = min(expense_codes, key=lambda x: x['plus_minus_budget']) if expense_codes else None
    largest_mat_over = max(material_codes, key=lambda x: x['plus_minus_budget']) if material_codes else None
    largest_mat_save = min(material_codes, key=lambda x: x['plus_minus_budget']) if material_codes else None

    return {
        'cost_by_source_pr_amount': round(pr_amt, 2),
        'cost_by_source_ap_amount': round(ap_amt, 2),
        'cost_by_source_gl_amount': round(gl_amt, 2),
        'direct_cost': round(direct, 2),
        'net_profit': round(revenue - direct, 2),
        'pr_pct_of_revenue': round(pr_amt / revenue, 6) if revenue else 0,
        'ap_pct_of_revenue': round(ap_amt / revenue, 6) if revenue else 0,
        'direct_cost_pct_of_revenue': round(direct / revenue, 6) if revenue else 0,
        'total_labor_hours': round(total_hours, 2),
        'pr_src_cost_per_hr': round(pr_amt / total_hours, 4) if total_hours else 0,
        'fully_loaded_wage': round((pr_amt + burden_total) / total_hours, 4) if total_hours else 0,
        'burden_multiplier': round((pr_amt + burden_total) / pr_amt, 4) if pr_amt else 0,
        'straight_time_rate': round(reg_amt_sum / reg_hours_sum, 4) if reg_hours_sum else 0,
        'total_workers': len(workers),
        'nominal_wage_percentiles': {
            'p10': round(percentile(nominal_rates, 10), 2),
            'p25': round(percentile(nominal_rates, 25), 2),
            'p50': round(percentile(nominal_rates, 50), 2),
            'p75': round(percentile(nominal_rates, 75), 2),
            'p90': round(percentile(nominal_rates, 90), 2),
        },
        'burden_total': round(burden_total, 2),
        'material_spend_total': round(mat_spend, 2),
        'material_budget_total': round(mat_budget, 2),
        'material_codes_tracked': len(material_codes),
        'material_codes_over_budget': len(mat_over),
        'material_codes_under_budget': len(mat_under),
        'largest_material_overrun': f"{largest_mat_over['code']} {largest_mat_over['description']}: ${largest_mat_over['actual_amount']:,.0f} vs ${largest_mat_over['current_budget']:,.0f} budget (+${largest_mat_over['plus_minus_budget']:,.0f})" if largest_mat_over else None,
        'largest_material_savings': f"{largest_mat_save['code']} {largest_mat_save['description']}: ${largest_mat_save['actual_amount']:,.0f} vs ${largest_mat_save['current_budget']:,.0f} budget (${largest_mat_save['plus_minus_budget']:,.0f})" if largest_mat_save else None,
        'material_vendor_count': len(ap_by_vendor),
        'top_vendor_spend': f"{top_vendor[0]}: ${top_vendor[1]:,.0f} ({top_vendor[1] / ap_amt * 100:.0f}% of AP)" if ap_amt else '(none)',
        'phases_over_budget': len(phases_over),
        'phases_under_budget': len(phases_under),
        'largest_overrun': f"{largest_overrun['code']} {largest_overrun['description']}: ${largest_overrun['actual_amount']:,.0f} vs ${largest_overrun['current_budget']:,.0f} budget (+${largest_overrun['plus_minus_budget']:,.0f})" if largest_overrun else None,
        'largest_savings': f"{largest_savings['code']} {largest_savings['description']}: ${largest_savings['actual_amount']:,.0f} vs ${largest_savings['current_budget']:,.0f} budget (${largest_savings['plus_minus_budget']:,.0f})" if largest_savings else None,
        'total_jtd_cost': round(sum(cs['actual_amount'] for cs in expense_codes), 2),
        'total_budget': round(total_budget, 2),
        'overall_pct_budget_consumed': round(sum(cs['actual_amount'] for cs in expense_codes) / total_budget, 4) if total_budget else 0,
        'total_over_under_budget': round(total_budget - sum(cs['actual_amount'] for cs in expense_codes), 2),
        'labor_unit_cost_per_hr': round(labor_jtd / total_hours, 4) if total_hours else 0,
        'revenue_per_labor_hour': round(revenue / total_hours, 4) if total_hours else 0,
        'material_price_variance': round(mat_budget - mat_spend, 2),
    }


def main():
    print('parsing 2047 JDR PDF...')
    line_items, code_summaries, job_totals = parse()
    print(f'  line_items: {len(line_items)}')
    print(f'  cost_code_summaries: {len(code_summaries)}')
    print(f'  job_totals: {job_totals}')

    workers = build_worker_wages(line_items)
    print(f'  workers (PR-distinct): {len(workers)}')

    recon = build_reconciliation(line_items, code_summaries)
    passed = sum(1 for r in recon if r['status'] == 'PASS')
    print(f'  reconciliation: {passed}/{len(recon)} PASS')

    derived = build_derived(line_items, code_summaries, workers, job_totals)
    print(f'  PR hours: {derived["total_labor_hours"]:,.2f}')
    print(f'  Revenue: ${abs(job_totals.get("revenues", 0)):,.2f}')
    print(f'  Expenses: ${job_totals.get("expenses", 0):,.2f}')
    print(f'  Direct (parsed): ${derived["direct_cost"]:,.2f}')

    data = {
        'schema': 'CORTEX_V2.2',
        'schema_note': 'Parsed direct from 94-page 2047 JDR PDF. Mirrors Sam v4 test-labels structure.',
        'project': PROJECT_META,
        'report_record': {
            'job_number': '2047',
            'job_name': 'GRE Greenlake',
            'report_date': '2026-04-03',
            'line_items_count': len(line_items),
            'cost_code_count': len(code_summaries),
            'job_totals_revenue': -abs(job_totals.get('revenues', 0)),
            'job_totals_expenses': job_totals.get('expenses', 0),
            'job_totals_net': job_totals.get('net', 0),
            'job_totals_net_due': job_totals.get('net_due', 0),
            'job_totals_retainage': job_totals.get('retainage', 0),
            'job_totals_by_source': job_totals.get('by_source', {}),
        },
        'line_items': line_items,
        'cost_code_summaries': code_summaries,
        'worker_wages': workers,
        'derived_fields': derived,
        'reconciliation': recon,
    }

    OUT.write_text(json.dumps(data, indent=2, default=str))
    print(f'\nwrote {OUT}  ({OUT.stat().st_size:,} bytes)')


if __name__ == '__main__':
    main()
