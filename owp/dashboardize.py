#!/usr/bin/env python3
"""
dashboardize.py — Generate the 16 dashboard JS data arrays for live projects.

Reads each project's _data.json + build script META, produces JS blocks
matching the format used by closed jobs in index.html.

Usage:
    python3 dashboardize.py          # prints JS blocks for all live jobs
    python3 dashboardize.py 2098     # prints JS block for one job
"""
import json
import sys
from collections import defaultdict
from pathlib import Path

_HERE = Path(__file__).resolve().parent

LABOR_CODES = {"011","100","101","110","111","112","113","120","130","140","141","142","143","145","150"}
MATERIAL_CODES = {"039","210","211","212","213","220","230","240","241","242","243","244","245","251"}
OVERHEAD_CODES = {"600","601","602","603","604","607"}
BURDEN_CODES = {"995","998"}

PHASE_NAMES = {
    "011": "Project Management", "100": "Supervision", "101": "Takeoff & Purchasing",
    "110": "Underground", "111": "Garage", "112": "Canout", "113": "Foundation Drain",
    "120": "Roughin", "130": "Finish", "140": "Gas", "141": "Water Main / Insulation",
    "142": "Mech Room", "143": "Condensation Drains", "145": "Fire Stopping", "150": "Warranty",
}

# ── META dicts for each live job (import from build scripts would be ideal,
#    but we inline them here for self-containment) ──

METAS = {}

def load_meta_from_build(job_id):
    """Try to extract META dict from build script. Tries both `owp-XXXX-live`
    (live projects) and `owp-XXXX` (closed projects)."""
    for folder in (f"owp-{job_id}-live", f"owp-{job_id}"):
        build_path = _HERE / folder / "cortex output files" / f"build_{job_id}.py"
        if build_path.exists():
            break
    else:
        return None
    text = build_path.read_text()
    # Find the META dict
    import re
    m = re.search(rf"META_{job_id}\s*=\s*(\{{.*?\n\}})", text, re.DOTALL)
    if not m:
        return None
    try:
        # Safe eval with restricted builtins
        meta = eval(m.group(1), {"__builtins__": {}}, {})
        return meta
    except:
        return None


def load_data(job_id):
    """Load and normalize parsed data. Tries both `owp-XXXX-live` (live) and
    `owp-XXXX` (closed) folder layouts."""
    for folder in (f"owp-{job_id}-live", f"owp-{job_id}"):
        data_path = _HERE / folder / "cortex output files" / f"{job_id}_data.json"
        if data_path.exists():
            break
    else:
        print(f"ERROR: {job_id}_data.json not found in either owp-{job_id}-live or owp-{job_id}", file=sys.stderr)
        return None
    raw = json.loads(data_path.read_text())

    code_totals = {}
    for cs in raw.get('cost_code_summaries', []):
        if cs['code'] is None:
            continue
        code_totals[cs['code']] = {
            'desc': cs['description'],
            'orig': cs['original_budget'],
            'rev': cs['current_budget'],
            'actual': cs['actual_amount'],
            'variance': cs['plus_minus_budget'],
            'net_due': cs['net_due'],
            'retainage': cs['retainage'],
            'reg_hours': cs.get('regular_hours', 0),
            'ot_hours': cs.get('overtime_hours', 0),
        }

    transactions = []
    for li in raw.get('line_items', []):
        hours = (li.get('regular_hours') or 0) + (li.get('overtime_hours') or 0) + \
                (li.get('doubletime_hours') or 0) + (li.get('other_hours') or 0)
        transactions.append({
            'src': li['source'],
            'party_name': li.get('name', ''),
            'amount': li.get('actual_amount', 0),
            'hours': hours,
            'code': li.get('cost_code', ''),
            'date': li.get('post_date', ''),
            'ref': li.get('reference', ''),
        })

    by_source = defaultdict(float)
    for t in transactions:
        by_source[t['src']] += t.get('amount', 0) or 0

    worker_wages = raw.get('worker_wages', [])
    derived = raw.get('derived_fields', {})
    job_totals = raw.get('totals', {})

    return {
        'code_totals': code_totals,
        'transactions': transactions,
        'by_source': dict(by_source),
        'worker_wages': worker_wages,
        'derived_fields': derived,
        'job_totals': job_totals,
        'raw': raw,
    }


def gen_crew_roster(data):
    """crewRoster: [[name, hours, earnings, rate, tier], ...]"""
    ww = data.get('worker_wages', [])
    roster = []
    for w in ww:
        name = w.get('name', '')
        # Support both formats: total_hours or regular_hours + overtime_hours
        hours = w.get('total_hours', 0)
        if not hours:
            hours = (w.get('regular_hours', 0) or 0) + (w.get('overtime_hours', 0) or 0) + \
                    (w.get('doubletime_hours', 0) or 0) + (w.get('other_hours', 0) or 0)
        earnings = w.get('total_earnings', 0)
        if not earnings:
            earnings = (w.get('regular_amount', 0) or 0) + (w.get('overtime_amount', 0) or 0)
        rate = round(earnings / hours, 2) if hours else 0
        if rate >= 28: tier = "Senior"
        elif rate >= 18: tier = "Mid"
        else: tier = "Apprentice"
        roster.append([name, hours, round(earnings), rate, tier])
    roster.sort(key=lambda x: -x[1])  # sort by hours desc
    return roster


def gen_tier_dist(roster):
    """tierDist: [[tier, count, hours, earnings, pct], ...]"""
    tiers = defaultdict(lambda: {'count': 0, 'hours': 0, 'earnings': 0})
    total_h = sum(r[1] for r in roster)
    for r in roster:
        t = r[4]
        tiers[t]['count'] += 1
        tiers[t]['hours'] += r[1]
        tiers[t]['earnings'] += r[2]
    result = []
    for tier in ["Senior", "Mid", "Apprentice"]:
        if tier in tiers:
            d = tiers[tier]
            pct = round(d['hours'] / total_h * 100, 1) if total_h else 0
            result.append([tier, d['count'], d['hours'], d['earnings'], pct])
    return result


def gen_wage_stats(roster):
    """wageStats: {bl, fl, bm, mn, mx, ah}"""
    if not roster:
        return {"bl": 0, "fl": 0, "bm": 1, "mn": 0, "mx": 0, "ah": 0}
    total_h = sum(r[1] for r in roster)
    total_e = sum(r[2] for r in roster)
    rates = [r[3] for r in roster if r[3] > 0]
    bl = round(total_e / total_h, 2) if total_h else 0
    mn = min(rates) if rates else 0
    mx = max(rates) if rates else 0
    fl = mx  # floor = max rate for senior
    bm = round(mx / mn, 3) if mn else 1
    ah = round(total_h / len(roster), 1) if roster else 0
    return {"bl": bl, "fl": fl, "bm": bm, "mn": mn, "mx": mx, "ah": ah}


def gen_all_vendors(data):
    """allVendors: [[vendor, invoices, spend, pct], ...]"""
    by_vendor = defaultdict(lambda: {'invoices': 0, 'spend': 0})
    for t in data.get('transactions', []):
        if t.get('src') == 'AP':
            name = (t.get('party_name') or '').strip()
            if not name:
                continue
            by_vendor[name]['invoices'] += 1
            by_vendor[name]['spend'] += t.get('amount', 0) or 0
    total_ap = sum(v['spend'] for v in by_vendor.values()) or 1
    result = []
    for name, d in sorted(by_vendor.items(), key=lambda x: -x[1]['spend']):
        pct = round(d['spend'] / total_ap * 100, 1)
        result.append([name, d['invoices'], round(d['spend']), pct])
    return result


def gen_phases(data, units):
    """phases: [[name, code, hours, pct, per_unit], ...]"""
    codes = data.get('code_totals', {})
    total_hours = 0
    phase_data = []
    for code in sorted(codes.keys()):
        if code not in LABOR_CODES or code == '999':
            continue
        info = codes[code]
        h = (info.get('reg_hours', 0) or 0) + (info.get('ot_hours', 0) or 0)
        if h > 0:
            total_hours += h
            phase_data.append((code, h))
    result = []
    for code, h in phase_data:
        name = PHASE_NAMES.get(code, codes[code].get('desc', code))
        pct = round(h / total_hours * 100, 1) if total_hours else 0
        per_unit = round(h / units, 2) if units else 0
        result.append([name, code, h, pct, per_unit])
    return result


def gen_cost_codes(data):
    """costCodes: [[code, desc, budget, actual, variance], ...]
    For live projects, budget = revised budget."""
    codes = data.get('code_totals', {})
    result = []
    for code in sorted(codes.keys()):
        if code == '999':
            continue
        info = codes[code]
        desc = info.get('desc', '')
        budget = info.get('rev', 0) or 0
        actual = info.get('actual', 0) or 0
        variance = budget - actual
        result.append([code, desc, round(budget), round(actual), round(variance)])
    return result


def gen_cost_cats(data, units):
    """costCats: [[category, count, budget, actual, pct, per_unit], ...]"""
    codes = data.get('code_totals', {})
    cats = defaultdict(lambda: {'count': 0, 'budget': 0, 'actual': 0})
    total_actual = 0
    for code in codes:
        if code == '999':
            continue
        info = codes[code]
        budget = info.get('rev', 0) or 0
        actual = info.get('actual', 0) or 0
        total_actual += actual
        if code in LABOR_CODES:
            cat = "Labor"
        elif code in MATERIAL_CODES:
            cat = "Material"
        elif code in OVERHEAD_CODES:
            cat = "Overhead"
        elif code in BURDEN_CODES:
            cat = "Burden"
        else:
            cat = "Other"
        cats[cat]['count'] += 1
        cats[cat]['budget'] += budget
        cats[cat]['actual'] += actual
    result = []
    for cat in ["Labor", "Material", "Overhead", "Burden", "Other"]:
        if cat in cats:
            d = cats[cat]
            pct = round(d['actual'] / total_actual * 100, 1) if total_actual else 0
            per_unit = round(d['actual'] / units) if units else 0
            result.append([cat, d['count'], round(d['budget']), round(d['actual']), pct, per_unit])
    return result


def gen_bva(data):
    """bva: [[label, budget, actual, variance, pct, null, null, status], ...]
    Uses revised budget for live projects."""
    codes = data.get('code_totals', {})
    result = []
    for code in sorted(codes.keys()):
        if code == '999':
            continue
        info = codes[code]
        desc = info.get('desc', '')
        budget = info.get('rev', 0) or 0
        actual = info.get('actual', 0) or 0
        variance = actual - budget
        if budget == 0:
            pct = 100.0 if actual > 0 else 0
            status = "UNBUDGETED" if actual > 0 else "ON"
        else:
            pct = round(variance / abs(budget) * 100, 1)
            if pct > 50:
                status = "CRITICAL"
            elif pct > 10:
                status = "OVER"
            elif pct < -10:
                status = "UNDER"
            else:
                status = "ON"
        label = f"{code} \u00b7 {desc}"
        result.append([label, round(budget), round(actual), round(variance), pct, None, None, status])
    return result


def gen_sov_data(meta, data):
    """sovData: {originalContract, changeOrders, finalContract, retainage, netPaid}"""
    codes = data.get('code_totals', {})
    rev_contract = abs(codes.get('999', {}).get('rev', 0))
    orig_contract = abs(codes.get('999', {}).get('orig', 0))
    co_net = rev_contract - orig_contract
    billed = meta.get('billed', 0)
    retention = meta.get('retention', 0)
    net_paid = billed - retention
    return {
        "originalContract": round(orig_contract),
        "changeOrders": round(co_net),
        "finalContract": round(rev_contract),
        "retainage": round(retention),
        "netPaid": round(net_paid)
    }


def gen_insights(meta, data, roster, vendors, cost_cats):
    """insights: [[title, body], ...] — 10 narrative insights."""
    codes = data.get('code_totals', {})
    rev_contract = abs(codes.get('999', {}).get('rev', 0))
    rev_expense = sum(codes[c].get('rev', 0) for c in codes if c != '999')
    actual_expense = sum(codes[c].get('actual', 0) for c in codes if c != '999')
    profit = rev_contract - actual_expense
    margin = profit / rev_contract * 100 if rev_contract else 0
    units = meta.get('units', 1) or 1
    fixtures = meta.get('total_fixtures', 0)
    billed = meta.get('billed', 0)
    retention = meta.get('retention', 0)

    total_hours = sum(r[1] for r in roster) if roster else 0
    total_workers = len(roster)
    total_ap = sum(v[2] for v in vendors) if vendors else 0

    insights = []

    # 1. Margin profile
    insights.append(["MARGIN PROFILE",
        f"{'Forecast' if meta.get('billed',0) < rev_contract else 'Net'} profit of ${profit:,.0f} on ${rev_contract:,.0f} revenue = {margin:.1f}% gross margin. "
        f"Direct cost ${actual_expense:,.0f} absorbs {100-margin:.1f}% of revenue."])

    # 2. Vendor concentration
    if vendors:
        top = vendors[0]
        insights.append(["VENDOR CONCENTRATION",
            f"{top[0]} alone = ${top[2]:,.0f} ({top[3]}% of all AP) across {top[1]} invoices. "
            f"Total {len(vendors)} AP vendors, ${total_ap:,.0f} total AP spend."])

    # 3. Crew profile
    if roster:
        top3_h = sum(r[1] for r in roster[:3])
        top3_pct = top3_h / total_hours * 100 if total_hours else 0
        insights.append(["CREW PROFILE",
            f"{total_workers} workers for {total_hours:,.0f} hours = {total_hours/total_workers:.0f} avg hrs/worker. "
            f"Top 3 workers logged {top3_h:,.0f} hrs combined ({top3_pct:.0f}% of all labor)."])

    # 4. Roughin dominance
    rl = codes.get('120', {})
    rl_h = (rl.get('reg_hours', 0) or 0) + (rl.get('ot_hours', 0) or 0)
    rl_pct = rl_h / total_hours * 100 if total_hours else 0
    insights.append(["ROUGHIN DOMINANCE",
        f"Code 120 Roughin Labor = {rl_h:,.0f} hrs ({rl_pct:.0f}% of all labor) and ${rl.get('actual',0):,.0f} actual. "
        f"Code 230 Finish Material = ${codes.get('230',{}).get('actual',0):,.0f}."])

    # 5. Budget delivery
    bud_var = rev_expense - actual_expense
    bud_var_pct = bud_var / rev_expense * 100 if rev_expense else 0
    insights.append(["BUDGET DELIVERY",
        f"Revised expense budget was ${rev_expense:,.0f}; actual direct cost ${actual_expense:,.0f} → "
        f"${bud_var:,.0f} variance ({bud_var_pct:+.1f}%)."])

    # 6. Contract growth
    orig = abs(codes.get('999', {}).get('orig', 0))
    co_net = rev_contract - orig
    co_pct = co_net / orig * 100 if orig else 0
    insights.append(["CONTRACT GROWTH",
        f"Contract {'grew' if co_net >= 0 else 'shrank'} from ${orig:,.0f} to ${rev_contract:,.0f} — "
        f"net {'+' if co_net >= 0 else ''}${co_net:,.0f} ({co_pct:+.1f}%). "
        f"Retainage of ${retention:,.0f} ({retention/billed*100:.0f}%) {'held' if retention > 0 else 'released'}." if billed else
        f"net {'+' if co_net >= 0 else ''}${co_net:,.0f} ({co_pct:+.1f}%)."])

    # 7. Burden ratio
    labor_actual = sum(codes[c].get('actual', 0) for c in codes if c in LABOR_CODES)
    burden_actual = sum(codes[c].get('actual', 0) for c in codes if c in BURDEN_CODES)
    burden_ratio = burden_actual / labor_actual * 100 if labor_actual else 0
    insights.append(["BURDEN RATIO",
        f"Burden codes 995 + 998 = ${burden_actual:,.0f} on ${labor_actual:,.0f} labor base = {burden_ratio:.1f}%. "
        f"{'In line with' if 45 < burden_ratio < 55 else 'Outside'} OWP's typical 48–50% labor burden."])

    # 8. Project scale
    rev_per_unit = rev_contract / units if units else 0
    h_per_unit = total_hours / units if units else 0
    profit_per_unit = profit / units if units else 0
    insights.append(["PROJECT SCALE",
        f"{units} units · {fixtures or '—'} fixtures. Revenue/unit ${rev_per_unit:,.0f}, "
        f"hours/unit {h_per_unit:.1f}, profit/unit ${profit_per_unit:,.0f}."])

    # 9. Productivity
    bl_rate = sum(r[2] for r in roster) / total_hours if total_hours else 0
    rev_per_hour = rev_contract / total_hours if total_hours else 0
    profit_per_hour = profit / total_hours if total_hours else 0
    insights.append(["PRODUCTIVITY",
        f"Blended labor rate ${bl_rate:.2f}/hr. Revenue/hour ${rev_per_hour:.0f}. "
        f"Profit/hour ${profit_per_hour:.0f}."])

    # 10. Fixture density
    if fixtures:
        mat_actual = sum(codes[c].get('actual', 0) for c in codes if c in MATERIAL_CODES)
        insights.append(["FIXTURE DENSITY",
            f"{fixtures} fixtures ÷ {units} units = {fixtures/units:.2f} fixtures/unit. "
            f"Material/fixture ${mat_actual/fixtures:.0f}, labor/fixture ${labor_actual/fixtures:.0f}."])
    else:
        insights.append(["CO PROFILE",
            f"{meta.get('executed_co_count', 0)} executed COs, {meta.get('cor_count', 0)} CORs. "
            f"Net CO impact ${co_net:,.0f} ({co_pct:+.1f}%)."])

    return insights


def gen_predictive_signals(meta, data):
    """predictiveSignals: [[label, value, threshold, status], ...]"""
    codes = data.get('code_totals', {})
    signals = []

    # BVA flags
    critical_count = 0
    over_count = 0
    for code in sorted(codes.keys()):
        if code == '999': continue
        info = codes[code]
        rev = info.get('rev', 0) or 0
        actual = info.get('actual', 0) or 0
        if rev == 0:
            if actual > 10000:
                critical_count += 1
        else:
            pct = (actual - rev) / rev
            if pct > 0.5: critical_count += 1
            elif pct > 0.1: over_count += 1

    signals.append(["BVA Critical Flags", str(critical_count), "<2", "ELEVATED" if critical_count > 1 else "HEALTHY"])
    signals.append(["BVA Over Flags", str(over_count), "<5", "ELEVATED" if over_count > 4 else "HEALTHY"])

    # CO count
    co_count = meta.get('executed_co_count', 0)
    signals.append(["Change Order Count", str(co_count), "<10", "ELEVATED" if co_count > 10 else "HEALTHY"])

    # Retention
    billed = meta.get('billed', 0)
    retention = meta.get('retention', 0)
    ret_pct = retention / billed * 100 if billed else 0
    signals.append(["Retention Held", f"${retention:,.0f}", "<6%", "HEALTHY" if ret_pct < 6 else "ELEVATED"])

    # Margin
    rev_contract = abs(codes.get('999', {}).get('rev', 0))
    rev_expense = sum(codes[c].get('rev', 0) for c in codes if c != '999')
    margin = (rev_contract - rev_expense) / rev_contract * 100 if rev_contract else 0
    signals.append(["Forecast Margin", f"{margin:.1f}%", ">25%", "HEALTHY" if margin > 25 else "ELEVATED"])

    # RFI count
    rfi = meta.get('rfi_count', 0)
    if rfi:
        signals.append(["RFI Count", str(rfi), "<100", "ELEVATED" if rfi > 100 else "HEALTHY"])

    return signals


def gen_change_meta(meta):
    """changeMeta: {total, costImpact, types: {CO: n, ...}}"""
    cos = meta.get('cos', [])
    change_log = meta.get('change_log', [])
    cost_impact = sum(co.get('amount', 0) for co in cos)
    types = defaultdict(int)
    for entry in change_log:
        if isinstance(entry, dict):
            types[entry.get('type', 'CO')] += 1
        elif isinstance(entry, (list, tuple)):
            types[entry[1] if len(entry) > 1 else 'CO'] += 1
    if not types and cos:
        types['CO'] = len(cos)
    return {
        "total": len(change_log) or len(cos),
        "costImpact": round(cost_impact),
        "types": dict(types)
    }


def gen_change_log(meta):
    """changeLog: [[ref, type, date, desc, responsible, amount, 0], ...]"""
    change_log = meta.get('change_log', [])
    cos = meta.get('cos', [])
    result = []

    if change_log:
        for entry in change_log:
            if isinstance(entry, dict):
                result.append([
                    entry.get('ref', ''),
                    entry.get('type', 'CO'),
                    entry.get('date', ''),
                    entry.get('desc', ''),
                    "GC",
                    entry.get('amount', 0),
                    0
                ])
    elif cos:
        for co in cos:
            co_ref = co['co']
            if isinstance(co_ref, int):
                co_ref_str = f"CO-{co_ref:03d}"
            else:
                co_ref_str = str(co_ref)
            result.append([
                co_ref_str,
                "CO",
                "",
                co.get('desc', ''),
                "GC",
                co.get('amount', 0),
                0
            ])
    return result


def gen_pay_apps(meta):
    """payApps from meta — many live projects don't have detailed pay app data."""
    return meta.get('pay_apps', [])


def gen_root_causes(meta):
    """rootCauses: [[cause, count, amount, responsible], ...]"""
    # For live projects, derive from COs if available
    cos = meta.get('cos', [])
    if not cos:
        return []
    # Simple categorization
    causes = defaultdict(lambda: {'count': 0, 'amount': 0})
    for co in cos:
        desc = co.get('desc', '').lower()
        if 'asi' in desc or 'design' in desc or 'revision' in desc or 'revised' in desc:
            cat = "Design change"
        elif 'rfi' in desc:
            cat = "Field condition (RFI)"
        elif 'credit' in desc or 'backcharge' in desc or 'deduct' in desc:
            cat = "Credit/backcharge"
        elif 'bond' in desc or 'permit' in desc or 'fee' in desc:
            cat = "Administrative"
        else:
            cat = "Owner/GC directive"
        causes[cat]['count'] += 1
        causes[cat]['amount'] += co.get('amount', 0)
    result = []
    for cat in sorted(causes.keys(), key=lambda x: -abs(causes[x]['amount'])):
        d = causes[cat]
        resp = "Designer" if "Design" in cat else "GC" if "Field" in cat else "Mixed"
        result.append([cat, d['count'], round(d['amount']), resp])
    return result


def gen_responsibility(root_causes):
    """responsibility: [[responsible, count, amount], ...]"""
    by_resp = defaultdict(lambda: {'count': 0, 'amount': 0})
    for rc in root_causes:
        by_resp[rc[3]]['count'] += rc[1]
        by_resp[rc[3]]['amount'] += rc[2]
    return [[r, d['count'], d['amount']] for r, d in sorted(by_resp.items(), key=lambda x: -abs(x[1]['amount']))]


def js_val(v):
    """Convert Python value to JS literal."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, str):
        return json.dumps(v)
    if isinstance(v, (int, float)):
        return str(v)
    if isinstance(v, dict):
        return json.dumps(v, separators=(',', ':'))
    if isinstance(v, (list, tuple)):
        return json.dumps(v, separators=(',', ':'))
    return json.dumps(v)


def generate_js_block(job_id, meta, data):
    """Generate all 16 data array assignments for a project."""
    units = meta.get('units', 1) or 1

    roster = gen_crew_roster(data)
    tier_dist = gen_tier_dist(roster)
    wage_stats = gen_wage_stats(roster)
    vendors = gen_all_vendors(data)
    phases = gen_phases(data, units)
    cost_codes = gen_cost_codes(data)
    cost_cats = gen_cost_cats(data, units)
    bva = gen_bva(data)
    sov_data = gen_sov_data(meta, data)
    insights = gen_insights(meta, data, roster, vendors, cost_cats)
    pay_apps = gen_pay_apps(meta)
    change_log_arr = gen_change_log(meta)
    root_causes = gen_root_causes(meta)
    responsibility = gen_responsibility(root_causes)
    pred_signals = gen_predictive_signals(meta, data)
    change_meta = gen_change_meta(meta)

    # Also compute the num fields that are null in the PROJECTS object
    codes = data.get('code_totals', {})
    labor_actual = sum(codes[c].get('actual', 0) for c in codes if c in LABOR_CODES)
    material_actual = sum(codes[c].get('actual', 0) for c in codes if c in MATERIAL_CODES)
    overhead_actual = sum(codes[c].get('actual', 0) for c in codes if c in OVERHEAD_CODES)
    burden_actual = sum(codes[c].get('actual', 0) for c in codes if c in BURDEN_CODES)
    total_hours = sum(r[1] for r in roster)
    total_workers = len(roster)
    total_ap = sum(v[2] for v in vendors)
    total_invoices = sum(v[1] for v in vendors)
    top_vendor = vendors[0][0] if vendors else '—'
    top_share = vendors[0][3] if vendors else 0

    lines = []
    lines.append(f"  // ── #{job_id} data arrays (auto-generated by dashboardize.py) ──")
    lines.append(f"  PROJECTS['{job_id}'].crewRoster = {js_val(roster)};")
    lines.append(f"  PROJECTS['{job_id}'].tierDist = {js_val(tier_dist)};")
    lines.append(f"  PROJECTS['{job_id}'].wageStats = {js_val(wage_stats)};")
    lines.append(f"  PROJECTS['{job_id}'].allVendors = {js_val(vendors)};")
    lines.append(f"  PROJECTS['{job_id}'].phases = {js_val(phases)};")
    lines.append(f"  PROJECTS['{job_id}'].costCodes = {js_val(cost_codes)};")
    lines.append(f"  PROJECTS['{job_id}'].costCats = {js_val(cost_cats)};")
    lines.append(f"  PROJECTS['{job_id}'].insights = {js_val(insights)};")
    lines.append(f"  PROJECTS['{job_id}'].bva = {js_val(bva)};")
    lines.append(f"  PROJECTS['{job_id}'].sovData = {js_val(sov_data)};")
    lines.append(f"  PROJECTS['{job_id}'].payApps = {js_val(pay_apps)};")
    lines.append(f"  PROJECTS['{job_id}'].changeLog = {js_val(change_log_arr)};")
    lines.append(f"  PROJECTS['{job_id}'].rootCauses = {js_val(root_causes)};")
    lines.append(f"  PROJECTS['{job_id}'].responsibility = {js_val(responsibility)};")
    lines.append(f"  PROJECTS['{job_id}'].predictiveSignals = {js_val(pred_signals)};")
    lines.append(f"  PROJECTS['{job_id}'].changeMeta = {js_val(change_meta)};")

    # Also emit num field patches
    lines.append(f"  // num field patches")
    lines.append(f"  if (PROJECTS['{job_id}'].num) {{")
    lines.append(f"    Object.assign(PROJECTS['{job_id}'].num, {{labor:{round(labor_actual)},material:{round(material_actual)},overhead:{round(overhead_actual)},burden:{round(burden_actual)},hours:{round(total_hours)},workers:{total_workers}}});")
    lines.append(f"  }}")
    # Vendor meta patches
    lines.append(f"  if (PROJECTS['{job_id}'].vendorMeta) {{")
    lines.append(f"    Object.assign(PROJECTS['{job_id}'].vendorMeta, {{count:{len(vendors)},invoices:{total_invoices},totalAp:{round(total_ap)},topVendor:{js_val(top_vendor)},topShare:{top_share}}});")
    lines.append(f"  }}")
    # Also populate vendors array in the PROJECTS object
    vendor_objs = [{"name": v[0], "invoices": v[1], "spend": v[2], "pct": v[3]} for v in vendors[:10]]
    lines.append(f"  PROJECTS['{job_id}'].vendors = {js_val(vendor_objs)};")

    return "\n".join(lines)


def main():
    jobs = sys.argv[1:] if len(sys.argv) > 1 else ["2098", "2103", "2104", "2105", "2106", "2107"]

    all_blocks = []
    for job_id in jobs:
        # Try -live folder first (active projects), then plain owp-XXXX (closed projects)
        for folder in (f"owp-{job_id}-live", f"owp-{job_id}"):
            data_path = _HERE / folder / "cortex output files" / f"{job_id}_data.json"
            if data_path.exists():
                break
        else:
            print(f"// SKIP {job_id}: no data.json", file=sys.stderr)
            continue

        meta = load_meta_from_build(job_id)
        if not meta:
            meta = {'job_id': job_id, 'units': 1, 'billed': 0, 'retention': 0}
            print(f"// WARN {job_id}: no META found in build script, using defaults", file=sys.stderr)

        data = load_data(job_id)
        if not data:
            continue

        block = generate_js_block(job_id, meta, data)
        all_blocks.append(block)
        print(f"// Generated {job_id}: {len(block)} chars", file=sys.stderr)

    print("\n".join(all_blocks))


if __name__ == "__main__":
    main()
