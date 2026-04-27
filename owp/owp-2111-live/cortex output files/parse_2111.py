#!/usr/bin/env python3
"""
parse_2111.py — Parse Job 2111 JDR PDF → 2111_data.json

Structured parse of the 209-page Sage Timberline Job Detail Report for
Job 2111 (Compass General, Northgate Station M3). Mirrors the per-project
parse convention established by parse_2098/parse_2103/parse_2104.

The underlying parsing logic is the generalized JDR parser shared with
2112 and earlier 2012 (see parse_live_jdr.py in workspace root). This
per-project script hardcodes the job_id, PDF path, and output path so
the full parse/build/notionize trio can be re-run from this folder.

Customer code: 2111CG · Client: Compass General Construction I, LLC ·
GC: Compass General · Owner: Simon Property Group.
"""
import json
import re
import sys
from pathlib import Path
import pdfplumber

_HERE = Path(__file__).resolve().parent
# owp-2111-live/cortex output files/ → owp/ → Gemini Accounting Reports/
PDF = _HERE.parent.parent / 'Gemini Accounting Reports' / '2111 Job Detail Report.pdf'
OUT = _HERE / '2111_data.json'

PROJECT_META = {
    'job_number': '2111',
    'job_name': 'Compass Northgate M3',
    'general_contractor': 'Compass General Construction I, LLC',
    'owner': 'Simon Property Group (Indianapolis, IN)',
    'location': 'Northgate Station PH2, Seattle, WA',
    'project_type': '6-story + basement multifamily (186 units)',
    'units': 186,
    'total_fixtures': 930,
    'architect': 'CPL/GGLO',
    'mep_engineer': 'Franklin Engineering',
    'start_date': '2025-10',
    'end_date': None,
    'duration_months': 13,
    'ri_foreman': 'Victor',
    'insurance': 'Standard (COI)',
    # Contract / retention values are pulled at runtime from the JDR
    # job_totals line; these are documentation defaults only.
    'contract_original': 4230000,
    'contract_final': 4512862,
    'total_cos_net': 282862,
    'billed_to_date': 2203668.00,
    'retention_amount': 110183.40,
    'retention_rate': 0.05,
}

MONEY_RE = r'[-]?[\d,]+\.\d{2}'


def parse_money(s):
    if s is None:
        return None
    s = s.strip().replace(',', '')
    neg = s.endswith('-')  # trailing minus = negative (Sage convention)
    if neg:
        s = s[:-1]
    if s.startswith('-'):
        neg = True
        s = s[1:]
    try:
        return -float(s) if neg else float(s)
    except ValueError:
        return None


def parse_date(s):
    # mm/dd/yy -> 20yy-mm-dd
    m = re.match(r'^(\d{2})/(\d{2})/(\d{2})$', s.strip())
    if not m:
        return None
    mm, dd, yy = m.groups()
    return f'20{yy}-{mm}-{dd}'


def parse_jdr(pdf_path, job_id):
    out = {
        'job_id': job_id,
        'job_name': None,
        'transactions': [],
        'code_totals': {},
        'job_totals': {},
        'by_source': {},
        'pages': 0,
    }
    with pdfplumber.open(pdf_path) as pdf:
        out['pages'] = len(pdf.pages)
        current_code = None
        current_code_desc = None
        current_party = None
        current_ref = None
        current_post_date = None
        current_org_num = None
        current_src = None
        pending_gl_date = None
        for pi, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ''
            lines = text.split('\n')
            for ln in lines:
                ln = ln.rstrip()
                if not ln.strip():
                    continue

                # Meta
                if out['job_name'] is None:
                    m = re.match(r'^\s*Job Name:\s+(.+?)\s*$', ln)
                    if m:
                        out['job_name'] = m.group(1).strip()

                # Cost code section header:  '100 Supervision'
                m = re.match(r'^(\d{3})\s+([A-Za-z][A-Za-z &/\-,\.]+?)\s*$', ln)
                if m and not re.search(MONEY_RE, ln) and 'hours' not in ln:
                    current_code = m.group(1)
                    current_code_desc = m.group(2).strip()
                    continue

                # Cost Code Totals:  'Cost Code Totals  ORIG  REV  +/-  ACTUAL  NETDUE  RETAINAGE'
                m = re.match(r'^\s*Cost Code Totals\s+(.+?)\s*$', ln)
                if m and current_code:
                    nums = re.findall(MONEY_RE, m.group(1))
                    if len(nums) >= 4:
                        vals = [parse_money(n) for n in nums]
                        while len(vals) < 6:
                            vals.append(0.0)
                        out['code_totals'][current_code] = {
                            'desc': current_code_desc,
                            'orig': vals[0], 'rev': vals[1], 'var': vals[2],
                            'actual': vals[3], 'net_due': vals[4], 'retainage': vals[5],
                        }
                    continue

                # Job Totals line
                m = re.match(r'^\s*Job Totals Revenues:\s+(.+?)\s*$', ln)
                if m:
                    nums = re.findall(MONEY_RE, m.group(1))
                    if len(nums) >= 5:
                        vals = [parse_money(n) for n in nums]
                        out['job_totals'] = {
                            'revenues': vals[0], 'expenses': vals[1],
                            'net': vals[2], 'net_due': vals[3], 'retainage': vals[4],
                        }
                    continue

                # by Source line
                m = re.match(
                    r'^\s*by Source:\s+GL:\s+([\d,.-]+)\s+AP:\s+([\d,.-]+)\s+PR:\s+([\d,.-]+)\s+AR:\s+([\d,.-]+)',
                    ln,
                )
                if m:
                    out['by_source'] = {
                        'GL': parse_money(m.group(1)), 'AP': parse_money(m.group(2)),
                        'PR': parse_money(m.group(3)), 'AR': parse_money(m.group(4)),
                    }
                    continue

                # Transaction line 1:  'PR 262 09/19/25 GE69 Gerard, Jeffrey S'
                m = re.match(r'^\s*(PR|AP|AR|GL)\s+(\d+)\s+(\d{2}/\d{2}/\d{2})\s+(\S+)\s+(.+?)\s*$', ln)
                if m:
                    current_src = m.group(1)
                    current_ref = m.group(2)
                    current_post_date = m.group(3)
                    current_org_num = m.group(4)
                    current_party = m.group(5).strip()
                    continue

                # Transaction line 2
                if current_src and current_ref:
                    # PR with hours
                    m = re.match(
                        r'^\s*(\d{2}/\d{2}/\d{2})\s+(Regular|Overtime|Double Time|Overhead \d*):?\s+([\d.]+)\s+hours\s+([\d,.\-]+)\s*(?:Ck\s*#:\s*(\d+))?',
                        ln,
                    )
                    if m and current_src == 'PR':
                        out['transactions'].append({
                            'src': 'PR', 'ref': current_ref, 'code': current_code,
                            'post_date': parse_date(current_post_date),
                            'doc_date': parse_date(m.group(1)),
                            'org_num': current_org_num, 'party_name': current_party,
                            'hour_type': m.group(2).split(':')[0],
                            'hours': float(m.group(3)),
                            'amount': parse_money(m.group(4)),
                            'ck_num': m.group(5),
                        })
                        current_ref = None
                        continue
                    # PR Overhead flat
                    m = re.match(r'^\s*(\d{2}/\d{2}/\d{2})\s+(Overhead \d*)\s+([\d,.\-]+)\s*$', ln)
                    if m and current_src == 'PR':
                        out['transactions'].append({
                            'src': 'PR', 'ref': current_ref, 'code': current_code,
                            'post_date': parse_date(current_post_date),
                            'doc_date': parse_date(m.group(1)),
                            'org_num': current_org_num, 'party_name': current_party,
                            'hour_type': m.group(2), 'hours': 0.0,
                            'amount': parse_money(m.group(3)), 'ck_num': None,
                        })
                        current_ref = None
                        continue
                    # AP line2
                    m = re.match(r'^\s*(\d{2}/\d{2}/\d{2})\s+Inv:\s+(\S+)\s+([\d,.\-]+)\s*$', ln)
                    if m and current_src == 'AP':
                        out['transactions'].append({
                            'src': 'AP', 'ref': current_ref, 'code': current_code,
                            'post_date': parse_date(current_post_date),
                            'doc_date': parse_date(m.group(1)),
                            'org_num': current_org_num, 'party_name': current_party,
                            'invoice_num': m.group(2),
                            'amount': parse_money(m.group(3)),
                        })
                        current_ref = None
                        continue
                    # AR line2
                    m = re.match(
                        r'^\s*(\d{2}/\d{2}/\d{2})\s+Invoice\s+(\S+)\s+([\d,.\-]+)(?:\s+([\d,.\-]+))?(?:\s+([\d,.\-]+))?',
                        ln,
                    )
                    if m and current_src == 'AR':
                        out['transactions'].append({
                            'src': 'AR', 'ref': current_ref, 'code': current_code,
                            'post_date': parse_date(current_post_date),
                            'doc_date': parse_date(m.group(1)),
                            'org_num': current_org_num, 'party_name': current_party,
                            'invoice_num': m.group(2),
                            'amount': parse_money(m.group(3)),
                            'net_due': parse_money(m.group(4)) if m.group(4) else None,
                            'retainage': parse_money(m.group(5)) if m.group(5) else None,
                        })
                        current_ref = None
                        continue
                    # GL line2: standalone date
                    m = re.match(r'^\s*(\d{2}/\d{2}/\d{2})\s*$', ln)
                    if m and current_src == 'GL':
                        pending_gl_date = parse_date(m.group(1))
                        continue
                    # GL line3: amount alone
                    m = re.match(r'^\s*([\d,.\-]+)\s*$', ln)
                    if m and current_src == 'GL' and pending_gl_date:
                        out['transactions'].append({
                            'src': 'GL', 'ref': current_ref, 'code': current_code,
                            'post_date': parse_date(current_post_date),
                            'doc_date': pending_gl_date,
                            'org_num': current_org_num, 'party_name': current_party,
                            'amount': parse_money(m.group(1)),
                        })
                        current_ref = None
                        pending_gl_date = None
                        continue
    return out


if __name__ == '__main__':
    data = parse_jdr(str(PDF), PROJECT_META['job_number'])
    OUT.write_text(json.dumps(data, indent=2, default=str))
    print(f'Job {data["job_id"]}: {data["pages"]} pages')
    print(f'  Job name:     {data["job_name"]}')
    print(f'  Transactions: {len(data["transactions"])}')
    print(f'  Cost codes:   {len(data["code_totals"])}')
    print(f'  Job totals:   {data["job_totals"]}')
    print(f'  by Source:    {data["by_source"]}')
    print(f'  Wrote:        {OUT}')
