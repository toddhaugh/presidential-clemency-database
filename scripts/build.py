#!/usr/bin/env python3
"""Build the Presidential Clemency Database website from the ICGE workbook.

    python scripts/build.py                 # uses the most recently added .xlsx in data/
    python scripts/build.py path/to/file.xlsx

Reads the "Clemency Data" sheet (plus two reference values from the README and
Summary sheets), computes every figure shown on the site, and writes:

    index.html, records.html, administrations.html, classification.html,
    methodology.html, about.html   - the six pages, with figures filled in
    records.json                   - the searchable records (named recipients)
    BUILD_REPORT.md                - what was built, plus checks

The page designs live in src/pages/<page>.template.html. Figures are written
into the templates where a {%name%} token appears ({%json:name%} for chart data).
Nothing here needs editing for a routine data update.
"""
import datetime as dt
import glob
import json
import math
import os
import re
import statistics
import subprocess
import sys
from collections import Counter, defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAGES = ['index', 'records', 'administrations', 'classification', 'methodology', 'about']

# ---------------------------------------------------------------------------
# Reference information that does not come from the workbook rows
# ---------------------------------------------------------------------------
# Presidential terms in display order. (data label, short key used by the site, start year)
TERMS = [('GHW Bush', 'ghwb', 1989), ('Clinton I', 'clinton1', 1993), ('Clinton II', 'clinton2', 1997),
         ('GW Bush I', 'gwb1', 2001), ('GW Bush II', 'gwb2', 2005), ('Obama I', 'obama1', 2009),
         ('Obama II', 'obama2', 2013), ('Trump I', 'trump1', 2017), ('Biden', 'biden', 2021),
         ('Trump II', 'trump2', 2025)]
TERM_NAMES = [t[0] for t in TERMS]
TERM_START = {t: dt.date(y, 1, 20) for t, _, y in TERMS}   # inauguration day
CURRENT_TERM = 'Trump II'          # the term still in progress (excluded from "completed terms")
FIRST_YEAR = 1989

GROUPS = {  # workbook "Categorical Group" -> label used on the site
    'None': '', 'Obama drug commutations': 'Obama drug commutations',
    'Jan 6 Proclamation': 'Jan. 6 proclamation', 'Biden family': 'Biden family pardons'}
STATUSES = ['Clean', 'Corrected', 'Needs review', 'Placeholder', 'Excluded']
CATEGORIES = ['White-collar crime', 'Corporate crime', 'Fraud', 'Unethical business practices']
CONFIDENCE = ['High', 'Medium', 'Low', 'Not rated']

# Short explanations shown on the Methodology page for excluded records. A newly
# excluded record that is not listed here shows the explanation from its
# Cleaning Notes instead.
EXCLUDED_REASONS = {
    'CR-2991': 'Superseded by a later grant (CR-3025)',
    'CR-3095': 'Duplicate entry',
    'CR-4707': 'Duplicate entry',
    'CR-4741': 'No pardon on DOJ record; commutation counted as CR-4858',
    'CR-4862': 'Amended re-issue of CR-4858',
    'CR-4900': 'Amended Jan. 20, 2026 warrant re-issuing a Jan. 15 pardon',
    'CR-4901': 'Amended Jan. 20, 2026 warrant re-issuing a Jan. 15 pardon',
    'CR-4902': 'Amended Jan. 20, 2026 warrant re-issuing a Jan. 15 pardon',
}
EXCLUDED_NAMES = {'CR-4902': 'Wanda Vázquez Garced'}  # accented spelling used on the site

REQUIRED_COLUMNS = ['Record ID', 'Administration', 'President', 'Clemency Type', 'Date Granted', 'Year',
                    'Term Year', 'Name', 'District', 'Sentence Year', 'Offense', 'White Collar?', 'Categories',
                    'Explanation', 'Confidence', 'Second Opinion (WC)', 'Categorical Group', 'Include in Analysis',
                    'Data Status', 'Cleaning Notes', 'Source', 'Years Sentence to Grant', 'USSC Offense Type',
                    'Offense Subtype']

# How each Sentencing Commission offense type reads inside a sentence
USSC_PHRASE = {
    'Administration of Justice': 'administration of justice offenses', 'Bribery/Corruption': 'bribery or corruption',
    'Burglary/Trespass': 'burglary or trespass', 'Commercialized Vice': 'commercialized vice',
    'Environmental': 'environmental offenses', 'Extortion/Racketeering': 'extortion or racketeering',
    'Food and Drug': 'food and drug offenses', 'Forgery/Counterfeiting/Copyright': 'forgery, counterfeiting or copyright',
    'Fraud/Theft/Embezzlement': 'fraud, theft or embezzlement', 'Individual Rights': 'individual rights and election offenses',
    'National Defense': 'national defense offenses', 'Other': 'other offenses', 'Prison Offenses': 'prison offenses',
    'Stalking/Harassing': 'stalking or harassing', 'Child Pornography': 'child pornography'}

MONTHS = ['Jan.', 'Feb.', 'March', 'April', 'May', 'June', 'July', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.']
NUMBER_WORDS = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
                'eleven', 'twelve']


class BuildError(Exception):
    pass


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
def n(x):
    """1234 -> '1,234'"""
    return f'{x:,}'


def pct(a, b):
    """Whole-number percentage, rounded half up like the site's JavaScript."""
    return int(math.floor(100 * a / b + 0.5)) if b else 0


def ap_date(d):
    return f'{MONTHS[d.month - 1]} {d.day}, {d.year}'


def ap_month_day(d):
    return f'{MONTHS[d.month - 1]} {d.day}'


def long_date(d):
    return d.strftime('%B ') + str(d.day) + d.strftime(', %Y')


def last_name_key(name):
    """Same ordering rule the Search records page uses for 'Recipient (last name)'."""
    s = re.split(r'\s*[(,]\s*(?:aka|fka|a\.k\.a\.|f\.k\.a\.)\b', name or '', flags=re.I)[0].split(' (')[0]
    parts = [w for w in s.replace(',', ' ').split() if not re.fullmatch(r'(jr\.?|sr\.?|ii|iii|iv|v|dr\.?|m\.d\.?)', w, re.I)]
    return ((parts[-1] if parts else s) + ' ' + ' '.join(parts[:-1])).lower()


# ---------------------------------------------------------------------------
# Reading the workbook
# ---------------------------------------------------------------------------
def pick_workbook():
    files = sorted(glob.glob(os.path.join(ROOT, 'data', '*.xlsx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files:
        raise BuildError('No .xlsx workbook found in the data/ folder.')
    if len(files) == 1:
        return files[0]

    def committed(f):  # newest commit time; files never committed count as newest
        try:
            out = subprocess.run(['git', 'log', '-1', '--format=%ct', '--', f], cwd=ROOT,
                                 capture_output=True, text=True).stdout.strip()
            return int(out) if out else 10 ** 12
        except Exception:
            return os.path.getmtime(f)
    return max(files, key=lambda f: (committed(f), os.path.getmtime(f)))


def clean(v):
    if v is None:
        return ''
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, str):
        v = v.strip()
        return '' if v in ('None', 'N/A') else v
    return v


def read_workbook(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sheet in ('Clemency Data', 'README', 'Summary'):
        if sheet not in wb.sheetnames:
            raise BuildError(f'The workbook has no "{sheet}" sheet.')
    rows = wb['Clemency Data'].iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else '' for h in next(rows)]
    missing = [c for c in REQUIRED_COLUMNS if c not in header]
    if missing:
        raise BuildError('Clemency Data is missing column(s): ' + ', '.join(missing))
    col = {h: i for i, h in enumerate(header)}
    recs, problems = [], []
    for rownum, r in enumerate(rows, start=2):
        if not any(c is not None and str(c).strip() for c in r):
            continue
        g = lambda k: clean(r[col[k]])  # noqa: E731
        rec = {k: g(k) for k in REQUIRED_COLUMNS}
        rec['_row'] = rownum
        rec['Categorical Group'] = rec['Categorical Group'] or 'None'
        # Year, Term Year and Years Sentence to Grant are formulas in the workbook. Recompute them here
        # (same formulas) so the build does not depend on Excel having saved calculated values.
        d = rec['Date Granted']
        if isinstance(d, dt.date) and rec['Administration'] in TERM_START:
            rec['Year'] = d.year
            rec['Term Year'] = min(4, int((d - TERM_START[rec['Administration']]).days / 365.25) + 1)
            sy = rec['Sentence Year']
            rec['Years Sentence to Grant'] = d.year - sy if isinstance(sy, (int, float)) and sy > 1900 else ''
        rec['Categories'] = rec['Categories'] or ''
        recs.append(rec)
        where = f'Clemency Data row {rownum} ({rec["Record ID"] or "no ID"})'
        if not re.fullmatch(r'CR-\d{4,}', str(rec['Record ID'])):
            problems.append(f'{where}: Record ID should look like CR-0001')
        if not isinstance(rec['Date Granted'], dt.date):
            problems.append(f'{where}: Date Granted is not a date')
        if rec['Administration'] not in TERM_NAMES:
            problems.append(f'{where}: unknown Administration "{rec["Administration"]}"')
        if rec['Clemency Type'] not in ('Pardon', 'Commutation'):
            problems.append(f'{where}: Clemency Type should be Pardon or Commutation')
        if rec['Include in Analysis'] not in ('Yes', 'No'):
            problems.append(f'{where}: Include in Analysis should be Yes or No')
        if rec['Data Status'] not in STATUSES:
            problems.append(f'{where}: unknown Data Status "{rec["Data Status"]}"')
        if rec['Categorical Group'] not in GROUPS:
            problems.append(f'{where}: unknown Categorical Group "{rec["Categorical Group"]}"')
        if rec['White Collar?'] not in ('Yes', 'No'):
            problems.append(f'{where}: White Collar? should be Yes or No')
        if rec['Confidence'] not in CONFIDENCE:
            problems.append(f'{where}: unknown Confidence "{rec["Confidence"]}"')
        if rec['Term Year'] not in (1, 2, 3, 4):
            problems.append(f'{where}: Term Year should be 1-4')
        for c in [c.strip() for c in rec['Categories'].split(';') if c.strip()]:
            if c not in CATEGORIES:
                problems.append(f'{where}: unknown category "{c}"')
        if rec['Data Status'] != 'Placeholder' and rec['Include in Analysis'] == 'Yes' and not rec['Name']:
            problems.append(f'{where}: named record has no Name')
    ids = Counter(r['Record ID'] for r in recs)
    problems += [f'Record ID {i} appears {c} times' for i, c in ids.items() if c > 1]
    if any('""' in p or 'should be Yes or No' in p for p in problems):
        problems.insert(0, 'Some cells that Excel calculates are blank. If the workbook was saved by a program other '
                           'than Excel, open it in Excel, save it, and upload it again.')
    if problems:
        raise BuildError('The workbook has problems to fix before the site can be built:\n  - '
                         + '\n  - '.join(problems[:60]) + ('\n  ...' if len(problems) > 60 else ''))

    # README: "Last checked against DOJ"
    checked = None
    for r in wb['README'].iter_rows(values_only=True):
        if r and r[0] and str(r[0]).strip().lower() == 'last checked against doj':
            v = r[1]
            if isinstance(v, dt.datetime):
                checked = v.date()
            elif isinstance(v, str):
                for fmt in ('%Y-%m-%d', '%B %d, %Y', '%b %d, %Y', '%m/%d/%Y'):
                    try:
                        checked = dt.datetime.strptime(v.strip(), fmt).date()
                        break
                    except ValueError:
                        pass
    if not checked:
        raise BuildError('README sheet: could not read the date next to "Last checked against DOJ".')

    # Summary section 8: Biden group-warrant commutations (counts only)
    biden = None
    srows = [r for r in wb['Summary'].iter_rows(values_only=True)]
    for i, r in enumerate(srows):
        if r and isinstance(r[0], str) and r[0].strip().startswith('8.') and 'Biden' in r[0]:
            batches = []
            for rr in srows[i + 1:i + 8]:
                if rr and len(rr) > 3 and isinstance(rr[1], dt.datetime) and isinstance(rr[2], (int, float)):
                    batches.append({'date': rr[1].date(), 'n': int(rr[2]), 'wc': int(rr[3])})
            if len(batches) >= 1:
                biden = batches
            break
    if not biden:
        raise BuildError('Summary sheet: could not read section 8 (Biden mass commutations).')

    # Website Map: the workbook's own live values, used only to cross-check this script
    wmap = {}
    if 'Website Map' in wb.sheetnames:
        section = None
        for r in wb['Website Map'].iter_rows(values_only=True):
            c0 = r[0] if r else None
            if isinstance(c0, str) and re.match(r'^W\d+ · ', c0):
                section = c0.split(' · ')[0]
            elif section and isinstance(c0, str) and len(r) > 1 and r[1] is not None and c0 != 'Number':
                v = r[1]
                if isinstance(v, dt.datetime):
                    v = v.date()
                wmap[(section, c0.strip())] = v
    return recs, checked, biden, wmap


# ---------------------------------------------------------------------------
# Figures
# ---------------------------------------------------------------------------
def compute(recs, checked, biden):
    F, checks = {}, {}
    inc = [r for r in recs if r['Include in Analysis'] == 'Yes']
    ind = [r for r in inc if r['Categorical Group'] == 'None']
    named = [r for r in inc if r['Data Status'] != 'Placeholder']
    wc = lambda rs: [r for r in rs if r['White Collar?'] == 'Yes']  # noqa: E731
    cats = lambda r: [c.strip() for c in r['Categories'].split(';') if c.strip()]  # noqa: E731
    ind_wc = wc(ind)
    latest = max(r['Date Granted'] for r in inc)
    latest_year = latest.year

    # --- dates and headline counts -------------------------------------------------
    F.update({
        'data_through': ap_date(latest), 'data_through_long': long_date(latest),
        'data_through_md': ap_month_day(latest), 'latest_year': str(latest_year),
        'doj_checked': ap_date(checked), 'doj_checked_long': long_date(checked),
        'doj_utc_args': f'{checked.year}, {checked.month - 1}, {checked.day}',
        'json:latest_year': latest_year,
        'json:partial_note': f' (through {ap_month_day(latest)})',
        'presidents_word': NUMBER_WORDS[len({r['President'] for r in inc})],
        'pardons_all': n(sum(r['Clemency Type'] == 'Pardon' for r in inc)),
        'commutations_all': n(sum(r['Clemency Type'] == 'Commutation' for r in inc)),
        'individual': n(len(ind)), 'mass': n(len(inc) - len(ind)),
        'named': n(len(named)),
        'placeholders': n(sum(r['Data Status'] == 'Placeholder' for r in recs)),
        'records_total': n(len(recs)),
    })
    checks['W1'] = {'Pardons (all recorded)': sum(r['Clemency Type'] == 'Pardon' for r in inc),
                    'Commutations (all recorded)': sum(r['Clemency Type'] == 'Commutation' for r in inc),
                    'Individual grants': len(ind), 'Grants in mass actions': len(inc) - len(ind),
                    'Total grants counted': len(inc), 'Latest grant in data': latest}

    # --- key facts -----------------------------------------------------------------
    completed = [r for r in ind if r['Administration'] != CURRENT_TERM]
    final = [r for r in completed if r['Term Year'] == 4]
    yrs = [r['Years Sentence to Grant'] for r in ind if isinstance(r['Years Sentence to Grant'], (int, float))]
    med = statistics.median(yrs)
    fraud = [r for r in ind_wc if 'Fraud' in cats(r)]
    F.update({
        'wc_ind': n(len(ind_wc)), 'wc_ind_pct': str(pct(len(ind_wc), len(ind))),
        'final_n': n(len(final)), 'completed_n': n(len(completed)), 'final_pct': str(pct(len(final), len(completed))),
        'median_years': str(int(med)) if float(med).is_integer() else f'{med:.1f}',
        'fraud_wc': n(len(fraud)), 'fraud_wc_pct': str(pct(len(fraud), len(ind_wc))),
    })
    checks['W2'] = {'White-collar individual grants': len(ind_wc),
                    'White-collar share of individual grants': len(ind_wc) / len(ind),
                    'Individual grants in final year, completed terms': len(final),
                    'Individual grants, completed terms': len(completed),
                    'Final-year share': len(final) / len(completed), 'Median years from sentencing to grant': med,
                    'White-collar grants tagged Fraud': len(fraud), 'Fraud share of white-collar grants': len(fraud) / len(ind_wc)}

    # --- by term -------------------------------------------------------------------
    by_term = {t: {'ind': 0, 'ind_wc': 0, 'all': 0, 'all_wc': 0, 'P': [0] * 4, 'C': [0] * 4, 'WP': [0] * 4, 'WC': [0] * 4,
                   'fraud': 0, 'corp': 0, 'uneth': 0} for t in TERM_NAMES}
    for r in inc:
        t, y = by_term[r['Administration']], r['Term Year'] - 1
        t['all'] += 1
        t['all_wc'] += r['White Collar?'] == 'Yes'
        p = r['Clemency Type'] == 'Pardon'
        t['P' if p else 'C'][y] += 1
        if r['White Collar?'] == 'Yes':
            t['WP' if p else 'WC'][y] += 1
    for r in ind:
        t = by_term[r['Administration']]
        t['ind'] += 1
        if r['White Collar?'] == 'Yes':
            t['ind_wc'] += 1
            c = cats(r)
            t['fraud'] += 'Fraud' in c
            t['corp'] += 'Corporate crime' in c
            t['uneth'] += 'Unethical business practices' in c
    F['json:ind_rows'] = [[t, by_term[t]['ind'], by_term[t]['ind_wc']] for t in TERM_NAMES]
    F['json:all_rows'] = [[t, by_term[t]['all'], by_term[t]['all_wc']] for t in TERM_NAMES]
    F['json:ind_by_term'] = {t: [by_term[t]['ind'], by_term[t]['ind_wc']] for t in TERM_NAMES}
    F['json:ind_max'] = max(by_term[t]['ind'] for t in TERM_NAMES)
    F['json:term_years'] = {k: {'P': by_term[t]['P'], 'C': by_term[t]['C'], 'WP': by_term[t]['WP'], 'WC': by_term[t]['WC'],
                                'ind': [by_term[t]['ind'], by_term[t]['ind_wc']]} for t, k, _ in TERMS}
    F['json:wc_cats_by_term'] = [[by_term[t]['ind_wc'], by_term[t]['fraud'], by_term[t]['corp'], by_term[t]['uneth']]
                                 for t in TERM_NAMES]
    checks['W3'], checks['W6'], checks['W7'], checks['W9'] = {}, {}, {}, {}
    for t in TERM_NAMES:
        b = by_term[t]
        checks['W3'].update({f'{t} – individual grants': b['ind'], f'{t} – individual white collar': b['ind_wc'],
                             f'{t} – individual share': b['ind_wc'] / b['ind'] if b['ind'] else 0,
                             f'{t} – all recorded grants': b['all'], f'{t} – all recorded white collar': b['all_wc']})
        checks['W6'].update({f'{t} – pardons': sum(b['P']), f'{t} – commutations': sum(b['C'])})
        for y in range(4):
            checks['W7'].update({f'{t} – year {y + 1} pardons': b['P'][y], f'{t} – year {y + 1} commutations': b['C'][y],
                                 f'{t} – year {y + 1} white collar': b['WP'][y] + b['WC'][y]})
        checks['W9'].update({f'{t} – fraud': b['fraud'], f'{t} – corporate crime': b['corp'],
                             f'{t} – unethical business': b['uneth']})

    # --- grants by calendar year (individual) ---------------------------------------
    term_start = {t: s for t, _, s in TERMS}
    term_end = {t: (s + 4 if t != CURRENT_TERM else latest_year) for t, _, s in TERMS}
    years = list(range(FIRST_YEAR, latest_year + 1))
    Y, BY = [], {}
    cnt = Counter((r['Date Granted'].year, r['Administration']) for r in ind)
    for y in years:
        tot = sum(1 for r in ind if r['Date Granted'].year == y)
        w = sum(1 for r in ind_wc if r['Date Granted'].year == y)
        Y.append([y, tot, w])
        BY[y] = {t: cnt[(y, t)] for t in TERM_NAMES if term_start[t] <= y <= term_end[t]}
    # each calendar year is labelled with the term that made most of its grants (ties: outgoing term)
    band = Counter()
    for y in years:
        terms_y = list(BY[y].keys())
        best = max(terms_y, key=lambda t: (BY[y][t], -TERM_NAMES.index(t)))
        band[best] += 1
    F['json:years'] = Y
    F['json:by_year_term'] = {str(k): v for k, v in BY.items()}
    F['json:bands'] = [[t, band[t]] for t in TERM_NAMES if band[t]]
    ymax = max(50, int(math.ceil(max(t for _, t, _ in Y) / 50.0) * 50))
    F['json:year_axis_max'] = ymax
    for i, f in enumerate([1, 0.75, 0.5, 0.25]):
        F[f'axis{i}'] = n(int(ymax * f))
    checks['W4'] = {}
    for y, t, w in Y:
        checks['W4'].update({f'{y} – individual grants': t, f'{y} – white collar': w})

    # --- most recent grants -----------------------------------------------------------
    recent = sorted(named, key=lambda r: r['Record ID'])
    recent.sort(key=lambda r: r['Date Granted'], reverse=True)
    short = lambda s: s if len(s) <= 130 else s[:127].strip() + '…'  # noqa: E731  (same rule as Search records)
    F['json:recent'] = [[ap_date(r['Date Granted']), r['Name'], r['Clemency Type'], short(str(r['Offense'])), r['White Collar?']]
                        for r in recent[:6]]
    checks['W5'] = {'Named recipients listed on Search records': len(named),
                    'Records marked Needs review': sum(r['Data Status'] == 'Needs review' for r in inc)}

    # --- mass actions ----------------------------------------------------------------------
    grp = lambda g, typ=None: sum(1 for r in inc if r['Categorical Group'] == g and (typ is None or r['Clemency Type'] == typ))  # noqa: E731
    F.update({
        'j6_pardons': n(grp('Jan 6 Proclamation', 'Pardon')), 'j6_commutations': n(grp('Jan 6 Proclamation', 'Commutation')),
        'j6_total': n(grp('Jan 6 Proclamation')),
        'obama_drug_pardons': n(grp('Obama drug commutations', 'Pardon')),
        'obama_drug': n(grp('Obama drug commutations', 'Commutation')), 'obama_drug_total': n(grp('Obama drug commutations')),
        'biden_family': n(grp('Biden family', 'Pardon')), 'biden_family_commutations': n(grp('Biden family', 'Commutation')),
        'biden_family_total': n(grp('Biden family')),
    })
    bt = sum(b['n'] for b in biden)
    bw = sum(b['wc'] for b in biden)
    F.update({'biden_warrants': n(bt), 'biden_warrants_wc': n(bw), 'biden_warrants_wc_pct': str(pct(bw, bt)),
              'biden_warrant_dates': ' and '.join(ap_date(b['date']) for b in biden),
              'biden_warrant_detail': ' and '.join(f'{ap_date(b["date"])} ({n(b["n"])} recipients)' for b in biden)})
    checks['W12'] = {'Jan 6 Proclamation – grants': grp('Jan 6 Proclamation'),
                     'Obama drug commutations – grants': grp('Obama drug commutations'),
                     'Biden family – grants': grp('Biden family')}

    # --- record status, review, exclusions --------------------------------------------------
    st = Counter(r['Data Status'] for r in recs)
    for s in STATUSES:
        F['status_' + s.lower().replace(' ', '_')] = n(st[s])
        checks['W12'][f'Status: {s}'] = st[s]
    checks['W12']['Total records'] = len(recs)
    F['needs_review'] = n(st['Needs review'])
    F['excluded'] = n(st['Excluded'])
    F['second_opinions'] = n(sum(1 for r in recs if r['Second Opinion (WC)']))
    exc = []
    for r in sorted([r for r in recs if r['Data Status'] == 'Excluded' or r['Include in Analysis'] == 'No'],
                    key=lambda r: r['Record ID']):
        reason = EXCLUDED_REASONS.get(r['Record ID'])
        if not reason:
            note = r['Cleaning Notes'] or 'Excluded'
            m = re.search(r'Excluded\s*[–-]\s*(.*)$', note)
            reason = (m.group(1) if m else note).rstrip('.')
        exc.append((r['Record ID'], EXCLUDED_NAMES.get(r['Record ID'], r['Name']), reason))
    F['excluded_rows'] = '\n'.join(
        f'<sc-raw-tr style="{"border-bottom: 1px solid #E4DFDE; " if k < len(exc) - 1 else ""}vertical-align: top">'
        f'<sc-raw-td style="padding: 8px 12px 8px 0">{esc(i)}</sc-raw-td><sc-raw-td style="padding: 8px 12px">{esc(nm)}'
        f'</sc-raw-td><sc-raw-td style="padding: 8px 0">{esc(why)}</sc-raw-td></sc-raw-tr>' for k, (i, nm, why) in enumerate(exc))

    # --- classification page ------------------------------------------------------------------
    ind_p = [r for r in ind if r['Clemency Type'] == 'Pardon']
    ind_c = [r for r in ind if r['Clemency Type'] == 'Commutation']
    nwc = len(ind) - len(ind_wc)
    F.update({
        'nwc_ind': n(nwc), 'nwc_ind_pct': str(pct(nwc, len(ind))),
        'json:wc_ind_n': len(ind_wc), 'json:nwc_ind_n': nwc, 'json:ind_n': len(ind),
        'json:wc_share_1dp': round(100 * len(ind_wc) / len(ind), 1),
        'json:nwc_share_1dp': round(100 - round(100 * len(ind_wc) / len(ind), 1), 1),
        'ind_pardons': n(len(ind_p)), 'ind_pardons_wc': n(len(wc(ind_p))), 'ind_pardons_pct': str(pct(len(wc(ind_p)), len(ind_p))),
        'ind_comm': n(len(ind_c)), 'ind_comm_wc': n(len(wc(ind_c))), 'ind_comm_pct': str(pct(len(wc(ind_c)), len(ind_c))),
    })
    tagc = Counter(c for r in ind_wc for c in cats(r))
    F.update({'tag_general': n(tagc['White-collar crime']), 'tag_fraud': n(tagc['Fraud']),
              'tag_fraud_pct': str(pct(tagc['Fraud'], len(ind_wc))), 'tag_corp': n(tagc['Corporate crime']),
              'tag_corp_pct': str(pct(tagc['Corporate crime'], len(ind_wc))),
              'tag_uneth': n(tagc['Unethical business practices']),
              'tag_uneth_pct': str(pct(tagc['Unethical business practices'], len(ind_wc)))})
    for k, c in [('fraud', 'Fraud'), ('corp', 'Corporate crime'), ('uneth', 'Unethical business practices')]:
        F[f'tag_{k}_w'] = f'{100 * tagc[c] / len(ind_wc):.1f}'
    combos = Counter(frozenset(c for c in cats(r) if c != 'White-collar crime') for r in ind_wc)
    F.update({'combo_fraud': n(combos[frozenset({'Fraud'})]), 'combo_general': n(combos[frozenset()]),
              'combo_corp_fraud': n(combos[frozenset({'Corporate crime', 'Fraud'})]),
              'combo_corp': n(combos[frozenset({'Corporate crime'})]),
              'combo_cfu': n(combos[frozenset({'Corporate crime', 'Fraud', 'Unethical business practices'})])})
    shown = [frozenset({'Fraud'}), frozenset(), frozenset({'Corporate crime', 'Fraud'}), frozenset({'Corporate crime'}),
             frozenset({'Corporate crime', 'Fraud', 'Unethical business practices'})]
    other = {k: v for k, v in combos.items() if k not in shown}
    F['combo_other_uneth'] = n(sum(other.values()))
    checks['W8'] = {'Individual pardons': len(ind_p), 'Individual pardons – white collar': len(wc(ind_p)),
                    'Individual commutations': len(ind_c), 'Individual commutations – white collar': len(wc(ind_c)),
                    'Individual grants – not white collar': nwc}
    checks['W9'].update({'Tagged Fraud': tagc['Fraud'], 'Tagged Corporate crime': tagc['Corporate crime'],
                         'Tagged Unethical business practices': tagc['Unethical business practices'],
                         'Tagged General white-collar crime tag': tagc['White-collar crime']})

    ut = Counter(r['USSC Offense Type'] for r in ind)
    uw = Counter(r['USSC Offense Type'] for r in ind_wc)
    U = sorted(ut, key=lambda k: (-ut[k], k))
    F['json:ussc'] = [[k, ut[k], uw[k]] for k in U]
    F['json:ussc_types'] = sorted(set(r['USSC Offense Type'] for r in named if r['USSC Offense Type'] and r['USSC Offense Type'] != 'Not recorded'))
    sub = Counter((r['USSC Offense Type'], r['Offense Subtype']) for r in ind)
    subs = []
    for typ, label in [('Drug Trafficking', 'Drug trafficking'), ('Drug Possession', 'Drug possession'),
                       ('Firearms', 'Firearms'), ('Robbery', 'Robbery')]:
        subs.append([label, ut[typ], 1])
        for (t, s), c in sorted(sub.items(), key=lambda kv: (-kv[1], kv[0][1])):
            if t == typ and s:
                subs.append([s, c, 0])
    F['json:subtypes'] = subs
    lv = lambda typ, s=None: (ut[typ] if s is None else sub[(typ, s)])  # noqa: E731
    F['json:levels_n'] = [lv('Drug Trafficking', 'Cocaine or crack'), lv('Drug Trafficking', 'Methamphetamine'),
                          lv('Drug Trafficking', 'Marijuana'), lv('Money Laundering'), lv('Robbery', 'Bank robbery'),
                          lv('Firearms', 'Possession by prohibited person'), lv('Assault'), lv('Kidnapping'),
                          lv('Robbery', 'Carjacking'), lv('Child Pornography'), lv('Burglary/Trespass', 'Trespass'),
                          lv('Immigration', 'Illegal reentry')]
    nw = lambda k: ut[k] - uw[k]  # noqa: E731
    drugs_nwc = nw('Drug Trafficking') + nw('Drug Possession')
    F.update({'fte_wc': n(uw['Fraud/Theft/Embezzlement']), 'drugs_nwc': n(drugs_nwc),
              'drugs_nwc_pct': str(pct(drugs_nwc, nwc)), 'rights_nwc': n(nw('Individual Rights')),
              'firearms_nwc': n(nw('Firearms'))})
    checks['W10'] = {}
    for k in U:
        checks['W10'].update({f'{k} – individual grants': ut[k], f'{k} – white collar': uw[k]})
    for (t, s), c in sub.items():
        if s:
            checks['W10'][f'{t} › {s}'] = c

    conf = Counter((r['White Collar?'], r['Confidence']) for r in ind)
    checks['W11'] = {f'{"White collar" if w == "Yes" else "Not white collar"} – {c}': conf[(w, c)]
                     for w in ('Yes', 'No') for c in ('High', 'Medium', 'Low')}

    # --- statements in the page text that depend on the data ------------------------------------
    nwc_groups = {'drugs': drugs_nwc}
    nwc_groups.update({k: nw(k) for k in ut if k not in ('Drug Trafficking', 'Drug Possession')})
    nwc_rank = sorted(nwc_groups, key=lambda k: (-nwc_groups[k], k))
    others = [k for k in nwc_rank if k != 'drugs']
    F.update({'nwc_second': USSC_PHRASE.get(others[0], others[0].lower()), 'nwc_second_n': n(nwc_groups[others[0]]),
              'nwc_third': USSC_PHRASE.get(others[1], others[1].lower()), 'nwc_third_n': n(nwc_groups[others[1]])})
    claims = [
        ('Classification page: drug trafficking and fraud/theft/embezzlement "together account for more than half" of individual grants',
         ut['Drug Trafficking'] + ut['Fraud/Theft/Embezzlement'] > len(ind) / 2),
        ('Classification page: drug offenses are the largest group of non-white-collar grants', nwc_rank[0] == 'drugs'),
        ('Classification page: fraud/theft/embezzlement is "the most common" white-collar offense type',
         max(uw, key=lambda k: uw[k]) == 'Fraud/Theft/Embezzlement'),
        ('Classification page: the only category combinations not listed in Table 5 involve unethical business practices',
         all('Unethical business practices' in k for k in other)),
    ]
    for key, label in [('Clinton II', 'Clinton second term'), ('Trump I', 'Trump first term')]:
        b = by_term[key]
        claims.append((f'By administration page: "Most grants came in the final year" ({label})',
                       (b['P'][3] + b['C'][3]) > (sum(b['P']) + sum(b['C'])) / 2))
    return F, checks, claims, named


def esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# ---------------------------------------------------------------------------
# records.json for the Search records page
# ---------------------------------------------------------------------------
def source_url(src):
    m = re.search(r'https?://[^\s,)"]+', src or '')
    if m:
        return m.group(0).rstrip('.')
    m = re.search(r'\b((?:www\.)?govinfo\.gov/[^\s,)"]+)', src or '')
    return 'https://' + m.group(1).rstrip('.') if m else ''


def records_json(named):
    rows = []
    for r in named:
        rows.append([r['Record ID'], r['Date Granted'].isoformat(), r['Administration'], r['Clemency Type'], r['Name'],
                     str(r['District']), str(r['Sentence Year']), str(r['Offense']), r['White Collar?'], r['Categories'],
                     str(r['Explanation']), r['Confidence'], GROUPS[r['Categorical Group']], r['Data Status'],
                     source_url(r['Source']), str(r['Years Sentence to Grant']), r['USSC Offense Type'],
                     r['Offense Subtype']])
    rows.sort(key=lambda a: a[0])
    rows.sort(key=lambda a: a[1], reverse=True)  # newest first; the page relies on this order
    return {'rows': rows}


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
TOKEN = re.compile(r'\{%(json:)?([a-z0-9_]+)%\}')


def render(template, F, page):
    missing = []

    def sub(m):
        key = ('json:' if m.group(1) else '') + m.group(2)
        if key not in F:
            missing.append(key)
            return m.group(0)
        v = F[key]
        if m.group(1):
            return json.dumps(v, ensure_ascii=False)
        return str(v)
    out = TOKEN.sub(sub, template)
    if missing:
        raise BuildError(f'{page}: template uses unknown figure(s): ' + ', '.join(sorted(set(missing))))
    return out


def write_pages(F):
    for p in PAGES:
        src = os.path.join(ROOT, 'src', 'pages', p)
        template = open(src + '.template.html', encoding='utf-8').read()
        shell = open(src + '.shell.html', encoding='utf-8').read()
        html = render(template, F, p)
        enc = json.dumps(html, ensure_ascii=False).replace('</', '<\\u002F')
        if shell.count('@@TEMPLATE@@') != 1:
            raise BuildError(f'src/pages/{p}.shell.html is damaged (template marker missing).')
        out = shell.replace('@@TEMPLATE@@', enc)
        with open(os.path.join(ROOT, p + '.html'), 'w', encoding='utf-8') as fh:
            fh.write(out)


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------
def compare(checks, wmap):
    rows, bad, unmatched = [], [], 0
    for (sec, label), live in sorted(wmap.items()):
        mine = checks.get(sec, {}).get(label)
        if mine is None:
            unmatched += 1
            continue
        if isinstance(live, (int, float)) and isinstance(mine, (int, float)):
            ok = abs(float(live) - float(mine)) < 1e-6
        else:
            ok = str(live) == str(mine)
        rows.append((sec, label, live, mine, ok))
        if not ok:
            bad.append((sec, label, live, mine))
    return rows, bad, unmatched


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else pick_workbook()
    print(f'Workbook: {os.path.relpath(path, ROOT)}')
    recs, checked, biden, wmap = read_workbook(path)
    F, checks, claims, named = compute(recs, checked, biden)
    data = records_json(named)
    with open(os.path.join(ROOT, 'records.json'), 'w', encoding='utf-8') as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(',', ':'))
    write_pages(F)
    rows, bad, unmatched = compare(checks, wmap)

    lines = ['# Build report', '',
             f'- Workbook: `{os.path.relpath(path, ROOT)}`',
             f'- Built: {dt.datetime.now(dt.timezone.utc):%Y-%m-%d %H:%M} UTC',
             f'- Records in workbook: {n(len(recs))}; named recipients on Search records: {n(len(named))}',
             f'- Data through: {F["data_through"]}; last checked against DOJ: {F["doj_checked"]}', '']
    lines += ['## Cross-check against the workbook\'s Website Map sheet', '']
    if not wmap:
        lines.append('The workbook has no Website Map sheet, so this check was skipped.')
    elif bad:
        lines.append(f'**{len(bad)} figure(s) differ from the workbook\'s own formulas.** Check these before relying on the site:')
        lines += [f'- {s} · {lbl}: workbook {live}, site {mine}' for s, lbl, live, mine in bad]
    else:
        lines.append(f'All {len(rows)} figures the Website Map lists match the workbook\'s own formulas.')
    lines += ['', '## Statements in the page text', '',
              'These sentences are written into the pages. If one is marked FALSE, the wording needs updating in `src/pages`.', '']
    lines += [f'- {"OK" if ok else "**FALSE**"} – {c}' for c, ok in claims]
    lines += ['', '## Text that is not computed', '',
              'Update these by hand in `src/pages/*.template.html` when they change: the Methodology change log and '
              '"Last revised" date, the count of 17 records too vague to classify, the "about 1,500" unnamed Jan. 6 '
              'recipients, the 77 Nov. 7, 2025 proclamation pardons, the cleaning history, and the news item on the About page.']
    with open(os.path.join(ROOT, 'BUILD_REPORT.md'), 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines) + '\n')
    print('\n'.join(lines))
    if bad:
        print(f'\nWARNING: {len(bad)} figure(s) differ from the Website Map.', file=sys.stderr)


if __name__ == '__main__':
    try:
        main()
    except BuildError as e:
        print('BUILD STOPPED\n' + str(e), file=sys.stderr)
        sys.exit(1)
