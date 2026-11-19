#!/usr/bin/env python3
"""Turn source_421a.py output (universe.csv + run-log.json) into the target-list workbook.

Usage: python3 scripts/source_report.py DIR [--name "Williamsburg 421-a targets.xlsx"]
"""
import argparse, csv, datetime as dt, json, os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

F = 'Arial'
HEAD = PatternFill('solid', start_color='1F3A5F')
BAND = PatternFill('solid', start_color='F3F6FA')
FILL_IN = PatternFill('solid', start_color='FFFF00')
THIN = Side(style='thin', color='D0D5DD')

COLS = [  # header, key, width, number format
    ('Address', 'address', 24, None), ('Zip', 'zipcode', 8, None), ('BBL', 'bbl', 13, None),
    ('Structure', 'structure', 20, None), ('Building SF', 'bldgarea', 12, '#,##0'), ('Res. units', 'unitsres', 9, '#,##0'),
    ('Year built', 'yearbuilt', 8, '0'), ('421-a type', 'exm_desc', 20, None), ('Benefit start (FY)', 'benefit_start_fy', 10, '0'),
    ('Term (yrs)', 'term_years', 7, '0'), ('Last exempt day', 'expiry_date', 12, 'mmm d, yyyy'),
    ('Years to expiry', None, 9, '0.0'), ('Exempt assessed value ($)', 'exempt_av', 15, '$#,##0'),
    ('Est. remaining tax step-up ($/yr)', None, 15, '$#,##0'), ('Step-up per unit ($/yr)', None, 12, '$#,##0'),
    ('Owner of record', 'owner', 30, None), ('HPD registered owner', 'hpd_owner_entity', 30, None),
    ('Principal (HPD head officer)', 'hpd_principal', 22, None), ('Owner business address', 'hpd_owner_address', 36, None),
    ('Managing agent', 'hpd_agent', 30, None), ('Agent address', 'hpd_agent_address', 36, None),
    ('Phone (fill in)', None, 16, None), ('Last deed', 'last_deed_date', 11, 'mmm yyyy'), ('Deed amount ($)', 'last_deed_amt', 14, '$#,##0'),
    ('Buyer on last deed', 'last_deed_to', 28, None), ('Last mortgage', 'last_mtge_date', 11, 'mmm yyyy'),
    ('Mortgage amount ($)', 'last_mtge_amt', 14, '$#,##0'), ('Lender', 'last_mtge_to', 28, None),
    ('50 to 150 units', None, 9, None), ('Why it is here / notes', 'note', 48, None)]


def f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def d(x):
    try:
        return dt.date.fromisoformat(x[:10])
    except (TypeError, ValueError):
        return None


def sheet(wb, title, rows, intro):
    ws = wb.create_sheet(title)
    ws['A1'] = title
    ws['A1'].font = Font(name=F, bold=True, size=14)
    ws['A2'] = intro
    ws['A2'].font = Font(name=F, italic=True, size=9, color='555555')
    hr = 4
    for j, (h, _, w, _) in enumerate(COLS, 1):
        c = ws.cell(hr, j, h)
        c.font, c.fill = Font(name=F, bold=True, color='FFFFFF', size=9), HEAD
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[hr].height = 42
    idx = {h: j for j, (h, *_) in enumerate(COLS, 1)}
    for i, r in enumerate(rows, hr + 1):
        for j, (h, k, _, fmt) in enumerate(COLS, 1):
            c = ws.cell(i, j)
            if k:
                v = r.get(k, '')
                if k in ('expiry_date', 'last_deed_date', 'last_mtge_date'):
                    v = d(v)
                elif fmt and k not in ('zipcode', 'bbl'):
                    v = f(v)
                    if k in ('last_deed_amt', 'last_mtge_amt') and not v:
                        v = None
                c.value = v
            elif h == 'Years to expiry':
                c.value = f"=IF({L(idx['Last exempt day'])}{i}=\"\",\"\",({L(idx['Last exempt day'])}{i}-Method!$B$4)/365.25)"
            elif h.startswith('Est. remaining'):
                c.value = f"=IF({L(idx['Exempt assessed value ($)'])}{i}=\"\",\"\",{L(idx['Exempt assessed value ($)'])}{i}*Method!$B$5)"
            elif h.startswith('Step-up per unit'):
                a, u = L(idx['Est. remaining tax step-up ($/yr)']), L(idx['Res. units'])
                c.value = f'=IF(OR({a}{i}="",{u}{i}=0),"",{a}{i}/{u}{i})'
            elif h == '50 to 150 units':
                u = L(idx['Res. units'])
                c.value = f'=IF(AND({u}{i}>=Method!$B$8,{u}{i}<=Method!$B$9),"Yes","No")'
            elif h.startswith('Phone'):
                c.fill = FILL_IN
            if fmt:
                c.number_format = fmt
            c.font = Font(name=F, size=9, color='008000' if (not k and not h.startswith('Phone')) else '000000')
            c.alignment = Alignment(vertical='top', wrap_text=h in ('Why it is here / notes',))
            c.border = Border(bottom=THIN)
            if (i - hr) % 2 == 0 and not h.startswith('Phone'):
                c.fill = BAND
    n = hr + len(rows)
    if rows:
        ws.cell(n + 1, 1, 'Total').font = Font(name=F, bold=True, size=9)
        for h in ('Building SF', 'Res. units', 'Exempt assessed value ($)', 'Est. remaining tax step-up ($/yr)'):
            c = ws.cell(n + 1, idx[h], f'=SUM({L(idx[h])}{hr + 1}:{L(idx[h])}{n})')
            c.font, c.number_format = Font(name=F, bold=True, size=9), COLS[idx[h] - 1][3]
        ws.auto_filter.ref = f'A{hr}:{L(len(COLS))}{n}'
    ws.freeze_panes = ws.cell(hr + 1, 2)
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('dir')
    ap.add_argument('--name', default='Williamsburg 421-a targets.xlsx')
    a = ap.parse_args()
    rows = list(csv.DictReader(open(os.path.join(a.dir, 'universe.csv'))))
    log = json.load(open(os.path.join(a.dir, 'run-log.json')))
    as_of = dt.date.fromisoformat(log['as_of'])
    p = log['params']
    for r in rows:
        r['owner'] = r['condo_top_owner'] if r['structure'] == 'rental in condo structure' and r['condo_top_owner'] else r['ownername']
        notes = []
        if r['structure'] == 'rental in condo structure':
            notes.append(f"Held as condo units; one owner holds {float(r['condo_top_owner_share']):.0%} of residential units.")
        dd = d(r['last_deed_date'])
        if dd and f(r['last_deed_amt']) and (as_of - dd).days < 5 * 365:
            notes.append(f"Traded {dd:%b %Y} for ${f(r['last_deed_amt']) / 1e6:.1f}M.")
        elif dd and (as_of - dd).days > 10 * 365:
            notes.append(f'Same owner since at least {dd.year}.')
        if 'end years' in r['match_basis']:
            notes.append('Unit lots carry different end years; check the bills.')
        if r['status'] == 'expired':
            r['exempt_av'] = ''
            notes.append('Benefit already ended; taxes have reset.')
        r['note'] = ' '.join(notes)
    rental = [r for r in rows if r['rental'] == 'True']
    big = lambda r: f(r['unitsres']) >= p['min_units']
    prim = lambda r: r['in_primary_zips'] == 'True'
    targets = [r for r in rental if r['status'] == 'in window' and prim(r) and big(r)]
    watch = [r for r in rental if r not in targets and r['status'] in ('in window', 'next up', 'expired')]
    for r in watch:
        why = {'next up': 'Expires 3 to 6 years out.', 'expired': ''}.get(r['status'], '')
        if r['status'] == 'in window' and not prim(r):
            why = 'In window but outside the primary zips.'
        elif r['status'] == 'in window':
            why = f"In window but under {p['min_units']} units."
        r['note'] = (why + ' ' + r['note']).strip()
    watch.sort(key=lambda r: (r['expiry_date'], -f(r['bldgarea'])))
    targets.sort(key=lambda r: (r['expiry_date'], -f(r['bldgarea'])))

    wb = Workbook()
    m = wb.active
    m.title = 'Method'
    c2 = lambda k: sum(1 for r in rows if r['status'] == k)
    lines = [
        ('421-a expiry sourcing', None), (None, None), ('Input', 'Value'),
        ('As-of date', as_of), ('Class 2 tax rate used for the step-up estimate', log['class_2_rate']),
        ('Minimum building SF', p['min_sf']), ('Minimum residential units', p['min_units']),
        ('Deal-size fit, low units', 50), ('Deal-size fit, high units', 150), ('Expiry window (years)', p['window_years']),
        (None, None), ('Funnel', 'Count'),
        (f"Buildings in Brooklyn CD {p['cd'][1:].lstrip('0')} with {int(p['min_sf']):,}+ SF and residential units", len(rows)),
        ('...with a 421-a exemption on the current roll', len(rows) - c2('no 421-a on roll') - c2('expired')),
        ('...expiring inside the window', c2('in window')),
        ('...rentals in the primary zips with enough units (Targets tab)', len(targets)),
        ('Watchlist tab (next up, edge cases, just expired)', len(watch)),
        (None, None), ('How to read it', None),
        ('Last exempt day', 'June 30 of fiscal year (benefit start + term - 1). Checked against 26 Quincy Street: DOF shows start 2022, term 36, giving June 30, 2057, the date in the offering memorandum.'),
        ('Est. remaining tax step-up', 'Exempt assessed value on the current roll times the class 2 rate. Benefits phase out in the final years, so part of the increase has already hit the bill. Screening figure, not a tax projection.'),
        ('Structure', 'Rental = one tax lot. Rental in condo structure = condo unit lots where one owner holds 80%+ of residential units. Individually owned condos are excluded from Targets and Watchlist and remain on the Universe tab.'),
        ('Owner fields', 'Owner of record is DOF/PLUTO. HPD fields come from the building registration the owner files each year: entity, head officer, managing agent and business addresses.'),
        ('Phone', 'Not in any public city dataset. Yellow column is for a paid lookup (PropertyShark or similar).'),
        ('Deed and mortgage', 'Latest ACRIS deed and mortgage on the lot. Multi-lot deeds show the full consideration. Zero-dollar deeds are internal transfers.'),
        (None, None), ('Sources (NYC Open Data)', None),
        ('PLUTO ' + str(log.get('pluto_version')), 'data.cityofnewyork.us/d/64uk-42ks'),
        ('DOF Property Exemption Detail, roll years ' + ' and '.join(log['exemption_roll_years']), 'data.cityofnewyork.us/d/muvi-b6kx'),
        ('DOF Exemption Classification Codes', 'data.cityofnewyork.us/d/myn9-hwsy'),
        ('DOF Property Valuation and Assessment Data, roll year ' + str(log.get('assessment_roll_year')), 'data.cityofnewyork.us/d/8y4t-faws'),
        ('DOF Property Tax Rates by Tax Class, ' + str(log.get('class_2_rate_year')), 'data.cityofnewyork.us/d/7zb8-7bpk'),
        ('HPD Multiple Dwelling Registrations and Contacts', 'data.cityofnewyork.us/d/tesw-yqqr, /d/feu5-w2e2'),
        ('ACRIS Real Property Legals, Master, Parties', 'data.cityofnewyork.us/d/8h5j-fqxa, /d/bnx9-e6tj, /d/636b-3b5g'),
        (None, None), ('Not yet verified', None),
        ('Tax bills', 'End years are derived from DOF exemption records, not read off each bill. Spot-check two or three bills before outreach.'),
        ('Rents, vacancy, value', 'Phase 2. Nothing here estimates rent or price.')]
    for i, (k, v) in enumerate(lines, 1):
        m.cell(i, 1, k).font = Font(name=F, size=10, bold=v is None or k in ('Input', 'Funnel'))
        c = m.cell(i, 2, v)
        c.font = Font(name=F, size=10, bold=k in ('Input', 'Funnel'), color='0000FF' if 4 <= i <= 10 else '000000')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    m['A1'].font = Font(name=F, bold=True, size=14)
    m['B4'].number_format, m['B5'].number_format, m['B6'].number_format = 'mmm d, yyyy', '0.000%', '#,##0'
    m.column_dimensions['A'].width, m.column_dimensions['B'].width = 58, 95
    for i in range(1, len(lines) + 1):
        m.cell(i, 1).alignment = Alignment(vertical='top', wrap_text=True)

    win = f"{p['window_years']:g}"
    sheet(wb, 'Targets', targets, f"Rentals of {int(p['min_sf']):,}+ SF and {p['min_units']}+ units in zips {p['zips']} whose 421-a ends within {win} years of {as_of:%b %d, %Y}. Sorted by expiry, then size.")
    sheet(wb, 'Watchlist', watch, 'Same screen, one step out: expiring in 3 to 6 years, outside the primary zips, under the unit minimum, or expired in the last two roll years.')
    order = {'in window': 0, 'next up': 1, 'expired': 2, 'later': 3, 'no 421-a on roll': 4}
    uni = sorted(rows, key=lambda r: (order[r['status']], r['expiry_date'] or '9', -f(r['bldgarea'])))
    for r in uni:
        if r not in targets and r not in watch:
            r['note'] = r['status'].capitalize() + ('. Individually owned condo.' if r['rental'] != 'True' else '.')
    sheet(wb, 'Universe', uni, 'Every building that passed the size screen, with or without 421-a, so the filter can be audited.')
    wb.move_sheet('Method', offset=3)
    wb.active = 0
    out = os.path.join(a.dir, a.name)
    wb.save(out)
    print(out, len(targets), len(watch), len(uni))


if __name__ == '__main__':
    main()
