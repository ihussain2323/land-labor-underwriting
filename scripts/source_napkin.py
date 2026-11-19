#!/usr/bin/env python3
"""Phase 2 of the sourcing screen: back-of-napkin value for each target.

source_report.py calls add_napkin() when the run folder holds rents.json, a list of
    {"address", "name", "website", "listings": [{"beds", "rent", "sf", "source_url"}],
     "units_available", "availability_source", "confidence", "notes"}
one object per target, collected by hand or by a research pass from the building's or
manager's own site. Every number on the Napkin tab is a formula over the Rent evidence
tab, the DOF assessed values and the blue inputs at the top of the tab.
"""
import json, os, re, urllib.parse, urllib.request
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

F = 'Arial'
HEAD = PatternFill('solid', start_color='1F3A5F')
BAND = PatternFill('solid', start_color='F3F6FA')
KEY = PatternFill('solid', start_color='FFFF00')
THIN = Side(style='thin', color='D0D5DD')
BLUE, GREEN, BLACK = '0000FF', '008000', '000000'

INPUTS = [  # label, value, format, note
    ('Operating expenses before RE tax ($/unit/yr)', 10000, '$#,##0', 'Deal lead, Nov 18, 2026.'),
    ('Alternative operating expenses before RE tax ($/SF/yr)', 13, '$#,##0.00', 'Deal lead, Nov 18, 2026. Read here as an alternative to $10,000/unit, with RE tax still added. Confirm.'),
    ('Cap rate', 0.0575, '0.00%', 'Deal lead, Nov 18, 2026: 5.75% or higher.'),
    ('Discount from asking rent to in-place rent', 0.0, '0.0%', 'Deal lead uses listed rents as-is and accepts they run high. Raise this to test.'),
    ('Minimum vacancy', 0.03, '0.0%', 'Assumption, not from the deal lead. Vacancy is units advertised / total units, floored here.'),
    ('Deal size, low ($)', 50e6, '$#,##0', 'Deal lead, Nov 18, 2026.'),
    ('Deal size, high ($)', 150e6, '$#,##0', 'Deal lead, Nov 18, 2026.'),
    ('Assumed unit mix: studio', 0.20, '0%', 'Assumption. Unit mix is not public; replace when a rent roll arrives.'),
    ('Assumed unit mix: 1 bedroom', 0.50, '0%', 'Assumption.'),
    ('Assumed unit mix: 2 bedroom', 0.30, '0%', 'Assumption.'),
    ('Assumed unit mix: 3 bedroom', 0.0, '0%', 'Assumption.'),
    ('Area asking rent, studio ($/mo)', "=IFERROR(AVERAGEIFS('Rent evidence'!$D:$D,'Rent evidence'!$C:$C,0),0)", '$#,##0', 'Average of every studio on the Rent evidence tab, all buildings.'),
    ('Area asking rent, 1 bedroom ($/mo)', "=IFERROR(AVERAGEIFS('Rent evidence'!$D:$D,'Rent evidence'!$C:$C,1),0)", '$#,##0', 'Same, 1 bedrooms.'),
    ('Area asking rent, 2 bedroom ($/mo)', "=IFERROR(AVERAGEIFS('Rent evidence'!$D:$D,'Rent evidence'!$C:$C,2),0)", '$#,##0', 'Same, 2 bedrooms.'),
    ('Area asking rent, 3 bedroom ($/mo)', "=IFERROR(AVERAGEIFS('Rent evidence'!$D:$D,'Rent evidence'!$C:$C,3),0)", '$#,##0', 'Same, 3 bedrooms.'),
    ('Area asking rent at the assumed mix ($/unit/mo)', None, '$#,##0', 'Mix-weighted. A building with no listings gets this; a building with listings gets this times its own index.'),
]
R0 = 4                      # first input row
HR = R0 + len(INPUTS) + 2   # table header row
IN = {k: f'$B${R0 + i}' for i, k in enumerate(('opex', 'opex_sf', 'cap', 'disc', 'vac', 'lo', 'hi', 'm0', 'm1', 'm2', 'm3', 'a0', 'a1', 'a2', 'a3', 'dflt'))}
INPUTS[-1] = (INPUTS[-1][0], f"=SUMPRODUCT({IN['m0']}:{IN['m3']},{IN['a0']}:{IN['a3']})") + INPUTS[-1][2:]


def norm(a):
    a = re.sub(r'(\d)(ST|ND|RD|TH)\b', r'\1', a.upper())
    return re.sub(r'\s+', ' ', a.replace('AVE.', 'AVENUE').replace('ST.', 'STREET')).strip()


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def load(run_dir):
    p = os.path.join(run_dir, 'rents.json')
    return {norm(b['address']): b for b in json.load(open(p))} if os.path.exists(p) else None


def pluto_areas(bbls):
    """Residential and commercial SF per lot from PLUTO; empty dict if the call fails."""
    try:
        q = urllib.parse.urlencode({'$select': 'bbl,resarea,comarea', '$where': 'bbl in (' + ','.join(bbls) + ')', '$limit': 5000})
        with urllib.request.urlopen('https://data.cityofnewyork.us/resource/64uk-42ks.json?' + q, timeout=60) as r:
            return {str(int(float(x['bbl']))): (num(x.get('resarea')), num(x.get('comarea'))) for x in json.load(r)}
    except Exception as e:
        print('PLUTO area lookup failed:', e)
        return {}


def head(ws, row, cols):
    for j, (h, w, _) in enumerate(cols, 1):
        c = ws.cell(row, j, h)
        c.font, c.fill = Font(name=F, bold=True, color='FFFFFF', size=9), HEAD
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[row].height = 42


def title(ws, t, intro):
    ws['A1'], ws['A2'] = t, intro
    ws['A1'].font = Font(name=F, bold=True, size=14)
    ws['A2'].font = Font(name=F, italic=True, size=9, color='555555')


def evidence(wb, targets, rents):
    ws = wb.create_sheet('Rent evidence')
    title(ws, 'Rent evidence', 'Every advertised unit behind the Napkin tab. Asking rents, not in-place rents. Add or correct rows here and the Napkin tab follows.')
    cols = [('Address', 24, None), ('Marketing name', 22, None), ('Beds', 6, '0'), ('Asking rent ($/mo)', 11, '$#,##0'), ('SF', 8, '#,##0'),
            ('Rent ($/SF/yr)', 9, '$#,##0.00'), ('Rent vs area for same bed count', 10, '0.00'), ('Source', 60, None)]
    head(ws, 4, cols)
    i = 5
    for r in targets:
        b = rents.get(norm(r['address']))
        for x in (b or {}).get('listings', []):
            rent = num(x.get('rent'))
            if not rent:
                continue
            vals = [r['address'], b.get('name'), num(x.get('beds')), rent, num(x.get('sf')), f'=IF(E{i}="","",D{i}*12/E{i})',
                    f'=IFERROR(D{i}/INDEX(Napkin!{IN["a0"]}:{IN["a3"]},C{i}+1),"")', x.get('source_url')]
            for j, v in enumerate(vals, 1):
                c = ws.cell(i, j, v)
                c.font = Font(name=F, size=9, color=BLUE if j in (3, 4, 5) else BLACK)
                if cols[j - 1][2]:
                    c.number_format = cols[j - 1][2]
            i += 1
    ws.freeze_panes = 'B5'
    ws.auto_filter.ref = f'A4:{L(len(cols))}{max(i - 1, 4)}'
    return i - 5


def add_napkin(wb, targets, rents, rate_ref='Method!$B$5'):
    n_list = evidence(wb, targets, rents)
    ws = wb.create_sheet('Napkin')
    title(ws, 'Back-of-napkin value', 'Asking rents x units, less vacancy, $/unit expenses and full post-expiry RE tax, capped. A first look to rank who to call, not an underwriting. Blue cells are inputs.')
    ws.cell(R0 - 1, 1, 'Input').font = ws.cell(R0 - 1, 2, 'Value').font = Font(name=F, bold=True, size=10)
    for i, (k, v, fmt, note) in enumerate(INPUTS, R0):
        ws.cell(i, 1, k).font = Font(name=F, size=10)
        c = ws.cell(i, 2, v)
        c.font, c.number_format = Font(name=F, size=10, color=GREEN if 'Rent evidence' in str(v) else BLACK if str(v).startswith('=') else BLUE), fmt
        if k in ('Cap rate', 'Discount from asking rent to in-place rent') or k.startswith('Assumed unit mix'):
            c.fill = KEY
        ws.cell(i, 3, note).font = Font(name=F, size=9, italic=True, color='555555')
    cols = [
        ('Address', 44, None), ('Marketing name', 20, None), ('Res. units', 8, '#,##0'), ('Building SF', 11, '#,##0'),
        ('Residential SF', 11, '#,##0'), ('Res. SF per unit', 8, '#,##0'), ('Commercial SF', 10, '#,##0'),
        ('Listings sampled', 8, '0'), ('Rent index vs area', 8, '0.00'), ('Asking rent ($/unit/mo)', 11, '$#,##0'), ('Rent basis', 12, None),
        ('Units advertised', 9, '0'), ('Vacancy', 8, '0.0%'), ('Gross rent ($/yr)', 14, '$#,##0'), ('Effective gross ($/yr)', 14, '$#,##0'),
        ('Operating expenses ($/yr)', 13, '$#,##0'), ('Taxable assessed value after expiry ($)', 15, '$#,##0'),
        ('RE tax after expiry ($/yr)', 13, '$#,##0'), ('RE tax / effective gross', 9, '0%'), ('NOI ($/yr)', 13, '$#,##0'),
        ('Napkin value ($)', 15, '$#,##0'), ('Value per unit ($)', 11, '$#,##0'), ('Value per SF ($)', 9, '$#,##0'),
        ('In deal-size range', 9, None), ('Value with $/SF expenses ($)', 15, '$#,##0'), ('Last mortgage ($)', 14, '$#,##0'),
        ('Mortgage / napkin value', 10, '0%'), ('Last deed', 10, 'mmm yyyy'), ('Last deed ($)', 14, '$#,##0'), ('Rent confidence', 10, None), ('Notes', 70, None)]
    head(ws, HR, cols)
    ix = {h: L(j) for j, (h, *_) in enumerate(cols, 1)}
    A, C, D, E, Fq, H, I, J, K, Lc, M, N, O, P, Q, V = (ix[k] for k in (
        'Address', 'Res. units', 'Building SF', 'Listings sampled', 'Asking rent ($/unit/mo)', 'Units advertised', 'Vacancy',
        'Gross rent ($/yr)', 'Effective gross ($/yr)', 'Operating expenses ($/yr)', 'Taxable assessed value after expiry ($)',
        'RE tax after expiry ($/yr)', 'RE tax / effective gross', 'NOI ($/yr)', 'Napkin value ($)', 'Last mortgage ($)'))
    ev = "'Rent evidence'!"
    areas = pluto_areas([r['bbl'] for r in targets])
    for i, r in enumerate(targets, HR + 1):
        b = rents.get(norm(r['address'])) or {}
        # full tax base once the 421-a comes off: today's taxable value plus the 421-a exempt value
        av = (num(r['assesstot']) or 0) - (num(r['exempttot']) or 0) + (num(r['exempt_av']) or 0)
        notes = [b.get('notes') or ('No rent evidence collected.' if not b else '')]
        if r['structure'] == 'rental in condo structure':
            notes.append('Condo structure: assessed value covers every lot in the condo, including any the main owner does not hold.')
        res, com = areas.get(str(r['bbl']), (None, None))
        if com and com > 500:
            notes.append(f'{com:,.0f} SF of commercial space; its income is not counted but its tax is.')
        if res and num(r['unitsres']) and res / num(r['unitsres']) > 1200:
            notes.append('Large units; a per-unit rent likely understates this building.')
        vals = {
            'Address': r['address'], 'Marketing name': b.get('name'), 'Res. units': num(r['unitsres']), 'Building SF': num(r['bldgarea']),
            'Residential SF': res, 'Commercial SF': com or None,
            'Res. SF per unit': f'=IF(OR({ix["Residential SF"]}{i}="",{C}{i}=0),"",{ix["Residential SF"]}{i}/{C}{i})',
            'Listings sampled': f'=COUNTIFS({ev}$A:$A,{A}{i},{ev}$D:$D,">0")',
            'Rent index vs area': f'=IF({E}{i}>0,AVERAGEIFS({ev}$G:$G,{ev}$A:$A,{A}{i}),1)',
            'Asking rent ($/unit/mo)': f'={ix["Rent index vs area"]}{i}*{IN["dflt"]}',
            'Rent basis': f'=IF({E}{i}>0,"Own listings","Area average")',
            'Units advertised': num(b.get('units_available')),
            'Vacancy': f'=MAX(IF({H}{i}="",0,{H}{i}/{C}{i}),{IN["vac"]})',
            'Gross rent ($/yr)': f'={C}{i}*{Fq}{i}*12*(1-{IN["disc"]})',
            'Effective gross ($/yr)': f'={J}{i}*(1-{I}{i})',
            'Operating expenses ($/yr)': f'={C}{i}*{IN["opex"]}',
            'Taxable assessed value after expiry ($)': av,
            'RE tax after expiry ($/yr)': f'={M}{i}*{rate_ref}',
            'RE tax / effective gross': f'=IF({K}{i}=0,"",{N}{i}/{K}{i})',
            'NOI ($/yr)': f'={K}{i}-{Lc}{i}-{N}{i}',
            'Napkin value ($)': f'={P}{i}/{IN["cap"]}',
            'Value per unit ($)': f'=IF({C}{i}=0,"",{Q}{i}/{C}{i})',
            'Value per SF ($)': f'=IF({D}{i}=0,"",{Q}{i}/{D}{i})',
            'In deal-size range': f'=IF(AND({Q}{i}>={IN["lo"]},{Q}{i}<={IN["hi"]}),"Yes",IF({Q}{i}<{IN["lo"]},"Below","Above"))',
            'Value with $/SF expenses ($)': f'=({K}{i}-{D}{i}*{IN["opex_sf"]}-{N}{i})/{IN["cap"]}',
            'Last mortgage ($)': num(r['last_mtge_amt']) or None,
            'Mortgage / napkin value': f'=IF(OR({V}{i}="",{Q}{i}<=0),"",{V}{i}/{Q}{i})',
            'Last deed': __import__('datetime').date.fromisoformat(r['last_deed_date'][:10]) if r['last_deed_date'] else None,
            'Last deed ($)': num(r['last_deed_amt']) or None,
            'Rent confidence': b.get('confidence') or 'none', 'Notes': ' '.join(x for x in notes if x)}
        for j, (h, _, fmt) in enumerate(cols, 1):
            v = vals[h]
            c = ws.cell(i, j, v)
            formula = isinstance(v, str) and v.startswith('=')
            c.font = Font(name=F, size=9, bold=h == 'Napkin value ($)',
                          color=BLUE if h == 'Units advertised' else GREEN if formula and ('Rent evidence' in v or 'Method' in v) else BLACK)
            c.alignment = Alignment(vertical='top', wrap_text=h == 'Notes')
            c.border = Border(bottom=THIN)
            if fmt:
                c.number_format = fmt
            if (i - HR) % 2 == 0:
                c.fill = BAND
    n = HR + len(targets)
    if targets:
        ws.cell(n + 1, 1, 'Total').font = Font(name=F, bold=True, size=9)
        for h in ('Res. units', 'Building SF', 'Gross rent ($/yr)', 'Effective gross ($/yr)', 'Operating expenses ($/yr)',
                  'RE tax after expiry ($/yr)', 'NOI ($/yr)', 'Napkin value ($)'):
            c = ws.cell(n + 1, cols.index(next(x for x in cols if x[0] == h)) + 1, f'=SUM({ix[h]}{HR + 1}:{ix[h]}{n})')
            c.font, c.number_format = Font(name=F, bold=True, size=9), next(x[2] for x in cols if x[0] == h)
        ws.auto_filter.ref = f'A{HR}:{L(len(cols))}{n}'
    ws.column_dimensions['C'].width = 11
    ws.freeze_panes = ws.cell(HR + 1, 2)
    return n_list
