#!/usr/bin/env python3
"""Source rental buildings whose 421-a exemption is about to expire.

Pulls public NYC records only (no listings, no paid data):
  PLUTO                         size, units, owner of record, assessed values
  DOF Property Exemption Detail 421-a code, benefit start, term
  DOF Exemption Classification  code descriptions
  HPD Registrations + Contacts  registered owner, head officer, agent, business address
  ACRIS                         last deed and last mortgage (date, amount, parties)

Benefit end rule: last exempt fiscal year = benefit start + term - 1, ending June 30.
Each roll year's records must be read with period 3 (final roll) where present.
Calibrated on 26 Quincy Street (BBL 3019720010): start 2022, term 36 -> June 30, 2057,
which matches the offering memorandum.

Usage:
  python3 scripts/source_421a.py --out DIR [--cd 301] [--zips 11211,11249,11206]
      [--min-sf 50000] [--min-units 40] [--window-years 3] [--as-of 2026-11-18]
"""
import argparse, csv, datetime as dt, json, os, sys, time, urllib.parse, urllib.request
from collections import Counter, defaultdict

BASE = 'https://data.cityofnewyork.us/resource/'
PLUTO, EXM, CODES = '64uk-42ks', 'muvi-b6kx', 'myn9-hwsy'
HPD_REG, HPD_CON = 'tesw-yqqr', 'feu5-w2e2'
AC_LEGAL, AC_MASTER, AC_PARTY = '8h5j-fqxa', 'bnx9-e6tj', '636b-3b5g'
AV, RATES = '8y4t-faws', '7zb8-7bpk'
# 421-a codes from the DOF classification table. Construction-period (-C) records share the code.
CODES_421A = ['5110', '5113', '5114', '5116', '5117', '5118', '5119', '5120', '5121', '5122', '5123']


def get(ds, **params):
    """Fetch every row for a SoQL query, paging 50k at a time, with retries."""
    rows, off = [], 0
    while True:
        q = {('$' + k): v for k, v in params.items()}
        q['$limit'], q['$offset'] = 50000, off
        url = BASE + ds + '.json?' + urllib.parse.urlencode(q)
        for attempt in range(5):
            try:
                with urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': 'sourcing-script'}), timeout=180) as r:
                    page = json.load(r)
                break
            except Exception as e:  # network or throttling
                if attempt == 4:
                    raise
                time.sleep(3 * (attempt + 1))
        rows += page
        if len(page) < 50000:
            return rows
        off += 50000


def chunks(xs, n):
    xs = list(xs)
    for i in range(0, len(xs), n):
        yield xs[i:i + n]


def inlist(xs):
    return '(' + ','.join("'" + str(x).replace("'", "''") + "'" for x in xs) + ')'


def num(x, default=0.0):
    try:
        return float(str(x).replace('+', ''))
    except (TypeError, ValueError):
        return default


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', required=True)
    ap.add_argument('--cd', default='301', help='PLUTO community district (301 = Brooklyn CD 1)')
    ap.add_argument('--zips', default='11211,11249,11206', help='primary zips; others in the CD go to the watchlist')
    ap.add_argument('--min-sf', type=float, default=50000)
    ap.add_argument('--min-units', type=int, default=40)
    ap.add_argument('--window-years', type=float, default=3)
    ap.add_argument('--as-of', default=dt.date.today().isoformat())
    ap.add_argument('--tax-rate', type=float, default=None, help='class 2 rate override; default is the latest published DOF rate')
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    as_of = dt.date.fromisoformat(a.as_of)
    zips = set(a.zips.split(','))
    boro = a.cd[0]
    log = {'as_of': a.as_of, 'params': vars(a).copy()}

    rates = sorted(get(RATES), key=lambda r: r['year'])
    rate = a.tax_rate or num(rates[-1]['class_2']) / 100
    log['class_2_rate'], log['class_2_rate_year'] = rate, rates[-1]['year']

    # 1. Universe
    pl = get(PLUTO, where=f"cd={a.cd} AND bldgarea>={a.min_sf} AND unitsres>0",
             select='bbl,borough,block,lot,address,zipcode,bldgclass,landuse,ownername,bldgarea,resarea,comarea,unitsres,unitstotal,numfloors,yearbuilt,yearalter1,assesstot,exempttot,condono,latitude,longitude,version')
    for p in pl:
        p['bbl'] = str(int(float(p['bbl'])))
        p['is_condo'] = p['bldgclass'].startswith('R') or int(p['lot']) >= 7500
    log['pluto_version'] = pl[0].get('version') if pl else None
    log['universe'] = len(pl)
    blocks = sorted({p['block'] for p in pl})

    # 2. 421-a exemptions on those blocks, latest two roll years
    years = sorted({r['year'] for r in get(EXM, select='year', group='year')})
    latest, prior = years[-1], years[-2]
    ex = []
    for ch in chunks(sorted({str(int(b)) for b in blocks}), 150):
        ex += get(EXM, where=f"boro='{boro}' AND year in ('{latest}','{prior}') AND exmp_code in {inlist(CODES_421A)} AND block in {inlist(ch)}",
                  select='parid,block,lot,year,period,exmp_code,exname,benftstart,no_years,curexmptot,status,bldg_class,extractdt')
    desc = {r['exempt_code']: r['description'] for r in get(CODES) if r['exempt_code'] in CODES_421A}
    log['exemption_roll_years'] = [prior, latest]
    log['exemption_rows'] = len(ex)

    def best(rows):
        """Latest roll year and period for each parcel and code, then the longest-running benefit."""
        if not rows:
            return None
        top = max((r['year'], r['period']) for r in rows)
        cur = [r for r in rows if (r['year'], r['period']) == top]
        r = max(cur, key=lambda r: num(r['benftstart']) + num(r['no_years']))
        r = dict(r)
        r['curexmptot_sum'] = sum(num(x['curexmptot']) for x in cur)
        r['end_fy'] = int(num(r['benftstart']) + num(r['no_years']) - 1)
        return r

    by_parid = defaultdict(list)
    for r in ex:
        by_parid[r['parid']].append(r)

    # Condo-structured buildings: the exemption sits on unit lots. Map unit lots to the
    # building through the condo number on the DOF assessment roll.
    condo_lots = defaultdict(list)
    cblocks = sorted({str(int(p['block'])) for p in pl if p['is_condo']})
    if cblocks:
        av_year = max(r['year'] for r in get(AV, select='year', group='year', where=f"boro='{boro}' AND block='{cblocks[0]}'"))
        for ch in chunks(cblocks, 60):
            for u in get(AV, where=f"boro='{boro}' AND year='{av_year}' AND period='3' AND lot>'1000' AND block in {inlist(ch)}",
                         select='parid,block,lot,condo_number,units,coop_apts,owner,bldg_class,curtaxclass'):
                if u.get('condo_number'):
                    condo_lots[(str(int(u['block'])), int(u['condo_number']) % 100000)].append(u)
        log['assessment_roll_year'] = av_year

    out = []
    for p in pl:
        row = dict(p)
        e, basis = best(by_parid.get(p['bbl'], [])), 'tax lot'
        if p['is_condo']:
            lots = condo_lots.get((str(int(p['block'])), int(num(p.get('condono')))), [])
            res = [u for u in lots if u.get('bldg_class', '').startswith('R') and u.get('bldg_class') not in ('RG', 'RP', 'RS', 'RK', 'RB', 'RA', 'RH', 'RW', 'RT')]
            n_units = sum(num(u.get('units')) for u in res) or len(res)
            owners = Counter()
            for u in res:
                owners[(u.get('owner') or '').strip().upper()] += max(num(u.get('units')), 1)
            top_owner, top_units = (owners.most_common(1)[0] if owners else ('', 0))
            share = top_units / n_units if n_units else 0
            row.update(condo_unit_lots=len(res), condo_top_owner=top_owner, condo_top_owner_share=round(share, 2))
            # Few unit lots for many apartments means the apartments are bundled into one or two lots.
            bundled = len(res) > 0 and len(res) < 0.5 * num(p['unitsres'])
            row['structure'] = 'rental in condo structure' if bundled or (share >= 0.8 and n_units >= 10) else 'condo, individually owned'
            rows = [r for u in lots for r in by_parid.get(u['parid'], [])]
            if rows:
                top = max((r['year'], r['period']) for r in rows)
                cur = [r for r in rows if (r['year'], r['period']) == top]
                ends = Counter(int(num(r['benftstart']) + num(r['no_years']) - 1) for r in cur)
                end = ends.most_common(1)[0][0]
                m = next(r for r in cur if int(num(r['benftstart']) + num(r['no_years']) - 1) == end)
                e = dict(m, end_fy=end, curexmptot_sum=sum(num(r['curexmptot']) for r in cur))
                basis = f'{len(cur)} condo unit lots' + ('' if len(ends) == 1 else f', {len(ends)} end years')
            else:
                e = None
        else:
            row['structure'] = 'rental'
        if e:
            row.update(exm_code=e['exmp_code'], exm_desc=desc.get(e['exmp_code'], ''), exm_name=e.get('exname', ''),
                       benefit_start_fy=int(num(e['benftstart'])), term_years=int(num(e['no_years'])), end_fy=e['end_fy'],
                       expiry_date=f"{e['end_fy']}-06-30", exempt_av=e['curexmptot_sum'], roll_year=e['year'], match_basis=basis)
            exp = dt.date(e['end_fy'], 6, 30)
            row['years_to_expiry'] = round((exp - as_of).days / 365.25, 2)
            row['est_tax_step_up'] = round(e['curexmptot_sum'] * rate)
        out.append(row)

    def status(r):
        if 'end_fy' not in r:
            return 'no 421-a on roll'
        y = r['years_to_expiry']
        if r['roll_year'] != latest or y < 0:
            return 'expired'
        if y <= a.window_years:
            return 'in window'
        if y <= a.window_years + 3:
            return 'next up'
        return 'later'
    for r in out:
        r['status'] = status(r)
        r['in_primary_zips'] = r.get('zipcode') in zips
        r['meets_units'] = num(r['unitsres']) >= a.min_units
        r['rental'] = r['structure'] != 'condo, individually owned'

    # 3. Owner contacts from HPD registrations (candidates only)
    cand = [r for r in out if r['status'] in ('in window', 'next up', 'expired')]
    regs = []
    for ch in chunks(sorted({str(int(r['block'])) for r in cand}), 150):
        regs += get(HPD_REG, where=f"boroid='{boro}' AND block in {inlist(ch)}", select='registrationid,block,lot,lastregistrationdate,registrationenddate')
    reg_by = defaultdict(list)
    for g in regs:
        reg_by[(str(int(g['block'])), str(int(g['lot'])))].append(g)
    want = {}
    for r in cand:
        gs = reg_by.get((str(int(r['block'])), str(int(r['lot']))), [])
        if gs:
            g = max(gs, key=lambda g: g.get('lastregistrationdate', ''))
            want[g['registrationid']] = r
            r['hpd_last_registered'] = g.get('lastregistrationdate', '')[:10]
    cons = []
    for ch in chunks(want, 300):
        cons += get(HPD_CON, where=f"registrationid in {inlist(ch)}")
    for c in cons:
        r = want[c['registrationid']]
        name = ' '.join(x for x in [c.get('firstname'), c.get('lastname')] if x)
        addr = ' '.join(x for x in [c.get('businesshousenumber'), c.get('businessstreetname'), c.get('businessapartment')] if x)
        addr = ', '.join(x for x in [addr, c.get('businesscity'), ' '.join(x for x in [c.get('businessstate'), c.get('businesszip')] if x)] if x)
        t = c.get('type')
        if t == 'CorporateOwner':
            r['hpd_owner_entity'], r['hpd_owner_address'] = c.get('corporationname', ''), addr
        elif t in ('HeadOfficer', 'IndividualOwner'):
            r.setdefault('hpd_principal', name)
            r.setdefault('hpd_principal_address', addr)
        elif t == 'Agent':
            r['hpd_agent'] = ' / '.join(x for x in [name, c.get('corporationname')] if x)
            r['hpd_agent_address'] = addr
        elif t == 'SiteManager':
            r['hpd_site_manager'] = name

    # 4. ACRIS last deed and mortgage (in-window and next-up rentals only)
    deep = [r for r in cand if r['status'] != 'expired' or r['in_primary_zips']]
    doc_to = defaultdict(list)
    for r in deep:
        for l in get(AC_LEGAL, where=f"borough='{boro}' AND block='{int(r['block'])}' AND lot='{int(r['lot'])}'", select='document_id'):
            doc_to[l['document_id']].append(r)
    masters = []
    for ch in chunks(doc_to, 200):
        masters += get(AC_MASTER, where=f"document_id in {inlist(ch)} AND doc_type in ('DEED','MTGE','AGMT')",
                       select='document_id,doc_type,document_date,recorded_datetime,document_amt')
    pick = {}
    for m in masters:
        d = (m.get('document_date') or m.get('recorded_datetime') or '')[:10]
        for r in doc_to[m['document_id']]:
            kind = 'deed' if m['doc_type'] == 'DEED' else 'mtge'
            if kind == 'deed' and num(m.get('document_amt')) <= 0 and r.get('_deed_amt', 0) > 0:
                continue
            if d > r.get(f'_{kind}_date', ''):
                r[f'_{kind}_date'], r[f'_{kind}_amt'], r[f'_{kind}_doc'] = d, num(m.get('document_amt')), m['document_id']
    need = {r[k]: (r, k) for r in deep for k in ('_deed_doc', '_mtge_doc') if r.get(k)}
    for ch in chunks(need, 200):
        for p in get(AC_PARTY, where=f"document_id in {inlist(ch)}", select='document_id,party_type,name'):
            r, k = need[p['document_id']]
            side = {'1': 'from', '2': 'to'}.get(p['party_type'])
            if side:
                key = ('last_deed_' if k == '_deed_doc' else 'last_mtge_') + side
                if p['name'] not in r.get(key, ''):
                    r[key] = (r.get(key, '') + '; ' + p['name']).strip('; ')
    for r in deep:
        r['last_deed_date'], r['last_deed_amt'] = r.pop('_deed_date', ''), r.pop('_deed_amt', '')
        r['last_mtge_date'], r['last_mtge_amt'] = r.pop('_mtge_date', ''), r.pop('_mtge_amt', '')
        r.pop('_deed_doc', None), r.pop('_mtge_doc', None)

    cols = ['status', 'address', 'zipcode', 'bbl', 'bldgclass', 'structure', 'rental', 'bldgarea', 'unitsres', 'unitstotal', 'numfloors', 'yearbuilt',
            'exm_code', 'exm_desc', 'benefit_start_fy', 'term_years', 'end_fy', 'expiry_date', 'years_to_expiry', 'exempt_av', 'est_tax_step_up',
            'match_basis', 'condo_unit_lots', 'condo_top_owner', 'condo_top_owner_share', 'ownername', 'exm_name', 'hpd_owner_entity', 'hpd_owner_address', 'hpd_principal', 'hpd_principal_address',
            'hpd_agent', 'hpd_agent_address', 'hpd_site_manager', 'hpd_last_registered', 'last_deed_date', 'last_deed_amt', 'last_deed_from',
            'last_deed_to', 'last_mtge_date', 'last_mtge_amt', 'last_mtge_to', 'in_primary_zips', 'meets_units', 'assesstot', 'exempttot',
            'latitude', 'longitude']
    order = {'in window': 0, 'next up': 1, 'expired': 2, 'later': 3, 'no 421-a on roll': 4}
    out.sort(key=lambda r: (order[r['status']], r.get('end_fy', 9999), -num(r['bldgarea'])))
    with open(os.path.join(a.out, 'universe.csv'), 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction='ignore')
        w.writeheader()
        w.writerows(out)
    log['counts'] = dict(Counter(r['status'] for r in out))
    log['targets'] = sum(1 for r in out if r['status'] == 'in window' and r['in_primary_zips'] and r['rental'])
    log['hpd_matched'] = sum(1 for r in cand if r.get('hpd_owner_entity') or r.get('hpd_principal'))
    log['candidates'] = len(cand)
    with open(os.path.join(a.out, 'run-log.json'), 'w') as f:
        json.dump(log, f, indent=1, default=str)
    print(json.dumps(log, indent=1, default=str))


if __name__ == '__main__':
    main()
