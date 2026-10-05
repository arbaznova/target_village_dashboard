"""Build data.js for the Target Village Dashboard.

Sources (all opened read-only, never modified):
  raw_data/*.xlsx                         target-village lists, booth bands, polling-station names
  form20 up/2022/AC*.xls(x)               Form-20 booth results for every candidate + electors
  data-analysis/cleaned_data/booth_level  approximate village coordinates (for "open in maps" links)

Run:  <python with openpyxl> build_data.py
"""
import csv, glob, json, os, re, shutil, subprocess, sys, tempfile
from datetime import date

import openpyxl

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
RAW = os.path.join(ROOT, 'raw_data')
F20 = os.path.join(ROOT, 'form20 up', '2022')
GEO = os.path.join(ROOT, 'data-analysis', 'cleaned_data', 'booth_level')
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data.js')

# key, AC no, name, district, workbook glob, form20 file, form20 sheet, electors col, turnout col, first slot col, slot width
ACS = [
    ('Deoband', 5, 'Deoband', 'Saharanpur', 'Deoband_AC5_*', 'AC5.xls', None, 3, 7, 10, 2),
    ('Budhana', 11, 'Budhana', 'Muzaffarnagar', 'Budhana_AC11_*', 'AC11.xls', '11-budhana', 3, 7, 10, 3),
    ('Charthawal', 12, 'Charthawal', 'Muzaffarnagar', 'Charthawal_AC12_*', 'AC12.xls', '12-Charthawal', 3, 7, 10, 3),
    ('Khatauli', 15, 'Khatauli', 'Muzaffarnagar', 'Khatauli_AC15_*', 'AC15.xls', '15-Khatauli', 3, 7, 10, 3),
    ('Kanth', 25, 'Kanth', 'Moradabad', 'Kanth_AC25_*', 'AC25.xlsx', None, 3, 7, 10, 3),
    ('Dhanaura', 39, 'Dhanaura (SC)', 'Amroha', 'Dhanaura_AC39_*', 'AC39.xls', 'ECI Format. (3)', 3, 7, 10, 3),
    ('Hastinapur', 45, 'Hastinapur (SC)', 'Meerut', 'Hastinapur_AC45_*', 'AC45.xlsx', None, 3, 7, 11, 3),
    ('Chhaprauli', 50, 'Chhaprauli', 'Baghpat', 'Chhaprauli_*', 'AC50.xls', 'Form 20 ECI', 3, 7, 10, 3),
    ('Baraut', 51, 'Baraut', 'Baghpat', 'Baraut_*', 'AC51.xls', None, 6, 10, 13, 3),
    ('Bagpat', 52, 'Baghpat', 'Baghpat', 'Bagpat_AC52_*', 'AC52.xls', None, 3, 7, 10, 3),
    ('Hapur', 59, 'Hapur (SC)', 'Hapur', 'Hapur_AC59_*', 'AC59.xlsx', None, 3, 7, 10, 3),
    ('Garhmukteshwar', 60, 'Garhmukteshwar', 'Hapur', 'Garhmukteshwar_AC60_*', 'AC60.xlsx', None, 3, 7, 10, 3),
]


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return None


def booth_key(v):
    """Normalise a booth number: 12 / 12.0 / '12' -> '12'; '304 A' -> '304A'."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r'\s+', '', str(v)).upper()
    return s or None


def booth_list(v):
    if v is None:
        return []
    if isinstance(v, (int, float)):
        return [booth_key(v)]
    s = str(v).strip()
    if not re.match(r'^[\d\s,A-Za-z–-]+$', s) or s.upper() in ('NA', 'MAPPED'):
        return []
    out = []
    for part in re.split(r'\s*,\s*', s):
        m = re.match(r'^(\d+)\s*[–-]\s*(\d+)$', part)
        if m:
            out += [str(i) for i in range(int(m.group(1)), int(m.group(2)) + 1)]
        elif re.match(r'^\d+\s*[A-Za-z]?$', part):
            out.append(booth_key(part))
    return out


def party_code(name):
    if name is None:
        return None
    n = re.sub(r'[\s\xa0]+', ' ', str(name)).strip().upper()
    if not n:
        return None
    if n in ('NOTA', 'NA') or 'NOTA' in n:
        return 'NOTA'
    if n == 'BJP' or ('BHARATIYA' in n or 'BHARTIYA' in n or 'BHARTIY' in n) and 'ANATA PARTY' in n.replace('JANTA', 'JANATA'):
        return 'BJP'
    if n == 'RLD' or 'LOK DAL' in n and 'RASHTRIYA' in n:
        return 'RLD'
    if n == 'SP' or 'SAMAJWADI' in n:
        return 'SP'
    if n == 'BSP' or 'BAHUJAN SAMAJ' in n:
        return 'BSP'
    if n == 'INC' or 'CONGRESS' in n:
        return 'INC'
    if n == 'AIMIM' or 'MAJLIS' in n:
        return 'AIMIM'
    if n == 'AAP' or 'AAM AADMI' in n:
        return 'AAP'
    if n.startswith('ASP') or 'AAZAD SAMAJ' in n:
        return 'ASP'
    if n.startswith('IND') or 'INDEPENDENT' in n or 'INDENDENT' in n:
        return 'IND'
    return n if len(n) <= 8 else 'OTH'


def open_wb(path):
    return openpyxl.load_workbook(path, read_only=True, data_only=True)


def rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


# ---------------------------------------------------------------- Form 20
def convert_xls(tmp, fname):
    src = os.path.join(F20, fname)
    if fname.endswith('.xlsx'):
        return src
    shutil.copy(src, tmp)
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', tmp,
                    os.path.join(tmp, fname)], check=True, capture_output=True)
    return os.path.join(tmp, fname[:-4] + '.xlsx')


def parse_form20(path, sheet, ecol, tcol, scol, width):
    wb = open_wb(path)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rs = rows(ws)
    data, header_rows = [], []
    for r in rs:
        r = r + [None] * 5
        b, nm, el = r[1], r[2], num(r[ecol])
        if (isinstance(b, (int, float)) or (isinstance(b, str) and re.match(r'^\d+\s*[A-Za-z]?$', b.strip()))) \
                and isinstance(nm, str) and el is not None:
            data.append(r)
        elif not data:
            header_rows.append(r)
    # party per slot: first non-empty party label in data rows, else header rows
    slots = []
    c = scol
    while c + width - 1 < len(data[0]):
        pcol, vcol = (c + 1, c + 2) if width == 3 else (c, c + 1)
        if width == 3 and num(data[0][c]) != len(slots) + 1:
            break
        party = None
        for r in data[:5] + header_rows[::-1]:
            if isinstance(r[pcol], str) and r[pcol].strip() and not re.search(r'party affil|votes', r[pcol], re.I):
                party = r[pcol]
                break
        if width == 2 and party is None:
            break
        slots.append((party_code(party), vcol))
        c += width
    if width == 2:
        slots = [s for s in slots if s[0]]
    out = {}
    for r in data:
        votes = {}
        for p, vc in slots:
            v = num(r[vc]) or 0
            p = p or 'OTH'
            votes[p] = votes.get(p, 0) + int(v)
        tv = int(num(r[tcol]) or 0)
        out[booth_key(r[1])] = {'electors': int(num(r[ecol]) or 0), 'turnout': tv, 'votes': votes}
    return out, [s[0] for s in slots]


# ---------------------------------------------------------------- workbook booths
def wb_booths(wb):
    ws = [w for w in wb.worksheets if re.match(r'(ALL BOOTHS|All \d+ Booths)', w.title)][0]
    rs = rows(ws)
    hi = next(i for i, r in enumerate(rs) if r and isinstance(r[0], str) and r[0].startswith('Booth')
              or (r and len(r) > 1 and r[1] == 'Booth No.'))
    hdr = [str(h or '').strip() for h in rs[hi]]

    def col(*names):
        for n in names:
            for i, h in enumerate(hdr):
                if h.lower() == n.lower():
                    return i
        for n in names:
            for i, h in enumerate(hdr):
                if h.lower().startswith(n.lower()):
                    return i
        return None
    cb = col('Booth No.')
    cs = col('Polling Station (Unicode Hindi)', 'Polling Station (Unicode)', 'Polling Station – Unicode Hindi',
             'Polling Station (Hindi)', 'Polling Station')
    cbjp, crld = col('BJP Votes', 'Bharatiya Janata Party Votes'), col('RLD Votes')
    out = {}
    for r in rs[hi + 1:]:
        if not r or r[cb] is None or num(r[cbjp]) is None:
            continue
        out[booth_key(r[cb])] = {'station': re.sub(r'\s+', ' ', str(r[cs] or '')).strip(),
                                 'bjp': int(num(r[cbjp])), 'rld': int(num(r[crld])) if crld is not None and num(r[crld]) is not None else None}
    return out


# ---------------------------------------------------------------- target villages
def clean(s):
    return re.sub(r'\s+', ' ', str(s)).strip() if s is not None else ''


def targets_from_source_col(wb, list_sheet, src_sheet_pat, name_col=0, info_col=1):
    """Bagpat / Baraut / Chhaprauli: booths are tagged in the Source Data sheet's 'Target Village' column."""
    ws = [w for w in wb.worksheets if re.search(src_sheet_pat, w.title)][0]
    rs = rows(ws)
    hi = next(i for i, r in enumerate(rs) if 'Target Village' in [clean(c) for c in r])
    hdr = [clean(c) for c in rs[hi]]
    cb, ct = hdr.index('Booth No.'), hdr.index('Target Village')
    tagged = {}
    for r in rs[hi + 1:]:
        if r[cb] is not None and r[ct]:
            tagged.setdefault(clean(r[ct]), []).append(booth_key(r[cb]))
    out = []
    lrs = rows(wb[list_sheet])
    hi = next(i for i, r in enumerate(lrs) if 'Target Village' in [clean(c) for c in r])
    hdr = [clean(c) for c in lrs[hi]]
    cn = hdr.index('Target Village')
    for r in lrs[hi + 1:]:
        nm = clean(r[cn]) if cn < len(r) else ''
        if not nm or nm.startswith('UNTRACED'):
            break
        info = clean(r[cn + 1]) if cn + 1 < len(r) else ''
        booths = tagged.get(nm, [])
        out.append({'name': nm, 'booths': booths, 'status': 'Mapped' if booths else 'NA',
                    'note': info.replace('\n', ' | ') if info and info != nm else '', 'annotation': ''})
    return out


def targets_from_table(wb):
    ws = wb['TARGET VILLAGES']
    rs = rows(ws)
    hi = next(i for i, r in enumerate(rs) if any(clean(c) in ('Target Village', 'Status') for c in r))
    hdr = [clean(c) for c in rs[hi]]
    ln = [h.lower() for h in hdr]

    def col(*keys):
        for k in keys:
            for i, h in enumerate(ln):
                if h.startswith(k.lower()):
                    return i
        return None
    cn = col('Target Village')
    if cn is None:  # Kanth's header is mislabelled 'Booth No.' but holds the village name
        cn = 0
    cst = col('Status')
    cbs = [col('Matched Booths', 'Booth(s)', 'Booths')]
    if cbs[0] is None:
        cbs = [i for i, h in enumerate(hdr) if re.match(r'Booth \d', h)]
    cnote = col('Notes', 'Related', 'Source Match', 'Note')
    cann = col('Target Annotation')
    out, seen = [], set()
    for r in rs[hi + 1:]:
        r = r + [None] * 12
        nm = clean(r[cn])
        if not nm or r[cst] is None:
            if out and not nm:
                continue
            if out:
                break
            continue
        if nm in seen or re.match(r'^Booth \d', nm) or nm.lower().startswith('clickable'):
            continue
        seen.add(nm)
        st = clean(r[cst])
        booths = []
        for c in cbs:
            booths += booth_list(r[c])
        note = clean(r[cnote]) if cnote is not None else ''
        if st.upper() == 'NA' or not booths:
            longs = [clean(c) for c in r if isinstance(c, str) and len(clean(c)) > 12 and clean(c) != nm]
            note = max(longs, key=len) if longs else note
        if (st.upper() == 'NA' or not booths) and not note:
            note = 'Not directly traceable in polling-station names; not guessed.'
        ann = clean(r[cann]) if cann is not None else ''
        if ann and booth_list(ann):  # Budhana: annotation column shifted when empty
            ann = ''
        out.append({'name': nm, 'booths': [] if st.upper() == 'NA' else booths,
                    'status': 'NA' if st.upper() == 'NA' or not booths else st,
                    'note': note, 'annotation': ann})
    return out


def targets_for(key, wb):
    if key == 'Bagpat':
        t = targets_from_source_col(wb, 'Target Village Analysis', r'Source Data')
    elif key in ('Baraut', 'Chhaprauli'):
        t = targets_from_source_col(wb, 'Target Village Analysis', r'Source Data')
    else:
        t = targets_from_table(wb)
    return t


# ---------------------------------------------------------------- notes / NA hints
def wb_notes(wb):
    for w in wb.worksheets:
        if re.search(r'notes', w.title, re.I):
            lines = []
            for r in rows(w):
                cells = [clean(c) for c in r if c is not None and not str(c).startswith('←')]
                if cells:
                    lines.append(' — '.join(cells))
            return lines[1:]
    return []


def geo_for(key, acno):
    fs = glob.glob(os.path.join(GEO, f'{key}_AC{acno}_*.csv'))
    out = {}
    if not fs:
        return out
    with open(fs[0], encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            b = booth_key(r.get('Booth No.'))
            lat, lon = num(r.get('Latitude')), num(r.get('Longitude'))
            out[b] = {'village_en': (r.get('Village') or '').strip(), 'lat': lat, 'lon': lon}
    return out


def main():
    tmp = tempfile.mkdtemp()
    acs_out, problems = [], []
    try:
        for key, acno, name, dist, pat, f20, sheet, ecol, tcol, scol, width in ACS:
            wbp = glob.glob(os.path.join(RAW, pat + '.xlsx'))[0]
            wb = open_wb(wbp)
            wbb = wb_booths(wb)
            f20d, parties = parse_form20(convert_xls(tmp, f20), sheet, ecol, tcol, scol, width)
            geo = geo_for(key, acno)
            # cross-check workbook BJP votes against Form 20
            mism = [b for b in wbb if b in f20d and f20d[b]['votes'].get('BJP') != wbb[b]['bjp']]
            missing = [b for b in wbb if b not in f20d]
            if mism or missing:
                problems.append(f'{key}: {len(mism)} BJP mismatches {mism[:5]}, {len(missing)} booths missing in Form20 {missing[:5]}')
            booths = {}
            for b, w in wbb.items():
                f = f20d.get(b, {})
                votes = dict(f.get('votes') or {'BJP': w['bjp']})
                if w['rld'] is not None and 'RLD' not in votes:
                    votes['RLD'] = w['rld']
                g = geo.get(b, {})
                booths[b] = {'no': b, 'station': w['station'], 'electors': f.get('electors'),
                             'turnout': f.get('turnout') or sum(votes.values()), 'votes': votes,
                             'village_en': g.get('village_en', ''), 'lat': g.get('lat'), 'lon': g.get('lon')}
            targets = targets_for(key, wb)
            for t in targets:
                lost = [b for b in t['booths'] if b not in booths]
                if lost:
                    problems.append(f'{key}/{t["name"]}: booths not found {lost}')
                t['booths'] = [b for b in t['booths'] if b in booths]
            used = sorted({b for t in targets for b in t['booths']}, key=lambda x: (int(re.match(r'\d+', x).group()), x))
            # AC totals over all booths (for "vs AC average" comparisons)
            tot = {'electors': 0, 'turnout': 0, 'votes': {}}
            for b in booths.values():
                tot['electors'] += b['electors'] or 0
                tot['turnout'] += b['turnout'] or 0
                for p, v in b['votes'].items():
                    tot['votes'][p] = tot['votes'].get(p, 0) + v
            persp = 'BJP × RLD' if any(b['votes'].get('RLD') for b in booths.values()) else 'BJP'
            acs_out.append({'key': key, 'no': acno, 'name': name, 'district': dist, 'workbook': os.path.basename(wbp),
                            'perspective': persp, 'boothCount': len(booths), 'totals': tot,
                            'targets': targets, 'booths': {b: booths[b] for b in used}, 'notes': wb_notes(wb)})
            print(f'{key:15} booths={len(booths):4} targets={len(targets):3} mapped={sum(1 for t in targets if t["booths"]):3} '
                  f'targetBooths={len(used):3} parties={parties}')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    payload = {'generated': date.today().isoformat(), 'election': 'UP Vidhan Sabha 2022', 'acs': acs_out, 'buildWarnings': problems}
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('window.DASH_DATA = ' + json.dumps(payload, ensure_ascii=False, separators=(',', ':')) + ';\n')
    print('\nwarnings:', *problems, sep='\n  ')
    print('wrote', OUT, os.path.getsize(OUT), 'bytes')


if __name__ == '__main__':
    main()
