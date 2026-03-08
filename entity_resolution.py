import csv
import datetime as dt
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def normalize_target(target: str) -> str:
    t = target.lstrip('/')
    return t if t.startswith('xl/') else f'xl/{t}'


def read_xlsx_rows(path: str, sheet_name: str):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(f'{NS}si'):
                shared.append(''.join((t.text or '') for t in si.iter(f'{NS}t')))

        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rel = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid_to_target = {r.attrib['Id']: normalize_target(r.attrib['Target']) for r in rel}

        target = None
        for sh in wb.findall(f'{NS}sheets/{NS}sheet'):
            if sh.attrib['name'] == sheet_name:
                target = rid_to_target[sh.attrib[REL_NS]]
                break
        if target is None:
            raise ValueError(f'Sheet {sheet_name} not found in {path}')

        sheet = ET.fromstring(z.read(target))
        out = []
        for row in sheet.findall(f'.//{NS}sheetData/{NS}row'):
            vals = {}
            for c in row.findall(f'{NS}c'):
                ref = c.attrib.get('r', '')
                idx = col_to_idx(''.join(ch for ch in ref if ch.isalpha()))
                t = c.attrib.get('t')
                if t == 'inlineStr':
                    tn = c.find(f'{NS}is/{NS}t')
                    val = tn.text if tn is not None else ''
                else:
                    v = c.find(f'{NS}v')
                    if v is None or v.text is None:
                        val = ''
                    elif t == 's':
                        val = shared[int(v.text)]
                    else:
                        val = v.text
                vals[idx] = val
            if vals:
                row_out = [''] * (max(vals) + 1)
                for i, v in vals.items():
                    row_out[i] = v
                out.append(row_out)
        return out


def clean_text(s: str) -> str:
    return re.sub(r'\s+', ' ', (s or '').strip().lower())


def alpha(s: str) -> str:
    return re.sub(r'[^a-z]', '', clean_text(s))


def parse_phone(s: str) -> str:
    digits = re.sub(r'\D', '', s or '')
    if len(digits) < 10:
        return ''
    if len(digits) > 11:
        digits = digits[-11:]
    return digits


def parse_date_any(s: str) -> str:
    s = (s or '').strip()
    if not s:
        return ''

    iso_match = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if iso_match:
        return f"{iso_match.group(1)}-{iso_match.group(2)}-{iso_match.group(3)}"

    for fmt in ('%d.%m.%Y', '%d.%m.%y', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y', '%Y-%m-%d'):
        try:
            d = dt.datetime.strptime(s, fmt).date()
            return d.isoformat()
        except ValueError:
            pass

    m = re.match(r'^(\d{2})\.(\d{2})\.(\d{2})$', s)
    if m:
        day, mon, yy = map(int, m.groups())
        year = 2000 + yy if yy <= 26 else 1900 + yy
        try:
            return dt.date(year, mon, day).isoformat()
        except ValueError:
            return ''
    return ''


def normalize_street(s: str) -> str:
    s = clean_text(s)
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\bstreet\b', 'st', s)
    s = re.sub(r'\bavenue\b', 'ave', s)
    s = re.sub(r'\bboulevard\b', 'blvd', s)
    s = re.sub(r'\bdrive\b', 'dr', s)
    s = re.sub(r'\broad\b', 'rd', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def parse_vehicle_address(s: str):
    parts = [p.strip() for p in (s or '').split(',')]
    street = normalize_street(parts[0]) if parts else ''
    city = clean_text(parts[1]) if len(parts) > 1 else ''
    state = clean_text(parts[2]) if len(parts) > 2 else ''
    return street, city, state


def sim(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    return SequenceMatcher(None, a, b).ratio()


def score_pair(a, b):
    score = 0.0

    name_score = 0.6 * sim(a['first_name'], b['first_name']) + 0.4 * sim(a['last_name'], b['last_name'])
    score += 0.35 * name_score

    if a['phone'] and b['phone']:
        if a['phone'] == b['phone']:
            score += 0.35
        elif a['phone'][-10:] == b['phone'][-10:]:
            score += 0.25

    if a['date'] and b['date']:
        if a['date'] == b['date']:
            score += 0.20
        elif a['date'][:7] == b['date'][:7]:
            score += 0.08

    city_sim = sim(a['city'], b['city'])
    if city_sim > 0.9:
        score += 0.05
    state_bonus = 0.03 if a['state'] and b['state'] and a['state'] == b['state'] else 0.0
    score += state_bonus

    addr_sim = sim(a['street'], b['street'])
    if addr_sim > 0.85:
        score += 0.07
    elif addr_sim > 0.7:
        score += 0.03

    return score


def build_records():
    people_rows = read_xlsx_rows('people_dataset_15000_challenges_with_overlap.xlsx', 'Data')
    vehicle_rows = read_xlsx_rows('fake_vehicle_people_dataset_15000_challenges.xlsx', 'Data')

    people_header = people_rows[0]
    vehicle_header = vehicle_rows[0]

    people = []
    for i, row in enumerate(people_rows[1:], start=1):
        rec = {people_header[j]: row[j] if j < len(row) else '' for j in range(len(people_header))}
        people.append({
            'a_idx': i,
            'first_name': alpha(rec.get('name', '')),
            'middle_name': alpha(rec.get('secondname', '')),
            'last_name': alpha(rec.get('surname', '')),
            'full_name': clean_text(rec.get('originalname', '')),
            'city': clean_text(rec.get('city', '')),
            'state': clean_text(rec.get('state', '')),
            'country': clean_text(rec.get('country', '')),
            'zipcode': clean_text(rec.get('zipcode', '')),
            'street': normalize_street(rec.get('address', '')),
            'date': parse_date_any(rec.get('date', '')),
            'gender': clean_text(rec.get('gender', '')),
            'phone': parse_phone(rec.get('phone #', '')),
            'email': clean_text(rec.get('email', '')),
            'raw': rec,
        })

    vehicles = []
    for i, row in enumerate(vehicle_rows[1:], start=1):
        rec = {vehicle_header[j]: row[j] if j < len(row) else '' for j in range(len(vehicle_header))}
        names = clean_text(rec.get('name', '')).split()
        first = alpha(names[0]) if names else ''
        last = alpha(names[-1]) if len(names) > 1 else ''
        street, city, state = parse_vehicle_address(rec.get('address', ''))
        vehicles.append({
            'b_idx': i,
            'id': rec.get('id', ''),
            'first_name': first,
            'last_name': last,
            'name_raw': clean_text(rec.get('name', '')),
            'city': city,
            'state': state,
            'street': street,
            'date': parse_date_any(rec.get('date', '')),
            'phone': parse_phone(rec.get('phone_number', '')),
            'passport': clean_text(rec.get('passport', '')),
            'plate_number': clean_text(rec.get('plate_number', '')),
            'car_model': clean_text(rec.get('car_model', '')),
            'car_color': clean_text(rec.get('car_color', '')),
            'car_year': clean_text(rec.get('car_year', '')),
            'car_vin': clean_text(rec.get('car_vin', '')),
            'raw': rec,
        })

    return people, vehicles


def resolve(people, vehicles):
    by_phone = defaultdict(list)
    by_last_first = defaultdict(list)
    by_last_city = defaultdict(list)
    for b in vehicles:
        if b['phone']:
            by_phone[b['phone']].append(b)
        by_last_first[(b['last_name'], b['first_name'][:1])].append(b)
        by_last_city[(b['last_name'], b['city'])].append(b)

    candidates = []
    for a in people:
        cand = {}

        if a['phone'] and a['phone'] in by_phone:
            for b in by_phone[a['phone']]:
                cand[b['b_idx']] = b

        key_lf = (a['last_name'], a['first_name'][:1])
        pool_lf = by_last_first.get(key_lf, [])
        if len(pool_lf) <= 120:
            for b in pool_lf:
                cand[b['b_idx']] = b

        key_lc = (a['last_name'], a['city'])
        pool_lc = by_last_city.get(key_lc, [])
        if len(pool_lc) <= 80:
            for b in pool_lc:
                cand[b['b_idx']] = b

        if not cand:
            continue

        for b in cand.values():
            s = score_pair(a, b)
            if s >= 0.60:
                candidates.append((s, a['a_idx'], b['b_idx']))

    candidates.sort(reverse=True)

    used_a = set()
    used_b = set()
    matches = []
    for s, a_idx, b_idx in candidates:
        if a_idx in used_a or b_idx in used_b:
            continue
        used_a.add(a_idx)
        used_b.add(b_idx)
        matches.append((a_idx, b_idx, round(s, 4)))

    return matches


def write_outputs(people, vehicles, matches):
    people_map = {p['a_idx']: p for p in people}
    vehicle_map = {v['b_idx']: v for v in vehicles}

    out_pairs = Path('matched_pairs.csv')
    with out_pairs.open('w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow([
            'a_idx', 'b_idx', 'match_score',
            'a_name', 'b_name', 'a_phone', 'b_phone', 'a_date', 'b_date',
            'a_address', 'b_address',
        ])
        for a_idx, b_idx, s in matches:
            a = people_map[a_idx]
            b = vehicle_map[b_idx]
            w.writerow([
                a_idx, b_idx, s,
                f"{a['raw'].get('name', '')} {a['raw'].get('surname', '')}".strip(),
                b['raw'].get('name', ''),
                a['raw'].get('phone #', ''),
                b['raw'].get('phone_number', ''),
                a['raw'].get('date', ''),
                b['raw'].get('date', ''),
                a['raw'].get('address', ''),
                b['raw'].get('address', ''),
            ])

    matched_a = {a for a, _, _ in matches}
    matched_b = {b for _, b, _ in matches}

    out_master = Path('master_people_dataset.csv')
    with out_master.open('w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow([
            'entity_id', 'source_a_idx', 'source_b_idx', 'match_score',
            'first_name', 'last_name', 'middle_name', 'full_name',
            'date_of_birth', 'gender', 'phone', 'email',
            'street', 'city', 'state', 'zipcode', 'country',
            'passport', 'vehicle_plate', 'vehicle_model', 'vehicle_color', 'vehicle_year', 'vehicle_vin',
            'source_coverage'
        ])

        entity_id = 1
        for a_idx, b_idx, s in matches:
            a = people_map[a_idx]
            b = vehicle_map[b_idx]
            w.writerow([
                entity_id, a_idx, b_idx, s,
                a['raw'].get('name', '') or b['raw'].get('name', '').split(' ')[0],
                a['raw'].get('surname', '') or b['raw'].get('name', '').split(' ')[-1],
                a['raw'].get('secondname', ''),
                a['raw'].get('originalname', '') or b['raw'].get('name', ''),
                a['date'] or b['date'],
                a['raw'].get('gender', ''),
                a['raw'].get('phone #', '') or b['raw'].get('phone_number', ''),
                a['raw'].get('email', ''),
                a['raw'].get('address', '') or b['raw'].get('address', ''),
                a['raw'].get('city', '') or b['city'],
                a['raw'].get('state', '') or b['state'],
                a['raw'].get('zipcode', ''),
                a['raw'].get('country', ''),
                b['raw'].get('passport', ''),
                b['raw'].get('plate_number', ''),
                b['raw'].get('car_model', ''),
                b['raw'].get('car_color', ''),
                b['raw'].get('car_year', ''),
                b['raw'].get('car_vin', ''),
                'A+B'
            ])
            entity_id += 1

        for p in people:
            if p['a_idx'] in matched_a:
                continue
            w.writerow([
                entity_id, p['a_idx'], '', '',
                p['raw'].get('name', ''), p['raw'].get('surname', ''), p['raw'].get('secondname', ''), p['raw'].get('originalname', ''),
                p['date'], p['raw'].get('gender', ''), p['raw'].get('phone #', ''), p['raw'].get('email', ''),
                p['raw'].get('address', ''), p['raw'].get('city', ''), p['raw'].get('state', ''), p['raw'].get('zipcode', ''), p['raw'].get('country', ''),
                '', '', '', '', '', '',
                'A'
            ])
            entity_id += 1

        for v in vehicles:
            if v['b_idx'] in matched_b:
                continue
            w.writerow([
                entity_id, '', v['b_idx'], '',
                (v['raw'].get('name', '').split(' ')[0] if v['raw'].get('name', '') else ''),
                (v['raw'].get('name', '').split(' ')[-1] if v['raw'].get('name', '') else ''),
                '', v['raw'].get('name', ''),
                v['date'], '', v['raw'].get('phone_number', ''), '',
                v['raw'].get('address', ''), v['city'], v['state'], '', '',
                v['raw'].get('passport', ''), v['raw'].get('plate_number', ''), v['raw'].get('car_model', ''),
                v['raw'].get('car_color', ''), v['raw'].get('car_year', ''), v['raw'].get('car_vin', ''),
                'B'
            ])
            entity_id += 1

    summary = Path('resolution_summary.txt')
    with summary.open('w', encoding='utf-8') as f:
        f.write('Entity Resolution Summary\n')
        f.write('=========================\n')
        f.write(f'Total dataset A records: {len(people)}\n')
        f.write(f'Total dataset B records: {len(vehicles)}\n')
        f.write(f'Matched cross-dataset entities: {len(matches)}\n')
        f.write(f'Unmatched dataset A records: {len(people) - len(matched_a)}\n')
        f.write(f'Unmatched dataset B records: {len(vehicles) - len(matched_b)}\n')
        f.write(f'Total master entities: {len(people) + len(vehicles) - len(matches)}\n')


def main():
    people, vehicles = build_records()
    matches = resolve(people, vehicles)
    write_outputs(people, vehicles, matches)
    print(f'people={len(people)} vehicles={len(vehicles)} matches={len(matches)}')


if __name__ == '__main__':
    main()
