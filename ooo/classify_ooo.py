"""Parse the T114 "Out of Order by Reason" PDF (text layer via pdftotext -layout)
and classify each record by its Remarks text.

Usage: python3 classify_ooo.py <report.pdf> [--json out.json]
Output: JSON summary (report date, unique rooms, rows, category counts, sub-category counts).
"""
import json, re, subprocess, sys, collections

ROW = re.compile(
    r'^(?P<code>[A-Z]{2,6})\s+(?P<desc>\S.*?)\s{2,}(?P<room>\d{4})\s+(?P<rtype>[A-Z]+)\s+'
    r'(?P<status>\S+)\s+(?P<ret>\S+)\s+(?P<frm>\d{2}-\d{2}-\d{2})\s+(?P<to>\d{2}-\d{2}-\d{2})')
DATE_HDR = re.compile(r'(\d{2}-\d{2}-\d{2})')

# ---- Classification ----
# 1) internal use and routine cleaning are checked first
INTERNAL = [('內部使用', ['倉庫', 'storage', '消防演練'])]
CLEANING = [('清潔保養', ['公清'])]

# 2) maintenance sub-categories (first match wins; a remark that matches any of these
#    and no status keyword is treated as 維修/工程)
MAINT_SUB = [
    ('霉味/異味', ['霉', '沼氣', '異味', '味重', '菸味', '煙味', '除味', '除臭']),
    ('蟲害', ['蟲']),
    ('玻璃門窗', ['玻璃', '落地窗', '氣窗', '窗戶', '窗台', '紗簾']),
    ('大理石', ['大理石']),
    ('木地板/木作裝修', ['木地板', '木作', '衣櫃', '衣櫥', '櫃門', '壁紙']),
    ('衛浴設備', ['浴缸', '馬桶', '洗手台', '淋浴', '補漆', '滲水', '漏水']),
    ('冷氣濾網更換(例行)', ['濾網']),
    ('冷氣/通風', ['冷氣', '冷風機', '抽風機', '出風口']),
    ('家具/五金', ['圓桌', '椅', '保險箱', '按鈕', '鐵片', '家具']),
    ('水電/燈具', ['燈', '線路', '插座']),
]
MAINT_GENERIC = ['報修', '未修', '仍有', '更換', '破損', '噪音', '不穩', '凹陷', '裂', '翹起', '故障', '未裝回', '維修', '待料']

# 3) status categories (used when the remark is not a maintenance note)
STATUS = [
    ('活動佈置', None, ['搬傢俱', '搬家具', '傢俱房', '傢俱恢復', '場地恢復', '房間恢復', '護膚團']),
    ('晚退', None, ['晚退', '晚褪']),
    ('預關/預進', 'VIP/訂房保留', ['董事長', 'CEO', 'PTY', '貴賓', '助總', '副總', '皇家', 'CRS', 'RV#', 'RQ Early', '訂房', 'booking', '張董']),
    ('預關/預進', '團體預進', ['pre block', 'pre-block', 'preblock', '大團預進', '團體預關']),
    ('預關/預進', '一般預關', ['預關', 'preblocked']),
]
STATUS_KW = [k for _, _, kws in STATUS for k in kws]

def _hit(text, kws):
    t = text.lower()
    return any(k.lower() in t for k in kws)

NEGATED = [r'未發霉', r'無霉', r'未有霉']

def _items(remark):
    """Split a multi-item remark ("1.xxx 2.yyy") into its numbered items."""
    parts = [p.strip() for p in re.split(r'(?:^|\s)(?=\d\.)', remark) if p.strip()]
    return parts or [remark]

def classify(remark: str):
    for pat in NEGATED:
        remark = re.sub(pat, '', remark)
    if _hit(remark, [k for _, kws in INTERNAL for k in kws]):
        return '內部使用', None
    if _hit(remark, [k for _, kws in CLEANING for k in kws]):
        return '清潔保養', None
    has_status = _hit(remark, STATUS_KW)
    # the room's maintenance sub-category comes from its first item that matches
    maint_sub = None
    for item in _items(remark):
        maint_sub = next((sub for sub, kws in MAINT_SUB if _hit(item, kws)), None)
        if maint_sub:
            break
    if maint_sub and not has_status:
        return '維修/工程', maint_sub
    if _hit(remark, MAINT_GENERIC) and not has_status:
        return '維修/工程', '其他維修'
    for cat, sub, kws in STATUS:
        if _hit(remark, kws):
            return cat, sub
    if maint_sub:
        return '維修/工程', maint_sub
    return '未分類', None

def parse(pdf_path):
    txt = subprocess.run(['pdftotext', '-layout', pdf_path, '-'],
                         capture_output=True, text=True, check=True).stdout
    lines = txt.splitlines()
    report_date = None
    for ln in lines[:5]:
        m = DATE_HDR.search(ln)
        if m:
            report_date = m.group(1); break
    recs = []
    for ln in lines:
        s = ln.rstrip()
        m = ROW.match(s.strip())
        if m:
            tail = s[s.find(m.group('to')) + len(m.group('to')):].strip()
            recs.append({'room': m.group('room'), 'rtype': m.group('rtype'),
                         'ret': m.group('ret'), 'from': m.group('frm'), 'to': m.group('to'),
                         'remark': tail})
        elif recs and s.strip() and not re.match(r'^\s*(Filter|Page|Room Class|Reason All|Room Status)', s.strip()) \
                and s.startswith(' ' * 20) and not re.search(r'\d{4}\s+[A-Z]{3,5}\s+OO', s):
            # continuation line of a multi-line remark (numbered items)
            if recs[-1]['remark']:
                recs[-1]['remark'] += ' ' + s.strip()
            else:
                recs[-1]['remark'] = s.strip()
    return report_date, recs

def _mdy(s):
    m, dd, y = s.split('-')
    return (2000 + int(y), int(m), int(dd))

def is_active(rec, report_date):
    """A record is active on the report date when From <= date <= To.
    Future pre-blocks (From after the report date) are not active."""
    if not report_date:
        return True
    d = _mdy(report_date)
    return _mdy(rec['from']) <= d <= _mdy(rec['to'])

def summarize(pdf_path):
    report_date, recs = parse(pdf_path)
    cat_c = collections.Counter(); sub_c = collections.Counter()
    room_cat = {}
    for r in recs:
        cat, sub = classify(r['remark'])
        r['category'], r['sub'] = cat, sub
        cat_c[cat] += 1
        if sub: sub_c[f'{cat}｜{sub}'] += 1
        room_cat.setdefault(r['room'], set()).add(cat)
    # unique rooms per category: a room counts in every category it has a record for
    rooms_by_cat = collections.Counter(c for cats in room_cat.values() for c in cats)
    active = [r for r in recs if is_active(r, report_date)]
    for r in recs:
        r['active'] = is_active(r, report_date)
    active_cat = collections.defaultdict(set)
    active_sub = collections.defaultdict(set)
    for r in active:
        active_cat[r['category']].add(r['room'])
        if r['sub']:
            key = ('維修/工程｜' if r['category'] == '維修/工程' else r['category'] + '｜') + r['sub']
            active_sub[key].add(r['room'])
    return {
        'report_date': report_date,
        'rows': len(recs),
        'unique_rooms': len({r['room'] for r in recs}),
        'active_rows': len(active),
        'active_unique_rooms': len({r['room'] for r in active}),
        'active_category_unique_rooms': {k: len(v) for k, v in active_cat.items()},
        'active_sub_unique_rooms': {k: len(v) for k, v in active_sub.items()},
        'category_rows': dict(cat_c.most_common()),
        'category_unique_rooms': dict(rooms_by_cat.most_common()),
        'sub_rows': dict(sub_c.most_common()),
        'records': recs,
    }

if __name__ == '__main__':
    out = summarize(sys.argv[1])
    if '--json' in sys.argv:
        with open(sys.argv[sys.argv.index('--json') + 1], 'w', encoding='utf-8') as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
    print(json.dumps({k: v for k, v in out.items() if k != 'records'}, ensure_ascii=False, indent=1))
