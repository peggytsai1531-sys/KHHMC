#!/usr/bin/env python3
"""Archive one day's T114 OOO result into ooo/history/<YYYY-MM>/ and rebuild the month's files.

Usage: python3 -I ooo/build_history.py <ooo_result.json> [--root <repo ooo/history dir>]
Input : JSON written by classify_ooo.py (report_date, records[...]).
Output: <date>_rooms.csv, <date>.json, <date>_diff.csv, summary.csv, OOO_<YYYY-MM>.xlsx
"""
import csv, json, os, re, sys, glob, datetime as dt
from collections import OrderedDict

CATS = ['維修/工程', '預關/預進', '內部使用', '清潔保養', '晚退', '活動佈置', '未分類']
ROOM_COLS = ['date', 'room', 'room_type', 'category', 'sub', 'period', 'return_status', 'remark', '連續天數']
DIFF_COLS = ['date', 'room', '狀態', 'category', '連續天數', '前一日基準日']

def iso(s):  # MM-DD-YY -> YYYY-MM-DD
    m = re.match(r'(\d\d)-(\d\d)-(\d\d)$', s)
    return f"20{m.group(3)}-{m.group(1)}-{m.group(2)}"

def rooms_from_result(res):
    by = OrderedDict()
    for r in res['records']:
        by.setdefault(r['room'], []).append(r)
    out = []
    for room, rs in by.items():
        act = [r for r in rs if r['active']]
        use = act or rs
        cats = list(OrderedDict.fromkeys(r['category'] for r in use))
        subs = list(OrderedDict.fromkeys(r['sub'] for r in use if r.get('sub')))
        frm = min(iso(r['from']) for r in use); to = max(iso(r['to']) for r in use)
        remark = ' | '.join(OrderedDict.fromkeys(' '.join(r['remark'].split()) for r in use))
        out.append(dict(room=room, room_type=use[0]['rtype'], active=bool(act),
                        category='／'.join(cats), sub='／'.join(subs),
                        period=f"{frm}~{to}", return_status=use[0]['ret'], remark=remark))
    return sorted(out, key=lambda x: x['room'])

def read_csv(p):
    with open(p, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))

def write_csv(p, cols, rows):
    with open(p, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction='ignore'); w.writeheader(); w.writerows(rows)

def rebuild_month(root, month):
    mdir = os.path.join(root, month)
    days = sorted(os.path.basename(p)[:10] for p in glob.glob(os.path.join(mdir, '*_rooms.csv')))
    summ = []
    for d in days:
        rooms = read_csv(os.path.join(mdir, f'{d}_rooms.csv'))
        meta = json.load(open(os.path.join(mdir, f'{d}.json'), encoding='utf-8'))
        dp = os.path.join(mdir, f'{d}_diff.csv')
        diff = read_csv(dp) if os.path.exists(dp) else None
        row = OrderedDict(date=d, OOO總計=len(rooms), 未來預排=meta['future_only_rooms'])
        for k in ('新增', '解除', '持續'):
            row[k] = sum(1 for x in diff if x['狀態'] == k) if diff is not None else ''
        for c in CATS:
            row[c] = sum(1 for x in rooms if c in x['category'].split('／'))
        row['前一日基準日'] = (diff[0]['前一日基準日'] if diff else '') if diff is not None else ''
        summ.append(row)
    cols = list(summ[0].keys()) if summ else []
    if summ: write_csv(os.path.join(mdir, 'summary.csv'), cols, summ)
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print('openpyxl missing: xlsx skipped'); return
    wb = Workbook(); ws = wb.active; ws.title = '每日摘要'
    def fill(ws, cols, rows):
        ws.append(cols)
        for r in rows: ws.append([r.get(c, '') for c in cols])
        for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='5B2C83')
        ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
        for i, c in enumerate(cols, 1):
            w = max(len(str(c)) * 2, *(min(len(str(r.get(c, ''))), 40) for r in rows)) + 2 if rows else 12
            ws.column_dimensions[get_column_letter(i)].width = min(max(w, 8), 60)
    fill(ws, cols, summ)
    alld = []
    for d in reversed(days):
        dp = os.path.join(mdir, f'{d}_diff.csv')
        if os.path.exists(dp): alld += read_csv(dp)
    fill(wb.create_sheet('前日差異'), DIFF_COLS, alld)
    for d in days:
        fill(wb.create_sheet(d[5:]), ROOM_COLS, read_csv(os.path.join(mdir, f'{d}_rooms.csv')))
    wb.save(os.path.join(mdir, f'OOO_{month}.xlsx'))

def main():
    res = json.load(open(sys.argv[1], encoding='utf-8'))
    root = sys.argv[sys.argv.index('--root') + 1] if '--root' in sys.argv else os.path.join(os.path.dirname(__file__), 'history')
    d = iso(res['report_date']); month = d[:7]
    mdir = os.path.join(root, month); os.makedirs(mdir, exist_ok=True)
    allr = rooms_from_result(res)
    act = [r for r in allr if r['active']]
    # previous report: latest *_rooms.csv strictly before this date, any month folder
    prevs = sorted(p for p in glob.glob(os.path.join(root, '*', '*_rooms.csv')) if os.path.basename(p)[:10] < d)
    prev_date, prev = None, {}
    if prevs:
        prev_date = os.path.basename(prevs[-1])[:10]
        prev = {x['room']: x for x in read_csv(prevs[-1])}
    gap = (dt.date.fromisoformat(d) - dt.date.fromisoformat(prev_date)).days if prev_date else 0
    cur = {r['room']: r for r in act}
    rows = []
    for r in act:
        p = prev.get(r['room'])
        days = int(p['連續天數'] or 1) + gap if p else 1
        rows.append({**r, 'date': d, '連續天數': days})
    write_csv(os.path.join(mdir, f'{d}_rooms.csv'), ROOM_COLS, rows)
    json.dump(dict(report_date=d, unique_rooms=len(allr), active_unique_rooms=len(act),
                   future_only_rooms=len(allr) - len(act), rooms=[r if r['active'] else {k: v for k, v in r.items() if k != 'remark'} for r in allr]),
              open(os.path.join(mdir, f'{d}.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    dpath = os.path.join(mdir, f'{d}_diff.csv')
    if os.path.exists(dpath): os.remove(dpath)
    if prev_date:
        diff = []
        for x in rows:
            diff.append(dict(date=d, room=x['room'], 狀態='持續' if x['room'] in prev else '新增', category=x['category'],
                             連續天數=x['連續天數'], 前一日基準日=prev_date))
        for room, p in prev.items():
            if room not in cur:
                diff.append(dict(date=d, room=room, 狀態='解除', category=p['category'], 連續天數=0, 前一日基準日=prev_date))
        order = {'新增': 0, '解除': 1, '持續': 2}
        diff.sort(key=lambda x: (order[x['狀態']], x['room']))
        write_csv(dpath, DIFF_COLS, diff)
        n = {k: sum(1 for x in diff if x['狀態'] == k) for k in order}
        print(f'{d}: OOO={len(act)} future={len(allr)-len(act)} vs {prev_date}: 新增{n["新增"]} 解除{n["解除"]} 持續{n["持續"]}')
    else:
        print(f'{d}: OOO={len(act)} future={len(allr)-len(act)} (首日基準，無前日可比)')
    rebuild_month(root, month)

if __name__ == '__main__':
    main()
