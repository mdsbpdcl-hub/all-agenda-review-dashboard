"""Save the sheets the build reads from an .xlsx as JSON (a list of rows per sheet), so it can be kept in the artifact.
Usage: python3 -I xlsx2json.py <in.xlsx> <out.json> [sheet ...]   (no sheet names = all sheets)"""
import json, sys, datetime, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], read_only=True, data_only=True)
want = sys.argv[3:]
def conv(v):
    if isinstance(v, (datetime.datetime, datetime.date)): return {'__dt': (v if isinstance(v, datetime.datetime) else datetime.datetime(v.year, v.month, v.day)).isoformat()}
    if isinstance(v, (datetime.time, datetime.timedelta)): return str(v)
    return v
out = {}
for ws in wb.worksheets:
    if want and ws.title not in want: continue
    rows = [[conv(c) for c in r] for r in ws.iter_rows(values_only=True)]
    while rows and not any(c not in (None, '') for c in rows[-1]): rows.pop()
    out[ws.title] = rows
json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, separators=(',', ':'))
# revenue: keep only the latest period (all the build reads)
if sys.argv[2].endswith('revenue.json'):
    for t, rows in out.items():
        if rows and 'Period' in rows[0]:
            i = rows[0].index('Period'); last = max(str(r[i]) for r in rows[1:] if r[i])
            out[t] = [rows[0]] + [r for r in rows[1:] if str(r[i]) == last]
    json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, separators=(',', ':'))
