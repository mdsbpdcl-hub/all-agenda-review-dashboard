"""Download every source for the review with a Google service account (GitHub Actions).
Needs env GOOGLE_SA_KEY (the service account's JSON key). Every sheet/file/folder in sources.json must be shared
with the service account's e-mail as Viewer. A source that cannot be read falls back to yesterday's copy in fallback/."""
import json, os, io, sys, shutil, re, urllib.request, datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
B = os.path.dirname(os.path.abspath(__file__))
S = json.load(open(os.path.join(B, 'sources.json')))
for d in ('data', 'ext', 'pib', 'news_raw', 'fallback'): os.makedirs(os.path.join(B, d), exist_ok=True)
cred = service_account.Credentials.from_service_account_info(json.loads(os.environ['GOOGLE_SA_KEY']),
        scopes=['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly'])
drive = build('drive', 'v3', credentials=cred, cache_discovery=False)
sheets = build('sheets', 'v4', credentials=cred, cache_discovery=False)
report = {'fresh': [], 'fallback': [], 'failed': []}
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
def save(req, path):
    buf = io.FileIO(path, 'wb'); dl = MediaIoBaseDownload(buf, req, chunksize=8 * 1024 * 1024); done = False
    while not done: _, done = dl.next_chunk()
    buf.close()
def fallback(name, dest):
    src = os.path.join(B, 'fallback', name)
    if os.path.exists(src): shutil.copy(src, dest); report['fallback'].append(name); return True
    report['failed'].append(name); return False
# 1. sheets -> xlsx (export), else yesterday's JSON copy
for name, fid in S['sheets_to_xlsx'].items():
    path = os.path.join(B, 'data', name)
    for attempt in (1, 2):
        try:
            save(drive.files().export_media(fileId=fid, mimeType=XLSX), path); report['fresh'].append(name); break
        except Exception as e:
            print('export failed', name, attempt, str(e)[:160])
            if os.path.exists(path): os.remove(path)
    else:
        fallback(name.replace('.xlsx', '.json'), os.path.join(B, 'data', name.replace('.xlsx', '.json')))
# 2. DM feedback rows
try:
    fb = S['dm_feedback_sheet']
    v = sheets.spreadsheets().values().get(spreadsheetId=fb['id'], range=fb['range']).execute()
    json.dump(v, open(os.path.join(B, fb['save_as']), 'w'), ensure_ascii=False); report['fresh'].append('dm_feedback')
except Exception as e: print('dm feedback failed', e); report['failed'].append('dm_feedback')
# 3. pending issues JSON
pj = S['pending_issues_json']
try: save(drive.files().get_media(fileId=pj['id']), os.path.join(B, pj['save_as'])); report['fresh'].append('pending issues')
except Exception as e: print('pending issues failed', e); fallback('summary.json', os.path.join(B, pj['save_as']))
# 4. NFMS / RDSS
for key, out in (('nfms_latest', 'nfms_latest.json'), ('rdss_latest', 'rdss_latest.json')):
    try: urllib.request.urlretrieve(S[key], os.path.join(B, 'ext', out)); report['fresh'].append(out)
    except Exception as e: print(key, 'failed', e); fallback(out, os.path.join(B, 'ext', out))
# 5. news: newest prepared manifest and its parts
news = {'date': None}
try:
    q = "'%s' in parents and name contains 'manifest' and trashed = false" % S['news_prepared_folder']
    files = drive.files().list(q=q, orderBy='createdTime desc', pageSize=5, fields='files(id,name,createdTime)').execute().get('files', [])
    if files:
        man = files[0]; m = re.match(r'prep_(.+)_(\d+)\.manifest\.json$', man['name']); paper = m.group(1)
        save(drive.files().get_media(fileId=man['id']), os.path.join(B, 'news_raw', man['name']))
        parts = json.load(open(os.path.join(B, 'news_raw', man['name'])))['parts']
        for p in parts:
            f = drive.files().list(q="'%s' in parents and name = '%s'" % (S['news_prepared_folder'], p), fields='files(id)').execute()['files'][0]
            save(drive.files().get_media(fileId=f['id']), os.path.join(B, 'news_raw', p))
        d = None
        try:
            title = drive.files().get(fileId=paper, fields='name').execute()['name']
            d = re.search(r'(\d{1,2})[ ._-](\d{1,2})[ ._-](20\d\d)', title)
        except Exception as e: print('paper title not readable, using the clippings date', str(e)[:120])
        if d: news['date'] = datetime.date(int(d.group(3)), int(d.group(2)), int(d.group(1))).strftime('%d %b %Y')
        else:
            made = drive.files().get(fileId=man['id'], fields='createdTime').execute()['createdTime'][:10]
            news['date'] = datetime.date.fromisoformat(made).strftime('%d %b %Y')
        news['paper'] = paper; report['fresh'].append('news ' + news['date'])
except Exception as e: print('news failed', e); report['failed'].append('news')
json.dump(news, open(os.path.join(B, 'news_raw', 'paper.json'), 'w'))
json.dump(report, open(os.path.join(B, 'fetch_report.json'), 'w'), indent=1)
print(json.dumps(report, indent=1))
