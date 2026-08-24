"""
update.py
---------
폴더 내 모든 .xlsx 파일을 읽어 dashboard.html 데이터를 자동 갱신합니다.
- 기존 07.01-08.23 베이스 파일 + 새로 추가되는 일별 파일 모두 합산
- 새 파일을 폴더에 추가한 뒤 이 스크립트를 실행하면 됩니다
"""

import openpyxl, glob, re, json
from collections import defaultdict

DASHBOARD = 'dashboard.html'
INDEX     = 'index.html'

# ── 유틸 ────────────────────────────────────────────────────────

def n(v):
    try: return float(str(v).replace(',', ''))
    except: return 0

def to_date(v):
    if v is None: return ''
    if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
    s = str(v)
    return s[:10] if len(s) >= 10 else ''

def shorten(name, maxlen=22):
    return name[:maxlen].strip()

def assign_cat(name):
    if '와이드' in name:                             return ('리빙', '와이드')
    if '이불 압축' in name or '이불압축' in name:     return ('리빙', '이불')
    if '아우터 압축' in name or '아우터압축' in name:  return ('리빙', '아우터')
    if '행잉 오거나이저' in name or '행잉오거나이저' in name: return ('리빙', '오거나이저')
    if '세이프 데이즈' in name or '데이즈' in name:   return ('여행', '데이즈')
    if '세이프 플러스' in name or '플러스 라인' in name: return ('여행', '플러스')
    if '세이프' in name:                              return ('여행', '세이프')
    if '압축 파우치 라이트' in name:                   return ('여행', '라이트')
    if '여행 압축 파우치' in name:                     return ('여행', '여행')
    if '패커블' in name or '폴더블' in name:           return ('여행', '패커블')
    return ('여행', '기타')

# ── 1. 전체 xlsx 파일 읽기 ──────────────────────────────────────

xlsx_files = sorted(glob.glob('*.xlsx'))
if not xlsx_files:
    print('ERROR: .xlsx 파일이 없습니다'); exit(1)

print(f'파일 {len(xlsx_files)}개 발견:')

daily    = defaultdict(lambda: {'s': 0, 'o': 0, 'r': 0, 'v': 0})
prod_all = defaultdict(lambda: {'s': 0, 'o': 0})
prod_mon = defaultdict(lambda: defaultdict(lambda: {'s': 0, 'o': 0}))
prod_map = {}  # code -> {name, s}

for fname in xlsx_files:
    try:
        wb = openpyxl.load_workbook(fname, data_only=True)
        if 'SALES' not in wb.sheetnames:
            print(f'  SKIP {fname}  (SALES 시트 없음)'); wb.close(); continue
        ws = wb['SALES']
        cnt = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            if len(row) < 16: continue
            date = to_date(row[0])
            if not date or date == 'None': continue
            mm = date[5:7]
            s, o, r, v = n(row[7]), n(row[3]), n(row[4]), n(row[15])
            d = daily[date]
            d['s'] += s; d['o'] += o; d['r'] += r; d['v'] += v
            name = str(row[1]).strip() if row[1] else ''
            code = str(row[2]).strip() if row[2] else ''
            if name:
                prod_all[name]['s'] += s; prod_all[name]['o'] += o
                prod_mon[mm][name]['s'] += s; prod_mon[mm][name]['o'] += o
            if code and code != 'None' and name:
                if code not in prod_map:
                    prod_map[code] = {'name': name, 's': 0}
                prod_map[code]['s'] += s
            cnt += 1
        print(f'  OK  {fname}  ({cnt:,}행)')
        wb.close()
    except Exception as e:
        print(f'  ERR {fname}: {e}')

if not daily:
    print('ERROR: 유효한 데이터가 없습니다'); exit(1)

months     = sorted(set(d[5:7] for d in daily.keys()))
start_date = min(daily.keys())
end_date   = max(daily.keys())
print(f'\n기간: {start_date} ~ {end_date}  |  {len(daily)}일  |  월: {", ".join(months)}')

# ── 2. JS 블록 생성 ─────────────────────────────────────────────

# ALL: 일별 데이터
rows = []
for d in sorted(daily.keys()):
    lb = f'{d[5:7]}/{d[8:10]}'
    v  = daily[d]
    rows.append(f'  ["{lb}",{int(v["s"])},{int(v["o"])},{int(v["r"])},{int(v["v"])}]')
all_js = 'const ALL = [ // [label, sales, orders, refunds, visits]\n' + ',\n'.join(rows) + '\n];'

# PEAKS: 일평균의 2배 초과 날짜 자동 감지
avg_s = sum(v['s'] for v in daily.values()) / len(daily)
peaks = {f'{d[5:7]}/{d[8:10]}': 1 for d, v in daily.items() if v['s'] > avg_s * 2}
peaks_js = 'const PEAKS = ' + json.dumps(peaks, ensure_ascii=False, separators=(',', ':')) + ';'

# TOP_S / TOP_O
def fmt_list(pairs):
    return '[' + ','.join(f'["{shorten(nm)}",{int(val)}]' for nm, val in pairs) + ']'

def top10(dct, key):
    return [(nm, v[key]) for nm, v in sorted(dct.items(), key=lambda x: x[1][key], reverse=True)[:10]]

s_parts = ['  all:' + fmt_list(top10(prod_all, 's'))]
o_parts = ['  all:' + fmt_list(top10(prod_all, 'o'))]
for mm in months:
    s_parts.append(f"  '{mm}':" + fmt_list(top10(prod_mon[mm], 's')))
    o_parts.append(f"  '{mm}':" + fmt_list(top10(prod_mon[mm], 'o')))
top_s_js = 'const TOP_S = {\n' + ',\n'.join(s_parts) + '\n};'
top_o_js = 'const TOP_O = {\n' + ',\n'.join(o_parts) + '\n};'

# PROD_LIST
pl_rows = []
for code, v in sorted(prod_map.items(), key=lambda x: x[1]['s'], reverse=True):
    nm      = v['name'].replace('\\', '\\\\').replace('"', '\\"')
    maj, mn = assign_cat(v['name'])
    pl_rows.append(f'  ["{code}","{nm}","{maj}","{mn}"]')
prod_list_js = 'const PROD_LIST=[\n' + ',\n'.join(pl_rows) + '\n];'

# ── 3. dashboard.html 패치 ──────────────────────────────────────

def replace_block(html, start, end, new_js):
    si = html.find(start)
    if si == -1:
        raise ValueError(f'시작 마커를 찾을 수 없음: {repr(start[:50])}')
    ei = html.find(end, si)
    if ei == -1:
        raise ValueError(f'종료 마커를 찾을 수 없음 (시작: {repr(start[:50])})')
    return html[:si] + new_js + html[ei + len(end):]

try:
    with open(DASHBOARD, 'r', encoding='utf-8') as f:
        html = f.read()
except FileNotFoundError:
    print(f'ERROR: {DASHBOARD} 파일이 없습니다. 같은 폴더에서 실행해주세요.'); exit(1)

# 신규 상품 감지: 기존 PROD_LIST 코드 추출
old_codes = set(re.findall(r'\["(\d+)"', html[html.find('const PROD_LIST=['):html.find('\n];', html.find('const PROD_LIST=['))]))
new_codes  = [code for code, _ in sorted(prod_map.items(), key=lambda x: x[1]['s'], reverse=True)]
added      = [c for c in new_codes if c not in old_codes]

# NEW_PRODS JS 생성
new_prods_js = 'const NEW_PRODS=' + json.dumps(added, ensure_ascii=False) + ';'

html = replace_block(html, 'const ALL = [',     '\n];', all_js)
html = replace_block(html, 'const TOP_S = {',   '\n};', top_s_js)
html = replace_block(html, 'const TOP_O = {',   '\n};', top_o_js)
html = replace_block(html, 'const PEAKS = ',    ';',    peaks_js)
html = replace_block(html, 'const NEW_PRODS=',  ';',    new_prods_js)
html = replace_block(html, 'const PROD_LIST=[\n', '\n];', prod_list_js)

# 타이틀 날짜 범위 업데이트 (예: 2026.07–08)
yr, m1, m2 = start_date[:4], start_date[5:7], end_date[5:7]
new_range = f'{yr}.{m1}–{m2}' if m1 != m2 else f'{yr}.{m1}'
html = re.sub(r'(매출 추이 대시보드 · )[\d.–\-]+', r'\g<1>' + new_range, html)

# PROD_LIST 주석 개수 업데이트
html = re.sub(
    r'// PRODUCT LIST DATA \(\d+개 상품\)',
    f'// PRODUCT LIST DATA ({len(pl_rows)}개 상품)',
    html
)

with open(DASHBOARD, 'w', encoding='utf-8') as f:
    f.write(html)
with open(INDEX, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'\n✓ {DASHBOARD} + {INDEX} 업데이트 완료')
print(f'  일별: {len(rows)}일  |  피크: {len(peaks)}일  |  상품: {len(pl_rows)}개')
if added:
    print(f'\n  ★ 신규 상품 {len(added)}개 — 대시보드에서 카테고리 지정 필요:')
    for c in added:
        print(f'    - [{c}] {prod_map[c]["name"]}')
else:
    print(f'  신규 상품 없음')
