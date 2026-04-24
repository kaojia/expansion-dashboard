import os

def fnum(v):
    try:
        return float(v)
    except Exception:
        return 0

def parse_csv(filepath):
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        raw = f.read()
    data = {}
    lines = [l.strip() for l in raw.replace('\r\n', '\n').split('\n')]
    for i, line in enumerate(lines):
        if line.startswith(',Retail Net OPS'):
            data['sw'] = lines[i+1].split(',')
            data['pw'] = lines[i+2].split(',')
            data['py'] = lines[i+3].split(',')
            break
    daily = {}
    for i, line in enumerate(lines):
        if line.startswith('Period,Merchant Type,Metric Type,'):
            for j in range(i+1, len(lines)):
                r = lines[j]
                if not r or r.startswith('Product Group'):
                    break
                parts = r.split(',')
                if len(parts) >= 10:
                    daily[parts[0]+'|'+parts[1]+'|'+parts[2]] = parts[3:]
            break
    data['daily'] = daily
    pg_ops, pg_units = {}, {}
    metric = None
    for line in lines:
        if line.startswith('Product Group breakdown,NET_OPS'):
            metric = 'ops'
            continue
        if line.startswith('Product Group breakdown,NET_UNITS'):
            metric = 'units'
            continue
        if line.startswith('Top ASINs'):
            metric = None
            continue
        if metric and line:
            parts = line.split(',', 1)
            if len(parts) == 2 and parts[0]:
                try:
                    if metric == 'ops':
                        pg_ops[parts[0]] = float(parts[1])
                    else:
                        pg_units[parts[0]] = int(float(parts[1]))
                except ValueError:
                    pass
    data['pg_ops'] = pg_ops
    data['pg_units'] = pg_units
    top_asins = []
    in_asins = False
    for line in lines:
        if line.startswith('Top ASINs,NET_OPS'):
            in_asins = True
            continue
        if line.startswith('Top ASINs,NET_UNITS'):
            in_asins = False
            continue
        if in_asins and line and line[0] == 'B':
            parts = line.split(',', 2)
            if len(parts) >= 2:
                try:
                    top_asins.append({'asin': parts[0], 'value': float(parts[1]), 'title': parts[2] if len(parts) > 2 else ''})
                except ValueError:
                    pass
    data['top_asins'] = top_asins
    return data

w13 = parse_csv('W13/2026_AU30_w13.csv')
w12 = parse_csv('W13/2026_AU30_w12.csv')
ly = parse_csv('W13/2025_AU30_w13.csv')
sw, pw, py_ = w13['sw'], w13['pw'], w13['py']
dd = w13['daily']
asins = sorted(w13.get('top_asins', []), key=lambda x: x['value'], reverse=True)

def pct(a, b):
    return (a - b) / b if b else 0
def fm(v):
    return '${:,.0f}'.format(v)
def fi(v):
    return '{:,.0f}'.format(v)
def fp(v):
    return '{:+.1f}%'.format(v * 100)
def pc(v):
    return 'pos' if v >= 0 else 'neg'
def ar(v):
    return '&#9650;' if v >= 0 else '&#9660;'
def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

# Write HTML directly to file to avoid triple-quote issues
f = open('W13_WBR_AU30_Pipeline.html', 'w', encoding='utf-8')

f.write('<!DOCTYPE html>\n<html lang="en">\n<head>\n')
f.write('<meta charset="UTF-8">\n')
f.write('<meta name="viewport" content="width=device-width, initial-scale=1.0">\n')
f.write('<title>WBR - AU30 Pipeline W13 2026</title>\n')
f.write('<style>\n')
f.write(':root{--blue:#4472C4;--bl:#D6E4F0;--g:#006100;--gb:#C6EFCE;--r:#9C0006;--rb:#FFC7CE;--gy:#F2F2F2;--dk:#1a1a2e;--bd:#dee2e6;}\n')
f.write('*{margin:0;padding:0;box-sizing:border-box;}\n')
f.write('body{font-family:"Segoe UI",system-ui,sans-serif;background:#f0f2f5;color:#333;line-height:1.5;}\n')
f.write('.ctn{max-width:1400px;margin:0 auto;padding:20px;}\n')
f.write('.hdr{background:linear-gradient(135deg,var(--dk),var(--blue));color:#fff;padding:30px 40px;border-radius:12px;margin-bottom:24px;}\n')
f.write('.hdr h1{font-size:24px;font-weight:700;margin-bottom:4px;}\n')
f.write('.hdr p{opacity:.85;font-size:14px;}\n')
f.write('.tabs{display:flex;gap:4px;margin-bottom:20px;flex-wrap:wrap;}\n')
f.write('.tab{padding:10px 20px;background:#fff;border:1px solid var(--bd);border-radius:8px 8px 0 0;cursor:pointer;font-size:13px;font-weight:600;color:#666;transition:all .2s;}\n')
f.write('.tab:hover{color:var(--blue);}\n')
f.write('.tab.active{background:var(--blue);color:#fff;border-color:var(--blue);}\n')
f.write('.pnl{display:none;}.pnl.active{display:block;}\n')
f.write('.card{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:24px;margin-bottom:20px;}\n')
f.write('.card h2{font-size:16px;color:var(--blue);margin-bottom:16px;border-bottom:2px solid var(--bl);padding-bottom:8px;}\n')
f.write('.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;margin-bottom:24px;}\n')
f.write('.kpi{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);border-left:4px solid var(--blue);}\n')
f.write('.kpi .lb{font-size:12px;color:#888;text-transform:uppercase;letter-spacing:.5px;}\n')
f.write('.kpi .vl{font-size:28px;font-weight:700;margin:4px 0;}\n')
f.write('.kpi .ch{font-size:13px;font-weight:600;}\n')
f.write('.pos{color:var(--g);}.neg{color:var(--r);}\n')
f.write('.pos-bg{background:var(--gb);color:var(--g);}\n')
f.write('.neg-bg{background:var(--rb);color:var(--r);}\n')
f.write('table{width:100%;border-collapse:collapse;font-size:13px;}\n')
f.write('th{background:var(--blue);color:#fff;padding:10px 12px;text-align:center;font-weight:600;white-space:nowrap;}\n')
f.write('td{padding:8px 12px;border-bottom:1px solid #eee;}\n')
f.write('tr:nth-child(even){background:var(--gy);}\n')
f.write('.tr{text-align:right;}.tc{text-align:center;}.tl{text-align:left;}\n')
f.write('.bd{display:inline-block;padding:2px 8px;border-radius:4px;font-weight:600;font-size:12px;}\n')
f.write('.at{max-width:400px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px;}\n')
f.write('@media(max-width:768px){.kg{grid-template-columns:1fr;}table{font-size:11px;}th,td{padding:6px 8px;}.hdr{padding:20px;}.hdr h1{font-size:18px;}}\n')
f.write('</style>\n</head>\n<body>\n<div class="ctn">\n')

# Header
f.write('<div class="hdr">\n')
f.write('  <h1>WBR &mdash; AU30 Pipeline Weekly Business Review</h1>\n')
f.write('  <p>Week 13: 2026-03-22 ~ 2026-03-28 &nbsp;|&nbsp; amazon.com.au &nbsp;|&nbsp; USD</p>\n')
f.write('</div>\n')

# KPI cards
so = fnum(sw[5]); po = fnum(pw[5]); yo = fnum(py_[5])
su = fnum(sw[6]); pu = fnum(pw[6]); yu = fnum(py_[6])
wo = pct(so, po); wu = pct(su, pu); yoo = pct(so, yo); yu2 = pct(su, yu)
tpg = list(w13['pg_ops'].keys())[0] if w13['pg_ops'] else 'N/A'
tpv = list(w13['pg_ops'].values())[0] if w13['pg_ops'] else 0
ta = asins[0]['asin'] if asins else 'N/A'
tv = asins[0]['value'] if asins else 0

f.write('<div class="kg">\n')
# KPI 1: Net OPS
f.write('<div class="kpi"><div class="lb">Net OPS (USD)</div>')
f.write('<div class="vl">' + fm(so) + '</div>')
f.write('<div class="ch ' + pc(wo) + '">WoW: ' + ar(wo) + ' ' + fp(wo) + '</div>')
f.write('<div class="ch ' + pc(yoo) + '">YoY: ' + ar(yoo) + ' ' + fp(yoo) + '</div>')
f.write('</div>\n')
# KPI 2: Net Units
f.write('<div class="kpi"><div class="lb">Net Units</div>')
f.write('<div class="vl">' + fi(su) + '</div>')
f.write('<div class="ch ' + pc(wu) + '">WoW: ' + ar(wu) + ' ' + fp(wu) + '</div>')
f.write('<div class="ch ' + pc(yu2) + '">YoY: ' + ar(yu2) + ' ' + fp(yu2) + '</div>')
f.write('</div>\n')
# KPI 3: Top PG
f.write('<div class="kpi"><div class="lb">Top Product Group (OPS)</div>')
f.write('<div class="vl" style="font-size:20px">' + tpg + '</div>')
f.write('<div class="ch">' + fm(tpv) + '</div>')
f.write('</div>\n')
# KPI 4: Top ASIN
f.write('<div class="kpi"><div class="lb">Top ASIN</div>')
f.write('<div class="vl" style="font-size:16px">' + ta + '</div>')
f.write('<div class="ch">' + fm(tv) + '</div>')
f.write('</div>\n')
f.write('</div>\n')

# Tabs
f.write('<div class="tabs">\n')
tab_names = ['Summary', 'Daily Trend', 'Product Group', 'Top ASINs']
for i, t in enumerate(tab_names):
    cls = ' active' if i == 0 else ''
    f.write('<div class="tab' + cls + '" onclick="showTab(' + str(i) + ')">' + t + '</div>\n')
f.write('</div>\n')

# ── Panel 0: Summary ───────────────────────────────────────────
f.write('<div class="pnl active" id="p0"><div class="card">\n')
f.write('<h2>Weekly Summary</h2>\n')
f.write('<table><thead><tr>')
for h in ['Metric', 'W13 2026', 'W12 2026', 'WoW Delta', 'WoW %', 'W13 2025', 'YoY Delta', 'YoY %']:
    f.write('<th>' + h + '</th>')
f.write('</tr></thead><tbody>\n')

mlist = [('3P Net OPS (USD)', 3, True), ('3P Net Units', 4, False), ('All Net OPS (USD)', 5, True), ('All Net Units', 6, False)]
for name, idx, is_m in mlist:
    s = fnum(sw[idx]); p = fnum(pw[idx]); y = fnum(py_[idx])
    wd = s - p; wp = pct(s, p); yd = s - y; yp = pct(s, y)
    ff = fm if is_m else fi
    f.write('<tr>')
    f.write('<td><strong>' + name + '</strong></td>')
    f.write('<td class="tr">' + ff(s) + '</td>')
    f.write('<td class="tr">' + ff(p) + '</td>')
    f.write('<td class="tr">' + ff(wd) + '</td>')
    f.write('<td class="tc"><span class="bd ' + pc(wp) + '-bg">' + ar(wp) + ' ' + fp(wp) + '</span></td>')
    f.write('<td class="tr">' + ff(y) + '</td>')
    f.write('<td class="tr">' + ff(yd) + '</td>')
    f.write('<td class="tc"><span class="bd ' + pc(yp) + '-bg">' + ar(yp) + ' ' + fp(yp) + '</span></td>')
    f.write('</tr>\n')
f.write('</tbody></table>\n')
f.write('</div></div>\n')

# ── Panel 1: Daily Trend ───────────────────────────────────────
f.write('<div class="pnl" id="p1">\n')
dlbl = ['Sun 3/22', 'Mon 3/23', 'Tue 3/24', 'Wed 3/25', 'Thu 3/26', 'Fri 3/27', 'Sat 3/28']

def write_daily(title, ksuf, is_m=True):
    ff = fm if is_m else fi
    f.write('<div class="card"><h2>' + title + '</h2>\n')
    f.write('<table><thead><tr><th>Period</th>')
    for d in dlbl:
        f.write('<th>' + d + '</th>')
    f.write('<th>Total</th></tr></thead><tbody>\n')
    pairs = [('W13 2026', 'Selected week|All|' + ksuf), ('W12 2026', 'Previous week|All|' + ksuf), ('W13 2025', 'Previous year|All|' + ksuf)]
    all_v = {}
    for lbl, key in pairs:
        vals = dd.get(key, ['0'] * 7)
        rv = [fnum(vals[j]) for j in range(7)]
        all_v[lbl] = rv
        tot = sum(rv)
        f.write('<tr><td><strong>' + lbl + '</strong></td>')
        for v in rv:
            f.write('<td class="tr">' + ff(v) + '</td>')
        f.write('<td class="tr"><strong>' + ff(tot) + '</strong></td></tr>\n')
    sv = all_v.get('W13 2026', [0]*7)
    pv = all_v.get('W12 2026', [0]*7)
    yv = all_v.get('W13 2025', [0]*7)
    f.write('<tr><td><strong><em>WoW %</em></strong></td>')
    for j in range(7):
        p = pct(sv[j], pv[j])
        f.write('<td class="tc"><span class="bd ' + pc(p) + '-bg">' + fp(p) + '</span></td>')
    f.write('<td></td></tr>\n')
    f.write('<tr><td><strong><em>YoY %</em></strong></td>')
    for j in range(7):
        p = pct(sv[j], yv[j])
        f.write('<td class="tc"><span class="bd ' + pc(p) + '-bg">' + fp(p) + '</span></td>')
    f.write('<td></td></tr>\n')
    f.write('</tbody></table></div>\n')

write_daily('Daily Net OPS (USD)', 'NET_OPS', True)
write_daily('Daily Net Units', 'NET_UNITS', False)
f.write('</div>\n')

# ── Panel 2: Product Group ─────────────────────────────────────
f.write('<div class="pnl" id="p2">\n')

def write_pg(title, d13, d12, dly, is_m=True):
    ff = fm if is_m else fi
    f.write('<div class="card"><h2>' + title + '</h2>\n')
    f.write('<table><thead><tr>')
    for h in ['Product Group', 'W13 2026', 'W12 2026', 'W13 2025', 'WoW %', 'YoY %']:
        f.write('<th>' + h + '</th>')
    f.write('</tr></thead><tbody>\n')
    for pg, val in d13.items():
        v12 = d12.get(pg, 0); vly = dly.get(pg, 0)
        wp = pct(val, v12); yp = pct(val, vly)
        f.write('<tr>')
        f.write('<td><strong>' + pg + '</strong></td>')
        f.write('<td class="tr">' + ff(val) + '</td>')
        f.write('<td class="tr">' + ff(v12) + '</td>')
        f.write('<td class="tr">' + ff(vly) + '</td>')
        f.write('<td class="tc"><span class="bd ' + pc(wp) + '-bg">' + fp(wp) + '</span></td>')
        f.write('<td class="tc"><span class="bd ' + pc(yp) + '-bg">' + fp(yp) + '</span></td>')
        f.write('</tr>\n')
    f.write('</tbody></table></div>\n')

write_pg('Net OPS by Product Group (USD)', w13['pg_ops'], w12['pg_ops'], ly['pg_ops'], True)
write_pg('Net Units by Product Group', w13['pg_units'], w12['pg_units'], ly['pg_units'], False)
f.write('</div>\n')

# ── Panel 3: Top ASINs ────────────────────────────────────────
f.write('<div class="pnl" id="p3"><div class="card">\n')
f.write('<h2>Top 30 ASINs by Net OPS</h2>\n')
f.write('<table><thead><tr>')
for h in ['#', 'ASIN', 'Title', 'Net OPS (USD)', 'Share']:
    f.write('<th>' + h + '</th>')
f.write('</tr></thead><tbody>\n')
tot_ops = fnum(sw[5])
for i, a in enumerate(asins[:30]):
    sh = a['value'] / tot_ops if tot_ops else 0
    tt = esc(a['title'][:70])
    f.write('<tr>')
    f.write('<td class="tc">' + str(i+1) + '</td>')
    f.write('<td>' + a['asin'] + '</td>')
    f.write('<td class="at">' + tt + '</td>')
    f.write('<td class="tr">' + fm(a['value']) + '</td>')
    f.write('<td class="tc">{:.1f}%</td>'.format(sh * 100))
    f.write('</tr>\n')
f.write('</tbody></table>\n')
f.write('</div></div>\n')

# JavaScript
f.write('<script>\n')
f.write('function showTab(i){')
f.write('document.querySelectorAll(".tab").forEach((t,j)=>t.classList.toggle("active",j===i));')
f.write('document.querySelectorAll(".pnl").forEach((p,j)=>p.classList.toggle("active",j===i));')
f.write('}\n')
f.write('</script>\n')
f.write('</div>\n</body>\n</html>\n')
f.close()

print('HTML WBR saved: W13_WBR_AU30_Pipeline.html (' + str(os.path.getsize('W13_WBR_AU30_Pipeline.html')) + ' bytes)')
