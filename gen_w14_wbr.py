import openpyxl, glob, re, sys
from collections import defaultdict
import os, datetime

# ── Auto-detect latest week folder ─────────────────────────────
week_dirs = sorted([d for d in glob.glob('W[0-9]*') if os.path.isdir(d)], key=lambda d: int(re.sub(r'\D', '', d) or '0'))
if not week_dirs:
    print("ERROR: No W## folders found."); sys.exit(1)
latest_dir = week_dirs[-1]
WEEK_NUM = int(re.sub(r'\D', '', latest_dir))
print("Detected latest week folder: %s (W%d)" % (latest_dir, WEEK_NUM))

# Find the xlsx file inside
xlsx_files = glob.glob(os.path.join(latest_dir, '*.xlsx'))
if not xlsx_files:
    print("ERROR: No .xlsx file found in %s" % latest_dir); sys.exit(1)
data_file = xlsx_files[0]
print("Loading: %s" % data_file)

wb = openpyxl.load_workbook(data_file, read_only=True, data_only=True)
ws = wb.active

AU_MP = 111172
CUR_YEAR = 2026
PREV_YEAR = 2025

def decode_year(v):
    if isinstance(v, (int, float)):
        return int(v)
    if isinstance(v, datetime.datetime):
        return (v - datetime.datetime(1899, 12, 30)).days
    return 0

# Key = (year, week), e.g. (2026, 14), (2025, 14)
totals = defaultdict(lambda: {'gms':0,'units':0,'sellers':0,'fba_gms':0,'fba_units':0})
pg_data = defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0}))
cat_data = defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0}))
seller_agg = defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0,'fba_gms':0,'cats':set(),'pgs':set(),'active':0}))

for r in ws.iter_rows(min_row=2, values_only=True):
    if r[10] != AU_MP:
        continue
    yr = decode_year(r[1])
    wk = r[2]
    key = (yr, wk)
    gms = r[90] or 0
    units = r[93] or 0
    fba_gms = r[91] or 0
    fba_units = r[94] or 0
    name = str(r[25] or r[21] or 'Unknown')
    cat = r[34] or 'Other'
    pg = r[36] or 'Other'
    active = r[63] or 0

    totals[key]['gms'] += gms
    totals[key]['units'] += units
    totals[key]['fba_gms'] += fba_gms
    totals[key]['fba_units'] += fba_units
    if active:
        totals[key]['sellers'] += 1

    pg_data[pg][key]['gms'] += gms
    pg_data[pg][key]['units'] += units

    cat_data[cat][key]['gms'] += gms
    cat_data[cat][key]['units'] += units

    d = seller_agg[name][key]
    d['gms'] += gms
    d['units'] += units
    d['fba_gms'] += fba_gms
    d['cats'].add(cat)
    d['pgs'].add(pg)
    d['active'] += active

wb.close()
print("Data loaded.")

# Debug: print totals
for k in sorted(totals.keys()):
    t = totals[k]
    print('  %s: GMS=$%s Units=%s Sellers=%d' % (k, '{:,.0f}'.format(t['gms']), '{:,.0f}'.format(t['units']), t['sellers']))

# ── Shorthand keys ─────────────────────────────────────────────
CW = (CUR_YEAR, WEEK_NUM)       # Current week
PW = (CUR_YEAR, WEEK_NUM - 1)   # Previous week
LY = (PREV_YEAR, WEEK_NUM)      # Last year same week
empty = {'gms':0,'units':0,'sellers':0,'fba_gms':0,'fba_units':0}

T_CW = totals.get(CW, empty)
T_PW = totals.get(PW, empty)
T_LY = totals.get(LY, empty)
T_W12 = totals.get((CUR_YEAR, WEEK_NUM - 2), empty)

# Dynamic labels
CW_LBL = 'W%d %d' % (WEEK_NUM, CUR_YEAR)
PW_LBL = 'W%d %d' % (WEEK_NUM - 1, CUR_YEAR)
LY_LBL = 'W%d %d' % (WEEK_NUM, PREV_YEAR)
W12_LBL = 'W%d %d' % (WEEK_NUM - 2, CUR_YEAR)
OUT_FILE = 'W%d_WBR_AU30_Pipeline.html' % WEEK_NUM

# ── Helpers ────────────────────────────────────────────────────
def fm(v): return '${:,.0f}'.format(v)
def fi(v): return '{:,.0f}'.format(v)
def fp(v): return '{:+.1f}%'.format(v * 100)
def pct(a, b): return (a - b) / b if b else 0
def pc(v): return 'pos' if v >= 0 else 'neg'
def ar(v): return '&#9650;' if v >= 0 else '&#9660;'
def esc(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
def badge(v):
    c = pc(v)
    return '<span class="bd %s-bg">%s %s</span>' % (c, ar(v), fp(v))

wow_gms = pct(T_CW['gms'], T_PW['gms'])
wow_units = pct(T_CW['units'], T_PW['units'])
yoy_gms = pct(T_CW['gms'], T_LY['gms'])
yoy_units = pct(T_CW['units'], T_LY['units'])

# ── Prepare seller list ────────────────────────────────────────
sellers = []
for name, wdata in seller_agg.items():
    g_cw = wdata[CW]['gms']
    g_pw = wdata[PW]['gms']
    g_ly = wdata[LY]['gms']
    g_w12 = wdata[(CUR_YEAR, WEEK_NUM - 2)]['gms']
    u_cw = wdata[CW]['units']
    u_pw = wdata[PW]['units']
    u_ly = wdata[LY]['units']
    cats = wdata[CW]['cats'] or wdata[PW]['cats']
    pgs = wdata[CW]['pgs'] or wdata[PW]['pgs']
    if g_cw > 0 or g_pw > 0:
        sellers.append({
            'name': name,
            'g_cw': g_cw, 'g_pw': g_pw, 'g_ly': g_ly, 'g_w12': g_w12,
            'u_cw': u_cw, 'u_pw': u_pw, 'u_ly': u_ly,
            'cats': cats, 'pgs': pgs,
            'wow_d': g_cw - g_pw,
            'wow': pct(g_cw, g_pw),
            'yoy_d': g_cw - g_ly,
            'yoy': pct(g_cw, g_ly),
        })

sellers_by_gms = sorted(sellers, key=lambda x: -x['g_cw'])
movers_up = sorted([s for s in sellers if s['g_cw'] >= 300 or s['g_pw'] >= 300], key=lambda x: -x['wow_d'])[:15]
movers_down = sorted([s for s in sellers if s['g_cw'] >= 300 or s['g_pw'] >= 300], key=lambda x: x['wow_d'])[:15]

# Product groups
pgs_list = []
for pg, wdata in pg_data.items():
    g_cw = wdata[CW]['gms']; g_pw = wdata[PW]['gms']; g_ly = wdata[LY]['gms']
    u_cw = wdata[CW]['units']; u_pw = wdata[PW]['units']; u_ly = wdata[LY]['units']
    if g_cw > 0 or g_pw > 0:
        pgs_list.append({'pg': pg, 'g_cw': g_cw, 'g_pw': g_pw, 'g_ly': g_ly, 'u_cw': u_cw, 'u_pw': u_pw, 'u_ly': u_ly})
pgs_by_gms = sorted(pgs_list, key=lambda x: -x['g_cw'])

# Categories
cats_list = []
for cat, wdata in cat_data.items():
    g_cw = wdata[CW]['gms']; g_pw = wdata[PW]['gms']; g_ly = wdata[LY]['gms']
    u_cw = wdata[CW]['units']; u_pw = wdata[PW]['units']; u_ly = wdata[LY]['units']
    if g_cw > 0 or g_pw > 0:
        cats_list.append({'cat': cat, 'g_cw': g_cw, 'g_pw': g_pw, 'g_ly': g_ly, 'u_cw': u_cw, 'u_pw': u_pw, 'u_ly': u_ly})
cats_by_gms = sorted(cats_list, key=lambda x: -x['g_cw'])

total_g_cw = T_CW['gms']

# ══════════════════════════════════════════════════════════════
# Write HTML
# ══════════════════════════════════════════════════════════════
print("Generating HTML...")
f = open(OUT_FILE, 'w', encoding='utf-8')

f.write('<!DOCTYPE html>\n<html lang="en">\n<head>\n')
f.write('<meta charset="UTF-8">\n')
f.write('<meta name="viewport" content="width=device-width, initial-scale=1.0">\n')
f.write('<title>WBR - TW2AU Pipeline %s</title>\n' % CW_LBL)
f.write('<style>\n')
f.write(':root{--blue:#4472C4;--bl:#D6E4F0;--g:#006100;--gb:#C6EFCE;--r:#9C0006;--rb:#FFC7CE;--gy:#F2F2F2;--dk:#1a1a2e;--bd:#dee2e6}\n')
f.write('*{margin:0;padding:0;box-sizing:border-box}\n')
f.write('body{font-family:"Segoe UI",system-ui,sans-serif;background:#f0f2f5;color:#333;line-height:1.5}\n')
f.write('.ctn{max-width:1500px;margin:0 auto;padding:20px}\n')
f.write('.hdr{background:linear-gradient(135deg,var(--dk),var(--blue));color:#fff;padding:30px 40px;border-radius:12px;margin-bottom:24px}\n')
f.write('.hdr h1{font-size:24px;font-weight:700;margin-bottom:4px}\n')
f.write('.hdr p{opacity:.85;font-size:14px}\n')
f.write('.tabs{display:flex;gap:4px;margin-bottom:20px;flex-wrap:wrap}\n')
f.write('.tab{padding:10px 20px;background:#fff;border:1px solid var(--bd);border-radius:8px 8px 0 0;cursor:pointer;font-size:13px;font-weight:600;color:#666;transition:all .2s}\n')
f.write('.tab:hover{color:var(--blue)}\n')
f.write('.tab.active{background:var(--blue);color:#fff;border-color:var(--blue)}\n')
f.write('.pnl{display:none}.pnl.active{display:block}\n')
f.write('.card{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:24px;margin-bottom:20px}\n')
f.write('.card h2{font-size:16px;color:var(--blue);margin-bottom:16px;border-bottom:2px solid var(--bl);padding-bottom:8px}\n')
f.write('.card h3{font-size:14px;color:#555;margin:16px 0 8px}\n')
f.write('.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px;margin-bottom:24px}\n')
f.write('.kpi{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);border-left:4px solid var(--blue)}\n')
f.write('.kpi .lb{font-size:12px;color:#888;text-transform:uppercase;letter-spacing:.5px}\n')
f.write('.kpi .vl{font-size:28px;font-weight:700;margin:4px 0}\n')
f.write('.kpi .ch{font-size:13px;font-weight:600}\n')
f.write('.pos{color:var(--g)}.neg{color:var(--r)}\n')
f.write('.pos-bg{background:var(--gb);color:var(--g)}\n')
f.write('.neg-bg{background:var(--rb);color:var(--r)}\n')
f.write('table{width:100%;border-collapse:collapse;font-size:13px}\n')
f.write('th{background:var(--blue);color:#fff;padding:10px 12px;text-align:center;font-weight:600;white-space:nowrap}\n')
f.write('td{padding:8px 12px;border-bottom:1px solid #eee}\n')
f.write('tr:nth-child(even){background:var(--gy)}\n')
f.write('.tr{text-align:right}.tc{text-align:center}.tl{text-align:left}\n')
f.write('.bd{display:inline-block;padding:2px 8px;border-radius:4px;font-weight:600;font-size:12px}\n')
f.write('.at{max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px}\n')
f.write('.mv-up{border-left:4px solid var(--g)}.mv-dn{border-left:4px solid var(--r)}\n')
f.write('.mv-card{background:#fff;border-radius:10px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:12px}\n')
f.write('.mv-card .nm{font-size:15px;font-weight:700;margin-bottom:4px}\n')
f.write('.mv-card .dt{font-size:13px;color:#555}\n')
f.write('.bar-container{display:flex;align-items:center;gap:8px;margin:4px 0}\n')
f.write('.bar{height:18px;border-radius:3px;min-width:2px}\n')
f.write('.bar-cw{background:var(--blue)}.bar-pw{background:#B4C7E7}.bar-ly{background:#FFC000}\n')
f.write('@media(max-width:768px){.kg{grid-template-columns:1fr}table{font-size:11px}th,td{padding:6px 8px}.hdr{padding:20px}.hdr h1{font-size:18px}}\n')
f.write('</style>\n</head>\n<body>\n<div class="ctn">\n')

# ── Header ─────────────────────────────────────────────────────
f.write('<div class="hdr">\n')
f.write('  <h1>WBR &mdash; TW2AU Pipeline Weekly Business Review</h1>\n')
f.write('  <p>%s &nbsp;|&nbsp; AU Marketplace (111172) &nbsp;|&nbsp; USD &nbsp;|&nbsp; Source: MCID Weekly Data</p>\n' % CW_LBL)
f.write('</div>\n')

# ── KPI Cards (with WoW + YoY) ────────────────────────────────
f.write('<div class="kg">\n')
f.write('<div class="kpi"><div class="lb">Total Ordered GMS (USD)</div>')
f.write('<div class="vl">%s</div>' % fm(T_CW['gms']))
f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wow_gms), ar(wow_gms), fp(wow_gms)))
f.write('<div class="ch %s">YoY: %s %s</div>' % (pc(yoy_gms), ar(yoy_gms), fp(yoy_gms)))
f.write('</div>\n')

f.write('<div class="kpi"><div class="lb">Total Ordered Units</div>')
f.write('<div class="vl">%s</div>' % fi(T_CW['units']))
f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wow_units), ar(wow_units), fp(wow_units)))
f.write('<div class="ch %s">YoY: %s %s</div>' % (pc(yoy_units), ar(yoy_units), fp(yoy_units)))
f.write('</div>\n')

f.write('<div class="kpi"><div class="lb">Active Sellers</div>')
f.write('<div class="vl">%d</div>' % T_CW['sellers'])
wow_s = pct(T_CW['sellers'], T_PW['sellers'])
yoy_s = pct(T_CW['sellers'], T_LY['sellers'])
f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wow_s), ar(wow_s), fp(wow_s)))
f.write('<div class="ch %s">YoY: %s %s</div>' % (pc(yoy_s), ar(yoy_s), fp(yoy_s)))
f.write('</div>\n')

top_seller = sellers_by_gms[0]['name'] if sellers_by_gms else 'N/A'
top_seller_v = sellers_by_gms[0]['g_cw'] if sellers_by_gms else 0
f.write('<div class="kpi"><div class="lb">Top Seller</div>')
f.write('<div class="vl" style="font-size:16px">%s</div>' % esc(top_seller))
f.write('<div class="ch">%s</div>' % fm(top_seller_v))
f.write('</div>\n')
f.write('</div>\n')

# ── Executive Summary (English) ────────────────────────────────
# Build dynamic narrative
top3_up = movers_up[:3]
top3_dn = movers_down[:3]
top_pg_name = pgs_by_gms[0]['pg'] if pgs_by_gms else 'N/A'
top_pg_share = pgs_by_gms[0]['g_cw'] / total_g_cw * 100 if pgs_by_gms and total_g_cw else 0

gainers_str = ', '.join(['%s (%s)' % (esc(s['name']), fp(s['wow'])) for s in top3_up])
decliners_str = ', '.join(['%s (%s)' % (esc(s['name']), fp(s['wow'])) for s in top3_dn])

fba_ratio = T_CW['fba_gms'] / T_CW['gms'] * 100 if T_CW['gms'] else 0

f.write('<div class="card" style="border-left:4px solid var(--blue);margin-bottom:20px">\n')
f.write('<h2>&#128221; Executive Summary</h2>\n')
f.write('<div style="font-size:14px;line-height:1.8;color:#333">\n')

f.write('<p><strong>Overall Performance:</strong> ')
f.write('%s total ordered GMS came in at <strong>%s</strong>, ' % (CW_LBL, fm(T_CW['gms'])))
wow_dir = 'down' if wow_gms < 0 else 'up'
yoy_dir = 'down' if yoy_gms < 0 else 'up'
f.write('%s <strong>%s WoW</strong> from %s (%s) and ' % (wow_dir, fp(wow_gms), fm(T_PW['gms']), PW_LBL))
f.write('%s <strong>%s YoY</strong> vs %s (%s). ' % (yoy_dir, fp(yoy_gms), fm(T_LY['gms']), LY_LBL))
f.write('Total ordered units were <strong>%s</strong> (WoW %s, YoY %s). ' % (fi(T_CW['units']), fp(wow_units), fp(yoy_units)))
f.write('FBA accounted for <strong>%.1f%%</strong> of total GMS.</p>\n' % fba_ratio)

f.write('<p><strong>Seller Landscape:</strong> ')
f.write('<strong>%d</strong> active sellers in %s (WoW %s, YoY %s). ' % (T_CW['sellers'], CW_LBL, fp(wow_s), fp(yoy_s)))
f.write('The top seller was <strong>%s</strong> with %s in GMS (%.1f%% share).</p>\n' % (
    esc(top_seller), fm(top_seller_v), top_seller_v / total_g_cw * 100 if total_g_cw else 0))

f.write('<p><strong>Top Product Group:</strong> ')
f.write('<strong>%s</strong> led with %.1f%% of total GMS.</p>\n' % (esc(top_pg_name), top_pg_share))

f.write('<p><strong>Movers &amp; Shakers:</strong> ')
f.write('Top WoW gainers include %s. ' % gainers_str)
f.write('Notable decliners include %s.</p>\n' % decliners_str)

# Key callouts
f.write('<p><strong>Key Callouts:</strong></p>\n')
f.write('<ul style="margin:4px 0 0 20px;font-size:13px">\n')

# Auto-detect notable patterns
if wow_gms < -0.05:
    f.write('<li style="margin-bottom:4px">GMS declined %s WoW &mdash; monitor whether this is seasonal or structural.</li>\n' % fp(wow_gms))
elif wow_gms > 0.05:
    f.write('<li style="margin-bottom:4px">Strong GMS growth of %s WoW &mdash; identify drivers to sustain momentum.</li>\n' % fp(wow_gms))

if yoy_gms > 0:
    f.write('<li style="margin-bottom:4px">Positive YoY trajectory (%s) indicates healthy pipeline growth vs last year.</li>\n' % fp(yoy_gms))
else:
    f.write('<li style="margin-bottom:4px">YoY GMS is %s &mdash; investigate root causes and recovery plan.</li>\n' % fp(yoy_gms))

# Highlight big movers
for s in movers_up[:2]:
    if s['wow_d'] > 1000:
        f.write('<li style="margin-bottom:4px"><strong>%s</strong> surged %s WoW (+%s) &mdash; worth investigating what drove the spike.</li>\n' % (esc(s['name']), fp(s['wow']), fm(s['wow_d'])))
for s in movers_down[:2]:
    if s['wow_d'] < -1000:
        f.write('<li style="margin-bottom:4px"><strong>%s</strong> dropped %s WoW (%s) &mdash; follow up on potential issues.</li>\n' % (esc(s['name']), fp(s['wow']), fm(s['wow_d'])))

f.write('</ul>\n')
f.write('</div>\n')
f.write('</div>\n')

# ── Tabs ───────────────────────────────────────────────────────
f.write('<div class="tabs">\n')
tab_names = ['Summary', 'Category / PG', 'Top Sellers', 'Movers &amp; Shakers', 'Seller Deep Dive']
for i, t in enumerate(tab_names):
    cls = ' active' if i == 0 else ''
    f.write('<div class="tab%s" onclick="showTab(%d)">%s</div>\n' % (cls, i, t))
f.write('</div>\n')

# ══════════════════════════════════════════════════════════════
# Panel 0: Summary (with YoY)
# ══════════════════════════════════════════════════════════════
f.write('<div class="pnl active" id="p0"><div class="card">\n')
f.write('<h2>Weekly Summary (AU Marketplace)</h2>\n')
f.write('<table><thead><tr>')
for h in ['Metric', CW_LBL, PW_LBL, 'WoW Delta', 'WoW %', LY_LBL, 'YoY Delta', 'YoY %']:
    f.write('<th>%s</th>' % h)
f.write('</tr></thead><tbody>\n')

summary_rows = [
    ('Ordered GMS (USD)', T_CW['gms'], T_PW['gms'], T_LY['gms'], True),
    ('Ordered Units', T_CW['units'], T_PW['units'], T_LY['units'], False),
    ('FBA GMS (USD)', T_CW['fba_gms'], T_PW['fba_gms'], T_LY['fba_gms'], True),
    ('FBA Units', T_CW['fba_units'], T_PW['fba_units'], T_LY['fba_units'], False),
    ('Active Sellers', T_CW['sellers'], T_PW['sellers'], T_LY['sellers'], False),
]
for name, v_cw, v_pw, v_ly, is_m in summary_rows:
    ff = fm if is_m else fi
    wd = v_cw - v_pw; wp = pct(v_cw, v_pw)
    yd = v_cw - v_ly; yp = pct(v_cw, v_ly)
    f.write('<tr>')
    f.write('<td><strong>%s</strong></td>' % name)
    f.write('<td class="tr">%s</td>' % ff(v_cw))
    f.write('<td class="tr">%s</td>' % ff(v_pw))
    f.write('<td class="tr">%s</td>' % ff(wd))
    f.write('<td class="tc">%s</td>' % badge(wp))
    f.write('<td class="tr">%s</td>' % ff(v_ly))
    f.write('<td class="tr">%s</td>' % ff(yd))
    f.write('<td class="tc">%s</td>' % badge(yp))
    f.write('</tr>\n')
f.write('</tbody></table>\n')
f.write('</div></div>\n')

# ══════════════════════════════════════════════════════════════
# Panel 1: Category / Product Group (with YoY)
# ══════════════════════════════════════════════════════════════
f.write('<div class="pnl" id="p1">\n')

# Category table
f.write('<div class="card"><h2>GMS by Account Primary Category</h2>\n')
f.write('<table><thead><tr>')
for h in ['Category', CW_LBL, PW_LBL, 'WoW %', LY_LBL, 'YoY %', CW_LBL + ' Units', 'Share']:
    f.write('<th>%s</th>' % h)
f.write('</tr></thead><tbody>\n')
for c in cats_by_gms[:20]:
    wp = pct(c['g_cw'], c['g_pw']); yp = pct(c['g_cw'], c['g_ly'])
    share = c['g_cw'] / total_g_cw if total_g_cw else 0
    f.write('<tr>')
    f.write('<td><strong>%s</strong></td>' % esc(c['cat']))
    f.write('<td class="tr">%s</td>' % fm(c['g_cw']))
    f.write('<td class="tr">%s</td>' % fm(c['g_pw']))
    f.write('<td class="tc">%s</td>' % badge(wp))
    f.write('<td class="tr">%s</td>' % fm(c['g_ly']))
    f.write('<td class="tc">%s</td>' % badge(yp))
    f.write('<td class="tr">%s</td>' % fi(c['u_cw']))
    f.write('<td class="tc">%.1f%%</td>' % (share * 100))
    f.write('</tr>\n')
f.write('</tbody></table></div>\n')

# PG table
f.write('<div class="card"><h2>GMS by SP Primary Product Group</h2>\n')
f.write('<table><thead><tr>')
for h in ['Product Group', CW_LBL, PW_LBL, 'WoW %', LY_LBL, 'YoY %', CW_LBL + ' Units', 'Share']:
    f.write('<th>%s</th>' % h)
f.write('</tr></thead><tbody>\n')
for p in pgs_by_gms[:15]:
    wp = pct(p['g_cw'], p['g_pw']); yp = pct(p['g_cw'], p['g_ly'])
    share = p['g_cw'] / total_g_cw if total_g_cw else 0
    f.write('<tr>')
    f.write('<td><strong>%s</strong></td>' % esc(p['pg']))
    f.write('<td class="tr">%s</td>' % fm(p['g_cw']))
    f.write('<td class="tr">%s</td>' % fm(p['g_pw']))
    f.write('<td class="tc">%s</td>' % badge(wp))
    f.write('<td class="tr">%s</td>' % fm(p['g_ly']))
    f.write('<td class="tc">%s</td>' % badge(yp))
    f.write('<td class="tr">%s</td>' % fi(p['u_cw']))
    f.write('<td class="tc">%.1f%%</td>' % (share * 100))
    f.write('</tr>\n')
f.write('</tbody></table></div>\n')
f.write('</div>\n')

# ══════════════════════════════════════════════════════════════
# Panel 2: Top Sellers (with YoY)
# ══════════════════════════════════════════════════════════════
f.write('<div class="pnl" id="p2"><div class="card">\n')
f.write('<h2>Top 30 Sellers by %s Ordered GMS</h2>\n' % CW_LBL)
f.write('<table><thead><tr>')
for h in ['#', 'Seller', CW_LBL, PW_LBL, 'WoW %', LY_LBL, 'YoY %', 'Units', 'Category', 'Share']:
    f.write('<th>%s</th>' % h)
f.write('</tr></thead><tbody>\n')
for i, s in enumerate(sellers_by_gms[:30]):
    wp = s['wow']; yp = s['yoy']
    share = s['g_cw'] / total_g_cw if total_g_cw else 0
    cats_str = ', '.join(list(s['cats'] - {'Other'})[:2]) or 'Other'
    f.write('<tr>')
    f.write('<td class="tc">%d</td>' % (i + 1))
    f.write('<td><strong>%s</strong></td>' % esc(s['name']))
    f.write('<td class="tr">%s</td>' % fm(s['g_cw']))
    f.write('<td class="tr">%s</td>' % fm(s['g_pw']))
    f.write('<td class="tc">%s</td>' % badge(wp))
    f.write('<td class="tr">%s</td>' % fm(s['g_ly']))
    f.write('<td class="tc">%s</td>' % badge(yp))
    f.write('<td class="tr">%s</td>' % fi(s['u_cw']))
    f.write('<td class="at">%s</td>' % esc(cats_str))
    f.write('<td class="tc">%.1f%%</td>' % (share * 100))
    f.write('</tr>\n')
f.write('</tbody></table>\n')
f.write('</div></div>\n')

# ══════════════════════════════════════════════════════════════
# Panel 3: Movers & Shakers (with YoY)
# ══════════════════════════════════════════════════════════════
f.write('<div class="pnl" id="p3">\n')

# Gainers table
f.write('<div class="card"><h2>&#128293; Top Gainers (Biggest WoW GMS Increase)</h2>\n')
f.write('<table><thead><tr>')
for h in ['#', 'Seller', CW_LBL, PW_LBL, 'WoW Delta', 'WoW %', LY_LBL, 'YoY %', 'Category']:
    f.write('<th>%s</th>' % h)
f.write('</tr></thead><tbody>\n')
for i, s in enumerate(movers_up[:10]):
    cats_str = ', '.join(list(s['cats'] - {'Other'})[:2]) or 'Other'
    f.write('<tr>')
    f.write('<td class="tc">%d</td>' % (i + 1))
    f.write('<td><strong>%s</strong></td>' % esc(s['name']))
    f.write('<td class="tr">%s</td>' % fm(s['g_cw']))
    f.write('<td class="tr">%s</td>' % fm(s['g_pw']))
    f.write('<td class="tr" style="color:var(--g);font-weight:700">%s</td>' % fm(s['wow_d']))
    f.write('<td class="tc">%s</td>' % badge(s['wow']))
    f.write('<td class="tr">%s</td>' % fm(s['g_ly']))
    f.write('<td class="tc">%s</td>' % badge(s['yoy']))
    f.write('<td class="at">%s</td>' % esc(cats_str))
    f.write('</tr>\n')
f.write('</tbody></table></div>\n')

# Decliners table
f.write('<div class="card"><h2>&#128308; Top Decliners (Biggest WoW GMS Decrease)</h2>\n')
f.write('<table><thead><tr>')
for h in ['#', 'Seller', CW_LBL, PW_LBL, 'WoW Delta', 'WoW %', LY_LBL, 'YoY %', 'Category']:
    f.write('<th>%s</th>' % h)
f.write('</tr></thead><tbody>\n')
for i, s in enumerate(movers_down[:10]):
    cats_str = ', '.join(list(s['cats'] - {'Other'})[:2]) or 'Other'
    f.write('<tr>')
    f.write('<td class="tc">%d</td>' % (i + 1))
    f.write('<td><strong>%s</strong></td>' % esc(s['name']))
    f.write('<td class="tr">%s</td>' % fm(s['g_cw']))
    f.write('<td class="tr">%s</td>' % fm(s['g_pw']))
    f.write('<td class="tr" style="color:var(--r);font-weight:700">%s</td>' % fm(s['wow_d']))
    f.write('<td class="tc">%s</td>' % badge(s['wow']))
    f.write('<td class="tr">%s</td>' % fm(s['g_ly']))
    f.write('<td class="tc">%s</td>' % badge(s['yoy']))
    f.write('<td class="at">%s</td>' % esc(cats_str))
    f.write('</tr>\n')
f.write('</tbody></table></div>\n')

# Visual comparison bars
all_movers = movers_up[:5] + movers_down[:5]
max_gms = max(max(s['g_cw'], s['g_pw'], s['g_ly']) for s in all_movers) if all_movers else 1

f.write('<div class="card"><h2>Movers &amp; Shakers Visual Comparison</h2>\n')
f.write('<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">\n')

# Gainers visual
f.write('<div>\n')
f.write('<h3 style="color:var(--g);margin-bottom:12px">&#9650; Top Gainers</h3>\n')
for s in movers_up[:7]:
    f.write('<div class="mv-card mv-up">\n')
    f.write('<div class="nm">%s</div>\n' % esc(s['name']))
    for lbl, val, cls in [(CW_LBL, s['g_cw'], 'bar-cw'), (PW_LBL, s['g_pw'], 'bar-pw'), (LY_LBL, s['g_ly'], 'bar-ly')]:
        w = val / max_gms * 100 if max_gms else 0
        f.write('<div class="bar-container"><span style="width:45px;font-size:10px">%s</span><div class="bar %s" style="width:%.1f%%"></div><span style="font-size:11px">%s</span></div>\n' % (lbl, cls, w, fm(val)))
    f.write('<div class="dt">WoW: %s | YoY: %s</div>\n' % (fp(s['wow']), fp(s['yoy'])))
    f.write('</div>\n')
f.write('</div>\n')

# Decliners visual
f.write('<div>\n')
f.write('<h3 style="color:var(--r);margin-bottom:12px">&#9660; Top Decliners</h3>\n')
for s in movers_down[:7]:
    f.write('<div class="mv-card mv-dn">\n')
    f.write('<div class="nm">%s</div>\n' % esc(s['name']))
    for lbl, val, cls in [(CW_LBL, s['g_cw'], 'bar-cw'), (PW_LBL, s['g_pw'], 'bar-pw'), (LY_LBL, s['g_ly'], 'bar-ly')]:
        w = val / max_gms * 100 if max_gms else 0
        f.write('<div class="bar-container"><span style="width:45px;font-size:10px">%s</span><div class="bar %s" style="width:%.1f%%"></div><span style="font-size:11px">%s</span></div>\n' % (lbl, cls, w, fm(val)))
    f.write('<div class="dt">WoW: %s | YoY: %s</div>\n' % (fp(s['wow']), fp(s['yoy'])))
    f.write('</div>\n')
f.write('</div>\n')
f.write('</div></div>\n')
f.write('</div>\n')

# ══════════════════════════════════════════════════════════════
# Panel 4: Seller Deep Dive (top 10, with YoY)
# ══════════════════════════════════════════════════════════════
f.write('<div class="pnl" id="p4">\n')

for s in sellers_by_gms[:10]:
    wp = s['wow']; yp = s['yoy']
    cats_str = ', '.join(list(s['cats'] - {'Other'})[:3]) or 'Other'
    pgs_str = ', '.join(list(s['pgs'] - {'Other'})[:3]) or 'Other'

    f.write('<div class="card">\n')
    f.write('<h2>%s</h2>\n' % esc(s['name']))
    f.write('<div class="kg" style="margin-bottom:16px">\n')

    f.write('<div class="kpi"><div class="lb">%s GMS</div>' % CW_LBL)
    f.write('<div class="vl" style="font-size:22px">%s</div>' % fm(s['g_cw']))
    f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wp), ar(wp), fp(wp)))
    f.write('<div class="ch %s">YoY: %s %s</div>' % (pc(yp), ar(yp), fp(yp)))
    f.write('</div>\n')

    f.write('<div class="kpi"><div class="lb">%s GMS</div>' % PW_LBL)
    f.write('<div class="vl" style="font-size:22px">%s</div>' % fm(s['g_pw']))
    f.write('</div>\n')

    f.write('<div class="kpi"><div class="lb">%s GMS</div>' % LY_LBL)
    f.write('<div class="vl" style="font-size:22px">%s</div>' % fm(s['g_ly']))
    f.write('</div>\n')

    wu = pct(s['u_cw'], s['u_pw']); yu = pct(s['u_cw'], s['u_ly'])
    f.write('<div class="kpi"><div class="lb">%s Units</div>' % CW_LBL)
    f.write('<div class="vl" style="font-size:22px">%s</div>' % fi(s['u_cw']))
    f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wu), ar(wu), fp(wu)))
    f.write('<div class="ch %s">YoY: %s %s</div>' % (pc(yu), ar(yu), fp(yu)))
    f.write('</div>\n')

    f.write('</div>\n')  # kg

    f.write('<div style="display:flex;gap:20px;font-size:13px;color:#555">')
    f.write('<div><strong>Category:</strong> %s</div>' % esc(cats_str))
    f.write('<div><strong>Product Group:</strong> %s</div>' % esc(pgs_str))
    f.write('<div><strong>GMS Share:</strong> %.1f%%</div>' % (s['g_cw'] / total_g_cw * 100 if total_g_cw else 0))
    f.write('</div>\n')

    # Trend bars
    max_v = max(s['g_cw'], s['g_pw'], s['g_ly'], s['g_w12']) or 1
    f.write('<div style="margin-top:12px">')
    f.write('<h3 style="font-size:13px;color:#888;margin-bottom:6px">GMS Comparison</h3>')
    for lbl, val, color in [
        (W12_LBL, s['g_w12'], '#B4C7E7'),
        (PW_LBL, s['g_pw'], '#7BA0D4'),
        (CW_LBL, s['g_cw'], 'var(--blue)'),
        (LY_LBL, s['g_ly'], '#FFC000'),
    ]:
        w = val / max_v * 100 if max_v else 0
        f.write('<div class="bar-container"><span style="width:45px;font-size:11px;font-weight:600">%s</span>' % lbl)
        f.write('<div class="bar" style="width:%.1f%%;background:%s"></div>' % (w, color))
        f.write('<span style="font-size:11px">%s</span></div>\n' % fm(val))
    f.write('</div>\n')

    f.write('</div>\n')  # card

f.write('</div>\n')  # panel

# ── JavaScript ─────────────────────────────────────────────────
f.write('<script>\n')
f.write('function showTab(i){\n')
f.write('document.querySelectorAll(".tab").forEach((t,j)=>t.classList.toggle("active",j===i));\n')
f.write('document.querySelectorAll(".pnl").forEach((p,j)=>p.classList.toggle("active",j===i));\n')
f.write('}\n')
f.write('</script>\n')
f.write('</div>\n</body>\n</html>\n')
f.close()

size = os.path.getsize(OUT_FILE)
print('HTML WBR saved: %s (%d bytes)' % (OUT_FILE, size))
print('Done!')
