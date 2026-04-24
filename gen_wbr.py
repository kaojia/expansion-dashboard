"""
Weekly Business Review (WBR) Generator — Multi-Marketplace
Auto-detects latest W## folder, generates one HTML per marketplace.
Marketplaces: AU (111172), AE (338801), SA (338811)
"""
import openpyxl, glob, re, sys, os, datetime
from collections import defaultdict

# ── Config ─────────────────────────────────────────────────────
MARKETPLACES = [
    {'id': 111172, 'code': 'AU', 'name': 'Australia'},
    {'id': 338801, 'code': 'AE', 'name': 'UAE'},
    {'id': 338811, 'code': 'SA', 'name': 'Saudi Arabia'},
]
CUR_YEAR = 2026
PREV_YEAR = 2025

# ── Auto-detect latest week folder ─────────────────────────────
week_dirs = sorted(
    [d for d in glob.glob('W[0-9]*') if os.path.isdir(d)],
    key=lambda d: int(re.sub(r'\D', '', d) or '0')
)
if not week_dirs:
    print("ERROR: No W## folders found."); sys.exit(1)
latest_dir = week_dirs[-1]
WEEK_NUM = int(re.sub(r'\D', '', latest_dir))
print("Detected latest week folder: %s (W%d)" % (latest_dir, WEEK_NUM))

xlsx_files = glob.glob(os.path.join(latest_dir, '*.xlsx'))
if not xlsx_files:
    print("ERROR: No .xlsx file found in %s" % latest_dir); sys.exit(1)
data_file = xlsx_files[0]
print("Loading: %s" % data_file)

# ── Shorthand keys ─────────────────────────────────────────────
CW = (CUR_YEAR, WEEK_NUM)
PW = (CUR_YEAR, WEEK_NUM - 1)
LY = (PREV_YEAR, WEEK_NUM)
W_2 = (CUR_YEAR, WEEK_NUM - 2)
CW_LBL = 'W%d %d' % (WEEK_NUM, CUR_YEAR)
PW_LBL = 'W%d %d' % (WEEK_NUM - 1, CUR_YEAR)
LY_LBL = 'W%d %d' % (WEEK_NUM, PREV_YEAR)
W2_LBL = 'W%d %d' % (WEEK_NUM - 2, CUR_YEAR)

# ── Helpers ────────────────────────────────────────────────────
def decode_year(v):
    if isinstance(v, (int, float)): return int(v)
    if isinstance(v, datetime.datetime): return (v - datetime.datetime(1899, 12, 30)).days
    return 0

def fm(v): return '${:,.0f}'.format(v)
def fi(v): return '{:,.0f}'.format(v)
def fp(v): return '{:+.1f}%'.format(v * 100)
def pct(a, b): return (a - b) / b if b else 0
def pc(v): return 'pos' if v >= 0 else 'neg'
def ar(v): return '&#9650;' if v >= 0 else '&#9660;'
def esc(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
def badge(v): return '<span class="bd %s-bg">%s %s</span>' % (pc(v), ar(v), fp(v))
EMPTY = {'gms':0,'units':0,'sellers':0,'fba_gms':0,'fba_units':0,'ytd_gms':0,'ytd_units':0}

# ── Read all data once, keyed by (marketplace_id, year, week) ──
print("Reading all rows...")
wb = openpyxl.load_workbook(data_file, read_only=True, data_only=True)
ws = wb.active
mp_ids = {m['id'] for m in MARKETPLACES}

# nested: mp -> structure
raw = defaultdict(lambda: {
    'totals': defaultdict(lambda: {'gms':0,'units':0,'sellers':0,'fba_gms':0,'fba_units':0,'ytd_gms':0,'ytd_units':0}),
    'pg': defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0,'ytd_gms':0,'ytd_units':0})),
    'cat': defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0,'ytd_gms':0,'ytd_units':0})),
    'seller': defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0,'fba_gms':0,'ytd_gms':0,'ytd_units':0,'cats':set(),'pgs':set(),'active':0,'ytd_launch':0})),
    'cohort_totals': defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0,'sellers':0,'fba_gms':0,'ytd_gms':0,'ytd_units':0})),
    'cohort_seller': defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {'gms':0,'units':0,'fba_gms':0,'ytd_gms':0,'ytd_units':0,'cats':set(),'pgs':set(),'active':0}))),
    'seller_channel': defaultdict(lambda: ''),  # seller name -> launch channel (resolved after loading)
    'seller_channel_week': defaultdict(lambda: (0, '')),  # name -> (latest_week_in_CUR_YEAR, channel)
    'seller_launch_date': {},  # seller name -> launch_date (datetime)
    'seller_mcid': {},  # seller name -> merchant_customer_id (str)
})

def get_cohort(launch_channel):
    if launch_channel in ('DSR', 'SSR'): return 'NSR'
    if launch_channel == 'ESM': return 'ESM'
    return 'Other'

for r in ws.iter_rows(min_row=2, values_only=True):
    mp = r[10]
    if mp not in mp_ids: continue
    yr = decode_year(r[1]); wk = r[2]; key = (yr, wk)
    gms = r[90] or 0; units = r[93] or 0
    fba_gms = r[91] or 0; fba_units = r[94] or 0
    ytd_gms = r[96] or 0; ytd_units = r[99] or 0
    name = str(r[25] or r[21] or 'Unknown')
    cat = r[34] or 'Other'; pg = r[36] or 'Other'
    active = r[63] or 0
    cohort = get_cohort(r[12])

    d = raw[mp]
    t = d['totals'][key]
    t['gms'] += gms; t['units'] += units; t['fba_gms'] += fba_gms; t['fba_units'] += fba_units
    t['ytd_gms'] += ytd_gms; t['ytd_units'] += ytd_units
    if active: t['sellers'] += 1
    d['pg'][pg][key]['gms'] += gms; d['pg'][pg][key]['units'] += units
    d['pg'][pg][key]['ytd_gms'] += ytd_gms; d['pg'][pg][key]['ytd_units'] += ytd_units
    d['cat'][cat][key]['gms'] += gms; d['cat'][cat][key]['units'] += units
    d['cat'][cat][key]['ytd_gms'] += ytd_gms; d['cat'][cat][key]['ytd_units'] += ytd_units
    s = d['seller'][name][key]
    s['gms'] += gms; s['units'] += units; s['fba_gms'] += fba_gms
    s['ytd_gms'] += ytd_gms; s['ytd_units'] += ytd_units
    s['cats'].add(cat); s['pgs'].add(pg); s['active'] += active
    # Store ytd_launch flag per seller per week
    if r[54] == 1:
        s['ytd_launch'] = 1
    # Track channel from CUR_YEAR data (use latest week)
    if r[12] and yr == CUR_YEAR:
        prev_wk, _ = d['seller_channel_week'][name]
        if wk >= prev_wk:
            d['seller_channel_week'][name] = (wk, str(r[12]))
    # Store launch date per seller
    if name not in d['seller_launch_date'] and r[28]:
        d['seller_launch_date'][name] = r[28]
    # Store merchant_customer_id per seller (as plain integer string)
    if name not in d['seller_mcid'] and r[21]:
        d['seller_mcid'][name] = str(int(r[21])) if isinstance(r[21], (int, float)) else str(r[21])
    # Cohort aggregation
    ct = d['cohort_totals'][cohort][key]
    ct['gms'] += gms; ct['units'] += units; ct['fba_gms'] += fba_gms
    ct['ytd_gms'] += ytd_gms; ct['ytd_units'] += ytd_units
    if active: ct['sellers'] = ct.get('sellers', 0) + 1
    cs = d['cohort_seller'][cohort][name][key]
    cs['gms'] += gms; cs['units'] += units; cs['fba_gms'] += fba_gms
    cs['ytd_gms'] += ytd_gms; cs['ytd_units'] += ytd_units
    cs['cats'].add(cat); cs['pgs'].add(pg); cs['active'] += active

wb.close()

# Resolve seller_channel: use the channel from the latest week in CUR_YEAR
for mp_id, d in raw.items():
    for name, (wk, ch) in d['seller_channel_week'].items():
        if ch:
            d['seller_channel'][name] = ch

print("Data loaded for %d marketplaces." % len(raw))

# ══════════════════════════════════════════════════════════════
# HTML generation function — one call per marketplace
# ══════════════════════════════════════════════════════════════
def generate_wbr(mp_id, mp_code, mp_name):
    d = raw[mp_id]
    totals = d['totals']
    T_CW = totals.get(CW, EMPTY); T_PW = totals.get(PW, EMPTY)
    T_LY = totals.get(LY, EMPTY); T_W2 = totals.get(W_2, EMPTY)

    if T_CW['gms'] == 0 and T_PW['gms'] == 0:
        print("  [%s] No GMS data — skipping." % mp_code)
        return

    wow_gms = pct(T_CW['gms'], T_PW['gms']); wow_units = pct(T_CW['units'], T_PW['units'])
    yoy_gms = pct(T_CW['gms'], T_LY['gms']); yoy_units = pct(T_CW['units'], T_LY['units'])
    wow_s = pct(T_CW['sellers'], T_PW['sellers']); yoy_s = pct(T_CW['sellers'], T_LY['sellers'])

    # Sellers
    sellers = []
    for name, wdata in d['seller'].items():
        g_cw=wdata[CW]['gms']; g_pw=wdata[PW]['gms']; g_ly=wdata[LY]['gms']; g_w2=wdata[W_2]['gms']
        u_cw=wdata[CW]['units']; u_pw=wdata[PW]['units']; u_ly=wdata[LY]['units']
        ytd_cw=wdata[CW]['ytd_gms']; ytd_ly=wdata[LY]['ytd_gms']
        ytd_u_cw=wdata[CW]['ytd_units']
        cats=wdata[CW]['cats'] or wdata[PW]['cats']; pgs=wdata[CW]['pgs'] or wdata[PW]['pgs']
        if g_cw > 0 or g_pw > 0:
            sellers.append({'name':name,'g_cw':g_cw,'g_pw':g_pw,'g_ly':g_ly,'g_w2':g_w2,
                'u_cw':u_cw,'u_pw':u_pw,'u_ly':u_ly,'cats':cats,'pgs':pgs,
                'ytd':ytd_cw,'ytd_ly':ytd_ly,'ytd_u':ytd_u_cw,
                'mcid':d['seller_mcid'].get(name,''),
                'wow_d':g_cw-g_pw,'wow':pct(g_cw,g_pw),'yoy_d':g_cw-g_ly,'yoy':pct(g_cw,g_ly)})
    sellers_by_gms = sorted(sellers, key=lambda x: -x['g_cw'])
    min_mover = 100 if mp_code in ('AE','SA') else 300
    movers_up = sorted([s for s in sellers if s['g_cw']>=min_mover or s['g_pw']>=min_mover], key=lambda x:-x['wow_d'])[:15]
    movers_down = sorted([s for s in sellers if s['g_cw']>=min_mover or s['g_pw']>=min_mover], key=lambda x:x['wow_d'])[:15]

    # PG / Cat
    pgs_list = []
    for pg, wdata in d['pg'].items():
        gc=wdata[CW]['gms'];gp=wdata[PW]['gms'];gl=wdata[LY]['gms']
        uc=wdata[CW]['units'];up_=wdata[PW]['units'];ul=wdata[LY]['units']
        ytd_c=wdata[CW].get('ytd_gms',0);ytd_l=wdata[LY].get('ytd_gms',0)
        if gc>0 or gp>0: pgs_list.append({'pg':pg,'g_cw':gc,'g_pw':gp,'g_ly':gl,'u_cw':uc,'u_pw':up_,'u_ly':ul,'ytd':ytd_c,'ytd_ly':ytd_l})
    pgs_by_gms = sorted(pgs_list, key=lambda x:-x['g_cw'])

    cats_list = []
    for cat, wdata in d['cat'].items():
        gc=wdata[CW]['gms'];gp=wdata[PW]['gms'];gl=wdata[LY]['gms']
        uc=wdata[CW]['units'];up_=wdata[PW]['units'];ul=wdata[LY]['units']
        ytd_c=wdata[CW].get('ytd_gms',0);ytd_l=wdata[LY].get('ytd_gms',0)
        if gc>0 or gp>0: cats_list.append({'cat':cat,'g_cw':gc,'g_pw':gp,'g_ly':gl,'u_cw':uc,'u_pw':up_,'u_ly':ul,'ytd':ytd_c,'ytd_ly':ytd_l})
    cats_by_gms = sorted(cats_list, key=lambda x:-x['g_cw'])

    total_g = T_CW['gms']
    total_ytd = sum(s['ytd'] for s in sellers)
    top_seller = sellers_by_gms[0]['name'] if sellers_by_gms else 'N/A'
    top_seller_v = sellers_by_gms[0]['g_cw'] if sellers_by_gms else 0
    top_pg_name = pgs_by_gms[0]['pg'] if pgs_by_gms else 'N/A'
    top_pg_share = pgs_by_gms[0]['g_cw']/total_g*100 if pgs_by_gms and total_g else 0
    fba_ratio = T_CW['fba_gms']/T_CW['gms']*100 if T_CW['gms'] else 0

    out_file = os.path.join(latest_dir, 'W%d_WBR_%s_Pipeline.html' % (WEEK_NUM, mp_code))
    title = 'WBR - TW2%s Pipeline %s' % (mp_code, CW_LBL)
    subtitle = '%s &nbsp;|&nbsp; %s Marketplace (%d) &nbsp;|&nbsp; USD' % (CW_LBL, mp_name, mp_id)

    f = open(out_file, 'w', encoding='utf-8')
    # ── Head + CSS ─────────────────────────────────────────────
    f.write('<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n')
    f.write('<meta name="viewport" content="width=device-width, initial-scale=1.0">\n')
    f.write('<title>%s</title>\n' % title)
    f.write('<style>\n')
    f.write(':root{--blue:#4472C4;--bl:#D6E4F0;--g:#006100;--gb:#C6EFCE;--r:#9C0006;--rb:#FFC7CE;--gy:#F2F2F2;--dk:#1a1a2e;--bd:#dee2e6}\n')
    f.write('*{margin:0;padding:0;box-sizing:border-box}\n')
    f.write('body{font-family:"Segoe UI",system-ui,sans-serif;background:#f0f2f5;color:#333;line-height:1.5}\n')
    f.write('.ctn{max-width:1500px;margin:0 auto;padding:20px}\n')
    f.write('.hdr{background:linear-gradient(135deg,var(--dk),var(--blue));color:#fff;padding:30px 40px;border-radius:12px;margin-bottom:24px}\n')
    f.write('.hdr h1{font-size:24px;font-weight:700;margin-bottom:4px}.hdr p{opacity:.85;font-size:14px}\n')
    f.write('.tabs{display:flex;gap:4px;margin-bottom:20px;flex-wrap:wrap}\n')
    f.write('.tab{padding:10px 20px;background:#fff;border:1px solid var(--bd);border-radius:8px 8px 0 0;cursor:pointer;font-size:13px;font-weight:600;color:#666;transition:all .2s}\n')
    f.write('.tab:hover{color:var(--blue)}.tab.active{background:var(--blue);color:#fff;border-color:var(--blue)}\n')
    f.write('.pnl{display:none}.pnl.active{display:block}\n')
    f.write('.card{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:24px;margin-bottom:20px}\n')
    f.write('.card h2{font-size:16px;color:var(--blue);margin-bottom:16px;border-bottom:2px solid var(--bl);padding-bottom:8px}\n')
    f.write('.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px;margin-bottom:24px}\n')
    f.write('.kpi{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);border-left:4px solid var(--blue)}\n')
    f.write('.kpi .lb{font-size:12px;color:#888;text-transform:uppercase;letter-spacing:.5px}\n')
    f.write('.kpi .vl{font-size:28px;font-weight:700;margin:4px 0}.kpi .ch{font-size:13px;font-weight:600}\n')
    f.write('.pos{color:var(--g)}.neg{color:var(--r)}\n')
    f.write('.pos-bg{background:var(--gb);color:var(--g)}.neg-bg{background:var(--rb);color:var(--r)}\n')
    f.write('table{width:100%;border-collapse:collapse;font-size:13px}\n')
    f.write('th{background:var(--blue);color:#fff;padding:10px 12px;text-align:center;font-weight:600;white-space:nowrap}\n')
    f.write('td{padding:8px 12px;border-bottom:1px solid #eee}tr:nth-child(even){background:var(--gy)}\n')
    f.write('.tr{text-align:right}.tc{text-align:center}\n')
    f.write('.bd{display:inline-block;padding:2px 8px;border-radius:4px;font-weight:600;font-size:12px}\n')
    f.write('.at{max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px}\n')
    f.write('.mv-up{border-left:4px solid var(--g)}.mv-dn{border-left:4px solid var(--r)}\n')
    f.write('.mv-card{background:#fff;border-radius:10px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:12px}\n')
    f.write('.mv-card .nm{font-size:15px;font-weight:700;margin-bottom:4px}.mv-card .dt{font-size:13px;color:#555}\n')
    f.write('.bar-container{display:flex;align-items:center;gap:8px;margin:4px 0}\n')
    f.write('.bar{height:18px;border-radius:3px;min-width:2px}\n')
    f.write('.bar-cw{background:var(--blue)}.bar-pw{background:#B4C7E7}.bar-ly{background:#FFC000}\n')
    f.write('@media(max-width:768px){.kg{grid-template-columns:1fr}table{font-size:11px}th,td{padding:6px 8px}}\n')
    f.write('.cpbtn{display:inline-block;padding:6px 14px;background:var(--blue);color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;margin-bottom:12px;transition:all .2s}.cpbtn:hover{opacity:.85}\n')
    f.write('</style>\n</head>\n<body>\n<div class="ctn">\n')

    # ── Header ─────────────────────────────────────────────────
    f.write('<div class="hdr"><h1>WBR &mdash; TW2%s Pipeline Weekly Business Review</h1>\n' % mp_code)
    f.write('<p>%s</p></div>\n' % subtitle)

    # ── KPI Cards ──────────────────────────────────────────────
    f.write('<div class="kg">\n')
    for lbl, val, wow, yoy, is_m in [
        ('Total Ordered GMS (USD)', T_CW['gms'], wow_gms, yoy_gms, True),
        ('Total Ordered Units', T_CW['units'], wow_units, yoy_units, False),
    ]:
        ff = fm if is_m else fi
        f.write('<div class="kpi"><div class="lb">%s</div><div class="vl">%s</div>' % (lbl, ff(val)))
        f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wow), ar(wow), fp(wow)))
        f.write('<div class="ch %s">YoY: %s %s</div></div>\n' % (pc(yoy), ar(yoy), fp(yoy)))
    f.write('<div class="kpi"><div class="lb">Active Sellers</div><div class="vl">%d</div>' % T_CW['sellers'])
    f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(wow_s), ar(wow_s), fp(wow_s)))
    f.write('<div class="ch %s">YoY: %s %s</div></div>\n' % (pc(yoy_s), ar(yoy_s), fp(yoy_s)))
    f.write('<div class="kpi"><div class="lb">Top Seller</div>')
    f.write('<div class="vl" style="font-size:16px">%s</div><div class="ch">%s</div></div>\n' % (esc(top_seller), fm(top_seller_v)))
    ytd_yoy_gms = pct(T_CW.get('ytd_gms',0), T_LY.get('ytd_gms',0))
    ytd_yoy_units = pct(T_CW.get('ytd_units',0), T_LY.get('ytd_units',0))
    f.write('<div class="kpi"><div class="lb">YTD GMS (USD)</div><div class="vl">%s</div>' % fm(T_CW.get('ytd_gms',0)))
    f.write('<div class="ch %s">YoY: %s %s</div></div>\n' % (pc(ytd_yoy_gms), ar(ytd_yoy_gms), fp(ytd_yoy_gms)))
    f.write('</div>\n')

    # ── Executive Summary ──────────────────────────────────────
    top3_up = movers_up[:3]; top3_dn = movers_down[:3]
    gainers_str = ', '.join(['%s (%s)' % (esc(s['name']), fp(s['wow'])) for s in top3_up]) or 'N/A'
    decliners_str = ', '.join(['%s (%s)' % (esc(s['name']), fp(s['wow'])) for s in top3_dn]) or 'N/A'

    f.write('<div class="card" style="border-left:4px solid var(--blue)">\n')
    f.write('<h2>&#128221; Executive Summary</h2>\n')
    f.write('<div style="font-size:14px;line-height:1.8">\n')

    wow_dir = 'down' if wow_gms < 0 else 'up'
    yoy_dir = 'down' if yoy_gms < 0 else 'up'
    f.write('<p><strong>Overall Performance:</strong> %s total ordered GMS came in at <strong>%s</strong>, ' % (CW_LBL, fm(T_CW['gms'])))
    f.write('%s <strong>%s WoW</strong> from %s (%s) and ' % (wow_dir, fp(wow_gms), fm(T_PW['gms']), PW_LBL))
    f.write('%s <strong>%s YoY</strong> vs %s (%s). ' % (yoy_dir, fp(yoy_gms), fm(T_LY['gms']), LY_LBL))
    f.write('Units: <strong>%s</strong> (WoW %s, YoY %s). ' % (fi(T_CW['units']), fp(wow_units), fp(yoy_units)))
    f.write('FBA: <strong>%.1f%%</strong> of GMS. ' % fba_ratio)
    ytd_g_cw = T_CW.get('ytd_gms',0); ytd_g_ly = T_LY.get('ytd_gms',0)
    ytd_u_cw = T_CW.get('ytd_units',0); ytd_u_ly = T_LY.get('ytd_units',0)
    ytd_g_yoy = pct(ytd_g_cw, ytd_g_ly); ytd_u_yoy = pct(ytd_u_cw, ytd_u_ly)
    f.write('YTD GMS: <strong>%s</strong> (YoY %s), YTD Units: <strong>%s</strong> (YoY %s).</p>\n' % (fm(ytd_g_cw), fp(ytd_g_yoy), fi(ytd_u_cw), fp(ytd_u_yoy)))

    f.write('<p><strong>Seller Landscape:</strong> <strong>%d</strong> active sellers (WoW %s, YoY %s). ' % (T_CW['sellers'], fp(wow_s), fp(yoy_s)))
    nsr_s = d['cohort_totals']['NSR'].get(CW, {'sellers':0}).get('sellers', 0)
    esm_s = d['cohort_totals']['ESM'].get(CW, {'sellers':0}).get('sellers', 0)
    f.write('NSR: <strong>%d</strong>, ESM: <strong>%d</strong>. ' % (nsr_s, esm_s))
    f.write('Top seller: <strong>%s</strong> — %s (%.1f%% share).</p>\n' % (esc(top_seller), fm(top_seller_v), top_seller_v/total_g*100 if total_g else 0))

    # DSR seller count (CW vs LY) — based on ytd_launch=1 & channel=DSR
    dsr_cw_cnt = 0; dsr_ly_cnt = 0
    for sname, wdata in d['seller'].items():
        ch = d['seller_channel'].get(sname, '')
        if ch == 'DSR':
            if wdata[CW].get('ytd_launch', 0): dsr_cw_cnt += 1
            if wdata[LY].get('ytd_launch', 0): dsr_ly_cnt += 1
    dsr_yoy = pct(dsr_cw_cnt, dsr_ly_cnt)
    dsr_dir = ar(dsr_yoy) + ' ' + fp(dsr_yoy)
    f.write('<p><strong>DSR Pipeline:</strong> <strong>%d</strong> active DSR sellers in %s vs <strong>%d</strong> in %s (<span class="%s">%s</span>).</p>\n' % (dsr_cw_cnt, CW_LBL, dsr_ly_cnt, LY_LBL, pc(dsr_yoy), dsr_dir))

    f.write('<p><strong>Top Product Group:</strong> <strong>%s</strong> — %.1f%% of GMS.</p>\n' % (esc(top_pg_name), top_pg_share))

    f.write('<p><strong>Movers &amp; Shakers:</strong> Gainers: %s. Decliners: %s.</p>\n' % (gainers_str, decliners_str))

    f.write('<p><strong>Key Callouts:</strong></p><ul style="margin:4px 0 0 20px;font-size:13px">\n')
    if wow_gms < -0.05:
        f.write('<li>GMS declined %s WoW — monitor for seasonal vs structural causes.</li>\n' % fp(wow_gms))
    elif wow_gms > 0.05:
        f.write('<li>Strong GMS growth of %s WoW — identify drivers to sustain.</li>\n' % fp(wow_gms))
    if yoy_gms > 0:
        f.write('<li>Positive YoY (%s) — healthy pipeline growth vs last year.</li>\n' % fp(yoy_gms))
    else:
        f.write('<li>YoY GMS %s — investigate root causes.</li>\n' % fp(yoy_gms))
    for s in movers_up[:2]:
        if s['wow_d'] > (500 if mp_code in ('AE','SA') else 1000):
            f.write('<li><strong>%s</strong> surged %s WoW (+%s).</li>\n' % (esc(s['name']), fp(s['wow']), fm(s['wow_d'])))
    for s in movers_down[:2]:
        if s['wow_d'] < (-500 if mp_code in ('AE','SA') else -1000):
            f.write('<li><strong>%s</strong> dropped %s WoW (%s).</li>\n' % (esc(s['name']), fp(s['wow']), fm(s['wow_d'])))
    f.write('</ul></div></div>\n')

    # ── Tabs ───────────────────────────────────────────────────
    f.write('<div class="tabs">\n')
    for i, t in enumerate(['Summary', 'Category / PG', 'Top Sellers', 'Movers &amp; Shakers', 'Seller Deep Dive', 'Cohort (NSR vs ESM)', 'All Sellers (DSR/ESM)', 'DSR Launches']):
        f.write('<div class="tab%s" onclick="showTab(%d)">%s</div>\n' % (' active' if i==0 else '', i, t))
    f.write('</div>\n')

    # ── Panel 0: Summary ──────────────────────────────────────
    f.write('<div class="pnl active" id="p0"><div class="card">\n')
    f.write('<h2>Weekly Summary (%s)</h2>\n' % mp_name)
    f.write('<table><thead><tr>')
    for h in ['Metric', CW_LBL, PW_LBL, 'WoW Delta', 'WoW %', LY_LBL, 'YoY Delta', 'YoY %']:
        f.write('<th>%s</th>' % h)
    f.write('</tr></thead><tbody>\n')
    for nm, vc, vp, vl, is_m in [
        ('Ordered GMS (USD)', T_CW['gms'], T_PW['gms'], T_LY['gms'], True),
        ('Ordered Units', T_CW['units'], T_PW['units'], T_LY['units'], False),
        ('FBA GMS (USD)', T_CW['fba_gms'], T_PW['fba_gms'], T_LY['fba_gms'], True),
        ('FBA Units', T_CW['fba_units'], T_PW['fba_units'], T_LY['fba_units'], False),
        ('Active Sellers', T_CW['sellers'], T_PW['sellers'], T_LY['sellers'], False),
    ]:
        ff = fm if is_m else fi
        wd=vc-vp; wp=pct(vc,vp); yd=vc-vl; yp=pct(vc,vl)
        f.write('<tr><td><strong>%s</strong></td>' % nm)
        f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td>' % (ff(vc),ff(vp),ff(wd),badge(wp)))
        f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td></tr>\n' % (ff(vl),ff(yd),badge(yp)))
    # YTD rows
    ytd_gms_cw = T_CW.get('ytd_gms',0); ytd_gms_ly = T_LY.get('ytd_gms',0)
    ytd_units_cw = T_CW.get('ytd_units',0); ytd_units_ly = T_LY.get('ytd_units',0)
    for nm, vc, vl, is_m in [
        ('YTD GMS (USD)', ytd_gms_cw, ytd_gms_ly, True),
        ('YTD Units', ytd_units_cw, ytd_units_ly, False),
    ]:
        ff = fm if is_m else fi
        yd = vc - vl; yp = pct(vc, vl)
        f.write('<tr style="background:#E8F0FE"><td><strong>%s</strong></td>' % nm)
        f.write('<td class="tr">%s</td><td class="tr">&mdash;</td><td class="tr">&mdash;</td><td class="tc">&mdash;</td>' % ff(vc))
        f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td></tr>\n' % (ff(vl), ff(yd), badge(yp)))
    f.write('</tbody></table></div></div>\n')

    # ── Panel 1: Category / PG ────────────────────────────────
    f.write('<div class="pnl" id="p1">\n')
    for title_t, data_list, key_name in [
        ('GMS by Account Primary Category', cats_by_gms[:20], 'cat'),
        ('GMS by SP Primary Product Group', pgs_by_gms[:15], 'pg'),
    ]:
        f.write('<div class="card"><h2>%s</h2>\n<table><thead><tr>' % title_t)
        for h in [key_name.upper() if key_name=='pg' else 'Category', CW_LBL, PW_LBL, 'WoW %', LY_LBL, 'YoY %', 'Units', 'Share', 'YTD GMS', 'YTD YoY']:
            f.write('<th>%s</th>' % h)
        f.write('</tr></thead><tbody>\n')
        for item in data_list:
            wp=pct(item['g_cw'],item['g_pw']); yp=pct(item['g_cw'],item['g_ly'])
            share=item['g_cw']/total_g if total_g else 0
            ytd_yoy=pct(item.get('ytd',0),item.get('ytd_ly',0))
            f.write('<tr><td><strong>%s</strong></td>' % esc(item[key_name]))
            f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td>' % (fm(item['g_cw']),fm(item['g_pw']),badge(wp)))
            f.write('<td class="tr">%s</td><td class="tc">%s</td>' % (fm(item['g_ly']),badge(yp)))
            f.write('<td class="tr">%s</td><td class="tc">%.1f%%</td>' % (fi(item['u_cw']),share*100))
            f.write('<td class="tr">%s</td><td class="tc">%s</td></tr>\n' % (fm(item.get('ytd',0)),badge(ytd_yoy)))
        f.write('</tbody></table></div>\n')
    f.write('</div>\n')

    # ── Panel 2: Top Sellers ──────────────────────────────────
    f.write('<div class="pnl" id="p2"><div class="card">\n')
    f.write('<h2>Top 30 Sellers by %s GMS</h2>\n' % CW_LBL)
    _mcids_p2 = ','.join(s.get('mcid','') for s in sellers_by_gms[:30] if s.get('mcid',''))
    f.write('<button class="cpbtn" data-mcids="%s" onclick="copyMcids(this)">&#128203; Copy MCIDs (%d)</button>\n' % (_mcids_p2, len([s for s in sellers_by_gms[:30] if s.get('mcid','')])))
    f.write('<table><thead><tr>')
    for h in ['#','Seller','MCID',CW_LBL,PW_LBL,'WoW %',LY_LBL,'YoY %','Units','YTD GMS','YTD YoY','YTD Share','Category','Wk Share']:
        f.write('<th>%s</th>' % h)
    f.write('</tr></thead><tbody>\n')
    for i, s in enumerate(sellers_by_gms[:30]):
        cats_str = ', '.join(list(s['cats']-{'Other'})[:2]) or 'Other'
        share = s['g_cw']/total_g if total_g else 0
        ytd_share = s['ytd']/total_ytd if total_ytd else 0
        ytd_yoy = pct(s['ytd'], s['ytd_ly'])
        f.write('<tr><td class="tc">%d</td><td><strong>%s</strong></td><td class="tc" style="font-size:11px">%s</td>' % (i+1, esc(s['name']), s.get('mcid','')))
        f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td>' % (fm(s['g_cw']),fm(s['g_pw']),badge(s['wow'])))
        f.write('<td class="tr">%s</td><td class="tc">%s</td>' % (fm(s['g_ly']),badge(s['yoy'])))
        f.write('<td class="tr">%s</td>' % fi(s['u_cw']))
        f.write('<td class="tr">%s</td><td class="tc">%s</td><td class="tc">%.1f%%</td>' % (fm(s['ytd']), badge(ytd_yoy), ytd_share*100))
        f.write('<td class="at">%s</td><td class="tc">%.1f%%</td></tr>\n' % (esc(cats_str),share*100))
    f.write('</tbody></table></div></div>\n')

    # ── Panel 3: Movers & Shakers ─────────────────────────────
    f.write('<div class="pnl" id="p3">\n')
    for label, icon, mlist, color_var in [
        ('Top Gainers (Biggest WoW GMS Increase)', '&#128293;', movers_up[:10], '--g'),
        ('Top Decliners (Biggest WoW GMS Decrease)', '&#128308;', movers_down[:10], '--r'),
    ]:
        _mcids_mv = ','.join(s.get('mcid','') for s in mlist if s.get('mcid',''))
        f.write('<div class="card"><h2>%s %s</h2>\n' % (icon, label))
        f.write('<button class="cpbtn" data-mcids="%s" onclick="copyMcids(this)">&#128203; Copy MCIDs (%d)</button>\n' % (_mcids_mv, len([s for s in mlist if s.get('mcid','')])))
        f.write('<table><thead><tr>')
        for h in ['#','Seller','MCID',CW_LBL,PW_LBL,'WoW Delta','WoW %',LY_LBL,'YoY %','YTD GMS','YTD YoY','Category']:
            f.write('<th>%s</th>' % h)
        f.write('</tr></thead><tbody>\n')
        for i, s in enumerate(mlist):
            cats_str = ', '.join(list(s['cats']-{'Other'})[:2]) or 'Other'
            ytd_yoy = pct(s.get('ytd',0), s.get('ytd_ly',0))
            f.write('<tr><td class="tc">%d</td><td><strong>%s</strong></td><td class="tc" style="font-size:11px">%s</td>' % (i+1, esc(s['name']), s.get('mcid','')))
            f.write('<td class="tr">%s</td><td class="tr">%s</td>' % (fm(s['g_cw']),fm(s['g_pw'])))
            f.write('<td class="tr" style="color:var(%s);font-weight:700">%s</td>' % (color_var, fm(s['wow_d'])))
            f.write('<td class="tc">%s</td>' % badge(s['wow']))
            f.write('<td class="tr">%s</td><td class="tc">%s</td>' % (fm(s['g_ly']),badge(s['yoy'])))
            f.write('<td class="tr">%s</td><td class="tc">%s</td>' % (fm(s.get('ytd',0)), badge(ytd_yoy)))
            f.write('<td class="at">%s</td></tr>\n' % esc(cats_str))
        f.write('</tbody></table></div>\n')

    # Visual bars
    all_mv = movers_up[:5] + movers_down[:5]
    max_gms = max((max(s['g_cw'],s['g_pw'],s['g_ly']) for s in all_mv), default=1) or 1
    f.write('<div class="card"><h2>Movers &amp; Shakers Visual</h2>\n')
    f.write('<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">\n')
    for side_label, side_list, side_cls in [('&#9650; Top Gainers', movers_up[:7], 'mv-up'), ('&#9660; Top Decliners', movers_down[:7], 'mv-dn')]:
        color_h = 'var(--g)' if 'up' in side_cls else 'var(--r)'
        f.write('<div><h3 style="color:%s;margin-bottom:12px;font-size:14px">%s</h3>\n' % (color_h, side_label))
        for s in side_list:
            f.write('<div class="mv-card %s"><div class="nm">%s</div>\n' % (side_cls, esc(s['name'])))
            for lbl, val, cls in [(CW_LBL,s['g_cw'],'bar-cw'),(PW_LBL,s['g_pw'],'bar-pw'),(LY_LBL,s['g_ly'],'bar-ly')]:
                w = val/max_gms*100 if max_gms else 0
                f.write('<div class="bar-container"><span style="width:55px;font-size:10px">%s</span><div class="bar %s" style="width:%.1f%%"></div><span style="font-size:11px">%s</span></div>\n' % (lbl,cls,w,fm(val)))
            f.write('<div class="dt">WoW: %s | YoY: %s</div></div>\n' % (fp(s['wow']),fp(s['yoy'])))
        f.write('</div>\n')
    f.write('</div></div>\n')
    f.write('</div>\n')  # panel 3

    # ── Panel 4: Seller Deep Dive ─────────────────────────────
    f.write('<div class="pnl" id="p4">\n')
    for s in sellers_by_gms[:10]:
        wp=s['wow']; yp=s['yoy']
        cats_str = ', '.join(list(s['cats']-{'Other'})[:3]) or 'Other'
        pgs_str = ', '.join(list(s['pgs']-{'Other'})[:3]) or 'Other'
        s_mcid = s.get('mcid','')
        f.write('<div class="card"><h2>%s <span style="font-size:12px;color:#888;font-weight:400">MCID: %s</span></h2>\n<div class="kg" style="margin-bottom:16px">\n' % (esc(s['name']), s_mcid))
        # KPI cards
        f.write('<div class="kpi"><div class="lb">%s GMS</div><div class="vl" style="font-size:22px">%s</div>' % (CW_LBL, fm(s['g_cw'])))
        f.write('<div class="ch %s">WoW: %s %s</div><div class="ch %s">YoY: %s %s</div></div>\n' % (pc(wp),ar(wp),fp(wp),pc(yp),ar(yp),fp(yp)))
        f.write('<div class="kpi"><div class="lb">%s GMS</div><div class="vl" style="font-size:22px">%s</div></div>\n' % (PW_LBL, fm(s['g_pw'])))
        f.write('<div class="kpi"><div class="lb">%s GMS</div><div class="vl" style="font-size:22px">%s</div></div>\n' % (LY_LBL, fm(s['g_ly'])))
        wu=pct(s['u_cw'],s['u_pw']); yu=pct(s['u_cw'],s['u_ly'])
        f.write('<div class="kpi"><div class="lb">%s Units</div><div class="vl" style="font-size:22px">%s</div>' % (CW_LBL, fi(s['u_cw'])))
        f.write('<div class="ch %s">WoW: %s %s</div><div class="ch %s">YoY: %s %s</div></div>\n' % (pc(wu),ar(wu),fp(wu),pc(yu),ar(yu),fp(yu)))
        f.write('</div>\n')  # kg
        f.write('<div style="display:flex;gap:20px;font-size:13px;color:#555">')
        f.write('<div><strong>Category:</strong> %s</div><div><strong>PG:</strong> %s</div>' % (esc(cats_str),esc(pgs_str)))
        f.write('<div><strong>Wk Share:</strong> %.1f%%</div>' % (s['g_cw']/total_g*100 if total_g else 0))
        f.write('<div><strong>YTD GMS:</strong> %s</div>' % fm(s['ytd']))
        ytd_sh = s['ytd']/total_ytd*100 if total_ytd else 0
        s_ytd_yoy = pct(s['ytd'], s['ytd_ly'])
        f.write('<div><strong>YTD YoY:</strong> <span class="%s">%s %s</span></div>' % (pc(s_ytd_yoy), ar(s_ytd_yoy), fp(s_ytd_yoy)))
        f.write('<div><strong>YTD Share:</strong> %.1f%%</div></div>\n' % ytd_sh)
        # Trend bars
        max_v = max(s['g_cw'],s['g_pw'],s['g_ly'],s['g_w2']) or 1
        f.write('<div style="margin-top:12px"><h3 style="font-size:13px;color:#888;margin-bottom:6px">GMS Comparison</h3>')
        for lbl, val, color in [(W2_LBL,s['g_w2'],'#B4C7E7'),(PW_LBL,s['g_pw'],'#7BA0D4'),(CW_LBL,s['g_cw'],'var(--blue)'),(LY_LBL,s['g_ly'],'#FFC000')]:
            w = val/max_v*100 if max_v else 0
            f.write('<div class="bar-container"><span style="width:55px;font-size:11px;font-weight:600">%s</span><div class="bar" style="width:%.1f%%;background:%s"></div><span style="font-size:11px">%s</span></div>\n' % (lbl,w,color,fm(val)))
        f.write('</div></div>\n')
    f.write('</div>\n')  # panel 4

    # ── Panel 5: Cohort Analysis (NSR vs ESM) ─────────────────
    f.write('<div class="pnl" id="p5">\n')

    COHORTS = ['NSR', 'ESM']
    cohort_empty = {'gms':0,'units':0,'sellers':0,'fba_gms':0}

    # Cohort summary table
    f.write('<div class="card"><h2>Cohort Overview: NSR (DSR+SSR) vs ESM</h2>\n')
    f.write('<table><thead><tr>')
    for h in ['Cohort', CW_LBL+' GMS', PW_LBL+' GMS', 'WoW %', LY_LBL+' GMS', 'YoY %', CW_LBL+' Units', 'Sellers', 'FBA %', 'YTD GMS', 'YTD YoY', 'YTD Share', 'Wk Share']:
        f.write('<th>%s</th>' % h)
    f.write('</tr></thead><tbody>\n')
    for coh in COHORTS:
        ct_cw = d['cohort_totals'][coh].get(CW, cohort_empty)
        ct_pw = d['cohort_totals'][coh].get(PW, cohort_empty)
        ct_ly = d['cohort_totals'][coh].get(LY, cohort_empty)
        c_wow = pct(ct_cw['gms'], ct_pw['gms'])
        c_yoy = pct(ct_cw['gms'], ct_ly['gms'])
        c_fba = ct_cw['fba_gms']/ct_cw['gms']*100 if ct_cw['gms'] else 0
        c_share = ct_cw['gms']/total_g*100 if total_g else 0
        c_ytd = ct_cw.get('ytd_gms', 0)
        c_ytd_ly = ct_ly.get('ytd_gms', 0)
        c_ytd_yoy = pct(c_ytd, c_ytd_ly)
        # YTD share: cohort YTD / sum of all cohort YTDs
        all_coh_ytd = sum(d['cohort_totals'][c].get(CW, cohort_empty).get('ytd_gms',0) for c in COHORTS)
        c_ytd_share = c_ytd/all_coh_ytd*100 if all_coh_ytd else 0
        f.write('<tr><td><strong>%s</strong></td>' % coh)
        f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td>' % (fm(ct_cw['gms']),fm(ct_pw['gms']),badge(c_wow)))
        f.write('<td class="tr">%s</td><td class="tc">%s</td>' % (fm(ct_ly['gms']),badge(c_yoy)))
        f.write('<td class="tr">%s</td><td class="tc">%d</td>' % (fi(ct_cw['units']),ct_cw.get('sellers',0)))
        f.write('<td class="tc">%.1f%%</td>' % c_fba)
        f.write('<td class="tr">%s</td><td class="tc">%s</td><td class="tc">%.1f%%</td>' % (fm(c_ytd), badge(c_ytd_yoy), c_ytd_share))
        f.write('<td class="tc">%.1f%%</td></tr>\n' % c_share)
    f.write('</tbody></table></div>\n')

    # Per-cohort top sellers
    for coh in COHORTS:
        coh_sellers = []
        for sname, wdata in d['cohort_seller'][coh].items():
            gc=wdata[CW]['gms']; gp=wdata[PW]['gms']; gl=wdata[LY]['gms']
            uc=wdata[CW]['units']; up_=wdata[PW]['units']; ul=wdata[LY]['units']
            ytd_c=wdata[CW]['ytd_gms']; ytd_l=wdata[LY]['ytd_gms']
            cats=wdata[CW]['cats'] or wdata[PW]['cats']
            if gc > 0 or gp > 0:
                coh_sellers.append({'name':sname,'g_cw':gc,'g_pw':gp,'g_ly':gl,
                    'u_cw':uc,'u_pw':up_,'u_ly':ul,'cats':cats,
                    'ytd':ytd_c,'ytd_ly':ytd_l,
                    'mcid':d['seller_mcid'].get(sname,''),
                    'wow':pct(gc,gp),'yoy':pct(gc,gl),'wow_d':gc-gp})
        coh_sellers.sort(key=lambda x:-x['g_cw'])

        coh_total = sum(s['g_cw'] for s in coh_sellers)
        coh_ytd_total = sum(s['ytd'] for s in coh_sellers)
        coh_label = 'NSR (DSR + SSR)' if coh == 'NSR' else 'ESM'

        # Top sellers table
        _mcids_coh = ','.join(s.get('mcid','') for s in coh_sellers[:20] if s.get('mcid',''))
        f.write('<div class="card"><h2>%s — Top 20 Sellers</h2>\n' % coh_label)
        f.write('<button class="cpbtn" data-mcids="%s" onclick="copyMcids(this)">&#128203; Copy MCIDs (%d)</button>\n' % (_mcids_coh, len([s for s in coh_sellers[:20] if s.get('mcid','')])))
        f.write('<table><thead><tr>')
        for h in ['#','Seller','MCID',CW_LBL,PW_LBL,'WoW %',LY_LBL,'YoY %','Units','YTD GMS','YTD YoY','YTD Share','Category','Wk Share']:
            f.write('<th>%s</th>' % h)
        f.write('</tr></thead><tbody>\n')
        for i, s in enumerate(coh_sellers[:20]):
            cats_str = ', '.join(list(s['cats']-{'Other'})[:2]) or 'Other'
            share = s['g_cw']/coh_total*100 if coh_total else 0
            ytd_sh = s['ytd']/coh_ytd_total*100 if coh_ytd_total else 0
            ytd_yoy = pct(s['ytd'], s.get('ytd_ly',0))
            f.write('<tr><td class="tc">%d</td><td><strong>%s</strong></td><td class="tc" style="font-size:11px">%s</td>' % (i+1, esc(s['name']), s.get('mcid','')))
            f.write('<td class="tr">%s</td><td class="tr">%s</td><td class="tc">%s</td>' % (fm(s['g_cw']),fm(s['g_pw']),badge(s['wow'])))
            f.write('<td class="tr">%s</td><td class="tc">%s</td>' % (fm(s['g_ly']),badge(s['yoy'])))
            f.write('<td class="tr">%s</td>' % fi(s['u_cw']))
            f.write('<td class="tr">%s</td><td class="tc">%s</td><td class="tc">%.1f%%</td>' % (fm(s['ytd']), badge(ytd_yoy), ytd_sh))
            f.write('<td class="at">%s</td><td class="tc">%.1f%%</td></tr>\n' % (esc(cats_str),share))
        f.write('</tbody></table></div>\n')

        # Movers within cohort
        coh_min = 50 if mp_code in ('AE','SA') else 200
        coh_up = sorted([s for s in coh_sellers if s['g_cw']>=coh_min or s['g_pw']>=coh_min], key=lambda x:-x['wow_d'])[:5]
        coh_dn = sorted([s for s in coh_sellers if s['g_cw']>=coh_min or s['g_pw']>=coh_min], key=lambda x:x['wow_d'])[:5]

        f.write('<div class="card"><h2>%s — Movers &amp; Shakers</h2>\n' % coh_label)
        f.write('<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">\n')
        for side_lbl, side_list in [('&#9650; Gainers', coh_up), ('&#9660; Decliners', coh_dn)]:
            is_up = 'Gainers' in side_lbl
            f.write('<div><h3 style="color:var(%s);font-size:14px;margin-bottom:8px">%s</h3>\n' % ('--g' if is_up else '--r', side_lbl))
            f.write('<table style="font-size:12px"><thead><tr><th>Seller</th><th>%s</th><th>%s</th><th>Delta</th><th>WoW</th></tr></thead><tbody>\n' % (CW_LBL, PW_LBL))
            for s in side_list:
                clr = 'var(--g)' if s['wow_d']>=0 else 'var(--r)'
                f.write('<tr><td><strong>%s</strong></td><td class="tr">%s</td><td class="tr">%s</td>' % (esc(s['name']),fm(s['g_cw']),fm(s['g_pw'])))
                f.write('<td class="tr" style="color:%s;font-weight:700">%s</td><td class="tc">%s</td></tr>\n' % (clr,fm(s['wow_d']),badge(s['wow'])))
            f.write('</tbody></table></div>\n')
        f.write('</div></div>\n')

    f.write('</div>\n')  # panel 5

    # ── Panel 6: All Sellers (DSR/ESM) ────────────────────────
    f.write('<div class="pnl" id="p6">\n')

    # Build full seller list with channel info
    all_sellers_list = []
    for name, wdata in d['seller'].items():
        g_cw=wdata[CW]['gms']; g_pw=wdata[PW]['gms']; g_ly=wdata[LY]['gms']
        u_cw=wdata[CW]['units']; u_pw=wdata[PW]['units']; u_ly=wdata[LY]['units']
        ytd_c=wdata[CW]['ytd_gms']; ytd_l=wdata[LY]['ytd_gms']
        fba_c=wdata[CW]['fba_gms']
        cats=wdata[CW]['cats'] or wdata[PW]['cats']
        pgs=wdata[CW]['pgs'] or wdata[PW]['pgs']
        channel = d['seller_channel'].get(name, '')
        if g_cw > 0 or g_pw > 0:
            all_sellers_list.append({
                'name':name, 'channel':channel,
                'g_cw':g_cw, 'g_pw':g_pw, 'g_ly':g_ly,
                'u_cw':u_cw, 'u_pw':u_pw, 'u_ly':u_ly,
                'fba_gms':fba_c,
                'ytd':ytd_c, 'ytd_ly':ytd_l,
                'cats':cats, 'pgs':pgs,
                'mcid':d['seller_mcid'].get(name,''),
                'wow':pct(g_cw,g_pw), 'yoy':pct(g_cw,g_ly),
                'wow_d':g_cw-g_pw,
            })

    # Separate by channel group
    dsr_sellers = sorted([s for s in all_sellers_list if s['channel'] in ('DSR','SSR')], key=lambda x:-x['g_cw'])
    esm_sellers = sorted([s for s in all_sellers_list if s['channel'] == 'ESM'], key=lambda x:-x['g_cw'])

    for grp_label, grp_list, grp_tag in [
        ('DSR / SSR (NSR) — All Sellers', dsr_sellers, 'DSR'),
        ('ESM — All Sellers', esm_sellers, 'ESM'),
    ]:
        grp_total_gms = sum(s['g_cw'] for s in grp_list)
        grp_total_ytd = sum(s['ytd'] for s in grp_list)
        grp_total_ytd_ly = sum(s['ytd_ly'] for s in grp_list)
        grp_total_gms_pw = sum(s['g_pw'] for s in grp_list)
        grp_total_gms_ly = sum(s['g_ly'] for s in grp_list)
        grp_wow = pct(grp_total_gms, grp_total_gms_pw)
        grp_yoy = pct(grp_total_gms, grp_total_gms_ly)
        grp_ytd_yoy = pct(grp_total_ytd, grp_total_ytd_ly)

        # Summary KPI row
        f.write('<div class="card">\n')
        f.write('<h2>%s (%d sellers)</h2>\n' % (grp_label, len(grp_list)))
        f.write('<div class="kg" style="margin-bottom:16px">\n')
        f.write('<div class="kpi"><div class="lb">%s GMS</div><div class="vl">%s</div>' % (CW_LBL, fm(grp_total_gms)))
        f.write('<div class="ch %s">WoW: %s %s</div>' % (pc(grp_wow), ar(grp_wow), fp(grp_wow)))
        f.write('<div class="ch %s">YoY: %s %s</div></div>\n' % (pc(grp_yoy), ar(grp_yoy), fp(grp_yoy)))
        f.write('<div class="kpi"><div class="lb">YTD GMS</div><div class="vl">%s</div>' % fm(grp_total_ytd))
        f.write('<div class="ch %s">YoY: %s %s</div></div>\n' % (pc(grp_ytd_yoy), ar(grp_ytd_yoy), fp(grp_ytd_yoy)))
        f.write('<div class="kpi"><div class="lb">Mkt Share</div><div class="vl">%.1f%%</div></div>\n' % (grp_total_gms/total_g*100 if total_g else 0))
        f.write('</div>\n')

        # Full seller table
        _mcids_grp = ','.join(s.get('mcid','') for s in grp_list if s.get('mcid',''))
        f.write('<button class="cpbtn" data-mcids="%s" onclick="copyMcids(this)">&#128203; Copy MCIDs (%d)</button>\n' % (_mcids_grp, len([s for s in grp_list if s.get('mcid','')])))
        f.write('<table><thead><tr>')
        for h in ['#','Seller','MCID','Channel',CW_LBL,PW_LBL,'WoW Delta','WoW %',LY_LBL,'YoY %','Units','FBA GMS','YTD GMS','YTD YoY','Category','PG']:
            f.write('<th>%s</th>' % h)
        f.write('</tr></thead><tbody>\n')
        for i, s in enumerate(grp_list):
            cats_str = ', '.join(list(s['cats']-{'Other'})[:2]) or 'Other'
            pgs_str = ', '.join(list(s['pgs']-{'Other'})[:2]) or 'Other'
            ytd_yoy = pct(s['ytd'], s['ytd_ly'])
            f.write('<tr>')
            f.write('<td class="tc">%d</td>' % (i+1))
            f.write('<td><strong>%s</strong></td>' % esc(s['name']))
            f.write('<td class="tc" style="font-size:11px">%s</td>' % s.get('mcid',''))
            f.write('<td class="tc">%s</td>' % esc(s['channel']))
            f.write('<td class="tr">%s</td>' % fm(s['g_cw']))
            f.write('<td class="tr">%s</td>' % fm(s['g_pw']))
            f.write('<td class="tr" style="color:var(%s);font-weight:700">%s</td>' % ('--g' if s['wow_d']>=0 else '--r', fm(s['wow_d'])))
            f.write('<td class="tc">%s</td>' % badge(s['wow']))
            f.write('<td class="tr">%s</td>' % fm(s['g_ly']))
            f.write('<td class="tc">%s</td>' % badge(s['yoy']))
            f.write('<td class="tr">%s</td>' % fi(s['u_cw']))
            f.write('<td class="tr">%s</td>' % fm(s['fba_gms']))
            f.write('<td class="tr">%s</td>' % fm(s['ytd']))
            f.write('<td class="tc">%s</td>' % badge(ytd_yoy))
            f.write('<td class="at">%s</td>' % esc(cats_str))
            f.write('<td class="at">%s</td>' % esc(pgs_str))
            f.write('</tr>\n')
        f.write('</tbody></table>\n')
        f.write('</div>\n')

    f.write('</div>\n')  # panel 6

    # ── Panel 7: DSR New Launches (ytd_launch=1 & channel=DSR) ─
    f.write('<div class="pnl" id="p7">\n')

    import datetime as _dt

    new_launches = []
    for name, wdata in d['seller'].items():
        ch = d['seller_channel'].get(name, '')
        if ch != 'DSR':
            continue
        # Check ytd_launch flag in current week
        if not wdata[CW].get('ytd_launch', 0):
            continue
        g_cw=wdata[CW]['gms']; g_pw=wdata[PW]['gms']; g_ly=wdata[LY]['gms']
        u_cw=wdata[CW]['units']; u_pw=wdata[PW]['units']
        ytd_c=wdata[CW]['ytd_gms']; ytd_l=wdata[LY]['ytd_gms']
        fba_c=wdata[CW]['fba_gms']
        cats=wdata[CW]['cats'] or wdata[PW]['cats']
        pgs=wdata[CW]['pgs'] or wdata[PW]['pgs']
        ld = d['seller_launch_date'].get(name)
        ld_str = ''
        if ld:
            if isinstance(ld, _dt.datetime): ld_str = ld.strftime('%Y-%m-%d')
            elif isinstance(ld, _dt.date): ld_str = ld.strftime('%Y-%m-%d')
        new_launches.append({
            'name':name, 'launch_date_str':ld_str,
            'g_cw':g_cw, 'g_pw':g_pw, 'g_ly':g_ly,
            'u_cw':u_cw, 'u_pw':u_pw,
            'fba_gms':fba_c,
            'ytd':ytd_c, 'ytd_ly':ytd_l,
            'cats':cats, 'pgs':pgs,
            'mcid':d['seller_mcid'].get(name,''),
            'wow':pct(g_cw,g_pw), 'wow_d':g_cw-g_pw,
            'yoy':pct(g_cw,g_ly),
        })
    new_launches.sort(key=lambda x: x['launch_date_str'], reverse=True)

    nl_total_gms = sum(s['g_cw'] for s in new_launches)
    nl_total_ytd = sum(s['ytd'] for s in new_launches)
    nl_active = sum(1 for s in new_launches if s['g_cw'] > 0)

    f.write('<div class="card">\n')
    f.write('<h2>&#128640; DSR New Launches (YTD %d)</h2>\n' % CUR_YEAR)
    f.write('<p style="font-size:13px;color:#666;margin-bottom:16px">DSR sellers with <code>ytd_launch = 1</code> in %s. Total: <strong>%d</strong> sellers, <strong>%d</strong> with GMS this week.</p>\n' % (CW_LBL, len(new_launches), nl_active))

    # KPI summary
    f.write('<div class="kg" style="margin-bottom:16px">\n')
    f.write('<div class="kpi"><div class="lb">DSR YTD Launches</div><div class="vl">%d</div>' % len(new_launches))
    f.write('<div class="ch">%d active this week</div></div>\n' % nl_active)
    f.write('<div class="kpi"><div class="lb">%s GMS</div><div class="vl">%s</div>' % (CW_LBL, fm(nl_total_gms)))
    nl_share = nl_total_gms/total_g*100 if total_g else 0
    f.write('<div class="ch">%.1f%% of total GMS</div></div>\n' % nl_share)
    f.write('<div class="kpi"><div class="lb">YTD GMS</div><div class="vl">%s</div></div>\n' % fm(nl_total_ytd))
    f.write('</div>\n')

    # Full table
    if new_launches:
        _mcids_nl = ','.join(s.get('mcid','') for s in new_launches if s.get('mcid',''))
        f.write('<button class="cpbtn" data-mcids="%s" onclick="copyMcids(this)">&#128203; Copy MCIDs (%d)</button>\n' % (_mcids_nl, len([s for s in new_launches if s.get('mcid','')])))
        f.write('<table><thead><tr>')
        for h in ['#','Seller','MCID','Launch Date',CW_LBL,PW_LBL,'WoW Delta','WoW %','Units','FBA GMS','YTD GMS','YTD YoY','Category','PG']:
            f.write('<th>%s</th>' % h)
        f.write('</tr></thead><tbody>\n')
        for i, s in enumerate(new_launches):
            cats_str = ', '.join(list(s['cats']-{'Other'})[:2]) or 'Other'
            pgs_str = ', '.join(list(s['pgs']-{'Other'})[:2]) or 'Other'
            ytd_yoy = pct(s['ytd'], s['ytd_ly'])
            f.write('<tr>')
            f.write('<td class="tc">%d</td>' % (i+1))
            f.write('<td><strong>%s</strong></td>' % esc(s['name']))
            f.write('<td class="tc" style="font-size:11px">%s</td>' % s.get('mcid',''))
            f.write('<td class="tc">%s</td>' % s['launch_date_str'])
            f.write('<td class="tr">%s</td>' % fm(s['g_cw']))
            f.write('<td class="tr">%s</td>' % fm(s['g_pw']))
            f.write('<td class="tr" style="color:var(%s);font-weight:700">%s</td>' % ('--g' if s['wow_d']>=0 else '--r', fm(s['wow_d'])))
            f.write('<td class="tc">%s</td>' % badge(s['wow']))
            f.write('<td class="tr">%s</td>' % fi(s['u_cw']))
            f.write('<td class="tr">%s</td>' % fm(s['fba_gms']))
            f.write('<td class="tr">%s</td>' % fm(s['ytd']))
            f.write('<td class="tc">%s</td>' % badge(ytd_yoy))
            f.write('<td class="at">%s</td>' % esc(cats_str))
            f.write('<td class="at">%s</td>' % esc(pgs_str))
            f.write('</tr>\n')
        f.write('</tbody></table>\n')
    else:
        f.write('<p style="text-align:center;color:#888;padding:40px">No DSR launches found (ytd_launch = 1) in %s.</p>\n' % CW_LBL)

    f.write('</div>\n')
    f.write('</div>\n')  # panel 7

    # ── JS ─────────────────────────────────────────────────────
    f.write('<script>function showTab(i){document.querySelectorAll(".tab").forEach((t,j)=>t.classList.toggle("active",j===i));document.querySelectorAll(".pnl").forEach((p,j)=>p.classList.toggle("active",j===i));}\n')
    f.write('function copyMcids(btn){var t=btn.getAttribute("data-mcids").split(",").join("\\n");navigator.clipboard.writeText(t).then(function(){var o=btn.textContent;btn.textContent="Copied!";btn.style.background="var(--gb)";btn.style.color="var(--g)";setTimeout(function(){btn.textContent=o;btn.style.background="";btn.style.color="";},1500);});}\n')
    f.write('</script>\n')
    f.write('</div>\n</body>\n</html>\n')
    f.close()
    sz = os.path.getsize(out_file)
    print("  [%s] Saved: %s (%d bytes)" % (mp_code, out_file, sz))

# ══════════════════════════════════════════════════════════════
# Generate all marketplaces
# ══════════════════════════════════════════════════════════════
for mp in MARKETPLACES:
    print("Generating WBR for %s (%s)..." % (mp['name'], mp['code']))
    generate_wbr(mp['id'], mp['code'], mp['name'])

print("\nAll done!")
