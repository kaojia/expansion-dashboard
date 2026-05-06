#!/usr/bin/env python3
"""
gen_wbr.py – Generate WBR (Weekly Business Review) HTML reports from xlsx data.

Auto-detects the latest W## folder, reads the xlsx, and generates per-market
HTML reports under wbr/W{N}/.
"""

import os, re, math, html
from pathlib import Path
from datetime import datetime

import openpyxl

# ── Configuration ──────────────────────────────────────────────────────────
MARKETS = {
    111172: ("AU", "Australia Marketplace (111172)", "🇦🇺 AU"),
    338801: ("AE", "UAE Marketplace (338801)", "🇦🇪 UAE"),
    338811: ("SA", "Saudi Arabia Marketplace (338811)", "🇸🇦 SA"),
}

# Column indices (0-based)
COL_CAL_TYPE = 0
COL_YEAR = 1
COL_WEEK = 2
COL_ORIGIN = 3
COL_MKT = 10
COL_CHANNEL = 12
COL_MCID = 21
COL_NAME = 25
COL_LAUNCH_DATE = 28
COL_CATEGORY = 34
COL_PG = 36
COL_YTD_LAUNCH = 54
COL_ACTIVE = 63
COL_WTD_GMS = 90
COL_WTD_FBA_GMS = 91
COL_WTD_UNITS = 93
COL_WTD_FBA_UNITS = 94
COL_YTD_GMS = 96
COL_YTD_FBA_GMS = 97
COL_YTD_UNITS = 99


# ── Helpers ────────────────────────────────────────────────────────────────
def safe_float(v):
    """Convert a cell value to float, defaulting to 0."""
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def safe_int(v):
    """Convert a cell value to int, defaulting to 0."""
    return int(round(safe_float(v)))


def safe_str(v):
    """Convert a cell value to string."""
    if v is None:
        return ""
    return str(v).strip()


def fmt_money(v):
    """Format as $X,XXX (integer)."""
    iv = int(round(v))
    if iv < 0:
        return f"$-{abs(iv):,}"
    return f"${iv:,}"


def fmt_money_delta(v):
    """Format delta as $X,XXX or $-X,XXX."""
    iv = int(round(v))
    if iv < 0:
        return f"$-{abs(iv):,}"
    return f"${iv:,}"


def fmt_pct(v):
    """Format as +X.X% or -X.X%."""
    if v >= 0:
        return f"+{v:.1f}%"
    return f"{v:.1f}%"


def pct_change(cur, prev):
    """Calculate percentage change. Returns 0 if prev is 0."""
    if prev == 0:
        return 0.0
    return (cur - prev) / abs(prev) * 100


def badge(val, show_arrow=True):
    """Return a WoW/YoY badge span."""
    cls = "pos-bg" if val >= 0 else "neg-bg"
    arrow = "&#9650; " if val >= 0 else "&#9660; "
    if not show_arrow:
        arrow = ""
    return f'<span class="bd {cls}">{arrow}{fmt_pct(val)}</span>'


def arrow_text(val):
    """Return arrow text for KPI cards."""
    if val >= 0:
        return f'&#9650; {fmt_pct(val)}'
    return f'&#9660; {fmt_pct(val)}'


def pos_neg_class(val):
    """Return 'pos' or 'neg' CSS class."""
    return "pos" if val >= 0 else "neg"


def h(s):
    """HTML-escape a string."""
    return html.escape(str(s))


# ── Data Loading ───────────────────────────────────────────────────────────
def find_latest_week_folder():
    """Find the highest W## folder in the workspace root."""
    root = Path(".")
    week_dirs = []
    for d in root.iterdir():
        if d.is_dir():
            m = re.match(r"^W(\d+)$", d.name)
            if m:
                week_dirs.append((int(m.group(1)), d))
    if not week_dirs:
        raise FileNotFoundError("No W## folders found")
    week_dirs.sort(key=lambda x: x[0])
    return week_dirs[-1]  # (week_num, path)


def load_data(week_num, week_dir):
    """Load xlsx and return list of row dicts."""
    xlsx_name = f"WBR page 0 MCID data_weekly_w{week_num}_2026.xlsx"
    xlsx_path = week_dir / xlsx_name
    if not xlsx_path.exists():
        raise FileNotFoundError(f"File not found: {xlsx_path}")

    print(f"  Loading {xlsx_path} ...")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["raw"]

    rows = []
    header_skipped = False
    for row in ws.iter_rows(values_only=True):
        if not header_skipped:
            header_skipped = True
            continue
        rows.append(row)
    wb.close()
    print(f"  Loaded {len(rows)} rows")
    return rows


def filter_data(rows, marketplace_id):
    """Filter for seller_origin=TW and given marketplace_id. Return list of tuples."""
    filtered = []
    for r in rows:
        origin = safe_str(r[COL_ORIGIN])
        mkt = safe_int(r[COL_MKT]) if r[COL_MKT] is not None else 0
        if origin == "TW" and mkt == marketplace_id:
            filtered.append(r)
    return filtered


# ── Data Aggregation ───────────────────────────────────────────────────────
def build_datasets(rows, current_week):
    """
    Build aggregated datasets from filtered rows.
    Returns a dict with all needed data structures.
    """
    prev_week = current_week - 1

    # Separate rows by (year, week)
    cw26 = [r for r in rows if safe_int(r[COL_YEAR]) == 2026 and safe_int(r[COL_WEEK]) == current_week]
    pw26 = [r for r in rows if safe_int(r[COL_YEAR]) == 2026 and safe_int(r[COL_WEEK]) == prev_week]
    cw25 = [r for r in rows if safe_int(r[COL_YEAR]) == 2025 and safe_int(r[COL_WEEK]) == current_week]

    # ── Seller-level aggregation ──
    def agg_sellers(row_list):
        """Aggregate by MCID. Returns dict of MCID -> seller dict."""
        sellers = {}
        for r in row_list:
            mcid = safe_str(r[COL_MCID])
            if not mcid:
                continue
            if mcid not in sellers:
                sellers[mcid] = {
                    "mcid": mcid,
                    "name": safe_str(r[COL_NAME]),
                    "category": safe_str(r[COL_CATEGORY]),
                    "pg": safe_str(r[COL_PG]),
                    "channel": safe_str(r[COL_CHANNEL]),
                    "launch_date": r[COL_LAUNCH_DATE],
                    "wtd_gms": 0, "wtd_fba_gms": 0,
                    "wtd_units": 0, "wtd_fba_units": 0,
                    "ytd_gms": 0, "ytd_fba_gms": 0, "ytd_units": 0,
                    "active": 0, "ytd_launch": 0,
                }
            s = sellers[mcid]
            s["wtd_gms"] += safe_float(r[COL_WTD_GMS])
            s["wtd_fba_gms"] += safe_float(r[COL_WTD_FBA_GMS])
            s["wtd_units"] += safe_float(r[COL_WTD_UNITS])
            s["wtd_fba_units"] += safe_float(r[COL_WTD_FBA_UNITS])
            s["ytd_gms"] += safe_float(r[COL_YTD_GMS])
            s["ytd_fba_gms"] += safe_float(r[COL_YTD_FBA_GMS])
            s["ytd_units"] += safe_float(r[COL_YTD_UNITS])
            if safe_int(r[COL_ACTIVE]) == 1:
                s["active"] = 1
            if safe_int(r[COL_YTD_LAUNCH]) == 1:
                s["ytd_launch"] = 1
            # Keep the most informative name/category/channel
            name = safe_str(r[COL_NAME])
            if name and (not s["name"] or len(name) > len(s["name"])):
                s["name"] = name
            cat = safe_str(r[COL_CATEGORY])
            if cat and cat != "Other":
                s["category"] = cat
            pg = safe_str(r[COL_PG])
            if pg and pg != "Other":
                s["pg"] = pg
            ch = safe_str(r[COL_CHANNEL])
            if ch:
                s["channel"] = ch
            ld = r[COL_LAUNCH_DATE]
            if ld is not None:
                s["launch_date"] = ld
        return sellers

    cw26_sellers = agg_sellers(cw26)
    pw26_sellers = agg_sellers(pw26)
    cw25_sellers = agg_sellers(cw25)

    # ── Totals ──
    def totals(sellers):
        t = {"gms": 0, "fba_gms": 0, "units": 0, "fba_units": 0,
             "ytd_gms": 0, "ytd_fba_gms": 0, "ytd_units": 0, "active": 0}
        for s in sellers.values():
            t["gms"] += s["wtd_gms"]
            t["fba_gms"] += s["wtd_fba_gms"]
            t["units"] += s["wtd_units"]
            t["fba_units"] += s["wtd_fba_units"]
            t["ytd_gms"] += s["ytd_gms"]
            t["ytd_fba_gms"] += s["ytd_fba_gms"]
            t["ytd_units"] += s["ytd_units"]
            if s["active"]:
                t["active"] += 1
        return t

    t_cw = totals(cw26_sellers)
    t_pw = totals(pw26_sellers)
    t_ly = totals(cw25_sellers)

    # ── Category aggregation ──
    def agg_by_field(sellers_dict, field):
        groups = {}
        for s in sellers_dict.values():
            key = s[field] or "Other"
            if key not in groups:
                groups[key] = {"gms": 0, "units": 0, "ytd_gms": 0}
            groups[key]["gms"] += s["wtd_gms"]
            groups[key]["units"] += s["wtd_units"]
            groups[key]["ytd_gms"] += s["ytd_gms"]
        return groups

    cat_cw = agg_by_field(cw26_sellers, "category")
    cat_pw = agg_by_field(pw26_sellers, "category")
    cat_ly = agg_by_field(cw25_sellers, "category")

    pg_cw = agg_by_field(cw26_sellers, "pg")
    pg_pw = agg_by_field(pw26_sellers, "pg")
    pg_ly = agg_by_field(cw25_sellers, "pg")

    # ── Merged seller list (all MCIDs across all periods) ──
    all_mcids = set(cw26_sellers.keys()) | set(pw26_sellers.keys()) | set(cw25_sellers.keys())

    def get_seller(mcid):
        """Build a merged seller record."""
        cw = cw26_sellers.get(mcid, {})
        pw = pw26_sellers.get(mcid, {})
        ly = cw25_sellers.get(mcid, {})
        # Pick best name
        name = cw.get("name") or pw.get("name") or ly.get("name") or mcid
        cat = cw.get("category") or pw.get("category") or ly.get("category") or "Other"
        pg = cw.get("pg") or pw.get("pg") or ly.get("pg") or "Other"
        channel = cw.get("channel") or pw.get("channel") or ly.get("channel") or ""
        launch_date = cw.get("launch_date") or pw.get("launch_date") or ly.get("launch_date")
        ytd_launch = cw.get("ytd_launch", 0) or pw.get("ytd_launch", 0)
        active = cw.get("active", 0)

        gms_cw = cw.get("wtd_gms", 0)
        gms_pw = pw.get("wtd_gms", 0)
        gms_ly = ly.get("wtd_gms", 0)
        units_cw = cw.get("wtd_units", 0)
        units_pw = pw.get("wtd_units", 0)
        units_ly = ly.get("wtd_units", 0)
        fba_cw = cw.get("wtd_fba_gms", 0)
        fba_pw = pw.get("wtd_fba_gms", 0)
        fba_units_cw = cw.get("wtd_fba_units", 0)
        fba_units_pw = pw.get("wtd_fba_units", 0)
        ytd_gms = cw.get("ytd_gms", 0)
        ytd_gms_ly = ly.get("ytd_gms", 0)
        ytd_units = cw.get("ytd_units", 0)
        ytd_units_ly = ly.get("ytd_units", 0)

        wow_delta = gms_cw - gms_pw
        wow_pct = pct_change(gms_cw, gms_pw)
        yoy_pct = pct_change(gms_cw, gms_ly)
        ytd_yoy = pct_change(ytd_gms, ytd_gms_ly)

        return {
            "mcid": mcid, "name": name, "category": cat, "pg": pg,
            "channel": channel, "launch_date": launch_date,
            "ytd_launch": ytd_launch, "active": active,
            "gms_cw": gms_cw, "gms_pw": gms_pw, "gms_ly": gms_ly,
            "units_cw": units_cw, "units_pw": units_pw, "units_ly": units_ly,
            "fba_cw": fba_cw, "fba_pw": fba_pw,
            "fba_units_cw": fba_units_cw, "fba_units_pw": fba_units_pw,
            "ytd_gms": ytd_gms, "ytd_gms_ly": ytd_gms_ly,
            "ytd_units": ytd_units, "ytd_units_ly": ytd_units_ly,
            "wow_delta": wow_delta, "wow_pct": wow_pct,
            "yoy_pct": yoy_pct, "ytd_yoy": ytd_yoy,
        }

    sellers = [get_seller(m) for m in all_mcids]

    return {
        "cw26_sellers": cw26_sellers, "pw26_sellers": pw26_sellers,
        "cw25_sellers": cw25_sellers,
        "t_cw": t_cw, "t_pw": t_pw, "t_ly": t_ly,
        "cat_cw": cat_cw, "cat_pw": cat_pw, "cat_ly": cat_ly,
        "pg_cw": pg_cw, "pg_pw": pg_pw, "pg_ly": pg_ly,
        "sellers": sellers,
        "current_week": current_week, "prev_week": prev_week,
        "pw2_sellers": {},  # Will be populated if previous xlsx is available
    }


# ── HTML Generation ────────────────────────────────────────────────────────
def generate_html(data, market_code, market_label, week_num):
    """Generate the complete HTML report string."""
    d = data
    t_cw, t_pw, t_ly = d["t_cw"], d["t_pw"], d["t_ly"]
    sellers = d["sellers"]
    cw = week_num
    pw = week_num - 1

    total_gms = t_cw["gms"]
    total_ytd = t_cw["ytd_gms"]

    # ── Derived metrics ──
    wow_gms = pct_change(t_cw["gms"], t_pw["gms"])
    yoy_gms = pct_change(t_cw["gms"], t_ly["gms"])
    wow_units = pct_change(t_cw["units"], t_pw["units"])
    yoy_units = pct_change(t_cw["units"], t_ly["units"])
    wow_active = pct_change(t_cw["active"], t_pw["active"])
    yoy_active = pct_change(t_cw["active"], t_ly["active"])
    wow_fba_gms = pct_change(t_cw["fba_gms"], t_pw["fba_gms"])
    yoy_fba_gms = pct_change(t_cw["fba_gms"], t_ly["fba_gms"])
    wow_fba_units = pct_change(t_cw["fba_units"], t_pw["fba_units"])
    yoy_fba_units = pct_change(t_cw["fba_units"], t_ly["fba_units"])
    ytd_yoy_gms = pct_change(t_cw["ytd_gms"], t_ly["ytd_gms"])
    ytd_yoy_units = pct_change(t_cw["ytd_units"], t_ly["ytd_units"])
    fba_pct = (t_cw["fba_gms"] / t_cw["gms"] * 100) if t_cw["gms"] else 0

    # Top seller by weekly GMS
    top_seller = max(sellers, key=lambda s: s["gms_cw"]) if sellers else None

    # Top 30 by YTD GMS
    top30 = sorted(sellers, key=lambda s: s["ytd_gms"], reverse=True)[:30]

    # Top 5 by YTD GMS for deep dive
    top5 = sorted(sellers, key=lambda s: s["ytd_gms"], reverse=True)[:5]

    # Movers: need sellers active in both CW and PW (or at least one)
    active_sellers = [s for s in sellers if s["gms_cw"] != 0 or s["gms_pw"] != 0]
    gainers = sorted(active_sellers, key=lambda s: s["wow_delta"], reverse=True)[:10]
    # Filter to only positive deltas
    gainers = [s for s in gainers if s["wow_delta"] > 0][:10]
    decliners = sorted(active_sellers, key=lambda s: s["wow_delta"])[:10]
    decliners = [s for s in decliners if s["wow_delta"] < 0][:10]

    # NSR = DSR + SSR
    nsr_sellers = [s for s in sellers if s["channel"] in ("DSR", "SSR")]
    esm_sellers = [s for s in sellers if s["channel"] == "ESM"]

    # NSR/ESM totals
    def cohort_totals(slist):
        t = {"gms": 0, "units": 0, "fba_gms": 0, "ytd_gms": 0, "active": 0,
             "gms_pw": 0, "gms_ly": 0, "ytd_gms_ly": 0}
        for s in slist:
            t["gms"] += s["gms_cw"]
            t["units"] += s["units_cw"]
            t["fba_gms"] += s["fba_cw"]
            t["ytd_gms"] += s["ytd_gms"]
            t["active"] += 1 if s["active"] else 0
            t["gms_pw"] += s["gms_pw"]
            t["gms_ly"] += s["gms_ly"]
            t["ytd_gms_ly"] += s["ytd_gms_ly"]
        return t

    nsr_t = cohort_totals(nsr_sellers)
    esm_t = cohort_totals(esm_sellers)

    # Fix YTD YoY for cohorts: use each cohort's original 2025 channel
    # to avoid misattribution when sellers change channel between years.
    # Compute ytd_gms_ly directly from 2025 data grouped by 2025 channel.
    cw25_sellers = d["cw25_sellers"]
    esm_ytd_ly_direct = sum(s["ytd_gms"] for s in cw25_sellers.values()
                            if s["channel"] == "ESM")
    nsr_ytd_ly_direct = sum(s["ytd_gms"] for s in cw25_sellers.values()
                            if s["channel"] in ("DSR", "SSR"))
    esm_t["ytd_gms_ly"] = esm_ytd_ly_direct
    nsr_t["ytd_gms_ly"] = nsr_ytd_ly_direct

    # DSR launches (ytd_launch=1, NSR channel only)
    dsr_launches = [s for s in sellers if s["ytd_launch"] == 1 and s["channel"] in ("DSR", "SSR")]
    dsr_launches.sort(key=lambda s: _launch_sort_key(s), reverse=True)

    # All DSR/SSR sellers for tab 6 (all with ytd_launch=1 for DSR section)
    nsr_all = sorted(nsr_sellers, key=lambda s: s["gms_cw"], reverse=True)
    esm_all = sorted(esm_sellers, key=lambda s: s["gms_cw"], reverse=True)

    # NSR top 20 by YTD GMS
    nsr_top20 = sorted(nsr_sellers, key=lambda s: s["ytd_gms"], reverse=True)[:20]
    esm_top20 = sorted(esm_sellers, key=lambda s: s["ytd_gms"], reverse=True)[:20]

    # NSR movers
    nsr_active = [s for s in nsr_sellers if s["gms_cw"] != 0 or s["gms_pw"] != 0]
    nsr_gainers = sorted(nsr_active, key=lambda s: s["wow_delta"], reverse=True)
    nsr_gainers = [s for s in nsr_gainers if s["wow_delta"] > 0][:5]
    nsr_decliners = sorted(nsr_active, key=lambda s: s["wow_delta"])
    nsr_decliners = [s for s in nsr_decliners if s["wow_delta"] < 0][:5]

    # ESM movers
    esm_active = [s for s in esm_sellers if s["gms_cw"] != 0 or s["gms_pw"] != 0]
    esm_gainers = sorted(esm_active, key=lambda s: s["wow_delta"], reverse=True)
    esm_gainers = [s for s in esm_gainers if s["wow_delta"] > 0][:5]
    esm_decliners = sorted(esm_active, key=lambda s: s["wow_delta"])
    esm_decliners = [s for s in esm_decliners if s["wow_delta"] < 0][:5]

    # DSR active count for summary
    dsr_active_cw = sum(1 for s in nsr_sellers if s["active"])
    dsr_active_ly = sum(1 for s in sellers if s["channel"] in ("DSR", "SSR")
                        and s["gms_ly"] > 0)

    # Top PG
    pg_cw = d["pg_cw"]
    top_pg = max(pg_cw.items(), key=lambda x: x[1]["gms"])[0] if pg_cw else "N/A"
    top_pg_share = (pg_cw[top_pg]["gms"] / total_gms * 100) if total_gms and top_pg in pg_cw else 0

    # ── Build HTML ──
    parts = []
    parts.append(_html_head(market_code, market_label, week_num))
    parts.append(_html_kpis(t_cw, t_pw, t_ly, wow_gms, yoy_gms, wow_units, yoy_units,
                            wow_active, yoy_active, top_seller, fba_pct, ytd_yoy_gms, week_num))
    parts.append(_html_exec_summary(t_cw, t_pw, t_ly, wow_gms, yoy_gms, wow_units, yoy_units,
                                     fba_pct, ytd_yoy_gms, ytd_yoy_units, top_seller,
                                     dsr_active_cw, dsr_active_ly, top_pg, top_pg_share,
                                     gainers, decliners, market_label, week_num))
    parts.append(_html_tabs())
    parts.append(_html_tab0_summary(t_cw, t_pw, t_ly, wow_gms, yoy_gms, wow_units, yoy_units,
                                     wow_fba_gms, yoy_fba_gms, wow_fba_units, yoy_fba_units,
                                     wow_active, yoy_active, ytd_yoy_gms, ytd_yoy_units,
                                     market_label, week_num))
    parts.append(_html_tab1_category(d, total_gms, week_num))
    parts.append(_html_tab2_top_sellers(top30, total_gms, total_ytd, week_num))
    parts.append(_html_tab3_movers(gainers, decliners, total_gms, week_num))
    parts.append(_html_tab4_deep_dive(top5, total_gms, total_ytd, d, week_num))
    parts.append(_html_tab5_cohort(nsr_t, esm_t, nsr_top20, esm_top20,
                                    nsr_gainers, nsr_decliners, esm_gainers, esm_decliners,
                                    total_gms, total_ytd, week_num))
    parts.append(_html_tab6_all_sellers(nsr_all, esm_all, nsr_t, esm_t, total_gms, total_ytd, week_num))
    parts.append(_html_tab7_dsr_launches(dsr_launches, nsr_t, total_gms, week_num))
    parts.append(_html_footer())

    return "\n".join(parts)


def _launch_sort_key(s):
    """Sort key for DSR launches: by launch_date descending."""
    ld = s.get("launch_date")
    if ld is None:
        return datetime.min
    if isinstance(ld, datetime):
        return ld
    try:
        return datetime.strptime(str(ld)[:10], "%Y-%m-%d")
    except Exception:
        return datetime.min


def _fmt_launch_date(ld):
    """Format launch date for display."""
    if ld is None:
        return "N/A"
    if isinstance(ld, datetime):
        return ld.strftime("%Y-%m-%d")
    s = str(ld).strip()
    return s[:10] if len(s) >= 10 else s


# ── HTML Head / CSS ────────────────────────────────────────────────────────
def _html_head(market_code, market_label, wk):
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<script src="../auth.js"></script>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WBR - TW2{market_code} Pipeline W{wk} 2026</title>
<style>
:root{{--blue:#4472C4;--bl:#D6E4F0;--g:#006100;--gb:#C6EFCE;--r:#9C0006;--rb:#FFC7CE;--gy:#F2F2F2;--dk:#1a1a2e;--bd:#dee2e6}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:"Segoe UI",system-ui,sans-serif;background:#f0f2f5;color:#333;line-height:1.5}}
.ctn{{max-width:1500px;margin:0 auto;padding:20px}}
.hdr{{background:linear-gradient(135deg,var(--dk),var(--blue));color:#fff;padding:30px 40px;border-radius:12px;margin-bottom:24px}}
.hdr h1{{font-size:24px;font-weight:700;margin-bottom:4px}}.hdr p{{opacity:.85;font-size:14px}}
.tabs{{display:flex;gap:4px;margin-bottom:20px;flex-wrap:wrap}}
.tab{{padding:10px 20px;background:#fff;border:1px solid var(--bd);border-radius:8px 8px 0 0;cursor:pointer;font-size:13px;font-weight:600;color:#666;transition:all .2s}}
.tab:hover{{color:var(--blue)}}.tab.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
.pnl{{display:none}}.pnl.active{{display:block}}
.card{{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:24px;margin-bottom:20px}}
.card h2{{font-size:16px;color:var(--blue);margin-bottom:16px;border-bottom:2px solid var(--bl);padding-bottom:8px}}
.kg{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px;margin-bottom:24px}}
.kpi{{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);border-left:4px solid var(--blue)}}
.kpi .lb{{font-size:12px;color:#888;text-transform:uppercase;letter-spacing:.5px}}
.kpi .vl{{font-size:28px;font-weight:700;margin:4px 0}}.kpi .ch{{font-size:13px;font-weight:600}}
.pos{{color:var(--g)}}.neg{{color:var(--r)}}
.pos-bg{{background:var(--gb);color:var(--g)}}.neg-bg{{background:var(--rb);color:var(--r)}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{background:var(--blue);color:#fff;padding:10px 12px;text-align:center;font-weight:600;white-space:nowrap}}
td{{padding:8px 12px;border-bottom:1px solid #eee}}tr:nth-child(even){{background:var(--gy)}}
.tr{{text-align:right}}.tc{{text-align:center}}
.bd{{display:inline-block;padding:2px 8px;border-radius:4px;font-weight:600;font-size:12px}}
.at{{max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px}}
.mv-up{{border-left:4px solid var(--g)}}.mv-dn{{border-left:4px solid var(--r)}}
.mv-card{{background:#fff;border-radius:10px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,.08);margin-bottom:12px}}
.mv-card .nm{{font-size:15px;font-weight:700;margin-bottom:4px}}.mv-card .dt{{font-size:13px;color:#555}}
.bar-container{{display:flex;align-items:center;gap:8px;margin:4px 0}}
.bar{{height:18px;border-radius:3px;min-width:2px}}
.bar-cw{{background:var(--blue)}}.bar-pw{{background:#B4C7E7}}.bar-ly{{background:#FFC000}}
@media(max-width:768px){{.kg{{grid-template-columns:1fr}}table{{font-size:11px}}th,td{{padding:6px 8px}}}}
.cpbtn{{display:inline-block;padding:6px 14px;background:var(--blue);color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;margin-bottom:12px;transition:all .2s}}.cpbtn:hover{{opacity:.85}}
</style>
</head>
<body>
<div class="ctn">
<div class="hdr" style="position:relative"><a href="../index.html" style="position:absolute;top:20px;right:20px;padding:8px 18px;background:rgba(255,255,255,0.15);color:#fff;border-radius:6px;font-size:13px;font-weight:600;text-decoration:none;transition:all .2s;border:1px solid rgba(255,255,255,0.3)" onmouseover="this.style.background=\'rgba(255,255,255,0.3)\'" onmouseout="this.style.background=\'rgba(255,255,255,0.15)\'">&#127968; Home</a><h1>WBR &mdash; TW2{market_code} Pipeline Weekly Business Review</h1>
<p>W{wk} 2026 &nbsp;|&nbsp; {market_label} &nbsp;|&nbsp; USD</p></div>'''


# ── KPI Cards ──────────────────────────────────────────────────────────────
def _html_kpis(t_cw, t_pw, t_ly, wow_gms, yoy_gms, wow_units, yoy_units,
               wow_active, yoy_active, top_seller, fba_pct, ytd_yoy_gms, wk):
    ts_name = h(top_seller["name"]) if top_seller else "N/A"
    ts_gms = fmt_money(top_seller["gms_cw"]) if top_seller else "$0"
    return f'''<div class="kg">
<div class="kpi"><div class="lb">Total Ordered GMS (USD)</div><div class="vl">{fmt_money(t_cw["gms"])}</div><div class="ch {pos_neg_class(wow_gms)}">WoW: {arrow_text(wow_gms)}</div><div class="ch {pos_neg_class(yoy_gms)}">YoY: {arrow_text(yoy_gms)}</div></div>
<div class="kpi"><div class="lb">Total Ordered Units</div><div class="vl">{safe_int(t_cw["units"]):,}</div><div class="ch {pos_neg_class(wow_units)}">WoW: {arrow_text(wow_units)}</div><div class="ch {pos_neg_class(yoy_units)}">YoY: {arrow_text(yoy_units)}</div></div>
<div class="kpi"><div class="lb">Active Sellers</div><div class="vl">{t_cw["active"]:,}</div><div class="ch {pos_neg_class(wow_active)}">WoW: {arrow_text(wow_active)}</div><div class="ch {pos_neg_class(yoy_active)}">YoY: {arrow_text(yoy_active)}</div></div>
<div class="kpi"><div class="lb">Top Seller</div><div class="vl" style="font-size:16px">{ts_name}</div><div class="ch">{ts_gms}</div></div>
<div class="kpi"><div class="lb">YTD GMS (USD)</div><div class="vl">{fmt_money(t_cw["ytd_gms"])}</div><div class="ch {pos_neg_class(ytd_yoy_gms)}">YoY: {arrow_text(ytd_yoy_gms)}</div></div>
</div>'''


# ── Executive Summary ──────────────────────────────────────────────────────
def _html_exec_summary(t_cw, t_pw, t_ly, wow_gms, yoy_gms, wow_units, yoy_units,
                        fba_pct, ytd_yoy_gms, ytd_yoy_units, top_seller,
                        dsr_active_cw, dsr_active_ly, top_pg, top_pg_share,
                        gainers, decliners, market_label, wk):
    pw = wk - 1
    ts_name = h(top_seller["name"]) if top_seller else "N/A"
    ts_gms = fmt_money(top_seller["gms_cw"]) if top_seller else "$0"
    ts_share = (top_seller["gms_cw"] / t_cw["gms"] * 100) if top_seller and t_cw["gms"] else 0

    nsr_active = sum(1 for s in [] if True)  # placeholder, computed in caller
    # Count NSR vs ESM active
    # We'll use the data from t_cw
    esm_active = t_cw["active"] - dsr_active_cw

    dsr_yoy_pct = pct_change(dsr_active_cw, dsr_active_ly)
    dsr_yoy_cls = pos_neg_class(dsr_yoy_pct)

    # Gainers/decliners text
    gainer_items = []
    for g in gainers[:3]:
        gainer_items.append(f'<li><strong>{h(g["name"])}</strong> surged {fmt_pct(g["wow_pct"])} WoW ({fmt_money_delta(g["wow_delta"])}).</li>')
    decliner_items = []
    for d in decliners[:3]:
        decliner_items.append(f'<li><strong>{h(d["name"])}</strong> dropped {fmt_pct(d["wow_pct"])} WoW ({fmt_money_delta(d["wow_delta"])}).</li>')

    callouts = "\n".join(gainer_items + decliner_items)
    yoy_direction = "Positive" if yoy_gms >= 0 else "Negative"
    yoy_comment = "healthy pipeline growth vs last year." if yoy_gms >= 0 else "pipeline contraction vs last year."

    gainer_summary = ", ".join([f'{h(g["name"])} ({fmt_pct(g["wow_pct"])})' for g in gainers[:3]]) if gainers else "None"
    decliner_summary = ", ".join([f'{h(d["name"])} ({fmt_pct(d["wow_pct"])})' for d in decliners[:3]]) if decliners else "None"

    return f'''<div class="card" style="border-left:4px solid var(--blue)">
<h2>&#128221; Executive Summary</h2>
<div style="font-size:14px;line-height:1.8">
<p><strong>Overall Performance:</strong> W{wk} 2026 total ordered GMS came in at <strong>{fmt_money(t_cw["gms"])}</strong>, {"up" if wow_gms >= 0 else "down"} <strong>{fmt_pct(wow_gms)} WoW</strong> from {fmt_money(t_pw["gms"])} (W{pw} 2026) and {"up" if yoy_gms >= 0 else "down"} <strong>{fmt_pct(yoy_gms)} YoY</strong> vs {fmt_money(t_ly["gms"])} (W{wk} 2025). Units: <strong>{safe_int(t_cw["units"]):,}</strong> (WoW {fmt_pct(wow_units)}, YoY {fmt_pct(yoy_units)}). FBA: <strong>{fba_pct:.1f}%</strong> of GMS. YTD GMS: <strong>{fmt_money(t_cw["ytd_gms"])}</strong> (YoY {fmt_pct(ytd_yoy_gms)}), YTD Units: <strong>{safe_int(t_cw["ytd_units"]):,}</strong> (YoY {fmt_pct(ytd_yoy_units)}).</p>
<p><strong>Seller Landscape:</strong> <strong>{t_cw["active"]}</strong> active sellers (WoW {fmt_pct(pct_change(t_cw["active"], t_pw["active"]))}, YoY {fmt_pct(pct_change(t_cw["active"], t_ly["active"]))}). NSR: <strong>{dsr_active_cw}</strong>, ESM: <strong>{esm_active}</strong>. Top seller: <strong>{ts_name}</strong> &mdash; {ts_gms} ({ts_share:.1f}% share).</p>
<p><strong>DSR Pipeline:</strong> <strong>{dsr_active_cw}</strong> active DSR sellers in W{wk} 2026 vs <strong>{dsr_active_ly}</strong> in W{wk} 2025 (<span class="{dsr_yoy_cls}">{arrow_text(dsr_yoy_pct)}</span>).</p>
<p><strong>Top Product Group:</strong> <strong>{h(top_pg)}</strong> &mdash; {top_pg_share:.1f}% of GMS.</p>
<p><strong>Movers &amp; Shakers:</strong> Gainers: {gainer_summary}. Decliners: {decliner_summary}.</p>
<p><strong>Key Callouts:</strong></p><ul style="margin:4px 0 0 20px;font-size:13px">
<li>{yoy_direction} YoY ({fmt_pct(yoy_gms)}) &mdash; {yoy_comment}</li>
{callouts}
</ul></div></div>'''


# ── Tabs ───────────────────────────────────────────────────────────────────
def _html_tabs():
    return '''<div class="tabs">
<div class="tab active" onclick="showTab(0)">Summary</div>
<div class="tab" onclick="showTab(1)">Category / PG</div>
<div class="tab" onclick="showTab(2)">Top Sellers</div>
<div class="tab" onclick="showTab(3)">Movers &amp; Shakers</div>
<div class="tab" onclick="showTab(4)">Seller Deep Dive</div>
<div class="tab" onclick="showTab(5)">Cohort (NSR vs ESM)</div>
<div class="tab" onclick="showTab(6)">All Sellers (DSR/ESM)</div>
<div class="tab" onclick="showTab(7)">DSR Launches</div>
</div>'''


# ── Tab 0: Summary ────────────────────────────────────────────────────────
def _html_tab0_summary(t_cw, t_pw, t_ly, wow_gms, yoy_gms, wow_units, yoy_units,
                        wow_fba_gms, yoy_fba_gms, wow_fba_units, yoy_fba_units,
                        wow_active, yoy_active, ytd_yoy_gms, ytd_yoy_units,
                        market_label, wk):
    pw = wk - 1

    def row(label, cw_val, pw_val, ly_val, wow, yoy, is_money=True, is_ytd=False):
        fmt = fmt_money if is_money else lambda v: f"{safe_int(v):,}"
        if is_ytd:
            return f'''<tr style="background:#E8F0FE"><td><strong>{label}</strong></td><td class="tr">{fmt(cw_val)}</td><td class="tr">&mdash;</td><td class="tr">&mdash;</td><td class="tc">&mdash;</td><td class="tr">{fmt(ly_val)}</td><td class="tr">{fmt(cw_val - ly_val) if is_money else f"{safe_int(cw_val - ly_val):,}"}</td><td class="tc">{badge(yoy)}</td></tr>'''
        delta = cw_val - pw_val
        yoy_delta = cw_val - ly_val
        return f'''<tr><td><strong>{label}</strong></td><td class="tr">{fmt(cw_val)}</td><td class="tr">{fmt(pw_val)}</td><td class="tr">{fmt(delta) if is_money else f"{safe_int(delta):,}"}</td><td class="tc">{badge(wow)}</td><td class="tr">{fmt(ly_val)}</td><td class="tr">{fmt(yoy_delta) if is_money else f"{safe_int(yoy_delta):,}"}</td><td class="tc">{badge(yoy)}</td></tr>'''

    market_name = market_label.split(" (")[0] if " (" in market_label else market_label

    return f'''<div class="pnl active" id="p0"><div class="card">
<h2>Weekly Summary ({market_name})</h2>
<table><thead><tr><th>Metric</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW Delta</th><th>WoW %</th><th>W{wk} 2025</th><th>YoY Delta</th><th>YoY %</th></tr></thead><tbody>
{row("Ordered GMS (USD)", t_cw["gms"], t_pw["gms"], t_ly["gms"], wow_gms, yoy_gms)}
{row("Ordered Units", t_cw["units"], t_pw["units"], t_ly["units"], wow_units, yoy_units, is_money=False)}
{row("FBA GMS (USD)", t_cw["fba_gms"], t_pw["fba_gms"], t_ly["fba_gms"], wow_fba_gms, yoy_fba_gms)}
{row("FBA Units", t_cw["fba_units"], t_pw["fba_units"], t_ly["fba_units"], wow_fba_units, yoy_fba_units, is_money=False)}
{row("Active Sellers", t_cw["active"], t_pw["active"], t_ly["active"], wow_active, yoy_active, is_money=False)}
{row("YTD GMS (USD)", t_cw["ytd_gms"], 0, t_ly["ytd_gms"], 0, ytd_yoy_gms, is_ytd=True)}
{row("YTD Units", t_cw["ytd_units"], 0, t_ly["ytd_units"], 0, ytd_yoy_units, is_money=False, is_ytd=True)}
</tbody></table></div></div>'''


# ── Tab 1: Category / PG ──────────────────────────────────────────────────
def _html_tab1_category(data, total_gms, wk):
    pw = wk - 1

    def build_table(title, cw_dict, pw_dict, ly_dict, total):
        rows_data = []
        for key in cw_dict:
            gms_cw = cw_dict[key]["gms"]
            gms_pw = pw_dict.get(key, {}).get("gms", 0)
            gms_ly = ly_dict.get(key, {}).get("gms", 0)
            units = cw_dict[key]["units"]
            ytd_cw = cw_dict[key]["ytd_gms"]
            ytd_ly = ly_dict.get(key, {}).get("ytd_gms", 0)
            rows_data.append((key, gms_cw, gms_pw, gms_ly, units, ytd_cw, ytd_ly))
        rows_data.sort(key=lambda x: x[1], reverse=True)

        rows_html = []
        for key, gms_cw, gms_pw, gms_ly, units, ytd_cw, ytd_ly in rows_data:
            wow = pct_change(gms_cw, gms_pw)
            yoy = pct_change(gms_cw, gms_ly)
            share = (gms_cw / total * 100) if total else 0
            ytd_yoy = pct_change(ytd_cw, ytd_ly)
            rows_html.append(
                f'<tr><td><strong>{h(key)}</strong></td>'
                f'<td class="tr">{fmt_money(gms_cw)}</td>'
                f'<td class="tr">{fmt_money(gms_pw)}</td>'
                f'<td class="tc">{badge(wow)}</td>'
                f'<td class="tr">{fmt_money(gms_ly)}</td>'
                f'<td class="tc">{badge(yoy)}</td>'
                f'<td class="tr">{safe_int(units):,}</td>'
                f'<td class="tc">{share:.1f}%</td>'
                f'<td class="tr">{fmt_money(ytd_cw)}</td>'
                f'<td class="tc">{badge(ytd_yoy)}</td></tr>'
            )

        label_col = "Category" if "Category" in title else "PG"
        return f'''<div class="card"><h2>{title}</h2>
<table><thead><tr><th>{label_col}</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW %</th><th>W{wk} 2025</th><th>YoY %</th><th>Units</th><th>Share</th><th>YTD GMS</th><th>YTD YoY</th></tr></thead><tbody>
{"".join(rows_html)}
</tbody></table></div>'''

    cat_html = build_table("GMS by Account Primary Category",
                           data["cat_cw"], data["cat_pw"], data["cat_ly"], total_gms)
    pg_html = build_table("GMS by SP Primary Product Group",
                          data["pg_cw"], data["pg_pw"], data["pg_ly"], total_gms)

    return f'<div class="pnl" id="p1">\n{cat_html}\n{pg_html}\n</div>'


# ── Tab 2: Top 30 Sellers ─────────────────────────────────────────────────
def _html_tab2_top_sellers(top30, total_gms, total_ytd, wk):
    pw = wk - 1
    mcids = ",".join(s["mcid"] for s in top30)
    rows = []
    for i, s in enumerate(top30, 1):
        wow = s["wow_pct"]
        yoy = s["yoy_pct"]
        ytd_yoy = s["ytd_yoy"]
        wk_share = (s["gms_cw"] / total_gms * 100) if total_gms else 0
        ytd_share = (s["ytd_gms"] / total_ytd * 100) if total_ytd else 0
        ytd_arrow = f'&#9650; {fmt_pct(ytd_yoy)}' if ytd_yoy >= 0 else f'&#9660; {fmt_pct(ytd_yoy)}'
        rows.append(
            f'<tr><td class="tc">{i}</td>'
            f'<td><strong>{h(s["name"])}</strong></td>'
            f'<td class="tc" style="font-size:11px">{h(s["mcid"])}</td>'
            f'<td class="tr">{fmt_money(s["gms_cw"])}</td>'
            f'<td class="tr">{fmt_money(s["gms_pw"])}</td>'
            f'<td class="tc">{badge(wow)}</td>'
            f'<td class="tr">{fmt_money(s["gms_ly"])}</td>'
            f'<td class="tc">{badge(yoy)}</td>'
            f'<td class="tr">{safe_int(s["units_cw"]):,}</td>'
            f'<td class="tr">{fmt_money(s["ytd_gms"])}</td>'
            f'<td class="tc">{ytd_arrow}</td>'
            f'<td class="tc">{ytd_share:.1f}%</td>'
            f'<td class="at">{h(s["category"])}</td>'
            f'<td class="tc">{wk_share:.1f}%</td></tr>'
        )

    return f'''<div class="pnl" id="p2"><div class="card">
<h2>Top 30 Sellers by YTD GMS</h2>
<button class="cpbtn" data-mcids="{mcids}" onclick="copyMcids(this)">&#128203; Copy MCIDs ({len(top30)})</button>
<table><thead><tr><th>#</th><th>Seller</th><th>MCID</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW %</th><th>W{wk} 2025</th><th>YoY %</th><th>Units</th><th>YTD GMS</th><th>YTD YoY</th><th>YTD Share</th><th>Category</th><th>Wk Share</th></tr></thead><tbody>
{"".join(rows)}
</tbody></table></div></div>'''


# ── Tab 3: Movers & Shakers ───────────────────────────────────────────────
def _html_tab3_movers(gainers, decliners, total_gms, wk):
    pw = wk - 1

    def mover_table(title, movers, is_gainer=True):
        mcids = ",".join(s["mcid"] for s in movers)
        color_var = "var(--g)" if is_gainer else "var(--r)"
        rows = []
        for i, s in enumerate(movers, 1):
            rows.append(
                f'<tr><td class="tc">{i}</td>'
                f'<td><strong>{h(s["name"])}</strong></td>'
                f'<td class="tc" style="font-size:11px">{h(s["mcid"])}</td>'
                f'<td class="tr">{fmt_money(s["gms_cw"])}</td>'
                f'<td class="tr">{fmt_money(s["gms_pw"])}</td>'
                f'<td class="tr" style="color:{color_var};font-weight:700">{fmt_money_delta(s["wow_delta"])}</td>'
                f'<td class="tc">{badge(s["wow_pct"])}</td>'
                f'<td class="tr">{fmt_money(s["gms_ly"])}</td>'
                f'<td class="tc">{badge(s["yoy_pct"])}</td>'
                f'<td class="tr">{fmt_money(s["ytd_gms"])}</td>'
                f'<td class="tc">{badge(s["ytd_yoy"])}</td>'
                f'<td class="at">{h(s["category"])}</td></tr>'
            )
        icon = "&#128293;" if is_gainer else "&#128308;"
        return f'''<div class="card"><h2>{icon} {title}</h2>
<button class="cpbtn" data-mcids="{mcids}" onclick="copyMcids(this)">&#128203; Copy MCIDs ({len(movers)})</button>
<table><thead><tr><th>#</th><th>Seller</th><th>MCID</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW Delta</th><th>WoW %</th><th>W{wk} 2025</th><th>YoY %</th><th>YTD GMS</th><th>YTD YoY</th><th>Category</th></tr></thead><tbody>
{"".join(rows)}
</tbody></table></div>'''

    # Visual cards
    max_gms = max([s["gms_cw"] for s in gainers] + [s["gms_pw"] for s in gainers] +
                  [s["gms_ly"] for s in gainers] +
                  [s["gms_cw"] for s in decliners] + [s["gms_pw"] for s in decliners] +
                  [s["gms_ly"] for s in decliners] + [1]) if (gainers or decliners) else 1

    def visual_card(s, is_up=True):
        cls = "mv-up" if is_up else "mv-dn"
        bar_max = max(abs(s["gms_cw"]), abs(s["gms_pw"]), abs(s["gms_ly"]), 1)
        # Use max_gms for consistent scaling
        scale = max_gms if max_gms > 0 else 1
        w_cw = abs(s["gms_cw"]) / scale * 100
        w_pw = abs(s["gms_pw"]) / scale * 100
        w_ly = abs(s["gms_ly"]) / scale * 100
        return f'''<div class="mv-card {cls}"><div class="nm">{h(s["name"])}</div>
<div class="bar-container"><span style="width:55px;font-size:10px">W{wk} 2026</span><div class="bar bar-cw" style="width:{w_cw:.1f}%"></div><span style="font-size:11px">{fmt_money(s["gms_cw"])}</span></div>
<div class="bar-container"><span style="width:55px;font-size:10px">W{pw} 2026</span><div class="bar bar-pw" style="width:{w_pw:.1f}%"></div><span style="font-size:11px">{fmt_money(s["gms_pw"])}</span></div>
<div class="bar-container"><span style="width:55px;font-size:10px">W{wk} 2025</span><div class="bar bar-ly" style="width:{w_ly:.1f}%"></div><span style="font-size:11px">{fmt_money(s["gms_ly"])}</span></div>
<div class="dt">WoW: {fmt_pct(s["wow_pct"])} | YoY: {fmt_pct(s["yoy_pct"])}</div></div>'''

    gainer_cards = "\n".join(visual_card(s, True) for s in gainers[:7])
    decliner_cards = "\n".join(visual_card(s, False) for s in decliners[:7])

    visual = f'''<div class="card"><h2>Movers &amp; Shakers Visual</h2>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
<div><h3 style="color:var(--g);margin-bottom:12px;font-size:14px">&#9650; Top Gainers</h3>
{gainer_cards}
</div>
<div><h3 style="color:var(--r);margin-bottom:12px;font-size:14px">&#9660; Top Decliners</h3>
{decliner_cards}
</div>
</div></div>'''

    return f'''<div class="pnl" id="p3">
{mover_table("Top Gainers (Biggest WoW GMS Increase)", gainers, True)}
{mover_table("Top Decliners (Biggest WoW GMS Decrease)", decliners, False)}
{visual}
</div>'''


# ── Tab 4: Seller Deep Dive ───────────────────────────────────────────────
def _html_tab4_deep_dive(top5, total_gms, total_ytd, data, wk):
    pw = wk - 1
    pw2 = wk - 2  # W-2 for the 3-week bar chart

    cards = []
    for s in top5:
        wow_gms = s["wow_pct"]
        yoy_gms = s["yoy_pct"]
        wow_units = pct_change(s["units_cw"], s["units_pw"])
        yoy_units = pct_change(s["units_cw"], s["units_ly"])
        wk_share = (s["gms_cw"] / total_gms * 100) if total_gms else 0
        ytd_share = (s["ytd_gms"] / total_ytd * 100) if total_ytd else 0
        ytd_yoy = s["ytd_yoy"]
        ytd_cls = pos_neg_class(ytd_yoy)

        # Get W-2 data if available
        gms_w2 = 0
        pw2_sellers = data.get("pw2_sellers", {})
        if s["mcid"] in pw2_sellers:
            gms_w2 = pw2_sellers[s["mcid"]].get("wtd_gms", 0)

        # Bar chart: W-2, W-1, W current, LY
        bar_vals = [gms_w2, s["gms_pw"], s["gms_cw"], s["gms_ly"]]
        bar_max = max(abs(v) for v in bar_vals) if any(bar_vals) else 1
        if bar_max == 0:
            bar_max = 1

        bars = f'''<div class="bar-container"><span style="width:55px;font-size:11px;font-weight:600">W{pw2} 2026</span><div class="bar" style="width:{abs(gms_w2)/bar_max*100:.1f}%;background:#B4C7E7"></div><span style="font-size:11px">{fmt_money(gms_w2)}</span></div>
<div class="bar-container"><span style="width:55px;font-size:11px;font-weight:600">W{pw} 2026</span><div class="bar" style="width:{abs(s["gms_pw"])/bar_max*100:.1f}%;background:#7BA0D4"></div><span style="font-size:11px">{fmt_money(s["gms_pw"])}</span></div>
<div class="bar-container"><span style="width:55px;font-size:11px;font-weight:600">W{wk} 2026</span><div class="bar" style="width:{abs(s["gms_cw"])/bar_max*100:.1f}%;background:var(--blue)"></div><span style="font-size:11px">{fmt_money(s["gms_cw"])}</span></div>
<div class="bar-container"><span style="width:55px;font-size:11px;font-weight:600">W{wk} 2025</span><div class="bar" style="width:{abs(s["gms_ly"])/bar_max*100:.1f}%;background:#FFC000"></div><span style="font-size:11px">{fmt_money(s["gms_ly"])}</span></div>'''

        cards.append(f'''<div class="card"><h2>{h(s["name"])} <span style="font-size:12px;color:#888;font-weight:400">MCID: {h(s["mcid"])}</span></h2>
<div class="kg" style="margin-bottom:16px">
<div class="kpi"><div class="lb">W{wk} 2026 GMS</div><div class="vl" style="font-size:22px">{fmt_money(s["gms_cw"])}</div><div class="ch {pos_neg_class(wow_gms)}">WoW: {arrow_text(wow_gms)}</div><div class="ch {pos_neg_class(yoy_gms)}">YoY: {arrow_text(yoy_gms)}</div></div>
<div class="kpi"><div class="lb">W{pw} 2026 GMS</div><div class="vl" style="font-size:22px">{fmt_money(s["gms_pw"])}</div></div>
<div class="kpi"><div class="lb">W{wk} 2025 GMS</div><div class="vl" style="font-size:22px">{fmt_money(s["gms_ly"])}</div></div>
<div class="kpi"><div class="lb">W{wk} 2026 Units</div><div class="vl" style="font-size:22px">{safe_int(s["units_cw"]):,}</div><div class="ch {pos_neg_class(wow_units)}">WoW: {arrow_text(wow_units)}</div><div class="ch {pos_neg_class(yoy_units)}">YoY: {arrow_text(yoy_units)}</div></div>
</div>
<div style="display:flex;gap:20px;font-size:13px;color:#555"><div><strong>Category:</strong> {h(s["category"])}</div><div><strong>PG:</strong> {h(s["pg"])}</div><div><strong>Wk Share:</strong> {wk_share:.1f}%</div><div><strong>YTD GMS:</strong> {fmt_money(s["ytd_gms"])}</div><div><strong>YTD YoY:</strong> <span class="{ytd_cls}">{arrow_text(ytd_yoy)}</span></div><div><strong>YTD Share:</strong> {ytd_share:.1f}%</div></div>
<div style="margin-top:12px"><h3 style="font-size:13px;color:#888;margin-bottom:6px">GMS Comparison</h3>{bars}
</div></div>''')

    return f'<div class="pnl" id="p4">\n' + "\n".join(cards) + '\n</div>'


# ── Tab 5: Cohort (NSR vs ESM) ────────────────────────────────────────────
def _html_tab5_cohort(nsr_t, esm_t, nsr_top20, esm_top20,
                       nsr_gainers, nsr_decliners, esm_gainers, esm_decliners,
                       total_gms, total_ytd, wk):
    pw = wk - 1

    def cohort_row(label, t, total_gms_val, total_ytd_val):
        wow = pct_change(t["gms"], t["gms_pw"])
        yoy = pct_change(t["gms"], t["gms_ly"])
        fba_pct = (t["fba_gms"] / t["gms"] * 100) if t["gms"] else 0
        ytd_yoy = pct_change(t["ytd_gms"], t["ytd_gms_ly"])
        ytd_share = (t["ytd_gms"] / total_ytd_val * 100) if total_ytd_val else 0
        wk_share = (t["gms"] / total_gms_val * 100) if total_gms_val else 0
        return (f'<tr><td><strong>{label}</strong></td>'
                f'<td class="tr">{fmt_money(t["gms"])}</td>'
                f'<td class="tr">{fmt_money(t["gms_pw"])}</td>'
                f'<td class="tc">{badge(wow)}</td>'
                f'<td class="tr">{fmt_money(t["gms_ly"])}</td>'
                f'<td class="tc">{badge(yoy)}</td>'
                f'<td class="tr">{safe_int(t["units"]):,}</td>'
                f'<td class="tc">{t["active"]}</td>'
                f'<td class="tc">{fba_pct:.1f}%</td>'
                f'<td class="tr">{fmt_money(t["ytd_gms"])}</td>'
                f'<td class="tc">{badge(ytd_yoy)}</td>'
                f'<td class="tc">{ytd_share:.1f}%</td>'
                f'<td class="tc">{wk_share:.1f}%</td></tr>')

    overview = f'''<div class="card"><h2>Cohort Overview: NSR (DSR+SSR) vs ESM</h2>
<table><thead><tr><th>Cohort</th><th>W{wk} 2026 GMS</th><th>W{pw} 2026 GMS</th><th>WoW %</th><th>W{wk} 2025 GMS</th><th>YoY %</th><th>W{wk} 2026 Units</th><th>Sellers</th><th>FBA %</th><th>YTD GMS</th><th>YTD YoY</th><th>YTD Share</th><th>Wk Share</th></tr></thead><tbody>
{cohort_row("NSR", nsr_t, total_gms, total_ytd)}
{cohort_row("ESM", esm_t, total_gms, total_ytd)}
</tbody></table></div>'''

    def seller_table(title, sellers, cohort_gms, cohort_ytd):
        mcids = ",".join(s["mcid"] for s in sellers)
        rows = []
        for i, s in enumerate(sellers, 1):
            wow = s["wow_pct"]
            yoy = s["yoy_pct"]
            ytd_yoy = s["ytd_yoy"]
            ytd_share = (s["ytd_gms"] / cohort_ytd * 100) if cohort_ytd else 0
            wk_share = (s["gms_cw"] / cohort_gms * 100) if cohort_gms else 0
            rows.append(
                f'<tr><td class="tc">{i}</td>'
                f'<td><strong>{h(s["name"])}</strong></td>'
                f'<td class="tc" style="font-size:11px">{h(s["mcid"])}</td>'
                f'<td class="tr">{fmt_money(s["gms_cw"])}</td>'
                f'<td class="tr">{fmt_money(s["gms_pw"])}</td>'
                f'<td class="tc">{badge(wow)}</td>'
                f'<td class="tr">{fmt_money(s["gms_ly"])}</td>'
                f'<td class="tc">{badge(yoy)}</td>'
                f'<td class="tr">{safe_int(s["units_cw"]):,}</td>'
                f'<td class="tr">{fmt_money(s["ytd_gms"])}</td>'
                f'<td class="tc">{badge(ytd_yoy)}</td>'
                f'<td class="tc">{ytd_share:.1f}%</td>'
                f'<td class="at">{h(s["category"])}</td>'
                f'<td class="tc">{wk_share:.1f}%</td></tr>'
            )
        return f'''<div class="card"><h2>{title}</h2>
<button class="cpbtn" data-mcids="{mcids}" onclick="copyMcids(this)">&#128203; Copy MCIDs ({len(sellers)})</button>
<table><thead><tr><th>#</th><th>Seller</th><th>MCID</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW %</th><th>W{wk} 2025</th><th>YoY %</th><th>Units</th><th>YTD GMS</th><th>YTD YoY</th><th>YTD Share</th><th>Category</th><th>Wk Share</th></tr></thead><tbody>
{"".join(rows)}
</tbody></table></div>'''

    def movers_section(title, g_list, d_list):
        def mover_rows(movers, is_gain):
            color = "var(--g)" if is_gain else "var(--r)"
            rows = []
            for s in movers:
                rows.append(
                    f'<tr><td><strong>{h(s["name"])}</strong></td>'
                    f'<td class="tr">{fmt_money(s["gms_cw"])}</td>'
                    f'<td class="tr">{fmt_money(s["gms_pw"])}</td>'
                    f'<td class="tr" style="color:{color};font-weight:700">{fmt_money_delta(s["wow_delta"])}</td>'
                    f'<td class="tc">{badge(s["wow_pct"])}</td></tr>'
                )
            return "".join(rows)

        return f'''<div class="card"><h2>{title}</h2>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
<div><h3 style="color:var(--g);font-size:14px;margin-bottom:8px">&#9650; Gainers</h3>
<table style="font-size:12px"><thead><tr><th>Seller</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>Delta</th><th>WoW</th></tr></thead><tbody>
{mover_rows(g_list, True)}
</tbody></table></div>
<div><h3 style="color:var(--r);font-size:14px;margin-bottom:8px">&#9660; Decliners</h3>
<table style="font-size:12px"><thead><tr><th>Seller</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>Delta</th><th>WoW</th></tr></thead><tbody>
{mover_rows(d_list, False)}
</tbody></table></div>
</div></div>'''

    nsr_table = seller_table("NSR (DSR + SSR) &mdash; Top 20 Sellers", nsr_top20,
                              nsr_t["gms"], nsr_t["ytd_gms"])
    nsr_movers = movers_section("NSR (DSR + SSR) &mdash; Movers &amp; Shakers",
                                 nsr_gainers, nsr_decliners)
    esm_table = seller_table("ESM &mdash; Top 20 Sellers", esm_top20,
                              esm_t["gms"], esm_t["ytd_gms"])
    esm_movers = movers_section("ESM &mdash; Movers &amp; Shakers",
                                 esm_gainers, esm_decliners)

    return f'''<div class="pnl" id="p5">
{overview}
{nsr_table}
{nsr_movers}
{esm_table}
{esm_movers}
</div>'''


# ── Tab 6: All Sellers (DSR/ESM) ──────────────────────────────────────────
def _html_tab6_all_sellers(nsr_all, esm_all, nsr_t, esm_t, total_gms, total_ytd, wk):
    pw = wk - 1

    def all_sellers_section(title, sellers, cohort_t, channel_label):
        count = len(sellers)
        wow = pct_change(cohort_t["gms"], cohort_t["gms_pw"])
        yoy = pct_change(cohort_t["gms"], cohort_t["gms_ly"])
        ytd_yoy = pct_change(cohort_t["ytd_gms"], cohort_t["ytd_gms_ly"])
        mkt_share = (cohort_t["gms"] / total_gms * 100) if total_gms else 0

        mcids = ",".join(s["mcid"] for s in sellers)
        rows = []
        for i, s in enumerate(sellers, 1):
            wow_s = s["wow_pct"]
            yoy_s = s["yoy_pct"]
            ytd_yoy_s = s["ytd_yoy"]
            delta = s["wow_delta"]
            delta_color = "var(--g)" if delta >= 0 else "var(--r)"
            ch = s["channel"] if s["channel"] else channel_label
            rows.append(
                f'<tr><td class="tc">{i}</td>'
                f'<td><strong>{h(s["name"])}</strong></td>'
                f'<td class="tc" style="font-size:11px">{h(s["mcid"])}</td>'
                f'<td class="tc">{h(ch)}</td>'
                f'<td class="tr">{fmt_money(s["gms_cw"])}</td>'
                f'<td class="tr">{fmt_money(s["gms_pw"])}</td>'
                f'<td class="tr" style="color:{delta_color};font-weight:700">{fmt_money_delta(delta)}</td>'
                f'<td class="tc">{badge(wow_s)}</td>'
                f'<td class="tr">{fmt_money(s["gms_ly"])}</td>'
                f'<td class="tc">{badge(yoy_s)}</td>'
                f'<td class="tr">{safe_int(s["units_cw"]):,}</td>'
                f'<td class="tr">{fmt_money(s["fba_cw"])}</td>'
                f'<td class="tr">{fmt_money(s["ytd_gms"])}</td>'
                f'<td class="tc">{badge(ytd_yoy_s)}</td>'
                f'<td class="at">{h(s["category"])}</td>'
                f'<td class="at">{h(s["pg"])}</td></tr>'
            )

        return f'''<div class="card">
<h2>{title} ({count} sellers)</h2>
<div class="kg" style="margin-bottom:16px">
<div class="kpi"><div class="lb">W{wk} 2026 GMS</div><div class="vl">{fmt_money(cohort_t["gms"])}</div><div class="ch {pos_neg_class(wow)}">WoW: {arrow_text(wow)}</div><div class="ch {pos_neg_class(yoy)}">YoY: {arrow_text(yoy)}</div></div>
<div class="kpi"><div class="lb">YTD GMS</div><div class="vl">{fmt_money(cohort_t["ytd_gms"])}</div><div class="ch {pos_neg_class(ytd_yoy)}">YoY: {arrow_text(ytd_yoy)}</div></div>
<div class="kpi"><div class="lb">Mkt Share</div><div class="vl">{mkt_share:.1f}%</div></div>
</div>
<button class="cpbtn" data-mcids="{mcids}" onclick="copyMcids(this)">&#128203; Copy MCIDs ({count})</button>
<table><thead><tr><th>#</th><th>Seller</th><th>MCID</th><th>Channel</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW Delta</th><th>WoW %</th><th>W{wk} 2025</th><th>YoY %</th><th>Units</th><th>FBA GMS</th><th>YTD GMS</th><th>YTD YoY</th><th>Category</th><th>PG</th></tr></thead><tbody>
{"".join(rows)}
</tbody></table>
</div>'''

    dsr_section = all_sellers_section("DSR / SSR (NSR) &mdash; All Sellers", nsr_all, nsr_t, "DSR")
    esm_section = all_sellers_section("ESM &mdash; All Sellers", esm_all, esm_t, "ESM")

    return f'''<div class="pnl" id="p6">
{dsr_section}
{esm_section}
</div>'''


# ── Tab 7: DSR Launches ───────────────────────────────────────────────────
def _html_tab7_dsr_launches(dsr_launches, nsr_t, total_gms, wk):
    pw = wk - 1
    count = len(dsr_launches)
    nsr_gms = nsr_t["gms"]
    share = (nsr_gms / total_gms * 100) if total_gms else 0

    mcids = ",".join(s["mcid"] for s in dsr_launches)
    rows = []
    for i, s in enumerate(dsr_launches, 1):
        wow = s["wow_pct"]
        delta = s["wow_delta"]
        delta_color = "var(--g)" if delta >= 0 else "var(--r)"
        ytd_yoy = s["ytd_yoy"]
        ld = _fmt_launch_date(s["launch_date"])
        rows.append(
            f'<tr><td class="tc">{i}</td>'
            f'<td><strong>{h(s["name"])}</strong></td>'
            f'<td class="tc" style="font-size:11px">{h(s["mcid"])}</td>'
            f'<td class="tc">{h(ld)}</td>'
            f'<td class="tr">{fmt_money(s["gms_cw"])}</td>'
            f'<td class="tr">{fmt_money(s["gms_pw"])}</td>'
            f'<td class="tr" style="color:{delta_color};font-weight:700">{fmt_money_delta(delta)}</td>'
            f'<td class="tc">{badge(wow)}</td>'
            f'<td class="tr">{safe_int(s["units_cw"]):,}</td>'
            f'<td class="tr">{fmt_money(s["fba_cw"])}</td>'
            f'<td class="tr">{fmt_money(s["ytd_gms"])}</td>'
            f'<td class="tc">{badge(ytd_yoy)}</td>'
            f'<td class="at">{h(s["category"])}</td>'
            f'<td class="at">{h(s["pg"])}</td></tr>'
        )

    return f'''<div class="pnl" id="p7">
<div class="card">
<h2>DSR Launches &mdash; YTD ({count} sellers)</h2>
<div class="kg" style="margin-bottom:16px">
<div class="kpi"><div class="lb">W{wk} 2026 GMS</div><div class="vl">{fmt_money(nsr_gms)}</div><div class="ch">{share:.1f}% of total GMS</div></div>
<div class="kpi"><div class="lb">YTD GMS</div><div class="vl">{fmt_money(nsr_t["ytd_gms"])}</div></div>
</div>
<button class="cpbtn" data-mcids="{mcids}" onclick="copyMcids(this)">&#128203; Copy MCIDs ({count})</button>
<table><thead><tr><th>#</th><th>Seller</th><th>MCID</th><th>Launch Date</th><th>W{wk} 2026</th><th>W{pw} 2026</th><th>WoW Delta</th><th>WoW %</th><th>Units</th><th>FBA GMS</th><th>YTD GMS</th><th>YTD YoY</th><th>Category</th><th>PG</th></tr></thead><tbody>
{"".join(rows)}
</tbody></table>
</div>
</div>'''


# ── Footer / JS ────────────────────────────────────────────────────────────
def _html_footer():
    return '''<script>function showTab(i){document.querySelectorAll(".tab").forEach((t,j)=>t.classList.toggle("active",j===i));document.querySelectorAll(".pnl").forEach((p,j)=>p.classList.toggle("active",j===i));}
function copyMcids(btn){var t=btn.getAttribute("data-mcids").split(",").join("\\n");navigator.clipboard.writeText(t).then(function(){var o=btn.textContent;btn.textContent="Copied!";btn.style.background="var(--gb)";btn.style.color="var(--g)";setTimeout(function(){btn.textContent=o;btn.style.background="";btn.style.color="";},1500);});}
</script>
</div>
</body>
</html>'''


# ── Index Update ───────────────────────────────────────────────────────────
def update_index(week_num):
    """Update wbr/index.html to include the new week if not already present."""
    index_path = Path("wbr/index.html")
    if not index_path.exists():
        print("  Warning: wbr/index.html not found, skipping index update")
        return

    content = index_path.read_text(encoding="utf-8")
    week_tag = f'"W{week_num}"'

    # Check if already present
    if week_tag in content:
        print(f"  W{week_num} already in index.html")
        return

    # Find the weeks array and add the new week
    # Pattern: const weeks = [ ... ];
    pattern = r'(const weeks = \[)'
    new_entry = f'  {{ week: "W{week_num}", year: 2026, markets: ["AE", "AU", "SA"] }},'
    # Insert after the opening bracket
    replacement = f'\\1\n  {new_entry}'

    # More robust: find the last entry and add after it
    # Look for the pattern: { week: "W##", ... },
    # and add our new entry after the last one
    lines = content.split("\n")
    new_lines = []
    inserted = False
    in_weeks = False
    last_entry_idx = -1

    for i, line in enumerate(lines):
        new_lines.append(line)
        if "const weeks = [" in line:
            in_weeks = True
        if in_weeks and "week:" in line:
            last_entry_idx = len(new_lines) - 1
        if in_weeks and "];" in line:
            in_weeks = False
            if not inserted and last_entry_idx >= 0:
                # Insert after the last entry
                entry_line = f'  {{ week: "W{week_num}", year: 2026, markets: ["AE", "AU", "SA"] }},'
                new_lines.insert(last_entry_idx + 1, entry_line)
                inserted = True

    if inserted:
        index_path.write_text("\n".join(new_lines), encoding="utf-8")
        print(f"  Added W{week_num} to index.html")
    else:
        print(f"  Warning: Could not find insertion point in index.html")


# ── Main ───────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("WBR Report Generator")
    print("=" * 60)

    # 1. Find latest week folder
    week_num, week_dir = find_latest_week_folder()
    print(f"\n[1] Latest week folder: {week_dir.name} (Week {week_num})")

    # 2. Load data
    print(f"\n[2] Loading data...")
    rows = load_data(week_num, week_dir)

    # 3. Determine current week from data
    weeks_2026 = set()
    for r in rows:
        if safe_int(r[COL_YEAR]) == 2026:
            w = safe_int(r[COL_WEEK])
            if w > 0:
                weeks_2026.add(w)
    current_week = max(weeks_2026) if weeks_2026 else week_num
    print(f"  Current week (from data): W{current_week}")

    # 4. Generate reports for each market
    output_dir = Path(f"wbr/W{current_week}")
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"\n[3] Output directory: {output_dir}")

    for mkt_id, (mkt_code, mkt_label, mkt_short) in MARKETS.items():
        print(f"\n[4] Generating {mkt_code} report...")
        filtered = filter_data(rows, mkt_id)
        print(f"  Filtered rows: {len(filtered)}")

        if not filtered:
            print(f"  WARNING: No data for {mkt_code} (marketplace_id={mkt_id}). Skipping.")
            continue

        data = build_datasets(filtered, current_week)

        # Try to load W-2 data from previous week's xlsx for Deep Dive bars
        pw2_week = current_week - 2
        prev_xlsx_dir = Path(f"W{current_week - 1}")
        prev_xlsx_name = f"WBR page 0 MCID data_weekly_w{current_week - 1}_2026.xlsx"
        prev_xlsx_path = prev_xlsx_dir / prev_xlsx_name
        if prev_xlsx_path.exists():
            try:
                if not hasattr(main, '_prev_rows'):
                    print(f"  Loading previous xlsx for W-2 data: {prev_xlsx_path}")
                    main._prev_rows = load_data(current_week - 1, prev_xlsx_dir)
                prev_filtered = filter_data(main._prev_rows, mkt_id)
                # Get W-2 sellers from previous xlsx
                pw2_rows = [r for r in prev_filtered
                            if safe_int(r[COL_YEAR]) == 2026 and safe_int(r[COL_WEEK]) == pw2_week]
                pw2_sellers = {}
                for r in pw2_rows:
                    mcid = safe_str(r[COL_MCID])
                    if not mcid:
                        continue
                    if mcid not in pw2_sellers:
                        pw2_sellers[mcid] = {"wtd_gms": 0}
                    pw2_sellers[mcid]["wtd_gms"] += safe_float(r[COL_WTD_GMS])
                data["pw2_sellers"] = pw2_sellers
            except Exception as e:
                print(f"  Note: Could not load W-2 data: {e}")

        html_content = generate_html(data, mkt_code, mkt_label, current_week)

        output_file = output_dir / f"W{current_week}_WBR_{mkt_code}_Pipeline.html"
        output_file.write_text(html_content, encoding="utf-8")
        print(f"  Written: {output_file}")

    # 5. Update index
    print(f"\n[5] Updating index...")
    update_index(current_week)

    print(f"\n{'=' * 60}")
    print(f"Done! Reports generated in {output_dir}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
