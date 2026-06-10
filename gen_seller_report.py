#!/usr/bin/env python3
"""
Generate WBR W23 2026 Seller Report Dashboard (seller-report.html)
Reads: W23/WBR page 0 MCID data_weekly_w23_2026.xlsx
Outputs: seller-report.html (encrypted with CryptoJS-compatible AES)
"""

import os
import sys
import hashlib
import base64
from datetime import datetime
from pathlib import Path

import openpyxl

# ─── Configuration ───────────────────────────────────────────────────────
CURRENT_WEEK = 23
PREV_WEEK = 22
CURRENT_YEAR = 2026
LAST_YEAR = 2025
PASSWORD = "expansionwbr"

BASE_DIR = Path(r"C:\Users\chiawenk\Documents\AU Launch Preparation\AU WBR")
XLSX_PATH = BASE_DIR / "W23" / "WBR page 0 MCID data_weekly_w23_2026.xlsx"
OUTPUT_PATH = BASE_DIR / "seller-report.html"

# Column indices (0-based)
COL_CAL_TYPE = 0
COL_REPORTING_YEAR = 1
COL_REPORTING_WEEK = 2
COL_SELLER_ORIGIN = 3
COL_MARKETPLACE_ID = 10
COL_LAUNCH_CHANNEL = 12
COL_MCID = 21
COL_MERCHANT_NAME = 25
COL_OWNER = 26
COL_LAUNCH_DATE = 28
COL_CATEGORY = 34
COL_SP_PG = 36
COL_WTD_LAUNCH = 53
COL_YTD_LAUNCH = 54
COL_WTD_ACTIVE = 63
COL_WTD_GMS = 90
COL_WTD_FBA_GMS = 91
COL_WTD_UNITS = 93
COL_WTD_FBA_UNITS = 94
COL_YTD_GMS = 96
COL_YTD_FBA_GMS = 97
COL_YTD_UNITS = 99

# Marketplace IDs (verified from data)
MP_IDS = {
    'AU': [111172],
    'AE': [338801],
    'SA': [338811],
    'JP': [6],
    'UK': [3],
    'DE': [4],
    'FR': [5],
    'IT': [35691],
    'ES': [44551],
    'NL': [44571],
    'SE': [328451],
    'BE': [338851],
}

# Actually let me build from data - AE marketplace_id
# From the reference: AU=111172, AE and SA might share a region

# EU5 = UK + DE + FR + IT + ES
EU5_MPS = ['UK', 'DE', 'FR', 'IT', 'ES']


def safe_float(v):
    """Convert cell value to float, handling None/empty/str"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def safe_str(v):
    """Convert cell value to string"""
    if v is None:
        return ""
    return str(v).strip()


def load_data():
    """Load xlsx and return list of row dicts"""
    print(f"Loading {XLSX_PATH}...")
    wb = openpyxl.load_workbook(str(XLSX_PATH), read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        # Skip header
        if row[COL_CAL_TYPE] != "SUMUP":
            continue
        rows.append(row)
    wb.close()
    print(f"  Loaded {len(rows)} SUMUP rows")
    return rows


def filter_tw(rows):
    """Filter rows where seller_origin = TW"""
    return [r for r in rows if safe_str(r[COL_SELLER_ORIGIN]).upper() == "TW"]


def get_week_data(rows, year, week):
    """Filter rows for a specific year+week"""
    result = []
    for r in rows:
        ry = r[COL_REPORTING_YEAR]
        rw = r[COL_REPORTING_WEEK]
        if (isinstance(ry, (int, float)) and int(ry) == year and
            isinstance(rw, (int, float)) and int(rw) == week):
            result.append(r)
    return result


def get_marketplace_id(row):
    """Get marketplace_id from row"""
    v = row[COL_MARKETPLACE_ID]
    if v is None:
        return None
    try:
        return int(float(v))
    except:
        return None


def get_channel(row):
    """Get launch_channel (DSR/SSR/ESM)"""
    return safe_str(row[COL_LAUNCH_CHANNEL]).upper()


def fmt_money(v):
    """Format as $XX,XXX"""
    v = round(v)
    if v < 0:
        return f"$-{abs(v):,}"
    return f"${v:,}"


def fmt_pct(v, with_arrow=False):
    """Format as percentage with color coding"""
    if v is None:
        return '<span class="na">N/A</span>'
    if abs(v) < 0.05:
        if with_arrow:
            return f'<span class="flat">0.0%</span>'
        return f'<span class="flat">0.0%</span>'
    if v > 0:
        if with_arrow:
            return f'<span class="up">&#9650; {v:.1f}%</span>'
        return f'<span class="up">+{v:.1f}%</span>'
    else:
        if with_arrow:
            return f'<span class="down">&#9660; {abs(v):.1f}%</span>'
        return f'<span class="down">{v:.1f}%</span>'


def fmt_pct_exp(v):
    """Format percentage for expansion table (no arrow)"""
    if v is None:
        return '<span class="na">-</span>'
    if abs(v) < 0.05:
        return '<span class="flat">0.0%</span>'
    if v > 0:
        return f'<span class="up">{v:.1f}%</span>'
    else:
        return f'<span class="down">{v:.1f}%</span>'


def calc_wow(cur, prev):
    """Calculate WoW%"""
    if prev == 0:
        return None
    return (cur - prev) / prev * 100


def calc_yoy(cur, ly):
    """Calculate YoY%"""
    if ly == 0:
        return None
    return (cur - ly) / ly * 100


def html_escape(s):
    """Basic HTML escape"""
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#x27;"))


# ─── Build marketplace data structures ───────────────────────────────────

def build_seller_data(tw_rows):
    """
    Build seller-level data from TW-filtered rows.
    Returns dict keyed by (marketplace_id, mcid) -> seller info dict
    """
    # Get current week, prev week, last year same week data
    cur_data = get_week_data(tw_rows, CURRENT_YEAR, CURRENT_WEEK)
    prev_data = get_week_data(tw_rows, CURRENT_YEAR, PREV_WEEK)
    ly_data = get_week_data(tw_rows, LAST_YEAR, CURRENT_WEEK)

    print(f"  Current week (Y{CURRENT_YEAR} W{CURRENT_WEEK}): {len(cur_data)} rows")
    print(f"  Prev week (Y{CURRENT_YEAR} W{PREV_WEEK}): {len(prev_data)} rows")
    print(f"  Last year (Y{LAST_YEAR} W{CURRENT_WEEK}): {len(ly_data)} rows")

    # Build lookup for prev week and last year
    prev_lookup = {}
    for r in prev_data:
        mp = get_marketplace_id(r)
        mcid = safe_str(r[COL_MCID])
        if mp and mcid:
            prev_lookup[(mp, mcid)] = r

    ly_lookup = {}
    for r in ly_data:
        mp = get_marketplace_id(r)
        mcid = safe_str(r[COL_MCID])
        if mp and mcid:
            ly_lookup[(mp, mcid)] = r

    # Build seller records from current week
    sellers = {}
    for r in cur_data:
        mp = get_marketplace_id(r)
        mcid = safe_str(r[COL_MCID])
        if not mp or not mcid:
            continue

        key = (mp, mcid)
        channel = get_channel(r)
        name = safe_str(r[COL_MERCHANT_NAME])
        owner = safe_str(r[COL_OWNER])
        category = safe_str(r[COL_CATEGORY])
        launch_date = safe_str(r[COL_LAUNCH_DATE])
        wtd_gms = safe_float(r[COL_WTD_GMS])
        wtd_units = safe_float(r[COL_WTD_UNITS])
        ytd_gms = safe_float(r[COL_YTD_GMS])
        ytd_units = safe_float(r[COL_YTD_UNITS])
        wtd_launch = safe_float(r[COL_WTD_LAUNCH])
        ytd_launch = safe_float(r[COL_YTD_LAUNCH])
        wtd_active = safe_float(r[COL_WTD_ACTIVE])

        # Get prev week data
        prev_r = prev_lookup.get(key)
        prev_gms = safe_float(prev_r[COL_WTD_GMS]) if prev_r else 0
        prev_units = safe_float(prev_r[COL_WTD_UNITS]) if prev_r else 0

        # Get last year data
        ly_r = ly_lookup.get(key)
        ly_gms = safe_float(ly_r[COL_WTD_GMS]) if ly_r else 0
        ly_ytd_gms = safe_float(ly_r[COL_YTD_GMS]) if ly_r else 0

        sellers[key] = {
            'mp': mp,
            'mcid': mcid,
            'channel': channel,
            'name': name,
            'owner': owner,
            'category': category,
            'launch_date': launch_date,
            'wtd_gms': wtd_gms,
            'wtd_units': wtd_units,
            'ytd_gms': ytd_gms,
            'ytd_units': ytd_units,
            'wtd_launch': wtd_launch,
            'ytd_launch': ytd_launch,
            'wtd_active': wtd_active,
            'prev_gms': prev_gms,
            'prev_units': prev_units,
            'ly_gms': ly_gms,
            'ly_ytd_gms': ly_ytd_gms,
        }

    # Also include sellers from prev_week that are NOT in current week
    for key, r in prev_lookup.items():
        if key not in sellers:
            mp, mcid = key
            channel = get_channel(r)
            name = safe_str(r[COL_MERCHANT_NAME])
            owner = safe_str(r[COL_OWNER])
            category = safe_str(r[COL_CATEGORY])
            launch_date = safe_str(r[COL_LAUNCH_DATE])
            ytd_launch_val = safe_float(r[COL_YTD_LAUNCH])

            # For sellers only in prev week, their current week values are 0
            # but we need to pull their ytd from current week data (they might still have ytd > 0)
            # Actually if they're not in cur_data, their current week stats are 0
            ly_r = ly_lookup.get(key)
            ly_gms = safe_float(ly_r[COL_WTD_GMS]) if ly_r else 0
            ly_ytd_gms = safe_float(ly_r[COL_YTD_GMS]) if ly_r else 0

            sellers[key] = {
                'mp': mp,
                'mcid': mcid,
                'channel': channel,
                'name': name,
                'owner': owner,
                'category': category,
                'launch_date': launch_date,
                'wtd_gms': 0,
                'wtd_units': 0,
                'ytd_gms': safe_float(r[COL_YTD_GMS]),  # use prev week's YTD as approx
                'ytd_units': safe_float(r[COL_YTD_UNITS]),
                'wtd_launch': 0,
                'ytd_launch': ytd_launch_val,
                'wtd_active': 0,
                'prev_gms': safe_float(r[COL_WTD_GMS]),
                'prev_units': safe_float(r[COL_WTD_UNITS]),
                'ly_gms': ly_gms,
                'ly_ytd_gms': ly_ytd_gms,
            }

    # Include last-year-only sellers for completeness in YoY calculations
    for key, r in ly_lookup.items():
        if key not in sellers:
            mp, mcid = key
            channel = get_channel(r)
            name = safe_str(r[COL_MERCHANT_NAME])
            owner = safe_str(r[COL_OWNER])
            category = safe_str(r[COL_CATEGORY])
            launch_date = safe_str(r[COL_LAUNCH_DATE])

            sellers[key] = {
                'mp': mp,
                'mcid': mcid,
                'channel': channel,
                'name': name,
                'owner': owner,
                'category': category,
                'launch_date': launch_date,
                'wtd_gms': 0,
                'wtd_units': 0,
                'ytd_gms': 0,
                'ytd_units': 0,
                'wtd_launch': 0,
                'ytd_launch': 0,
                'wtd_active': 0,
                'prev_gms': 0,
                'prev_units': 0,
                'ly_gms': safe_float(r[COL_WTD_GMS]),
                'ly_ytd_gms': safe_float(r[COL_YTD_GMS]),
            }

    print(f"  Total seller records: {len(sellers)}")
    return sellers


def discover_marketplace_ids(sellers):
    """Discover actual marketplace IDs from data"""
    mp_ids = set()
    for (mp, mcid), s in sellers.items():
        mp_ids.add(mp)
    return sorted(mp_ids)


def get_sellers_for_mp(sellers, mp_ids, channels=None, nsr_only_ytd=False):
    """Get sellers for given marketplace IDs and optional channel filter"""
    result = []
    for (mp, mcid), s in sellers.items():
        if mp not in mp_ids:
            continue
        if channels and s['channel'] not in channels:
            continue
        if nsr_only_ytd and s['ytd_launch'] <= 0:
            continue
        result.append(s)
    return result


# ─── Discover marketplace IDs from actual data ──────────────────────────

def auto_detect_mp_ids(sellers):
    """Detect marketplace IDs by looking at the data"""
    # We'll group by marketplace_id and check what we have
    mp_counts = {}
    for (mp, mcid), s in sellers.items():
        if mp not in mp_counts:
            mp_counts[mp] = 0
        mp_counts[mp] += 1

    print(f"  Marketplace IDs found: {sorted(mp_counts.keys())}")
    for mp_id in sorted(mp_counts.keys()):
        print(f"    MP {mp_id}: {mp_counts[mp_id]} sellers")
    return mp_counts


# ─── Expansion DSR Summary Table ────────────────────────────────────────

def build_expansion_summary(sellers, tw_rows):
    """Build the EXP summary table data"""
    # We need aggregate data for each region/IC

    # Helper: sum metrics for a subset of sellers
    def agg(seller_list):
        wtd_gms = sum(s['wtd_gms'] for s in seller_list)
        prev_gms = sum(s['prev_gms'] for s in seller_list)
        ly_gms = sum(s['ly_gms'] for s in seller_list)
        ytd_gms = sum(s['ytd_gms'] for s in seller_list)
        ly_ytd_gms = sum(s['ly_ytd_gms'] for s in seller_list)
        wtd_launch = sum(s['wtd_launch'] for s in seller_list)
        prev_launch = 0  # We'll compute from prev week data
        ytd_launch = 0  # Max ytd_launch among sellers
        # For launches, count sellers with wtd_launch > 0
        wtd_launch_count = sum(1 for s in seller_list if s['wtd_launch'] > 0)
        # For YTD launch, we use sum of ytd_launch (which represents cumulative unique)
        # Actually ytd_launch at the row level means "this seller is a YTD launch (0 or 1)"
        ytd_launch_count = sum(1 for s in seller_list if s['ytd_launch'] > 0)
        return {
            'wtd_gms': wtd_gms,
            'prev_gms': prev_gms,
            'ly_gms': ly_gms,
            'ytd_gms': ytd_gms,
            'ly_ytd_gms': ly_ytd_gms,
            'wtd_launch': wtd_launch_count,
            'ytd_launch': ytd_launch_count,
        }

    # We need prev_week wtd_launch counts too
    # Let's compute from tw_rows directly
    prev_week_rows = get_week_data(tw_rows, CURRENT_YEAR, PREV_WEEK)
    ly_week_rows = get_week_data(tw_rows, LAST_YEAR, CURRENT_WEEK)

    def count_launches_in_rows(rows, mp_ids, channels):
        """Count wtd_launch and ytd_launch from raw rows"""
        wtd = 0
        ytd = 0
        seen_ytd = set()
        for r in rows:
            mp = get_marketplace_id(r)
            if mp not in mp_ids:
                continue
            ch = get_channel(r)
            if channels and ch not in channels:
                continue
            mcid = safe_str(r[COL_MCID])
            wl = safe_float(r[COL_WTD_LAUNCH])
            yl = safe_float(r[COL_YTD_LAUNCH])
            if wl > 0:
                wtd += 1
            if yl > 0 and mcid not in seen_ytd:
                seen_ytd.add(mcid)
                ytd += 1
        return wtd, ytd

    def count_soa_in_rows(rows, mp_ids, channels):
        """Count SOA launches (wtd_active and ytd equivalents) - using same fields"""
        # SOA = Selection on Amazon. From the reference, SOA columns appear to track same as seller launch
        # Actually in the reference table, SOA seems to be separate but uses same structure
        # Looking at the W22 data: SOA tracks marketplace-level selection count
        # For simplicity, we'll use wtd_active for SOA weekly, and a ytd version
        # Actually from the data columns provided, there's no separate SOA column
        # Looking at the W22 reference more carefully:
        # SOA Launch = selection/offering activated (could be same seller launching in multiple MPs)
        # The data might count per-marketplace appearances
        # Let's use the same launch data but count per (mcid, mp) pair for SOA
        wtd = 0
        ytd = 0
        for r in rows:
            mp = get_marketplace_id(r)
            if mp not in mp_ids:
                continue
            ch = get_channel(r)
            if channels and ch not in channels:
                continue
            wl = safe_float(r[COL_WTD_LAUNCH])
            yl = safe_float(r[COL_YTD_LAUNCH])
            if wl > 0:
                wtd += 1
            if yl > 0:
                ytd += 1
        return wtd, ytd

    # Define regions
    # EU5 marketplace IDs: UK=3, DE=4, FR=5, IT=35691, ES=44551
    eu5_ids = [3, 4, 5, 35691, 44551]
    # Additional EU: NL=44571, PL=?, SE=328451, BE=338851
    # JP=6, AU=111172
    # AE and SA: need to discover from data
    # From the reference, AE uses one ID and SA uses another
    # Let's check - from reference: AU ID: 111172, JP ID: 6
    # AE and SA might be 771771 and 771770 or different

    # Actually let me look at the marketplace IDs from actual data
    # We'll map them dynamically

    au_ids = [111172]
    jp_ids = [6]

    # For AE/SA we need to discover - common IDs are:
    # AE (UAE) = 771771 or similar, SA = 771770
    # Let me check what's in the data that's NOT AU/JP/EU5
    known_ids = set(eu5_ids + au_ids + jp_ids + [44571, 328451, 338851])
    unknown_ids = set()
    for (mp, mcid), s in sellers.items():
        if mp not in known_ids:
            unknown_ids.add(mp)

    # AE (UAE) = 338801, SA (Saudi Arabia) = 338811 (verified from owner analysis)
    ae_ids = [338801]
    sa_ids = [338811]

    print(f"  Using MP IDs: AU={au_ids}, JP={jp_ids}, AE={ae_ids}, SA={sa_ids}")
    print(f"  EU5 IDs: {eu5_ids}")

    mena_ids = ae_ids + sa_ids

    # Get current week rows for launch counting
    cur_week_rows = get_week_data(tw_rows, CURRENT_YEAR, CURRENT_WEEK)

    # Build row data for the expansion table
    # Each row: label, is_mp_row (vs ic_row), metrics
    def make_row(label, mp_ids_list, channels, is_esm=False):
        """Compute expansion row data"""
        # Get sellers matching criteria
        cur_sellers = []
        for (mp, mcid), s in sellers.items():
            if mp not in mp_ids_list:
                continue
            ch = s['channel']
            if channels and ch not in channels:
                continue
            cur_sellers.append(s)

        wtd_gms = sum(s['wtd_gms'] for s in cur_sellers)
        prev_gms = sum(s['prev_gms'] for s in cur_sellers)
        ly_gms = sum(s['ly_gms'] for s in cur_sellers)
        ytd_gms = sum(s['ytd_gms'] for s in cur_sellers)
        ly_ytd_gms = sum(s['ly_ytd_gms'] for s in cur_sellers)

        wow_gms = calc_wow(wtd_gms, prev_gms)
        yoy_gms = calc_yoy(wtd_gms, ly_gms)
        ytd_yoy_gms = calc_yoy(ytd_gms, ly_ytd_gms)

        if is_esm:
            # ESM rows don't show seller/SOA launch
            return {
                'label': label,
                'prev_gms': prev_gms,
                'wtd_gms': wtd_gms,
                'wow_gms': wow_gms,
                'yoy_gms': yoy_gms,
                'ytd_gms': ytd_gms,
                'ytd_yoy_gms': ytd_yoy_gms,
                'prev_launch': '-',
                'wtd_launch': '-',
                'wow_launch': None,
                'ytd_launch': '-',
                'ytd_yoy_launch': None,
                'prev_soa': '-',
                'wtd_soa': '-',
                'wow_soa': None,
                'ytd_soa': '-',
                'ytd_yoy_soa': None,
                'is_esm': True,
            }

        # Count launches
        # Current week launches
        wtd_launch_cur, ytd_launch_cur = count_launches_in_rows(cur_week_rows, mp_ids_list, channels)
        wtd_launch_prev, ytd_launch_prev = count_launches_in_rows(prev_week_rows, mp_ids_list, channels)
        wtd_launch_ly, ytd_launch_ly = count_launches_in_rows(ly_week_rows, mp_ids_list, channels)

        # SOA (same as launch counting per marketplace)
        wtd_soa_cur, ytd_soa_cur = count_soa_in_rows(cur_week_rows, mp_ids_list, channels)
        wtd_soa_prev, ytd_soa_prev = count_soa_in_rows(prev_week_rows, mp_ids_list, channels)
        wtd_soa_ly, ytd_soa_ly = count_soa_in_rows(ly_week_rows, mp_ids_list, channels)

        wow_launch = calc_wow(wtd_launch_cur, wtd_launch_prev)
        ytd_yoy_launch = calc_yoy(ytd_launch_cur, ytd_launch_ly)

        wow_soa = calc_wow(wtd_soa_cur, wtd_soa_prev)
        ytd_yoy_soa = calc_yoy(ytd_soa_cur, ytd_soa_ly)

        return {
            'label': label,
            'prev_gms': prev_gms,
            'wtd_gms': wtd_gms,
            'wow_gms': wow_gms,
            'yoy_gms': yoy_gms,
            'ytd_gms': ytd_gms,
            'ytd_yoy_gms': ytd_yoy_gms,
            'prev_launch': wtd_launch_prev,
            'wtd_launch': wtd_launch_cur,
            'wow_launch': wow_launch,
            'ytd_launch': ytd_launch_cur,
            'ytd_yoy_launch': ytd_yoy_launch,
            'prev_soa': wtd_soa_prev,
            'wtd_soa': wtd_soa_cur,
            'wow_soa': wow_soa,
            'ytd_soa': ytd_soa_cur,
            'ytd_yoy_soa': ytd_yoy_soa,
            'is_esm': False,
        }

    def make_ic_row(label, mp_ids_list, channels, owner_name):
        """Compute IC-specific row filtered by owner"""
        cur_sellers = []
        for (mp, mcid), s in sellers.items():
            if mp not in mp_ids_list:
                continue
            ch = s['channel']
            if channels and ch not in channels:
                continue
            if s['owner'] != owner_name:
                continue
            cur_sellers.append(s)

        wtd_gms = sum(s['wtd_gms'] for s in cur_sellers)
        prev_gms = sum(s['prev_gms'] for s in cur_sellers)
        ly_gms = sum(s['ly_gms'] for s in cur_sellers)
        ytd_gms = sum(s['ytd_gms'] for s in cur_sellers)
        ly_ytd_gms = sum(s['ly_ytd_gms'] for s in cur_sellers)

        wow_gms = calc_wow(wtd_gms, prev_gms)
        yoy_gms = calc_yoy(wtd_gms, ly_gms)
        ytd_yoy_gms = calc_yoy(ytd_gms, ly_ytd_gms)

        # Count launches for this owner
        def count_owner_launches(rows, mp_list, ch_list, owner):
            wtd = 0
            ytd_set = set()
            for r in rows:
                mp = get_marketplace_id(r)
                if mp not in mp_list:
                    continue
                ch = get_channel(r)
                if ch_list and ch not in ch_list:
                    continue
                ow = safe_str(r[COL_OWNER])
                if ow != owner:
                    continue
                mcid = safe_str(r[COL_MCID])
                wl = safe_float(r[COL_WTD_LAUNCH])
                yl = safe_float(r[COL_YTD_LAUNCH])
                if wl > 0:
                    wtd += 1
                if yl > 0 and mcid not in ytd_set:
                    ytd_set.add(mcid)
                    ytd += 1  # noqa - this would be wrong
            return wtd, len(ytd_set)

        def count_owner_launches_fixed(rows, mp_list, ch_list, owner):
            wtd = 0
            ytd_set = set()
            for r in rows:
                mp = get_marketplace_id(r)
                if mp not in mp_list:
                    continue
                ch = get_channel(r)
                if ch_list and ch not in ch_list:
                    continue
                ow = safe_str(r[COL_OWNER])
                if ow != owner:
                    continue
                mcid = safe_str(r[COL_MCID])
                wl = safe_float(r[COL_WTD_LAUNCH])
                yl = safe_float(r[COL_YTD_LAUNCH])
                if wl > 0:
                    wtd += 1
                if yl > 0 and mcid not in ytd_set:
                    ytd_set.add(mcid)
            return wtd, len(ytd_set)

        def count_owner_soa(rows, mp_list, ch_list, owner):
            wtd = 0
            ytd = 0
            for r in rows:
                mp = get_marketplace_id(r)
                if mp not in mp_list:
                    continue
                ch = get_channel(r)
                if ch_list and ch not in ch_list:
                    continue
                ow = safe_str(r[COL_OWNER])
                if ow != owner:
                    continue
                wl = safe_float(r[COL_WTD_LAUNCH])
                yl = safe_float(r[COL_YTD_LAUNCH])
                if wl > 0:
                    wtd += 1
                if yl > 0:
                    ytd += 1
            return wtd, ytd

        wtd_launch_cur, ytd_launch_cur = count_owner_launches_fixed(cur_week_rows, mp_ids_list, channels, owner_name)
        wtd_launch_prev, _ = count_owner_launches_fixed(prev_week_rows, mp_ids_list, channels, owner_name)
        _, ytd_launch_ly = count_owner_launches_fixed(ly_week_rows, mp_ids_list, channels, owner_name)

        wtd_soa_cur, ytd_soa_cur = count_owner_soa(cur_week_rows, mp_ids_list, channels, owner_name)
        wtd_soa_prev, _ = count_owner_soa(prev_week_rows, mp_ids_list, channels, owner_name)
        _, ytd_soa_ly = count_owner_soa(ly_week_rows, mp_ids_list, channels, owner_name)

        wow_launch = calc_wow(wtd_launch_cur, wtd_launch_prev)
        ytd_yoy_launch = calc_yoy(ytd_launch_cur, ytd_launch_ly)
        wow_soa = calc_wow(wtd_soa_cur, wtd_soa_prev)
        ytd_yoy_soa = calc_yoy(ytd_soa_cur, ytd_soa_ly)

        return {
            'label': label,
            'prev_gms': prev_gms,
            'wtd_gms': wtd_gms,
            'wow_gms': wow_gms,
            'yoy_gms': yoy_gms,
            'ytd_gms': ytd_gms,
            'ytd_yoy_gms': ytd_yoy_gms,
            'prev_launch': wtd_launch_prev,
            'wtd_launch': wtd_launch_cur,
            'wow_launch': wow_launch,
            'ytd_launch': ytd_launch_cur,
            'ytd_yoy_launch': ytd_yoy_launch,
            'prev_soa': wtd_soa_prev,
            'wtd_soa': wtd_soa_cur,
            'wow_soa': wow_soa,
            'ytd_soa': ytd_soa_cur,
            'ytd_yoy_soa': ytd_yoy_soa,
            'is_esm': False,
        }

    # Build rows
    # EU = DSR only for EU5
    eu_dsr = ['DSR']
    # JP = DSR only
    jp_dsr = ['DSR']
    # AU/MENA = DSR + SSR
    au_mena_ch = ['DSR', 'SSR']

    exp_rows = []
    # 1. EU
    exp_rows.append(('mp', make_row('EU', eu5_ids, eu_dsr)))
    # 2. Eddie Chu (EU)
    exp_rows.append(('ic', make_ic_row('Eddie Chu', eu5_ids, eu_dsr, 'Eddie Chu')))
    # 3. Jerry Kan (EU)
    exp_rows.append(('ic', make_ic_row('Jerry Kan', eu5_ids, eu_dsr, 'Jerry Kan')))
    # 4. JP
    exp_rows.append(('mp', make_row('JP', jp_ids, jp_dsr)))
    # 5. Shelly Huang (JP)
    exp_rows.append(('ic', make_ic_row('Shelly Huang', jp_ids, jp_dsr, 'Shelly Huang')))
    # 6. Silvia Lien (JP)
    exp_rows.append(('ic', make_ic_row('Silvia Lien', jp_ids, jp_dsr, 'Silvia Lien')))
    # 7. AU (DSR+SSR)
    exp_rows.append(('mp', make_row('AU', au_ids, au_mena_ch)))
    # 8. Jenny Kao (AU)
    exp_rows.append(('ic', make_ic_row('Jenny Kao', au_ids, au_mena_ch, 'Jenny Kao')))
    # 9. AU ESM
    exp_rows.append(('mp', make_row('AU ESM', au_ids, ['ESM'], is_esm=True)))
    # 10. MENA (AE+SA DSR+SSR)
    exp_rows.append(('mp', make_row('MENA', mena_ids, au_mena_ch)))
    # 11. AE (DSR+SSR)
    exp_rows.append(('mp', make_row('AE', ae_ids, au_mena_ch)))
    # 12. Jenny Kao (AE)
    exp_rows.append(('ic', make_ic_row('Jenny Kao', ae_ids, au_mena_ch, 'Jenny Kao')))
    # 13. SA (DSR+SSR)
    exp_rows.append(('mp', make_row('SA', sa_ids, au_mena_ch)))
    # 14. Jenny Kao (SA)
    exp_rows.append(('ic', make_ic_row('Jenny Kao', sa_ids, au_mena_ch, 'Jenny Kao')))
    # 15. MENA ESM
    exp_rows.append(('mp', make_row('MENA ESM', mena_ids, ['ESM'], is_esm=True)))

    # 16. Summary: DSR Expansion (Without AU/MENA ESM)
    # = EU(DSR) + JP(DSR) + AU(DSR+SSR) + MENA(DSR+SSR)
    dsr_all_ids = eu5_ids + jp_ids + au_ids + mena_ids
    sum_row = make_row('DSR Expansion (Without AU/MENA ESM)', dsr_all_ids, ['DSR', 'SSR'])
    # But EU and JP are DSR only - need to combine properly
    # Actually let's just sum the component rows
    components_excl_esm = [exp_rows[0][1], exp_rows[3][1], exp_rows[6][1], exp_rows[9][1]]  # EU, JP, AU, MENA
    sum16 = {
        'label': 'DSR Expansion (Without AU/MENA ESM)',
        'prev_gms': sum(r['prev_gms'] for r in components_excl_esm),
        'wtd_gms': sum(r['wtd_gms'] for r in components_excl_esm),
        'ytd_gms': sum(r['ytd_gms'] for r in components_excl_esm),
        'is_esm': False,
    }
    sum16['wow_gms'] = calc_wow(sum16['wtd_gms'], sum16['prev_gms'])
    ly_gms_sum = sum(sum(s['ly_gms'] for (mp, mcid), s in sellers.items()
                        if mp in dsr_all_ids and s['channel'] in ['DSR', 'SSR'])
                     for _ in [1])  # flatten
    # Recompute properly
    ly_gms_16 = 0
    ly_ytd_16 = 0
    for (mp, mcid), s in sellers.items():
        if mp in eu5_ids and s['channel'] == 'DSR':
            ly_gms_16 += s['ly_gms']
            ly_ytd_16 += s['ly_ytd_gms']
        elif mp in jp_ids and s['channel'] == 'DSR':
            ly_gms_16 += s['ly_gms']
            ly_ytd_16 += s['ly_ytd_gms']
        elif mp in au_ids and s['channel'] in ['DSR', 'SSR']:
            ly_gms_16 += s['ly_gms']
            ly_ytd_16 += s['ly_ytd_gms']
        elif mp in mena_ids and s['channel'] in ['DSR', 'SSR']:
            ly_gms_16 += s['ly_gms']
            ly_ytd_16 += s['ly_ytd_gms']
    sum16['yoy_gms'] = calc_yoy(sum16['wtd_gms'], ly_gms_16)
    sum16['ytd_yoy_gms'] = calc_yoy(sum16['ytd_gms'], ly_ytd_16)

    # Launches
    sum16['prev_launch'] = sum(r['prev_launch'] for r in components_excl_esm if isinstance(r['prev_launch'], (int, float)))
    sum16['wtd_launch'] = sum(r['wtd_launch'] for r in components_excl_esm if isinstance(r['wtd_launch'], (int, float)))
    sum16['wow_launch'] = calc_wow(sum16['wtd_launch'], sum16['prev_launch'])
    sum16['ytd_launch'] = sum(r['ytd_launch'] for r in components_excl_esm if isinstance(r['ytd_launch'], (int, float)))
    ytd_launch_ly_16 = 0
    for r in components_excl_esm:
        if isinstance(r.get('ytd_yoy_launch'), (int, float)) and r['ytd_yoy_launch'] is not None:
            # back-calculate ly from yoy
            pass
    # Use count from ly rows
    _, ytd_l_eu_ly = count_launches_in_rows(ly_week_rows, eu5_ids, eu_dsr)
    _, ytd_l_jp_ly = count_launches_in_rows(ly_week_rows, jp_ids, jp_dsr)
    _, ytd_l_au_ly = count_launches_in_rows(ly_week_rows, au_ids, au_mena_ch)
    _, ytd_l_mena_ly = count_launches_in_rows(ly_week_rows, mena_ids, au_mena_ch)
    ytd_launch_ly_16 = ytd_l_eu_ly + ytd_l_jp_ly + ytd_l_au_ly + ytd_l_mena_ly
    sum16['ytd_yoy_launch'] = calc_yoy(sum16['ytd_launch'], ytd_launch_ly_16)

    # SOA
    sum16['prev_soa'] = sum(r['prev_soa'] for r in components_excl_esm if isinstance(r['prev_soa'], (int, float)))
    sum16['wtd_soa'] = sum(r['wtd_soa'] for r in components_excl_esm if isinstance(r['wtd_soa'], (int, float)))
    sum16['wow_soa'] = calc_wow(sum16['wtd_soa'], sum16['prev_soa'])
    sum16['ytd_soa'] = sum(r['ytd_soa'] for r in components_excl_esm if isinstance(r['ytd_soa'], (int, float)))
    _, ytd_s_eu_ly = count_soa_in_rows(ly_week_rows, eu5_ids, eu_dsr)
    _, ytd_s_jp_ly = count_soa_in_rows(ly_week_rows, jp_ids, jp_dsr)
    _, ytd_s_au_ly = count_soa_in_rows(ly_week_rows, au_ids, au_mena_ch)
    _, ytd_s_mena_ly = count_soa_in_rows(ly_week_rows, mena_ids, au_mena_ch)
    ytd_soa_ly_16 = ytd_s_eu_ly + ytd_s_jp_ly + ytd_s_au_ly + ytd_s_mena_ly
    sum16['ytd_yoy_soa'] = calc_yoy(sum16['ytd_soa'], ytd_soa_ly_16)

    exp_rows.append(('summary', sum16))

    # 17. DSR Expansion + AU/MENA ESM
    au_esm_row = exp_rows[8][1]  # AU ESM
    mena_esm_row = exp_rows[14][1]  # MENA ESM
    sum17 = {
        'label': 'DSR Expansion + AU/MENA ESM',
        'prev_gms': sum16['prev_gms'] + au_esm_row['prev_gms'] + mena_esm_row['prev_gms'],
        'wtd_gms': sum16['wtd_gms'] + au_esm_row['wtd_gms'] + mena_esm_row['wtd_gms'],
        'ytd_gms': sum16['ytd_gms'] + au_esm_row['ytd_gms'] + mena_esm_row['ytd_gms'],
        'is_esm': False,
    }
    sum17['wow_gms'] = calc_wow(sum17['wtd_gms'], sum17['prev_gms'])
    ly_gms_esm = 0
    ly_ytd_esm = 0
    for (mp, mcid), s in sellers.items():
        if mp in au_ids and s['channel'] == 'ESM':
            ly_gms_esm += s['ly_gms']
            ly_ytd_esm += s['ly_ytd_gms']
        elif mp in mena_ids and s['channel'] == 'ESM':
            ly_gms_esm += s['ly_gms']
            ly_ytd_esm += s['ly_ytd_gms']
    sum17['yoy_gms'] = calc_yoy(sum17['wtd_gms'], ly_gms_16 + ly_gms_esm)
    sum17['ytd_yoy_gms'] = calc_yoy(sum17['ytd_gms'], ly_ytd_16 + ly_ytd_esm)

    # Launches same as row 16 (ESM doesn't have launches)
    sum17['prev_launch'] = sum16['prev_launch']
    sum17['wtd_launch'] = sum16['wtd_launch']
    sum17['wow_launch'] = sum16['wow_launch']
    sum17['ytd_launch'] = sum16['ytd_launch']
    sum17['ytd_yoy_launch'] = sum16['ytd_yoy_launch']
    sum17['prev_soa'] = sum16['prev_soa']
    sum17['wtd_soa'] = sum16['wtd_soa']
    sum17['wow_soa'] = sum16['wow_soa']
    sum17['ytd_soa'] = sum16['ytd_soa']
    sum17['ytd_yoy_soa'] = sum16['ytd_yoy_soa']

    exp_rows.append(('summary2', sum17))

    return exp_rows, {
        'eu5_ids': eu5_ids,
        'jp_ids': jp_ids,
        'au_ids': au_ids,
        'ae_ids': ae_ids,
        'sa_ids': sa_ids,
        'mena_ids': mena_ids,
    }


def render_exp_row_html(idx, row_type, row_data):
    """Render a single expansion table row as HTML"""
    d = row_data
    css_class = {
        'mp': 'mp-row',
        'ic': 'ic-row',
        'summary': 'summary-row',
        'summary2': 'summary-row2',
    }[row_type]

    label = d['label']

    def fmt_launch_val(v):
        if isinstance(v, str):
            return v
        return str(int(v))

    is_esm = d.get('is_esm', False)

    # GMS columns
    gms_cells = (
        f'<td class="sep-gms">{fmt_money(d["prev_gms"])}</td>'
        f'<td>{fmt_money(d["wtd_gms"])}</td>'
        f'<td>{fmt_pct_exp(d["wow_gms"])}</td>'
        f'<td>{fmt_pct_exp(d["yoy_gms"])}</td>'
        f'<td>{fmt_money(d["ytd_gms"])}</td>'
        f'<td>{fmt_pct_exp(d["ytd_yoy_gms"])}</td>'
    )

    # Seller Launch columns
    if is_esm:
        launch_cells = (
            f'<td class="sep-seller">-</td><td>-</td>'
            f'<td><span class="na">-</span></td><td>-</td>'
            f'<td><span class="na">-</span></td>'
        )
        soa_cells = (
            f'<td class="sep-soa">-</td><td>-</td>'
            f'<td><span class="na">-</span></td><td>-</td>'
            f'<td><span class="na">-</span></td>'
        )
    else:
        launch_cells = (
            f'<td class="sep-seller">{fmt_launch_val(d["prev_launch"])}</td>'
            f'<td>{fmt_launch_val(d["wtd_launch"])}</td>'
            f'<td>{fmt_pct_exp(d["wow_launch"])}</td>'
            f'<td>{fmt_launch_val(d["ytd_launch"])}</td>'
            f'<td>{fmt_pct_exp(d["ytd_yoy_launch"])}</td>'
        )
        soa_cells = (
            f'<td class="sep-soa">{fmt_launch_val(d["prev_soa"])}</td>'
            f'<td>{fmt_launch_val(d["wtd_soa"])}</td>'
            f'<td>{fmt_pct_exp(d["wow_soa"])}</td>'
            f'<td>{fmt_launch_val(d["ytd_soa"])}</td>'
            f'<td>{fmt_pct_exp(d["ytd_yoy_soa"])}</td>'
        )

    return (
        f'<tr class="{css_class}">'
        f'<td class="center">{idx}</td><td class="left">{label}</td>'
        f'{gms_cells}{launch_cells}{soa_cells}'
        f'</tr>'
    )


# ─── Movers & Shakers ────────────────────────────────────────────────────

def build_movers_shakers(sellers, mp_ids_list, channels, mp_label, show_mp_prefix=False, mp_name_map=None):
    """Build top 10 gainers and decliners for a marketplace group"""
    # Get sellers with delta
    seller_deltas = []
    for (mp, mcid), s in sellers.items():
        if mp not in mp_ids_list:
            continue
        ch = s['channel']
        if channels and ch not in channels:
            continue
        delta = s['wtd_gms'] - s['prev_gms']
        if delta == 0 and s['wtd_gms'] == 0 and s['prev_gms'] == 0:
            continue
        name = s['name']
        if show_mp_prefix and mp_name_map:
            prefix = mp_name_map.get(mp, '')
            if prefix:
                name = f"{prefix} | {name}"
        seller_deltas.append({
            'name': name,
            'mcid': mcid,
            'wtd_gms': s['wtd_gms'],
            'prev_gms': s['prev_gms'],
            'delta': delta,
        })

    # Sort by delta
    gainers = sorted([s for s in seller_deltas if s['delta'] > 0], key=lambda x: -x['delta'])[:10]
    decliners = sorted([s for s in seller_deltas if s['delta'] < 0], key=lambda x: x['delta'])[:10]

    return gainers, decliners


def render_ms_card(title_type, items, week_label):
    """Render a Movers & Shakers card"""
    if title_type == 'gainer':
        header_html = f'<div class="ms-card-header gainer"><span>&#128314; Top 10 Gainers</span><span>WoW Delta</span></div>'
    else:
        header_html = f'<div class="ms-card-header decliner"><span>&#128315; Top 10 Decliners</span><span>WoW Delta</span></div>'

    rows_html = ""
    for i, item in enumerate(items, 1):
        name_esc = html_escape(item['name'])
        mcid = item['mcid']
        wtd = fmt_money(item['wtd_gms'])
        prev = fmt_money(item['prev_gms'])
        delta = item['delta']
        if delta > 0:
            delta_html = f'<span class="up">${abs(round(delta)):,}</span>'
        elif delta < 0:
            delta_html = f'<span class="down">-${abs(round(delta)):,}</span>'
        else:
            delta_html = '<span class="flat">$0</span>'

        wow = calc_wow(item['wtd_gms'], item['prev_gms'])
        if wow is None:
            wow_html = '<span class="na">N/A</span>'
        elif wow > 0:
            wow_html = f'<span class="up">&#9650; {wow:.1f}%</span>'
        elif wow < 0:
            wow_html = f'<span class="down">&#9660; {abs(wow):.1f}%</span>'
        else:
            wow_html = '<span class="flat">0.0%</span>'

        rows_html += (
            f'<tr><td>{i}</td><td title="{name_esc}">{name_esc}</td><td>{mcid}</td>\n'
            f'<td>{wtd}</td><td>{prev}</td>\n'
            f'<td>{delta_html}</td><td>{wow_html}</td></tr>\n'
        )

    return (
        f'<div class="ms-card">\n{header_html}\n'
        f'<table><thead><tr><th>#</th><th>Seller</th><th>MCID</th>'
        f'<th>W{CURRENT_WEEK}</th><th>W{PREV_WEEK}</th><th>Delta</th><th>WoW %</th></tr></thead>'
        f'<tbody>\n{rows_html}</tbody></table></div>\n'
    )


# ─── Market Tabs (Seller Tables) ─────────────────────────────────────────

def render_market_tab(tab_id, title, mp_id_val, sellers_list, is_esm=False, region_name="MEA"):
    """Render a full market tab panel with table"""
    # Sort by YTD GMS descending
    sellers_sorted = sorted(sellers_list, key=lambda s: -s['ytd_gms'])
    count = len(sellers_sorted)

    # Compute total for share calculations
    total_wtd_gms = sum(s['wtd_gms'] for s in sellers_sorted)
    total_ytd_gms = sum(s['ytd_gms'] for s in sellers_sorted)

    # Discover owners
    owners = sorted(set(s['owner'] for s in sellers_sorted if s['owner']))

    # Panel header
    html = f'<div class="tab-panel{"" if tab_id.endswith("_first") else ""}" id="panel-{tab_id}">\n'
    html += f'<div class="panel-header"><span class="title">{title}</span><span class="badge">{count} sellers</span></div>\n'

    # Toolbar with search
    html += f'<div class="toolbar">\n'
    html += f'  <input type="text" placeholder="&#128269; Search seller name or MCID..." oninput="applyFilters(\'{tab_id}\')" id="search-{tab_id}">\n'
    html += f'  <span class="info" id="count-{tab_id}">Showing {count} of {count}</span>\n'
    html += f'</div>\n'

    # Button bar with filters
    html += f'<div class="btn-bar">\n'
    if not is_esm:
        # Channel filter
        html += f'  <span class="filter-label">Channel:</span>\n'
        html += f'  <button class="filter-btn active" data-filter="ch" data-val="all" data-tab="{tab_id}" onclick="setChFilter(this)">All</button>\n'
        html += f'  <button class="filter-btn" data-filter="ch" data-val="DSR" data-tab="{tab_id}" onclick="setChFilter(this)">DSR</button>\n'
        html += f'  <button class="filter-btn" data-filter="ch" data-val="SSR" data-tab="{tab_id}" onclick="setChFilter(this)">SSR</button>\n'
        html += f'  <span class="filter-sep">|</span>\n'

    # Owner filter
    html += f'  <span class="filter-label">Owner:</span>\n'
    html += f'  <div class="owner-dropdown" id="od-{tab_id}">\n'
    html += f'    <button class="filter-btn" onclick="toggleOwnerDropdown(\'{tab_id}\')">All Owners &#9662;</button>\n'
    html += f'    <div class="owner-menu" id="om-{tab_id}">\n'
    html += f'      <label class="owner-opt"><input type="checkbox" value="__all__" checked onchange="ownerSelectAll(\'{tab_id}\',this)"> <b>Select All</b></label>\n'
    for owner in owners:
        html += f'      <label class="owner-opt"><input type="checkbox" value="{html_escape(owner)}" checked onchange="ownerChanged(\'{tab_id}\')"> {html_escape(owner)}</label>\n'
    html += f'    </div>\n'
    html += f'  </div>\n'
    html += f'  <span class="filter-sep">|</span>\n'
    html += f'  <button class="copy-btn" onclick="copyMcids(\'{tab_id}\',\'newline\',this)">&#128203; Copy MCIDs (&#25563;&#34892;)</button>\n'
    html += f'  <button class="copy-btn" onclick="copyMcids(\'{tab_id}\',\'comma\',this)">&#128203; Copy MCIDs (&#36887;&#34399;)</button>\n'
    html += f'</div>\n'

    # Table header
    html += f'<div class="table-wrap"><table><thead><tr>\n'
    html += f'  <th class="al">#</th><th class="al">Merchant Customer ID</th>\n'
    if not is_esm:
        html += f'  <th class="al">Channel</th>\n'
    html += f'  <th class="al">Opportunity Owner</th>\n'
    html += f'  <th class="al">Seller</th>\n'
    html += f'  <th>W{CURRENT_WEEK} {CURRENT_YEAR}</th><th>W{PREV_WEEK} {CURRENT_YEAR}</th><th>WoW %</th>\n'
    html += f'  <th>W{CURRENT_WEEK} {LAST_YEAR}</th><th>YoY %</th><th>Units</th>\n'
    html += f'  <th>YTD GMS</th><th>YTD YoY</th><th>YTD Share</th>\n'
    html += f'  <th>Category</th><th>Wk Share</th>\n'
    html += f'</tr></thead><tbody>\n'

    # Table rows
    for i, s in enumerate(sellers_sorted, 1):
        name_esc = html_escape(s['name'])
        name_lower = s['name'].lower()
        mcid = s['mcid']
        owner = s['owner']
        channel = s['channel']
        search_str = f"{name_lower} {mcid} {owner.lower()}"

        ch_class = f"ch-{channel.lower()}" if channel in ['DSR', 'SSR'] else ""

        # Calculate metrics
        wow = calc_wow(s['wtd_gms'], s['prev_gms'])
        yoy = calc_yoy(s['wtd_gms'], s['ly_gms'])
        ytd_yoy = calc_yoy(s['ytd_gms'], s['ly_ytd_gms'])

        ytd_share = (s['ytd_gms'] / total_ytd_gms * 100) if total_ytd_gms != 0 else 0
        wk_share = (s['wtd_gms'] / total_wtd_gms * 100) if total_wtd_gms != 0 else 0

        # Format WoW
        if wow is None:
            wow_html = '<span class="na">N/A</span>'
        elif wow > 0:
            wow_html = f'<span class="up">&#9650; {wow:.1f}%</span>'
        elif wow < 0:
            wow_html = f'<span class="down">&#9660; {abs(wow):.1f}%</span>'
        else:
            wow_html = '<span class="na">N/A</span>'

        # Format YoY
        if yoy is None:
            yoy_html = '<span class="na">N/A</span>'
        elif yoy > 0:
            yoy_html = f'<span class="up">&#9650; {yoy:.1f}%</span>'
        elif yoy < 0:
            yoy_html = f'<span class="down">&#9660; {abs(yoy):.1f}%</span>'
        else:
            yoy_html = '<span class="flat">0.0%</span>'

        # Format YTD YoY
        if ytd_yoy is None:
            ytd_yoy_html = '<span class="na">N/A</span>'
        elif ytd_yoy > 0:
            ytd_yoy_html = f'<span class="up">&#9650; {ytd_yoy:.1f}%</span>'
        elif ytd_yoy < 0:
            ytd_yoy_html = f'<span class="down">&#9660; {abs(ytd_yoy):.1f}%</span>'
        else:
            ytd_yoy_html = '<span class="flat">0.0%</span>'

        cat_html = f'<span class="cat">{html_escape(s["category"])}</span>' if s['category'] else ''

        data_ch = channel if not is_esm else ""
        html += f'<tr data-s="{html_escape(search_str)}" data-ch="{data_ch}" data-ow="{html_escape(owner)}">\n'
        html += f'<td class="rn">{i}</td><td class="mcid">{mcid}</td>\n'
        if not is_esm:
            html += f'<td class="channel {ch_class}">{channel}</td>\n'
        html += f'<td class="owner">{html_escape(owner)}</td>\n'
        html += f'<td class="seller" title="{name_esc}">{name_esc}</td>\n'
        html += f'<td data-v="{s["wtd_gms"]}">{fmt_money(s["wtd_gms"])}</td>'
        html += f'<td data-v="{s["prev_gms"]}">{fmt_money(s["prev_gms"])}</td>'
        html += f'<td>{wow_html}</td>\n'
        html += f'<td data-v="{s["ly_gms"]}">{fmt_money(s["ly_gms"])}</td>'
        html += f'<td>{yoy_html}</td>'
        html += f'<td data-v="{int(s["wtd_units"])}">{int(s["wtd_units"])}</td>\n'
        html += f'<td data-v="{s["ytd_gms"]}" data-ly="{s["ly_ytd_gms"]}">{fmt_money(s["ytd_gms"])}</td>'
        html += f'<td>{ytd_yoy_html}</td>'
        html += f'<td data-v="{ytd_share}">{ytd_share:.1f}%</td>\n'
        html += f'<td>{cat_html}</td>'
        html += f'<td data-v="{wk_share}">{wk_share:.1f}%</td>\n'
        html += f'</tr>\n'

    # Summary row
    s_cur = sum(s['wtd_gms'] for s in sellers_sorted)
    s_prev = sum(s['prev_gms'] for s in sellers_sorted)
    s_ly = sum(s['ly_gms'] for s in sellers_sorted)
    s_units = sum(int(s['wtd_units']) for s in sellers_sorted)
    s_ytd = sum(s['ytd_gms'] for s in sellers_sorted)
    s_ytd_ly = sum(s['ly_ytd_gms'] for s in sellers_sorted)

    s_wow = calc_wow(s_cur, s_prev)
    s_yoy = calc_yoy(s_cur, s_ly)
    s_ytd_yoy = calc_yoy(s_ytd, s_ytd_ly)

    if s_wow is None:
        s_wow_html = '<span class="na">-</span>'
    elif s_wow > 0:
        s_wow_html = f'<span class="up">&#9650; {s_wow:.1f}%</span>'
    elif s_wow < 0:
        s_wow_html = f'<span class="down">&#9660; {abs(s_wow):.1f}%</span>'
    else:
        s_wow_html = '<span class="flat">0.0%</span>'

    if s_yoy is None:
        s_yoy_html = '<span class="na">N/A</span>'
    elif s_yoy > 0:
        s_yoy_html = f'<span class="up">&#9650; {s_yoy:.1f}%</span>'
    elif s_yoy < 0:
        s_yoy_html = f'<span class="down">&#9660; {abs(s_yoy):.1f}%</span>'
    else:
        s_yoy_html = '<span class="flat">0.0%</span>'

    if s_ytd_yoy is None:
        s_ytd_yoy_html = '<span class="na">N/A</span>'
    elif s_ytd_yoy > 0:
        s_ytd_yoy_html = f'<span class="up">&#9650; {s_ytd_yoy:.1f}%</span>'
    elif s_ytd_yoy < 0:
        s_ytd_yoy_html = f'<span class="down">&#9660; {abs(s_ytd_yoy):.1f}%</span>'
    else:
        s_ytd_yoy_html = '<span class="flat">0.0%</span>'

    html += f'<tr class="summary-row"><td class="rn"></td><td class="mcid"></td>'
    if not is_esm:
        html += '<td></td>'
    html += f'<td></td><td class="seller" data-col="label">Total ({count} sellers)</td>\n'
    html += f'<td data-col="cur">{fmt_money(s_cur)}</td>'
    html += f'<td data-col="prev">{fmt_money(s_prev)}</td>'
    html += f'<td data-col="wow">{s_wow_html}</td>\n'
    html += f'<td data-col="yoy_g">{fmt_money(s_ly)}</td>'
    html += f'<td data-col="yoy">{s_yoy_html}</td>'
    html += f'<td data-col="units">{s_units}</td>\n'
    html += f'<td data-col="ytd">{fmt_money(s_ytd)}</td>'
    html += f'<td data-col="ytd_yoy">{s_ytd_yoy_html}</td>'
    html += f'<td data-col="ytd_sh">100.0%</td>\n'
    html += f'<td></td><td data-col="wk_sh">100.0%</td></tr>\n'

    html += f'</tbody></table></div></div>\n'
    return html


# ─── Executive Summary ────────────────────────────────────────────────────

def render_executive_summary(exp_rows):
    """Generate the executive summary section"""
    # Extract key metrics
    eu = exp_rows[0][1]
    jp = exp_rows[3][1]
    au = exp_rows[6][1]
    au_esm = exp_rows[8][1]
    mena = exp_rows[9][1]
    ae = exp_rows[10][1]
    sa = exp_rows[12][1]
    mena_esm = exp_rows[14][1]
    sum16 = exp_rows[15][1]
    sum17 = exp_rows[16][1]

    def color_span(val, fmt_str=None):
        if val is None:
            return ''
        if val > 0:
            return f'<span style="color:#006100; font-weight:600;">+{val:.1f}%</span>'
        elif val < 0:
            return f'<span style="color:#9c0006; font-weight:600;">{val:.1f}%</span>'
        else:
            return f'<span style="color:#888; font-weight:600;">0.0%</span>'

    def wow_str(val):
        if val is None:
            return ''
        return f'WoW {color_span(val)}'

    def yoy_str(val):
        if val is None:
            return ''
        return f'YoY {color_span(val)}'

    # Get AU ESM ytd_launch from data (count of sellers with ytd_launch > 0)
    # For ESM, we track active sellers not launches
    au_esm_ytd_note = ""
    mena_esm_ytd_note = ""

    html = f'''
<div style="margin-top:28px; background:#fff; border:1px solid #ccc; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,.08); padding:24px 32px;">
<h2 style="color:#1a3a5c; font-size:18px; margin-bottom:16px; border-bottom:2px solid #2d5f8a; padding-bottom:8px;">&#128202; W{CURRENT_WEEK} {CURRENT_YEAR} Executive Summary</h2>

<h3 style="color:#2d5f8a; font-size:14px; margin:14px 0 8px;">Overall DSR Expansion (Excl. AU/MENA ESM)</h3>
<ul style="font-size:13px; line-height:1.8; color:#333; padding-left:20px;">
  <li>W{CURRENT_WEEK} GMS reached <b>{fmt_money(sum16["wtd_gms"])}</b>, {color_span(sum16["wow_gms"])} WoW and {color_span(sum16["yoy_gms"])} YoY.</li>
  <li>YTD GMS stands at <b>{fmt_money(sum16["ytd_gms"])}</b>, {color_span(sum16["ytd_yoy_gms"])} YoY.</li>
  <li>W{CURRENT_WEEK} Seller Launch: <b>{sum16["wtd_launch"]}</b> new sellers ({color_span(sum16["wow_launch"])} WoW); YTD total <b>{sum16["ytd_launch"]}</b> sellers ({color_span(sum16["ytd_yoy_launch"])} YoY).</li>
</ul>

<h3 style="color:#2d5f8a; font-size:14px; margin:14px 0 8px;">Including AU &amp; MENA ESM</h3>
<ul style="font-size:13px; line-height:1.8; color:#333; padding-left:20px;">
  <li>Total W{CURRENT_WEEK} GMS <b>{fmt_money(sum17["wtd_gms"])}</b>, {color_span(sum17["wow_gms"])} WoW and {color_span(sum17["yoy_gms"])} YoY.</li>
  <li>YTD GMS <b>{fmt_money(sum17["ytd_gms"])}</b>, {color_span(sum17["ytd_yoy_gms"])} YoY.</li>
</ul>

<h3 style="color:#2d5f8a; font-size:14px; margin:14px 0 8px;">Regional Highlights</h3>
<ul style="font-size:13px; line-height:1.8; color:#333; padding-left:20px;">
  <li><b>EU</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(eu["wtd_gms"])}</b> ({color_span(eu["wow_gms"])} WoW, {color_span(eu["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(eu["ytd_gms"])}</b> ({color_span(eu["ytd_yoy_gms"])} YoY). YTD Seller Launch {eu["ytd_launch"]} ({color_span(eu["ytd_yoy_launch"])} YoY).</li>
  <li><b>JP</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(jp["wtd_gms"])}</b> ({color_span(jp["wow_gms"])} WoW, {color_span(jp["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(jp["ytd_gms"])}</b> ({color_span(jp["ytd_yoy_gms"])} YoY). YTD Seller Launch {jp["ytd_launch"]} ({color_span(jp["ytd_yoy_launch"])} YoY).</li>
  <li><b>AU</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(au["wtd_gms"])}</b> ({color_span(au["wow_gms"])} WoW, {color_span(au["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(au["ytd_gms"])}</b> ({color_span(au["ytd_yoy_gms"])} YoY). YTD Seller Launch {au["ytd_launch"]} ({color_span(au["ytd_yoy_launch"])} YoY).</li>
  <li><b>AU ESM</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(au_esm["wtd_gms"])}</b> ({color_span(au_esm["wow_gms"])} WoW, {color_span(au_esm["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(au_esm["ytd_gms"])}</b> ({color_span(au_esm["ytd_yoy_gms"])} YoY).</li>
  <li><b>MENA DSR (AE+SA)</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(mena["wtd_gms"])}</b> ({color_span(mena["wow_gms"])} WoW, {color_span(mena["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(mena["ytd_gms"])}</b> ({color_span(mena["ytd_yoy_gms"])} YoY). YTD Seller Launch {mena["ytd_launch"]} ({color_span(mena["ytd_yoy_launch"])} YoY).</li>
  <li><b>AE</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(ae["wtd_gms"])}</b> ({color_span(ae["wow_gms"])} WoW, {color_span(ae["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(ae["ytd_gms"])}</b> ({color_span(ae["ytd_yoy_gms"])} YoY). YTD Seller Launch {ae["ytd_launch"]} ({color_span(ae["ytd_yoy_launch"])} YoY).</li>
  <li><b>SA</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(sa["wtd_gms"])}</b> ({color_span(sa["wow_gms"])} WoW, {color_span(sa["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(sa["ytd_gms"])}</b> ({color_span(sa["ytd_yoy_gms"])} YoY). YTD Seller Launch {sa["ytd_launch"]} ({color_span(sa["ytd_yoy_launch"])} YoY).</li>
  <li><b>MENA ESM</b> &mdash; W{CURRENT_WEEK} GMS <b>{fmt_money(mena_esm["wtd_gms"])}</b> ({color_span(mena_esm["wow_gms"])} WoW, {color_span(mena_esm["yoy_gms"])} YoY). YTD GMS <b>{fmt_money(mena_esm["ytd_gms"])}</b> ({color_span(mena_esm["ytd_yoy_gms"])} YoY).</li>
</ul>

<h3 style="color:#2d5f8a; font-size:14px; margin:14px 0 8px;">Key Risks &amp; Watch Items</h3>
<ul style="font-size:13px; line-height:1.8; color:#333; padding-left:20px;">
'''

    # Auto-generate risk items based on negative metrics
    risks = []
    for label, data in [('AE', ae), ('SA', sa), ('EU', eu), ('JP', jp), ('AU', au)]:
        if data['wow_gms'] is not None and data['wow_gms'] < -15:
            risks.append(f'  <li>{label} GMS declining {color_span(data["wow_gms"])} WoW &mdash; needs attention.</li>')
        if data['ytd_yoy_gms'] is not None and data['ytd_yoy_gms'] < -20:
            risks.append(f'  <li>{label} YTD GMS down {color_span(data["ytd_yoy_gms"])} &mdash; largest gap to close.</li>')

    if not risks:
        risks.append('  <li>No major risks this week.</li>')

    html += '\n'.join(risks)
    html += '\n</ul>\n</div>\n'
    return html


# ─── Full HTML Generation ─────────────────────────────────────────────────

def generate_dashboard_html(sellers, tw_rows):
    """Generate the complete dashboard HTML"""
    print("Building expansion summary...")
    exp_rows, mp_ids_map = build_expansion_summary(sellers, tw_rows)

    eu5_ids = mp_ids_map['eu5_ids']
    jp_ids = mp_ids_map['jp_ids']
    au_ids = mp_ids_map['au_ids']
    ae_ids = mp_ids_map['ae_ids']
    sa_ids = mp_ids_map['sa_ids']
    mena_ids = mp_ids_map['mena_ids']

    # EU marketplace name map (for Movers & Shakers prefixes)
    eu_mp_names = {3: '&#127468;&#127463; UK', 4: '&#127465;&#127466; DE', 5: '&#127467;&#127479; FR',
                   35691: '&#127470;&#127481; IT', 44551: '&#127466;&#127480; ES',
                   44571: '&#127475;&#127473; NL', 328451: '&#127480;&#127466; SE', 338851: '&#127463;&#127466; BE'}

    # Additional EU marketplace IDs
    eu_all_ids = eu5_ids + [44571, 328451, 338851]

    # ─── CSS ───
    css = '''*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f0f2f5;color:#333;padding:20px}
.container{max-width:1500px;margin:0 auto}
h1{text-align:center;color:#1a3a5c;margin-bottom:6px;font-size:26px}
.subtitle{text-align:center;color:#666;margin-bottom:20px;font-size:13px}
.region-bar{display:flex;gap:6px;margin-bottom:16px;justify-content:center}
.region-btn{padding:8px 22px;font-size:14px;font-weight:700;cursor:pointer;
  border:2px solid #1a3a5c;border-radius:20px;background:#fff;color:#1a3a5c;transition:all .15s}
.region-btn:hover{background:#e8eef4}
.region-btn.active{background:#1a3a5c;color:#fff}
.region-panel{display:none}
.region-panel.active{display:block}
.tab-bar{display:flex;flex-wrap:wrap;gap:4px;border-bottom:3px solid #2d5f8a}
.tab-btn{padding:10px 18px;font-size:13px;font-weight:600;cursor:pointer;
  border:1px solid #ccc;border-bottom:none;border-radius:6px 6px 0 0;
  background:#e4e8ec;color:#555;transition:all .15s;white-space:nowrap}
.tab-btn:hover{background:#d0d8e0}
.tab-btn.active{background:#2d5f8a;color:#fff;border-color:#2d5f8a}
.tab-panel{display:none;background:#fff;border:1px solid #ddd;border-top:none;
  border-radius:0 0 8px 8px;box-shadow:0 2px 8px rgba(0,0,0,.08);overflow:hidden}
.tab-panel.active{display:block}
.panel-header{background:linear-gradient(135deg,#1a3a5c,#2d5f8a);color:#fff;padding:14px 24px;
  display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
.panel-header .title{font-size:18px;font-weight:600}
.badge{background:rgba(255,255,255,.22);padding:3px 14px;border-radius:12px;font-size:12px}
.toolbar{padding:10px 20px;background:#f8f9fa;border-bottom:1px solid #eee;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.toolbar input{padding:6px 12px;border:1px solid #ccc;border-radius:4px;font-size:13px;width:280px}
.toolbar .info{font-size:12px;color:#888}
.copy-btn{padding:5px 12px;border:1px solid #2d5f8a;border-radius:4px;background:#fff;color:#2d5f8a;
  font-size:12px;cursor:pointer;transition:all .15s;white-space:nowrap}
.copy-btn:hover{background:#2d5f8a;color:#fff}
.copy-btn.copied{background:#27ae60;border-color:#27ae60;color:#fff}
.btn-bar{padding:6px 20px 10px;background:#f8f9fa;border-bottom:1px solid #eee;display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.filter-label{font-size:12px;font-weight:600;color:#555}
.filter-btn{padding:4px 12px;border:1px solid #ccc;border-radius:4px;background:#fff;color:#555;
  font-size:12px;cursor:pointer;transition:all .12s}
.filter-btn:hover{background:#e8eef4}
.filter-btn.active{background:#2d5f8a;color:#fff;border-color:#2d5f8a}
.filter-sep{color:#ccc;font-size:14px;margin:0 2px}
.owner-dropdown{position:relative;display:inline-block}
.owner-menu{display:none;position:absolute;top:100%;left:0;background:#fff;border:1px solid #ccc;border-radius:6px;
  box-shadow:0 4px 16px rgba(0,0,0,.15);padding:8px 4px;z-index:99;max-height:300px;overflow-y:auto;min-width:200px}
.owner-menu.open{display:block}
.owner-opt{display:block;padding:3px 10px;font-size:12px;cursor:pointer;white-space:nowrap}
.owner-opt:hover{background:#f0f4f8}
.owner-opt input{margin-right:6px}
td.channel{text-align:center;font-size:11px;font-weight:600}
.ch-dsr{color:#2d5f8a}
.ch-ssr{color:#8e44ad}
.table-wrap{overflow-x:auto;max-height:78vh;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:#34495e;color:#fff;padding:9px 12px;text-align:right;font-weight:600;
  white-space:nowrap;position:sticky;top:0;z-index:2}
thead th.al{text-align:left}
tbody td{padding:7px 12px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap}
td.rn{text-align:center;color:#888;font-weight:600;width:36px}
td.mcid{text-align:left;font-family:monospace;font-size:11px;color:#666}
td.seller{text-align:left;font-weight:500;max-width:240px;overflow:hidden;text-overflow:ellipsis}
td.owner{text-align:left;font-size:12px;color:#555}
tbody tr:hover{background:#f0f5fa}
tbody tr:nth-child(even){background:#fafbfc}
tbody tr:nth-child(even):hover{background:#eaf0f6}
.up{color:#27ae60;font-weight:600}.down{color:#e74c3c;font-weight:600}.flat{color:#888}.na{color:#bbb}
.summary-row{background:#e4ecf4!important;font-weight:700;position:sticky;bottom:0;z-index:1}
.summary-row td{border-top:2px solid #34495e}
.cat{display:inline-block;background:#f0f0f0;padding:2px 8px;border-radius:4px;font-size:11px;color:#555}
.ms-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;padding:16px 20px}
@media(max-width:900px){.ms-grid{grid-template-columns:1fr}}
.ms-card{background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden}
.ms-card-header{padding:10px 16px;font-size:14px;font-weight:700;display:flex;justify-content:space-between;align-items:center}
.ms-card-header.gainer{background:#e8f5e9;color:#2e7d32}
.ms-card-header.decliner{background:#fce4ec;color:#c62828}
.ms-card table{width:100%;border-collapse:collapse;font-size:12px}
.ms-card th{background:#e2e6ea;padding:6px 10px;text-align:right;font-weight:700;font-size:11px;color:#1a3a5c}
.ms-card th:nth-child(1),.ms-card th:nth-child(2),.ms-card th:nth-child(3){text-align:left}
.ms-card td{padding:5px 10px;border-bottom:1px solid #f0f0f0;text-align:right}
.ms-card td:nth-child(1){text-align:center;color:#888;width:28px}
.ms-card td:nth-child(2){text-align:left;font-weight:500;max-width:200px;overflow:hidden;text-overflow:ellipsis}
.ms-card td:nth-child(3){text-align:left;font-family:monospace;font-size:10px;color:#888}
.ms-card tr:hover{background:#fafafa}
.ms-section-title{padding:14px 20px 6px;font-size:16px;font-weight:700;color:#1a3a5c;border-bottom:2px solid #2d5f8a;margin:0 0 0 0}
.exp-embed{padding:16px 0}
.exp-embed .container{max-width:100%}
.exp-embed h1{font-size:18px;margin-bottom:4px}
.exp-embed .subtitle{margin-bottom:12px;font-size:12px}
.exp-embed .table-wrap{border-radius:6px;overflow-x:auto}
.exp-embed table{font-size:12px}
@media(max-width:768px){body{padding:8px}table{font-size:11px}thead th,tbody td{padding:5px 6px}.tab-btn{padding:7px 10px;font-size:11px}.region-btn{padding:6px 14px;font-size:12px}}
.export-btn{padding:6px 16px;border:1px solid #2d5f8a;border-radius:4px;background:#fff;color:#2d5f8a;
  font-size:12px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap;display:inline-flex;align-items:center;gap:4px}
.export-btn:hover{background:#2d5f8a;color:#fff}
.export-btn.exported{background:#27ae60;border-color:#27ae60;color:#fff}
.export-bar{padding:8px 20px;background:#f8f9fa;border-bottom:1px solid #eee;display:flex;justify-content:flex-end;gap:8px;align-items:center}
.exp-export-bar{margin-top:12px;display:flex;justify-content:flex-end;gap:8px}'''

    css2 = '''
thead tr.header-year td { background: #d9e1f2; font-weight: 700; padding: 4px 8px; border: 1px solid #b4c6e7; text-align: center; }
thead tr.header-section td { background: #e2efda; font-weight: 700; padding: 4px 8px; border: 1px solid #a9d18e; text-align: center; font-size: 11px; }
thead tr.header-section td.gms-header { background: #e2efda; }
thead tr.header-section td.seller-header { background: #fce4d6; }
thead tr.header-section td.soa-header { background: #d6dce4; }
thead tr.header-col td { background: #f2f2f2; font-weight: 700; padding: 5px 8px; border: 1px solid #d0d0d0; text-align: center; font-size: 11px; }
.exp-embed tbody td { padding: 4px 8px; border: 1px solid #e0e0e0; text-align: right; font-size: 12px; }
.exp-embed tbody td.left { text-align: left; }
.exp-embed tbody td.center { text-align: center; }
.exp-embed tbody tr.summary-row { background: #c5d9a4 !important; font-weight: 700; }
.exp-embed tbody tr.summary-row2 { background: #a8c97a !important; font-weight: 700; }
.exp-embed tbody tr.mp-row { background: #eef3f8; }
.exp-embed tbody tr.ic-row { background: #fff; }
td.sep-gms { border-left: 3px solid #548235; }
td.sep-seller { border-left: 3px solid #c55a11; }
td.sep-soa { border-left: 3px solid #4472c4; }
'''

    # ─── Build HTML ───
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WBR W{CURRENT_WEEK} {CURRENT_YEAR} - Expansion Dashboard</title>
<style>
{css}
</style>
<style>
{css2}
</style>

</head>
<body>
<div class="container">
<h1>WBR W{CURRENT_WEEK} {CURRENT_YEAR} &mdash; Expansion Dashboard</h1>
<p class="subtitle">MEA &nbsp;|&nbsp; EU &nbsp;|&nbsp; JP &nbsp;|&nbsp; NSR (DSR+SSR) &middot; ESM &nbsp;|&nbsp; All Sellers</p>

<!-- Region tabs -->
<div class="region-bar">
  <div class="region-btn active" data-region="EXP">&#128200; Expansion DSR</div>
  <div class="region-btn" data-region="MS">&#128202; Movers &amp; Shakers</div>
  <div class="region-btn" data-region="MEA">MEA - Middle East & Australia</div>
  <div class="region-btn" data-region="EU">EU - Europe</div>
  <div class="region-btn" data-region="JP">JP - Japan</div>
</div>
'''

    # ─── Region: EXP ───
    print("Rendering EXP region...")
    html += '<div class="region-panel active" id="region-EXP">\n'
    html += '<div class="exp-embed">\n<div class="container">\n'
    html += f'<h1>W{CURRENT_WEEK} {CURRENT_YEAR} &mdash; TW2 Expansion DSR GS WBR</h1>\n'
    html += f'<p class="subtitle">Calculated from WBR Page 0 MCID raw data | Filters: EU/JP=DSR only, AU/MENA=DSR+SSR, seller_origin=TW, is_seller_fraud=0</p>\n'
    html += '<div class="table-wrap">\n<table>\n<thead>\n'
    html += f'<tr class="header-section">\n'
    html += f'<td>{CURRENT_WEEK}</td><td>TW2 Expansion DSR GS WBR</td>\n'
    html += f'<td class="gms-header sep-gms" colspan="6">{CURRENT_YEAR} GMS</td>\n'
    html += f'<td class="seller-header sep-seller" colspan="5">{CURRENT_YEAR} Seller Launch</td>\n'
    html += f'<td class="soa-header sep-soa" colspan="5">{CURRENT_YEAR} SOA Launch</td>\n'
    html += f'</tr>\n'
    html += f'<tr class="header-col">\n'
    html += f'<td>#</td><td>MP</td>\n'
    html += f'<td class="sep-gms">W{PREV_WEEK}</td><td>W{CURRENT_WEEK}</td><td>WoW</td><td>YoY</td><td>YTD</td><td>YTD-YoY</td>\n'
    html += f'<td class="sep-seller">W{PREV_WEEK}</td><td>W{CURRENT_WEEK}</td><td>WoW</td><td>W{CURRENT_WEEK}YTD</td><td>YTD-YoY</td>\n'
    html += f'<td class="sep-soa">W{PREV_WEEK}</td><td>W{CURRENT_WEEK}</td><td>WoW</td><td>YTD</td><td>YTD-YoY</td>\n'
    html += f'</tr>\n</thead>\n<tbody>\n'

    for idx, (row_type, row_data) in enumerate(exp_rows, 1):
        html += render_exp_row_html(idx, row_type, row_data) + '\n'

    html += '</tbody>\n</table>\n</div>\n'

    # Executive summary
    html += render_executive_summary(exp_rows)

    html += '</div>\n</div>\n</div>\n'  # close exp-embed, container, region-EXP

    # ─── Region: MS (Movers & Shakers) ───
    print("Rendering MS region...")
    html += '<div class="region-panel" id="region-MS">\n'

    # EU5 NSR
    gainers, decliners = build_movers_shakers(sellers, eu5_ids, ['DSR', 'SSR'], 'EU5 NSR',
                                              show_mp_prefix=True, mp_name_map=eu_mp_names)
    html += f'<div class="ms-section-title">&#127466;&#127482; EU5 &mdash; NSR (DSR+SSR)</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # EU5 ESM
    gainers, decliners = build_movers_shakers(sellers, eu5_ids, ['ESM'], 'EU5 ESM',
                                              show_mp_prefix=True, mp_name_map=eu_mp_names)
    html += f'<div class="ms-section-title">&#127466;&#127482; EU5 &mdash; ESM</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # JP NSR
    gainers, decliners = build_movers_shakers(sellers, jp_ids, ['DSR', 'SSR'], 'JP NSR')
    html += f'<div class="ms-section-title">&#127471;&#127477; JP &mdash; NSR (DSR+SSR)</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # JP ESM
    gainers, decliners = build_movers_shakers(sellers, jp_ids, ['ESM'], 'JP ESM')
    html += f'<div class="ms-section-title">&#127471;&#127477; JP &mdash; ESM</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # AU NSR
    gainers, decliners = build_movers_shakers(sellers, au_ids, ['DSR', 'SSR'], 'AU NSR')
    html += f'<div class="ms-section-title">&#127462;&#127482; AU &mdash; NSR (DSR+SSR)</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # AU ESM
    gainers, decliners = build_movers_shakers(sellers, au_ids, ['ESM'], 'AU ESM')
    html += f'<div class="ms-section-title">&#127462;&#127482; AU &mdash; ESM</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # AE NSR
    gainers, decliners = build_movers_shakers(sellers, ae_ids, ['DSR', 'SSR'], 'AE NSR')
    html += f'<div class="ms-section-title">&#127462;&#127466; AE &mdash; NSR (DSR+SSR)</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # AE ESM
    gainers, decliners = build_movers_shakers(sellers, ae_ids, ['ESM'], 'AE ESM')
    html += f'<div class="ms-section-title">&#127462;&#127466; AE &mdash; ESM</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # SA NSR
    gainers, decliners = build_movers_shakers(sellers, sa_ids, ['DSR', 'SSR'], 'SA NSR')
    html += f'<div class="ms-section-title">&#127480;&#127462; SA &mdash; NSR (DSR+SSR)</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    # SA ESM
    gainers, decliners = build_movers_shakers(sellers, sa_ids, ['ESM'], 'SA ESM')
    html += f'<div class="ms-section-title">&#127480;&#127462; SA &mdash; ESM</div>\n'
    html += '<div class="ms-grid">\n'
    html += render_ms_card('gainer', gainers, f'W{CURRENT_WEEK}')
    html += render_ms_card('decliner', decliners, f'W{CURRENT_WEEK}')
    html += '</div>\n'

    html += '</div>\n'  # close region-MS

    # ─── Region: MEA ───
    print("Rendering MEA region...")
    html += '<div class="region-panel" id="region-MEA">\n'
    html += '<div class="tab-bar">\n'
    html += f'  <div class="tab-btn active" data-tab="AU_NSR" data-region="MEA">&#127462;&#127482; AU - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="AU_ESM" data-region="MEA">&#127462;&#127482; AU - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="AE_NSR" data-region="MEA">&#127462;&#127466; AE - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="AE_ESM" data-region="MEA">&#127462;&#127466; AE - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="SA_NSR" data-region="MEA">&#127480;&#127462; SA - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="SA_ESM" data-region="MEA">&#127480;&#127462; SA - ESM</div>\n'
    html += '</div>\n'

    # AU NSR tab
    au_nsr_sellers = [s for (mp, _), s in sellers.items()
                      if mp in au_ids and s['channel'] in ['DSR', 'SSR'] and s['ytd_launch'] > 0]
    tab_html = render_market_tab('AU_NSR',
        f'&#127462;&#127482; AU &mdash; Australia (ID: 111172) &mdash; NSR (DSR+SSR)',
        111172, au_nsr_sellers, is_esm=False, region_name="MEA")
    # Make first tab active
    tab_html = tab_html.replace('id="panel-AU_NSR"', 'id="panel-AU_NSR" class="tab-panel active"', 1)
    tab_html = tab_html.replace('<div class="tab-panel" id="panel-AU_NSR" class="tab-panel active"',
                                '<div class="tab-panel active" id="panel-AU_NSR"', 1)
    # Fix: the render_market_tab already starts with <div class="tab-panel...
    # Let's just handle the active state:
    html += tab_html.replace('<div class="tab-panel"', '<div class="tab-panel active"', 1)

    # AU ESM tab
    au_esm_sellers = [s for (mp, _), s in sellers.items()
                      if mp in au_ids and s['channel'] == 'ESM']
    html += render_market_tab('AU_ESM',
        f'&#127462;&#127482; AU &mdash; Australia (ID: 111172) &mdash; ESM',
        111172, au_esm_sellers, is_esm=True, region_name="MEA")

    # AE NSR tab
    ae_nsr_sellers = [s for (mp, _), s in sellers.items()
                      if mp in ae_ids and s['channel'] in ['DSR', 'SSR'] and s['ytd_launch'] > 0]
    ae_id_str = ae_ids[0] if ae_ids else '?'
    html += render_market_tab('AE_NSR',
        f'&#127462;&#127466; AE &mdash; UAE (ID: {ae_id_str}) &mdash; NSR (DSR+SSR)',
        ae_id_str, ae_nsr_sellers, is_esm=False, region_name="MEA")

    # AE ESM tab
    ae_esm_sellers = [s for (mp, _), s in sellers.items()
                      if mp in ae_ids and s['channel'] == 'ESM']
    html += render_market_tab('AE_ESM',
        f'&#127462;&#127466; AE &mdash; UAE (ID: {ae_id_str}) &mdash; ESM',
        ae_id_str, ae_esm_sellers, is_esm=True, region_name="MEA")

    # SA NSR tab
    sa_nsr_sellers = [s for (mp, _), s in sellers.items()
                      if mp in sa_ids and s['channel'] in ['DSR', 'SSR'] and s['ytd_launch'] > 0]
    sa_id_str = sa_ids[0] if sa_ids else '?'
    html += render_market_tab('SA_NSR',
        f'&#127480;&#127462; SA &mdash; Saudi Arabia (ID: {sa_id_str}) &mdash; NSR (DSR+SSR)',
        sa_id_str, sa_nsr_sellers, is_esm=False, region_name="MEA")

    # SA ESM tab
    sa_esm_sellers = [s for (mp, _), s in sellers.items()
                      if mp in sa_ids and s['channel'] == 'ESM']
    html += render_market_tab('SA_ESM',
        f'&#127480;&#127462; SA &mdash; Saudi Arabia (ID: {sa_id_str}) &mdash; ESM',
        sa_id_str, sa_esm_sellers, is_esm=True, region_name="MEA")

    html += '</div>\n'  # close region-MEA

    # ─── Region: EU ───
    print("Rendering EU region...")
    html += '<div class="region-panel" id="region-EU">\n'
    html += '<div class="tab-bar">\n'
    html += f'  <div class="tab-btn active" data-tab="EU5_NSR" data-region="EU">&#127466;&#127482; EU5 - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="EU5_ESM" data-region="EU">&#127466;&#127482; EU5 - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="UK_NSR" data-region="EU">&#127468;&#127463; UK - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="UK_ESM" data-region="EU">&#127468;&#127463; UK - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="DE_NSR" data-region="EU">&#127465;&#127466; DE - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="DE_ESM" data-region="EU">&#127465;&#127466; DE - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="FR_NSR" data-region="EU">&#127467;&#127479; FR - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="FR_ESM" data-region="EU">&#127467;&#127479; FR - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="IT_NSR" data-region="EU">&#127470;&#127481; IT - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="IT_ESM" data-region="EU">&#127470;&#127481; IT - ESM</div>\n'
    html += f'  <div class="tab-btn" data-tab="ES_NSR" data-region="EU">&#127466;&#127480; ES - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="ES_ESM" data-region="EU">&#127466;&#127480; ES - ESM</div>\n'
    html += '</div>\n'

    eu_mp_config = [
        ('EU5', 'EU5 Combined (UK+DE+FR+IT+ES)', eu5_ids, 'EU5'),
        ('UK', 'UK', [3], '3'),
        ('DE', 'DE', [4], '4'),
        ('FR', 'FR', [5], '5'),
        ('IT', 'IT', [35691], '35691'),
        ('ES', 'ES', [44551], '44551'),
    ]

    first_eu_tab = True
    for mp_code, mp_name, mp_id_list, mp_id_display in eu_mp_config:
        # NSR
        nsr_sellers = [s for (mp, _), s in sellers.items()
                       if mp in mp_id_list and s['channel'] in ['DSR', 'SSR'] and s['ytd_launch'] > 0]
        # For EU5 combined, add MP prefix to seller names
        if mp_code == 'EU5':
            for s in nsr_sellers:
                mp_prefix = eu_mp_names.get(s['mp'], '')
                if mp_prefix and '|' not in s['name']:
                    s['name'] = f"{mp_prefix} | {s['name']}"

        tab_id = f'{mp_code}_NSR'
        title = f'&#127466;&#127482; {mp_name} (ID: {mp_id_display}) &mdash; NSR (DSR+SSR)' if mp_code == 'EU5' else f'{mp_name} (ID: {mp_id_display}) &mdash; NSR (DSR+SSR)'
        tab_content = render_market_tab(tab_id, title, mp_id_display, nsr_sellers, is_esm=False, region_name="EU")
        if first_eu_tab:
            tab_content = tab_content.replace('<div class="tab-panel"', '<div class="tab-panel active"', 1)
            first_eu_tab = False
        html += tab_content

        # ESM
        esm_sellers = [s for (mp, _), s in sellers.items()
                       if mp in mp_id_list and s['channel'] == 'ESM']
        if mp_code == 'EU5':
            for s in esm_sellers:
                mp_prefix = eu_mp_names.get(s['mp'], '')
                if mp_prefix and '|' not in s['name']:
                    s['name'] = f"{mp_prefix} | {s['name']}"

        tab_id = f'{mp_code}_ESM'
        title = f'&#127466;&#127482; {mp_name} (ID: {mp_id_display}) &mdash; ESM' if mp_code == 'EU5' else f'{mp_name} (ID: {mp_id_display}) &mdash; ESM'
        html += render_market_tab(tab_id, title, mp_id_display, esm_sellers, is_esm=True, region_name="EU")

    html += '</div>\n'  # close region-EU

    # ─── Region: JP ───
    print("Rendering JP region...")
    html += '<div class="region-panel" id="region-JP">\n'
    html += '<div class="tab-bar">\n'
    html += f'  <div class="tab-btn active" data-tab="JP_NSR" data-region="JP">&#127471;&#127477; JP - NSR (DSR+SSR)</div>\n'
    html += f'  <div class="tab-btn" data-tab="JP_ESM" data-region="JP">&#127471;&#127477; JP - ESM</div>\n'
    html += '</div>\n'

    # JP NSR
    jp_nsr_sellers = [s for (mp, _), s in sellers.items()
                      if mp in jp_ids and s['channel'] in ['DSR', 'SSR'] and s['ytd_launch'] > 0]
    tab_content = render_market_tab('JP_NSR',
        f'&#127471;&#127477; JP &mdash; Japan (ID: 6) &mdash; NSR (DSR+SSR)',
        6, jp_nsr_sellers, is_esm=False, region_name="JP")
    tab_content = tab_content.replace('<div class="tab-panel"', '<div class="tab-panel active"', 1)
    html += tab_content

    # JP ESM
    jp_esm_sellers = [s for (mp, _), s in sellers.items()
                      if mp in jp_ids and s['channel'] == 'ESM']
    html += render_market_tab('JP_ESM',
        f'&#127471;&#127477; JP &mdash; Japan (ID: 6) &mdash; ESM',
        6, jp_esm_sellers, is_esm=True, region_name="JP")

    html += '</div>\n'  # close region-JP

    # ─── Footer ───
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    html += f'\n<p style="text-align:center;color:#999;font-size:11px;margin-top:16px">Generated from WBR page 0 MCID data_weekly_w{CURRENT_WEEK}_{CURRENT_YEAR}.xlsx &mdash; {now}</p>\n'
    html += '</div>\n'  # close container

    # ─── JavaScript ───
    html += '''<script>
/* Region tab switching */
document.querySelectorAll('.region-btn').forEach(btn=>{
  btn.addEventListener('click',()=>{
    document.querySelectorAll('.region-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.region-panel').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('region-'+btn.dataset.region).classList.add('active');
  });
});
/* Market/cohort tab switching (scoped within region) */
document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click',()=>{
    const region=btn.dataset.region;
    const panel=document.getElementById('region-'+region);
    panel.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    panel.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('panel-'+btn.dataset.tab).classList.add('active');
  });
});
function filterTable(input,tabId){
  document.getElementById('search-'+tabId).value=input.value;
  applyFilters(tabId);
}
function applyFilters(tabId){
  const panel=document.getElementById('panel-'+tabId);
  const q=(document.getElementById('search-'+tabId)||{}).value||'';
  const ql=q.toLowerCase().trim();
  let chFilter='all';
  const chBtns=panel.querySelectorAll('[data-filter="ch"]');
  chBtns.forEach(b=>{if(b.classList.contains('active'))chFilter=b.dataset.val;});
  const owMenu=document.getElementById('om-'+tabId);
  let owSet=null;
  if(owMenu){
    owSet=new Set();
    owMenu.querySelectorAll('input[type=checkbox]').forEach(cb=>{
      if(cb.value!=='__all__'&&cb.checked)owSet.add(cb.value);
    });
  }
  const rows=panel.querySelectorAll('tbody tr:not(.summary-row)');
  let shown=0;
  rows.forEach(tr=>{
    let vis=true;
    if(ql){const s=tr.getAttribute('data-s')||'';if(!s.includes(ql))vis=false;}
    if(vis&&chFilter!=='all'){const ch=tr.getAttribute('data-ch')||'';if(!ch.includes(chFilter))vis=false;}
    if(vis&&owSet){const ow=tr.getAttribute('data-ow')||'';if(ow&&!owSet.has(ow))vis=false;}
    tr.style.display=vis?'':'none';
    if(vis)shown++;
  });
  document.getElementById('count-'+tabId).textContent='Showing '+shown+' of '+rows.length;
  updateSummaryRow(panel,shown);
}
function updateSummaryRow(panel,shown){
  const sr=panel.querySelector('tr.summary-row');
  if(!sr) return;
  const rows=panel.querySelectorAll('tbody tr:not(.summary-row)');
  let firstRow=null;
  for(let i=0;i<rows.length;i++){if(rows[i].style.display!=='none'){firstRow=rows[i];break;}}
  if(!firstRow) return;
  const refTds=[...firstRow.querySelectorAll('td')];
  const colMap={};
  refTds.forEach((td,i)=>{
    if(td.hasAttribute('data-v')){
      if(td.hasAttribute('data-ly')) colMap['ytd']=i;
      else if(!colMap.hasOwnProperty('cur')) colMap['cur']=i;
      else if(!colMap.hasOwnProperty('prev')) colMap['prev']=i;
      else if(!colMap.hasOwnProperty('yoy_g')) colMap['yoy_g']=i;
      else if(!colMap.hasOwnProperty('units')) colMap['units']=i;
      else if(!colMap.hasOwnProperty('ytd_sh')) colMap['ytd_sh']=i;
      else if(!colMap.hasOwnProperty('wk_sh')) colMap['wk_sh']=i;
    }
  });
  let sCur=0,sPrev=0,sYoyG=0,sUnits=0,sYtd=0,sYtdLy=0;
  rows.forEach(tr=>{
    if(tr.style.display==='none') return;
    const tds=[...tr.querySelectorAll('td')];
    if(colMap.cur!==undefined) sCur+=parseFloat(tds[colMap.cur].getAttribute('data-v'))||0;
    if(colMap.prev!==undefined) sPrev+=parseFloat(tds[colMap.prev].getAttribute('data-v'))||0;
    if(colMap.yoy_g!==undefined) sYoyG+=parseFloat(tds[colMap.yoy_g].getAttribute('data-v'))||0;
    if(colMap.units!==undefined) sUnits+=parseFloat(tds[colMap.units].getAttribute('data-v'))||0;
    if(colMap.ytd!==undefined){
      sYtd+=parseFloat(tds[colMap.ytd].getAttribute('data-v'))||0;
      sYtdLy+=parseFloat(tds[colMap.ytd].getAttribute('data-ly'))||0;
    }
  });
  const wow=(sPrev>0)?((sCur-sPrev)/sPrev*100):null;
  const yoy=(sYoyG>0)?((sCur-sYoyG)/sYoyG*100):null;
  const ytdYoy=(sYtdLy>0)?((sYtd-sYtdLy)/sYtdLy*100):null;
  function fmtD(v){return '$'+Math.round(v).toLocaleString('en-US');}
  function fmtP(v){
    if(v===null) return '<span class="na">-</span>';
    const cls=v>0?'up':v<0?'down':'flat';
    const arrow=v>0?'\\u25B2 ':v<0?'\\u25BC ':'';
    return '<span class="'+cls+'">'+arrow+v.toFixed(1)+'%</span>';
  }
  const labelTd=sr.querySelector('td[data-col="label"]');
  if(labelTd) labelTd.textContent='Total ('+shown+' sellers)';
  const colCur=sr.querySelector('td[data-col="cur"]');
  const colPrev=sr.querySelector('td[data-col="prev"]');
  const colWow=sr.querySelector('td[data-col="wow"]');
  const colYoyG=sr.querySelector('td[data-col="yoy_g"]');
  const colYoy=sr.querySelector('td[data-col="yoy"]');
  const colUnits=sr.querySelector('td[data-col="units"]');
  const colYtd=sr.querySelector('td[data-col="ytd"]');
  const colYtdYoy=sr.querySelector('td[data-col="ytd_yoy"]');
  const colYtdSh=sr.querySelector('td[data-col="ytd_sh"]');
  const colWkSh=sr.querySelector('td[data-col="wk_sh"]');
  if(colCur) colCur.textContent=fmtD(sCur);
  if(colPrev) colPrev.textContent=fmtD(sPrev);
  if(colWow) colWow.innerHTML=fmtP(wow);
  if(colYoyG) colYoyG.textContent=fmtD(sYoyG);
  if(colYoy) colYoy.innerHTML=fmtP(yoy);
  if(colUnits) colUnits.textContent=Math.round(sUnits).toLocaleString('en-US');
  if(colYtd) colYtd.textContent=fmtD(sYtd);
  if(colYtdYoy) colYtdYoy.innerHTML=fmtP(ytdYoy);
  if(colYtdSh) colYtdSh.textContent='100.0%';
  if(colWkSh) colWkSh.textContent='100.0%';
}
function setChFilter(btn){
  const tabId=btn.dataset.tab;
  const panel=document.getElementById('panel-'+tabId);
  panel.querySelectorAll('[data-filter="ch"]').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  applyFilters(tabId);
}
function toggleOwnerDropdown(tabId){
  const menu=document.getElementById('om-'+tabId);
  menu.classList.toggle('open');
  if(menu.classList.contains('open')){
    setTimeout(()=>{
      const handler=e=>{
        if(!menu.contains(e.target)&&!e.target.closest('#od-'+tabId)){
          menu.classList.remove('open');
          document.removeEventListener('click',handler);
        }
      };
      document.addEventListener('click',handler);
    },0);
  }
}
function ownerSelectAll(tabId,masterCb){
  const menu=document.getElementById('om-'+tabId);
  menu.querySelectorAll('input[type=checkbox]').forEach(cb=>{cb.checked=masterCb.checked;});
  updateOwnerLabel(tabId);
  applyFilters(tabId);
}
function ownerChanged(tabId){
  const menu=document.getElementById('om-'+tabId);
  const all=menu.querySelectorAll('input[type=checkbox]:not([value=__all__])');
  const checked=[...all].filter(cb=>cb.checked);
  const masterCb=menu.querySelector('input[value=__all__]');
  masterCb.checked=checked.length===all.length;
  updateOwnerLabel(tabId);
  applyFilters(tabId);
}
function updateOwnerLabel(tabId){
  const menu=document.getElementById('om-'+tabId);
  const dd=document.getElementById('od-'+tabId);
  const btn=dd.querySelector('.filter-btn');
  const all=menu.querySelectorAll('input[type=checkbox]:not([value=__all__])');
  const checked=[...all].filter(cb=>cb.checked);
  if(checked.length===all.length)btn.textContent='All Owners \\u25BE';
  else if(checked.length===0)btn.textContent='No Owner \\u25BE';
  else if(checked.length<=2)btn.textContent=checked.map(cb=>cb.value).join(', ')+' \\u25BE';
  else btn.textContent=checked.length+' owners \\u25BE';
}
function copyMcids(tabId,mode,btn){
  const panel=document.getElementById('panel-'+tabId);
  const rows=panel.querySelectorAll('tbody tr:not(.summary-row)');
  const mcids=[];
  rows.forEach(tr=>{
    if(tr.style.display==='none')return;
    const td=tr.querySelector('td.mcid');
    if(td)mcids.push(td.textContent.trim());
  });
  const sep=mode==='comma'?',':'\\n';
  const text=mcids.join(sep);
  navigator.clipboard.writeText(text).then(()=>{
    const orig=btn.textContent;
    btn.textContent='\\u2705 Copied '+mcids.length+' MCIDs';
    btn.classList.add('copied');
    setTimeout(()=>{btn.textContent=orig;btn.classList.remove('copied');},2000);
  });
}

/* Export to CSV */
function escCsv(val){
  val=String(val).replace(/\\s+/g,' ').trim();
  if(val.includes(',')||val.includes('"')||val.includes('\\n'))
    return '"'+val.replace(/"/g,'""')+'"';
  return val;
}
function tableToCSV(table,visibleOnly){
  const rows=[];
  table.querySelectorAll('thead tr').forEach(tr=>{
    const cells=[];
    tr.querySelectorAll('th,td').forEach(td=>{
      const colspan=parseInt(td.getAttribute('colspan'))||1;
      const text=td.textContent.trim();
      cells.push(escCsv(text));
      for(let i=1;i<colspan;i++) cells.push('');
    });
    rows.push(cells.join(','));
  });
  table.querySelectorAll('tbody tr').forEach(tr=>{
    if(visibleOnly && tr.style.display==='none') return;
    const cells=[];
    tr.querySelectorAll('td').forEach(td=>{
      cells.push(escCsv(td.textContent));
    });
    rows.push(cells.join(','));
  });
  return rows.join('\\n');
}
function downloadCSV(csv,filename){
  const bom='\\uFEFF';
  const blob=new Blob([bom+csv],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function exportPanel(containerId,filename,btn){
  const container=document.getElementById(containerId);
  const table=container.querySelector('table');
  if(!table){alert('No exportable table');return;}
  const csv=tableToCSV(table,true);
  downloadCSV(csv,filename);
  const orig=btn.innerHTML;
  btn.innerHTML='\\u2705 Exported';btn.classList.add('exported');
  setTimeout(()=>{btn.innerHTML=orig;btn.classList.remove('exported');},2000);
}
function exportMoversShakers(btn){
  const panel=document.getElementById('region-MS');
  const tables=panel.querySelectorAll('.ms-card table');
  const sections=panel.querySelectorAll('.ms-section-title');
  let allCsv='';
  let si=0;
  sections.forEach((sec,idx)=>{
    const title=sec.textContent.trim();
    allCsv+=title+'\\n';
    for(let c=0;c<2&&si<tables.length;c++,si++){
      const header=tables[si].closest('.ms-card').querySelector('.ms-card-header').textContent.trim();
      allCsv+=header+'\\n';
      allCsv+=tableToCSV(tables[si],false)+'\\n\\n';
    }
  });
  downloadCSV(allCsv,'WBR_W''' + str(CURRENT_WEEK) + '''_Movers_Shakers.csv');
  const orig=btn.innerHTML;
  btn.innerHTML='\\u2705 Exported';btn.classList.add('exported');
  setTimeout(()=>{btn.innerHTML=orig;btn.classList.remove('exported');},2000);
}

/* Dynamically inject export buttons */
(function injectExportButtons(){
  const expPanel=document.getElementById('region-EXP');
  if(expPanel){
    const tableWrap=expPanel.querySelector('.table-wrap');
    if(tableWrap){
      const bar=document.createElement('div');
      bar.className='exp-export-bar';
      bar.innerHTML='<button class="export-btn" onclick="exportPanel(\\'region-EXP\\',\\'WBR_W''' + str(CURRENT_WEEK) + '''_Expansion_DSR.csv\\',this)">\\u{1F4E5} Export CSV</button>';
      tableWrap.parentNode.insertBefore(bar,tableWrap.nextSibling);
    }
  }
  const msPanel=document.getElementById('region-MS');
  if(msPanel){
    const bar=document.createElement('div');
    bar.className='exp-export-bar';
    bar.style.padding='12px 20px';
    bar.innerHTML='<button class="export-btn" onclick="exportMoversShakers(this)">\\u{1F4E5} Export CSV</button>';
    msPanel.insertBefore(bar,msPanel.firstChild);
  }
  document.querySelectorAll('.tab-panel').forEach(panel=>{
    const panelId=panel.id;
    const tabId=panelId.replace('panel-','');
    const header=panel.querySelector('.panel-header');
    if(!header) return;
    const filename='WBR_W''' + str(CURRENT_WEEK) + '''_'+tabId+'.csv';
    let btnBar=panel.querySelector('.btn-bar');
    if(btnBar){
      const sep=document.createElement('span');
      sep.className='filter-sep';sep.textContent='|';
      btnBar.appendChild(sep);
      const btn=document.createElement('button');
      btn.className='export-btn';
      btn.innerHTML='\\u{1F4E5} Export CSV';
      btn.onclick=function(){exportPanel(panelId,filename,this);};
      btnBar.appendChild(btn);
    } else {
      const bar=document.createElement('div');
      bar.className='export-bar';
      bar.innerHTML='<button class="export-btn" onclick="exportPanel(\\''+panelId+'\\',\\''+filename+'\\',this)">\\u{1F4E5} Export CSV</button>';
      const toolbar=panel.querySelector('.toolbar');
      if(toolbar) toolbar.parentNode.insertBefore(bar,toolbar.nextSibling);
      else header.parentNode.insertBefore(bar,header.nextSibling);
    }
  });
})();
</script>
'''

    html += '</body>\n</html>'
    return html


# ─── AES Encryption (CryptoJS compatible) ────────────────────────────────

def evp_bytes_to_key(password, salt, key_len=32, iv_len=16):
    """CryptoJS-compatible key derivation (EVP_BytesToKey with MD5)"""
    pwd = password.encode('utf-8')
    derived = b''
    block = b''
    while len(derived) < key_len + iv_len:
        block = hashlib.md5(block + pwd + salt).digest()
        derived += block
    return derived[:key_len], derived[key_len:key_len+iv_len]


def aes_encrypt_cryptojs(plaintext, password):
    """Encrypt text compatible with CryptoJS.AES.decrypt()"""
    from Crypto.Cipher import AES
    from Crypto.Util.Padding import pad

    salt = os.urandom(8)
    key, iv = evp_bytes_to_key(password, salt)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    padded = pad(plaintext.encode('utf-8'), AES.block_size)
    ciphertext = cipher.encrypt(padded)
    # CryptoJS format: "Salted__" + salt + ciphertext
    result = b'Salted__' + salt + ciphertext
    return base64.b64encode(result).decode('ascii')


# ─── Login wrapper ────────────────────────────────────────────────────────

def wrap_with_login(encrypted_data):
    """Create the login HTML wrapper"""
    return f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WBR W{CURRENT_WEEK} {CURRENT_YEAR} - Seller Report</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body.login-mode {{
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  background: #f0f2f5; color: #333;
  display: flex; justify-content: center; align-items: center; min-height: 100vh;
}}
.login-container {{
  background: #fff; padding: 40px 36px; border-radius: 12px;
  box-shadow: 0 4px 24px rgba(0,0,0,.12); text-align: center;
  max-width: 400px; width: 90%;
}}
.login-container h1 {{ font-size: 20px; color: #1a3a5c; margin-bottom: 8px; }}
.login-container p {{ font-size: 13px; color: #888; margin-bottom: 24px; }}
.pw-wrap {{ position: relative; width: 100%; }}
.pw-wrap input {{
  width: 100%; padding: 12px 16px; padding-right: 44px; border: 1px solid #ccc; border-radius: 6px;
  font-size: 14px; outline: none; transition: border-color .2s;
}}
.pw-wrap input:focus {{ border-color: #2d5f8a; }}
.pw-toggle {{
  position: absolute; right: 12px; top: 50%; transform: translateY(-50%);
  background: none; border: none; cursor: pointer; font-size: 18px; color: #888;
  padding: 0; line-height: 1;
}}
.pw-toggle:hover {{ color: #333; }}
.login-container > button {{
  width: 100%; padding: 12px; margin-top: 16px; background: #2d5f8a; color: #fff;
  border: none; border-radius: 6px; font-size: 14px; cursor: pointer; transition: background .2s;
}}
.login-container > button:hover {{ background: #1a3a5c; }}
.error-msg {{ color: #c0392b; font-size: 13px; margin-top: 12px; display: none; }}
.lock-icon {{ font-size: 48px; margin-bottom: 16px; }}
</style>
</head>
<body class="login-mode">
<div class="login-container" id="loginBox">
  <div class="lock-icon">&#x1F512;</div>
  <h1>WBR W{CURRENT_WEEK} {CURRENT_YEAR} Expansion Dashboard</h1>
  <p>This content is password protected.</p>
  <div class="pw-wrap">
    <input type="password" id="pwInput" placeholder="Enter password" autocomplete="off"
           onkeydown="if(event.key==='Enter')unlock()">
    <button type="button" class="pw-toggle" onclick="togglePw()" title="Show/Hide password">&#x1F441;</button>
  </div>
  <button onclick="unlock()">Unlock</button>
  <div class="error-msg" id="errMsg">Incorrect password. Please try again.</div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/crypto-js/4.2.0/crypto-js.min.js"></script>
<script>
var ED="{encrypted_data}";
function togglePw(){{
  var inp=document.getElementById("pwInput");
  inp.type=inp.type==="password"?"text":"password";
  inp.focus();
}}
function unlock(){{
  var pw=document.getElementById("pwInput").value;
  var err=document.getElementById("errMsg");
  err.style.display="none";
  if(typeof CryptoJS==="undefined"){{err.textContent="Error: encryption library failed to load. Please refresh.";err.style.display="block";return;}}
  try{{
    var dec=CryptoJS.AES.decrypt(ED,pw).toString(CryptoJS.enc.Utf8);
    if(!dec||dec.length<50)throw new Error("empty result");
    document.open();document.write(dec);document.close();
  }}catch(e){{
    err.style.display="block";
    document.getElementById("pwInput").value="";
    document.getElementById("pwInput").focus();
  }}
}}
</script>
</body>
</html>'''


# ─── Main ─────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print(f"  WBR W{CURRENT_WEEK} {CURRENT_YEAR} Seller Report Generator")
    print("=" * 60)

    # Load data
    all_rows = load_data()

    # Filter TW sellers
    print("\nFiltering seller_origin=TW...")
    tw_rows = filter_tw(all_rows)
    print(f"  TW rows: {len(tw_rows)}")

    # Build seller data
    print("\nBuilding seller data structures...")
    sellers = build_seller_data(tw_rows)

    # Auto-detect marketplace IDs
    print("\nDetecting marketplace IDs...")
    auto_detect_mp_ids(sellers)

    # Generate HTML
    print("\nGenerating dashboard HTML...")
    dashboard_html = generate_dashboard_html(sellers, tw_rows)
    print(f"  Dashboard HTML size: {len(dashboard_html):,} chars")

    # Encrypt
    print("\nEncrypting with AES...")
    try:
        encrypted = aes_encrypt_cryptojs(dashboard_html, PASSWORD)
        print(f"  Encrypted data size: {len(encrypted):,} chars")
    except ImportError:
        print("  WARNING: pycryptodome not available, trying pycryptodomex...")
        try:
            from Cryptodome.Cipher import AES
            from Cryptodome.Util.Padding import pad

            salt = os.urandom(8)
            key, iv = evp_bytes_to_key(PASSWORD, salt)
            cipher = AES.new(key, AES.MODE_CBC, iv)
            padded = pad(dashboard_html.encode('utf-8'), AES.block_size)
            ciphertext = cipher.encrypt(padded)
            result = b'Salted__' + salt + ciphertext
            encrypted = base64.b64encode(result).decode('ascii')
            print(f"  Encrypted data size: {len(encrypted):,} chars")
        except ImportError:
            print("  ERROR: Neither pycryptodome nor pycryptodomex found!")
            print("  Installing pycryptodome...")
            import subprocess
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pycryptodome', '-q'])
            from Crypto.Cipher import AES
            from Crypto.Util.Padding import pad

            salt = os.urandom(8)
            key, iv = evp_bytes_to_key(PASSWORD, salt)
            cipher = AES.new(key, AES.MODE_CBC, iv)
            padded = pad(dashboard_html.encode('utf-8'), AES.block_size)
            ciphertext = cipher.encrypt(padded)
            result = b'Salted__' + salt + ciphertext
            encrypted = base64.b64encode(result).decode('ascii')
            print(f"  Encrypted data size: {len(encrypted):,} chars")

    # Generate final HTML with login wrapper
    print("\nGenerating final output...")
    final_html = wrap_with_login(encrypted)
    print(f"  Final HTML size: {len(final_html):,} chars")

    # Write output
    with open(str(OUTPUT_PATH), 'w', encoding='utf-8') as f:
        f.write(final_html)
    print(f"\n  Written to: {OUTPUT_PATH}")
    print(f"  File size: {os.path.getsize(str(OUTPUT_PATH)):,} bytes")
    print("\nDone!")


if __name__ == '__main__':
    main()
