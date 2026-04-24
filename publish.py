"""
publish.py — Auto-update index.html and push WBR reports to GitHub.

Usage:
    python publish.py          # auto-detect all W## folders
    python publish.py W17      # only publish W17 (still updates index with all weeks)

What it does:
  1. Scans all W## folders for HTML reports
  2. Rebuilds the `weeks` array in index.html
  3. git add + commit + push
"""
import glob, os, re, sys

# ── 1. Discover all weeks with HTML reports ────────────────────
week_dirs = sorted(
    [d for d in glob.glob('W[0-9]*') if os.path.isdir(d)],
    key=lambda d: int(re.sub(r'\D', '', d) or '0')
)

weeks = []
for d in week_dirs:
    wnum = int(re.sub(r'\D', '', d))
    htmls = glob.glob(os.path.join(d, 'W*_WBR_*_Pipeline.html'))
    if not htmls:
        continue
    # Extract market codes from filenames like W16_WBR_AE_Pipeline.html
    markets = sorted(set(
        re.search(r'_WBR_(\w+)_Pipeline', os.path.basename(f)).group(1)
        for f in htmls
        if re.search(r'_WBR_(\w+)_Pipeline', os.path.basename(f))
    ))
    if markets:
        weeks.append({'week': 'W%d' % wnum, 'num': wnum, 'year': 2026, 'markets': markets})

if not weeks:
    print("ERROR: No WBR HTML reports found in any W## folder.")
    sys.exit(1)

print("Found %d weeks with reports:" % len(weeks))
for w in weeks:
    print("  %s %d — markets: %s" % (w['week'], w['year'], ', '.join(w['markets'])))

# ── 2. Update index.html ──────────────────────────────────────
INDEX_FILE = 'index.html'
if not os.path.exists(INDEX_FILE):
    print("ERROR: %s not found." % INDEX_FILE)
    sys.exit(1)

with open(INDEX_FILE, 'r', encoding='utf-8') as f:
    html = f.read()

# Build new weeks JS array
js_entries = []
for w in weeks:
    markets_str = ', '.join('"%s"' % m for m in w['markets'])
    js_entries.append('  { week: "%s", year: %d, markets: [%s] },' % (w['week'], w['year'], markets_str))

new_weeks_block = 'const weeks = [\n' + '\n'.join(js_entries) + '\n];'

# Replace the existing weeks array in index.html
pattern = r'const weeks = \[[\s\S]*?\];'
if not re.search(pattern, html):
    print("ERROR: Could not find 'const weeks = [...]' in index.html")
    sys.exit(1)

html_new = re.sub(pattern, new_weeks_block, html, count=1)

with open(INDEX_FILE, 'w', encoding='utf-8') as f:
    f.write(html_new)

print("\nUpdated index.html with %d weeks." % len(weeks))

# ── 3. Git add, commit, push ─────────────────────────────────
# Determine which week to mention in commit message
if len(sys.argv) > 1:
    target_week = sys.argv[1].upper()
else:
    target_week = weeks[-1]['week']

# Stage HTML files and index
html_files = []
for w in weeks:
    html_files.extend(glob.glob(os.path.join(w['week'], '*.html')))

stage_files = html_files + [INDEX_FILE]
stage_cmd = 'git add ' + ' '.join('"%s"' % f for f in stage_files)
print("\n" + stage_cmd)
os.system(stage_cmd)

# Check if there are changes to commit
status = os.popen('git status --porcelain').read().strip()
if not status:
    print("\nNo changes to commit. Everything is up to date.")
    sys.exit(0)

commit_msg = "Update WBR: %s reports" % target_week
print('git commit -m "%s"' % commit_msg)
os.system('git commit -m "%s"' % commit_msg)

print('git push')
ret = os.system('git push')
if ret == 0:
    print("\n✅ Published successfully! Check: https://kaojia.github.io/expansion-dashboard/")
else:
    print("\n⚠️  Push failed. You may need to run 'git push' manually.")
