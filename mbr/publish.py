# -*- coding: utf-8 -*-
"""
publish.py - Stage a locally generated MBR dashboard into mbr/<Mon>/ for GitHub Pages.

Usage:
    python mbr/publish.py March June          # publish specific months
    python mbr/publish.py                     # publish every month found locally

What it does for each month:
1. Copies 2026/<Mon>/MBR_<Mon>_2026_Expansion_Dashboard_local.html
   to mbr/<Mon>/MBR_<Mon>_2026_Expansion_Dashboard.html
2. Injects <script src="../auth.js"></script> right after <meta charset>
   so the page redirects to the login screen when not authenticated
3. Reports which months are now staged, so index.html can be checked

The README used to describe this as a manual copy-and-remember-auth.js step. That is
how mbr/June went live without the guard, leaving it publicly readable, so the
injection is automated here instead.
"""
import os
import re
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))       # .../temp_repo/mbr
REPO_DIR = os.path.dirname(SCRIPT_DIR)                        # .../temp_repo
LOCAL_ROOT = os.path.dirname(REPO_DIR)                        # .../2026

AUTH_TAG = '<script src="../auth.js"></script>'


def local_source(month):
    """Path of the generator's no-auth output for a month, or None."""
    path = os.path.join(
        LOCAL_ROOT, month, "MBR_%s_2026_Expansion_Dashboard_local.html" % month)
    return path if os.path.exists(path) else None


def inject_auth(html):
    """Ensure exactly one auth.js tag, placed just after <meta charset>."""
    # Drop any existing tag first so re-publishing cannot stack duplicates.
    html = re.sub(r'<script\s+src=["\']\.\.?/?auth\.js["\']\s*>\s*</script>\s*\n?',
                  '', html)
    m = re.search(r'<meta\s+charset=[^>]*>\s*\n?', html, re.I)
    if not m:
        raise SystemExit("No <meta charset> found -- cannot place auth.js safely.")
    return html[:m.end()] + AUTH_TAG + "\n" + html[m.end():]


def publish(month):
    src = local_source(month)
    if not src:
        print("  skip %s (no local dashboard found)" % month)
        return False

    dst_dir = os.path.join(SCRIPT_DIR, month)
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, "MBR_%s_2026_Expansion_Dashboard.html" % month)

    with open(src, "r", encoding="utf-8") as f:
        html = f.read()
    html = inject_auth(html)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(html)

    print("  OK {}/{} ({:,} bytes)".format(
        month, os.path.basename(dst), os.path.getsize(dst)))
    return True


def main():
    months = sys.argv[1:]
    if not months:
        months = [m for m in ("January", "February", "March", "April", "May",
                              "June", "July", "August", "September", "October",
                              "November", "December") if local_source(m)]
        if not months:
            raise SystemExit("No local MBR dashboards found under %s" % LOCAL_ROOT)

    print("Publishing: %s" % ", ".join(months))
    published = [m for m in months if publish(m)]

    print("")
    print("Staged %d month(s). Verify mbr/index.html lists them, then:" % len(published))
    print("  git add mbr/ && git commit && git push origin master")


if __name__ == "__main__":
    main()
