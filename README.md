# TW2 Expansion Dashboard

WBR / MBR Pipeline reports are hosted on GitHub Pages with password protection.

🏠 **首頁（各報表入口）：** https://kaojia.github.io/expansion-dashboard/

🔒 **Amazon 內部 Midway 版（Protozoa mirror，免密碼）：** https://protozoa.amazon.dev/prototypes/f8410751-2ea6-4ad8-b478-a9af0e867bba/ — 每次 push 後由 `../tools/protozoa_sync.py` 同步

🔗 **NSR Dashboard (Latest)：** https://kaojia.github.io/expansion-dashboard/seller-report.html

🔗 **NSR Dashboard (All Weeks)：** https://kaojia.github.io/expansion-dashboard/nsr/

🔗 **EM WBR：** https://kaojia.github.io/expansion-dashboard/wbr/

🔗 **NSR MBR Dashboard：** https://kaojia.github.io/expansion-dashboard/mbr/

🔗 **Decliner Analysis (Weekly)：** https://kaojia.github.io/expansion-dashboard/decliner/

> 需要輸入密碼才能查看內容。

## NSR MBR Dashboard 內容

- 📈 **Expansion DSR** — TW2 Expansion DSR GS MBR 總表（Monthly：MoM / YoY）+ Executive Summary
- 📊 **Movers & Shakers** — EU5/JP/AU/MENA Top 10 Gainers & Decliners（MoM Delta）
- **MEA / EU / JP** — 各市場 NSR/ESM Seller-Level GMS 明細（含 Channel、Owner 篩選、Copy MCIDs、匯出 CSV）

### 目前可用報告

| Month | Link |
|-------|------|
| Mar 2026 | [MBR Mar 2026](https://kaojia.github.io/expansion-dashboard/mbr/March/MBR_March_2026_Expansion_Dashboard.html) |
| Apr 2026 | [MBR Apr 2026](https://kaojia.github.io/expansion-dashboard/mbr/Apr/MBR_Apr_2026_Expansion_Dashboard.html) |
| May 2026 | [MBR May 2026](https://kaojia.github.io/expansion-dashboard/mbr/May/MBR_May_2026_Expansion_Dashboard.html) |
| June 2026 | [MBR June 2026](https://kaojia.github.io/expansion-dashboard/mbr/June/MBR_June_2026_Expansion_Dashboard.html) |
| Jul 2026 | [MBR Jul 2026](https://kaojia.github.io/expansion-dashboard/mbr/Jul/MBR_Jul_2026_Expansion_Dashboard.html) |
| Aug 2026 | [MBR Aug 2026](https://kaojia.github.io/expansion-dashboard/mbr/August/MBR_August_2026_Expansion_Dashboard.html) |

## WBR Dashboard 內容

- 📈 **Expansion DSR** — TW2 Expansion DSR GS WBR 總表 + Executive Summary
- 📊 **Movers & Shakers** — EU5/JP/AU/AE/SA Top 10 Gainers & Decliners
- **MEA / EU / JP** — 各市場 NSR/ESM Seller GMS 明細（含 Channel、Owner 篩選）

## 每週更新流程

### NSR Dashboard

由 `nsr-weekly-dashboard` skill 執行（`../tools/gen_dashboard.py` + `../tools/update_seller_panels.py`），輸出 `seller-report.html` + `nsr/W##_NSR_Dashboard.html`，push 後跑 `../tools/protozoa_sync.py`。

> 舊流程 `generate_weekly_report.py`（整頁 AES 加密的根目錄 `index.html`）已停用，最後版本為 2026-08-12；根目錄現為報表入口頁。

### WBR Pipeline

```bash
# 1. 將新的 WBR HTML 放到 wbr/W##/ 資料夾
# 2. Push 到 GitHub
git add wbr/
git commit -m "W## 2026 update"
git push origin master

# 3. 產生本地無密碼版本
python wbr/publish.py
```

## NSR MBR 更新流程

```bash
# 1. 產生本地版（無密碼）
cd 2026 && MBR_MONTH=3 python gen_mbr_dashboard.py

# 2. 複製到 mbr/<Mon>/ 並自動注入 auth.js
python mbr/publish.py March          # 不給參數則發佈所有找到的月份

# 3. 確認 mbr/index.html 的 months 陣列有列出該月份

# 4. Push 到 GitHub
git add mbr/
git commit -m "MBR March 2026 update"
git push origin master
```

`mbr/publish.py` 會自動注入 `auth.js`，不需手動確認。本地未加密版本保留在
`2026/<Mon>/MBR_<Mon>_2026_Expansion_Dashboard_local.html`。
