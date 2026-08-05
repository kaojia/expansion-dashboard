# TW2 Expansion Dashboard

WBR / MBR Pipeline reports are hosted on GitHub Pages with password protection.

🔗 **NSR Dashboard (Latest)：** https://kaojia.github.io/expansion-dashboard/seller-report.html

🔗 **NSR Dashboard (All Weeks)：** https://kaojia.github.io/expansion-dashboard/nsr/

🔗 **EM WBR：** https://kaojia.github.io/expansion-dashboard/wbr/

🔗 **MBR Dashboard：** https://kaojia.github.io/expansion-dashboard/mbr/

🔗 **Decliner Analysis (Weekly)：** https://kaojia.github.io/expansion-dashboard/decliner/

> 需要輸入密碼才能查看內容（Decliner Analysis 除外）。

## MBR Dashboard 內容

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

## WBR Dashboard 內容

- 📈 **Expansion DSR** — TW2 Expansion DSR GS WBR 總表 + Executive Summary
- 📊 **Movers & Shakers** — EU5/JP/AU/AE/SA Top 10 Gainers & Decliners
- **MEA / EU / JP** — 各市場 NSR/ESM Seller GMS 明細（含 Channel、Owner 篩選）

## 每週更新流程

### Expansion Dashboard

```bash
python generate_weekly_report.py          # 自動偵測最新週次
python generate_weekly_report.py W17      # 指定週次
```

腳本會自動生成加密版推送到 GitHub Pages，同時產生本地無密碼版。

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

## MBR 更新流程

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
