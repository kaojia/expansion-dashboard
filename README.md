# TW2 Expansion Dashboard

WBR / MBR Pipeline reports are hosted on GitHub Pages with password protection.

🔗 **Expansion Dashboard：** https://kaojia.github.io/expansion-dashboard/seller-report.html

🔗 **WBR Pipeline：** https://kaojia.github.io/expansion-dashboard/wbr/

🔗 **MBR Dashboard：** https://kaojia.github.io/expansion-dashboard/mbr/

> 需要輸入密碼才能查看內容。

## MBR Dashboard 內容

- 📈 **Expansion DSR** — TW2 Expansion DSR GS MBR 總表（Monthly：MoM / YoY）+ Executive Summary
- 📊 **Movers & Shakers** — EU5/JP/AU/MENA Top 10 Gainers & Decliners（MoM Delta）
- **MEA / EU / JP** — 各市場 NSR/ESM Seller-Level GMS 明細（含 Channel、Owner 篩選、Copy MCIDs、匯出 CSV）

### 目前可用報告

| Month | Link |
|-------|------|
| Apr 2026 | [MBR Apr 2026](https://kaojia.github.io/expansion-dashboard/mbr/Apr/MBR_Apr_2026_Expansion_Dashboard.html) |

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
# 1. 將 MBR HTML 放到 mbr/Mon/ 資料夾（例如 mbr/Apr/）
# 2. 確認檔案頂部有 auth.js 引用
# 3. Push 到 GitHub
git add mbr/
git commit -m "MBR Apr 2026 update"
git push origin master
```

本地未加密版本保留在 `April/MBR_Apr_2026_Expansion_Dashboard_local.html`。
