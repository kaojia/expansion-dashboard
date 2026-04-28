# TW2 Expansion WBR Dashboard

WBR Pipeline reports are hosted on GitHub Pages with password protection.

🔗 **Expansion Dashboard：** https://kaojia.github.io/expansion-dashboard/seller-report.html

🔗 **WBR Pipeline：** https://kaojia.github.io/expansion-dashboard/wbr/

> 需要輸入密碼才能查看內容。

## Expansion Dashboard 內容

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
