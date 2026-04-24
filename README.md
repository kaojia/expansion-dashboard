# TW2 Expansion DSR GS WBR Dashboard

🔗 **Dashboard 網址：** https://kaojia.github.io/expansion-dashboard/

🔗 **Seller Report 網址：** https://kaojia.github.io/expansion-dashboard/seller-report.html

🔗 **WBR Pipeline 網址：** https://kaojia.github.io/expansion-dashboard/wbr/

> 需要輸入密碼才能查看內容。

## 每週更新流程

```bash
# 1. 加密新的 HTML 報告
python encrypt_html.py "WBR_W##_2026_Expansion_DSR_from_Page0.html" "expansionwbr" "C:\Users\chiawenk\Documents\expansion-dashboard\index.html"

# 2. Push 到 GitHub（網址不變）
cd C:\Users\chiawenk\Documents\expansion-dashboard
git add index.html
git commit -m "W## 2026 update"
git push
```
