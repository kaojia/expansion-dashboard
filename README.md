# TW2 Expansion WBR Dashboard

WBR Pipeline reports are hosted on GitHub Pages with password protection.

🔗 **WBR Pipeline：** https://kaojia.github.io/expansion-dashboard/wbr/

> 需要輸入密碼才能查看內容。

## 本地無密碼版本

```bash
python wbr/publish.py
```

產生的檔案在 `wbr/local/`，直接開 `wbr/local/index.html` 即可瀏覽。

## 每週更新流程

```bash
# 1. 將新的 WBR HTML 放到 wbr/W##/ 資料夾
# 2. Push 到 GitHub
git add wbr/
git commit -m "W## 2026 update"
git push origin master

# 3. 產生本地無密碼版本
python wbr/publish.py
```
