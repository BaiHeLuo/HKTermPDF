# 輸入勞工終止僱傭合約通知書 — 批量生成器

批量生成香港「補充勞工優化計劃」下的 **輸入勞工終止僱傭合約通知書** (ESLS-13) PDF 文件。

上傳一份包含勞工名單的 Excel 表格，選擇簽署日期，即可一鍵產出所有勞工的通知書。

> **完全離線運行** — 所有數據僅在本地處理，不會上傳至任何伺服器。

---

## 功能特點

- 智能識別 Excel 表頭（列順序不限，自動按名稱匹配）
- 逐位勞工數據驗證，前端顯示詳細錯誤提示
- 批量生成 PDF，以 ZIP 壓縮包下載
- 支援 Windows 7 / 8 / 10 / 11
- 無需安裝 Python 或任何依賴（打包後的 exe）

---

## 文件說明

| 文件 | 說明 |
|---|---|
| `app.py` | Flask Web 應用，包含前端界面和後端邏輯 |
| `pdf_generator.py` | PDF 生成核心邏輯，Excel 讀取與數據驗證 |
| `app.spec` | PyInstaller 打包配置文件 |
| `build_win7.py` | Win7 兼容打包腳本（自動安裝 Python 3.8 並打包） |
| `使用说明.txt` | 面向最終用戶的簡體中文使用說明 |

---

## 使用前提

使用本程序前，你需要準備：

1. **PDF 模板** — 從勞工處官網下載 ESLS-13 表格 PDF，命名為 `勞工終止合約通知書-.pdf`，放在程序同目錄下
2. **Excel 文件** — 包含以下表頭列（列順序不限）：
   - 中文姓名
   - 英文姓名
   - 約滿日期（格式：`yyyy.mm.dd`，如 `2026.06.20`）
   - 入境事務處檔號
   - 香港身份證號碼
   - SLS（僱傭合約號碼，如 `ESLS004030`）
   - 勞工處檔案編號

---

## 開發環境運行

```bash
# 安裝依賴（Python 3.9+）
pip install flask openpyxl PyMuPDF

# 啟動
python app.py
```

瀏覽器訪問 http://localhost:5000

---

## 打包為 exe

### Windows 10/11（使用當前 Python）

```bash
pip install pyinstaller
pyinstaller app.spec --noconfirm
```

### Windows 7 兼容（自動使用 Python 3.8）

```bash
python build_win7.py
```

腳本會自動下載 Python 3.8、安裝依賴、打包為獨立 exe。
生成的 exe 位於 `dist/勞工通知書生成器.exe`。

---

## 技術棧

- **Python** + **Flask** — Web 框架
- **PyMuPDF (fitz)** — PDF 表單填寫
- **openpyxl** — Excel 讀取
- **PyInstaller** — 打包為獨立可執行文件

---

## 授權

MIT License — 自由使用、修改和分發。
