# Maintenance ONE

> 廠內電路圖、設備程序書、IO 對照表與維修紀錄的統一入口。
> Unified portal for plant schematics, equipment procedures, IO description lists, and maintenance records.

**🔗 正式網站：<https://amtoppme.github.io/Maintenance-Smart/>**

---

## 這是什麼

一個**純靜態網頁**系統，把原本散落各處的技術文件集中在一個入口，並提供維修紀錄的線上登錄與查詢。

- 沒有伺服器要維護
- 沒有資料庫主機
- 沒有編譯（build）流程
- 更新內容 = 換檔案 → 用 GitHub Desktop 上傳

---

## 功能模組

| 模組 | 說明 |
|---|---|
| **Updated Schematic** | 電路圖檢視（F12 / F13 / F16 / F201 × Line 1~4） |
| **IO Description List** | IO 對照表互動查詢（CSV 即時載入） |
| **Procedure** | 設備程序書 —— Blending、Extruder、Die、Chill Roll、MDO、TDO、PRS、Treaters、Winder、Grinder、Reclaim |
| **Training** | 教育訓練資料 |
| **SCADA PC / Inventory** | SCADA 電腦資產清單 |
| **Records** | 維修紀錄查詢與編輯 |
| **Report** | 新增維修報告 |

---

## 系統架構

```
瀏覽器
  └─ index.html（Portal 外殼：選單、主題、iframe）
       ├─ 靜態內容 ── PDF / CSV（本 repo，由 GitHub Pages 提供）
       ├─ localStorage ── 主題設定與紀錄快取
       └─ Google Apps Script ×2 ── Google Sheets（維修紀錄、SCADA PC 清單）
```

技術組成：原生 HTML / CSS / JavaScript（樣式與腳本內嵌於各頁面）、[PDF.js](https://mozilla.github.io/pdf.js/)、Google Apps Script + Google Sheets。

---

## 目錄結構

```
.
├── index.html                  Portal 外殼
├── Home.html                   首頁 / Quality Policy
├── Updated Schematic.html      電路圖檢視器
├── IO_Description_List.html    IO 描述查詢
├── IO_List.html                IO 清單
├── CalibrationProcedure.html   Procedure 目錄
├── Training.html               教育訓練
├── SCADA_PC_Inventory.html     SCADA PC 資產
├── Maintenance_Records.html    維修紀錄
├── New_Report.html             新增報告
├── about.html                  關於
├── export_json.py              維修紀錄 → Excel 統計報表
│
├── assets/                     背景圖與圖示
├── pdfjs/                      內建 PDF 檢視器（勿更動）
│
├── docs/
│   ├── schematic/              電路圖 PDF
│   ├── procedures/             設備程序書 PDF
│   └── MAINTENANCE_PROCEDURE.md  ★ 維護交接手冊
│
└── IODescriptionList/
    ├── index.json              CSV 路徑對照表（手動維護）
    └── F12/ F13/ F16/ F201/    各 Section 的 CSV 與 PDF
```

---

## 更新內容

### ⚠️ 最重要的一條規則

> ### **新檔案的名稱必須與舊檔案「一模一樣」，直接覆蓋。**
>
> 網頁裡的連結是寫死檔名的。差一個空白、一個底線、一個大小寫，
> 使用者點下去就會變成 404。原檔名有錯字也照抄。

### 標準流程

1. GitHub Desktop → **Pull origin**
2. 把新檔案改成與舊檔**完全相同**的檔名，覆蓋進對應資料夾
   - 電路圖 → `docs/schematic/`
   - 程序書 → `docs/procedures/`
   - IO 清單 → `IODescriptionList/<Section>/<Category>/`
3. **本機驗證**（不要雙擊 HTML，`file://` 會讓 CSV/JSON 載入失敗）：
   ```bash
   python -m http.server 8080
   ```
   開 <http://localhost:8080/index.html>，走一次使用者的操作路徑
4. GitHub Desktop → 填寫 Summary → **Commit to main** → **Push origin**
5. 等 1~3 分鐘，開正式網站按 `Ctrl + F5` 確認

> 新增（而非取代）CSV 時，還要一併編輯 `IODescriptionList/index.json`。

### Commit 訊息格式

```
<動作>: <對象> - <內容>
```

| ✅ 好 | ❌ 不好 |
|---|---|
| `Update: Line-4 F13 schematic - rev B 2026-08` | `Update index.html` |
| `Fix: CalibrationProcedure - 修正 Bardac PDF 檔名` | `updated file` |

---

## 健康檢查

在 repo 根目錄用 Git Bash 執行。

**PDF 連結完整性：**
```bash
grep -oh -E "docs/(schematic|procedures)/[^\"']*\.pdf" *.html | grep -v '\$' | sort -u | while IFS= read -r f; do [ -f "$f" ] || echo "MISSING: $f"; done; echo "--- 檢查完成 ---"
```

**IO 對照表完整性：**
```bash
python -X utf8 -c "import json,os,glob;d=json.load(open('IODescriptionList/index.json',encoding='utf-8'));t=[];w=lambda o:[w(v) for v in o.values()] if isinstance(o,dict) else (t.append(o) if isinstance(o,str) and o.endswith('.csv') else None);w(d);m=[p for p in t if not os.path.isfile(p)];print(f'登錄 {len(t)} / 磁碟 {len(glob.glob(\"IODescriptionList/**/*.csv\",recursive=True))} / 找不到 {len(m)} 筆');[print('  MISSING:',x) for x in m]"
```

> `-X utf8` 不可省略 —— Windows 主控台預設編碼無法輸出中文，省略會出現 `UnicodeEncodeError`。

第一項應回報零缺漏；第二項目前已知有 11 筆待處理，詳見[交接手冊 §8.5](docs/MAINTENANCE_PROCEDURE.md)。

---

## 統計報表

由 Records 頁面下載 `maintenance_logs.json` 後：

```bash
pip install pandas openpyxl
```
```bash
python export_json.py
```

產出 `Maintenance_Report_From_JSON.xlsx`，含依產線、月份、類別、設備的停機統計共 7 個工作表。

---

## 維護與交接

完整的維護程序、備份還原、故障排除與已知問題，請見：

### 📘 [系統維護交接手冊](docs/MAINTENANCE_PROCEDURE.md)

內容包含：接手第一週待辦、權限清單、日/週/月/季維護排程、更新 SOP、
Google Sheets 備份與還原、故障排除對照表、9 項已知風險與處理方案。

> ⚠️ **接手者請先讀該手冊的 §0 與 §8。**

---

## 注意事項

- 本網站與 repo 內全部文件目前**公開於網際網路**（GitHub Pages 免費方案僅支援 public repository）
- Repo 體積較大（`.git` 約 1.9 GB），首次 clone 請使用有線網路
- 維修紀錄的真實資料存放於 Google Sheets，**不在本 repo 內**，需另行備份

---

<sub>Inteplast — Maintenance ONE (Maintenance Smart)</sub>
