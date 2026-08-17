# Maintenance ONE (Maintenance Smart) — 系統維護交接手冊

> **這份文件是寫給接手本系統維護工作的人。**
> 你不需要事先了解這個系統，也不需要會寫程式。照著步驟做即可。
> 遇到本文件沒寫到的狀況，請先看 §7 故障排除，再看 §10 求助管道。

| 項目 | 內容 |
|---|---|
| 系統名稱 | Maintenance ONE（頁面標題有時顯示 Maintenance Smart，是同一套系統） |
| 用途 | 廠內電路圖、設備 Procedure、IO 清單、維修紀錄的統一入口 |
| **正式網站** | <https://amtoppme.github.io/Maintenance-Smart/> |
| 託管方式 | GitHub Pages（免費、由 GitHub 直接提供，公司內沒有伺服器需要維護） |
| GitHub Repository | `AMTOPPME/Maintenance-Smart`（分支：`main`） |
| 本機工作資料夾 | `C:\Users\Franklin\Documents\GitHub2\Maintenance-Smart` |
| 主要工具 | **GitHub Desktop**（不需要用命令列） |
| 文件版本 | v1.4 |
| 最後更新 | 2026-08-17 |

---

## 0. 接手第一週該做的事

按順序完成，全部做完你就能獨立維護這套系統。

| # | 事項 | 對應章節 |
|---|---|---|
| 1 | 讀完 §1，理解系統由哪些部分組成 | §1 |
| 2 | 取得 GitHub 帳號存取權，安裝並登入 GitHub Desktop | §2、§3 |
| 3 | 確認本機資料夾能開啟，且 GitHub Desktop 看得到這個 repo | §3.2 |
| 4 | 取得 Google 帳號存取權（維修紀錄的資料存在這裡） | §2 |
| 5 | **實際做一次**檔案更新演練（找一個不重要的檔案試） | §4 |
| 6 | 做一次資料備份，確認你會操作 | §5.1 |
| 7 | 執行一次檢查腳本，了解目前系統健康狀況 | §6 |
| 8 | 閱讀 §8 已知問題，了解目前有哪些未解決的風險 | §8 |
| 9 | 與前任確認 §9.4 的待確認清單，填上答案 | §9.4 |

---

## 1. 這套系統是什麼

### 1.1 一句話說明

這是一個**純網頁**的資料入口，把廠內散落的電路圖、操作程序書、IO 對照表集中在一個網站上，另外提供維修紀錄的登錄與查詢。

### 1.2 它由三個部分組成

```
┌─────────────────────────────────────────────────────┐
│  ①  網頁檔案 + PDF + CSV                             │
│      存放在 GitHub，也就是你電腦上的那個資料夾          │
│      → 更新電路圖、Procedure、IO 清單 = 改這裡         │
└─────────────────────────────────────────────────────┘
┌─────────────────────────────────────────────────────┐
│  ②  Google Sheets（試算表）                          │
│      維修紀錄、SCADA PC 清單的「真正資料庫」            │
│      → 使用者在網頁上填的資料，最後都存到這裡           │
└─────────────────────────────────────────────────────┘
┌─────────────────────────────────────────────────────┐
│  ③  Google Apps Script                              │
│      網頁與 Google Sheets 之間的橋樑                   │
│      → 平常不用動它。動它之前先看 §4.5                 │
└─────────────────────────────────────────────────────┘
```

**沒有伺服器、沒有資料庫主機、不用安裝任何服務。**
你的日常維護工作 ≈ 90% 是「換檔案 → 用 GitHub Desktop 上傳」。

### 1.2.1 網站是怎麼上線的

網站由 **GitHub Pages** 直接提供服務：

```
你在本機換檔案
      ↓
GitHub Desktop → Commit → Push
      ↓
GitHub 自動重新發布（約 1～3 分鐘）
      ↓
https://amtoppme.github.io/Maintenance-Smart/  ← 使用者看到新版本
```

**重點：沒有測試環境。Push 出去就是正式上線。** 所以 §4.6 的本機驗證絕對不能跳過。

> ⚠️ **這個網站是公開在網際網路上的**，任何人知道網址就能瀏覽，不需要登入，也不限公司網路。
> 也就是說廠內電路圖、設備程序書目前對外公開。這是 GitHub Pages 免費方案的限制（免費版只支援公開 repo）。
> 相關風險與處理選項見 §8.7。

### 1.3 網頁模組一覽

| 使用者看到的名稱 | 對應檔案 | 資料從哪來 |
|---|---|---|
| （首頁）Inteplast Quality Policy | `Home.html` | 固定內容 |
| Updated Schematic（電路圖） | `Updated Schematic.html` | `docs/schematic/` 的 PDF |
| IO Description List | `IO_Description_List.html` | `IODescriptionList/` 的 CSV |
| IO List | `IO_List.html` | 固定內容 |
| Procedure（Blending / Extruder / Die / Chill Roll / MDO / TDO / PRS / Treaters / Winder / Grinder / Reclaim） | `CalibrationProcedure.html` | `docs/procedures/` 的 PDF |
| Training | `Training.html` | 固定內容 |
| SCADA PC / Inventory | `SCADA_PC_Inventory.html` | Google Sheets |
| Records（維修紀錄） | `Maintenance_Records.html` | Google Sheets |
| Report（新增報告） | `New_Report.html` | Google Sheets |
| About | `about.html` | 固定內容 |

`index.html` 是最外層的框架（左側選單、主題切換），其他頁面是被它用 iframe 嵌進來的。

### 1.4 後端連線位址（不用背，知道在哪即可）

程式碼裡有兩組 Google Apps Script 網址：

| 用途 | 寫在哪些檔案 | 網址前綴 |
|---|---|---|
| 維修紀錄 | `Maintenance_Records.html`、`New_Report.html` | `…/AKfycbx7CDsQJdIOmrgRF…/exec` |
| SCADA PC 清單 | `SCADA_PC_Inventory.html`、`New_Report.html` | `…/AKfycbznKRp5Q1y0XvdEN…/exec` |

### 1.5 瀏覽器本機儲存（localStorage）

網頁會在使用者瀏覽器存幾筆資料：

| 名稱 | 用途 | 清掉會怎樣 |
|---|---|---|
| `ms-maintenance-logs` | 維修紀錄的**快取** | 沒事，重新整理會從 Google Sheets 重抓 |
| `ms-theme` / `ms-style` | 深淺色、介面樣式 | 回到預設外觀 |
| `ms-quote-idx` / `ms-home-fx` | 首頁標語、特效 | 回到預設 |

> 使用者說「我的紀錄不見了」時，**第一步是叫他按 `Ctrl + F5` 重新整理**，不是叫他重打。真正的資料在 Google Sheets。

---

## 2. 你需要的權限與帳號

接手時務必逐項確認你已取得：

| # | 項目 | 用途 | 取得狀態 |
|---|---|---|---|
| 1 | GitHub 帳號 `AMTOPPME` 的存取權 | 上傳檔案更新 | ☐ |
| 2 | GitHub Desktop 已安裝並登入 | 日常操作工具 | ☐ |
| 3 | Google 帳號（維修紀錄 Sheets 的擁有者／編輯者） | 資料備份、查修 | ☐ |
| 4 | 維修紀錄 Google Sheets 連結 | 備份 | ☐ |
| 5 | SCADA PC Inventory Google Sheets 連結 | 備份 | ☐ |
| 6 | Google Apps Script 專案存取權 | 後端維護 | ☐ |

> ⚠️ **目前所有權限集中在一個帳號，是單點故障風險。** 接手後請與主管討論指派代理人（見 §8.3）。

---

## 3. 工具準備

### 3.1 安裝 GitHub Desktop

1. 到 <https://desktop.github.com/> 下載安裝
2. 開啟後選 **Sign in to GitHub.com**，用 `AMTOPPME` 帳號登入
3. 登入後選 **File → Clone repository → GitHub.com**，找到 `Maintenance-Smart`

> **注意：** 這個 repo 很大（約 2 GB），第一次 clone 需要較長時間，請用有線網路。
> 如果電腦上已有 `C:\Users\Franklin\Documents\GitHub2\Maintenance-Smart` 資料夾，改用 **File → Add local repository** 指到該資料夾即可，不用重新下載。

### 3.2 認識 GitHub Desktop 畫面

| 位置 | 名稱 | 用途 |
|---|---|---|
| 左上 | Current repository | 確認選的是 `Maintenance-Smart` |
| 左上中 | Current branch | 必須是 `main` |
| 右上 | Fetch / Pull / Push origin | 與網路上的版本同步 |
| 左側清單 | Changes | 你改了哪些檔案 |
| 左下 | Summary + Description | 寫這次改了什麼 |
| 左下 | **Commit to main** 按鈕 | 確認變更 |

### 3.3 每次開工前的動作

> **開始任何修改前，一定先按 `Pull origin`。**
> 這會把別人（或你在別台電腦）做的更新抓下來，避免衝突。

畫面右上若顯示 `Pull origin ↓ 3`，代表有 3 筆更新待下載，按下去。
若顯示 `Fetch origin`，按一下，等幾秒，它會自動變成 Pull 或維持原狀（代表已是最新）。

---

## 4. 內容更新程序（最常用，請熟記）

### 🔑 最重要的一條規則

> ## **新檔案的名稱必須與舊檔案「一模一樣」，直接覆蓋。**

**為什麼：** 網頁裡的連結是寫死檔名的。只要檔名差一個字、一個空白、一個大小寫，
使用者點下去就會變成「找不到檔案」。

**包含：**
- 大小寫（`Line` ≠ `line`）
- 空白與底線（`Download Procedure` ≠ `Download_Procedure`）
- 連字號位置
- 副檔名（`.pdf` ≠ `.PDF`）
- **即使原檔名有錯字也照抄**（例如 F16 的檔名 `Esixting` 是拼錯的 `Existing`，**不要改**）

> **如果新檔案的名稱不同怎麼辦？**
> 最簡單的做法：**把新檔案改名成舊檔案的名字**，再覆蓋。
> 真的必須換名字時，你就得同時去改 HTML 裡的連結 —— 見 §4.4，屬於進階作業。

---

### 4.1 更新 Schematic（電路圖）

**情境：** 現場線路改了，工程師給你一份新的 F13 Line 4 電路圖 PDF。

| 步驟 | 動作 | 說明 |
|---|---|---|
| 1 | 開 GitHub Desktop，按 **Pull origin** | 確保是最新版 |
| 2 | 開啟資料夾 `…\Maintenance-Smart\docs\schematic\` | GitHub Desktop 選單：Repository → Show in Explorer |
| 3 | 找到要取代的舊檔，**把檔名複製起來** | 例如 `Line-4-F13-Schematic.pdf` |
| 4 | 把新 PDF 改成**完全相同**的檔名 | 建議做法：先刪掉舊檔，把新檔拖進來後貼上剛複製的檔名 |
| 5 | Windows 詢問是否取代 → **選「取代目的地中的檔案」** | |
| 6 | 開啟網頁確認 | 見 §4.6 |
| 7 | 回 GitHub Desktop → Commit → Push | 見 §4.7 |

**檔名規則對照表：**

| Section | Line 1 | Line 2 | Line 3 | Line 4 |
|---|---|---|---|---|
| F12 | `Line-1-F12-Schematic.pdf` | `Line-2-F12-Schematic.pdf` | `Line-3-F12-Schematic.pdf` | `Line-4-F12-Schematic.pdf` |
| F13 | `Line-1-F13-Schematic.pdf` | `Line-2-F13-Schematic.pdf` | `Line-3-F13-Schematic.pdf` | `Line-4-F13-Schematic.pdf` |
| F16 | `Line-1-F16-Esixting Updated Schematic.pdf` | `Line-2-F16-…` | `Line-3-F16-…` | `Line-4-F16-…` |
| F201 | — | `Line-2-F201-Schematic.pdf` | — | — |

> F16 的四個檔案目前內容相同（都是 3.6 MB），如果只更新一條線，記得只換那一個檔。
> F201 目前只有 Line 2 一個檔案。

---

### 4.2 更新 Procedure（設備程序書 PDF）

**情境：** Bardac 驅動器的下載程序更新了。

| 步驟 | 動作 |
|---|---|
| 1 | GitHub Desktop → **Pull origin** |
| 2 | 開啟 `…\Maintenance-Smart\docs\procedures\` |
| 3 | 找到舊檔，**複製檔名** |
| 4 | 新檔改成完全相同檔名 → 覆蓋舊檔 |
| 5 | 網頁確認（§4.6）→ Commit & Push（§4.7） |

目前 `docs/procedures/` 有 40 個 PDF，涵蓋：AB PLC、Mitsubishi PLC、Siemens S5/S7、Bardac Drive、Dynisco UPR 900、Magelis HMI、Measurex、Maguire、Blending、Load Cell 校正等。

---

### 4.3 更新 IO Description List（CSV）

**情境：** IO 對照表有異動。

**⚠️ 這一項比較特別：除了換檔案，可能還要改一個對照檔。**

| 步驟 | 動作 |
|---|---|
| 1 | GitHub Desktop → **Pull origin** |
| 2 | 開啟 `…\Maintenance-Smart\IODescriptionList\<Section>\<Category>\` |
| 3 | 找到舊 CSV，複製檔名，新檔改成相同名稱後覆蓋 |
| 4 | **確認 CSV 存檔編碼是 UTF-8**（見下方警告） |
| 5 | 執行 §6.2 檢查腳本 |
| 6 | 網頁確認（§4.6）→ Commit & Push（§4.7） |

> ⚠️ **CSV 編碼陷阱：** 用 Excel 編輯後直接存檔，中文會變亂碼。
> 正確做法：Excel → 另存新檔 → 檔案類型選 **「CSV UTF-8 (逗號分隔)」**。

**如果是新增（而非取代）CSV：**
還必須手動編輯 `IODescriptionList\index.json`，加入新的一行，例如：
```json
"XE05": "IODescriptionList/F12/BIT/F12 - 11 - BIT - XE05.csv"
```
路徑字串要與實體檔名逐字元相同，且注意 JSON 格式的逗號（最後一項後面不能有逗號）。
改完務必執行 §6.2 驗證。

---

### 4.4 【進階】修改網頁本身

> 這一節需要基本 HTML 知識。若不熟悉，建議先找人協助，或參考 §10。

所有 CSS 與 JavaScript 都直接寫在各個 `.html` 檔案裡，沒有編譯流程 —— 存檔即生效。

**什麼情況需要改網頁：**
- 新增一條產線或 Section（要改 `index.html` 的選單）
- 新增一份 Procedure PDF（要在 `CalibrationProcedure.html` 加連結）
- 檔名真的必須更改（要同步改 HTML 裡的字串）
- 調整外觀或文字

**注意事項：**
- 改之前先確認 GitHub Desktop 的 Changes 清單是乾淨的，這樣改壞了可以用 **Discard changes** 一鍵還原
- `index.html` 有 82 KB、2000 多行，修改前先用搜尋（Ctrl+F）定位
- 改完一定要測淺色/深色兩種主題，以及手機寬度

---

### 4.5 【高風險】修改 Google Apps Script

> ⚠️ **動這個之前，一定先做 §5.1 的資料備份。**

| 步驟 | 動作 |
|---|---|
| 1 | 開啟對應的 Google Sheets → 擴充功能 → Apps Script |
| 2 | **先建立版本**：部署 → 管理部署作業 → 建立新版本（這是你的回滾點） |
| 3 | 修改程式碼 |
| 4 | 部署 → **編輯現有部署作業** → 版本選「新版本」 → 部署 |
| 5 | 測試四種操作：新增 / 查詢 / 修改 / 刪除，全部通過才算完成 |
| 6 | 有問題 → 立刻回到步驟 4，版本選回舊的 |

> 🔴 **千萬不要選「新增部署作業」**。那會產生**新的網址**，舊網址失效，網站立刻壞掉。
> 如果不小心做了，必須把新網址填回這 4 個地方：
>
> | 檔案 | 大約行號 | 變數名稱 |
> |---|---|---|
> | `Maintenance_Records.html` | 390 | `SHEET_API_URL` |
> | `New_Report.html` | 337 | `API_BASE_URL` |
> | `New_Report.html` | 342 | `INVENTORY_API_URL` |
> | `SCADA_PC_Inventory.html` | 725 | `PC_SHEET_API_URL` |

---

### 4.6 上傳前的驗證（每次都要做）

**❌ 不要用滑鼠雙擊 HTML 檔來測試。**
那是 `file://` 模式，網頁讀 CSV／JSON 會被瀏覽器安全機制擋住，你會看到「壞掉」的假象。

**✅ 正確做法 —— 啟動本機預覽：**

1. 在 GitHub Desktop 選 **Repository → Open in Command Prompt**（或在資料夾按住 Shift + 右鍵 → 在此處開啟 PowerShell 視窗）
2. 貼上並執行：

```bash
python -m http.server 8080
```

3. 瀏覽器開 <http://localhost:8080/index.html>
4. **測試你剛換的那個檔案**：走一次使用者的操作路徑（選單 → Section → Line → 確認 PDF 打得開、頁數正確）
5. 測完在命令列視窗按 `Ctrl + C` 關閉

> 若電腦沒有 Python，到 <https://www.python.org/downloads/> 安裝，安裝時**勾選 Add Python to PATH**。

---

### 4.7 用 GitHub Desktop 上傳（Commit & Push）

| 步驟 | 動作 | 畫面上看到什麼 |
|---|---|---|
| 1 | 回到 GitHub Desktop | 左側 **Changes** 出現你改的檔案 |
| 2 | **檢查清單** | 只應該有你剛才動的檔案。有不認識的 → 停下來釐清，別急著上傳 |
| 3 | 確認每個檔案左邊的**勾選框都有勾** | 沒勾的不會上傳 |
| 4 | 左下 **Summary** 欄填寫這次改了什麼 | 見下方格式 |
| 5 | （選填）**Description** 補充細節 | 例如變更原因、圖面版次 |
| 6 | 按 **Commit to main** | Changes 清單清空 |
| 7 | 右上按 **Push origin** | 上傳到 GitHub |
| 8 | 等 1～3 分鐘，開啟正式網站確認 | <https://amtoppme.github.io/Maintenance-Smart/> |

> 步驟 8 開網站時請按 **`Ctrl + F5`**（強制重新整理）。
> 瀏覽器會快取舊的 PDF，直接重新整理常常還是看到舊版，會讓你誤以為沒更新成功。
> 若按了 `Ctrl + F5` 仍是舊版，再等 2 分鐘 —— GitHub Pages 發布需要時間。

**Summary 撰寫格式：**

```
<動作>: <對象> - <內容>
```

| 好的例子 | 不好的例子 |
|---|---|
| `Update: Line-4 F13 schematic - rev B 2026-08` | `Update index.html` |
| `Update: Bardac drive procedure - 新增參數表` | `updated file` |
| `Update: F12 BIT XE01 IO list` | `修改` |

> 目前 repo 裡有大量 `Update index.html` 這種訊息，導致無法從歷史紀錄看出改了什麼。
> **請從你接手開始改善這一點** —— 未來出問題時，這是你唯一的線索。

**如果 Push 失敗，看 §7.5。**

---

## 5. 備份與還原

### 5.1 Google Sheets 資料備份（每月，最重要）

> **這是唯一無法從 GitHub 還原的資料。網頁檔案再怎麼壞都救得回來，維修紀錄弄丟就沒了。**

**方式 A — 直接下載（建議，最簡單）**

1. 開啟維修紀錄 Google Sheets
2. 檔案 → 下載 → **Microsoft Excel (.xlsx)**
3. 檔名改為 `MaintenanceLogs_YYYYMM.xlsx`
4. 存到公司檔案伺服器：`\\<伺服器>\Maintenance-Smart-Backup\YYYY-MM\`
5. **同樣步驟再做一次 SCADA PC Inventory 的 Sheets**

**方式 B — 產生統計報表（需要月報時）**

1. 在 Records 頁面下載 `maintenance_logs.json`
2. 把它跟 `export_json.py` 放在同一個資料夾
3. 開啟命令列，執行：

```bash
pip install pandas openpyxl
```

```bash
python export_json.py
```

4. 產生 `Maintenance_Report_From_JSON.xlsx`，內含 7 個工作表：

| 工作表 | 內容 |
|---|---|
| `01_ByLine_Count` | 各產線維修次數 |
| `02_ByLine_DowntimeMin` | 各產線停機分鐘數 |
| `03_ByMonth_DowntimeMin` | 每月停機分鐘數 |
| `04_ByCategory` | 各類別次數／總停機／平均停機 |
| `05_TopEquipment_Count` | 維修次數前 20 名設備 |
| `06_TopEquipment_DowntimeMin` | 停機時間前 20 名設備 |
| `99_Raw_Logs` | 原始紀錄 |

**保留策略：** 月備份保留 24 個月；每年 12 月的備份永久保留。

### 5.2 程式碼備份

GitHub 上的版本本身就是備份。額外做離線備份（每季一次）：

```bash
cd "C:/Users/Franklin/Documents/GitHub2" && git clone --mirror https://github.com/AMTOPPME/Maintenance-Smart.git Maintenance-Smart-mirror
```

把產生的 `Maintenance-Smart-mirror` 資料夾壓縮後存到檔案伺服器。

### 5.3 容量監控（每月）

```bash
cd "C:/Users/Franklin/Documents/GitHub2/Maintenance-Smart" && du -sm .git . && git count-objects -vH
```

| `.git` 大小 | 狀態 | 動作 |
|---|---|---|
| < 500 MB | 正常 | 無 |
| 500 MB ~ 1 GB | 注意 | 開始規劃（見 §8.1） |
| **> 1 GB** | ⚠️ **警戒** | GitHub 會寄容量警告信 |
| > 5 GB | 🚨 危險 | Push 可能被拒絕，系統無法更新 |

**2026-08-17 實測：`.git` = 1,972 MB（已在警戒區），實際內容僅 275 MB。**
詳見 §8.1 —— 這是你接手後需要處理的第一優先項目。

### 5.4 還原程序

| 情境 | 怎麼做 |
|---|---|
| **檔案換錯了，還沒 Commit** | GitHub Desktop → Changes 清單 → 右鍵該檔 → **Discard changes** |
| **已經 Commit 但還沒 Push** | GitHub Desktop → History 分頁 → 右鍵該筆 → **Undo commit** |
| **已經 Push 出去了** | History → 右鍵該筆 → **Revert changes** → 再 Push |
| **本機資料夾整個不見** | GitHub Desktop → Clone repository → 選 `Maintenance-Smart` |
| **Sheets 資料誤刪（當下發現）** | Google Sheets → 檔案 → 版本記錄 → 查看版本記錄 → 還原 |
| **Sheets 資料誤刪（很久以後才發現）** | 用 §5.1 的月備份 .xlsx 重建 |
| **Apps Script 改壞** | 部署 → 管理部署作業 → 切回舊版本 |
| **GitHub 帳號進不去** | 用 §5.2 的 mirror 備份重建 |

> ⚠️ **每季實際演練一次還原**（隨便找一個檔案改壞，再還原回來）。
> 只有文件沒演練，等於沒有備份。

---

## 6. 定期健康檢查

### 6.1 PDF 連結檢查（每月）

檢查有沒有「網頁上有連結，但檔案不存在」的情況。

在 repo 資料夾開啟 **Git Bash**（GitHub Desktop → Repository → Open in Git Bash），貼上：

```bash
grep -oh -E "docs/(schematic|procedures)/[^\"']*\.pdf" *.html | grep -v '\$' | sort -u | while IFS= read -r f; do [ -f "$f" ] || echo "MISSING: $f"; done; echo "--- 檢查完成 ---"
```

**判讀：** 只出現「檢查完成」= 正常。出現 `MISSING:` = 該連結點下去會 404，要修。

### 6.2 IO 對照表檢查（每月／每次改 CSV 後）

```bash
cd "C:/Users/Franklin/Documents/GitHub2/Maintenance-Smart" && python -X utf8 -c "import json,os,glob;d=json.load(open('IODescriptionList/index.json',encoding='utf-8'));t=[];w=lambda o:[w(v) for v in o.values()] if isinstance(o,dict) else (t.append(o) if isinstance(o,str) and o.endswith('.csv') else None);w(d);m=[p for p in t if not os.path.isfile(p)];print(f'登錄 {len(t)} 筆 / 磁碟 {len(glob.glob(\"IODescriptionList/**/*.csv\",recursive=True))} 個 / 找不到 {len(m)} 筆');[print('  MISSING:',x) for x in m]"
```

> `-X utf8` 不可省略。Windows 主控台預設編碼（cp1252）無法輸出中文，
> 省略會直接噴 `UnicodeEncodeError` 而不是給你檢查結果。

**判讀：** 「找不到 0 筆」= 正常。有 MISSING = 使用者選到該項目會看到空白表格。
目前已知有 11 筆，詳見 §8.5。

### 6.3 後端連線檢查（Records 頁面出問題時用）

在瀏覽器按 `F12` 開啟開發者工具 → Console 分頁，貼上：

```javascript
fetch('https://script.google.com/macros/s/AKfycbx7CDsQJdIOmrgRFiFAyvZVcLvJrqr2dF1RM4oqy0Aj9FPiZDmTRUP75o8ZUp3bFHXvfg/exec').then(r=>r.json()).then(d=>console.log('OK, 筆數 =', d.data?d.data.length:d)).catch(e=>console.error('FAIL',e));
```

**判讀：** 出現 `OK, 筆數 = 123` = 後端正常。出現 `FAIL` = 後端有問題，看 §7.2。

---

## 7. 故障排除

### 7.1 網站打不開／整頁空白

網址：<https://amtoppme.github.io/Maintenance-Smart/>

| 順序 | 檢查 | 判斷 |
|---|---|---|
| 1 | 問其他同事是否也打不開 | 只有一人 → 請他按 `Ctrl + F5`，或換瀏覽器／換網路 |
| 2 | 開 <https://www.githubstatus.com/> | 紅色/黃色（尤其 Pages 項目）→ GitHub 故障，等待恢復，無法自行處理 |
| 3 | GitHub 網站 → repo → Settings → Pages，確認仍為 Deploy from branch `main` | 被關掉 → 重新啟用 |
| 4 | GitHub 網站 → repo → Actions 分頁，看最近一次 `pages build and deployment` | 紅色 ✗ → 發布失敗，點進去看錯誤 |
| 5 | 想想最近一次 Push 改了什麼 | 剛 Push 完就壞 → 用 §5.4 的 **Revert changes** |
| 6 | 公司網路是否擋了 `github.io` | 手機開行動網路測試，能開 → 是公司防火牆，找 IT |

### 7.2 Records / Report 讀不到或存不進資料

| 順序 | 檢查 |
|---|---|
| 1 | 按 F12 看 Console 有沒有紅色錯誤訊息 |
| 2 | 執行 §6.3 的連線檢查 |
| 3 | 開 Google Apps Script → 左側「執行項目」，看有沒有失敗紀錄 |
| 4 | 部署 → 管理部署作業，確認「誰可以存取」是**任何人** |
| 5 | 確認 Google 帳號沒有超過 Apps Script 每日配額 |
| 6 | 確認網頁裡的網址與目前部署的網址一致（§4.5 的表格） |

> 🚨 **使用者端急救：** 如果使用者說「我填了一堆但存不進去」，
> **千萬不要叫他清除瀏覽器資料或重開機**。資料可能還在 localStorage 裡。
> 請他按 F12 → Console，貼上這行，然後把結果貼給你：
> ```javascript
> copy(localStorage.getItem('ms-maintenance-logs'))
> ```

### 7.3 PDF 打不開

| 症狀 | 原因 | 處理 |
|---|---|---|
| 404 / 空白頁 | 檔名不符（最常見） | 執行 §6.1；99% 是上次更新沒照 §4「檔名一模一樣」的規則 |
| 轉圈很久才開 | 檔案太大（有些 20 MB） | 見 §8.2 |
| 手機上顯示異常 | pdfjs 相容性 | 改用瀏覽器原生 PDF 檢視器測試 |

### 7.4 IO Description 表格空白

1. 執行 §6.2，確認該項目不在「找不到」清單中
2. 確認是用 `http://localhost:8080` 開的，不是雙擊檔案（§4.6）
3. 確認 CSV 編碼是 UTF-8（中文變亂碼就是這個問題，§4.3）

### 7.5 GitHub Desktop 上傳失敗

| 訊息（大意） | 原因 | 處理 |
|---|---|---|
| 檔案超過 100 MB | 單一檔案太大 | 壓縮 PDF，或改用 Git LFS（§8.1） |
| repository 超過容量 | repo 太肥 | 見 §8.1 |
| 需要先 Pull / 有衝突 | 遠端有別人的更新 | 按 **Pull origin** 後再 Push |
| 認證失敗 | 密碼或 Token 過期 | GitHub Desktop → File → Options → Accounts → 重新登入 |

### 7.6 GitHub Desktop 出現「Conflict（衝突）」

發生在你和別人改到同一個檔案時。

1. **不要慌，也不要按任何看不懂的按鈕**
2. 如果衝突的是 PDF/CSV 這類檔案：GitHub Desktop 會問你要留哪一個 → 選你剛換的那個
3. 如果衝突的是 `.html` 或 `.json`：建議找人協助，或先 **Discard changes** 放棄自己的修改，重新 Pull 後再改一次
4. 預防方法：**每次開工前先按 Pull origin**（§3.3）

---

## 8. 已知問題與待處理事項

> 以下是 2026-08-17 實際掃描本機 repo 得到的結果，不是假設。
> 接手後請依優先順序處理。

### 8.1 🔴 Git 歷史檔案肥大（第一優先）

**現況：** `.git` 資料夾 1,972 MB，但網站實際內容只有 275 MB。約 **86% 是舊版 PDF 的歷史殘留**。

**成因：** 每次更新 schematic 都是整檔取代（單檔 13～21 MB），而 Git 會**永久保留每一個舊版本**。
這是「檔名一模一樣、直接覆蓋」這個規則的必然副作用 —— 規則本身是對的（不然連結會斷），但代價是歷史會不斷長大。

**影響：** 已超過 GitHub 建議上限 1 GB。繼續下去：clone 極慢、新人接手困難、最終可能無法 Push。

**可行方案（需與主管討論後執行，不要自行決定）：**

| 方案 | 做法 | 優點 | 缺點 |
|---|---|---|---|
| A. Git LFS | 把 `*.pdf` 移到 Large File Storage | 保留現有使用流程 | 需改寫歷史；GitHub 免費額度 1 GB/月流量可能不夠 |
| B. 大檔外移 | Schematic 改放檔案伺服器/SharePoint，網頁改成連結過去 | 最徹底解決 | 外網或行動裝置存取方式會改變 |
| C. 清理歷史 | 用 `git filter-repo` 刪掉舊版 PDF | repo 立刻瘦身 | 所有人要重新 clone；**執行前務必先做 §5.2 mirror 備份** |
| D. 壓縮 PDF | 降低解析度後再上傳 | 簡單 | 治標不治本；**電路圖壓縮後細節可能看不清，不建議** |

**在方案定案前的緩解措施：**
更新 schematic 前先確認「這份圖真的有改」，避免內容相同卻重複上傳，白白增加 1 個 20 MB 的歷史版本。

### 8.2 🟡 大型 PDF 影響載入速度

以下 4 個檔案介於 16～21 MB，在無線網路或行動裝置上開啟會明顯卡頓：

- `Line-2-F13-Schematic.pdf`（20.0 MB）
- `Line-1-F13-Schematic.pdf`（19.9 MB）
- `Line-3-F13-Schematic.pdf`（19.9 MB）
- `Line-2-F12-Schematic.pdf` / `Line-3-F12` / `Line-4-F12`（各 16.3 MB）

**建議：** 評估拆成分頁 PDF，或提供「低解析度線上預覽 + 高解析度下載」兩種版本。

### 8.3 🔴 後端資料無存取控制

兩組 Google Apps Script 網址是**明文寫在網頁裡**的，而且部署成「任何人皆可存取」。
任何拿到網址的人，都能直接新增、修改、**刪除**維修紀錄，沒有任何驗證，也查不出是誰做的。

**建議處理（由淺入深）：**

| 階段 | 做法 | 成本 |
|---|---|---|
| 立即 | 確實執行 §5.1 月備份，並確認 Sheets 版本記錄可用 —— 這是目前**唯一的防線** | 低 |
| 短期 | Apps Script 加入共用密鑰檢查 | 中（可擋隨機掃描，但密鑰仍在前端可見） |
| 中期 | Apps Script 增加「操作日誌」工作表，記錄每次寫入的時間與內容 | 中（至少能追溯） |
| 長期 | 改為需 Google 帳號登入（限公司網域）的部署方式 | 高（使用者要登入） |

### 8.4 ✅ 已存在的斷鏈（1 筆）—— 已於 2026-08-17 修正

`CalibrationProcedure.html` 第 386 行原本寫的是：
```
docs/procedures/Bardac Drive Download Procedure_SimpleVersion.pdf
```
但實際檔名是：
```
docs/procedures/Bardac Drive Download_Procedure_SimpleVersion.pdf
                                    ↑ 這裡是底線，不是空白
```

**處理方式：** 已改 HTML 裡的字串使其符合實際檔名（不改檔名，以免既有書籤失效）。
修正後執行 §6.1 檢查，PDF 連結為零錯誤。

> 這是最典型的「檔名沒有一模一樣」案例，留在文件裡當作範例。

### 8.5 🟡 IO 對照表有 11 筆無效登錄

`IODescriptionList/index.json` 登錄了 102 筆 CSV，但磁碟上只有 101 個檔案，其中 11 筆對不上。
使用者選到這些項目會看到**空白表格**。

已逐筆比對磁碟實況，**其中 10 筆是 `index.json` 路徑打錯，檔案其實都在**，只要改字串就好：

| # | index.json 目前寫的（錯） | 磁碟上實際的（對） |
|---|---|---|
| 1–5 | `F12/WORD1BIT1 - missing pg 6/F12 - 8 - WORD1BIT1 - …` | `F12/WORDBIT1 - missing pg 6/F12 - 8 - WORDBIT1 - …`　（多打一個 `1`，共 ORIGINAL + XE01~XE04 五筆） |
| 6 | `F13 - 8 - WORDBIT - 4750-4808.csv` | `F13 - 8 - WORDBIT - 4750 - 4808.csv`　（數字間少了空格） |
| 7 | `F13 - 12 - TEMP CONTROL - DATA.csv` | `F13 - 12 - TEMP CONTROL - DATA WORD.csv` |
| 8 | `F13 - 12 - TEMP CONTROL - TC 50.csv` | `F13 - 12 - TEMP CONTROL - TC 50 WORDS.csv` |
| 9 | `F13 - 12 - TEMP CONTROL - TEMP RANGE.csv` | `F13 - 12 - TEMP CONTROL - TEMP REGULATOR 50 WORD.csv`　⚠️ 名稱差異較大，**請先確認是否為同一份資料再改** |
| 10 | `F201 - 4 - CW - CW assign for winder.csv` | `F201 - 4 - CW - CW assign for winder W.csv`　（結尾少一個 `W`） |

**剩下 1 筆是檔案真的不存在：**

| # | 項目 | 說明 |
|---|---|---|
| 11 | `F16/GoOnline/Master Form/… - TIMER.csv` | `Master Form` 資料夾裡沒有 TIMER，但隔壁 `Master Form Updated` 有。<br>需判斷：改指向 Updated 版，或從 `index.json` 移除這筆 |

**另外發現 3 個檔案存在磁碟上但沒有登錄**，使用者在網頁上選不到：
- `F13/TEMP CONTROL/F13 - 12 - TEMP CONTROL - TC 16 BIT.csv`
- `F201/CW/F201 - 4 - CW - cdw.csv`
- `F201/CW/F201 - 4 - CW - cw.csv`

**處理原則：**
1. 第 1~8、10 項可直接修正 `index.json` 路徑字串
2. 第 9、11 項需先與熟悉該產線的工程師確認內容是否對應
3. 未登錄的 3 個檔案，確認是否應該讓使用者看到，是的話補進 `index.json`
4. 每改一次都執行 §6.2 驗證，目標是「找不到 0 筆」

### 8.6 🟡 帳號安全

2026-08-17 曾發生 GitHub 帳號 `AMTOPPME` 密碼以明文外流的事件。

**必辦清單：**
- [ ] 更換 GitHub 密碼
- [ ] 啟用 2FA 兩步驟驗證
- [ ] Settings → Sessions，登出所有不明裝置
- [ ] Settings → Developer settings → Personal access tokens，撤銷不明 token
- [ ] 檢視 repo 的 Security → Audit log，確認沒有異常 Push

**日後守則：密碼、Token、API Key 一律不寫進 repo、不貼在對話或郵件中。**

### 8.7 🔴 網站與所有廠內文件公開在網際網路上

**現況：** 網站託管於 GitHub Pages 免費方案，該方案**只支援公開（public）repository**。
因此以下內容目前任何人都能存取，不需登入、不限公司網路：

- 全部 13 份電路圖 PDF（`docs/schematic/`）
- 全部 40 份設備程序書 PDF（`docs/procedures/`）
- 全部 IO 對照 CSV 與 PDF（`IODescriptionList/`）
- 網頁原始碼（含 §1.4 的兩組後端網址）
- **完整的 Git 歷史 —— 包含所有曾經上傳過、後來被覆蓋掉的舊版檔案**

最後一項要特別注意：即使日後把某份文件刪掉，它仍然留在 Git 歷史中，可被下載。

**需要確認的問題（請與主管／IT 確認，不要自行判斷）：**
- 電路圖與設備程序書是否屬於公司機密或營業秘密？
- 是否受 ISO 或客戶稽核的文件管制要求？
- 是否包含供應商的受著作權保護資料（例如 Bardac、Dynisco、Measurex 的原廠手冊）？

**若確認不應公開，處理選項：**

| 方案 | 做法 | 代價 |
|---|---|---|
| A. 改為私有 repo + GitHub Pages | 需 GitHub Enterprise 方案 | 需付費 |
| B. 移到公司內網 | 檔案放內部檔案伺服器或 IIS，網站改為內網存取 | 需 IT 資源；外出時無法存取 |
| C. 移到 SharePoint / OneDrive | 用公司既有的 Microsoft 365 權限控管 | 需重做前端連結 |
| D. 分級處理 | 敏感文件移到內網，非敏感的留在現址 | 需逐份文件判定 |

> 這一項與 §8.1（repo 過大）可以合併考慮 —— 方案 B / C 同時解決兩個問題。

### 8.8 ✅ 缺少專案說明文件 —— 已於 2026-08-17 補上

已新增 `README.md`，內容包含：專案用途、功能模組、系統架構、目錄結構、
更新 SOP 摘要、健康檢查指令、統計報表產生方式，並連結至本手冊。

GitHub 會自動把 `README.md` 顯示在 repo 首頁，是新接手者第一眼會看到的東西。
**日後目錄結構或更新流程有變動時，記得同步修改 `README.md`。**

### 8.9 🟢 Commit 訊息品質

歷史紀錄中充斥 `Update index.html`、`updated file` 等無意義訊息，出問題時無法回溯。
請從接手起改用 §4.7 的格式。

---

## 9. 附錄

### 9.1 資料夾結構

```
Maintenance-Smart/
├── index.html                    ← 外層框架（選單、主題），82 KB
├── Home.html                     首頁 / Quality Policy
├── about.html                    關於（彈出視窗）
├── Updated Schematic.html        電路圖檢視器
├── IO_Description_List.html      IO 描述查詢
├── IO_List.html                  IO 清單
├── CalibrationProcedure.html     Procedure PDF 目錄
├── Training.html                 教育訓練
├── SCADA_PC_Inventory.html       SCADA PC 資產
├── Maintenance_Records.html      維修紀錄查詢 / 編輯
├── New_Report.html               新增維修報告
├── export_json.py                月報產生工具（§5.1 方式 B）
│
├── assets/                       背景圖與圖示（12 MB）
├── pdfjs/                        內建 PDF 檢視器程式（20 MB，不要動）
│
├── docs/
│   ├── procedures/               ← 【常更新】40 個設備程序書 PDF（80 MB）
│   ├── schematic/                ← 【常更新】13 個電路圖 PDF（150 MB）
│   └── MAINTENANCE_PROCEDURE.md  ← 本文件
│
└── IODescriptionList/
    ├── index.json                ← 【改 CSV 時要一起改】路徑對照表
    ├── F12/ ├─ BIT/ INPUTBIT/ INPUTWD/ OUTPUT/ TIMER/ TXT/ WORD1~4/ WORDBIT1/
    ├── F13/ ├─ BIT/ INPUTBIT/ INPUTWD/ MODULELIST/ OUTPUTBIT/ OUTPUTWD/
    │        └─ TEMP CONTROL/ TIMER/ TXT/ WORD/ WORD6/ WORDBIT/
    ├── F16/ ├─ BIT/ GoOnline/ INPUT/ INPUTOUTPUT/ OUTPUT/
    └── F201/└─ BIT/ CW/ INPUTOUTPUT/ NEWWINDER/ PLC-LIST/ STRUCTURE PROGRAMME/
              TIMER/ WORD/
```

### 9.2 命名規則速查

| 類型 | 規則 | 範例 |
|---|---|---|
| 電路圖 | `Line-<1~4>-<F12\|F13\|F201>-Schematic.pdf` | `Line-4-F13-Schematic.pdf` |
| 電路圖（F16） | `Line-<n>-F16-Esixting Updated Schematic.pdf` | `Esixting` 是原始拼字，**勿更正** |
| IO 描述 PDF | `<Section>_Line<n>_IODescription.pdf` | `F13_Line2_IODescription.pdf` |
| IO CSV | `<Section> - <編號> - <類別> - <版本>.csv` | `F12 - 11 - BIT - XE01.csv` |
| Procedure PDF | 無固定規則（沿用原始文件名） | — |

### 9.3 GitHub Desktop 日常操作速查

| 我想做什麼 | 怎麼做 |
|---|---|
| 開工前同步 | 右上 **Fetch origin** → 若變成 **Pull origin** 就按下去 |
| 打開資料夾 | Repository → **Show in Explorer** |
| 開啟命令列 | Repository → **Open in Command Prompt / Git Bash** |
| 看我改了什麼 | 左側 **Changes** 分頁 |
| 放棄我的修改 | Changes → 右鍵檔案 → **Discard changes** |
| 上傳 | 填 Summary → **Commit to main** → **Push origin** |
| 看歷史紀錄 | 左側 **History** 分頁 |
| 撤銷已上傳的變更 | History → 右鍵該筆 → **Revert changes** → Push |

### 9.4 待確認事項（請與前任逐項確認並填寫）

- [x] ~~正式網站網址~~ → <https://amtoppme.github.io/Maintenance-Smart/>（GitHub Pages，已確認）
- [ ] Google Apps Script / Sheets 擁有者帳號：`_________________`
- [ ] 維修紀錄 Sheets 連結：`_________________`
- [ ] SCADA PC Inventory Sheets 連結：`_________________`
- [ ] 系統使用者範圍（哪些部門、約幾人）：`_________________`
- [ ] 資料保存年限要求（是否受 ISO 或客戶稽核規範）：`_________________`
- [ ] 維護代理人：`_________________`
- [ ] 電路圖 / Procedure 的原始檔（非 PDF）存放位置：`_________________`
- [ ] 誰有權核准電路圖更新：`_________________`

### 9.5 修訂記錄

| 版本 | 日期 | 修訂者 | 內容 |
|---|---|---|---|
| v1.0 | 2026-08-17 | — | 初版建立 |
| v1.1 | 2026-08-17 | — | 改寫為交接手冊；操作流程改以 GitHub Desktop 為主；明確化「檔名一模一樣、直接覆蓋」規則 |
| v1.2 | 2026-08-17 | — | 補上正式網站網址與 GitHub Pages 發布流程；新增 §8.7 公開存取風險 |
| v1.3 | 2026-08-17 | — | §8.4 斷鏈已修正；§8.5 補上 11 筆無效登錄的逐筆對應與 3 筆未登錄檔案 |
| v1.4 | 2026-08-17 | — | §8.8 已新增 `README.md` |

---

## 10. 求助管道

| 問題類型 | 找誰 / 去哪 |
|---|---|
| 電路圖內容正確性 | 現場電控工程師 |
| Procedure 內容正確性 | 該設備負責工程師 |
| GitHub Desktop 操作 | <https://docs.github.com/desktop> |
| GitHub 服務狀態 | <https://www.githubstatus.com/> |
| Google Apps Script | <https://developers.google.com/apps-script> |
| 前任維護人員 | `_________________`（請填寫聯絡方式與可諮詢期限） |

---

**文件結束。有任何步驟做不通，請不要「試試看」——先看 §7，或依 §10 求助。**
**這套系統壞掉的成本，遠低於資料弄丟的成本。**
