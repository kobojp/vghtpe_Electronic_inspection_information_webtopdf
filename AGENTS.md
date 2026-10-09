# AGENTS.md — AI 開發與操作規則

本文件供在此儲存庫工作的 AI Agent 使用。先讀本文件，再依使用者當次需求確認範圍。與使用者溝通、工作說明及專案文件使用繁體中文。使用者當次明確指示優先；不要把本文件中的歷史規則當成再次詢問已授權工作的理由。

## 1. 專案定位與已採用的架構

這是「臺北榮總水電消防報表下載系統」，目標交付 Windows 桌面 EXE。使用者已同意：

- UI 使用 React + TypeScript + Vite。
- Python / FastAPI 提供只監聽本機的 HTTP API，處理資料維護、報表下載及 PDF 擷取。
- pywebview / Microsoft WebView2 提供 Windows 視窗、原生對話框與檔案拖曳。
- 介面採左側導覽、中央工作區、底部共用任務面板。
- 四個頁面依序為報表下載、PDF 擷取、報表管理、設定。
- 外部 `data.json` 由使用者維護，與 EXE 同層，不嵌入 EXE。

`main.py` 現在只是桌面啟動入口，舊版超過 2200 行的 Tkinter 介面已被拆分。不要重新加入 Tkinter、PyQt，或依舊文件要求建立 `gui.py`。目前不使用資料庫、Docker、Rust 或雲端部署；根目錄殘留的 Cargo 檔案不代表目前程式以 Rust 建置。沒有明確需求時不要導入這些技術、登入系統、OCR、自動更新或其他額外功能。

原則：架構清楚、複雜度低、每一步能執行。先確認實際程式，再修改相關模組，不順手重構其他功能。

## 2. 開工前與工作區保護

1. 檢查 `git status --short`，辨識既有修改、未追蹤檔案與本機產出；不要還原、刪除或提交無關內容。
2. 讀取本次相關模組及測試。優先使用 `rg` / `rg --files` 搜尋，不假設文件中的歷史行數、版本或測試數仍正確。
3. 用簡短步驟說明預期結果；已授權的實作、檢查與修正直接完成。
4. 有資訊不足時先進行不依賴答案的工作，再提出必要的釐清；普通、可逆的修改不需反覆要求確認。
5. Windows 檔案操作使用原生 PowerShell cmdlet 與 `-LiteralPath`。遞迴刪除前確認絕對路徑位於本次預期目錄；不要跨 shell 拼接刪除命令。
6. 建置、測試背景程序用隱藏視窗。記錄自己建立的程序並在測試結束後停止；不得終止使用者的不明程序。
7. 不把 API 權杖、測試秘密或帳密寫入文件、提交內容及工作輸出。

## 3. 檔案與責任分界

```text
main.py                         # 呼叫 desktop.launcher.run
backend/
  app.py                        # API、工作階段保護、請求驗證及組合核心
  catalog.py                    # data.json CRUD、版本比對、備份、原子保存
  reports.py                    # 七種類型、月份範圍、網址、wkhtmltopdf 下載
  pdf_core.py                   # PyMuPDF 文字搜尋、頁面擷取、輸出名稱
  tasks.py                      # 單一任務、執行緒、取消、進度與紀錄
  config.py                     # 外部／內嵌路徑、設定、日誌與開啟資料夾
desktop/
  launcher.py                   # loopback 服務、WebView2、選檔與程式關閉
  drop.py                       # 原生拖曳路徑、PDF 驗證、React 通知
frontend/
  src/App.tsx                   # 四頁導覽、共用資料及操作狀態
  src/pages/                    # Downloads、PdfExtract、ReportManager、Settings
  src/components/               # 共用圖示與任務面板
  src/api/client.ts             # API 呼叫及 X-App-Token
  src/hooks/useTask.ts          # 取消可清理的每秒任務查詢
  src/types/                    # API 與前端型別
  src/styles.css                # 共用樣式、CSS 變數、拖曳／響應式提示
  e2e/                          # Playwright、隔離服務生命週期及案例
tests/                          # unittest 與測試用 FastAPI 服務
docs/                           # 架構、驗證範圍及發布說明
media/2026-10-09.png             # README 指定介面截圖
data.json                       # 外部預設報表目錄
Pipfile / Pipfile.lock           # Python 宣告及鎖定依賴
frontend/package*.json          # 前端宣告及鎖定依賴
build.spec                      # 主要 PyInstaller 打包設定
main.spec                       # 相容入口，委派 build.spec
.github/workflows/ci.yml         # PR / main 驗證
.github/workflows/release.yml    # v* 標籤建置與 GitHub Release
```

- React 不直接讀寫本機 JSON、啟動程序或處理 PDF；這些操作走 API 或桌面原生事件。
- `backend/` 核心不依賴 UI，不由背景工作直接更新 React 或操作視窗。
- 原生視窗整合只放在 `desktop/`；不要把 pywebview 平台細節塞進下載或 PDF 核心。
- API、TypeScript 型別與呼叫端一起修改；不要用 `any` 隱藏契約錯誤。
- 保留跨頁表單狀態及持續存在的共用任務面板。

## 4. 執行環境與依賴

- Python 版本以 `Pipfile` / CI 為準，目前是 3.13.3；前端開發及 CI 使用 Node.js 24。
- Python 用 Pipenv，前端用 npm。正常安裝依鎖檔執行 `pipenv sync --dev` 與 `npm.cmd ci --prefix frontend`，不在 CI 重新 resolve 依賴。
- 設定 `PIPENV_VENV_IN_PROJECT=1`，使 Python 環境位於專案 `.venv/`。Playwright 的 setup 直接啟動此環境的 Python，不能只建立全域 Pipenv 環境。
- 只有需要改依賴時才更新 manifest 與鎖檔；不要因一般 UI 修改順便升級套件。
- 鎖檔來源保持正式公開套件來源。下載限制下使用的本機代理、臨時快取不寫進正式鎖檔。
- 若 shell 受到政策限制，使用可用工具執行等價命令並說明限制；不得變更執行政策或要求提權來繞過限制。

首次準備及執行：

```powershell
$env:PIPENV_VENV_IN_PROJECT = "1"
python -m pip install pipenv
pipenv verify
pipenv sync --dev
npm.cmd ci --prefix frontend
npm.cmd --prefix frontend run build
pipenv run python main.py
```

即時開發使用兩個終端：

```powershell
# 終端一
npm.cmd --prefix frontend run dev

# 終端二
pipenv run python main.py --dev
```

預設桌面模式使用隨機本機 API port；`--dev` 使用 API 127.0.0.1:8765 與 Vite 127.0.0.1:5173。不要將 API 改成監聽 0.0.0.0。瀏覽器預覽沒有原生對話框或完整檔案拖曳路徑；桌面功能須以 EXE / pywebview 執行。

## 5. 路徑、持久化與日誌

區分三種位置，不能混用：

| 資料 | 原始碼執行 | EXE 執行 |
| --- | --- | --- |
| 可變報表目錄 `data.json` | 專案根目錄 | `sys.executable` 同層 |
| 內嵌前端、圖示、wkhtmltopdf | 根目錄／frontend/dist | PyInstaller 的 `sys._MEIPASS` 資源目錄 |
| 個人設定及日誌 | APPDATA 狀態目錄 | 同左 |

- `config.application_dir()` 是外部可變資料位置；`config.resource_dir()` 是內嵌資源位置。
- 設定為 `%APPDATA%/VghtpeReportDownloader/settings.json`；日誌為同目錄的 `logs/error.log`，使用 UTF-8 滾動日誌。
- `VGHTPE_STATE_DIR` 只改設定／日誌位置，**不會改變 data.json 的位置**。測試報表目錄須顯式傳入 `create_app(data_path=...)`。
- 預設下載資料夾為 EXE／專案同層 `水電消防報表`；擷取資料夾為 `報表合併pdf`。使用者可改用完整絕對路徑。
- 保存 JSON 或 PDF 先寫同資料夾暫存檔，再以 `os.replace` 替換；失敗要清理暫存檔並保留原資料。
- 記錄例外 traceback，對使用者顯示繁體中文訊息。攔截 `PermissionError`，提示關閉正在使用的 PDF 或確認資料夾可寫入；不能吞掉例外並回報成功。

## 6. data.json 與報表管理：必須維持的契約

### 資料結構

最上層為 JSON object。七個報表清單 key 必須保留且為 array：

| UI 類型 | JSON key | 輸出大類 |
| --- | --- | --- |
| 消防 | Fire_Equipment | 消防 |
| 電力每日 | electricity_every_day | 電力 |
| 電力每月 | electricity_every_month | 電力 |
| 電力每周 | electricity_every_week | 電力 |
| 排水每日 | drain_day | 排水 |
| 排水每月 | drain_month | 排水 |
| 排水每周 | drain_week | 排水 |

每筆至少含 `name`、`api_1`、`api_2`。保留 `squadName`、`Building_name`、所有其他根欄位，以及修改單筆資料時的額外欄位。不要根據前端已知欄位重建整份檔案。

### 保存與衝突

- `CatalogStore` 負責讀取、鎖定、SHA-256 revision、CRUD 及原子保存。
- `data.json.bak` 保留每次成功替換前的上一版完整內容，只保留一版；不列入發布 ZIP 或 Git。
- 備份失敗不可替換正式檔案；正式檔案替換失敗不可更新畫面為成功。
- 管理 API 的新增／修改／刪除都帶 `revision`。外部手動修改或其他畫面更新後，舊 revision 回 409 並提示重載，不能默默覆蓋。
- 現有 report id 為 `JSON key:index`，不是永久 UUID。刪除或移動類型會影響 index；不得用過期 id 直接操作新資料。
- 前端下載請求須帶 `catalog_revision`；保存或重載後同步清單並清除舊勾選，保留月份、篩選與搜尋。
- 下載任務持有啟動時的報表／設定快照。允許維護目錄，但不可讓正在執行的任務改用新的報表資料。
- 刪除報表只刪除目錄項目，不刪除已下載 PDF。UI 先顯示明確名稱，再由使用者確認刪除。
- 外部資料缺失／損壞需明確報錯，不能自動用內嵌預設資料覆蓋使用者資料。
- 更新 EXE 時保留現有 `data.json`；ZIP 的資料是首次使用預設目錄，不是升級時強制替換的內容。

### 輸入驗證

- 名稱為 1–180 字，前後無空白，禁止 Windows 非法字元、控制字元、尾端句點與保留裝置名。
- 同一輸出大類名稱不重複，含不同電力／排水頻率，避免輸出至相同月份時覆蓋；比較忽略大小寫。
- `api_1` 為 `/數字/`；`api_2` 為 `/數字/數字`，目前相容舊資料省略第一個斜線的格式。沒有實際網站驗證時，不順手更改舊參數。
- 第一版只管理既有七種類型；新增自訂類型須另行確認下載端點與資料結構。

## 7. 下載、PDF 與任務規則

### 報表下載

- 消防網址使用 `Report6{api_1}{YYYY-MM}{api_2}`；電力／排水使用 `Report6BatchAll{api_1}{月初}/{月底}{api_2}`，主機維持現有報表網站。
- 使用實際月份最後一天，含閏年；跨年度範圍包含起訖月。上個月用月初減一天，不以 31 日硬套另一月份。
- 「選取目前清單」只作用於篩選後可見報表。下載全部須全部類別、清空搜尋後選取；不要把既有按鈕改成無條件全目錄下載。
- 已存在的有效、可讀 PDF 跳過；零位元組、殘檔或非 PDF 要重新下載。
- wkhtmltopdf 以隱藏子程序執行，暫存輸出通過驗證後才換成正式檔。
- 有重試與逾時；取消時終止並回收子程序。不可在未確認子程序／worker 結束前把狀態重設成可開始。

### PDF 擷取與拖曳

- PyMuPDF 只搜尋有文字層的 PDF；不宣稱支援掃描 OCR。
- 關鍵字允許間隔空白／換行、不分大小寫。符合頁與選用簽核頁維持來源順序，不重複插入。
- 沒有符合關鍵字時不可只輸出承商簽核頁。
- 每份來源各輸出一份結果；不同月份的同名 PDF 使用月份或遞增編號，不覆蓋。
- 拖曳完整路徑來自 pywebview DOM drop 的 `pywebviewFullPath`。不得把瀏覽器 `File.name` 當成完整路徑。
- `desktop/drop.py` 在 React 就緒後訂閱拖曳區域，驗證實際 PDF 後送出 `vghtpe:pdf-drop` CustomEvent。
- 前端合併拖入及選檔清單，忽略 Windows 路徑大小寫與斜線差異；不同資料夾同名檔案仍是不同來源。
- 隱藏頁面／任務執行中拒絕加入；阻止拖放檔案造成視窗導覽；保留原生選檔按鈕。

### 任務與 UI

- `TaskManager` 同時僅允許一個下載／擷取 worker；每次任務獨立 `threading.Event`，不得共用或提前清除取消旗標。
- 狀態：`running → completed / failed`；取消：`running → cancelling → cancelled`。
- 使用 lock 保護狀態及 snapshot；回傳紀錄副本，不把可變內部物件直接交給 UI。
- 成功／跳過／失敗分開計數；已取消為未完成項目。100% 是處理結束，不保證全部成功。
- 前端每秒查詢，避免重疊請求；卸載或停止時清理 AbortController / timer。
- 關閉程式取消工作、關閉 uvicorn 與 socket。任務及最近紀錄只在記憶體，不自行加入續傳或永久歷史。
- 共用視覺以 `styles.css` 變數集中管理。新增控制項須有明確標籤、鍵盤焦點及 disabled 狀態。

## 8. API 安全與跨層一致性

- 服務僅限 loopback，保留 hostname / Origin 檢查。
- `/api/health` 與 `/api/bootstrap` 是啟動入口；其他 API 驗證每次啟動產生的 `X-App-Token`。
- 不啟用對外開放的 CORS、不拿掉工作階段保護，也不暴露任意 shell 執行接口。
- 開啟資料夾 API 只接受 download / extract / task 種類，不改成接受任意使用者命令。
- 桌面原生操作不可用時提供明確訊息，不捏造選檔結果。
- `tests/ui_server.py` 的 `/__test__/` 端點及 `VGHTPE_UI_TEST_TOKEN` 只供隔離測試，不複製到正式 API。

## 9. 驗證流程與證據

依修改範圍選擇必要檢查，通過後不無故重複整套測試：

| 修改 | 必要驗證 |
| --- | --- |
| 一般文字、圖片 | Markdown 路徑／圖片存在、文件內容與實作一致、git diff --check |
| backend / desktop | 相關 unittest、跨層 API 案例，必要時完整 Python suite |
| React / API 契約 | Vitest、TypeScript/Vite build、受影響的 Playwright 流程 |
| 拖曳／對話框 | 邏輯測試及前端通知流程，另標記原生 Windows 實測結果 |
| build.spec / 依賴 | 鎖檔一致、PyInstaller、打包 EXE headless smoke、ZIP 內容 |
| workflows | YAML / Actions 靜態檢查、PowerShell 語法及相關可在本機執行的步驟 |

常用指令：

```powershell
pipenv verify
pipenv run python -m unittest discover -s tests -v
npm.cmd --prefix frontend test
npm.cmd --prefix frontend run build
npm.cmd --prefix frontend run test:e2e
pipenv run python main.py --headless --smoke-test
pipenv run python main.py --smoke-test
git diff --check
```

- Vitest 僅執行 `src/**/*.test.{ts,tsx}`；Playwright 執行 `e2e/`。不要讓兩套 runner 混跑。
- Playwright 從 `.venv` 啟動 `tests/ui_server.py`，使用 127.0.0.1:18865。CI 安裝 Playwright Chromium；本機可使用 Chrome 或 `VGHTPE_TEST_BROWSER`。
- 所有寫入 data.json／PDF／設定的測試使用臨時資料夾；不要測試修改專案原始 data.json，也不連醫院網站進行大量下載。
- 下載轉換及原生對話框可隔離替代；紀錄哪些實際執行、哪些模擬。修 bug 先建立可重現案例，再驗證修正。
- 當前通過數與限制見 `docs/verification.md`，不是永久固定數字。
- headless smoke 驗證 API、目錄與前端包裝；不代表 WebView2 視窗、原生對話框、檔案總管拖曳已實測。
- 截至 2026-10-09，本機自動化環境原生 WebView2 未完成載入。沒有新證據不得宣稱此限制已解決；一般 Windows 桌面實測與實際醫院報表下載另行記錄。
- GitHub Actions 尚未執行時，只能說完成本機／靜態檢查，不能宣稱遠端 CI 已通過。
- 文件變更不要求新增程式或測試；涉及功能時交付可執行結果。回覆說明改動、檔案位置、驗證及實際限制。

## 10. 打包與 workflows

```powershell
npm.cmd --prefix frontend run build
pipenv run pyinstaller build.spec --clean --noconfirm
```

- 先建前端，再打包 EXE。不要直接打包遺留根目錄 EXE，也不要提交 `frontend/dist`。
- `build.spec` 包含編譯後前端、`app.ico`、根目錄 `wkhtmltopdf.exe` 與必要 WebView2 / uvicorn 模組；`data.json` 留在外部。
- 交付 ZIP 只含最新 `水電消防報表下載系統.exe` 與外部預設 `data.json`；排除個人設定、日誌、備份、報表 PDF 與測試產物。
- 打包後從 EXE 同層外部資料啟動 `--headless --smoke-test`，確實等待 Windows GUI EXE 的程序結束並檢查 exit code；不要直接呼叫 EXE 後就假設成功。
- `ci.yml`：main / PR 做依賴安裝、Python / 前端 / 瀏覽器測試、建置、EXE 資源啟動驗證；失敗保留診斷產物。
- `release.yml`：v* 標籤經格式檢查後做相同驗證、ZIP 檢查、發布 GitHub Release。不得以跳過測試解除發布失敗。
- workflow 使用 Windows runner。保留 `PIPENV_VENV_IN_PROJECT=1`；PowerShell 多個外部命令逐一檢查 `$LASTEXITCODE`。
- 不在 workflows 自動執行依賴一般桌面互動的原生視窗／拖曳測試；headless 與瀏覽器測試不能當成原生驗證。
- `ci.yml` 與 `release.yml` 的共同驗證步驟一起檢查；目前規模不額外建立大型 reusable-workflow 抽象。
- `media/2026-10-09.png` 是指定 README 圖片。沒有要求不要重做截圖、修改影像，或改回舊版介面圖片。

## 11. GitHub 更新、版號與純淨提交

只在使用者要求上傳／更新 GitHub／發布時執行 commit、push、tag、Release；普通開發、文件修改及流程檢查不自動發布。

### 程式變更發布：強制遞增版本

1. 讀取遠端現有版本標籤，不只看本機；使用已授權的 Git 或連接器。無法確認遠端時先完成本機準備並報告，不猜版號。
2. 正常發布以最新 `vMAJOR.MINOR.PATCH` 的 Patch 加一；例如 `v5.0.5 → v5.0.6`，這只是格式範例。
3. 確認此次要發布的全部相關變更。先前尚未上傳的程式變更若一起發布，不得把最後一輪文件修改誤認為純文件發布。
4. 依專案可用的 `vghtpe-release-version` 技能執行固定流程；需連接器發布時使用對應技能。技能不在預期位置時先搜尋，不能省略必要步驟。
5. 只 stage 本次相關檔案，提交程式與必要文件，將新版本 tag 指向該提交。
6. 推送程式分支及版本 tag；只有 push 分支無法觸發 Release。
7. 查驗遠端 tag、Actions 執行與 Release 資產；報告實際結果與連結，不把觸發中說成已完成。

以下為格式示例；確認實際遠端版號與提交後才執行：

```powershell
git tag v5.0.6
git push origin main
git push origin v5.0.6
```

### 純文件例外

若整批發布內容僅有說明文件（例如 AGENTS.md、README.md、docs 的文字說明）且完全沒有程式、依賴、資料、打包或流程變更，可依既有規則只推送文件，不遞增版本。修改 workflows 不屬於純文件。不要自行擴大例外範圍。

### 禁止提交

- `.venv/`、`.cache/`、node_modules、frontend/dist、build、dist。
- 報表／擷取產出 PDF、ZIP、執行日誌、個人設定、data.json.bak。
- 無關暫存檔、既有使用者改動及未要求加入的本機技能／工具目錄。
- 根目錄雖有歷史追蹤 EXE／PDF，沒有明確要求不要刪除或替換它們，也不要將新產物追加至版本控制。

完成後提供具體修改檔案、可執行／可重現的驗證方式、測試結果及未確認的事項。不要宣稱使用者未要求的發布、版本更新或正式生產驗收已完成。
