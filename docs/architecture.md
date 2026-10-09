# React Windows 桌面架構

## 模組責任

- React / TypeScript / Vite：左側導覽、四個工作頁面與共用任務面板。
- FastAPI：驗證輸入、讀取目錄與設定、建立與取消任務、提供原生檔案操作入口。
- Python 核心：月份與網址、wkhtmltopdf 子程序、PyMuPDF 擷取與原子檔案寫入。
- pywebview：WebView2 視窗、多選 PDF、原生拖曳完整路徑、資料夾選擇、關閉應用程式。

React 不直接存取檔案系統，也不在介面中執行 PDF 處理。後端核心不依賴 Tkinter 或 React。

## 任務生命週期

`running → completed / failed`；取消時為 `running → cancelling → cancelled`。

每個任務有自己的 `threading.Event`。後端在工作執行緒退出前拒絕建立另一個任務。
切換前端頁面不卸載工作表單，任務查詢位於 App 共用層，每秒查詢一次。
取消下載時終止並回收 wkhtmltopdf 子程序，重試等待可立即取消。

成功、跳過、失敗按檔案計數；已取消表示未完成項目數。進度到 100% 表示所有項目已處理，不代表每一份報表都下載成功，應同時檢查失敗數。
任務與最近 200 筆紀錄保存在記憶體，關閉程式後不保留，也不自動續傳。

## API

| 路徑 | 用途 |
|---|---|
| GET /api/bootstrap | 報表目錄、設定、上個月與工作階段權杖 |
| GET /api/catalog | 完整報表欄位、檔案位置與資料版本 |
| POST /api/catalog/reports | 驗證、備份並新增報表 |
| PUT /api/catalog/reports/{id} | 修改報表及類型 |
| DELETE /api/catalog/reports/{id} | 刪除目錄項目，PDF 保留 |
| GET /api/tasks/current | 目前或最近一個任務 |
| POST /api/tasks/download | 選定報表與月份區間下載 |
| POST /api/tasks/extract | 多個 PDF 的關鍵字擷取 |
| POST /api/tasks/{id}/cancel | 取消指定任務 |
| PUT /api/settings | 原子儲存設定 |
| POST /api/desktop/select-pdfs | Windows 多選 PDF |
| POST /api/desktop/select-folder | Windows 選擇資料夾 |
| POST /api/desktop/open-folder | 開啟下載、擷取或目前任務資料夾 |

正式模式由 FastAPI 提供 Vite 建置產物，前端與 API 同源。
開發模式使用 Vite 代理 API。服務限制在 loopback，API 驗證工作階段權杖與請求來源。

## 資源與持久化

- `data.json`：原始碼根目錄或 EXE 同層；由報表管理 API 驗證、保存並立即同步清單。
- `data.json.bak`：每次變更前保留上一版完整檔案，未知欄位不會被移除。資料使用 SHA-256 版本比對，防止舊清單覆蓋新資料。
- 靜態前端與 wkhtmltopdf：打包後由 PyInstaller 資源目錄載入。
- 設定與錯誤日誌：APPDATA，可用 `VGHTPE_STATE_DIR` 指定驗證目錄。
- 下載報表：類別／月份子資料夾，驗證暫存 PDF 後才替換正式檔案。
- 擷取報表：日期子資料夾。同名來源以月份或遞增編號區分，不覆蓋既有輸出。

## 驗證範圍

Python 測試涵蓋月份、網址、設定、下載殘檔、取消競態、PDF 頁面內容與 API。
Vitest 驗證前端月份區間與 Windows 路徑顯示。
Playwright 使用實際 FastAPI 加上隔離的轉換器與對話框，驗證搜尋、下載、跨頁進度、取消、PDF 擷取、設定保存，以及報表新增、修改、刪除確認、保存後重載與版本衝突。
WebView2 桌面啟動仍待一般 Windows 桌面實測，詳細限制見 verification.md。
真實報表網站的可用性與內容需另外以實際報表驗證。

## PDF 拖曳

`desktop/drop.py` 在 WebView2 載入後，等待 React 工作區就緒，再以 pywebview DOMEventHandler 訂閱 `#pdf-drop-zone` 的 drop 事件。每次頁面重新載入也會重新訂閱，同一元素不重複註冊。

原生事件提供 `pywebviewFullPath`；Python 驗證實際 PDF 與路徑後，以 `vghtpe:pdf-drop` CustomEvent 將清單傳回 React。前端依 Windows 路徑忽略大小寫與斜線差異，合併原清單，不將 File.name 當成本機路徑。拖曳區域隱藏或任務執行中時，兩端都拒絕加入。視窗攔截檔案拖放的預設導覽，避免開啟 PDF 後離開工作區。

原生事件實作依據 [pywebview 官方拖曳範例](https://pywebview.flowrl.com/examples/drag_drop)。瀏覽器流程測試模擬原生事件通知，實際 Windows 檔案總管到 WebView2 的拖曳仍需桌面實測。
