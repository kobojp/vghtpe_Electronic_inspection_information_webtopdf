# 專案 Agent 規則 (Project Agent Rules)

這份文件定義了 AI Agent 在本專案中必須遵守的開發與操作原則。

## 1. 更新 GitHub 必須伴隨版本號更新 (強制)
當使用者要求「上傳 GitHub」、「更新 GitHub」、「發布變更」時，Agent **絕對不能**只做單純的程式碼 Push，必須確保觸發版本號更新流程。

### 版本號更新機制：
- 本專案的 CI/CD (GitHub Actions) 依賴 **Git Tags (例如 `v5.0.6`)** 來觸發自動編譯與打包 `release.yml`。
- 因此，當 Agent 要上傳變更時，必須執行以下步驟（或直接呼叫專案內的 `vghtpe-release-version` 技能）：
  1. **確認最新版本**：列出遠端現有的標籤 (例如 `v5.0.5`)。
  2. **版號遞增**：將 Patch 號碼加 1 (變成 `v5.0.6`)。
  3. **打上標籤**：`git tag v5.0.6`
  4. **推送到遠端**：除了 push 程式碼，也必須 `git push origin v5.0.6`。

### 例外情況 (Exceptions)：
- 若本次變更**僅包含**說明文件（例如 `.md` 檔、`AGENTS.md`、`README.md` 等），且**完全沒有**碰到任何程式碼，則**不需要**遞增與發布版本號，只需單純推送到 GitHub 即可。

## 2. 開發規範
- **語言**：必須以繁體中文與使用者溝通。
- **測試**：打包發布前應確保沒有語法錯誤。
- **純淨提交**：上傳 GitHub 時，只 commit 和本次任務相關的檔案，切勿將本機的暫存檔、報表產出檔 (`.pdf`, `.zip`) 等加入版本控制。


## 3. 待重構與優化方向 (Future Refactoring Plans)
未來 AI Agent 若被要求進行系統優化或程式碼整理，請優先執行以下技術債處理：

1. **程式碼拆分模組化 (Modularization)**
   - 目前 `main.py` 已超過 2200 行。必須拆分為 `gui.py` (UI 介面)、`pdf_core.py` (PDF 合併與擷取邏輯)、`config.py` (設定檔存取)。
2. **導入錯誤日誌檔 (Error Logging)**
   - 將系統的 `Exception` 透過 Python `logging` 模組寫入本地端的 `logs/error.log`，確保 `.exe` 打包版發生閃退時有跡可循。
3. **Tkinter 執行緒安全 (Thread Safety)**
   - 所有的 UI 更新與 `messagebox` 彈跳視窗，若由背景執行緒 (如 `_merge_pdf_task`) 觸發，**絕對必須**使用 `self.after(0, ...)` 導回主執行緒，防止介面當機或凍結。
4. **UI 主題統一管理 (Theme Manager)**
   - 將 `#0078D7`、`#f5f5f5`、`微軟正黑體` 等魔法字串 (Magic Strings) 抽取為全域常數或設定檔，方便日後切換深色模式或統一調整視覺。
5. **檔案鎖定防呆檢查 (File Permission Handling)**
   - 在輸出或覆蓋 PDF 檔案前，攔截 `PermissionError`，並友善提示使用者「請先關閉正在閱讀的 PDF 檔案再試」。
