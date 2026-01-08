# 變更日誌

本專案的所有重要變更都會記錄在此檔案中。

格式基於 [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)，
並遵循 [語義化版本](https://semver.org/spec/v2.0.0.html) 規範。

> 📖 **[中文版本](CHANGELOG_zh-tw.md)** | **[English Version](CHANGELOG.md)**

## [1.0.2] - 2026-01-08

### 新增
- Lookup Column 選取功能：動態下拉選單選取 Excel 欄位表頭
- Lookup Column 選取功能的完整測試覆蓋
- 支援查找操作中的中文和特殊字元欄位名稱
- API 請求中特殊字元的自動 URL 編碼
- 透過視覺化欄位選取介面增強使用者體驗

## [1.0.1] - 2026-01-06

### 新增
- Update 和 Delete 操作的 **處理模式** 選項
  - `處理所有符合記錄`（預設）：處理所有符合查找條件的列
  - `僅處理第一筆`：僅處理第一筆符合的列
- 批次更新支援：使用查找一次更新多筆記錄
- 批次刪除支援：使用查找一次刪除多筆記錄
- 效能優化：對唯一識別碼查找使用「僅處理第一筆」模式

### 變更
- 預設行為：查找操作現在預設處理所有符合的記錄
- 改進零匹配情況的錯誤訊息
- 更新文件，包含兩種處理模式的詳細範例

### 測試
- 新增處理模式功能的完整測試案例
- 批次更新情境的測試覆蓋
- 批次刪除情境的測試覆蓋
- 處理模式參數處理的驗證測試

## [1.0.0] - 2025-12-23

### 新增
- n8n-nodes-excel-api 初始發布
- 五個核心操作：Append、Read、Update、Delete、Batch
- Append 操作的 **物件模式**
  - 自動讀取 Excel 表頭（第一列）
  - 智能欄位名稱對應
  - 忽略未知欄位並提供警告
  - 靈活的欄位順序
- Append 操作的 **陣列模式**
  - 精確的欄位順序指定
- **Read 操作**
  - 可選的儲存格範圍指定
  - 自動表頭偵測和物件轉換
- **Update 操作**
  - 依列號更新
  - 依查找更新（欄位值搜尋）
- **Delete 操作**
  - 依列號刪除
  - 依查找刪除（欄位值搜尋）
- **Batch 操作**
  - 在一次 API 呼叫中執行多個操作
  - 混合不同操作類型
- 與 [Excel API Server](https://github.com/code4Copilot/excel-api-server) 整合
- 透過 API 伺服器檔案鎖定提供並行安全保護
- 支援多使用者同時存取
- 完整的錯誤處理
- 詳細的文件與使用範例
- TypeScript 實作，具備完整型別安全
- ESLint 設定以確保程式碼品質
- Jest 測試框架設定
- Gulp 建置系統用於圖示處理

### 文件
- 完整的 README，包含安裝說明
- 常見情境的使用範例
- 所有操作的 API 參考
- 疑難排解指南
- 安全性最佳實踐
- 效能優化建議

### 開發者體驗
- TypeScript 支援與型別定義
- 自動化建置流程
- Linting 和格式化工具
- Jest 測試套件
- GitHub 儲存庫設定
- 準備就緒可發布至 npm

## [0.9.0] - 2025-12-20（預發布版）

### 新增
- 初始開發版本
- 基本 CRUD 操作
- Excel API Server 整合
- 開發和測試基礎設施

---

## 發布說明

### 版本 1.0.1 重點
此版本為 Update 和 Delete 操作引入了強大的批次處理功能。新的處理模式選項讓您可以控制要處理所有符合的記錄還是僅處理第一筆，使唯一識別碼查找和批次操作都更有效率。

**主要使用情境：**
- **唯一 ID 查找**：使用「僅處理第一筆」模式搭配員工編號、Email 等
- **批次更新**：使用「處理所有符合記錄」模式更新整個部門
- **批次刪除**：使用「處理所有符合記錄」模式進行資料清理操作

### 版本 1.0.0 重點
第一個穩定版本提供了在 n8n 工作流程中管理 Excel 檔案的完整解決方案，具備並行安全性。物件模式功能透過使用欄位名稱而不是記住欄位位置，讓 Excel 資料處理變得更容易。

**主要功能：**
- **並行安全**：內建檔案鎖定防止資料損毀
- **使用者友善**：具備自動表頭對應的物件模式
- **靈活**：支援簡單和複雜的使用情境
- **正式環境就緒**：完整的測試和文件

---

## 遷移指南

### 從 1.0.0 升級到 1.0.1

沒有破壞性變更。所有現有工作流程無需修改即可繼續運作。

**新增的可選參數：**
- Update 和 Delete 操作使用查找時新增 `processMode` 參數
- 預設值為 `"all"` 以維持向後相容性
- 要優化唯一識別碼查找的效能，明確設定為 `"first"`

**範例：**
```json
{
  "operation": "update",
  "identifyBy": "lookup",
  "lookupColumn": "員工編號",
  "lookupValue": "E100",
  "processMode": "first",  // 新增：加入此參數以獲得更好的效能
  "valuesToSet": {
    "薪資": "85000"
  }
}
```

---

## 連結

- [GitHub 儲存庫](https://github.com/code4Copilot/n8n-nodes-excel-api)
- [npm 套件](https://www.npmjs.com/package/n8n-nodes-excel-api)
- [Excel API Server](https://github.com/code4Copilot/excel-api-server)
- [n8n 文件](https://docs.n8n.io)

## 貢獻

我們歡迎貢獻！詳情請參閱我們的 [貢獻指南](CONTRIBUTING.md)。

## 支援

- 回報錯誤：[GitHub Issues](https://github.com/code4Copilot/n8n-nodes-excel-api/issues)
- 詢問問題：[n8n 社群論壇](https://community.n8n.io)
- Email：hueyan.chen@gmail.com
