# n8n-nodes-excel-api

[![npm version](https://badge.fury.io/js/n8n-nodes-excel-api.svg)](https://badge.fury.io/js/n8n-nodes-excel-api)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

n8n 社群節點，透過 API 存取 Excel 檔案，具備**並行安全保護**。完美適用於多使用者同時透過 n8n 工作流程存取相同 Excel 檔案的場景。

## 🎯 為什麼需要這個節點？

### 問題所在
直接在 n8n 中使用 Excel 檔案時：
- ❌ 多個工作流程同時存取同一檔案會導致檔案損毀
- ❌ 並行寫入時會發生資料覆蓋與遺失
- ❌ 缺乏檔案鎖定機制
- ❌ 難以處理多人同時提交的 Webhook 表單

### 解決方案
本節點搭配 [Excel API Server](https://github.com/code4Copilot/excel-api-server) 提供：
- ✅ **檔案鎖定** - 自動佇列管理並行請求
- ✅ **資料完整性** - 無資料遺失或損毀
- ✅ **多使用者支援** - 完美適用於多人提交的 HTML 表單
- ✅ **類似 Google Sheets 的介面** - 在 n8n 中熟悉的操作方式
- ✅ **批次操作** - 高效的大量更新

## 📦 安裝方式

### 方法 1：npm（推薦）

```bash
npm install n8n-nodes-excel-api
```

### 方法 2：手動安裝

```bash
# 1. 複製儲存庫
git clone https://github.com/code4Copilot/n8n-nodes-excel-api.git
cd n8n-nodes-excel-api

# 2. 安裝相依套件
npm install

# 3. 建置
npm run build

# 4. 連結到 n8n
npm link
cd ~/.n8n
npm link n8n-nodes-excel-api

# 5. 重新啟動 n8n
n8n start
```

### 方法 3：社群套件（發布後）

在 n8n 中：
1. 前往 **設定** → **社群節點**
2. 點擊 **安裝**
3. 輸入：`n8n-nodes-excel-api`
4. 點擊 **安裝**

## 🚀 前置需求

**必須先執行 Excel API Server！**

安裝並啟動 [Excel API Server](https://github.com/code4Copilot/excel-api-server)：

```bash
# 使用 Docker 快速啟動
docker run -d \
  -p 8000:8000 \
  -v $(pwd)/data:/app/data \
  -e API_TOKEN=your-secret-token \
  yourusername/excel-api-server
```

詳細資訊請參閱 [Excel API Server 文件](https://github.com/code4Copilot/excel-api-server)。

## 🔧 設定

### 1. 設定憑證

在 n8n 中：
1. 前往 **憑證** → **新增**
2. 搜尋「Excel API」
3. 填寫：
   - **API URL**：`http://localhost:8000`（您的 API 伺服器位址）
   - **API Token**：`your-secret-token`（來自 Excel API Server）
4. 點擊 **儲存**

### 2. 將節點加入工作流程

1. 建立或開啟工作流程
2. 點擊 **新增節點**
3. 搜尋「Excel API」
4. 選擇節點
5. 選擇您的憑證
6. 設定操作

## 📚 操作說明

### 1. Append（附加）
在工作表末端新增一列資料。

**兩種模式：**

#### Object Mode（物件模式）- 推薦
使用欄位名稱對應，更安全且易於維護。

**範例：**
```json
{
  "員工編號": "{{ $json.body.employeeId }}",
  "姓名": "{{ $json.body.name }}",
  "部門": "{{ $json.body.department }}",
  "職位": "{{ $json.body.position }}",
  "薪資": "{{ $json.body.salary }}"
}
```

**特色：**
- ✅ 自動讀取 Excel 表頭（第一列）
- ✅ 按照欄位名稱智能對應
- ✅ 忽略未知欄位，並在回應中提示
- ✅ 欄位順序可任意調整
- ✅ 缺少的欄位會自動填入空值

#### Array Mode（陣列模式）
依照精確的欄位順序指定值。

**範例：**
```json
["E100", "江小魚", "人資部", "經理", "70000"]
```

**注意：** 值的順序必須與 Excel 欄位順序完全對應。

### 2. Read（讀取）
從 Excel 檔案讀取資料。

**參數：**
- `file`：檔案名稱（例如：`employees.xlsx`）
- `sheet`：工作表名稱（預設：`Sheet1`）
- `range`：儲存格範圍（例如：`A1:D10`，留空讀取全部資料）

**輸出：**
- 若偵測到表頭，自動將第一列轉換為欄位名稱
- 回傳物件陣列，以表頭作為鍵值
- 若無表頭則回傳原始資料陣列

### 3. Update（更新）
更新現有列的資料。

**識別方式：**

#### 依列號（Row Number）
直接指定要更新的列號（從 2 開始，第 1 列為表頭）。

**範例：**
```json
{
  "operation": "update",
  "identifyBy": "rowNumber",
  "rowNumber": 5,
  "valuesToSet": {
    "狀態": "已完成",
    "更新日期": "2025-12-21"
  }
}
```

#### 依查找（Lookup）
透過查找特定欄位的值來找到要更新的列。

**範例：**
```json
{
  "operation": "update",
  "identifyBy": "lookup",
  "lookupColumn": "員工編號",
  "lookupValue": "E100",
  "valuesToSet": {
    "薪資": "80000",
    "職位": "資深經理"
  }
}
```

### 4. Delete（刪除）
從工作表中刪除一列。

**識別方式：**

#### 依列號
```json
{
  "operation": "delete",
  "identifyBy": "rowNumber",
  "rowNumber": 5
}
```

#### 依查找
```json
{
  "operation": "delete",
  "identifyBy": "lookup",
  "lookupColumn": "員工編號",
  "lookupValue": "E100"
}
```

### 5. Batch（批次）
一次執行多個操作（更有效率）。

**範例：**
```json
{
  "operations": [
    {
      "type": "append",
      "values": ["E010", "Alice", "行銷部", "專員", "65000"]
    },
    {
      "type": "update",
      "row": 5,
      "values": ["E005", "Updated Name", "IT部", "經理", "90000"]
    },
    {
      "type": "delete",
      "row": 10
    }
  ]
}
```

## 🎨 使用範例

## 🎨 使用範例

### 範例 1：Webhook 表單寫入 Excel

完美適用於多人同時提交表單的場景！

```
┌──────────────────┐
│ Webhook          │  接收表單提交
│ POST /submit     │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  操作：Append（物件模式）
│                  │  檔案：registrations.xlsx
│                  │  值：{
│                  │    "姓名": "{{ $json.body.name }}",
│                  │    "Email": "{{ $json.body.email }}",
│                  │    "電話": "{{ $json.body.phone }}",
│                  │    "提交時間": "{{ $now }}"
│                  │  }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Respond Webhook  │  回傳成功訊息
└──────────────────┘
```

**HTML 表單：**
```html
<form id="registrationForm">
  <input type="text" name="name" placeholder="姓名" required>
  <input type="email" name="email" placeholder="Email" required>
  <input type="tel" name="phone" placeholder="電話" required>
  <button type="submit">提交</button>
</form>

<script>
document.getElementById('registrationForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(e.target);
  await fetch('YOUR_WEBHOOK_URL', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(Object.fromEntries(formData))
  });
  alert('提交成功！');
});
</script>
```

### 範例 2：每日報表產生

```
┌──────────────────┐
│ Schedule         │  每天早上 9:00
│ 0 9 * * *        │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  操作：Read
│ (讀取)           │  檔案：sales.xlsx
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Filter           │  篩選今日記錄
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Send Email       │  發送每日報表
└──────────────────┘
```

### 範例 3：批次更新

```
┌──────────────────┐
│ Code             │  準備操作陣列
│                  │  operations = [...]
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  操作：Batch
│ (批次)           │  檔案：data.xlsx
│                  │  操作：{{ $json.operations }}
└──────────────────┘
```

### 範例 4：透過員工編號更新薪資

```
┌──────────────────┐
│ Webhook          │  接收更新請求
│ POST /update     │  { "employeeId": "E100", "salary": 85000 }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  操作：Update
│                  │  識別方式：Lookup
│                  │  查找欄位：員工編號
│                  │  查找值：{{ $json.body.employeeId }}
│                  │  設定值：{ "薪資": "{{ $json.body.salary }}" }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Respond Webhook  │  回傳更新結果
└──────────────────┘
```

## 🧪 並行測試

測試 10 個同時提交：

```javascript
// concurrent_test.js
const promises = [];
for (let i = 0; i < 10; i++) {
  promises.push(
    fetch('YOUR_WEBHOOK_URL', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        員工編號: `E${String(i).padStart(3, '0')}`,
        姓名: `測試使用者 ${i}`,
        時間戳記: new Date().toISOString()
      })
    })
  );
}

await Promise.all(promises);
console.log('所有請求完成！');
```

**結果：** 所有 10 筆記錄都會安全地寫入 Excel，不會有資料遺失或損毀！

## ⚠️ 常見問題

### 問題 1：節點未顯示在 n8n 中

**解決方法：**
```bash
# 重新啟動 n8n
pkill -f n8n
n8n start

# 或使用 pm2
pm2 restart n8n
```

### 問題 2：API 連線失敗

**解決方法：**
- 檢查 Excel API Server 是否正在執行：`curl http://localhost:8000/`
- 驗證憑證中的 API URL 是否正確
- 檢查 API Token 是否正確
- 檢查防火牆設定

### 問題 3：「找不到參數」錯誤

**原因：** 參數名稱設定錯誤

**解決方法：**
- 確認選擇了正確的 Append Mode（Object 或 Array）
- Object Mode：使用 `appendValuesObject` 參數
- Array Mode：使用 `appendValuesArray` 參數
- 檢查 JSON 格式是否正確

### 問題 4：「檔案鎖定」錯誤

**原因：** 並行請求過多或 API 伺服器問題

**解決方法：**
- 稍等片刻後重試
- 檢查 API 伺服器狀態
- 必要時重新啟動 Excel API Server

## 🔐 安全性

### 最佳實踐

1. **使用強式 API Token**
   ```bash
   # 產生安全的 token
   openssl rand -hex 32
   ```

2. **在正式環境使用 HTTPS**
   - 設定反向代理（Nginx）
   - 使用 SSL 憑證

3. **限制存取**
   - 僅允許信任的網路存取 API URL
   - 遠端存取時使用 VPN

4. **定期備份**
   - 設定 Excel 檔案自動備份
   - 將備份儲存在安全位置

## 📊 效能優化建議

### 1. 使用批次操作
```javascript
// ❌ 不好：多次單一操作
for (item of items) {
  await appendRow(item);
}

// ✅ 好：一次批次操作
await batchOperations(items.map(item => ({
  type: "append",
  values: item.values
})));
```

### 2. 讀取時指定範圍
```javascript
// ❌ 不好：讀取整個檔案
range: ""

// ✅ 好：只讀取需要的範圍
range: "A1:D100"
```

### 3. 使用高效的工作流程
- 在一個工作流程中組合相關操作
- 減少 API 呼叫次數
- 適當使用快取

## 🆕 最新功能

### Object Mode（物件模式）
- ✅ 使用 `/api/excel/append_object` API
- ✅ 自動讀取 Excel 表頭（第一列）
- ✅ 按照欄位名稱智能對應
- ✅ 忽略未知欄位，並在回應中提示
- ✅ 不需要記住欄位順序

### 進階更新與刪除
- ✅ 支援依列號直接操作
- ✅ 支援依查找欄位值來操作
- ✅ 可更新特定欄位而不影響其他欄位

## 🤝 貢獻

歡迎貢獻！

1. Fork 此儲存庫
2. 建立您的功能分支：`git checkout -b feature/AmazingFeature`
3. 提交您的變更：`git commit -m 'Add some AmazingFeature'`
4. 推送到分支：`git push origin feature/AmazingFeature`
5. 開啟 Pull Request

## 📄 授權

MIT 授權 - 詳見 [LICENSE](LICENSE) 檔案

## 🔗 相關專案

- [Excel API Server](https://github.com/code4Copilot/excel-api-server) - 後端 API 伺服器（必要）
- [n8n](https://github.com/n8n-io/n8n) - 工作流程自動化工具

## 📧 支援

- GitHub Issues：[回報問題](https://github.com/code4Copilot/n8n-nodes-excel-api/issues)
- Email：your.email@example.com
- n8n 社群：[n8n 論壇](https://community.n8n.io)

## ⭐ Star 歷史

如果這個專案對您有幫助，請給它一個 ⭐！

---

**用 ❤️ 為 n8n 社群打造**