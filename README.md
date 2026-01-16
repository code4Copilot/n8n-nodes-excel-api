# n8n-nodes-excel-api

[![npm version](https://badge.fury.io/js/n8n-nodes-excel-api.svg)](https://badge.fury.io/js/n8n-nodes-excel-api)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

An n8n community node for accessing Excel files via API with **concurrent safety protection**. Perfect for scenarios where multiple users simultaneously access the same Excel file through n8n workflows.

> 📖 **[中文文檔](README_zh-tw.md)** | **[English Documentation](README.md)**

## 🎯 Why This Node?

### The Problem
When working with Excel files directly in n8n:
- ❌ Multiple workflows accessing the same file cause file corruption
- ❌ Concurrent writes lead to data overwrite and loss
- ❌ No file locking mechanism
- ❌ Difficult to handle simultaneous webhook form submissions

### The Solution
This node works with [Excel API Server](https://github.com/code4Copilot/excel-api-server) to provide:
- ✅ **File Locking** - Automatically queue concurrent requests
- ✅ **Data Integrity** - No data loss or corruption
- ✅ **Multi-User Support** - Perfect for multi-user HTML form submissions
- ✅ **Google Sheets-like Interface** - Familiar operations in n8n
- ✅ **Batch Operations** - Efficient bulk updates

## 📦 Installation

### Method 1: npm (Recommended)

```bash
npm install n8n-nodes-excel-api
```

### Method 2: Manual Installation

```bash
# 1. Clone repository
git clone https://github.com/code4Copilot/n8n-nodes-excel-api.git
cd n8n-nodes-excel-api

# 2. Install dependencies
npm install

# 3. Build
npm run build

# 4. Link to n8n
npm link
cd ~/.n8n
npm link n8n-nodes-excel-api

# 5. Restart n8n
n8n start
```

### Method 3: Community Package (After Publication)

In n8n:
1. Go to **Settings** → **Community Nodes**
2. Click **Install**
3. Enter: `n8n-nodes-excel-api`
4. Click **Install**

## 🚀 Prerequisites

**You must run Excel API Server first!**

Install and start [Excel API Server](https://github.com/code4Copilot/excel-api-server):

```bash
# Quick start with Docker
docker run -d \
  -p 8000:8000 \
  -v $(pwd)/data:/app/data \
  -e API_TOKEN=your-secret-token \
  yourusername/excel-api-server
```

See [Excel API Server Documentation](https://github.com/code4Copilot/excel-api-server) for details.

## 🔧 Configuration

### 1. Set Up Credentials

In n8n:
1. Go to **Credentials** → **New**
2. Search for "Excel API"
3. Fill in:
   - **API URL**: `http://localhost:8000` (Your API server address)
   - **API Token**: `your-secret-token` (From Excel API Server)
4. Click **Save**

### 2. Add Node to Workflow

1. Create or open a workflow
2. Click **Add Node**
3. Search for "Excel API"
4. Select the node
5. Choose your credential
6. Configure operation

## 📚 Operations

### 1. Append
Add a new row to the end of the sheet.

**Two Modes:**

#### Object Mode - Recommended
Map values by column names, safer and easier to maintain.

**Example:**
```json
{
  "Employee ID": "{{ $json.body.employeeId }}",
  "Name": "{{ $json.body.name }}",
  "Department": "{{ $json.body.department }}",
  "Position": "{{ $json.body.position }}",
  "Salary": "{{ $json.body.salary }}"
}
```

**Features:**
- ✅ Automatically read Excel headers (first row)
- ✅ Intelligently map by column names
- ✅ Ignore unknown columns with warnings in response
- ✅ Column order can be arbitrary
- ✅ Missing columns automatically filled with empty values

#### Array Mode
Specify values in exact column order.

**Example:**
```json
["E100", "John Doe", "HR", "Manager", "70000"]
```

**Note:** Value order must exactly match Excel column order.

### 2. Read
Read data from Excel file.

**Parameters:**
- `file`: File name (e.g., `employees.xlsx`)
- `sheet`: Sheet name (default: `Sheet1`)
- `range`: Cell range (e.g., `A1:D10`, leave empty to read all)

**Output:**
- Auto-convert first row to column names if headers detected
- Return array of objects with headers as keys
- Return raw data array if no headers

### 3. Update
Update existing row data.

**Identify Methods:**

#### By Row Number
Directly specify row number to update (starts from 2, row 1 is header).

**Example:**
```json
{
  "operation": "update",
  "identifyBy": "rowNumber",
  "rowNumber": 5,
  "valuesToSet": {
    "Status": "Completed",
    "Update Date": "2025-12-21"
  }
}
```

#### By Lookup
Find rows to update by looking up specific column values.

**Process Modes:**

##### All Matching Records - Default
Update all matching rows, suitable for batch update scenarios.

**Example: Update all IT department employees**
```json
{
  "operation": "update",
  "identifyBy": "lookup",
  "lookupColumn": "Department",
  "lookupValue": "IT",
  "processMode": "all",
  "valuesToSet": {
    "Status": "Reviewed",
    "Review Date": "2026-01-06"
  }
}
```

##### First Match Only
Update only the first matching record, suitable for unique identifier lookups.

**Example: Update specific employee data**
```json
{
  "operation": "update",
  "identifyBy": "lookup",
  "lookupColumn": "Employee ID",
  "lookupValue": "E100",
  "processMode": "first",
  "valuesToSet": {
    "Salary": "80000",
    "Position": "Senior Manager"
  }
}
```

**💡 Usage Tips:**
- When looking up by unique identifiers (Employee ID, Email), use `processMode: "first"` for better performance
- Use `processMode: "all"` when batch updating multiple records
- Default is `"all"` to ensure no matching records are missed

### 4. Delete
Delete a row from the sheet.

**Identify Methods:**

#### By Row Number
```json
{
  "operation": "delete",
  "identifyBy": "rowNumber",
  "rowNumber": 5
}
```

#### By Lookup
Find rows to delete by looking up specific column values.

**Process Modes:**

##### All Matching Records - Default
Delete all matching rows.

**Example: Delete all terminated employees**
```json
{
  "operation": "delete",
  "identifyBy": "lookup",
  "lookupColumn": "Status",
  "lookupValue": "Terminated",
  "processMode": "all"
}
```

##### First Match Only
Delete only the first matching record.

**Example: Delete specific employee**
```json
{
  "operation": "delete",
  "identifyBy": "lookup",
  "lookupColumn": "Employee ID",
  "lookupValue": "E100",
  "processMode": "first"
}
```

**⚠️ Important:**
- Delete operations cannot be undone, use with caution
- When looking up by unique identifiers, use `processMode: "first"`
- Verify lookup conditions carefully when batch deleting to avoid accidental data loss

### 5. Batch
Execute multiple operations at once (more efficient).

**Example:**
```json
{
  "operations": [
    {
      "type": "append",
      "values": ["E010", "Alice", "Marketing", "Specialist", "65000"]
    },
    {
      "type": "update",
      "row": 5,
      "values": ["E005", "Updated Name", "IT", "Manager", "90000"]
    },
    {
      "type": "delete",
      "row": 10
    }
  ]
}
```

## 🎨 Usage Examples

### Example 1: Webhook Form to Excel

Perfect for scenarios with multiple simultaneous form submissions!

```
┌──────────────────┐
│ Webhook          │  Receive form submission
│ POST /submit     │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  Operation: Append (Object Mode)
│                  │  File: registrations.xlsx
│                  │  Values: {
│                  │    "Name": "{{ $json.body.name }}",
│                  │    "Email": "{{ $json.body.email }}",
│                  │    "Phone": "{{ $json.body.phone }}",
│                  │    "Submit Time": "{{ $now }}"
│                  │  }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Respond Webhook  │  Return success message
└──────────────────┘
```

**HTML Form:**
```html
<form id="registrationForm">
  <input type="text" name="name" placeholder="Name" required>
  <input type="email" name="email" placeholder="Email" required>
  <input type="tel" name="phone" placeholder="Phone" required>
  <button type="submit">Submit</button>
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
  alert('Submitted successfully!');
});
</script>
```

### Example 2: Daily Report Generation

```
┌──────────────────┐
│ Schedule         │  Every day at 9:00 AM
│ 0 9 * * *        │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  Operation: Read
│ (Read)           │  File: sales.xlsx
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Filter           │  Filter today's records
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Send Email       │  Send daily report
└──────────────────┘
```

### Example 3: Batch Updates

```
┌──────────────────┐
│ Code             │  Prepare operations array
│                  │  operations = [...]
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  Operation: Batch
│ (Batch)          │  File: data.xlsx
│                  │  Operations: {{ $json.operations }}
└──────────────────┘
```

### Example 4: Update Salary by Employee ID

```
┌──────────────────┐
│ Webhook          │  Receive update request
│ POST /update     │  { "employeeId": "E100", "salary": 85000 }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  Operation: Update
│                  │  Identify By: Lookup
│                  │  Lookup Column: Employee ID
│                  │  Lookup Value: {{ $json.body.employeeId }}
│                  │  Process Mode: First Match Only
│                  │  Values To Set: { "Salary": "{{ $json.body.salary }}" }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Respond Webhook  │  Return update result
└──────────────────┘
```

### Example 5: Batch Department Status Update

**Use Case:** Review all employees in a department at once

```
┌──────────────────┐
│ Webhook          │  Receive batch review request
│ POST /approve    │  { "department": "IT", "status": "Reviewed" }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  Operation: Update
│                  │  Identify By: Lookup
│                  │  Lookup Column: Department
│                  │  Lookup Value: {{ $json.body.department }}
│                  │  Process Mode: All Matching Records
│                  │  Values To Set: {
│                  │    "Status": "{{ $json.body.status }}",
│                  │    "Review Date": "{{ $now.format('YYYY-MM-DD') }}",
│                  │    "Reviewer": "{{ $json.body.reviewer }}"
│                  │  }
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Respond Webhook  │  Return: Updated N records
└──────────────────┘
```

### Example 6: Clean Up Expired Data

**Use Case:** Periodically delete employee records terminated over a year ago

```
┌──────────────────┐
│ Schedule         │  Execute on 1st of month
│ 0 0 1 * *        │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Excel API        │  Operation: Delete
│                  │  Identify By: Lookup
│                  │  Lookup Column: Status
│                  │  Lookup Value: Terminated
│                  │  Process Mode: All Matching Records
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Send Email       │  Notify admin: Cleaned N records
└──────────────────┘
```

## 🧪 Concurrent Testing

Test 10 simultaneous submissions:

```javascript
// concurrent_test.js
const promises = [];
for (let i = 0; i < 10; i++) {
  promises.push(
    fetch('YOUR_WEBHOOK_URL', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        EmployeeID: `E${String(i).padStart(3, '0')}`,
        Name: `Test User ${i}`,
        Timestamp: new Date().toISOString()
      })
    })
  );
}

await Promise.all(promises);
console.log('All requests completed!');
```

**Result:** All 10 records will be safely written to Excel without data loss or corruption!

## ⚠️ Troubleshooting

### Issue 1: Node Not Showing in n8n

**Solution:**
```bash
# Restart n8n
pkill -f n8n
n8n start

# Or with pm2
pm2 restart n8n
```

### Issue 2: API Connection Failed

**Solution:**
- Check if Excel API Server is running: `curl http://localhost:8000/`
- Verify API URL in credentials is correct
- Check API Token is correct
- Check firewall settings

### Issue 3: "Parameter Not Found" Error

**Cause:** Incorrect parameter name configuration

**Solution:**
- Confirm correct Append Mode is selected (Object or Array)
- Object Mode: Use `appendValuesObject` parameter
- Array Mode: Use `appendValuesArray` parameter
- Check JSON format is correct

### Issue 4: "File Lock" Error

**Cause:** Too many concurrent requests or API server issues

**Solution:**
- Wait a moment and retry
- Check API server status
- Restart Excel API Server if necessary

## 🔐 Security

### Best Practices

1. **Use Strong API Token**
   ```bash
   # Generate secure token
   openssl rand -hex 32
   ```

2. **Use HTTPS in Production**
   - Set up reverse proxy (Nginx)
   - Use SSL certificate

3. **Restrict Access**
   - Allow only trusted networks to access API URL
   - Use VPN for remote access

4. **Regular Backups**
   - Set up automatic backups of Excel files
   - Store backups in secure location

## 📊 Performance Optimization Tips

### 1. Use Batch Operations
```javascript
// ❌ Bad: Multiple single operations
for (item of items) {
  await appendRow(item);
}

// ✅ Good: One batch operation
await batchOperations(items.map(item => ({
  type: "append",
  values: item.values
})));
```

### 2. Specify Range When Reading
```javascript
// ❌ Bad: Read entire file
range: ""

// ✅ Good: Only read needed range
range: "A1:D100"
```

### 3. Use Efficient Workflows
- Combine related operations in one workflow
- Reduce number of API calls
- Use caching appropriately

## 🆕 Latest Features

### 🎉 Automatic Type Conversion (v1.0.3)
- ✅ **Smart Type Detection**: Automatically convert strings to appropriate data types
- ✅ **Number Conversion**: `"123"` → `123`, `"45.67"` → `45.67`
- ✅ **Boolean Conversion**: `"true"` → `true`, `"false"` → `false`
- ✅ **Null Conversion**: `"null"` or empty string → `null`
- ✅ **Date Conversion**: ISO format date strings auto-convert (`"2024-01-15"`)
- ✅ **Preserve Typed Values**: Numbers, booleans, etc. remain unchanged
- ✅ **All Operations**: Supported in both Append and Update operations

**Example:**
```json
{
  "EmployeeID": "E001",      // Remains string
  "Age": "30",               // Auto-converts to number 30
  "Salary": "50000.50",     // Auto-converts to 50000.50
  "IsActive": "true",        // Auto-converts to boolean true
  "TerminationDate": "null", // Auto-converts to null
  "HireDate": "2020-01-15"   // Auto-converts to date format
}
```

### Object Mode
- ✅ Uses `/api/excel/append_object` API
- ✅ Automatically reads Excel headers (first row)
- ✅ Intelligently maps by column names
- ✅ Ignores unknown columns with warnings in response
- ✅ No need to remember column order

### Advanced Update and Delete
- ✅ Support operations by row number
- ✅ Support operations by column value lookup
- ✅ Can update specific columns without affecting others
- ✅ Batch processing support with process modes

### Lookup Column Selection
- ✅ Dynamic dropdown selection of Excel headers
- ✅ Support for Chinese and special character column names
- ✅ Automatic URL encoding for special characters
- ✅ Enhanced user experience with visual column selection

## 🤝 Contributing

Contributions are welcome!

1. Fork this repository
2. Create your feature branch: `git checkout -b feature/AmazingFeature`
3. Commit your changes: `git commit -m 'Add some AmazingFeature'`
4. Push to the branch: `git push origin feature/AmazingFeature`
5. Open a Pull Request

## 📄 License

MIT License - see [LICENSE](LICENSE) file

## 🔗 Related Projects

- [Excel API Server](https://github.com/code4Copilot/excel-api-server) - Backend API server (Required)
- [n8n](https://github.com/n8n-io/n8n) - Workflow automation tool

## 📧 Support

- GitHub Issues: [Report Issues](https://github.com/code4Copilot/n8n-nodes-excel-api/issues)
- Email: hueyan.chen@gmail.com
- n8n Community: [n8n Forum](https://community.n8n.io)

## ⭐ Star History

If this project helps you, please give it a ⭐!

---

**Built with ❤️ for the n8n community**
