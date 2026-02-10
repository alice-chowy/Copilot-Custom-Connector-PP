# Copilot Custom Connector - Project Portal

透過 Microsoft Graph External Connector API，將內部專案管理系統（Project Portal）的資料索引至 Microsoft 365，使其可在 Microsoft Search 和 Microsoft 365 Copilot 中被搜尋與引用。

## 架構概覽

```
PostgreSQL (Project Portal DB)
        │
        ▼
  data_sync.py ──► Microsoft Graph API ──► Microsoft Search / Copilot
        ▲
        │
  .env (credentials)
```

## 索引的資料類型

| 類型 | 說明 | 主要欄位 |
|------|------|----------|
| Project | 專案 | 名稱、代碼、狀態、進度、預算、負責人 |
| Milestone | 里程碑 | 標題、截止日、階段、是否為關鍵路徑 |
| Risk | 風險 | 機率、影響、緩解措施 |
| Issue | 問題 | 嚴重程度、根本原因 |

## 執行步驟

### 1. 環境設定

```bash
pip install requests python-dotenv psycopg2-binary
```

建立 `.env` 檔案：

```
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
```

### 2. 建立 External Connection

```bash
python connection_create.py
```

### 3. 註冊 Schema

```bash
python schema_register.py
```

Schema 建立為非同步操作，約需 5–15 分鐘。腳本會自動輪詢狀態直到完成。

### 4. 同步資料

```bash
# 正式模式（需要 PostgreSQL 資料庫）
python data_sync.py

# 測試模式（使用內建測試資料，不需資料庫）
python data_sync.py --test
```

資料庫連線可透過環境變數設定：`DB_HOST`、`DB_PORT`、`DB_NAME`、`DB_USER`、`DB_PASSWORD`。

## 輔助工具

| 檔案 | 用途 |
|------|------|
| `check_status.py` | 檢查 Connection、Schema 狀態與已同步項目數量 |
| `if_connect_success.py` | 列出所有已建立的 External Connections |
| `sdk_psuedo.py` | Microsoft Graph SDK 寫法參考（pseudo code） |

## 前置需求

- Python 3.8+
- Azure AD 應用程式註冊，並授予 `ExternalConnection.ReadWrite.All`、`ExternalItem.ReadWrite.All` 權限
- Microsoft 365 租戶（含 Microsoft Search）

## 專案結構

```
├── config.py               # 統一讀取環境變數
├── connection_create.py     # 步驟 2：建立 External Connection
├── schema_register.py       # 步驟 3：註冊 Schema（30 個欄位）
├── data_sync.py             # 步驟 4：從 DB 同步資料至 Graph API
├── check_status.py          # 檢查連線與同步狀態
├── if_connect_success.py    # 列出所有 Connections
├── sdk_psuedo.py            # Graph SDK 參考寫法
├── .env                     # 機密設定（不納入版控）
└── .gitignore
```
