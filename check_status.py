# 建立檔案: check_status.py
import requests
import json
from config import CONFIG
CONNECTION_ID = "ProjectPortalConnection"

def get_token():
    url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "client_id": CONFIG["client_id"],
        "client_secret": CONFIG["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    })
    return resp.json()["access_token"]

def check_connection():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    # 1. 檢查 Connection 狀態
    conn_url = f"https://graph.microsoft.com/v1.0/external/connections/{CONNECTION_ID}"
    conn_resp = requests.get(conn_url, headers=headers)
    print("=== Connection 狀態 ===")
    print(json.dumps(conn_resp.json(), indent=2, ensure_ascii=False))
    
    # 2. 檢查 Schema 狀態
    schema_url = f"{conn_url}/schema"
    schema_resp = requests.get(schema_url, headers=headers)
    print("\n=== Schema 狀態 ===")
    if schema_resp.ok:
        schema = schema_resp.json()
        print(f"欄位數量: {len(schema.get('properties', []))}")
        print(f"狀態: {schema.get('status', 'N/A')}")
    else:
        print(f"錯誤: {schema_resp.status_code}")
        print(schema_resp.text)
    
    # 3. 檢查已同步的項目數量
    items_url = f"{conn_url}/items"
    items_resp = requests.get(items_url, headers=headers)
    print("\n=== 已同步項目 ===")
    if items_resp.ok:
        items = items_resp.json()
        print(f"項目數量: {len(items.get('value', []))}")
    else:
        print(f"無法取得項目列表")

if __name__ == "__main__":
    check_connection()