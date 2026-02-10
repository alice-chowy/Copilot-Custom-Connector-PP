"""
步驟 2：建立 External Connection（使用 requests）
"""
import requests
import json

from config import CONFIG

CONNECTION = {
    "id": "ProjectPortalConnection",
    "name": "Project Portal Connector",
    "description": "Connection to index Project Portal system",
}

# ============================================
# 取得 Access Token
# ============================================
def get_access_token():
    url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    
    payload = {
        "client_id": CONFIG["client_id"],
        "client_secret": CONFIG["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    response = requests.post(url, data=payload)
    data = response.json()
    
    if not response.ok:
        print(f"❌ 取得 Token 失敗：{data}")
        raise Exception(data.get("error_description", "Token request failed"))
    
    print("✅ Access Token 取得成功")
    return data["access_token"]

# ============================================
# 建立 External Connection
# ============================================
def create_connection(token):
    url = "https://graph.microsoft.com/v1.0/external/connections"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.post(url, headers=headers, json=CONNECTION)
    data = response.json()
    
    if not response.ok:
        print(f"❌ 建立 Connection 失敗：{data}")
        raise Exception(data.get("error", {}).get("message", "Failed"))
    
    return data

# ============================================
# 執行
# ============================================
def main():
    print("=" * 50)
    print("開始建立 External Connection...")
    print("=" * 50)
    
    token = get_access_token()
    connection = create_connection(token)
    
    print("\n✅ Connection 建立成功！")
    print(json.dumps(connection, indent=2, ensure_ascii=False))
    print(f"\n狀態: {connection.get('state')}")

if __name__ == "__main__":
    main()