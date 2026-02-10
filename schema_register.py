"""
步驟 3：註冊 Schema（使用 requests）
Schema 建立是非同步操作，需要 5-15 分鐘完成
"""
import requests
import json
import time

from config import CONFIG

# 你在步驟 2 建立的 Connection ID
CONNECTION_ID = "ProjectPortalConnection"

# ============================================
# Schema 定義
# ============================================
SCHEMA = {
    "baseType": "microsoft.graph.externalItem",
    "properties": [
        # === 必要語意標籤欄位（Copilot 需要） ===
        {
            "name": "itemType",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "title",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
            "labels": ["title"],
        },
        {
            "name": "description",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": False,
        },
        {
            "name": "url",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": False,
            "labels": ["url"],
        },
        {
            "name": "iconUrl",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": False,
            "labels": ["iconUrl"],
        },
        {
            "name": "lastModifiedDateTime",
            "type": "DateTime",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
            "labels": ["lastModifiedDateTime"],
        },
        {
            "name": "createdDateTime",
            "type": "DateTime",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
            "labels": ["createdDateTime"],
        },
        {
            "name": "lastModifiedBy",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "labels": ["lastModifiedBy"],
        },
        # === 專案相關欄位 ===
        {
            "name": "projectCode",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "projectName",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
        },
        {
            "name": "projectId",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
        },
        {
            "name": "status",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "priority",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "progress",
            "type": "Int64",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "startDate",
            "type": "DateTime",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "endDate",
            "type": "DateTime",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "dueDate",
            "type": "DateTime",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "category",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "phase",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        # === 人員相關欄位 ===
        {
            "name": "owners",
            "type": "StringCollection",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
        },
        {
            "name": "managers",
            "type": "StringCollection",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
        },
        {
            "name": "teamMembers",
            "type": "StringCollection",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": False,
        },
        {
            "name": "tags",
            "type": "StringCollection",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        # === 風險/問題專用欄位 ===
        {
            "name": "severity",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "probability",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "impact",
            "type": "String",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "isCriticalPath",
            "type": "Boolean",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
        },
        {
            "name": "mitigation",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": False,
        },
        {
            "name": "rootCause",
            "type": "String",
            "isSearchable": True,
            "isRetrievable": True,
            "isQueryable": False,
        },
        # === 財務欄位 ===
        {
            "name": "budget",
            "type": "Double",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
            "isRefinable": True,
        },
        {
            "name": "budgetUsed",
            "type": "Double",
            "isSearchable": False,
            "isRetrievable": True,
            "isQueryable": True,
        },
    ],
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
        "grant_type": "client_credentials",
    }

    response = requests.post(url, data=payload)
    data = response.json()

    if not response.ok:
        print(f"❌ 取得 Token 失敗：{data}")
        raise Exception(data.get("error_description", "Token request failed"))

    print("✅ Access Token 取得成功")
    return data["access_token"]


# ============================================
# 註冊 Schema（非同步操作）
# ============================================
def register_schema(token):
    url = f"https://graph.microsoft.com/v1.0/external/connections/{CONNECTION_ID}/schema"

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    response = requests.patch(url, headers=headers, json=SCHEMA)

    # 成功會回傳 202 Accepted
    if response.status_code == 202:
        operation_url = response.headers.get("Location")
        print("✅ Schema 註冊請求已送出")
        print(f"📍 Operation URL: {operation_url}")
        return operation_url
    else:
        data = response.json()
        print(f"❌ Schema 註冊失敗：{json.dumps(data, indent=2, ensure_ascii=False)}")
        raise Exception(data.get("error", {}).get("message", "Failed"))


# ============================================
# 輪詢 Schema 建立狀態（已修正錯誤處理）
# ============================================
def poll_schema_status(token, operation_url):
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(operation_url, headers=headers)
    
    if not response.ok:
        print(f"⚠️ 輪詢請求失敗: {response.status_code}")
        return {"status": "unknown", "error": f"HTTP {response.status_code}"}
    
    data = response.json()
    
    # 安全取得 error message
    error_obj = data.get("error")
    error_msg = None
    if error_obj and isinstance(error_obj, dict):
        error_msg = error_obj.get("message")
    
    return {
        "status": data.get("status", "unknown"),
        "error": error_msg,
        "raw": data  # 保留原始回應以便除錯
    }


# ============================================
# 等待 Schema 建立完成
# ============================================
def wait_for_schema_ready(token, operation_url, max_wait_minutes=20, poll_interval_seconds=30):
    print(f"\n⏳ 等待 Schema 建立完成（最多 {max_wait_minutes} 分鐘）...")

    start_time = time.time()
    max_wait_seconds = max_wait_minutes * 60

    while time.time() - start_time < max_wait_seconds:
        result = poll_schema_status(token, operation_url)
        status = result["status"]

        if status == "completed":
            print("\n✅ Schema 建立完成！")
            return True

        if status == "failed":
            print(f"\n❌ Schema 建立失敗：{result['error']}")
            print(f"   詳細資訊：{json.dumps(result.get('raw', {}), indent=2, ensure_ascii=False)}")
            return False

        elapsed = int(time.time() - start_time)
        print(f"   狀態: {status} | 已等待: {elapsed}s")
        time.sleep(poll_interval_seconds)

    print("\n⚠️ 等待逾時，請稍後手動檢查狀態")
    return False


# ============================================
# 檢查現有 Schema
# ============================================
def get_current_schema(token):
    url = f"https://graph.microsoft.com/v1.0/external/connections/{CONNECTION_ID}/schema"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(url, headers=headers)

    if response.ok:
        return response.json()
    elif response.status_code == 404:
        return None
    else:
        data = response.json()
        print(f"⚠️ 查詢 Schema 失敗：{data}")
        return None


# ============================================
# 單獨檢查 Operation 狀態（可手動呼叫）
# ============================================
def check_operation_status(operation_id=None):
    """
    手動檢查 schema operation 狀態
    用法: check_operation_status("6068921f-5a6f-33d9-3966-1cac9df82949")
    """
    token = get_access_token()
    
    if operation_id:
        operation_url = f"https://graph.microsoft.com/v1.0/external/connections/{CONNECTION_ID}/operations/{operation_id}"
    else:
        # 取得所有 operations
        operation_url = f"https://graph.microsoft.com/v1.0/external/connections/{CONNECTION_ID}/operations"
    
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(operation_url, headers=headers)
    
    print(f"Status Code: {response.status_code}")
    print(json.dumps(response.json(), indent=2, ensure_ascii=False))
    return response.json()


# ============================================
# 執行
# ============================================
def main():
    print("=" * 60)
    print("步驟 3：註冊 Schema")
    print("=" * 60)

    token = get_access_token()

    # 先檢查是否已有 Schema
    print("\n📋 檢查現有 Schema...")
    existing = get_current_schema(token)
    if existing and existing.get("properties"):
        print(f"⚠️ 已存在 Schema，共 {len(existing['properties'])} 個欄位")
        confirm = input("是否要更新 Schema？(y/N): ")
        if confirm.lower() != "y":
            print("取消操作")
            return

    # 註冊 Schema
    print(f"\n📝 正在註冊 Schema 到 Connection: {CONNECTION_ID}")
    print(f"   欄位數量: {len(SCHEMA['properties'])}")

    operation_url = register_schema(token)

    # 等待完成
    success = wait_for_schema_ready(token, operation_url)

    if success:
        print("\n" + "=" * 60)
        print("🎉 Schema 註冊完成！")
        print("   下一步：執行步驟 4 - 同步資料 (data_sync.py)")
        print("=" * 60)


if __name__ == "__main__":
    main()