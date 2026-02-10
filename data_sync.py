"""
步驟 4：同步資料到 Microsoft Graph Connector
將 Projects, Milestones, Risks, Issues 同步到 M365
"""
import requests
import json
import os
from datetime import datetime
from typing import Optional, List, Dict, Any
import psycopg2
from psycopg2.extras import RealDictCursor

from config import CONFIG

CONNECTION_ID = "ProjectPortalConnection"
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# 你的應用程式 URL（用於生成連結）
APP_BASE_URL = os.environ.get("APP_BASE_URL", "https://project.adata-ai.com/")


# 資料庫連線設定（請修改為你的設定）
DATABASE_CONFIG = {
    "host": os.environ.get("DB_HOST", "localhost"),
    "port": os.environ.get("DB_PORT", "5432"),
    "database": os.environ.get("DB_NAME", "your_database"),
    "user": os.environ.get("DB_USER", "your_user"),
    "password": os.environ.get("DB_PASSWORD", "your_password"),
}


# ============================================
# Access Token
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
        raise Exception(f"Token 取得失敗：{data}")
    return data["access_token"]


# ============================================
# 資料庫連線
# ============================================
def get_db_connection():
    return psycopg2.connect(
        host=DATABASE_CONFIG["host"],
        port=DATABASE_CONFIG["port"],
        database=DATABASE_CONFIG["database"],
        user=DATABASE_CONFIG["user"],
        password=DATABASE_CONFIG["password"],
        cursor_factory=RealDictCursor,
    )


# ============================================
# 從資料庫讀取資料
# ============================================
def fetch_projects(conn) -> List[Dict]:
    with conn.cursor() as cur:
        cur.execute("""
            SELECT 
                p.id, p.name, p.code, p.description,
                p.start_date, p.end_date, p.status, p.progress,
                p.budget, p.budget_used, p.priority,
                p.managers, p.team_members, p.tags,
                p.created_at, p.updated_at,
                pc.label as category_label
            FROM projects p
            LEFT JOIN project_categories pc ON p.category_id = pc.id
        """)
        return cur.fetchall()


def fetch_milestones(conn) -> List[Dict]:
    with conn.cursor() as cur:
        cur.execute("""
            SELECT 
                m.id, m.project_id, m.title, m.description,
                m.due_date, m.status, m.priority, m.assigned_to,
                m.category, m.phase, m.is_critical_path,
                m.created_at, m.updated_at,
                p.name as project_name, p.code as project_code
            FROM milestones m
            JOIN projects p ON m.project_id = p.id
        """)
        return cur.fetchall()


def fetch_risks(conn) -> List[Dict]:
    with conn.cursor() as cur:
        cur.execute("""
            SELECT 
                r.id, r.project_ids, r.title, r.description,
                r.deadline, r.probability, r.impact, r.status,
                r.mitigation, r.owners, r.is_critical_path,
                r.created_at, r.updated_at
            FROM risks r
        """)
        return cur.fetchall()


def fetch_issues(conn) -> List[Dict]:
    with conn.cursor() as cur:
        cur.execute("""
            SELECT 
                i.id, i.project_ids, i.title, i.description,
                i.due_date, i.severity, i.status, i.owners,
                i.root_cause, i.is_critical_path,
                i.created_at, i.updated_at
            FROM issues i
        """)
        return cur.fetchall()


def fetch_project_names(conn, project_ids: List[str]) -> Dict[str, Dict]:
    """根據 project_ids 取得專案名稱和代碼"""
    if not project_ids:
        return {}
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id, name, code FROM projects WHERE id = ANY(%s)",
            (project_ids,)
        )
        rows = cur.fetchall()
        return {row["id"]: {"name": row["name"], "code": row["code"]} for row in rows}


# ============================================
# 日期格式轉換
# ============================================
def to_iso_string(dt) -> Optional[str]:
    if dt is None:
        return None
    if isinstance(dt, datetime):
        return dt.isoformat() + "Z" if dt.tzinfo is None else dt.isoformat()
    return str(dt)


# ============================================
# 資料轉換為 External Item
# ============================================
def transform_project(project: Dict) -> Dict:
    return {
        "id": f"project-{project['id']}",
        "properties": {
            "itemType": "project",
            "title": project["name"],
            "description": project.get("description") or "",
            "url": f"{APP_BASE_URL}/projects/{project['id']}",
            "lastModifiedDateTime": to_iso_string(project.get("updated_at")),
            "createdDateTime": to_iso_string(project.get("created_at")),
            "projectCode": project["code"],
            "projectName": project["name"],
            "projectId": project["id"],
            "status": project["status"],
            "priority": project.get("priority") or "medium",
            "progress": project.get("progress") or 0,
            "startDate": to_iso_string(project.get("start_date")),
            "endDate": to_iso_string(project.get("end_date")),
            "category": project.get("category_label") or "",
            "managers": project.get("managers") or [],
            "teamMembers": project.get("team_members") or [],
            "tags": project.get("tags") or [],
            "budget": float(project["budget"]) if project.get("budget") else None,
            "budgetUsed": float(project["budget_used"]) if project.get("budget_used") else None,
        },
        "content": {
            "type": "text",
            "value": "\n".join([
                f"專案名稱: {project['name']}",
                f"專案代碼: {project['code']}",
                f"狀態: {project['status']}",
                f"進度: {project.get('progress', 0)}%",
                f"優先級: {project.get('priority', 'medium')}",
                project.get("description") or "",
            ]),
        },
        "acl": [
            {"type": "everyone", "value": "everyone", "accessType": "grant"}
        ],
    }


def transform_milestone(milestone: Dict) -> Dict:
    return {
        "id": f"milestone-{milestone['id']}",
        "properties": {
            "itemType": "milestone",
            "title": milestone["title"],
            "description": milestone.get("description") or "",
            "url": f"{APP_BASE_URL}/projects/{milestone['project_id']}/milestones/{milestone['id']}",
            "lastModifiedDateTime": to_iso_string(milestone.get("updated_at")),
            "createdDateTime": to_iso_string(milestone.get("created_at")),
            "projectCode": milestone.get("project_code") or "",
            "projectName": milestone.get("project_name") or "",
            "projectId": milestone["project_id"],
            "status": milestone["status"],
            "priority": milestone.get("priority") or "medium",
            "dueDate": to_iso_string(milestone.get("due_date")),
            "category": milestone.get("category") or "",
            "phase": milestone.get("phase") or "",
            "owners": [milestone["assigned_to"]] if milestone.get("assigned_to") else [],
            "isCriticalPath": milestone.get("is_critical_path") or False,
        },
        "content": {
            "type": "text",
            "value": "\n".join([
                f"里程碑: {milestone['title']}",
                f"專案: {milestone.get('project_name', '')} ({milestone.get('project_code', '')})",
                f"狀態: {milestone['status']}",
                f"截止日期: {milestone.get('due_date', 'N/A')}",
                f"階段: {milestone.get('phase') or 'N/A'}",
                milestone.get("description") or "",
            ]),
        },
        "acl": [
            {"type": "everyone", "value": "everyone", "accessType": "grant"}
        ],
    }


def transform_risk(risk: Dict, project_map: Dict[str, Dict]) -> Dict:
    project_ids = risk.get("project_ids") or []
    project_names = ", ".join([project_map.get(pid, {}).get("name", "") for pid in project_ids if pid in project_map])
    project_codes = ", ".join([project_map.get(pid, {}).get("code", "") for pid in project_ids if pid in project_map])
    
    return {
        "id": f"risk-{risk['id']}",
        "properties": {
            "itemType": "risk",
            "title": risk["title"],
            "description": risk.get("description") or "",
            "url": f"{APP_BASE_URL}/risks/{risk['id']}",
            "lastModifiedDateTime": to_iso_string(risk.get("updated_at")),
            "createdDateTime": to_iso_string(risk.get("created_at")),
            "projectCode": project_codes,
            "projectName": project_names,
            "projectId": project_ids[0] if project_ids else "",
            "status": risk["status"],
            "dueDate": to_iso_string(risk.get("deadline")),
            "probability": risk["probability"],
            "impact": risk["impact"],
            "owners": risk.get("owners") or [],
            "isCriticalPath": risk.get("is_critical_path") or False,
            "mitigation": risk.get("mitigation") or "",
        },
        "content": {
            "type": "text",
            "value": "\n".join([
                f"風險: {risk['title']}",
                f"專案: {project_names}",
                f"狀態: {risk['status']}",
                f"機率: {risk['probability']} | 影響: {risk['impact']}",
                f"截止日期: {risk.get('deadline', 'N/A')}",
                f"緩解措施: {risk.get('mitigation') or 'N/A'}",
                risk.get("description") or "",
            ]),
        },
        "acl": [
            {"type": "everyone", "value": "everyone", "accessType": "grant"}
        ],
    }


def transform_issue(issue: Dict, project_map: Dict[str, Dict]) -> Dict:
    project_ids = issue.get("project_ids") or []
    project_names = ", ".join([project_map.get(pid, {}).get("name", "") for pid in project_ids if pid in project_map])
    project_codes = ", ".join([project_map.get(pid, {}).get("code", "") for pid in project_ids if pid in project_map])
    
    return {
        "id": f"issue-{issue['id']}",
        "properties": {
            "itemType": "issue",
            "title": issue["title"],
            "description": issue.get("description") or "",
            "url": f"{APP_BASE_URL}/issues/{issue['id']}",
            "lastModifiedDateTime": to_iso_string(issue.get("updated_at")),
            "createdDateTime": to_iso_string(issue.get("created_at")),
            "projectCode": project_codes,
            "projectName": project_names,
            "projectId": project_ids[0] if project_ids else "",
            "status": issue["status"],
            "dueDate": to_iso_string(issue.get("due_date")),
            "severity": issue.get("severity") or "medium",
            "owners": issue.get("owners") or [],
            "isCriticalPath": issue.get("is_critical_path") or False,
            "rootCause": issue.get("root_cause") or "",
        },
        "content": {
            "type": "text",
            "value": "\n".join([
                f"問題: {issue['title']}",
                f"專案: {project_names}",
                f"狀態: {issue['status']}",
                f"嚴重程度: {issue.get('severity', 'medium')}",
                f"截止日期: {issue.get('due_date', 'N/A')}",
                f"根本原因: {issue.get('root_cause') or 'N/A'}",
                issue.get("description") or "",
            ]),
        },
        "acl": [
            {"type": "everyone", "value": "everyone", "accessType": "grant"}
        ],
    }


# ============================================
# 上傳到 Microsoft Graph
# ============================================
def upsert_external_item(token: str, item: Dict) -> bool:
    """新增或更新 External Item"""
    url = f"{GRAPH_API_BASE}/external/connections/{CONNECTION_ID}/items/{item['id']}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    
    response = requests.put(url, headers=headers, json=item)
    
    if response.ok:
        return True
    else:
        print(f"   ❌ 上傳失敗 {item['id']}: {response.status_code}")
        try:
            error_data = response.json()
            print(f"      錯誤: {json.dumps(error_data, indent=2, ensure_ascii=False)}")
        except:
            print(f"      回應: {response.text[:200]}")
        return False


def delete_external_item(token: str, item_id: str) -> bool:
    """刪除 External Item"""
    url = f"{GRAPH_API_BASE}/external/connections/{CONNECTION_ID}/items/{item_id}"
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.delete(url, headers=headers)
    return response.ok or response.status_code == 404


# ============================================
# 主要同步邏輯
# ============================================
def sync_all_data():
    print("=" * 60)
    print("步驟 4：同步資料到 Microsoft Graph Connector")
    print("=" * 60)
    
    # 取得 Token
    print("\n🔑 取得 Access Token...")
    token = get_access_token()
    print("✅ Token 取得成功")
    
    # 連接資料庫
    print("\n📦 連接資料庫...")
    try:
        conn = get_db_connection()
        print("✅ 資料庫連接成功")
    except Exception as e:
        print(f"❌ 資料庫連接失敗: {e}")
        print("\n請確認 DATABASE_CONFIG 設定正確，或使用環境變數：")
        print("  DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD")
        return
    
    results = {"success": 0, "failed": 0, "errors": []}
    
    try:
        # 1. 同步 Projects
        print("\n📁 同步 Projects...")
        projects = fetch_projects(conn)
        print(f"   找到 {len(projects)} 個專案")
        
        for project in projects:
            item = transform_project(project)
            if upsert_external_item(token, item):
                results["success"] += 1
                print(f"   ✅ {item['id']}")
            else:
                results["failed"] += 1
                results["errors"].append(item["id"])
        
        # 2. 同步 Milestones
        print("\n📌 同步 Milestones...")
        milestones = fetch_milestones(conn)
        print(f"   找到 {len(milestones)} 個里程碑")
        
        for milestone in milestones:
            item = transform_milestone(milestone)
            if upsert_external_item(token, item):
                results["success"] += 1
                print(f"   ✅ {item['id']}")
            else:
                results["failed"] += 1
                results["errors"].append(item["id"])
        
        # 3. 同步 Risks
        print("\n⚠️ 同步 Risks...")
        risks = fetch_risks(conn)
        print(f"   找到 {len(risks)} 個風險")
        
        # 取得相關專案資訊
        all_risk_project_ids = []
        for risk in risks:
            all_risk_project_ids.extend(risk.get("project_ids") or [])
        project_map = fetch_project_names(conn, list(set(all_risk_project_ids)))
        
        for risk in risks:
            item = transform_risk(risk, project_map)
            if upsert_external_item(token, item):
                results["success"] += 1
                print(f"   ✅ {item['id']}")
            else:
                results["failed"] += 1
                results["errors"].append(item["id"])
        
        # 4. 同步 Issues
        print("\n🔴 同步 Issues...")
        issues = fetch_issues(conn)
        print(f"   找到 {len(issues)} 個問題")
        
        # 取得相關專案資訊
        all_issue_project_ids = []
        for issue in issues:
            all_issue_project_ids.extend(issue.get("project_ids") or [])
        project_map = fetch_project_names(conn, list(set(all_issue_project_ids)))
        
        for issue in issues:
            item = transform_issue(issue, project_map)
            if upsert_external_item(token, item):
                results["success"] += 1
                print(f"   ✅ {item['id']}")
            else:
                results["failed"] += 1
                results["errors"].append(item["id"])
        
    finally:
        conn.close()
    
    # 結果摘要
    print("\n" + "=" * 60)
    print("📊 同步結果摘要")
    print("=" * 60)
    print(f"   ✅ 成功: {results['success']}")
    print(f"   ❌ 失敗: {results['failed']}")
    
    if results["errors"]:
        print(f"\n   失敗項目:")
        for err in results["errors"][:10]:
            print(f"      - {err}")
        if len(results["errors"]) > 10:
            print(f"      ... 還有 {len(results['errors']) - 10} 個")
    
    print("\n🎉 同步完成！")
    print("   資料現在可以在 Microsoft Search 和 Copilot 中搜尋")


# ============================================
# 測試模式（不需要資料庫）
# ============================================
def sync_test_data():
    """使用測試資料進行同步，不需要資料庫連線"""
    print("=" * 60)
    print("步驟 4：同步測試資料到 Microsoft Graph Connector")
    print("=" * 60)
    
    token = get_access_token()
    print("✅ Token 取得成功")
    
    # 測試資料
    test_items = [
        {
            "id": "project-test-001",
            "properties": {
                "itemType": "project",
                "title": "測試專案 Alpha",
                "description": "這是一個測試專案",
                "url": f"{APP_BASE_URL}/projects/test-001",
                "projectCode": "ALPHA-001",
                "projectName": "測試專案 Alpha",
                "projectId": "test-001",
                "status": "C2",
                "priority": "high",
                "progress": 45,
                "category": "AI專案",
            },
            "content": {
                "type": "text",
                "value": "測試專案 Alpha\n專案代碼: ALPHA-001\n狀態: C2\n進度: 45%",
            },
            "acl": [{"type": "everyone", "value": "everyone", "accessType": "grant"}],
        },
        {
            "id": "milestone-test-001",
            "properties": {
                "itemType": "milestone",
                "title": "Alpha 里程碑 1",
                "description": "第一個里程碑",
                "url": f"{APP_BASE_URL}/projects/test-001/milestones/test-m-001",
                "projectCode": "ALPHA-001",
                "projectName": "測試專案 Alpha",
                "projectId": "test-001",
                "status": "in_progress",
                "priority": "high",
                "dueDate": "2025-01-15T00:00:00Z",
                "phase": "C2",
            },
            "content": {
                "type": "text",
                "value": "Alpha 里程碑 1\n專案: 測試專案 Alpha\n狀態: in_progress",
            },
            "acl": [{"type": "everyone", "value": "everyone", "accessType": "grant"}],
        },
        {
            "id": "risk-test-001",
            "properties": {
                "itemType": "risk",
                "title": "測試風險：時程延遲",
                "description": "可能因為資源不足導致時程延遲",
                "url": f"{APP_BASE_URL}/risks/test-r-001",
                "projectCode": "ALPHA-001",
                "projectName": "測試專案 Alpha",
                "status": "open",
                "probability": "medium",
                "impact": "high",
                "mitigation": "增加人力資源",
            },
            "content": {
                "type": "text",
                "value": "風險: 時程延遲\n機率: medium\n影響: high",
            },
            "acl": [{"type": "everyone", "value": "everyone", "accessType": "grant"}],
        },
        {
            "id": "issue-test-001",
            "properties": {
                "itemType": "issue",
                "title": "測試問題：API 效能問題",
                "description": "API 回應時間過長",
                "url": f"{APP_BASE_URL}/issues/test-i-001",
                "projectCode": "ALPHA-001",
                "projectName": "測試專案 Alpha",
                "status": "open",
                "severity": "high",
                "rootCause": "資料庫查詢未優化",
            },
            "content": {
                "type": "text",
                "value": "問題: API 效能問題\n嚴重程度: high\n狀態: open",
            },
            "acl": [{"type": "everyone", "value": "everyone", "accessType": "grant"}],
        },
    ]
    
    print(f"\n📤 上傳 {len(test_items)} 個測試項目...")
    
    success = 0
    for item in test_items:
        if upsert_external_item(token, item):
            print(f"   ✅ {item['id']}")
            success += 1
        else:
            print(f"   ❌ {item['id']}")
    
    print(f"\n🎉 測試同步完成！成功: {success}/{len(test_items)}")
    print("   你可以到 Microsoft Search 搜尋 '測試專案' 來驗證")


# ============================================
# 執行
# ============================================
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        # 測試模式：使用假資料
        sync_test_data()
    else:
        # 正式模式：從資料庫同步
        sync_all_data()