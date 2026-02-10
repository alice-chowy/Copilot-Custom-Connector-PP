"""
查詢已建立的 Connections
"""
import requests
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from config import CONFIG

def get_access_token():
    url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    payload = {
        "client_id": CONFIG["client_id"],
        "client_secret": CONFIG["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    response = requests.post(url, data=payload, verify=False)
    return response.json()["access_token"]

def list_connections(token):
    url = "https://graph.microsoft.com/v1.0/external/connections"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers, verify=False)
    return response.json()

def main():
    token = get_access_token()
    connections = list_connections(token)
    
    print("已建立的 Connections：")
    print("=" * 50)
    
    for conn in connections.get("value", []):
        print(f"ID: {conn.get('id')}")
        print(f"Name: {conn.get('name')}")
        print(f"State: {conn.get('state')}")
        print("-" * 50)

if __name__ == "__main__":
    main()