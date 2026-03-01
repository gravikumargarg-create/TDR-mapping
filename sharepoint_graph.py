"""
Microsoft Graph helpers for TDR app: list and download Excel files from a SharePoint folder.
Uses app-only auth (client credentials). Credentials from env or Streamlit secrets.
"""
import os
import urllib.parse

# Optional: use requests if available (Streamlit has it), else fallback
try:
    import requests
except ImportError:
    requests = None

# MSAL for token
try:
    import msal
except ImportError:
    msal = None

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
# SharePoint site and folder path for TDR (R2 Data)
SHAREPOINT_HOST = "amdocs.sharepoint.com"
SITE_PATH = "/sites/USCCTesting_Offshore"
# Folder path relative to the default document library (Shared Documents)
FOLDER_PATH = "Shared Documents/Release & OffCycle/UTMO - Migration (Data creation)/Data Creation/R2 Data"
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


def _get_secrets():
    """Tenant ID, Client ID, Client secret from env or Streamlit secrets."""
    try:
        import streamlit as st
        s = getattr(st, "secrets", None) or {}
        return (
            os.environ.get("AZURE_TENANT_ID") or s.get("AZURE_TENANT_ID"),
            os.environ.get("AZURE_CLIENT_ID") or s.get("AZURE_CLIENT_ID"),
            os.environ.get("AZURE_CLIENT_SECRET") or s.get("AZURE_CLIENT_SECRET"),
        )
    except Exception:
        return (
            os.environ.get("AZURE_TENANT_ID"),
            os.environ.get("AZURE_CLIENT_ID"),
            os.environ.get("AZURE_CLIENT_SECRET"),
        )


def has_sharepoint_credentials():
    """True if all three Azure credentials are set."""
    tenant, client, secret = _get_secrets()
    return bool(tenant and client and secret)


def get_token():
    """Get access token for Microsoft Graph (app-only)."""
    if not msal:
        return None
    tenant_id, client_id, client_secret = _get_secrets()
    if not all([tenant_id, client_id, client_secret]):
        return None
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token") if result else None


def _graph_get(token, path, params=None):
    if not requests or not token:
        return None
    url = f"{GRAPH_BASE}{path}" if path.startswith("/") else f"{GRAPH_BASE}/{path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params or {}, timeout=30)
    if r.status_code != 200:
        return None
    return r.json()


def _graph_get_bytes(token, path):
    if not requests or not token:
        return None
    url = f"{GRAPH_BASE}{path}" if path.startswith("/") else f"{GRAPH_BASE}/{path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code != 200:
        return None
    return r.content


def list_tdr_excel_files(token):
    """
    List Excel files (.xlsx, .xlsm) in the configured SharePoint folder.
    Returns list of dicts: [{"id": "...", "name": "file.xlsx", "drive_id": "..."}, ...]
    """
    if not token:
        return []
    # Site by path: /sites/{hostname}:{server-relative-path}
    site = _graph_get(token, f"/sites/{SHAREPOINT_HOST}:{SITE_PATH}")
    if not site or "id" not in site:
        return []
    site_id = site["id"]
    # Default document library (drive)
    drive_res = _graph_get(token, f"/sites/{site_id}/drive")
    if not drive_res or "id" not in drive_res:
        return []
    drive_id = drive_res["id"]
    # Folder by path (encode path segments)
    path_encoded = urllib.parse.quote(FOLDER_PATH, safe="/")
    folder = _graph_get(token, f"/drives/{drive_id}/root:/{path_encoded}")
    if not folder or "id" not in folder:
        return []
    folder_id = folder["id"]
    # Children
    children = _graph_get(token, f"/drives/{drive_id}/items/{folder_id}/children")
    if not children or "value" not in children:
        return []
    out = []
    for item in children.get("value", []):
        name = (item.get("name") or "").lower()
        if name.endswith(EXCEL_EXTENSIONS) and "file" in (item.get("file") or {}):
            out.append({
                "id": item["id"],
                "name": item.get("name", "?"),
                "drive_id": drive_id,
            })
    return sorted(out, key=lambda x: x["name"])


def download_file_content(token, drive_id, item_id):
    """Download file content as bytes."""
    if not token:
        return None
    return _graph_get_bytes(token, f"/drives/{drive_id}/items/{item_id}/content")
