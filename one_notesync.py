"""
OneNote → Markdown Sync Script
Uses Microsoft Graph API (delegated auth) to pull all your OneNote notebooks
and save them as .md files, keeping your VS Code workspace in sync.

Requirements:
    pip install msal requests html2text

Usage:
    1. Fill in CLIENT_ID below with your Azure App Registration client ID
    2. Set OUTPUT_DIR to your VS Code workspace folder
    3. Run: python onenote_sync.py
    4. On first run, a browser window opens for Microsoft login (one-time)
    5. Schedule this script to run periodically for continuous sync
"""

import os
import re
import json
import time
import html2text
import requests
import msal

# ── CONFIG ────────────────────────────────────────────────────────────────────
CLIENT_ID  = "YOUR_AZURE_APP_CLIENT_ID"   # paste your App Registration client ID
OUTPUT_DIR = "./onenote-notes"             # folder opened in VS Code
TOKEN_CACHE_FILE = ".onenote_token_cache.json"
# ─────────────────────────────────────────────────────────────────────────────

SCOPES    = ["Notes.Read", "offline_access"]
AUTHORITY = "https://login.microsoftonline.com/consumers"
GRAPH_URL = "https://graph.microsoft.com/v1.0/me/onenote"


# ── AUTH ──────────────────────────────────────────────────────────────────────

def get_token():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE).read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result:
        print("Opening browser for Microsoft login...")
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(flow["message"])  # prints: "Go to https://microsoft.com/devicelogin and enter code XXXX"
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description')}")

    # persist token cache so future runs skip login
    with open(TOKEN_CACHE_FILE, "w") as f:
        f.write(cache.serialize())

    return result["access_token"]


# ── GRAPH API HELPERS ─────────────────────────────────────────────────────────

def graph_get(token, url, params=None):
    headers = {"Authorization": f"Bearer {token}"}
    items, next_link = [], url
    while next_link:
        r = requests.get(next_link, headers=headers, params=params)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", [data]))
        next_link = data.get("@odata.nextLink")
        params = None  # only use params on first call
    return items


def get_page_content(token, page_id):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(f"{GRAPH_URL}/pages/{page_id}/content", headers=headers)
    r.raise_for_status()
    return r.text  # returns HTML


# ── CONVERSION ────────────────────────────────────────────────────────────────

def html_to_markdown(html_content):
    converter = html2text.HTML2Text()
    converter.ignore_links = False
    converter.ignore_images = False
    converter.body_width = 0       # no line wrapping
    converter.protect_links = True
    return converter.handle(html_content)


def safe_filename(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()


# ── SYNC ──────────────────────────────────────────────────────────────────────

def sync_all(token):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    synced, skipped, errors = 0, 0, 0

    print("Fetching notebooks...")
    notebooks = graph_get(token, f"{GRAPH_URL}/notebooks")

    for nb in notebooks:
        nb_name = safe_filename(nb["displayName"])
        nb_dir  = os.path.join(OUTPUT_DIR, nb_name)
        os.makedirs(nb_dir, exist_ok=True)
        print(f"\n📓 {nb_name}")

        sections = graph_get(token, f"{GRAPH_URL}/notebooks/{nb['id']}/sections")

        for sec in sections:
            sec_name = safe_filename(sec["displayName"])
            sec_dir  = os.path.join(nb_dir, sec_name)
            os.makedirs(sec_dir, exist_ok=True)
            print(f"  📄 {sec_name}")

            pages = graph_get(token, f"{GRAPH_URL}/sections/{sec['id']}/pages",
                              params={"$select": "id,title,lastModifiedDateTime"})

            for page in pages:
                page_title    = safe_filename(page.get("title") or "Untitled")
                page_modified = page.get("lastModifiedDateTime", "")
                md_path       = os.path.join(sec_dir, f"{page_title}.md")

                # Skip if file exists and hasn't changed (check modified time in metadata)
                meta_path = md_path + ".meta"
                if os.path.exists(md_path) and os.path.exists(meta_path):
                    with open(meta_path) as f:
                        if f.read().strip() == page_modified:
                            skipped += 1
                            continue

                try:
                    html = get_page_content(token, page["id"])
                    md   = f"# {page.get('title', 'Untitled')}\n\n" + html_to_markdown(html)

                    with open(md_path, "w", encoding="utf-8") as f:
                        f.write(md)
                    with open(meta_path, "w") as f:
                        f.write(page_modified)

                    print(f"    ✅ {page_title}.md")
                    synced += 1
                    time.sleep(0.3)  # be gentle with the API rate limit

                except Exception as e:
                    print(f"    ❌ {page_title}: {e}")
                    errors += 1

    print(f"\nDone. {synced} synced, {skipped} unchanged, {errors} errors.")
    print(f"Notes saved to: {os.path.abspath(OUTPUT_DIR)}")


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if CLIENT_ID == "YOUR_AZURE_APP_CLIENT_ID":
        print("ERROR: Set your CLIENT_ID in the CONFIG section at the top of this script.")
        print("Get it from: portal.azure.com → App registrations")
        exit(1)

    token = get_token()
    sync_all(token)

