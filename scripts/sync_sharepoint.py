#!/usr/bin/env python3
"""Pull files from a SharePoint document library folder into ./docs via Microsoft Graph."""
import os
import sys
import pathlib
import requests

TENANT = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
SITE_ID = os.environ["SHAREPOINT_SITE_ID"]
DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]

# Folder inside the library to sync, e.g. "BA-Uploads". Empty string = root.
SOURCE_FOLDER = os.environ.get("SHAREPOINT_FOLDER", "")
# Local destination in the repo.
DEST_DIR = pathlib.Path(os.environ.get("DEST_DIR", "docs"))

GRAPH = "https://graph.microsoft.com/v1.0"


def get_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=30)
    if not r.ok:
        print(f"Token request failed with status {r.status_code}")
        print(f"Response body: {r.text}")
        print(f"Tenant ID length: {len(TENANT)} (expect 36)")
        print(f"Client ID length: {len(CLIENT_ID)} (expect 36)")
        print(f"Client secret length: {len(CLIENT_SECRET)} (varies, usually 40)")
        r.raise_for_status()
    return r.json()["access_token"]


def list_children(token: str, item_path: str):
    headers = {"Authorization": f"Bearer {token}"}
    if item_path:
        url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{item_path}:/children"
    else:
        url = f"{GRAPH}/drives/{DRIVE_ID}/root/children"
    items = []
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        if not r.ok:
            print(f"Request failed: {r.status_code}")
            print(f"URL: {url}")
            print(f"Response: {r.text}")
        r.raise_for_status()
        body = r.json()
        items.extend(body.get("value", []))
        url = body.get("@odata.nextLink")
    return items


def download_file(token: str, item: dict, local_path: pathlib.Path):
    headers = {"Authorization": f"Bearer {token}"}
    download_url = item.get("@microsoft.graph.downloadUrl")
    if not download_url:
        # Fallback: fetch the item to get a fresh download URL.
        r = requests.get(
            f"{GRAPH}/drives/{DRIVE_ID}/items/{item['id']}",
            headers=headers, timeout=30,
        )
        r.raise_for_status()
        download_url = r.json()["@microsoft.graph.downloadUrl"]
    # downloadUrl is pre-authenticated; do NOT send the auth header.
    r = requests.get(download_url, timeout=120)
    r.raise_for_status()
    local_path.parent.mkdir(parents=True, exist_ok=True)
    local_path.write_bytes(r.content)
    print(f"  saved {local_path}")


def walk(token: str, remote_folder: str, local_base: pathlib.Path):
    """Recursively mirror remote_folder into local_base."""
    for item in list_children(token, remote_folder):
        name = item["name"]
        if "folder" in item:
            child_remote = f"{remote_folder}/{name}".lstrip("/")
            walk(token, child_remote, local_base / name)
        elif "file" in item:
            download_file(token, item, local_base / name)


def main():
    token = get_token()
    print(f"Syncing SharePoint folder '{SOURCE_FOLDER or '<root>'}' -> {DEST_DIR}/")
    # Clean destination so deletions on SharePoint propagate.
    if DEST_DIR.exists():
        import shutil
        shutil.rmtree(DEST_DIR)
    DEST_DIR.mkdir(parents=True, exist_ok=True)
    walk(token, SOURCE_FOLDER, DEST_DIR)
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
