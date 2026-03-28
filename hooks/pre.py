import json
import os
import sys

from sharepoint import get_graph_client, resolve_and_create_folder


input_data = json.loads(sys.argv[1])
path = input_data.get("path")

site_url = os.environ["SHAREPOINT_SITE_URL"]
drive_name = os.environ.get("SHAREPOINT_DRIVE_NAME")
base_path = os.environ.get("SHAREPOINT_BASE_PATH")

client = get_graph_client()
drive_id, folder_item_id, created = resolve_and_create_folder(
    client, site_url, drive_name, base_path, path
)

folder_path = "/".join(p for p in (base_path, path) if p)

context = {
    "folder_item_id": folder_item_id,
    "folder_path": folder_path,
    "drive_id": drive_id,
}

with open("/tmp/sharepoint_context.json", "w") as f:
    json.dump(context, f)

print(f"Using SharePoint site: {site_url}")

status = "Created" if created else "Using existing"
print(f"{status} SharePoint folder: {folder_path}")
