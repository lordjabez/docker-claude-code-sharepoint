import json
import os
import sys

from sharepoint import create_folder, get_graph_client


input_data = json.loads(sys.argv[1])
path = input_data["path"]

drive_id = os.environ["SHAREPOINT_DRIVE_ID"]
base_path = os.environ.get("SHAREPOINT_BASE_PATH", "")

folder_name = f"{base_path}/{path}" if base_path else path

client = get_graph_client()
folder_item_id, created = create_folder(client, drive_id, base_path, path)

context = {
    "folder_item_id": folder_item_id,
    "folder_path": folder_name,
    "drive_id": drive_id,
}

with open("/tmp/sharepoint_context.json", "w") as f:
    json.dump(context, f)

status = "Created" if created else "Using existing"
print(f"{status} SharePoint folder: {folder_name}")
