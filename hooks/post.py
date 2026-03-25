import json
import os

from sharepoint import get_graph_client, upload_files


WORKSPACE = "/home/claude/workspace"
SKIP_DIRS = {".venv"}
SKIP_FILES = {"pyproject.toml", "uv.lock"}

with open("/tmp/sharepoint_context.json") as f:
    context = json.load(f)

drive_id = context["drive_id"]
folder_path = context["folder_path"]

files = []
for root, dirs, filenames in os.walk(WORKSPACE):
    dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
    for file_name in filenames:
        if file_name not in SKIP_FILES:
            local_path = os.path.join(root, file_name)
            relative_path = os.path.relpath(local_path, WORKSPACE)
            files.append((local_path, relative_path))

client = get_graph_client()
upload_files(client, drive_id, folder_path, files)

print(f"All files uploaded to {folder_path}")
