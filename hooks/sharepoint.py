import asyncio
import os

from azure.identity import OnBehalfOfCredential
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphRequestAdapter, GraphServiceClient
from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.folder import Folder
from msgraph.generated.models.o_data_errors.o_data_error import ODataError


GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]


def resolve_drive_id(
    client: GraphServiceClient,
    site_url: str,
    drive_name: str | None = None,
) -> str:
    """Resolve a SharePoint site URL and optional drive name to a drive ID.

    The site_url should be a full URL like
    ``https://example.sharepoint.com/sites/marketing``.  If drive_name is
    omitted the default document library (``"Documents"``) is used.
    """
    return asyncio.run(_resolve_drive_id(client, site_url, drive_name))


def resolve_and_create_folder(
    client: GraphServiceClient,
    site_url: str,
    drive_name: str | None,
    parent_path: str | None,
    folder_name: str | None,
    download_dest: str | None = None,
) -> tuple[str, str, bool, list[str]]:
    """Resolve a drive and create a folder in a single event loop.

    When download_dest is provided and the folder already exists, its
    contents are downloaded into that directory.

    Returns (drive_id, folder_item_id, created, downloaded_files).
    """
    return asyncio.run(
        _resolve_and_create_folder(
            client, site_url, drive_name, parent_path, folder_name, download_dest
        )
    )


async def _resolve_and_create_folder(
    client: GraphServiceClient,
    site_url: str,
    drive_name: str | None,
    parent_path: str | None,
    folder_name: str | None,
    download_dest: str | None = None,
) -> tuple[str, str, bool, list[str]]:
    drive_id = await _resolve_drive_id(client, site_url, drive_name)
    folder_item_id, created = await _create_folder(
        client, drive_id, parent_path, folder_name
    )
    downloaded: list[str] = []
    if not created and download_dest:
        downloaded = await _download_files(
            client, drive_id, folder_item_id, download_dest, ""
        )
    return drive_id, folder_item_id, created, downloaded


async def _resolve_drive_id(
    client: GraphServiceClient,
    site_url: str,
    drive_name: str | None = None,
) -> str:
    from urllib.parse import urlparse

    parsed = urlparse(site_url)
    hostname = parsed.hostname
    site_path = parsed.path.rstrip("/")
    if not hostname or not site_path:
        raise ValueError(
            f"SHAREPOINT_SITE_URL must include hostname and path (got {site_url!r})"
        )

    site = await client.sites.by_site_id(f"{hostname}:{site_path}:").get()
    if site is None or site.id is None:
        raise RuntimeError(f"Could not resolve site for {site_url}")

    drives = await client.sites.by_site_id(site.id).drives.get()
    if drives is None or drives.value is None:
        raise RuntimeError(f"No drives found for site {site_url}")

    target_name = drive_name or "Documents"
    for drive in drives.value:
        if drive.name == target_name:
            if drive.id is None:
                raise RuntimeError(f"Drive '{target_name}' has no ID")
            return drive.id

    available = [d.name for d in drives.value if d.name]
    raise RuntimeError(
        f"Drive '{target_name}' not found in site {site_url}. "
        f"Available drives: {available}"
    )


def _item_ref(path: str | None) -> str:
    """Return a Graph API drive-item reference for the given path.

    Returns ``"root"`` for the drive root (when path is None or empty),
    or ``"root:/{path}:"`` for a subfolder.
    """
    if path:
        return f"root:/{path}:"
    return "root"


def get_graph_client() -> GraphServiceClient:
    tenant_id = os.environ["AZURE_TENANT_ID"]
    client_id = os.environ["AZURE_CLIENT_ID"]
    client_secret = os.environ["AZURE_CLIENT_SECRET"]
    user_token = os.environ["USER_ACCESS_TOKEN"]

    credential = OnBehalfOfCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        user_assertion=user_token,
    )

    auth_provider = AzureIdentityAuthenticationProvider(
        credential, scopes=GRAPH_SCOPES
    )
    adapter = GraphRequestAdapter(auth_provider)
    return GraphServiceClient(request_adapter=adapter)


def create_folder(
    client: GraphServiceClient,
    drive_id: str,
    parent_path: str | None,
    folder_name: str | None,
) -> tuple[str, bool]:
    """Create a folder (and any intermediate segments) under parent_path.

    Returns a tuple of (item_id, created) where created is False if the
    final folder already existed.
    """
    return asyncio.run(_create_folder(client, drive_id, parent_path, folder_name))


async def _create_folder(
    client: GraphServiceClient,
    drive_id: str,
    parent_path: str | None,
    folder_name: str | None,
) -> tuple[str, bool]:
    full_path = "/".join(p for p in (parent_path, folder_name) if p)

    if not full_path:
        root = await (
            client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id("root")
            .get()
        )
        if root is None or root.id is None:
            raise RuntimeError("Drive root not found")
        return root.id, False

    try:
        existing = await (
            client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(_item_ref(full_path))
            .get()
        )
        if existing is not None and existing.id is not None:
            return existing.id, False
    except ODataError:
        pass

    current_path: str | None = None
    created = False
    for segment in full_path.split("/"):
        body = DriveItem(
            name=segment,
            folder=Folder(),
            additional_data={"@microsoft.graph.conflictBehavior": "replace"},
        )

        try:
            result = await (
                client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(_item_ref(current_path))
                .children.post(body, request_configuration=RequestConfiguration())
            )
            created = True
        except ODataError as e:
            if e.response_status_code == 409:
                child_path = "/".join(p for p in (current_path, segment) if p)
                result = await (
                    client.drives.by_drive_id(drive_id)
                    .items.by_drive_item_id(_item_ref(child_path))
                    .get()
                )
            else:
                raise
        if result is None or result.id is None:
            raise RuntimeError(f"Failed to create folder '{segment}' under '{current_path}'")
        current_path = "/".join(p for p in (current_path, segment) if p)

    return result.id, created


async def _download_files(
    client: GraphServiceClient,
    drive_id: str,
    folder_item_id: str,
    dest_dir: str,
    relative_prefix: str,
) -> list[str]:
    downloaded: list[str] = []

    children = await (
        client.drives.by_drive_id(drive_id)
        .items.by_drive_item_id(folder_item_id)
        .children.get()
    )
    if children is None or children.value is None:
        return downloaded

    for item in children.value:
        if item.name is None or item.id is None:
            continue

        relative_path = f"{relative_prefix}/{item.name}" if relative_prefix else item.name

        if item.folder is not None:
            sub = await _download_files(
                client, drive_id, item.id, dest_dir, relative_path
            )
            downloaded.extend(sub)
        else:
            local_path = os.path.join(dest_dir, relative_path)
            os.makedirs(os.path.dirname(local_path), exist_ok=True)

            content = await (
                client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item.id)
                .content.get()
            )
            if content is not None:
                with open(local_path, "wb") as f:
                    f.write(content)
                downloaded.append(relative_path)
                print(f"Downloaded: {relative_path}")

    return downloaded


def upload_files(
    client: GraphServiceClient,
    drive_id: str,
    folder_path: str,
    files: list[tuple[str, str]],
) -> None:
    """Upload files to a SharePoint folder.

    Each entry in files is a (local_path, relative_path) tuple where
    relative_path preserves the desired folder structure under folder_path.
    """
    asyncio.run(_upload_files(client, drive_id, folder_path, files))


async def _upload_files(
    client: GraphServiceClient,
    drive_id: str,
    folder_path: str,
    files: list[tuple[str, str]],
) -> None:
    for local_path, relative_path in files:
        upload_path = f"{folder_path}/{relative_path}"

        with open(local_path, "rb") as f:
            content = f.read()

        await (
            client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(f"root:/{upload_path}:")
            .content.put(content)
        )
        print(f"Uploaded: {relative_path}")
