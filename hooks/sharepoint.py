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
