import sys
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "hooks"))

from sharepoint import (
    _item_ref,
    create_folder,
    get_graph_client,
    resolve_and_create_folder,
    resolve_drive_id,
    upload_files,
)


class TestItemRef:
    def test_with_path(self) -> None:
        assert _item_ref("base/folder") == "root:/base/folder:"

    def test_with_none(self) -> None:
        assert _item_ref(None) == "root"

    def test_with_empty_string(self) -> None:
        assert _item_ref("") == "root"


class TestGetGraphClient:
    @patch.dict(
        "os.environ",
        {
            "AZURE_TENANT_ID": "tenant-123",
            "AZURE_CLIENT_ID": "client-456",
            "AZURE_CLIENT_SECRET": "secret-789",
            "USER_ACCESS_TOKEN": "token-abc",
        },
    )
    @patch("sharepoint.GraphServiceClient")
    @patch("sharepoint.GraphRequestAdapter")
    @patch("sharepoint.AzureIdentityAuthenticationProvider")
    @patch("sharepoint.OnBehalfOfCredential")
    def test_returns_graph_client(
        self, mock_cred, mock_auth, mock_adapter, mock_client
    ) -> None:
        client = get_graph_client()
        mock_cred.assert_called_once_with(
            tenant_id="tenant-123",
            client_id="client-456",
            client_secret="secret-789",
            user_assertion="token-abc",
        )
        assert client is mock_client.return_value

    def test_missing_env_var_raises(self) -> None:
        with patch.dict("os.environ", {}, clear=True):
            with pytest.raises(KeyError):
                get_graph_client()


class TestCreateFolder:
    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    def test_delegates_to_async(self, mock_async) -> None:
        mock_async.return_value = ("item-id", True)
        client = MagicMock()
        result = create_folder(client, "drive-1", "base/path", "sub/folder")
        assert result == ("item-id", True)
        mock_async.assert_called_once_with(client, "drive-1", "base/path", "sub/folder")

    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    def test_none_folder_name(self, mock_async) -> None:
        mock_async.return_value = ("parent-id", False)
        client = MagicMock()
        result = create_folder(client, "drive-1", "base/path", None)
        assert result == ("parent-id", False)
        mock_async.assert_called_once_with(client, "drive-1", "base/path", None)


class TestCreateFolderGraph:
    """Tests that verify Graph API calls are constructed correctly for all
    combinations of parent_path and folder_name."""

    def _mock_client(self, item_id: str = "found-id") -> MagicMock:
        mock_item = MagicMock()
        mock_item.id = item_id

        mock_get = AsyncMock(return_value=mock_item)
        mock_post = AsyncMock(return_value=mock_item)

        client = MagicMock()
        drive = client.drives.by_drive_id.return_value
        items = drive.items.by_drive_item_id.return_value
        items.get = mock_get
        items.children.post = mock_post
        return client

    def test_no_base_path_with_folder_name(self) -> None:
        client = self._mock_client("new-id")
        create_folder(client, "drive-1", None, "reports")
        by_item = client.drives.by_drive_id.return_value.items.by_drive_item_id
        by_item.assert_any_call("root:/reports:")

    def test_neither_path_returns_root(self) -> None:
        client = self._mock_client("root-id")
        item_id, created = create_folder(client, "drive-1", None, None)
        assert item_id == "root-id"
        assert created is False
        by_item = client.drives.by_drive_id.return_value.items.by_drive_item_id
        by_item.assert_called_with("root")

    def test_base_path_no_folder_name(self) -> None:
        client = self._mock_client("base-id")
        item_id, created = create_folder(client, "drive-1", "base", None)
        assert item_id == "base-id"
        assert created is False
        by_item = client.drives.by_drive_id.return_value.items.by_drive_item_id
        by_item.assert_called_with("root:/base:")

    def test_both_paths_provided(self) -> None:
        client = self._mock_client("existing-id")
        item_id, created = create_folder(client, "drive-1", "base", "sub")
        assert item_id == "existing-id"
        assert created is False
        by_item = client.drives.by_drive_id.return_value.items.by_drive_item_id
        by_item.assert_any_call("root:/base/sub:")


class TestResolveDriveId:
    def _mock_client(
        self, site_id: str = "site-abc", drives: list[tuple[str, str]] | None = None
    ) -> MagicMock:
        if drives is None:
            drives = [("drive-1", "Documents")]

        mock_site = MagicMock()
        mock_site.id = site_id

        mock_drives_response = MagicMock()
        mock_drive_items = []
        for drive_id, name in drives:
            d = MagicMock()
            d.id = drive_id
            d.name = name
            mock_drive_items.append(d)
        mock_drives_response.value = mock_drive_items

        client = MagicMock()
        client.sites.by_site_id.return_value.get = AsyncMock(return_value=mock_site)
        client.sites.by_site_id.return_value.drives.get = AsyncMock(
            return_value=mock_drives_response
        )
        return client

    def test_resolves_default_documents_drive(self) -> None:
        client = self._mock_client(drives=[("drive-1", "Documents")])
        drive_id = resolve_drive_id(
            client, "https://example.sharepoint.com/sites/marketing"
        )
        assert drive_id == "drive-1"

    def test_resolves_named_drive(self) -> None:
        client = self._mock_client(
            drives=[("drive-1", "Documents"), ("drive-2", "Reports")]
        )
        drive_id = resolve_drive_id(
            client, "https://example.sharepoint.com/sites/marketing", "Reports"
        )
        assert drive_id == "drive-2"

    def test_raises_on_missing_drive(self) -> None:
        client = self._mock_client(drives=[("drive-1", "Documents")])
        with pytest.raises(RuntimeError, match="Drive 'Reports' not found"):
            resolve_drive_id(
                client,
                "https://example.sharepoint.com/sites/marketing",
                "Reports",
            )

    def test_raises_on_invalid_url(self) -> None:
        client = MagicMock()
        with pytest.raises(ValueError, match="must include hostname and path"):
            resolve_drive_id(client, "not-a-url")

    def test_site_lookup_uses_correct_format(self) -> None:
        client = self._mock_client()
        resolve_drive_id(
            client, "https://example.sharepoint.com/sites/marketing"
        )
        client.sites.by_site_id.assert_any_call(
            "example.sharepoint.com:/sites/marketing:"
        )


class TestResolveAndCreateFolder:
    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_combines_resolve_and_create(self, mock_resolve, mock_create) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", True)
        client = MagicMock()
        drive_id, folder_id, created, downloaded = resolve_and_create_folder(
            client, "https://example.sharepoint.com/sites/test", None, "base", "sub"
        )
        assert drive_id == "drive-1"
        assert folder_id == "folder-id"
        assert created is True
        assert downloaded == []
        mock_resolve.assert_called_once_with(
            client, "https://example.sharepoint.com/sites/test", None
        )
        mock_create.assert_called_once_with(client, "drive-1", "base", "sub")

    @patch("sharepoint._download_files", new_callable=AsyncMock)
    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_downloads_when_existing_and_dest_provided(
        self, mock_resolve, mock_create, mock_download
    ) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", False)
        mock_download.return_value = ["file.txt"]
        client = MagicMock()
        drive_id, folder_id, created, downloaded = resolve_and_create_folder(
            client,
            "https://example.sharepoint.com/sites/test",
            None,
            "base",
            "sub",
            download_dest="/tmp/dest",
        )
        assert created is False
        assert downloaded == ["file.txt"]
        mock_download.assert_called_once_with(
            client, "drive-1", "folder-id", "/tmp/dest", ""
        )

    @patch("sharepoint._download_files", new_callable=AsyncMock)
    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_skips_download_when_created(
        self, mock_resolve, mock_create, mock_download
    ) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", True)
        client = MagicMock()
        _, _, created, downloaded = resolve_and_create_folder(
            client,
            "https://example.sharepoint.com/sites/test",
            None,
            "base",
            "sub",
            download_dest="/tmp/dest",
        )
        assert created is True
        assert downloaded == []
        mock_download.assert_not_called()

    @patch("sharepoint._download_files", new_callable=AsyncMock)
    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_skips_download_when_no_dest(
        self, mock_resolve, mock_create, mock_download
    ) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", False)
        client = MagicMock()
        _, _, created, downloaded = resolve_and_create_folder(
            client,
            "https://example.sharepoint.com/sites/test",
            None,
            "base",
            "sub",
        )
        assert created is False
        assert downloaded == []
        mock_download.assert_not_called()


class TestDownloadFilesGraph:
    """Tests that verify _download_files Graph API calls via resolve_and_create_folder."""

    def _make_file_item(self, name: str, item_id: str) -> MagicMock:
        item = MagicMock()
        item.name = name
        item.id = item_id
        item.folder = None
        return item

    def _make_folder_item(self, name: str, item_id: str) -> MagicMock:
        item = MagicMock()
        item.name = name
        item.id = item_id
        item.folder = MagicMock()
        return item

    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_downloads_files(self, mock_resolve, mock_create, tmp_path) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", False)

        file_item = self._make_file_item("report.txt", "file-1")
        children_response = MagicMock()
        children_response.value = [file_item]

        client = MagicMock()
        drive = client.drives.by_drive_id.return_value
        items = drive.items.by_drive_item_id.return_value
        items.children.get = AsyncMock(return_value=children_response)
        items.content.get = AsyncMock(return_value=b"hello world")

        _, _, _, downloaded = resolve_and_create_folder(
            client, "https://example.sharepoint.com/sites/test",
            None, "base", "sub", download_dest=str(tmp_path),
        )
        assert downloaded == ["report.txt"]
        assert (tmp_path / "report.txt").read_bytes() == b"hello world"

    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_downloads_nested_folders(self, mock_resolve, mock_create, tmp_path) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", False)

        inner_file = self._make_file_item("data.csv", "file-2")
        subfolder = self._make_folder_item("sub", "folder-2")

        top_response = MagicMock()
        top_response.value = [subfolder]
        sub_response = MagicMock()
        sub_response.value = [inner_file]

        client = MagicMock()
        drive = client.drives.by_drive_id.return_value

        call_count = 0

        async def mock_children_get(*args, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 1:
                return top_response
            return sub_response

        drive.items.by_drive_item_id.return_value.children.get = AsyncMock(
            side_effect=mock_children_get
        )
        drive.items.by_drive_item_id.return_value.content.get = AsyncMock(
            return_value=b"csv data"
        )

        _, _, _, downloaded = resolve_and_create_folder(
            client, "https://example.sharepoint.com/sites/test",
            None, "base", "sub", download_dest=str(tmp_path),
        )
        assert downloaded == ["sub/data.csv"]
        assert (tmp_path / "sub" / "data.csv").read_bytes() == b"csv data"

    @patch("sharepoint._create_folder", new_callable=AsyncMock)
    @patch("sharepoint._resolve_drive_id", new_callable=AsyncMock)
    def test_empty_folder(self, mock_resolve, mock_create, tmp_path) -> None:
        mock_resolve.return_value = "drive-1"
        mock_create.return_value = ("folder-id", False)

        children_response = MagicMock()
        children_response.value = []

        client = MagicMock()
        drive = client.drives.by_drive_id.return_value
        drive.items.by_drive_item_id.return_value.children.get = AsyncMock(
            return_value=children_response
        )

        _, _, _, downloaded = resolve_and_create_folder(
            client, "https://example.sharepoint.com/sites/test",
            None, "base", "sub", download_dest=str(tmp_path),
        )
        assert downloaded == []


class TestUploadFiles:
    @patch("sharepoint._upload_files", new_callable=AsyncMock)
    def test_delegates_to_async(self, mock_async) -> None:
        client = MagicMock()
        files = [("/tmp/a.txt", "a.txt")]
        upload_files(client, "drive-1", "folder/path", files)
        mock_async.assert_called_once_with(client, "drive-1", "folder/path", files)
