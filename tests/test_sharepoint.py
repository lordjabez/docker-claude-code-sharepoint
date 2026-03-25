import sys
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "hooks"))

from sharepoint import _item_ref, create_folder, get_graph_client, upload_files


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


class TestUploadFiles:
    @patch("sharepoint._upload_files", new_callable=AsyncMock)
    def test_delegates_to_async(self, mock_async) -> None:
        client = MagicMock()
        files = [("/tmp/a.txt", "a.txt")]
        upload_files(client, "drive-1", "folder/path", files)
        mock_async.assert_called_once_with(client, "drive-1", "folder/path", files)
