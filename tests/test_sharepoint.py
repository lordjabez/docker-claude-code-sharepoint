import sys
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "hooks"))

from sharepoint import create_folder, get_graph_client, upload_files


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


class TestUploadFiles:
    @patch("sharepoint._upload_files", new_callable=AsyncMock)
    def test_delegates_to_async(self, mock_async) -> None:
        client = MagicMock()
        files = [("/tmp/a.txt", "a.txt")]
        upload_files(client, "drive-1", "folder/path", files)
        mock_async.assert_called_once_with(client, "drive-1", "folder/path", files)
