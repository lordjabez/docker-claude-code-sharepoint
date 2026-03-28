# Docker Claude Code SharePoint

[![Docker Hub](https://github.com/lordjabez/docker-claude-code-sharepoint/actions/workflows/publish.yml/badge.svg)](https://github.com/lordjabez/docker-claude-code-sharepoint/actions/workflows/publish.yml)

A Docker image that layers SharePoint integration on top of
[docker-claude-code](https://github.com/lordjabez/docker-claude-code). It
provides pre and post hooks that create a SharePoint folder before Claude runs
and upload all output files afterward, using the Microsoft Graph API with
on-behalf-of (OBO) authentication.

Downstream images extend this one to add domain-specific Claude instructions,
reference examples, and additional dependencies.

## Prerequisites

- Docker
- An Anthropic API key exported as `ANTHROPIC_API_KEY`
- An Entra ID app registration configured for OBO (see below)

## Usage

Build the image:

```bash
bin/build.bash
```

Run:

```bash
bin/run.bash '{"prompt": "Your prompt here", "path": "2025/my-output"}'
```

The container accepts a JSON object with two keys:

- `prompt` — the prompt passed to Claude (extracted by the `docker-claude-code` entrypoint)
- `path` — optional SharePoint folder path created under `SHAREPOINT_BASE_PATH`; when omitted, output goes to the base path (or drive root if `SHAREPOINT_BASE_PATH` is also unset)

The pre-hook creates the folder and the post-hook uploads all workspace output to it.

## Extending

Downstream images add their own Claude instructions, reference files, and
dependencies. A minimal downstream Dockerfile:

```dockerfile
FROM lordjabez/claude-code-sharepoint:latest

COPY examples ./examples
COPY .claude-settings/ /home/claude/.claude/
```

The downstream image can also override the hooks if needed by copying its own
scripts to `/home/claude/hooks/`.

## SharePoint Integration

Authentication uses the on-behalf-of (OBO) flow: an upstream REST endpoint
authenticates the user via Entra ID and passes their access token to the
container, which exchanges it for a Graph-scoped token. Uploaded files appear
as the calling user in SharePoint, and their permissions are respected.

### Entra ID Setup

1. In the Azure portal, go to **Entra ID > App registrations > New registration**
   and create a new app.

2. From the app's **Overview** page, note the **Application (client) ID** and
   **Directory (tenant) ID**.

3. Go to **Certificates & secrets > New client secret**. Copy the secret
   **value** (not the secret ID).

4. Go to **API permissions > Add a permission > Microsoft Graph > Delegated
   permissions** and add `Sites.ReadWrite.All`. Click **Grant admin consent**.

5. Go to **Expose an API**:

   - Click **Set** next to "Application ID URI" and accept the default
     (`api://<client-id>`).
   - Click **Add a scope** (e.g., name it `access_as_user`).
   - Under **Authorized client applications**, add the Azure CLI client ID
     (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`) and select your scope. This
     enables local testing. When you have an upstream REST endpoint app, add
     its client ID here too.

### Getting a User Access Token (Local Testing)

For local testing without the upstream REST endpoint, use the Azure CLI:

```bash
az login --scope api://<your-client-id>/.default
export USER_ACCESS_TOKEN=$(az account get-access-token --resource api://<your-client-id> --query accessToken -o tsv)
```

In production, the upstream REST endpoint obtains the user's token via Entra ID
OAuth2 and passes it to the container as `USER_ACCESS_TOKEN`.

### Environment Variables

| Variable | Required | Description |
| --- | --- | --- |
| `ANTHROPIC_API_KEY` | Yes | Anthropic API key |
| `AZURE_TENANT_ID` | Yes | Entra ID tenant ID |
| `AZURE_CLIENT_ID` | Yes | App registration client ID |
| `AZURE_CLIENT_SECRET` | Yes | App registration client secret |
| `USER_ACCESS_TOKEN` | Yes | User's JWT (from upstream endpoint or Azure CLI) |
| `SHAREPOINT_SITE_URL` | Yes | SharePoint site URL (e.g. `https://example.sharepoint.com/sites/marketing`) |
| `SHAREPOINT_DRIVE_NAME` | No | Document library name (default: `Documents`) |
| `SHAREPOINT_BASE_PATH` | No | Folder prefix in SharePoint (default: empty) |
| `SHAREPOINT_SKIP_PATTERNS` | No | Comma-separated glob patterns for files/dirs to exclude from upload (default: empty) |

## Project Structure

```text
bin/build.bash          Builds the Docker image
bin/run.bash            Runs the container with JSON input
bin/test.bash           Integration test: runs three containers against live SharePoint
hooks/sharepoint.py     Graph API helpers (OBO auth, site/drive resolution, folder creation, file upload)
hooks/pre.py            Pre-hook: creates SharePoint folder from JSON input path
hooks/post.py           Post-hook: uploads all workspace output to SharePoint
tests/                  Unit tests
Dockerfile              Extends lordjabez/claude-code with SharePoint hooks
pyproject.toml          Python dependencies
```

## CI/CD

A GitHub Actions workflow builds and pushes the image on every push to `main`.
Each build is tagged `latest` plus the installed Claude Code CLI version (e.g.
`2.1.83`). Images are built for both linux/amd64 and linux/arm64.

## Limitations

- **4MB upload limit**: The post-hook uses the Graph API's simple upload endpoint,
  which is limited to 4MB per file. If larger files are needed, the upload logic
  in `hooks/sharepoint.py` would need to switch to the
  [resumable upload session API](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession).

## License

MIT-0
