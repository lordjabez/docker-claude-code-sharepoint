# Development Guide

This project builds a Docker image that layers SharePoint integration on top of
[docker-claude-code](https://github.com/lordjabez/docker-claude-code). Downstream
images extend this one to add domain-specific Claude instructions and dependencies.

## Architecture

The base image (`lordjabez/claude-code`) provides a headless Claude Code runtime
with Python and uv pre-installed. This project adds:

- **`hooks/sharepoint.py`** - Graph API helpers: OBO authentication, folder
  creation, and file upload.
- **`hooks/pre.py`** - Pre-hook that extracts the optional `path` from the
  JSON input and creates the corresponding SharePoint folder (including any
  intermediate segments).
- **`hooks/post.py`** - Post-hook that uploads all workspace output files
  to the SharePoint folder created by the pre-hook.
- **`pyproject.toml` / `uv.lock`** - Python dependencies for Azure auth,
  Graph SDK, and document generation (openpyxl, pymupdf, pypandoc, python-docx).

## Build and Run

```bash
bin/build.bash
bin/run.bash '{"prompt": "Your prompt here", "path": "2025/my-output"}'
```

The container accepts a JSON object with `prompt` (passed to Claude by the
`docker-claude-code` entrypoint) and `path` (used by the pre-hook to create
the SharePoint folder). The `ANTHROPIC_API_KEY` and SharePoint environment
variables must be set.

## SharePoint Integration

The hooks use Microsoft Graph API to persist output to SharePoint:

- **Pre-hook** (`hooks/pre.py`): Extracts `path` from the JSON input
  (which may be `None`), creates the folder under `SHAREPOINT_BASE_PATH`
  (also optional), and writes folder metadata to `/tmp/sharepoint_context.json`.
  When both are unset the drive root is used.
- **Post-hook** (`hooks/post.py`): Reads the folder context, walks the
  workspace (skipping `.venv/`, `pyproject.toml`, and `uv.lock`), applies
  any additional `SHAREPOINT_SKIP_PATTERNS` glob patterns, and uploads all
  remaining output files.
- **Shared module** (`hooks/sharepoint.py`): Graph client setup, OBO auth,
  folder creation, and file upload helpers.

Authentication uses the on-behalf-of (OBO) flow. An upstream REST endpoint
authenticates the user via Entra ID and passes the user's access token to the
container as `USER_ACCESS_TOKEN`. The hooks exchange it for a Graph-scoped token.

### Required Environment Variables

| Variable | Description |
| --- | --- |
| `ANTHROPIC_API_KEY` | Anthropic API key for Claude |
| `AZURE_TENANT_ID` | Entra ID tenant |
| `AZURE_CLIENT_ID` | App registration client ID |
| `AZURE_CLIENT_SECRET` | App registration client secret |
| `USER_ACCESS_TOKEN` | User's JWT from the upstream endpoint |
| `SHAREPOINT_SITE_ID` | Target SharePoint site ID |
| `SHAREPOINT_DRIVE_ID` | Target document library drive ID |
| `SHAREPOINT_BASE_PATH` | Folder prefix (default: empty) |
| `SHAREPOINT_SKIP_PATTERNS` | Comma-separated glob patterns for files/dirs to exclude from upload (default: empty) |

### Entra App Registration

The app registration needs:

- `Sites.ReadWrite.All` delegated permission
- An exposed API with a scope (for OBO)
- The upstream app registered as an authorized client application

### Error Handling

Hooks fail closed. If SharePoint is unreachable or auth fails, the hook exits
non-zero and aborts the run.

## Key Files

- `hooks/sharepoint.py` contains shared Graph API helpers (OBO auth, folder
  creation, file upload).
- `hooks/pre.py` creates the SharePoint folder before Claude runs.
- `hooks/post.py` uploads all workspace output after Claude finishes.

## Extending

Downstream images add domain-specific content by extending this image:

```dockerfile
FROM lordjabez/claude-code-sharepoint:latest

COPY examples ./examples
COPY .claude-settings/ /home/claude/.claude/
```

Downstream images can override hooks by copying their own scripts to
`/home/claude/hooks/`.

## Adding Dependencies

Add Python packages to `pyproject.toml` and run `uv lock` to update `uv.lock`.
Both files are copied into the image and installed via `uv sync` during the build.

## Tests

Unit tests live in `tests/` (excluded from the Docker image via `.dockerignore`).
Run them with:

```bash
uv run pytest
```

An integration test script runs three containers in parallel against a live
SharePoint site, exercising no-path, single-segment, and nested-path scenarios.
It appends a UTC timestamp to `SHAREPOINT_BASE_PATH` so each run gets its own
folder. All SharePoint environment variables must be set:

```bash
bin/test.bash
```
