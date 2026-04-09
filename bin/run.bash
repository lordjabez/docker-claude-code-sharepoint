#!/usr/bin/env bash
set -e

cd "$(dirname "$0")/.."

input="${1}"

docker run --rm \
  -e ANTHROPIC_API_KEY \
  -e AZURE_TENANT_ID \
  -e AZURE_CLIENT_ID \
  -e AZURE_CLIENT_SECRET \
  -e USER_ACCESS_TOKEN \
  -e SHAREPOINT_SITE_URL \
  -e SHAREPOINT_DRIVE_NAME \
  -e SHAREPOINT_BASE_PATH \
  -e SHAREPOINT_SKIP_PATTERNS \
  --tty \
  lordjabez/claude-code-sharepoint:latest "${input}"
