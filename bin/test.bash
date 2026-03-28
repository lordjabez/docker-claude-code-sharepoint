#!/usr/bin/env bash
set -e

cd "$(dirname "$0")/.."

source .env

timestamp=$(date -u +"%Y%m%dT%H%M%SZ")
export SHAREPOINT_BASE_PATH="${SHAREPOINT_BASE_PATH}/${timestamp}"

echo "Tests will output to ${SHAREPOINT_BASE_PATH}"

args_file=$(mktemp)
cat > "${args_file}" <<'EOF'
'{"prompt": "write a short poem about programming into a markdown file"}'
'{"prompt": "write a short poem about programming into a markdown file", "path": "foo"}'
'{"prompt": "write a short poem about programming into a markdown file", "path": "foo/bar"}'
EOF

fanout bin/run.bash "${args_file}"
rm -f "${args_file}"
