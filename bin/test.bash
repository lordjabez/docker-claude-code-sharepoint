#!/usr/bin/env bash
set -e

cd "$(dirname "$0")/.."

timestamp=$(date -u +"%Y%m%dT%H%M%SZ")
export SHAREPOINT_BASE_PATH="${SHAREPOINT_BASE_PATH}/${timestamp}"

echo "Tests will output to ${SHAREPOINT_BASE_PATH}"

inputs=(
  '{"prompt": "write a short poem about programming into a markdown file"}'
  '{"prompt": "write a short poem about programming into a markdown file", "path": "foo"}'
  '{"prompt": "write a short poem about programming into a markdown file", "path": "foo/bar"}'
)

pids=()
for input in "${inputs[@]}"; do
  bin/run.bash "${input}" &
  pids+=($!)
done

failed=0
for pid in "${pids[@]}"; do
  if ! wait "${pid}"; then
    failed=$((failed + 1))
  fi
done

if [ "${failed}" -gt 0 ]; then
  echo "${failed} of ${#inputs[@]} runs failed"
  exit 1
fi

echo "All ${#inputs[@]} runs completed successfully"
