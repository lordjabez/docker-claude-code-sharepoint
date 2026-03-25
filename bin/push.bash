#!/usr/bin/env bash
set -e

cd "$(dirname "$0")/.."

docker push lordjabez/claude-code-sharepoint
