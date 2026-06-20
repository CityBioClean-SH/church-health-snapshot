#!/bin/bash
# Church Health Snapshot — Deploy Script
# Push, version, deploy — access setting preserved via appsscript.json

DEPLOY_ID="AKfycbw8UlQ_6ZOYzNne2unExCSf-SPPJb02phQiYxcnhbvIPKDlwUmjUBkWEVXdwTnx5eds"
DESC="${1:-Auto deploy}"

clasp push --force || { echo "Push failed"; exit 1; }

VERSION=$(clasp version "$DESC" 2>&1 | grep -oE '[0-9]+' | head -1)
[ -z "$VERSION" ] && { echo "Version failed"; exit 1; }

clasp deploy -i $DEPLOY_ID -V $VERSION -d "$DESC"
echo "=== Deployed v${VERSION} ==="
