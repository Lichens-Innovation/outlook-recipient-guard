#!/usr/bin/env bash
# Usage: ./set-base-url.sh https://my-host.example.com/path
set -euo pipefail
NEW="${1%/}"
OLD="https://lichens-innovation.github.io/outlook-recipient-guard"
ORIGIN=$(echo "$NEW" | sed -E 's|^(https://[^/]+).*|\1|')
sed -i.bak -e "s|$OLD|$NEW|g" -e "s|<AppDomain>https://lichens-innovation.github.io</AppDomain>|<AppDomain>$ORIGIN</AppDomain>|" manifest.xml
rm -f manifest.xml.bak
echo "manifest.xml -> $NEW"
