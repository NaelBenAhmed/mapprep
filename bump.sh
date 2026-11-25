#!/bin/sh
# Met une nouvelle version : les téléphones se mettront à jour tout seuls.
set -e
cd "$(dirname "$0")"
OLD=$(sed -n 's/.*"version":"\([0-9]*\)".*/\1/p' version.json)
NEW=$(date -u +%Y%m%d%H%M)
sed -i "s/$OLD/$NEW/g" version.json sw.js index.html
echo "Version $OLD -> $NEW"
