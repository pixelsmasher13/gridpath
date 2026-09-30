#!/usr/bin/env bash
# Builds gridpath-<version>.mcpb for Claude Desktop (double-click install).
# Requires the package to be built first (engine, ui, dist). Uses the
# @anthropic-ai/mcpb CLI via npx.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
PKG="$(cd "$HERE/.." && pwd)"
VERSION="$(node -p "require('$PKG/package.json').version")"
STAGE="$HERE/stage"
rm -rf "$STAGE" && mkdir -p "$STAGE/server"

for d in dist engine ui; do
  [ -d "$PKG/$d" ] || { echo "missing $PKG/$d — build the package first"; exit 1; }
  cp -R "$PKG/$d" "$STAGE/server/$d"
done
# Runtime deps only, installed inside the bundle (Claude Desktop ships Node, not npm).
node -e '
const p = require(process.argv[1]);
const out = { name: p.name, version: p.version, type: p.type, private: true, dependencies: p.dependencies };
require("fs").writeFileSync(process.argv[2], JSON.stringify(out, null, 2));
' "$PKG/package.json" "$STAGE/server/package.json"
(cd "$STAGE/server" && npm install --omit=dev --no-audit --no-fund --silent --legacy-peer-deps)
# The engine's CommonJS glue must stay CommonJS inside an ESM package.
[ -f "$STAGE/server/engine/package.json" ] || printf '{ "type": "commonjs" }\n' > "$STAGE/server/engine/package.json"

sed "s/\"version\": \"[^\"]*\"/\"version\": \"$VERSION\"/" "$HERE/manifest.json" > "$STAGE/manifest.json"
cp "$HERE/icon.png" "$STAGE/icon.png"

npx -y @anthropic-ai/mcpb validate "$STAGE/manifest.json"
npx -y @anthropic-ai/mcpb pack "$STAGE" "$HERE/gridpath-$VERSION.mcpb"
ls -la "$HERE/gridpath-$VERSION.mcpb"
