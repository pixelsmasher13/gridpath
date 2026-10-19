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
# The browser build (the package's main export) is not part of the server.
rm -rf "$STAGE/server/dist/web"
# Runtime deps only, installed inside the bundle (Claude Desktop ships Node, not npm).
node -e '
const p = require(process.argv[1]);
// Dependencies of the browser build stay out of the server bundle.
const { "styled-components": _web, ...dependencies } = p.dependencies;
const out = { name: p.name, version: p.version, type: p.type, private: true, dependencies };
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
