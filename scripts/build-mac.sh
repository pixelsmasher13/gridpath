#!/usr/bin/env bash
# Build, sign, notarize, and upload the GridPath macOS .dmg to S3.
#
# Required env vars (for signing + notarization):
#   APPLE_API_ISSUER         (App Store Connect API issuer ID)
#   APPLE_API_KEY            (App Store Connect API key ID)
#   APPLE_API_KEY_PATH       (absolute path to the AuthKey_*.p8 file)
#   APPLE_SIGNING_IDENTITY   (e.g. "Developer ID Application: Your Co (TEAMID)")
#   TAURI_SIGNING_PRIVATE_KEY (for updater artifacts)
#
# Optional:
#   TAURI_SIGNING_PRIVATE_KEY_PASSWORD (if the updater key has a passphrase)
#   S3_BUCKET                (default: gridpath)
#   S3_REGION                (default: us-east-1)
#   AWS_PROFILE              (default: none; uses your default AWS credentials)
#
# Usage:
#   ./scripts/build-mac.sh                # build + upload
#   ./scripts/build-mac.sh --no-upload    # build only (skip S3)

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

# ── Auto-load credentials from .env.build if present (gitignored) ──
# `set -a` auto-exports every variable defined while sourcing, so subprocesses
# (npm, tauri-bundler, codesign, notarytool) inherit them even if .env.build
# uses bare `KEY=value` lines without `export`.
if [[ -f "$ROOT_DIR/.env.build" ]]; then
  echo "==> Loading credentials from .env.build"
  set -a
  # shellcheck disable=SC1090
  source "$ROOT_DIR/.env.build"
  set +a
fi

# ── Require signing env vars ──
: "${APPLE_API_ISSUER:?APPLE_API_ISSUER must be set}"
: "${APPLE_API_KEY:?APPLE_API_KEY must be set}"
: "${APPLE_API_KEY_PATH:?APPLE_API_KEY_PATH must be set}"
: "${APPLE_SIGNING_IDENTITY:?APPLE_SIGNING_IDENTITY must be set}"

if [[ ! -f "$APPLE_API_KEY_PATH" ]]; then
  echo "❌ APPLE_API_KEY_PATH does not exist: $APPLE_API_KEY_PATH"
  exit 1
fi

S3_BUCKET="${S3_BUCKET:-gridpath}"
S3_REGION="${S3_REGION:-us-east-1}"
SKIP_UPLOAD=0
for arg in "$@"; do
  [[ "$arg" == "--no-upload" ]] && SKIP_UPLOAD=1
done

# ── Verify AWS CLI is available (unless skipping upload) ──
if [[ "$SKIP_UPLOAD" -eq 0 ]]; then
  if ! command -v aws >/dev/null 2>&1; then
    echo "❌ aws CLI not found. Install with: brew install awscli"
    exit 1
  fi
fi

# ── Build ──
# Use `npm run tauri build` directly — do NOT call scripts/build-platform.sh.
# That script does an unsafe `mv src-tauri/resources/windows ...-temp` during
# the macOS build, and if the build crashes before the restore step runs it
# leaves the working tree with thousands of "deleted" files in git status.
# Windows resources are only bundled into Windows targets, so the swap is
# unnecessary for a macOS-only build.
# Clean stale dmg artifacts from prior failed runs. Tauri's bundle_dmg.sh
# calls `hdiutil convert ... -o <target>` without `-ov`, so a leftover
# GridPath_*.dmg or rw.GridPath_*.dmg in bundle/macos/ aborts the bundling
# step before it ever produces a new image.
rm -f src-tauri/target/release/bundle/macos/*.dmg \
      src-tauri/target/release/bundle/macos/.DS_Store \
      src-tauri/target/release/bundle/dmg/*.dmg \
      src-tauri/target/release/bundle/dmg/rw.*.dmg

echo "==> Building GridPath for macOS (signed + notarized)..."
npm run tauri build

# ── Locate the DMG ──
DMG_PATH=$(find src-tauri/target/release/bundle/dmg -maxdepth 1 -name "*.dmg" | head -1)
if [[ -z "$DMG_PATH" ]]; then
  echo "❌ No .dmg produced under src-tauri/target/release/bundle/dmg"
  exit 1
fi

DMG_FILE=$(basename "$DMG_PATH")
DMG_SIZE=$(du -h "$DMG_PATH" | cut -f1)
echo "==> Built: $DMG_PATH ($DMG_SIZE)"

# Verify signature + notarization ticket
echo "==> Verifying signature..."
codesign --verify --deep --strict --verbose=2 "$DMG_PATH" 2>&1 | tail -3 || true

if [[ "$SKIP_UPLOAD" -eq 1 ]]; then
  echo "==> Skipping S3 upload (--no-upload)"
  exit 0
fi

# ── Upload to S3 ──
echo "==> Uploading to s3://$S3_BUCKET/$DMG_FILE"
aws s3 cp "$DMG_PATH" "s3://$S3_BUCKET/$DMG_FILE" \
  --region "$S3_REGION" \
  --content-type application/x-apple-diskimage

# Mirror to a stable "latest" URL for the website's download button
LATEST_KEY="GridPath-latest.dmg"
echo "==> Mirroring to s3://$S3_BUCKET/$LATEST_KEY"
aws s3 cp "$DMG_PATH" "s3://$S3_BUCKET/$LATEST_KEY" \
  --region "$S3_REGION" \
  --content-type application/x-apple-diskimage

# ── Updater artifacts (.app.tar.gz + latest.json) ──
# Tauri's updater works in-place: existing installs poll latest.json and
# swap themselves to the new .app.tar.gz without re-downloading the DMG.
# The DMG above is for FIRST installs (marketing site download). These
# two artifacts are for AUTO-UPGRADES of already-installed users.
TARBALL_PATH="src-tauri/target/release/bundle/macos/GridPath.app.tar.gz"
SIG_PATH="${TARBALL_PATH}.sig"
if [[ ! -f "$TARBALL_PATH" || ! -f "$SIG_PATH" ]]; then
  echo "⚠️  Updater artifacts missing ($TARBALL_PATH or .sig). Skipping latest.json."
  echo "   (DMG is published — first-install download still works.)"
else
  # Extract version from tauri.conf.json. node is available in any
  # tauri-build environment so this is safe.
  VERSION=$(node -p "require('./src-tauri/tauri.conf.json').version")
  TARBALL_KEY="GridPath_${VERSION}_aarch64.app.tar.gz"

  echo "==> Uploading updater bundle to s3://$S3_BUCKET/$TARBALL_KEY"
  aws s3 cp "$TARBALL_PATH" "s3://$S3_BUCKET/$TARBALL_KEY" \
    --region "$S3_REGION" \
    --content-type application/gzip

  # Read signature inline (Tauri's updater verifies it against the
  # pubkey in tauri.conf.json — no separate .sig file lookup needed).
  SIGNATURE=$(cat "$SIG_PATH")
  PUB_DATE=$(date -u +%Y-%m-%dT%H:%M:%SZ)
  TARBALL_URL="https://$S3_BUCKET.s3.$S3_REGION.amazonaws.com/$TARBALL_KEY"

  # Build latest.json. Both aarch64 and x86_64 keys point at the same
  # tarball — if you ever ship a universal binary you only need one
  # entry; if you split to per-arch builds, swap the urls/signatures.
  MANIFEST_TMP=$(mktemp -t gridpath-latest.json)
  cat > "$MANIFEST_TMP" <<EOF
{
  "version": "${VERSION}",
  "notes": "Update to version ${VERSION}",
  "pub_date": "${PUB_DATE}",
  "platforms": {
    "darwin-aarch64": {
      "signature": "${SIGNATURE}",
      "url": "${TARBALL_URL}"
    },
    "darwin-x86_64": {
      "signature": "${SIGNATURE}",
      "url": "${TARBALL_URL}"
    }
  }
}
EOF

  # latest.json lives at the URL configured in tauri.conf.json:
  #   https://$S3_BUCKET.s3.$S3_REGION.amazonaws.com/latest_app/latest.json
  echo "==> Uploading latest.json to s3://$S3_BUCKET/latest_app/latest.json"
  aws s3 cp "$MANIFEST_TMP" "s3://$S3_BUCKET/latest_app/latest.json" \
    --region "$S3_REGION" \
    --content-type application/json \
    --cache-control "no-cache, max-age=60"
  rm -f "$MANIFEST_TMP"
fi

echo
echo "✅ Done"
echo "   Versioned DMG:   https://$S3_BUCKET.s3.$S3_REGION.amazonaws.com/$DMG_FILE"
echo "   Latest DMG:      https://$S3_BUCKET.s3.$S3_REGION.amazonaws.com/$LATEST_KEY"
if [[ -n "${VERSION:-}" ]]; then
  echo "   Updater bundle:  https://$S3_BUCKET.s3.$S3_REGION.amazonaws.com/$TARBALL_KEY"
  echo "   Updater manifest: https://$S3_BUCKET.s3.$S3_REGION.amazonaws.com/latest_app/latest.json"
fi
