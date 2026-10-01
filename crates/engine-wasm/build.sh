#!/usr/bin/env bash
# Builds the engine for Node and drops the wasm-bindgen output into
# packages/gridpath/engine/. Requires: rustup target wasm32-unknown-unknown,
# wasm-bindgen-cli matching the version in Cargo.lock.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
cd "$HERE"
# Rust bakes source paths into panic-location strings; strip the machine-
# specific prefixes so the published .wasm carries no home directory or user.
REPO="$(cd ../.. && pwd)"
export RUSTFLAGS="${RUSTFLAGS:-} --remap-path-prefix=${HOME}/.cargo/registry/src=/cargo/registry --remap-path-prefix=${HOME}/.cargo/git/checkouts=/cargo/git --remap-path-prefix=${HOME}/.rustup=/rustup --remap-path-prefix=${REPO}=/gridpath --remap-path-prefix=${HOME}=/home"
cargo build --target wasm32-unknown-unknown --release
wasm-bindgen --target nodejs --out-dir ../../packages/gridpath/engine \
  target/wasm32-unknown-unknown/release/gridpath_engine.wasm
python3 "$HERE/cjs-snippets.py" ../../packages/gridpath/engine
# wasm-bindgen's nodejs target is CommonJS; the package is ESM.
printf '{ "type": "commonjs" }\n' > ../../packages/gridpath/engine/package.json
ls -la ../../packages/gridpath/engine/
