#!/usr/bin/env bash
# Собирает движок ENS-Table (форк IronCalc) в WebAssembly и кладёт в src/engine-pkg.
# Нужны: Rust с целью wasm32-unknown-unknown и wasm-bindgen 0.2.126 (см. tools/env.sh).
set -euo pipefail
here="$(cd "$(dirname "$0")/.." && pwd)"
root="$here/../.."
cd "$root/core/ironcalc"
cargo build --release --target wasm32-unknown-unknown -p wasm --features xlsx
mkdir -p "$here/src/engine-pkg"
wasm-bindgen --target web --out-dir "$here/src/engine-pkg" --out-name engine \
  target/wasm32-unknown-unknown/release/wasm.wasm
echo "Движок собран: $here/src/engine-pkg"
