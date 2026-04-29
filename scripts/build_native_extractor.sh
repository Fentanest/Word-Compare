#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CRATE_DIR="${ROOT_DIR}/native/docx-structure-extractor"
TARGET_DIR="${CRATE_DIR}/target/release"
OUTPUT_DIR="${ROOT_DIR}/build/native"
OUTPUT_NAME="word_compare_native_extractor"

if ! command -v cargo >/dev/null 2>&1 && [[ -f "${HOME}/.cargo/env" ]]; then
  # Support shells that have Rust installed but have not reloaded PATH yet.
  # shellcheck disable=SC1090
  source "${HOME}/.cargo/env"
fi

mkdir -p "${OUTPUT_DIR}"

cargo build --manifest-path "${CRATE_DIR}/Cargo.toml" --release

cp "${TARGET_DIR}/${OUTPUT_NAME}" "${OUTPUT_DIR}/${OUTPUT_NAME}"
if [[ -f "${TARGET_DIR}/${OUTPUT_NAME}.exe" ]]; then
  cp "${TARGET_DIR}/${OUTPUT_NAME}.exe" "${OUTPUT_DIR}/${OUTPUT_NAME}.exe"
fi

echo "Native extractor built into ${OUTPUT_DIR}"
