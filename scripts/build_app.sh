#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR="${ROOT_DIR}/venv"
PACKAGE_DIR="${ROOT_DIR}/build/package"
PYINSTALLER_BUILD_DIR="${ROOT_DIR}/build/main"
PYINSTALLER_ALT_BUILD_DIR="${ROOT_DIR}/build/WordCompare"

cd "${ROOT_DIR}"

echo "[build_app] Cleaning previous outputs"
rm -rf "${ROOT_DIR}/dist" "${VENV_DIR}" "${PACKAGE_DIR}" "${PYINSTALLER_BUILD_DIR}" "${PYINSTALLER_ALT_BUILD_DIR}"
mkdir -p "${PACKAGE_DIR}"

echo "[build_app] Creating virtual environment with ${PYTHON_BIN}"
"${PYTHON_BIN}" -m venv "${VENV_DIR}"
"${VENV_DIR}/bin/python" -m pip install --upgrade pip
"${VENV_DIR}/bin/pip" install -r requirements.txt
"${VENV_DIR}/bin/pip" install pyinstaller

echo "[build_app] Building native Rust extractor"
bash "${ROOT_DIR}/scripts/build_native_extractor.sh"

echo "[build_app] Running test suite"
"${VENV_DIR}/bin/python" -m unittest discover -s tests -v

echo "[build_app] Building PyInstaller package"
"${VENV_DIR}/bin/pyinstaller" main.spec

VERSION="$("${VENV_DIR}/bin/python" -c "from version import __version__; print(__version__)")"
PACKAGE_NAME="WordCompare-Linux-${VERSION}"
PACKAGE_ROOT="${PACKAGE_DIR}/${PACKAGE_NAME}"

rm -rf "${PACKAGE_ROOT}" "${PACKAGE_DIR}/${PACKAGE_NAME}.tar.gz"
mkdir -p "${PACKAGE_ROOT}"
cp "${ROOT_DIR}/dist/WordCompare" "${PACKAGE_ROOT}/WordCompare"
cp "${ROOT_DIR}/README.md" "${PACKAGE_ROOT}/README.md"
cp "${ROOT_DIR}/LICENSE" "${PACKAGE_ROOT}/LICENSE"
chmod +x "${PACKAGE_ROOT}/WordCompare"

echo "[build_app] Creating tar.gz package"
tar -C "${PACKAGE_DIR}" -czf "${PACKAGE_DIR}/${PACKAGE_NAME}.tar.gz" "${PACKAGE_NAME}"

echo "Packaged artifact: ${PACKAGE_DIR}/${PACKAGE_NAME}.tar.gz"
