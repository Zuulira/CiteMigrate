#!/bin/bash
# ──────────────────────────────────────────────
# Build script for CiteMigrate macOS app
# Creates a standalone .app bundle using PyInstaller
#
# Uses PyQt6 (no Tcl/Tk dependency - avoids the
# macOS system Tk 8.5 crash issue)
#
# Usage:  cd /path/to/citemigrate-github && bash build_app.sh
# Output: dist/CiteMigrate.app
# ──────────────────────────────────────────────

set -e

APP_NAME="CiteMigrate"
VERSION="1.0"
BUNDLE_ID="com.citemigrate.app"
PYTHON=${PYTHON:-python3}

echo "========================================"
echo "  Building ${APP_NAME}.app"
echo "========================================"

# Check Python version
PY_VERSION=$($PYTHON --version 2>&1)
echo "Using Python: $PY_VERSION"

# Ensure Python >= 3.9
$PYTHON -c "import sys; assert sys.version_info >= (3, 9), 'Python 3.9+ required'" || {
    echo "ERROR: Python 3.9 or higher is required."
    echo "Install with: brew install python@3.12"
    exit 1
}

# Install dependencies
echo ""
echo "Step 1: Installing dependencies..."
$PYTHON -m pip install -r requirements.txt --quiet

# Clean previous builds
echo ""
echo "Step 2: Cleaning previous builds..."
rm -rf build/ dist/ *.spec

# Prepare icon
echo ""
echo "Step 3: Preparing icon..."

ICON_FLAG=""
ADD_DATA_FLAG=""

# Generate .icns from icon.png if sips + iconutil are available (macOS)
if [ ! -f "${APP_NAME}.icns" ] && [ -f "icon.png" ]; then
    if command -v sips &>/dev/null && command -v iconutil &>/dev/null; then
        echo "  Generating .icns from icon.png..."
        ICONSET="${APP_NAME}.iconset"
        mkdir -p "${ICONSET}"
        sips -z 16 16     icon.png --out "${ICONSET}/icon_16x16.png"      >/dev/null 2>&1
        sips -z 32 32     icon.png --out "${ICONSET}/icon_16x16@2x.png"   >/dev/null 2>&1
        sips -z 32 32     icon.png --out "${ICONSET}/icon_32x32.png"      >/dev/null 2>&1
        sips -z 64 64     icon.png --out "${ICONSET}/icon_32x32@2x.png"   >/dev/null 2>&1
        sips -z 128 128   icon.png --out "${ICONSET}/icon_128x128.png"    >/dev/null 2>&1
        sips -z 256 256   icon.png --out "${ICONSET}/icon_128x128@2x.png" >/dev/null 2>&1
        sips -z 256 256   icon.png --out "${ICONSET}/icon_256x256.png"    >/dev/null 2>&1
        sips -z 512 512   icon.png --out "${ICONSET}/icon_256x256@2x.png" >/dev/null 2>&1
        sips -z 512 512   icon.png --out "${ICONSET}/icon_512x512.png"    >/dev/null 2>&1
        cp icon.png "${ICONSET}/icon_512x512@2x.png"
        iconutil -c icns "${ICONSET}" -o "${APP_NAME}.icns" 2>/dev/null && \
            echo "  Generated ${APP_NAME}.icns" || \
            echo "  WARNING: iconutil failed"
        rm -rf "${ICONSET}"
    else
        echo "  WARNING: sips/iconutil not found (not on macOS?) — building without .icns"
    fi
fi

if [ -f "${APP_NAME}.icns" ]; then
    ICON_FLAG="--icon=${APP_NAME}.icns"
    echo "  Using ${APP_NAME}.icns for Dock/Finder icon"
fi

if [ -f "icon.png" ]; then
    ADD_DATA_FLAG="--add-data=icon.png:."
    echo "  Bundling icon.png for window title bar icon"
fi

# Build with PyInstaller
echo ""
echo "Step 4: Building .app bundle..."

$PYTHON -m PyInstaller \
    --windowed \
    --onedir \
    --name "${APP_NAME}" \
    --osx-bundle-identifier "${BUNDLE_ID}" \
    $ICON_FLAG \
    $ADD_DATA_FLAG \
    --hidden-import=pyzotero \
    --hidden-import=lxml \
    --hidden-import=lxml.etree \
    --hidden-import=PyQt6 \
    --hidden-import=PyQt6.QtWidgets \
    --hidden-import=PyQt6.QtCore \
    --hidden-import=PyQt6.QtGui \
    --hidden-import=PyQt6.sip \
    --collect-all PyQt6 \
    --noconfirm \
    citemigrate.py

APP_PATH="dist/${APP_NAME}.app"

# Check result
if [ ! -d "${APP_PATH}" ]; then
    echo ""
    echo "BUILD FAILED - .app not found. Checking dist/:"
    ls -la dist/ 2>/dev/null || echo "  dist/ does not exist"
    exit 1
fi

# Patch Info.plist with version info
echo ""
echo "Step 5: Patching Info.plist..."
PLIST="${APP_PATH}/Contents/Info.plist"
if [ -f "${PLIST}" ]; then
    /usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString ${VERSION}" "${PLIST}" 2>/dev/null || \
        /usr/libexec/PlistBuddy -c "Add :CFBundleShortVersionString string ${VERSION}" "${PLIST}"
    /usr/libexec/PlistBuddy -c "Set :CFBundleVersion ${VERSION}" "${PLIST}" 2>/dev/null || \
        /usr/libexec/PlistBuddy -c "Add :CFBundleVersion string ${VERSION}" "${PLIST}"
    /usr/libexec/PlistBuddy -c "Set :CFBundleDisplayName ${APP_NAME}" "${PLIST}" 2>/dev/null || \
        /usr/libexec/PlistBuddy -c "Add :CFBundleDisplayName string ${APP_NAME}" "${PLIST}"
    echo "  Info.plist updated"
fi

# Ad-hoc code sign
echo ""
echo "Step 6: Code signing (ad-hoc)..."
if command -v codesign &>/dev/null; then
    codesign --force --deep --sign - "${APP_PATH}" 2>&1 && \
        echo "  Signed successfully" || \
        echo "  WARNING: codesign failed (try right-click > Open)"
fi

echo ""
echo "========================================"
echo "  BUILD SUCCESSFUL!"
echo "========================================"
echo ""
echo "  App location: ${APP_PATH}"
echo "  App size: $(du -sh "${APP_PATH}" | cut -f1)"
echo ""
echo "  To run directly:"
echo "    open ${APP_PATH}"
echo ""
echo "  To install:"
echo "    1. Drag '${APP_PATH}' to your Applications folder"
echo "    2. On first launch, right-click > Open (to bypass Gatekeeper)"
echo ""
echo ""
