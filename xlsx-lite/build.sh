#!/bin/bash
# Builds a universal (Apple Silicon + Intel) app. Usage: bash build.sh ["/path/to/XLSX Lite.app"]
set -euo pipefail
SRC="$(cd "$(dirname "$0")" && pwd)"
APP="${1:-/private/tmp/XLSX Lite.app}"
case "$APP" in /private/tmp/*) ;; *) echo "Build destination must be inside /private/tmp" >&2; exit 1 ;; esac
TMP="$(mktemp -d /private/tmp/xlsxlite-build.XXXXXX)"; trap 'rm -rf "$TMP"' EXIT
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
for ARCH in arm64 x86_64; do
    xcrun swiftc -O -whole-module-optimization -swift-version 5 \
        -target "$ARCH-apple-macosx13.0" -module-cache-path "$TMP/cache" \
        "$SRC/main.swift" -o "$TMP/XLSXLite-$ARCH"
done
lipo -create "$TMP/XLSXLite-arm64" "$TMP/XLSXLite-x86_64" -output "$APP/Contents/MacOS/XLSXLite"
strip -x "$APP/Contents/MacOS/XLSXLite"
cp "$SRC/Info.plist" "$APP/Contents/Info.plist"
cp "$SRC/AppIcon.icns" "$APP/Contents/Resources/AppIcon.icns"
codesign --force --sign - "$APP"
du -sh "$APP"
