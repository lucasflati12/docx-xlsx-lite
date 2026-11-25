#!/bin/bash
# Builds a universal (Apple Silicon + Intel) app. Usage: bash build.sh ["/path/to/DOCX Lite.app"]
set -euo pipefail
SRC="$(cd "$(dirname "$0")" && pwd)"
APP="${1:-$HOME/Desktop/DOCX Lite.app}"
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
for ARCH in arm64 x86_64; do
    xcrun swiftc -O -whole-module-optimization -swift-version 5 \
        -target "$ARCH-apple-macosx13.0" -module-cache-path "$TMP/cache" \
        "$SRC/main.swift" -o "$TMP/DOCXLite-$ARCH"
done
lipo -create "$TMP/DOCXLite-arm64" "$TMP/DOCXLite-x86_64" -output "$APP/Contents/MacOS/DOCXLite"
strip -x "$APP/Contents/MacOS/DOCXLite"
cp "$SRC/Info.plist" "$APP/Contents/Info.plist"
cp "$SRC/AppIcon.icns" "$APP/Contents/Resources/AppIcon.icns"
codesign --force --sign - "$APP"
du -sh "$APP"
