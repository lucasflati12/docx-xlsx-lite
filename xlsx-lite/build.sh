#!/bin/bash
set -euo pipefail
TASK_SOURCE_DIR="$(cd "$(dirname "$0")" && pwd)"
TASK_APP_PATH="${1:-$HOME/Desktop/Lettore XLSX.app}"
mkdir -p "$TASK_APP_PATH/Contents/MacOS" "$TASK_APP_PATH/Contents/Resources"
xcrun swiftc -O -whole-module-optimization -swift-version 5 \
    -target arm64-apple-macosx13.0 -module-cache-path /private/tmp/lettorexlsx-module-cache \
    "$TASK_SOURCE_DIR/main.swift" -o "$TASK_APP_PATH/Contents/MacOS/LettoreXLSX"
cp "$TASK_SOURCE_DIR/Info.plist" "$TASK_APP_PATH/Contents/Info.plist"
if [[ -f "$TASK_SOURCE_DIR/AppIcon.icns" ]]; then
    cp "$TASK_SOURCE_DIR/AppIcon.icns" "$TASK_APP_PATH/Contents/Resources/AppIcon.icns"
fi
strip -x "$TASK_APP_PATH/Contents/MacOS/LettoreXLSX"
codesign --force --sign - "$TASK_APP_PATH"
du -sh "$TASK_APP_PATH"
