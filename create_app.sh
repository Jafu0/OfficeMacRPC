#!/bin/bash

APP_NAME="OfficeRichPresence"
APP_DIR="$APP_NAME.app"
CONTENTS_DIR="$APP_DIR/Contents"
MACOS_DIR="$CONTENTS_DIR/MacOS"
RESOURCES_DIR="$CONTENTS_DIR/Resources"

echo "Cleaning up..."
pkill -f "OfficeRichPresence" || true
rm -rf "$APP_DIR"
rm -f office-rich-presence

echo "Compiling Swift..."
swiftc Sources/main.swift -target arm64-apple-macosx13.0 -o office-rich-presence || { echo "Swift compilation failed"; exit 1; }

echo "Creating App Bundle Structure..."
mkdir -p "$MACOS_DIR"
mkdir -p "$RESOURCES_DIR"

echo "Copying Executable..."
cp office-rich-presence "$MACOS_DIR/$APP_NAME"
chmod +x "$MACOS_DIR/$APP_NAME"

SRC_DIR="."

echo "Copying Resources..."
cp "$SRC_DIR/index.js" "$RESOURCES_DIR/"
cp "menubaricon.png" "$RESOURCES_DIR/menubar_icon.png"
cp package.json "$RESOURCES_DIR/"

if [ -d "node_modules" ]; then
    echo "Copying node_modules..."
    cp -R "node_modules" "$RESOURCES_DIR/"
else
    echo "Warning: node_modules not found. App may not work correctly."
fi

if [ -f "icon.png" ]; then
    echo "Creating Icon..."
    mkdir -p icon.iconset
    sips -z 16 16     icon.png --setProperty format png --out icon.iconset/icon_16x16.png
    sips -z 32 32     icon.png --setProperty format png --out icon.iconset/icon_16x16@2x.png
    sips -z 32 32     icon.png --setProperty format png --out icon.iconset/icon_32x32.png
    sips -z 64 64     icon.png --setProperty format png --out icon.iconset/icon_32x32@2x.png
    sips -z 128 128   icon.png --setProperty format png --out icon.iconset/icon_128x128.png
    sips -z 256 256   icon.png --setProperty format png --out icon.iconset/icon_128x128@2x.png
    sips -z 256 256   icon.png --setProperty format png --out icon.iconset/icon_256x256.png
    sips -z 512 512   icon.png --setProperty format png --out icon.iconset/icon_256x256@2x.png
    sips -z 512 512   icon.png --setProperty format png --out icon.iconset/icon_512x512.png
    sips -z 1024 1024 icon.png --setProperty format png --out icon.iconset/icon_512x512@2x.png
    
    iconutil -c icns icon.iconset
    cp icon.icns "$RESOURCES_DIR/AppIcon.icns"
    rm -rf icon.iconset
else
    echo "Warning: icon.png not found, skipping icon creation."
fi

VERSION=$(date +"2.0.%Y%m%d.%H%M")
echo "Building Version: $VERSION"

echo "Creating Info.plist..."
cat > "$CONTENTS_DIR/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>$APP_NAME</string>
    <key>CFBundleIdentifier</key>
    <string>cl.jafu.OfficeRichPresence</string>
    <key>CFBundleName</key>
    <string>$APP_NAME</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundleShortVersionString</key>
    <string>$VERSION</string>
    <key>CFBundleVersion</key>
    <string>$VERSION</string>
    <key>LSUIElement</key>
    <true/>
    <key>NSAppleEventsUsageDescription</key>
    <string>OfficeRichPresence needs to control Office applications to read the active document name.</string>
    <key>NSHumanReadableCopyright</key>
    <string>Copyright © 2025 Jafu. Released under the MIT License.</string>
    <key>LSMinimumSystemVersion</key>
    <string>13.0</string>
</dict>
</plist>
EOF

echo "Done! $APP_NAME created."
