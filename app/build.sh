#!/bin/sh
# Build 执笔. Default: this Mac's arch into ~/Applications/执笔.app, signed with an Apple Development identity when
# there is one (keeps the Accessibility grant across rebuilds), then (re)launch it.
# --release: universal and ad-hoc signed, zipped into dist/ for GitHub Releases and the Homebrew cask.
set -e
ROOT=$(cd "$(dirname "$0")/.." && pwd)
PLIST="$ROOT/app/Info.plist"
VERSION=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$PLIST")
MIN=$(/usr/libexec/PlistBuddy -c 'Print :LSMinimumSystemVersion' "$PLIST")
if [ "$1" = "--release" ]; then
  ARCHS="arm64 x86_64" APP="$ROOT/dist/Zhibi.app" ID=-
  rm -rf "$ROOT/dist"
else
  ARCHS=$(uname -m) APP="$HOME/Applications/执笔.app"
  [ -d "$HOME/Applications/OutlookAI.app" ] && APP="$HOME/Applications/OutlookAI.app"   # Accessibility grants are per path
  ID=$(security find-identity -v -p codesigning | awk -F'"' '/Apple Development|Developer ID Application/ {print $2; exit}')
fi
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

build() {  # build <output> <swift source>
  for a in $ARCHS; do swiftc -O -target "$a-apple-macos$MIN" "$2" -o "$TMP/$(basename "$1").$a"; done
  lipo -create "$TMP/$(basename "$1")".* -output "$1"
}

rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources/scripts" "$APP/Contents/Resources/references"
build "$APP/Contents/MacOS/Zhibi" "$ROOT/app/Zhibi.swift"
build "$APP/Contents/MacOS/outlook_compose" "$ROOT/scripts/outlook_compose.swift"

ICNS="$ROOT/.build/AppIcon.icns"
if [ ! -f "$ICNS" ] || [ "$ROOT/app/make_icon.swift" -nt "$ICNS" ]; then
  mkdir -p "$ROOT/.build"
  swiftc -O "$ROOT/app/make_icon.swift" -o "$TMP/make_icon"
  "$TMP/make_icon" "$TMP/AppIcon.iconset"
  iconutil -c icns "$TMP/AppIcon.iconset" -o "$ICNS"
fi
cp "$ICNS" "$ROOT/app/pane.html" "$APP/Contents/Resources/"
cp "$PLIST" "$APP/Contents/Info.plist"
cp "$ROOT"/scripts/*.py "$ROOT/scripts/outlook_compose.sh" "$APP/Contents/Resources/scripts/"
cp "$ROOT"/references/*.md "$APP/Contents/Resources/references/"

codesign --force --sign "${ID:--}" "$APP/Contents/MacOS/outlook_compose"
codesign --force --sign "${ID:--}" "$APP"

if [ "$1" = "--release" ]; then
  cd "$ROOT/dist"
  ditto -c -k --keepParent Zhibi.app "Zhibi-$VERSION.zip"
  shasum -a 256 "Zhibi-$VERSION.zip" | awk '{print $1}' > "Zhibi-$VERSION.zip.sha256"
  rm -rf Zhibi.app   # a second copy with the same bundle id confuses LaunchServices (URL scheme, open -a)
  echo "$ROOT/dist/Zhibi-$VERSION.zip $(cat "Zhibi-$VERSION.zip.sha256")"
else
  pkill -x Zhibi && sleep 1 || true
  open "$APP"
  echo "$APP"
fi
