#!/bin/sh
# Build 执笔. Default: this Mac's arch into ~/Applications/执笔.app, signed with an Apple Development identity when
# there is one (keeps the Accessibility grant across rebuilds), then (re)launch it.
# --release: one ad-hoc signed zip per arch in dist/ for GitHub Releases and the Homebrew cask.
set -e
ROOT=$(cd "$(dirname "$0")/.." && pwd)
PLIST="$ROOT/app/Info.plist"
VERSION=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$PLIST")
MIN=$(/usr/libexec/PlistBuddy -c 'Print :LSMinimumSystemVersion' "$PLIST")
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

ICNS="$ROOT/.build/AppIcon.icns"
if [ ! -f "$ICNS" ] || [ "$ROOT/app/make_icon.swift" -nt "$ICNS" ]; then
  mkdir -p "$ROOT/.build"
  swiftc -O "$ROOT/app/make_icon.swift" -o "$TMP/make_icon"
  "$TMP/make_icon" "$TMP/AppIcon.iconset"
  iconutil -c icns "$TMP/AppIcon.iconset" -o "$ICNS"
fi

# make_app <arch> <app path> <signing identity>
make_app() {
  arch=$1 app=$2 id=$3
  rm -rf "$app"
  mkdir -p "$app/Contents/MacOS" "$app/Contents/Resources/scripts" "$app/Contents/Resources/references"
  swiftc -O -target "$arch-apple-macos$MIN" "$ROOT/app/Zhibi.swift" -o "$app/Contents/MacOS/Zhibi"
  swiftc -O -target "$arch-apple-macos$MIN" "$ROOT/scripts/outlook_compose.swift" -o "$app/Contents/MacOS/outlook_compose"
  cp "$ICNS" "$ROOT/app/pane.html" "$app/Contents/Resources/"
  cp "$PLIST" "$app/Contents/Info.plist"
  cp "$ROOT"/scripts/*.py "$ROOT/scripts/outlook_compose.sh" "$app/Contents/Resources/scripts/"
  cp "$ROOT"/references/*.md "$app/Contents/Resources/references/"
  ditto "$("$ROOT/app/build_python.sh" "$arch")" "$app/Contents/Resources/python"
  rm -f "$app/Contents/Resources/python/.built"
  codesign --force --sign "$id" "$app/Contents/MacOS/outlook_compose"
  codesign --force --sign "$id" "$app"
}

if [ "$1" = "--release" ]; then
  rm -rf "$ROOT/dist"
  for arch in arm64 x86_64; do
    make_app "$arch" "$ROOT/dist/$arch/Zhibi.app" -
    (cd "$ROOT/dist/$arch" && ditto -c -k --keepParent Zhibi.app "../Zhibi-$VERSION-$arch.zip")
    rm -rf "$ROOT/dist/$arch"   # a second copy with the same bundle id confuses LaunchServices (URL scheme, open -a)
  done
  cd "$ROOT/dist"
  shasum -a 256 Zhibi-*.zip | tee SHA256SUMS
else
  ID=$(security find-identity -v -p codesigning | awk -F'"' '/Apple Development|Developer ID Application/ {print $2; exit}')
  APP="$HOME/Applications/执笔.app"
  make_app "$(uname -m)" "$APP" "${ID:--}"
  pkill -x Zhibi && sleep 1 || true
  open "$APP"
  echo "$APP"
fi
