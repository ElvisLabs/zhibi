#!/bin/sh
# Run the Swift helper: prebuilt inside the 执笔 app bundle, else built on first use (or after the source changes).
DIR=$(cd "$(dirname "$0")" && pwd)
[ -x "$DIR/../../MacOS/outlook_compose" ] && exec "$DIR/../../MacOS/outlook_compose" "$@"
BIN="$DIR/.build/outlook_compose"
if [ ! -x "$BIN" ] || [ "$DIR/outlook_compose.swift" -nt "$BIN" ]; then
  mkdir -p "$DIR/.build" && swiftc -O "$DIR/outlook_compose.swift" -o "$BIN" >&2 || exit 5
fi
exec "$BIN" "$@"
