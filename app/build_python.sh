#!/bin/sh
# Build the Python that 执笔 bundles (python-build-standalone + outlook-cli + attachment readers) for one arch into
# .build/python-<arch> and print that path. Wheels are fetched for the target arch, so this cross-builds too.
set -e
ROOT=$(cd "$(dirname "$0")/.." && pwd)
ARCH=$1
PBS=20261021 PY=3.13.17 OUTLOOK_CLI=1.1.14
EXTRAS="openpyxl xlrd pypdf"
DEST="$ROOT/.build/python-$ARCH"
CACHE="$ROOT/.build/cache"
[ -f "$DEST/.built" ] && [ "$DEST/.built" -nt "$0" ] && { echo "$DEST"; exit 0; }

triple() { [ "$1" = arm64 ] && echo aarch64 || echo x86_64; }
sha() { [ "$1" = arm64 ] && echo 8dc60610a82f1e73be7faf964383e02520f95f6e39592a5c37c3809f5a7b798e \
                        || echo ce9c748c8f7525c0052247edb2435fdb1cac89c77e687e0e199eaf2b827668c2; }
# Unpack python-build-standalone for <arch> into <dir>.
unpack() {
  f="$CACHE/cpython-$PY+$PBS-$(triple "$1")-apple-darwin-install_only_stripped.tar.gz"
  [ -f "$f" ] || curl -fsSL -o "$f" "https://github.com/astral-sh/python-build-standalone/releases/download/$PBS/$(basename "$f")"
  echo "$(sha "$1")  $f" | shasum -a 256 -c - >/dev/null
  rm -rf "$2" && mkdir -p "$2" && tar -xzf "$f" -C "$2" --strip-components 1
}
case $ARCH in arm64|x86_64) ;; *) echo "usage: build_python.sh arm64|x86_64" >&2; exit 2;; esac
mkdir -p "$CACHE/wheels"

# pip runs on this Mac's own arch; it only downloads and unpacks wheels for the target.
TOOL="$CACHE/tool-$(uname -m)"
[ -x "$TOOL/bin/python3" ] || unpack "$(uname -m)" "$TOOL"
WHEEL="$CACHE/wheels/outlook_cli-$OUTLOOK_CLI-py3-none-any.whl"
[ -f "$WHEEL" ] || "$TOOL/bin/python3" -m pip wheel -q --no-deps -w "$CACHE/wheels" \
  "https://github.com/fatecannotbealtered/outlook-cli/archive/refs/tags/v$OUTLOOK_CLI.tar.gz"

unpack "$ARCH" "$DEST"
cd "$DEST"
LIB=lib/python3.13
# Headless runtime: no Tk, IDLE, pip, headers or the shared libpython (the python binary is static).
rm -rf include share lib/pkgconfig lib/libpython* lib/libtcl* lib/tcl* lib/tk* lib/itcl* lib/thread* \
  bin/idle* bin/pip* bin/pydoc* bin/python*-config \
  $LIB/idlelib $LIB/tkinter $LIB/turtle* $LIB/turtledemo $LIB/ensurepip $LIB/site-packages/pip*
"$TOOL/bin/python3" -m pip install -q --disable-pip-version-check --only-binary=:all: --platform "macosx_14_0_$ARCH" \
  --python-version 3.13 --implementation cp --abi cp313 --target "$LIB/site-packages" "$WHEEL" $EXTRAS
rm -rf "$LIB/site-packages/bin"
"$TOOL/bin/python3" -m compileall -q -j0 "$LIB" >/dev/null
# Universal2 wheels carry the other arch too; keep one slice and re-sign every binary ad hoc.
find . -type f \( -name '*.so' -o -name '*.dylib' -o -perm -u+x \) | while read -r f; do
  file -b "$f" | grep -q Mach-O || continue
  if [ "$(lipo -archs "$f")" != "$ARCH" ]; then lipo "$f" -thin "$ARCH" -output "$f.thin" && mv "$f.thin" "$f"; fi
  codesign -f -s - "$f" 2>/dev/null
done
touch .built
echo "$DEST"
